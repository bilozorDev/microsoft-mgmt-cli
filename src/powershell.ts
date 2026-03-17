import type { Subprocess } from "bun";
import {
  reportPowerShellError,
  reportPowerShellTimeout,
  addBreadcrumb,
} from "./telemetry.ts";

const EXEC_MARKER = "---PROFULGENT-EXEC---";
const END_MARKER = "---PROFULGENT-END-MARKER---";
const ERROR_MARKER = "---PROFULGENT-ERROR---";

// PowerShell script that runs a command loop, reading commands from stdin
// line-by-line using .NET's Console.In.ReadLine() (which processes one line
// at a time even with piped stdin, unlike PowerShell's -Command - mode which
// buffers ALL stdin before executing).
// Commands are accumulated until the EXEC marker, then executed.
// Output is captured via Out-String and written through [Console]::Out so it
// stays ordered relative to the end/error markers (Invoke-Expression output
// goes through PowerShell's async formatting pipeline, which can race with
// direct [Console]::Out writes and leak into subsequent commands).
const LOOP_SCRIPT = `
$sb = [System.Text.StringBuilder]::new()
while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }
    if ($line -eq '${EXEC_MARKER}') {
        $cmd = $sb.ToString()
        [void]$sb.Clear()
        try {
            $ErrorActionPreference = 'Stop'
            $__out = @(Invoke-Expression $cmd) | Out-String
            if ($__out.Length -gt 0) { [Console]::Out.Write($__out) }
        } catch {
            [Console]::Out.WriteLine($_.Exception.Message)
            [Console]::Out.WriteLine('${ERROR_MARKER}')
        }
        [Console]::Out.WriteLine('${END_MARKER}')
        [Console]::Out.Flush()
    } else {
        [void]$sb.AppendLine($line)
    }
}
`;

export class PowerShellSession {
  private process: Subprocess<"pipe", "pipe", "inherit"> | null = null;
  private decoder = new TextDecoder();
  private grantedScopes = new Set<string>();
  private exchangeConnected = false;
  tenantId: string | null = null;
  tenantDomain: string | null = null;
  onTenantResolved: ((domain: string) => void) | null = null;

  get isExchangeConnected(): boolean {
    return this.exchangeConnected;
  }

  async start(): Promise<void> {
    this.process = Bun.spawn(["pwsh", "-NoLogo", "-NoProfile", "-Command", LOOP_SCRIPT], {
      stdin: "pipe",
      stdout: "pipe",
      stderr: "inherit",
    });
  }

  async runCommand(
    command: string,
    onOutput?: (accumulated: string) => void,
    timeout?: number,
  ): Promise<{ output: string; error: string }> {
    if (!this.process?.stdin || !this.process?.stdout) {
      throw new Error("PowerShell session not started");
    }

    this.process.stdin.write(command + "\n" + EXEC_MARKER + "\n");
    this.process.stdin.flush();

    let output: string;
    try {
      output = await this.readUntilMarker(this.process.stdout, END_MARKER, onOutput, timeout);
    } catch (e) {
      if (e instanceof Error && e.message.includes("timed out")) {
        reportPowerShellTimeout(command, timeout ?? 120_000);
      }
      throw e;
    }

    const hasError = output.includes(ERROR_MARKER);
    const cleanOutput = output
      .replace(ERROR_MARKER, "")
      .replace(END_MARKER, "")
      .trim();

    if (hasError) {
      reportPowerShellError(command, cleanOutput);
      return { output: cleanOutput, error: cleanOutput };
    }

    return { output: cleanOutput, error: "" };
  }

  private async readUntilMarker(
    stream: ReadableStream<Uint8Array>,
    marker: string,
    onOutput?: (accumulated: string) => void,
    timeoutMs = 120_000,
  ): Promise<string> {
    const reader = stream.getReader();
    let accumulated = "";

    const readLoop = async () => {
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        accumulated += this.decoder.decode(value, { stream: true });
        if (!accumulated.includes(marker)) {
          onOutput?.(accumulated);
        }
        if (accumulated.includes(marker)) break;
      }
      return accumulated;
    };

    let timer: ReturnType<typeof setTimeout>;
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(
        () => reject(new Error(`PowerShell command timed out after ${timeoutMs / 1000}s`)),
        timeoutMs,
      );
    });

    try {
      return await Promise.race([readLoop(), timeout]);
    } finally {
      clearTimeout(timer!);
      reader.releaseLock();
    }
  }

  async runCommandJson<T>(command: string): Promise<T | null> {
    const { output, error } = await this.runCommand(
      `${command} | ConvertTo-Json -Depth 5 -Compress`,
    );
    if (error) throw new Error(error);
    if (!output) return null;
    return JSON.parse(output) as T;
  }

  private isScopeCovered(scope: string): boolean {
    if (this.grantedScopes.has(scope)) return true;
    // ReadWrite covers Read (e.g. User.ReadWrite.All covers User.Read.All)
    const rwVariant = scope.replace(".Read.", ".ReadWrite.");
    if (rwVariant !== scope && this.grantedScopes.has(rwVariant)) return true;
    return false;
  }

  async ensureExchangeConnected(): Promise<void> {
    if (this.exchangeConnected) {
      // Retry tenant metadata if it failed on first connect
      if (!this.tenantDomain) {
        try {
          if (!this.tenantId) await this.extractTenantId();
          await this.getTenantDomain();
          if (this.tenantDomain) {
            this.onTenantResolved?.(this.tenantDomain);
          }
        } catch {
          // Best-effort metadata retry
        }
      }
      return;
    }

    await this.connectExchangeOnline();
    await this.extractTenantId();
    try {
      await this.getTenantDomain();
      if (this.tenantDomain) {
        this.onTenantResolved?.(this.tenantDomain);
      }
    } catch {
      // Non-fatal — domain is metadata, connection itself succeeded
    }
  }

  async ensureGraphConnected(requiredScopes: string[]): Promise<void> {
    const missing = requiredScopes.filter((s) => !this.isScopeCovered(s));
    if (missing.length === 0) return;

    // Try to get tenantId from Exchange if connected, but don't force Exchange connection
    if (!this.tenantId && this.exchangeConnected) {
      await this.extractTenantId();
    }

    if (this.grantedScopes.size > 0) {
      try {
        await this.runCommand("Disconnect-MgGraph *>$null");
      } catch {
        // Best-effort disconnect
      }
    }

    const allScopes = new Set([...this.grantedScopes, ...requiredScopes]);
    const scopeStr = [...allScopes].map((s) => `"${s}"`).join(",");
    let graphCmd = `Connect-MgGraph -Scopes ${scopeStr} -ContextScope Process -NoWelcome`;
    if (this.tenantId) {
      graphCmd += ` -TenantId '${this.tenantId.replace(/'/g, "''")}'`;
    }

    const { error } = await this.runCommand(graphCmd);
    if (error) {
      throw new Error(`Failed to connect to Microsoft Graph: ${error}`);
    }

    // Verify Graph connected to the same tenant as Exchange Online
    if (this.tenantId) {
      const { output: graphTenant } = await this.runCommand(
        "Get-MgContext | Select-Object -ExpandProperty TenantId",
      );
      if (graphTenant.trim() && graphTenant.trim() !== this.tenantId) {
        throw new Error(
          `Tenant mismatch: Exchange Online is connected to ${this.tenantId} but Graph connected to ${graphTenant.trim()}. Please restart the app and authenticate with the correct account.`,
        );
      }
    }

    this.grantedScopes = allScopes;
    addBreadcrumb({
      category: "lifecycle",
      message: "Graph connected",
      data: { scopeCount: allScopes.size },
    });
  }

  async getGraphAccessToken(): Promise<string> {
    const { output, error } = await this.runCommand(
      `(Get-MgContext).AccessToken`,
    );
    if (error || !output.trim()) {
      throw new Error(
        "Failed to retrieve Graph access token from PowerShell session",
      );
    }
    return output.trim();
  }

  async connectExchangeOnline(): Promise<void> {
    let cmd = "Connect-ExchangeOnline -ShowBanner:$false";

    // On Windows, disable WAM to prevent OS-level cached account auto-selection
    const useDisableWam = process.platform === "win32";
    if (useDisableWam) {
      cmd += " -DisableWAM";
    }

    const { error } = await this.runCommand(cmd);

    if (error) {
      // If -DisableWAM is not recognized (EXO module < 3.7.2), retry without it
      if (useDisableWam && error.includes("DisableWAM")) {
        const fallback = await this.runCommand(
          "Connect-ExchangeOnline -ShowBanner:$false",
        );
        if (fallback.error) {
          throw new Error(`Failed to connect to Exchange Online: ${fallback.error}`);
        }
        this.exchangeConnected = true;
        addBreadcrumb({ category: "lifecycle", message: "Exchange Online connected" });
        return;
      }
      throw new Error(`Failed to connect to Exchange Online: ${error}`);
    }
    this.exchangeConnected = true;
    addBreadcrumb({ category: "lifecycle", message: "Exchange Online connected" });
  }

  async extractTenantId(): Promise<string | null> {
    const { output, error } = await this.runCommand(
      "Get-ConnectionInformation | Select-Object -First 1 -ExpandProperty TenantID",
    );
    if (error || !output.trim()) {
      return null;
    }
    this.tenantId = output.trim();
    return this.tenantId;
  }

  async getTenantDomain(): Promise<string> {
    const { output, error } = await this.runCommand(
      "Get-AcceptedDomain | Where-Object {$_.Default -eq $true} | Select-Object -ExpandProperty DomainName"
    );
    if (error) {
      throw new Error(`Failed to get tenant domain: ${error}`);
    }
    this.tenantDomain = output.trim();
    return this.tenantDomain;
  }

  async disconnectTenant(): Promise<void> {
    if (this.grantedScopes.size > 0) {
      try {
        await this.runCommand("Disconnect-MgGraph *>$null");
      } catch {
        // Best-effort disconnect
      }
      this.grantedScopes = new Set();
    }

    if (this.exchangeConnected) {
      try {
        await this.runCommand("Disconnect-ExchangeOnline -Confirm:$false *>$null");
      } catch {
        // Best-effort disconnect
      }
    }

    this.exchangeConnected = false;
    this.tenantId = null;
    this.tenantDomain = null;
    addBreadcrumb({ category: "lifecycle", message: "Tenant disconnected" });
  }

  async end(): Promise<void> {
    if (!this.process) return;

    if (this.grantedScopes.size > 0) {
      try {
        await this.runCommand("Disconnect-MgGraph *>$null");
      } catch {
        // Best-effort disconnect
      }
      this.grantedScopes = new Set();
    }

    if (this.exchangeConnected) {
      try {
        await this.runCommand("Disconnect-ExchangeOnline -Confirm:$false *>$null");
      } catch {
        // Best-effort disconnect
      }
      this.exchangeConnected = false;
    }

    try {
      this.process.stdin?.end();
      this.process.kill();
    } catch {
      // Process may already be dead
    }

    this.process = null;
    this.tenantDomain = null;
    this.tenantId = null;
  }
}
