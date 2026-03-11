import type { Subprocess } from "bun";

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
  private graphScopeLevel: "none" | "read" | "readwrite" = "none";

  private static readonly READ_SCOPES = [
    "User.Read.All",
    "Organization.Read.All",
    "Directory.Read.All",
    "RoleManagement.Read.Directory",
    "Group.Read.All",
    "GroupMember.Read.All",
    "AuditLog.Read.All",
    "UserAuthenticationMethod.Read.All",
  ];

  private static readonly READWRITE_SCOPES = [
    "User.ReadWrite.All",
    "Organization.Read.All",
    "Directory.ReadWrite.All",
    "RoleManagement.Read.Directory",
    "Group.ReadWrite.All",
    "GroupMember.ReadWrite.All",
    "AuditLog.Read.All",
    "UserAuthenticationMethod.ReadWrite.All",
    "User-PasswordProfile.ReadWrite.All",
  ];
  tenantDomain: string | null = null;

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

    const output = await this.readUntilMarker(this.process.stdout, END_MARKER, onOutput, timeout);

    const hasError = output.includes(ERROR_MARKER);
    const cleanOutput = output
      .replace(ERROR_MARKER, "")
      .replace(END_MARKER, "")
      .trim();

    if (hasError) {
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

  async ensureGraphConnected(needsWrite = false): Promise<void> {
    if (this.graphScopeLevel === "readwrite") return;
    if (this.graphScopeLevel === "read" && !needsWrite) return;

    // Disconnect existing session if upgrading from read to readwrite
    if (this.graphScopeLevel === "read" && needsWrite) {
      try {
        await this.runCommand("Disconnect-MgGraph *>$null");
      } catch {
        // Best-effort disconnect
      }
      this.graphScopeLevel = "none";
    }

    const scopes = needsWrite
      ? PowerShellSession.READWRITE_SCOPES
      : PowerShellSession.READ_SCOPES;
    const scopeStr = scopes.map((s) => `"${s}"`).join(",");
    const { error } = await this.runCommand(
      `Connect-MgGraph -Scopes ${scopeStr} -NoWelcome`,
    );
    if (error) {
      throw new Error(`Failed to connect to Microsoft Graph: ${error}`);
    }
    this.graphScopeLevel = needsWrite ? "readwrite" : "read";
  }

  async connectExchangeOnline(): Promise<void> {
    const { error } = await this.runCommand(
      "Connect-ExchangeOnline -ShowBanner:$false",
    );
    if (error) {
      throw new Error(`Failed to connect to Exchange Online: ${error}`);
    }
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

  async switchTenant(): Promise<void> {
    if (this.graphScopeLevel !== "none") {
      try {
        await this.runCommand("Disconnect-MgGraph *>$null");
      } catch {
        // Best-effort disconnect
      }
      this.graphScopeLevel = "none";
    }

    try {
      await this.runCommand("Disconnect-ExchangeOnline -Confirm:$false *>$null");
    } catch {
      // Best-effort disconnect
    }

    await this.connectExchangeOnline();
    await this.getTenantDomain();
  }

  async end(): Promise<void> {
    if (!this.process) return;

    if (this.graphScopeLevel !== "none") {
      try {
        await this.runCommand("Disconnect-MgGraph *>$null");
      } catch {
        // Best-effort disconnect
      }
      this.graphScopeLevel = "none";
    }

    try {
      await this.runCommand("Disconnect-ExchangeOnline -Confirm:$false *>$null");
    } catch {
      // Best-effort disconnect
    }

    try {
      this.process.stdin?.end();
      this.process.kill();
    } catch {
      // Process may already be dead
    }

    this.process = null;
    this.tenantDomain = null;
  }
}
