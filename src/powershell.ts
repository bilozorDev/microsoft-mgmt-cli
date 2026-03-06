import type { Subprocess } from "bun";

const EXEC_MARKER = "---PROFULGENT-EXEC---";
const END_MARKER = "---PROFULGENT-END-MARKER---";
const ERROR_MARKER = "---PROFULGENT-ERROR---";

// PowerShell script that runs a command loop, reading commands from stdin
// line-by-line using .NET's Console.In.ReadLine() (which processes one line
// at a time even with piped stdin, unlike PowerShell's -Command - mode which
// buffers ALL stdin before executing).
// Commands are accumulated until the EXEC marker, then executed.
// Markers are written directly via [Console]::Out to bypass pipeline buffering.
const LOOP_SCRIPT = `
$sb = [System.Text.StringBuilder]::new()
while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }
    if ($line -eq '${EXEC_MARKER}') {
        $cmd = $sb.ToString()
        [void]$sb.Clear()
        try {
            Invoke-Expression $cmd
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
  ): Promise<{ output: string; error: string }> {
    if (!this.process?.stdin || !this.process?.stdout) {
      throw new Error("PowerShell session not started");
    }

    this.process.stdin.write(command + "\n" + EXEC_MARKER + "\n");
    this.process.stdin.flush();

    const output = await this.readUntilMarker(this.process.stdout, END_MARKER, onOutput);

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
  ): Promise<string> {
    const reader = stream.getReader();
    let accumulated = "";

    try {
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        accumulated += this.decoder.decode(value, { stream: true });
        if (!accumulated.includes(marker)) {
          onOutput?.(accumulated);
        }
        if (accumulated.includes(marker)) break;
      }
    } finally {
      reader.releaseLock();
    }

    return accumulated;
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

  async end(): Promise<void> {
    if (!this.process) return;

    try {
      await this.runCommand("Disconnect-ExchangeOnline -Confirm:$false");
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
