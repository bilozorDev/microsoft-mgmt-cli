import * as p from "@clack/prompts";

async function isPwshAvailable(): Promise<boolean> {
  try {
    const proc = Bun.spawn(["pwsh", "-v"], {
      stdout: "pipe",
      stderr: "pipe",
    });
    const code = await proc.exited;
    return code === 0;
  } catch {
    return false;
  }
}

function getPwshInstallHint(): string {
  switch (process.platform) {
    case "darwin":
      return "brew install powershell";
    case "win32":
      return "winget install Microsoft.PowerShell";
    default:
      return "See https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux";
  }
}

async function checkModules(): Promise<string[]> {
  const proc = Bun.spawn(
    [
      "pwsh",
      "-NoLogo",
      "-NoProfile",
      "-Command",
      `@{
  ExchangeOnline = [bool](Get-Module -ListAvailable -Name ExchangeOnlineManagement)
  MicrosoftGraph = [bool](Get-Module -ListAvailable -Name Microsoft.Graph)
} | ConvertTo-Json`,
    ],
    { stdout: "pipe", stderr: "pipe" },
  );

  const output = await new Response(proc.stdout).text();
  const code = await proc.exited;
  if (code !== 0) {
    throw new Error("Failed to check PowerShell modules");
  }

  const result = JSON.parse(output) as {
    ExchangeOnline: boolean;
    MicrosoftGraph: boolean;
  };

  const missing: string[] = [];
  if (!result.ExchangeOnline) missing.push("ExchangeOnlineManagement");
  if (!result.MicrosoftGraph) missing.push("Microsoft.Graph");
  return missing;
}

export async function checkRequirements(): Promise<void> {
  // 1. Check pwsh
  if (!(await isPwshAvailable())) {
    const hint = getPwshInstallHint();
    const canAutoInstall =
      process.platform === "darwin" || process.platform === "win32";

    p.log.error("PowerShell Core (pwsh) is not installed.");

    if (canAutoInstall) {
      const install = await p.confirm({
        message: `Install PowerShell Core now? (${hint})`,
      });

      if (p.isCancel(install) || !install) {
        p.log.info(`Install manually: ${hint}`);
        process.exit(1);
      }

      p.log.info("Installing PowerShell Core (this may take a minute)...");
      const cmd =
        process.platform === "darwin"
          ? ["brew", "install", "powershell"]
          : ["winget", "install", "Microsoft.PowerShell", "--accept-source-agreements", "--accept-package-agreements"];
      const installProc = Bun.spawn(cmd, {
        stdin: "inherit",
        stdout: "inherit",
        stderr: "inherit",
      });
      const code = await installProc.exited;
      if (code !== 0) {
        p.log.error(`Installation failed. Install manually: ${hint}`);
        process.exit(1);
      }
      p.log.success("PowerShell Core installed. Please close and reopen this app.");
      process.exit(0);
    } else {
      p.log.info(`Install manually:\n  ${hint}`);
      process.exit(1);
    }
  }

  // 2. Check modules
  const missing = await checkModules();
  if (missing.length === 0) return;

  // 3. Show what's missing and auto-install
  p.log.warn(`Missing module${missing.length > 1 ? "s" : ""}: ${missing.join(", ")}`);

  const install = await p.confirm({
    message: "Install missing modules now?",
  });

  if (p.isCancel(install) || !install) {
    process.exit(1);
  }

  for (const m of missing) {
    p.log.info(`Installing ${m} (this may take a minute)...`);
    const proc = Bun.spawn(
      [
        "pwsh",
        "-NoLogo",
        "-NoProfile",
        "-Command",
        `Install-Module -Name ${m} -Scope CurrentUser -Force -AllowClobber`,
      ],
      { stdin: "inherit", stdout: "inherit", stderr: "inherit" },
    );
    const code = await proc.exited;
    if (code !== 0) {
      p.log.error(`Failed to install ${m}.`);
      process.exit(1);
    }
    p.log.success(`${m} installed.`);
  }
}
