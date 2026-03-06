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
  SharePointOnline = [bool](Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)
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
    SharePointOnline: boolean;
  };

  const missing: string[] = [];
  if (!result.ExchangeOnline) missing.push("ExchangeOnlineManagement");
  if (!result.MicrosoftGraph) missing.push("Microsoft.Graph");
  if (!result.SharePointOnline) missing.push("Microsoft.Online.SharePoint.PowerShell");
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

      const spin = p.spinner();
      spin.start("Installing PowerShell Core...");
      const cmd =
        process.platform === "darwin"
          ? ["brew", "install", "powershell"]
          : ["winget", "install", "Microsoft.PowerShell"];
      const installProc = Bun.spawn(cmd, {
        stdout: "inherit",
        stderr: "inherit",
      });
      const code = await installProc.exited;
      if (code !== 0) {
        spin.stop("Installation failed.");
        p.log.error(`Install manually: ${hint}`);
        process.exit(1);
      }
      spin.stop("PowerShell Core installed.");
    } else {
      p.log.info(`Install manually:\n  ${hint}`);
      process.exit(1);
    }
  }

  // 2. Check modules
  const missing = await checkModules();
  if (missing.length === 0) return;

  // 3. Show what's missing and offer to install
  p.note(
    missing.map((m) => `- ${m}`).join("\n"),
    "Missing PowerShell modules",
  );

  const install = await p.confirm({
    message: "Install missing modules now?",
  });

  if (p.isCancel(install) || !install) {
    p.log.info("Install manually:");
    for (const m of missing) {
      p.log.info(
        `  Install-Module -Name ${m} -Scope CurrentUser -Force`,
      );
    }
    process.exit(1);
  }

  for (const m of missing) {
    const spin = p.spinner();
    spin.start(`Installing ${m}...`);
    const proc = Bun.spawn(
      [
        "pwsh",
        "-NoLogo",
        "-NoProfile",
        "-Command",
        `Install-Module -Name ${m} -Scope CurrentUser -Force`,
      ],
      { stdout: "inherit", stderr: "inherit" },
    );
    const code = await proc.exited;
    if (code !== 0) {
      spin.stop(`Failed to install ${m}.`);
      p.log.error(
        `Install manually: Install-Module -Name ${m} -Scope CurrentUser -Force`,
      );
      process.exit(1);
    }
    spin.stop(`${m} installed.`);
  }
}
