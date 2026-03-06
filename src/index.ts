import * as p from "@clack/prompts";
import { PowerShellSession } from "./powershell.ts";
import { checkRequirements } from "./requirements.ts";
import { run as whitelistDomain } from "./commands/whitelist-domain.ts";
import { run as createUser } from "./commands/create-user.ts";

const ps = new PowerShellSession();

async function cleanup() {
  p.log.info("Disconnecting from Exchange Online...");
  await ps.end();
  p.outro("Goodbye!");
  process.exit(0);
}

process.on("SIGINT", () => void cleanup());
process.on("SIGTERM", () => void cleanup());

async function main() {
  p.intro("Profulgent — Exchange Online Admin CLI");

  await checkRequirements();

  // Start PowerShell session
  const connectSpin = p.spinner();
  connectSpin.start("Starting PowerShell session...");
  try {
    await ps.start();
    connectSpin.stop("PowerShell session ready.");
  } catch (e) {
    connectSpin.stop("Failed to start PowerShell.");
    p.log.error(`Could not start pwsh. Is PowerShell Core installed?\n${e}`);
    process.exit(1);
  }

  // Connect to Exchange Online (opens browser for auth)
  const authSpin = p.spinner();
  authSpin.start("Connecting to Exchange Online (check your browser)...");
  try {
    await ps.connectExchangeOnline();
    authSpin.stop("Connected to Exchange Online.");
  } catch (e) {
    authSpin.stop("Connection failed.");
    p.log.error(`${e}`);
    await ps.end();
    process.exit(1);
  }

  // Fetch tenant domain
  try {
    const domain = await ps.getTenantDomain();
    p.log.success(`Connected to: ${domain}`);
  } catch {
    p.log.warn("Could not determine tenant domain.");
  }

  // Main menu loop
  while (true) {
    const action = await p.select({
      message: "What would you like to do?",
      options: [
        { value: "create-user", label: "Create user", hint: "will prompt to login" },
        { value: "whitelist-domain", label: "Whitelist domain(s)" },
        { value: "exit", label: "Exit" },
      ],
    });

    if (p.isCancel(action) || action === "exit") {
      break;
    }

    switch (action) {
      case "create-user":
        await createUser(ps);
        break;
      case "whitelist-domain":
        await whitelistDomain(ps);
        break;
    }
  }

  await cleanup();
}

main().catch(async (e) => {
  p.log.error(`Unexpected error: ${e}`);
  await ps.end();
  process.exit(1);
});
