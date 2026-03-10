import { dirname, join } from "path";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "./powershell.ts";

/** Directory where the executable (or script) lives — reports save relative to this. */
export function appDir(): string {
  return dirname(process.execPath);
}

export function escapePS(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Creates a one-time secret link via onetimesecret.com (anonymous, no auth needed).
 * Returns the shareable URL or null on failure.
 */
export async function createSecretLink(
  secret: string,
  ttl: number = 604800,
): Promise<string | null> {
  try {
    const res = await fetch("https://us.onetimesecret.com/api/v1/share", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ secret, ttl: String(ttl) }),
    });
    if (!res.ok) return null;
    const data = (await res.json()) as { secret_key?: string };
    if (!data.secret_key) return null;
    return `https://us.onetimesecret.com/secret/${data.secret_key}`;
  } catch {
    return null;
  }
}

/**
 * Detects the "Enable-OrganizationCustomization" error, offers to run it,
 * and polls until the tenant is ready. Returns true if fixed (caller should retry).
 */
export async function handleOrgCustomizationError(
  ps: PowerShellSession,
  error: string
): Promise<boolean> {
  if (!error.includes("Enable-OrganizationCustomization")) return false;

  p.log.warn(
    "This tenant requires organization customization before policies can be modified."
  );
  p.log.info(
    "Docs: https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/enable-organizationcustomization?view=exchange-ps"
  );

  const shouldEnable = await p.confirm({
    message: "Run Enable-OrganizationCustomization now?",
  });

  if (p.isCancel(shouldEnable) || !shouldEnable) return false;

  p.log.warn("This can take up to 5 minutes.");

  const spin = p.spinner();
  spin.start("Running Enable-OrganizationCustomization...");

  const timerInterval = setInterval(() => {
    const elapsed = Math.floor((Date.now() - enableStart) / 1000);
    const min = Math.floor(elapsed / 60);
    const sec = elapsed % 60;
    spin.message(
      `Running Enable-OrganizationCustomization... ${min}:${sec.toString().padStart(2, "0")}`
    );
  }, 1000);
  const enableStart = Date.now();

  const { error: enableError } = await ps.runCommand(
    "Enable-OrganizationCustomization"
  );

  clearInterval(timerInterval);

  const alreadyEnabled =
    enableError && /already.+enabled|not required/i.test(enableError);

  if (enableError && !alreadyEnabled) {
    spin.stop("Failed to enable organization customization.");
    p.log.error(enableError);
    return false;
  }

  const elapsed = Math.floor((Date.now() - enableStart) / 1000);

  if (alreadyEnabled) {
    spin.stop("Organization customization is already enabled.");
    p.log.warn(
      "If you just enabled it, changes can take up to an hour to propagate. Please try again later."
    );
    return false;
  }

  spin.stop(`Enable-OrganizationCustomization completed in ${elapsed}s.`);

  // Poll until Set- commands actually work (propagation can take a few minutes).
  // We test with a real write operation: add then remove a dummy marker domain.
  const pollSpin = p.spinner();
  pollSpin.start("Waiting for changes to propagate (up to 5 min)...");

  const maxWaitMs = 5 * 60 * 1000;
  const intervalMs = 15_000;
  const pollStart = Date.now();
  const testDomain = "org-customization-test.invalid";

  while (Date.now() - pollStart < maxWaitMs) {
    const { error: testError } = await ps.runCommand(
      `Set-HostedContentFilterPolicy -Identity 'Default' -AllowedSenderDomains @{Add='${testDomain}'}`
    );
    if (!testError) {
      // Clean up the test domain
      await ps.runCommand(
        `Set-HostedContentFilterPolicy -Identity 'Default' -AllowedSenderDomains @{Remove='${testDomain}'}`
      );
      pollSpin.stop("Organization customization is active.");
      return true;
    }
    if (!testError.includes("Enable-OrganizationCustomization")) {
      // Different error — Set- is accepted, propagation is done
      pollSpin.stop("Organization customization is active.");
      return true;
    }
    await Bun.sleep(intervalMs);
  }

  pollSpin.stop("Timed out waiting for propagation.");
  p.log.warn("You may need to wait a few more minutes and try again.");
  return false;
}
