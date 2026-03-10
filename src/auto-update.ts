import * as p from "@clack/prompts";
import { dirname, join } from "path";
import { mkdirSync, rmSync, renameSync, copyFileSync } from "fs";
import { tmpdir } from "os";
import { appDir } from "./utils.ts";
import pkg from "../package.json";

const REPO = "bilozorDev/microsoft-mgmt-cli";

interface GitHubAsset {
  name: string;
  browser_download_url: string;
}

interface GitHubRelease {
  tag_name: string;
  body?: string;
  assets: GitHubAsset[];
}

export async function checkForUpdates(): Promise<void> {
  // Skip in dev mode (not a compiled .exe)
  if (!process.execPath.endsWith(".exe")) return;

  const currentVersion = pkg.version;

  let release: GitHubRelease;
  try {
    const res = await fetch(
      `https://api.github.com/repos/${REPO}/releases/latest`,
      { headers: { Accept: "application/vnd.github.v3+json" } },
    );
    if (!res.ok) return; // no releases or network issue
    release = (await res.json()) as GitHubRelease;
  } catch {
    // Network error — continue without updating
    return;
  }

  const latestVersion = release.tag_name.replace(/^v/, "");
  if (latestVersion <= currentVersion) return;

  // Show update info
  p.log.info(`Update available: v${currentVersion} → v${latestVersion}`);
  if (release.body) {
    p.log.info(`Release notes:\n${release.body}`);
  }

  const shouldUpdate = await p.confirm({
    message: "Would you like to update now?",
  });
  if (p.isCancel(shouldUpdate) || !shouldUpdate) return;

  // Find the .zip asset
  const zipAsset = release.assets.find((a) => a.name.endsWith(".zip"));
  if (!zipAsset) {
    p.log.warn("No zip asset found in the release. Skipping update.");
    return;
  }

  const spin = p.spinner();
  spin.start("Downloading update...");

  const tempDir = join(tmpdir(), `m365-update-${Date.now()}`);
  const tempZip = join(tempDir, "update.zip");
  const extractDir = join(tempDir, "extracted");

  try {
    mkdirSync(tempDir, { recursive: true });

    // Download zip with progress
    const dlRes = await fetch(zipAsset.browser_download_url);
    if (!dlRes.ok || !dlRes.body) {
      spin.stop("Download failed.");
      p.log.warn("Could not download the update. Continuing with current version.");
      return;
    }
    const totalBytes = Number(dlRes.headers.get("content-length") || 0);
    let downloaded = 0;
    const chunks: Uint8Array[] = [];
    for await (const chunk of dlRes.body) {
      chunks.push(chunk);
      downloaded += chunk.length;
      if (totalBytes > 0) {
        const pct = Math.round((downloaded / totalBytes) * 100);
        const mb = (downloaded / 1024 / 1024).toFixed(1);
        const totalMb = (totalBytes / 1024 / 1024).toFixed(1);
        spin.message(`Downloading update… ${mb} / ${totalMb} MB (${pct}%)`);
      } else {
        const mb = (downloaded / 1024 / 1024).toFixed(1);
        spin.message(`Downloading update… ${mb} MB`);
      }
    }
    const buffer = Buffer.concat(chunks);
    await Bun.write(tempZip, buffer);

    // Extract using PowerShell
    spin.message("Extracting update...");
    mkdirSync(extractDir, { recursive: true });
    const extract = Bun.spawnSync([
      "pwsh",
      "-Command",
      `Expand-Archive -Path '${tempZip}' -DestinationPath '${extractDir}' -Force`,
    ]);
    if (extract.exitCode !== 0) {
      spin.stop("Extraction failed.");
      p.log.warn("Could not extract the update. Continuing with current version.");
      rmSync(tempDir, { recursive: true, force: true });
      return;
    }

    // Find the new exe inside extracted dir (may be nested in a subfolder)
    spin.message("Installing update...");
    const exeDir = appDir();
    const currentExe = process.execPath;

    // Walk extracted dir to find m365-admin.exe
    const newExe = findFile(extractDir, "m365-admin.exe");

    if (!newExe) {
      spin.stop("Update package missing m365-admin.exe.");
      p.log.warn("Invalid update package. Continuing with current version.");
      rmSync(tempDir, { recursive: true, force: true });
      return;
    }

    // Rename current exe (Windows allows rename of running exe)
    const oldExe = currentExe + ".old";
    renameSync(currentExe, oldExe);

    // Copy new files
    copyFileSync(newExe, join(exeDir, "m365-admin.exe"));

    // Cleanup
    rmSync(tempDir, { recursive: true, force: true });
    // Try to remove the old exe (may fail if still locked)
    try {
      rmSync(oldExe, { force: true });
    } catch {
      // Will be cleaned up on next update
    }

    spin.stop(`Updated to v${latestVersion}.`);
    p.log.success("Update installed. Please restart the app.");
    process.exit(0);
  } catch (e) {
    spin.stop("Update failed.");
    p.log.warn(`Could not complete the update: ${e}`);
    // Try to restore if we renamed the exe
    try {
      const oldExe = process.execPath + ".old";
      const { statSync } = await import("fs");
      statSync(oldExe);
      renameSync(oldExe, process.execPath);
    } catch {
      // Nothing to restore
    }
    rmSync(tempDir, { recursive: true, force: true });
  }
}

/** Recursively find a file by name in a directory. */
function findFile(dir: string, name: string): string | null {
  const { readdirSync, statSync } = require("fs") as typeof import("fs");
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    if (statSync(full).isDirectory()) {
      const found = findFile(full, name);
      if (found) return found;
    } else if (entry.toLowerCase() === name.toLowerCase()) {
      return full;
    }
  }
  return null;
}
