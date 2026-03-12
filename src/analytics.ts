import { PostHog } from "posthog-node";
import { mkdirSync, readFileSync, writeFileSync, existsSync } from "fs";
import { join } from "path";
import { homedir } from "os";
import { randomUUID } from "crypto";

const DEFAULT_API_KEY = "phc_OGe4vaw741KAdVtwNADqo8TDgwW1qSEwCiXfncflTap";
const ID_DIR = join(homedir(), ".m365-admin");
const ID_FILE = join(ID_DIR, "analytics-id");

let posthog: PostHog | null = null;
let distinctId: string = "";

function loadOrCreateId(): string {
  try {
    if (existsSync(ID_FILE)) {
      return readFileSync(ID_FILE, "utf-8").trim();
    }
  } catch {}

  const id = randomUUID();
  try {
    mkdirSync(ID_DIR, { recursive: true });
    writeFileSync(ID_FILE, id, "utf-8");
  } catch {}
  return id;
}

export function initAnalytics(): void {
  const apiKey = process.env.POSTHOG_API_KEY ?? DEFAULT_API_KEY;
  distinctId = loadOrCreateId();
  posthog = new PostHog(apiKey, {
    host: "https://us.i.posthog.com",
    flushAt: 1,
    flushInterval: 0,
  });
}

export function trackAppLaunch(version: string): void {
  posthog?.capture({
    distinctId,
    event: "app_launched",
    properties: {
      version,
      os: process.platform,
      $set: { os: process.platform, version },
      $set_once: { first_seen: new Date().toISOString() },
    },
  });
}

export function identifyTenant(domain: string): void {
  posthog?.identify({
    distinctId,
    properties: {
      tenant: domain,
      os: process.platform,
    },
  });
}

export function trackCommandEvent(command: string): void {
  posthog?.capture({
    distinctId,
    event: "command_executed",
    properties: {
      command,
      os: process.platform,
    },
  });
}

export async function shutdownAnalytics(): Promise<void> {
  await posthog?.shutdown();
}
