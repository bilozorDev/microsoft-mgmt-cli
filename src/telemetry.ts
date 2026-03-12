import * as Sentry from "@sentry/bun";
import pkg from "../package.json";

const DEFAULT_DSN =
  "https://2420ef87f1ec8422d7872f3e5502b5e1@o4511032494653440.ingest.us.sentry.io/4511032496750592";

// --- Sanitization ---

const EMAIL_RE = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
const TENANT_RE = /[a-zA-Z0-9\-]+\.onmicrosoft\.com/g;
const OTS_RE = /https?:\/\/onetimesecret\.com\/secret\/[^\s"')]+/g;
const WIN_PATH_RE = /[A-Z]:\\(?:[^\s\\:*?"<>|]+\\)*[^\s\\:*?"<>|]+/g;
const UNIX_PATH_RE = /\/(?:Users|home|tmp|var|etc)\/[^\s"')]+/g;
const SENSITIVE_KEYS = /password|secret|token|credential/i;

function sanitize(value: string): string {
  return value
    .replace(OTS_RE, "[OTS_URL]")
    .replace(TENANT_RE, "[TENANT].onmicrosoft.com")
    .replace(EMAIL_RE, "[EMAIL]")
    .replace(WIN_PATH_RE, "[PATH]")
    .replace(UNIX_PATH_RE, "[PATH]");
}

function sanitizeEventData(obj: Record<string, unknown>): void {
  for (const key of Object.keys(obj)) {
    if (SENSITIVE_KEYS.test(key)) {
      obj[key] = "[REDACTED]";
    } else if (typeof obj[key] === "string") {
      obj[key] = sanitize(obj[key] as string);
    } else if (obj[key] && typeof obj[key] === "object") {
      sanitizeEventData(obj[key] as Record<string, unknown>);
    }
  }
}

// --- Init ---

export function initTelemetry(): void {
  const dsn = process.env.SENTRY_DSN ?? DEFAULT_DSN;

  Sentry.init({
    dsn,
    release: `m365-admin-cli@${pkg.version}`,
    tracesSampleRate: 1.0,
    beforeSend(event) {
      if (event.message) {
        event.message = sanitize(event.message);
      }
      if (event.exception?.values) {
        for (const ex of event.exception.values) {
          if (ex.value) ex.value = sanitize(ex.value);
        }
      }
      if (event.extra) sanitizeEventData(event.extra as Record<string, unknown>);
      if (event.contexts) sanitizeEventData(event.contexts as Record<string, unknown>);
      return event;
    },
    beforeBreadcrumb(breadcrumb) {
      if (breadcrumb.message) {
        breadcrumb.message = sanitize(breadcrumb.message);
      }
      if (breadcrumb.data) {
        sanitizeEventData(breadcrumb.data as Record<string, unknown>);
      }
      return breadcrumb;
    },
  });

  Sentry.setTag("platform", process.platform);
  Sentry.setTag("runtime", "bun");
}

// --- Command tracking ---

export async function trackCommand<T>(name: string, fn: () => Promise<T>): Promise<T> {
  return Sentry.startSpan({ name: `command:${name}`, op: "command" }, async () => {
    Sentry.addBreadcrumb({
      category: "command",
      message: `Started: ${name}`,
      level: "info",
    });

    const start = performance.now();
    try {
      const result = await fn();
      const duration = Math.round(performance.now() - start);
      Sentry.addBreadcrumb({
        category: "command",
        message: `Completed: ${name}`,
        level: "info",
        data: { durationMs: duration },
      });
      return result;
    } catch (error) {
      Sentry.captureException(error, {
        tags: { command: name },
      });
      throw error;
    }
  });
}

// --- PowerShell error reporting ---

const SILENT_PATTERNS = [
  "-ErrorAction SilentlyContinue",
  "Disconnect-MgGraph",
  "Disconnect-ExchangeOnline",
];

export function reportPowerShellError(command: string, error: string): void {
  if (SILENT_PATTERNS.some((pat) => command.includes(pat))) return;

  Sentry.captureMessage("PowerShell error", {
    level: "warning",
    extra: {
      command: sanitize(command),
      error: sanitize(error),
    },
  });
}

export function reportPowerShellTimeout(command: string, timeoutMs: number): void {
  Sentry.captureMessage("PowerShell command timed out", {
    level: "error",
    extra: {
      command: sanitize(command),
      timeoutMs,
    },
  });
}

// --- Context ---

export function setTenantContext(domain: string): void {
  Sentry.setTag("tenant", sanitize(domain));
}

// --- Shutdown ---

export async function flushTelemetry(): Promise<void> {
  await Sentry.close(2000);
}

// --- Re-exports ---

export const addBreadcrumb = Sentry.addBreadcrumb.bind(Sentry);
export const captureException = Sentry.captureException.bind(Sentry);
