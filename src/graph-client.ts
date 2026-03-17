import type { PowerShellSession } from "./powershell.ts";
import { addBreadcrumb } from "./telemetry.ts";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const TOKEN_TTL_MS = 50 * 60 * 1000; // 50 minutes (tokens last ~60-75 min)

interface GraphRequestOptions {
  params?: Record<string, string>;
  headers?: Record<string, string>;
  method?: "GET" | "POST" | "PATCH" | "DELETE";
  body?: unknown;
}

interface GraphPageResponse<T> {
  "@odata.context"?: string;
  "@odata.count"?: number;
  "@odata.nextLink"?: string;
  value: T[];
}

export class GraphApiError extends Error {
  constructor(
    public status: number,
    public body: string,
    public path: string,
  ) {
    let detail = "";
    try {
      const parsed = JSON.parse(body);
      detail = parsed?.error?.message ?? body;
    } catch {
      detail = body;
    }
    super(`Graph API ${status} on ${path}: ${detail}`);
    this.name = "GraphApiError";
  }
}

export class GraphClient {
  private token: string | null = null;
  private tokenExpiresAt = 0;

  constructor(private ps: PowerShellSession) {}

  private async ensureToken(): Promise<string> {
    if (this.token && Date.now() < this.tokenExpiresAt) {
      return this.token;
    }
    this.token = await this.ps.getGraphAccessToken();
    this.tokenExpiresAt = Date.now() + TOKEN_TTL_MS;
    return this.token;
  }

  private invalidateToken(): void {
    this.token = null;
    this.tokenExpiresAt = 0;
  }

  async request<T>(
    path: string,
    options: GraphRequestOptions = {},
  ): Promise<T> {
    const { params, headers: extraHeaders, method = "GET", body } = options;

    const url = new URL(
      path.startsWith("http") ? path : `${GRAPH_BASE}${path}`,
    );
    if (params) {
      for (const [key, value] of Object.entries(params)) {
        url.searchParams.set(key, value);
      }
    }

    const doRequest = async (token: string): Promise<Response> => {
      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...extraHeaders,
      };
      return fetch(url.toString(), {
        method,
        headers,
        body: body ? JSON.stringify(body) : undefined,
      });
    };

    let token = await this.ensureToken();
    let res = await doRequest(token);

    // Retry once on 401 (token may have expired)
    if (res.status === 401) {
      this.invalidateToken();
      token = await this.ensureToken();
      res = await doRequest(token);
    }

    // Handle throttling
    if (res.status === 429) {
      const retryAfter = parseInt(
        res.headers.get("Retry-After") ?? "5",
        10,
      );
      await Bun.sleep((Number.isFinite(retryAfter) ? retryAfter : 5) * 1000);
      // Refresh token in case it expired during the wait
      token = await this.ensureToken();
      res = await doRequest(token);
      // Handle 401 after throttle retry (token may have expired during wait)
      if (res.status === 401) {
        this.invalidateToken();
        token = await this.ensureToken();
        res = await doRequest(token);
      }
    }

    if (!res.ok) {
      const errBody = await res.text();
      throw new GraphApiError(res.status, errBody, url.pathname);
    }

    // 204 No Content (e.g., DELETE or POST member ref)
    if (res.status === 204) return undefined as T;

    return (await res.json()) as T;
  }

  async getAll<T>(
    path: string,
    options: Omit<GraphRequestOptions, "method" | "body"> = {},
  ): Promise<T[]> {
    const results: T[] = [];
    const { params, headers: extraHeaders } = options;

    // Build initial URL with params
    const initialUrl = new URL(
      path.startsWith("http") ? path : `${GRAPH_BASE}${path}`,
    );
    if (params) {
      for (const [key, value] of Object.entries(params)) {
        initialUrl.searchParams.set(key, value);
      }
    }

    let nextUrl: string | null = initialUrl.toString();

    while (nextUrl) {
      const page: GraphPageResponse<T> = await this.request(nextUrl, {
        headers: extraHeaders,
      });
      if (page.value) {
        results.push(...page.value);
      }
      nextUrl = page["@odata.nextLink"] ?? null;
    }

    addBreadcrumb({
      category: "graph",
      message: `getAll ${path}`,
      data: { count: results.length },
    });

    return results;
  }

  async post<T>(
    path: string,
    body: unknown,
    headers?: Record<string, string>,
  ): Promise<T> {
    return this.request<T>(path, { method: "POST", body, headers });
  }
}
