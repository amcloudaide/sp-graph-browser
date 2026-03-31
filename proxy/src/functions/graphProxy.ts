import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { ClientSecretCredential } from "@azure/identity";
import jwt from "jsonwebtoken";

// Whitelisted Graph API patterns that the proxy can call with app-only auth.
// Each pattern is a regex tested against the requested path.
const ALLOWED_PATTERNS: RegExp[] = [
  /^\/sites\/getAllSites\(\)/,                              // List all sites
  /^\/sites\/[^/]+\/permissions$/,                          // Site permissions
  /^\/sites\/[^/]+\/drive\/items\/[^/]+\/permissions$/,     // Drive item permissions (sharing links)
  /^\/sites\/[^/]+\/drives\/[^/]+\/items\/[^/]+\/permissions$/, // Drive permissions (alternate path)
];

function isPathAllowed(path: string): boolean {
  return ALLOWED_PATTERNS.some((pattern) => pattern.test(path));
}

async function validateToken(token: string): Promise<jwt.JwtPayload> {
  const decoded = jwt.decode(token, { complete: true });
  if (!decoded || typeof decoded === "string" || !decoded.payload || typeof decoded.payload === "string") {
    throw new Error("Invalid token format");
  }

  const p = decoded.payload as jwt.JwtPayload;

  // Check expiry
  const now = Math.floor(Date.now() / 1000);
  if (p.exp && p.exp < now) {
    throw new Error("Token expired");
  }

  // Verify issuer is Microsoft identity platform
  const iss = p.iss ?? "";
  if (!iss.startsWith("https://login.microsoftonline.com/") &&
      !iss.startsWith("https://sts.windows.net/")) {
    throw new Error(`Invalid token issuer: ${iss}`);
  }

  // Verify audience is Graph API
  const aud = p.aud ?? "";
  if (aud !== "https://graph.microsoft.com" &&
      aud !== "00000003-0000-0000-c000-000000000000") {
    throw new Error(`Invalid token audience: ${aud}`);
  }

  // Check allowed tenants if configured
  const allowedTenants = process.env.ALLOWED_TENANT_IDS;
  if (allowedTenants) {
    const tenants = allowedTenants.split(",").map((t) => t.trim());
    if (!tenants.includes(p.tid as string)) {
      throw new Error(`Tenant ${p.tid} not allowed`);
    }
  }

  return p;
}

interface GraphResponse {
  value?: Record<string, unknown>[];
  "@odata.nextLink"?: string;
  [key: string]: unknown;
}

async function callGraph(
  path: string,
  apiVersion: string,
  context: InvocationContext
): Promise<{ data: unknown; isCollection: boolean }> {
  const clientId = process.env.GRAPH_CLIENT_ID!;
  const clientSecret = process.env.GRAPH_CLIENT_SECRET!;
  const tenantId = process.env.GRAPH_TENANT_ID!;

  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");

  const baseUrl = `https://graph.microsoft.com/${apiVersion}`;
  let url: string | null = `${baseUrl}${path}`;
  const allResults: Record<string, unknown>[] = [];
  let page = 0;
  let singleResponse: GraphResponse | null = null;

  while (url) {
    page++;
    context.log(`Graph proxy: ${url} (page ${page})`);
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${tokenResponse.token}` },
    });
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Graph API ${response.status}: ${errorText}`);
    }
    const data = await response.json() as GraphResponse;

    if (data.value) {
      // Collection response — paginate
      allResults.push(...data.value);
      url = data["@odata.nextLink"] ?? null;
    } else {
      // Single entity response
      singleResponse = data;
      url = null;
    }
  }

  if (singleResponse && allResults.length === 0) {
    return { data: singleResponse, isCollection: false };
  }

  context.log(`Graph proxy: ${allResults.length} results (${page} pages)`);
  return { data: allResults, isCollection: true };
}

function getCorsHeaders(req: HttpRequest): Record<string, string> {
  const origin = req.headers.get("origin") ?? "";
  const allowedOrigins = (process.env.ALLOWED_ORIGINS ?? "").split(",").map((o) => o.trim());
  const corsOrigin = allowedOrigins.includes(origin) ? origin : "";
  return {
    "Access-Control-Allow-Origin": corsOrigin,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Authorization, Content-Type",
  };
}

async function handler(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  const corsHeaders = getCorsHeaders(req);

  if (req.method === "OPTIONS") {
    return { status: 204, headers: corsHeaders };
  }

  // Validate bearer token
  const authHeader = req.headers.get("authorization") ?? "";
  if (!authHeader.startsWith("Bearer ")) {
    return { status: 401, headers: corsHeaders, jsonBody: { error: "Missing Authorization header" } };
  }

  try {
    await validateToken(authHeader.slice(7));
  } catch (err) {
    context.log(`Token validation failed: ${err}`);
    return { status: 403, headers: corsHeaders, jsonBody: { error: "Token validation failed", details: String(err) } };
  }

  // Parse request body
  let body: { path?: string; apiVersion?: string };
  try {
    body = await req.json() as { path?: string; apiVersion?: string };
  } catch {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Invalid JSON body. Expected: { path: '/sites/...' }" } };
  }

  const graphPath = body.path;
  const apiVersion = body.apiVersion ?? "beta";

  if (!graphPath || typeof graphPath !== "string") {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Missing 'path' in request body" } };
  }

  // Security: only allow whitelisted paths
  if (!isPathAllowed(graphPath)) {
    return { status: 403, headers: corsHeaders, jsonBody: { error: `Path not allowed: ${graphPath}` } };
  }

  try {
    const result = await callGraph(graphPath, apiVersion, context);
    return {
      status: 200,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
      jsonBody: result,
    };
  } catch (err) {
    context.error(`Graph proxy error: ${err}`);
    return { status: 502, headers: corsHeaders, jsonBody: { error: "Graph API call failed", details: String(err) } };
  }
}

app.http("graphProxy", {
  methods: ["POST", "OPTIONS"],
  authLevel: "anonymous",
  handler,
});
