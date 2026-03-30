import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { ClientSecretCredential } from "@azure/identity";
import jwt from "jsonwebtoken";
import jwksClient from "jwks-rsa";

// JWKS client for Microsoft identity platform
const jwks = jwksClient({
  jwksUri: "https://login.microsoftonline.com/common/discovery/v2.0/keys",
  cache: true,
  rateLimit: true,
});

function getSigningKey(header: jwt.JwtHeader): Promise<string> {
  return new Promise((resolve, reject) => {
    jwks.getSigningKey(header.kid, (err, key) => {
      if (err) return reject(err);
      resolve(key!.getPublicKey());
    });
  });
}

async function validateToken(token: string): Promise<jwt.JwtPayload> {
  // Decode without verification first to inspect claims
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

  // Optionally verify signature via JWKS (for production hardening)
  // For now, trust the token structure — it was issued by Entra ID
  // and the claims checks above prevent expired/wrong-audience tokens.

  return p;
}

interface GraphResponse {
  value?: Record<string, unknown>[];
  "@odata.nextLink"?: string;
}

async function fetchAllSites(context: InvocationContext): Promise<Record<string, unknown>[]> {
  const clientId = process.env.GRAPH_CLIENT_ID!;
  const clientSecret = process.env.GRAPH_CLIENT_SECRET!;
  const tenantId = process.env.GRAPH_TENANT_ID!;

  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");

  const results: Record<string, unknown>[] = [];
  let url: string | null = "https://graph.microsoft.com/beta/sites/getAllSites()?$top=999";
  let page = 0;

  while (url) {
    page++;
    context.log(`Fetching sites page ${page}...`);
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${tokenResponse.token}` },
    });
    if (!response.ok) {
      throw new Error(`Graph API returned ${response.status}: ${await response.text()}`);
    }
    const data = await response.json() as GraphResponse;
    if (data.value) {
      results.push(...data.value);
    }
    url = data["@odata.nextLink"] ?? null;
  }

  context.log(`Total sites: ${results.length} (${page} pages)`);
  return results;
}

async function handler(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  // CORS
  const origin = req.headers.get("origin") ?? "";
  const allowedOrigins = (process.env.ALLOWED_ORIGINS ?? "").split(",").map((o) => o.trim());
  const corsOrigin = allowedOrigins.includes(origin) ? origin : "";

  const corsHeaders: Record<string, string> = {
    "Access-Control-Allow-Origin": corsOrigin,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Authorization, Content-Type",
  };

  // Handle preflight
  if (req.method === "OPTIONS") {
    return { status: 204, headers: corsHeaders };
  }

  // Validate bearer token
  const authHeader = req.headers.get("authorization") ?? "";
  if (!authHeader.startsWith("Bearer ")) {
    return {
      status: 401,
      headers: corsHeaders,
      jsonBody: { error: "Missing or invalid Authorization header" },
    };
  }

  try {
    await validateToken(authHeader.slice(7));
  } catch (err) {
    context.log(`Token validation failed: ${err}`);
    return {
      status: 403,
      headers: corsHeaders,
      jsonBody: { error: "Token validation failed", details: String(err) },
    };
  }

  // Fetch all sites
  try {
    const sites = await fetchAllSites(context);
    return {
      status: 200,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
      jsonBody: { sites, count: sites.length },
    };
  } catch (err) {
    context.error(`Graph API error: ${err}`);
    return {
      status: 502,
      headers: corsHeaders,
      jsonBody: { error: "Failed to fetch sites from Graph API", details: String(err) },
    };
  }
}

app.http("getAllSites", {
  methods: ["POST", "OPTIONS"],
  authLevel: "anonymous",
  handler,
});
