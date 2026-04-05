import { createClient } from "jsr:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};

const supabaseUrl = Deno.env.get("SUPABASE_URL") || "";
const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "";
const googleClientId = Deno.env.get("GOOGLE_CLIENT_ID") || "";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return json({ ok: true });

  try {
    if (!supabaseUrl || !serviceRoleKey || !googleClientId) {
      throw new Error("Google Calendar connection is missing required secrets.");
    }

    const authorization = req.headers.get("Authorization") || req.headers.get("authorization") || "";
    const accessToken = authorization.replace(/^Bearer\s+/i, "").trim();
    const jwtPayload = decodeJwt(accessToken);
    const userId = String(jwtPayload?.sub || "").trim();
    if (!userId) {
      return json({ error: "Could not identify the signed-in user." }, 401);
    }

    const supabase = createClient(supabaseUrl, serviceRoleKey);
    const body = await req.json().catch(() => ({}));
    const returnTo = String(body.returnTo || "").trim() || "http://localhost:5173";
    const callbackUrl = `${supabaseUrl}/functions/v1/google-calendar-callback`;
    const stateToken = crypto.randomUUID();

    const { error } = await supabase.from("calendar_oauth_states").insert({
      user_id: userId,
      provider: "google",
      state_token: stateToken,
      return_to: returnTo
    });
    if (error) throw error;

    const scope = [
      "openid",
      "email",
      "profile",
      "https://www.googleapis.com/auth/calendar.events"
    ].join(" ");

    const url = new URL("https://accounts.google.com/o/oauth2/v2/auth");
    url.searchParams.set("client_id", googleClientId);
    url.searchParams.set("redirect_uri", callbackUrl);
    url.searchParams.set("response_type", "code");
    url.searchParams.set("access_type", "offline");
    url.searchParams.set("prompt", "consent");
    url.searchParams.set("scope", scope);
    url.searchParams.set("state", stateToken);
    url.searchParams.set("include_granted_scopes", "true");

    return json({ url: url.toString() });
  } catch (error) {
    return json({ error: error instanceof Error ? error.message : "Unexpected Google connect error." }, 500);
  }
});

function decodeJwt(token: string): Record<string, unknown> | null {
  try {
    const [, payload] = token.split(".");
    if (!payload) return null;
    const normalized = payload.replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
    return JSON.parse(atob(padded));
  } catch {
    return null;
  }
}

function json(payload: unknown, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      ...corsHeaders,
      "content-type": "application/json"
    }
  });
}
