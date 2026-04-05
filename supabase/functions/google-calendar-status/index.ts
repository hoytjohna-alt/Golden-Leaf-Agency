import { createClient } from "jsr:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};

const supabaseUrl = Deno.env.get("SUPABASE_URL") || "";
const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return json({ ok: true });

  try {
    if (!supabaseUrl || !serviceRoleKey) {
      throw new Error("Calendar status is missing required secrets.");
    }

    const authorization = req.headers.get("Authorization") || req.headers.get("authorization") || "";
    const accessToken = authorization.replace(/^Bearer\s+/i, "").trim();
    const jwtPayload = decodeJwt(accessToken);
    const userId = String(jwtPayload?.sub || "").trim();
    if (!userId) {
      return json({ error: "Could not identify the signed-in user." }, 401);
    }

    const supabase = createClient(supabaseUrl, serviceRoleKey);
    const { data, error } = await supabase
      .from("user_calendar_connections")
      .select("provider, provider_email, expires_at")
      .eq("user_id", userId)
      .eq("provider", "google")
      .maybeSingle();
    if (error) throw error;

    return json({
      connection: data
        ? {
            connected: true,
            provider: data.provider,
            email: data.provider_email,
            expiresAt: data.expires_at
          }
        : {
            connected: false
          }
    });
  } catch (error) {
    return json({ error: error instanceof Error ? error.message : "Unexpected calendar status error." }, 500);
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
