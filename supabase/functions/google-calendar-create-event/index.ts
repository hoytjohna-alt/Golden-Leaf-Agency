import { createClient } from "jsr:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};

const supabaseUrl = Deno.env.get("SUPABASE_URL") || "";
const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "";
const googleClientId = Deno.env.get("GOOGLE_CLIENT_ID") || "";
const googleClientSecret = Deno.env.get("GOOGLE_CLIENT_SECRET") || "";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return json({ ok: true });

  try {
    if (!supabaseUrl || !serviceRoleKey || !googleClientId || !googleClientSecret) {
      throw new Error("Google Calendar event sync is missing required secrets.");
    }

    const authorization = req.headers.get("Authorization") || req.headers.get("authorization") || "";
    const accessToken = authorization.replace(/^Bearer\s+/i, "").trim();
    const jwtPayload = decodeJwt(accessToken);
    const userId = String(jwtPayload?.sub || "").trim();
    if (!userId) {
      return json({ error: "Could not identify the signed-in user." }, 401);
    }

    const body = await req.json().catch(() => ({}));
    const summary = String(body.summary || "").trim();
    const description = String(body.description || "").trim();
    const startAt = String(body.startAt || "").trim();
    const endAt = String(body.endAt || "").trim();
    const opportunityId = String(body.opportunityId || "").trim();
    if (!summary || !startAt || !endAt) {
      return json({ error: "Summary, start, and end times are required." }, 400);
    }

    const supabase = createClient(supabaseUrl, serviceRoleKey);
    const { data: connection, error: connectionError } = await supabase
      .from("user_calendar_connections")
      .select("*")
      .eq("user_id", userId)
      .eq("provider", "google")
      .maybeSingle();
    if (connectionError || !connection) {
      return json({ error: "Connect Google Calendar first." }, 400);
    }

    const validAccessToken = await getValidGoogleToken(connection, supabase);
    const eventResponse = await fetch("https://www.googleapis.com/calendar/v3/calendars/primary/events", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        Authorization: `Bearer ${validAccessToken}`
      },
      body: JSON.stringify({
        summary,
        description,
        start: { dateTime: new Date(startAt).toISOString() },
        end: { dateTime: new Date(endAt).toISOString() }
      })
    });
    const eventPayload = await eventResponse.json();
    if (!eventResponse.ok) {
      return json({ error: eventPayload?.error?.message || "Google Calendar event creation failed." }, eventResponse.status);
    }

    if (opportunityId) {
      await supabase.from("opportunity_activity").insert({
        opportunity_id: opportunityId,
        actor_id: userId,
        actor_name: "Calendar Sync",
        activity_type: "Appointment",
        outcome: "Appointment Set",
        appointment_at: new Date(startAt).toISOString(),
        title: "Calendar event created",
        detail: `${summary} was added to Google Calendar.${eventPayload.htmlLink ? ` ${eventPayload.htmlLink}` : ""}`
      });
    }

    return json({ id: eventPayload.id, htmlLink: eventPayload.htmlLink || "" });
  } catch (error) {
    return json({ error: error instanceof Error ? error.message : "Unexpected Google event error." }, 500);
  }
});

async function getValidGoogleToken(connection: Record<string, string>, supabase: ReturnType<typeof createClient>) {
  const expiresAt = connection.expires_at ? new Date(connection.expires_at).valueOf() : 0;
  if (connection.access_token && expiresAt > Date.now() + 60_000) {
    return connection.access_token;
  }
  if (!connection.refresh_token) {
    throw new Error("Reconnect Google Calendar so the refresh token can be stored.");
  }

  const refreshResponse = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: {
      "content-type": "application/x-www-form-urlencoded"
    },
    body: new URLSearchParams({
      client_id: googleClientId,
      client_secret: googleClientSecret,
      refresh_token: connection.refresh_token,
      grant_type: "refresh_token"
    })
  });
  const refreshPayload = await refreshResponse.json();
  if (!refreshResponse.ok) {
    throw new Error(refreshPayload?.error_description || "Google token refresh failed.");
  }

  const newExpiresAt = refreshPayload.expires_in
    ? new Date(Date.now() + Number(refreshPayload.expires_in) * 1000).toISOString()
    : connection.expires_at;

  await supabase
    .from("user_calendar_connections")
    .update({
      access_token: String(refreshPayload.access_token || connection.access_token),
      expires_at: newExpiresAt,
      token_scope: String(refreshPayload.scope || connection.token_scope || "")
    })
    .eq("id", connection.id);

  return String(refreshPayload.access_token || connection.access_token);
}

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
