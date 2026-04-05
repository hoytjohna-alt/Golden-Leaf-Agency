import { createClient } from "jsr:@supabase/supabase-js@2";

const supabaseUrl = Deno.env.get("SUPABASE_URL") || "";
const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "";
const googleClientId = Deno.env.get("GOOGLE_CLIENT_ID") || "";
const googleClientSecret = Deno.env.get("GOOGLE_CLIENT_SECRET") || "";
const appUrl = Deno.env.get("APP_URL") || "http://localhost:5173";

Deno.serve(async (req) => {
  try {
    if (!supabaseUrl || !serviceRoleKey || !googleClientId || !googleClientSecret) {
      return redirectWithError("Google Calendar callback is missing required secrets.");
    }

    const url = new URL(req.url);
    const code = url.searchParams.get("code") || "";
    const stateToken = url.searchParams.get("state") || "";
    if (!code || !stateToken) {
      return redirectWithError("Google did not return the required OAuth data.");
    }

    const supabase = createClient(supabaseUrl, serviceRoleKey);
    const { data: stateRow, error: stateError } = await supabase
      .from("calendar_oauth_states")
      .select("id, user_id, return_to")
      .eq("state_token", stateToken)
      .eq("provider", "google")
      .maybeSingle();
    if (stateError || !stateRow) {
      return redirectWithError("This Google connection request expired. Start it again from the app.");
    }

    const callbackUrl = `${supabaseUrl}/functions/v1/google-calendar-callback`;
    const tokenResponse = await fetch("https://oauth2.googleapis.com/token", {
      method: "POST",
      headers: {
        "content-type": "application/x-www-form-urlencoded"
      },
      body: new URLSearchParams({
        code,
        client_id: googleClientId,
        client_secret: googleClientSecret,
        redirect_uri: callbackUrl,
        grant_type: "authorization_code"
      })
    });
    const tokenPayload = await tokenResponse.json();
    if (!tokenResponse.ok) {
      return redirectWithError(tokenPayload?.error_description || "Google token exchange failed.", stateRow.return_to);
    }

    const userInfoResponse = await fetch("https://openidconnect.googleapis.com/v1/userinfo", {
      headers: {
        Authorization: `Bearer ${tokenPayload.access_token}`
      }
    });
    const userInfo = await userInfoResponse.json().catch(() => ({}));

    const expiresAt = tokenPayload.expires_in
      ? new Date(Date.now() + Number(tokenPayload.expires_in) * 1000).toISOString()
      : null;

    const { error: upsertError } = await supabase.from("user_calendar_connections").upsert(
      {
        user_id: stateRow.user_id,
        provider: "google",
        provider_email: String(userInfo.email || ""),
        access_token: String(tokenPayload.access_token || ""),
        refresh_token: String(tokenPayload.refresh_token || ""),
        token_scope: String(tokenPayload.scope || ""),
        expires_at: expiresAt
      },
      { onConflict: "user_id,provider" }
    );
    if (upsertError) {
      return redirectWithError(upsertError.message || "Could not save the Google connection.", stateRow.return_to);
    }

    await supabase.from("calendar_oauth_states").delete().eq("id", stateRow.id);
    return Response.redirect(`${stateRow.return_to}?calendar=google-connected`, 302);
  } catch (error) {
    return redirectWithError(error instanceof Error ? error.message : "Unexpected Google callback error.");
  }
});

function redirectWithError(message: string, target = appUrl) {
  const safeMessage = encodeURIComponent(message);
  return Response.redirect(`${target}?calendar_error=${safeMessage}`, 302);
}
