import { createClient } from "jsr:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};

const supabaseUrl = Deno.env.get("SUPABASE_URL") || "";
const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "";
const resendApiKey = Deno.env.get("RESEND_API_KEY") || "";
const reminderFromEmail = Deno.env.get("REMINDER_FROM_EMAIL") || "";
const reminderReplyTo = Deno.env.get("REMINDER_REPLY_TO") || "";
const twilioAccountSid = Deno.env.get("TWILIO_ACCOUNT_SID") || "";
const twilioAuthToken = Deno.env.get("TWILIO_AUTH_TOKEN") || "";
const twilioFromNumber = Deno.env.get("TWILIO_FROM_NUMBER") || "";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return json({ ok: true });

  try {
    if (!supabaseUrl || !serviceRoleKey) {
      throw new Error("Reminder sender is missing required Supabase secrets.");
    }

    const authorization = req.headers.get("Authorization") || req.headers.get("authorization") || "";
    const accessToken = authorization.replace(/^Bearer\s+/i, "").trim();
    const jwtPayload = decodeJwt(accessToken);
    const userId = String(jwtPayload?.sub || "").trim();
    if (!userId) {
      return json({ error: "Could not identify the signed-in user." }, 401);
    }

    const body = await req.json().catch(() => ({}));
    const opportunityId = String(body.opportunityId || "").trim();
    const channel = String(body.channel || "").trim().toLowerCase();
    const subject = String(body.subject || "").trim();
    const messageBody = String(body.body || "").trim();
    const replyToEmail = String(body.replyToEmail || reminderReplyTo || "").trim();
    if (!opportunityId || !["email", "sms"].includes(channel) || !messageBody) {
      return json({ error: "Opportunity, channel, and message body are required." }, 400);
    }

    const supabase = createClient(supabaseUrl, serviceRoleKey);
    const { data: profile, error: profileError } = await supabase
      .from("profiles")
      .select("id, full_name, role, active")
      .eq("id", userId)
      .single();
    if (profileError || !profile || !profile.active) {
      return json({ error: "Could not validate the signed-in user." }, 401);
    }

    const { data: opportunity, error: opportunityError } = await supabase
      .from("opportunities")
      .select("id, assigned_user_id, contact_email, contact_phone, business_name")
      .eq("id", opportunityId)
      .single();
    if (opportunityError || !opportunity) {
      return json({ error: "That lead could not be found." }, 404);
    }
    if (profile.role !== "admin" && opportunity.assigned_user_id !== userId) {
      return json({ error: "You do not have access to message this lead." }, 403);
    }

    if (channel === "email") {
      if (!resendApiKey || !reminderFromEmail) {
        return json({ error: "Email provider is not configured yet." }, 400);
      }
      if (!opportunity.contact_email) {
        return json({ error: "This lead does not have a contact email." }, 400);
      }

      const response = await fetch("https://api.resend.com/emails", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${resendApiKey}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          from: reminderFromEmail,
          to: [opportunity.contact_email],
          subject: subject || `${opportunity.business_name} follow-up`,
          text: messageBody,
          ...(replyToEmail ? { reply_to: [replyToEmail] } : {})
        })
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        return json({ error: payload?.message || "Resend could not send the email." }, response.status);
      }
      return json({ ok: true, providerId: payload.id || "" });
    }

    if (!twilioAccountSid || !twilioAuthToken || !twilioFromNumber) {
      return json({ error: "SMS provider is not configured yet." }, 400);
    }
    if (!opportunity.contact_phone) {
      return json({ error: "This lead does not have a contact phone number." }, 400);
    }

    const twilioAuth = btoa(`${twilioAccountSid}:${twilioAuthToken}`);
    const smsResponse = await fetch(`https://api.twilio.com/2010-04-01/Accounts/${twilioAccountSid}/Messages.json`, {
      method: "POST",
      headers: {
        Authorization: `Basic ${twilioAuth}`,
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: new URLSearchParams({
        To: opportunity.contact_phone,
        From: twilioFromNumber,
        Body: messageBody
      })
    });
    const smsPayload = await smsResponse.json().catch(() => ({}));
    if (!smsResponse.ok) {
      return json({ error: smsPayload?.message || "Twilio could not send the text." }, smsResponse.status);
    }

    return json({ ok: true, providerId: smsPayload.sid || "" });
  } catch (error) {
    return json({ error: error instanceof Error ? error.message : "Unexpected reminder error." }, 500);
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
