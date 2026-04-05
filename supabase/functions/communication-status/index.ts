const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};

const resendApiKey = Deno.env.get("RESEND_API_KEY") || "";
const reminderFromEmail = Deno.env.get("REMINDER_FROM_EMAIL") || "";
const twilioAccountSid = Deno.env.get("TWILIO_ACCOUNT_SID") || "";
const twilioAuthToken = Deno.env.get("TWILIO_AUTH_TOKEN") || "";
const twilioFromNumber = Deno.env.get("TWILIO_FROM_NUMBER") || "";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return json({ ok: true });
  }

  return json({
    email: {
      configured: Boolean(resendApiKey && reminderFromEmail)
    },
    sms: {
      configured: Boolean(twilioAccountSid && twilioAuthToken && twilioFromNumber)
    }
  });
});

function json(payload: unknown, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      ...corsHeaders,
      "content-type": "application/json"
    }
  });
}
