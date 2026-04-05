import { createClient } from "jsr:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};

const supabaseUrl = Deno.env.get("SUPABASE_URL") || "";
const supabaseAnonKey = Deno.env.get("SUPABASE_ANON_KEY") || "";
const anthropicApiKey = Deno.env.get("ANTHROPIC_API_KEY") || "";
const anthropicModel = Deno.env.get("ANTHROPIC_MODEL") || "claude-sonnet-4-20250514";

type AssistantRequest = {
  question?: string;
  activeOpportunityId?: string | null;
  history?: Array<{ role: string; content: string }>;
  helpCenter?: unknown;
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    if (!supabaseUrl || !supabaseAnonKey || !anthropicApiKey) {
      throw new Error("Assistant function is missing required environment variables.");
    }

    const authorization = req.headers.get("Authorization") || req.headers.get("authorization") || "";
    if (!authorization) {
      return json({ error: "Missing authorization header." }, 401);
    }
    const accessToken = authorization.replace(/^Bearer\s+/i, "").trim();
    if (!accessToken) {
      return json({ error: "Missing bearer token." }, 401);
    }

    const supabase = createClient(supabaseUrl, supabaseAnonKey, {
      global: {
        headers: {
          Authorization: authorization
        }
      }
    });

    const jwtPayload = decodeJwt(accessToken);
    const userId = String(jwtPayload?.sub || "").trim();
    if (!userId) {
      return json({ error: "Could not read the signed-in user from the access token." }, 401);
    }

    const { data: profile, error: profileError } = await supabase
      .from("profiles")
      .select("id, email, full_name, role, active")
      .eq("id", userId)
      .single();
    if (profileError || !profile) {
      return json({ error: "Could not load the user profile." }, 403);
    }
    if (!profile.active) {
      return json({ error: "This account is inactive." }, 403);
    }

    const body = (await req.json()) as AssistantRequest;
    const question = String(body.question || "").trim();
    if (!question) {
      return json({ error: "Ask a question first." }, 400);
    }

    const [{ data: settings }, { data: opportunities }, { data: activities }] = await Promise.all([
      supabase
        .from("app_settings")
        .select("assumptions, lead_sources, statuses, products, carriers, routing_rules")
        .eq("singleton_key", "default")
        .maybeSingle(),
      supabase
        .from("opportunities")
        .select("id, lead_number, assigned_user_id, assigned_rep_name, date_received, lead_source, business_name, product_focus, target_niche, contact_name, carrier, policy_type, renewal_status, effective_date, expiration_date, next_task, next_follow_up_date, status, premium_quoted, premium_bound, notes, last_activity_date")
        .order("date_received", { ascending: false })
        .limit(profile.role === "admin" ? 250 : 150),
      supabase
        .from("opportunity_activity")
        .select("opportunity_id, actor_name, activity_type, outcome, next_follow_up_date, appointment_at, title, detail, created_at")
        .order("created_at", { ascending: false })
        .limit(profile.role === "admin" ? 300 : 180)
    ]);

    let activeOpportunity = null;
    let activeOpportunityActivities: unknown[] = [];
    if (body.activeOpportunityId) {
      const { data } = await supabase
        .from("opportunities")
        .select("id, lead_number, assigned_user_id, assigned_rep_name, date_received, lead_source, business_name, product_focus, target_niche, contact_name, contact_email, contact_phone, carrier, incumbent_carrier, policy_type, policy_term_months, renewal_status, effective_date, expiration_date, next_task, task_priority, next_follow_up_date, status, premium_quoted, premium_bound, notes, last_activity_date")
        .eq("id", body.activeOpportunityId)
        .maybeSingle();
      activeOpportunity = data || null;

      if (activeOpportunity) {
        const { data: detailActivities } = await supabase
          .from("opportunity_activity")
          .select("actor_name, activity_type, outcome, next_follow_up_date, appointment_at, title, detail, created_at")
          .eq("opportunity_id", body.activeOpportunityId)
          .order("created_at", { ascending: false })
          .limit(20);
        activeOpportunityActivities = detailActivities || [];
      }
    }

    const metrics = summarizeMetrics(opportunities || [], activities || []);
    const promptContext = {
      user: {
        name: profile.full_name,
        role: profile.role,
        email: profile.email
      },
      scope: profile.role === "admin" ? "agency-wide" : "rep-scoped to assigned leads only",
      metrics,
      settings: settings || {},
      helpCenter: body.helpCenter || null,
      opportunities: opportunities || [],
      recentActivities: activities || [],
      activeOpportunity,
      activeOpportunityActivities
    };

    const history = Array.isArray(body.history)
      ? body.history
          .filter((item) => item && typeof item.content === "string" && ["user", "assistant"].includes(item.role))
          .slice(-8)
          .map((item) => ({
            role: item.role as "user" | "assistant",
            content: item.content
          }))
      : [];

    const anthropicResponse = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": anthropicApiKey,
        "anthropic-version": "2023-06-01"
      },
      body: JSON.stringify({
        model: anthropicModel,
        max_tokens: 900,
        temperature: 0.2,
        system:
          "You are Claude inside Golden Leaf Agency HQ. You combine two roles: a calm agency IT/setup expert and an experienced independent agency owner who coaches teams on pipeline management, renewals, reporting, and producer habits. Answer only from the supplied agency data and help-center content. Never reveal data outside the caller's scope. Reps can only receive rep-scoped answers. For setup or workflow questions, use the supplied help-center content first and give practical step-by-step guidance. If the answer is not in the provided data or help center, say that clearly and suggest the nearest available metric, record, or next setup step. Keep answers concise, operational, and supportive.",
        messages: [
          ...history,
          {
            role: "user",
            content: `Agency data context:\n${JSON.stringify(promptContext)}\n\nQuestion:\n${question}`
          }
        ]
      })
    });

    const anthropicPayload = await anthropicResponse.json();
    if (!anthropicResponse.ok) {
      return json(
        { error: anthropicPayload?.error?.message || "Claude request failed." },
        anthropicResponse.status
      );
    }

    const answer = Array.isArray(anthropicPayload.content)
      ? anthropicPayload.content
          .filter((item: { type?: string }) => item?.type === "text")
          .map((item: { text?: string }) => item.text || "")
          .join("\n\n")
      : "";

    return json({ answer, role: profile.role, model: anthropicModel });
  } catch (error) {
    return json(
      { error: error instanceof Error ? error.message : "Unexpected assistant error." },
      500
    );
  }
});

function summarizeMetrics(opportunities: Array<Record<string, unknown>>, activities: Array<Record<string, unknown>>) {
  const totalLeads = opportunities.length;
  const bound = opportunities.filter((item) => item.status === "Bound").length;
  const quoted = opportunities.filter((item) => item.status === "Quoted" || item.status === "Pending Decision").length;
  const overdue = opportunities.filter((item) => {
    const nextFollowUp = String(item.next_follow_up_date || "");
    return nextFollowUp && nextFollowUp < new Date().toISOString().slice(0, 10) && item.status !== "Bound" && item.status !== "Lost";
  }).length;
  const renewalsDueSoon = opportunities.filter((item) => {
    const expiration = String(item.expiration_date || "");
    if (!expiration) return false;
    const diffDays = Math.round((new Date(`${expiration}T00:00:00`).valueOf() - new Date().valueOf()) / 86400000);
    return diffDays >= 0 && diffDays <= 60;
  }).length;
  const touches = activities.filter((item) => item.activity_type && item.activity_type !== "system").length;

  return {
    totalLeads,
    bound,
    quoted,
    overdue,
    renewalsDueSoon,
    touches
  };
}

function decodeJwt(token: string): Record<string, unknown> | null {
  try {
    const [, payload] = token.split(".");
    if (!payload) return null;
    const normalized = payload.replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
    const decoded = atob(padded);
    return JSON.parse(decoded);
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
