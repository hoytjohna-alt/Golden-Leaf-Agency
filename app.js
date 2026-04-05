import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

const APP_CONFIG = {
  supabaseUrl: import.meta.env.VITE_SUPABASE_URL || "",
  supabaseAnonKey: import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY || "",
  appUrl: import.meta.env.VITE_APP_URL || ""
};
const SUPABASE_READY = Boolean(APP_CONFIG.supabaseUrl && APP_CONFIG.supabaseAnonKey);
const APP_BUILD_ID = typeof __APP_BUILD_ID__ !== "undefined" ? __APP_BUILD_ID__ : "v1";
const APP_BUILD_STORAGE_KEY = "golden-leaf-app-build-id";
const APP_RECOVERY_STORAGE_KEY = "golden-leaf-app-recovery-build-id";
const AUTH_STORAGE_PREFIX = "golden-leaf-auth";
const AUTH_STORAGE_KEY = AUTH_STORAGE_PREFIX;
const VERSION_CHECK_PATH = "/version.json";
const WORKSPACE_TIMEOUT_MS = 10000;
const VERSION_CHECK_INTERVAL_MS = 45000;
const ATTACHMENTS_BUCKET = "opportunity-files";
const ASSISTANT_FUNCTION_NAME = "claude-assistant";
const GOOGLE_CALENDAR_STATUS_FUNCTION = "google-calendar-status";
const GOOGLE_CALENDAR_CONNECT_FUNCTION = "google-calendar-connect";
const GOOGLE_CALENDAR_DISCONNECT_FUNCTION = "google-calendar-disconnect";
const GOOGLE_CALENDAR_CREATE_EVENT_FUNCTION = "google-calendar-create-event";
const COMMUNICATION_STATUS_FUNCTION = "communication-status";
const SEND_REMINDER_FUNCTION = "send-reminder";

const seedSettings = {
  assumptions: {
    averageCommissionPct: 0.12,
    sameDayWorkedTargetPct: 0.9,
    contactRateTargetPct: 0.5,
    quoteRateTargetPct: 0.2,
    quoteToBindTargetPct: 0.25,
    leadToBindTargetPct: 0.05,
    crmComplianceTargetPct: 0.95,
    followUpDueWindowDays: 3,
    freshLeadWindowDays: 3
  },
  routingRules: {
    autoAssignEnabled: false,
    mode: "round_robin",
    roundRobinCursor: 0,
    sourceRules: []
  },
  communicationSettings: {
    emailSubjectTemplate: "{{businessName}} follow-up from Golden Leaf Agency",
    emailBodyTemplate: "Hi {{contactName}},\n\nThis is {{repName}} from Golden Leaf Agency following up on {{businessName}}. {{nextTaskSentence}}\n\nYou can reply here if you have questions.\n\nThanks,\n{{repName}}",
    smsBodyTemplate: "Hi {{contactName}}, this is {{repName}} from Golden Leaf Agency following up on {{businessName}}. {{nextTaskSentence}}",
    replyToEmail: ""
  },
  leadSources: [
    "Purchased Leads",
    "Warm Transfer",
    "Referral",
    "Website / Organic",
    "Partner / Network",
    "Recycled Lead",
    "Self-Generated"
  ],
  statuses: [
    "New Lead",
    "Attempted",
    "Contacted",
    "Qualified",
    "Quoted",
    "Pending Decision",
    "Bound",
    "Lost",
    "Nurture / Recycle"
  ],
  products: ["GL / BOP", "Workers Comp", "Package / Multi-Line"],
  carriers: [
    { name: "AmTrust", newPct: 0.16, renewalPct: 0.1, notes: "Typical" },
    { name: "Berxi", newPct: 0.13, renewalPct: 0.13, notes: "Typical" },
    { name: "Blitz", newPct: 0.125, renewalPct: 0.125, notes: "Typical" },
    { name: "Chubb", newPct: 0.14, renewalPct: 0.12, notes: "Typical" },
    { name: "Coterie", newPct: 0.12, renewalPct: 0.1, notes: "Typical" },
    { name: "First", newPct: 0.16, renewalPct: 0.16, notes: "Typical" },
    { name: "Hiscox", newPct: 0.14, renewalPct: 0.12, notes: "Typical" },
    { name: "Pathpoint", newPct: 0.11, renewalPct: 0.11, notes: "Typical" },
    { name: "Simply Business", newPct: 0.12, renewalPct: 0.12, notes: "Typical" },
    { name: "THREE", newPct: 0.12, renewalPct: 0.12, notes: "Typical" }
  ]
};

const state = {
  supabase: null,
  session: null,
  profile: null,
  setup: structuredClone(seedSettings),
  profiles: [],
  opportunities: [],
  opportunityActivities: [],
  opportunityAttachments: [],
  coachingNotes: [],
  calendarConnection: null,
  communicationStatus: null,
  ui: {
    loading: true,
    authLoading: false,
    recoveringSession: false,
    error: "",
    notice: "",
    timeframe: "all",
    search: "",
    repFilter: "All",
    sourceFilter: "All",
    statusFilter: "All",
    dateFrom: "",
    dateTo: "",
    activeOpportunityId: null,
    activeTab: "dashboard",
    bulkAssignUserId: "",
    selectedOpportunityIds: [],
    uploadProgress: 0,
    uploadingAttachmentName: "",
    uploadingAttachmentOpportunityId: "",
    opportunityView: "board",
    opportunityTab: "stage",
    setupTab: "users",
    carrierEditing: false,
    assumptionEditing: false,
    importPreviewRows: [],
    importFileName: "",
    importingCsv: false
    ,
    assistantOpen: false,
    assistantLoading: false,
    assistantError: "",
    assistantMessages: [],
    assistantInput: "",
    calendarSyncLoading: false,
    reminderSending: false,
    reminderEditing: false
  }
};

let versionCheckTimer = null;
let dragAutoScrollVelocity = 0;
let dragAutoScrollFrame = null;
let workspaceLoadPromise = null;

const appEl = document.getElementById("app");
const heroActionsEl = document.getElementById("heroActions");
const topNavEl = document.getElementById("topNav");
const assistantRootEl = document.getElementById("assistantRoot");

init();

async function init() {
  handleAppQueryState();
  handleBuildVersionChange();
  setupVersionWatchers();

  if (!SUPABASE_READY) {
    state.ui.loading = false;
    render();
    return;
  }

  state.supabase = createClient(APP_CONFIG.supabaseUrl, APP_CONFIG.supabaseAnonKey, {
    auth: {
      storageKey: AUTH_STORAGE_KEY,
      persistSession: true,
      autoRefreshToken: true,
      detectSessionInUrl: true
    }
  });

  state.supabase.auth.onAuthStateChange((event, session) => {
    state.session = session;
    if (session?.user) {
      const sameUser = state.profile?.id === session.user.id;
      if ((event === "SIGNED_IN" || event === "INITIAL_SESSION" || event === "USER_UPDATED") && !sameUser) {
        window.setTimeout(() => {
          state.profile = null;
          void loadWorkspace({ showLoading: true });
        }, 0);
      }
    } else {
      state.profile = null;
      state.ui.loading = false;
      render();
    }
  });

  const {
    data: { session }
  } = await state.supabase.auth.getSession();
  state.session = session;
  if (session?.user) {
    await loadWorkspace({ showLoading: true });
  } else {
    state.ui.loading = false;
    render();
  }
}

function handleAppQueryState() {
  const url = new URL(window.location.href);
  const connected = url.searchParams.get("calendar");
  const calendarError = url.searchParams.get("calendar_error");
  if (connected === "google-connected") {
    state.ui.notice = "Google Calendar connected successfully.";
  }
  if (calendarError) {
    state.ui.error = calendarError;
  }
  if (connected || calendarError) {
    url.searchParams.delete("calendar");
    url.searchParams.delete("calendar_error");
    window.history.replaceState({}, "", url.toString());
  }
}

async function loadWorkspace({ showLoading = true } = {}) {
  if (workspaceLoadPromise) {
    return workspaceLoadPromise;
  }

  workspaceLoadPromise = (async () => {
    try {
      if (showLoading) {
        state.ui.loading = true;
        state.ui.error = "";
        state.ui.notice = "";
        render();
      }

      const profile = await withTimeout(fetchProfile(state.session.user.id), WORKSPACE_TIMEOUT_MS);
      state.profile = profile;
      if (!isAdmin() && state.ui.activeTab === "dashboard") {
        state.ui.activeTab = "opportunities";
      }
      ensureActiveTab();

      if (!profile.active) {
        state.profiles = [profile];
        state.opportunities = [];
        state.coachingNotes = [];
        state.opportunityActivities = [];
        state.opportunityAttachments = [];
        state.ui.loading = false;
        render();
        return;
      }

      const [settings, opportunities, coachingNotes, profiles, opportunityActivities, opportunityAttachments, calendarConnection, communicationStatus] = await withTimeout(Promise.all([
        fetchSettings(),
        fetchOpportunities(),
        fetchCoachingNotes(),
        isAdmin() ? fetchProfiles() : Promise.resolve([profile]),
        fetchOpportunityActivities(),
        fetchOpportunityAttachments(),
        fetchCalendarConnectionStatus(),
        fetchCommunicationStatus()
      ]), WORKSPACE_TIMEOUT_MS);

      state.setup = settings;
      state.opportunities = opportunities;
      state.coachingNotes = coachingNotes;
      state.profiles = profiles;
      state.opportunityActivities = opportunityActivities;
      state.opportunityAttachments = opportunityAttachments;
      state.calendarConnection = calendarConnection;
      state.communicationStatus = communicationStatus;
      state.ui.recoveringSession = false;
      state.ui.selectedOpportunityIds = [];
      state.ui.loading = false;
      render();
    } catch (error) {
      console.error(error);
      const recovered = await attemptAutomaticRecovery(error);
      if (recovered) {
        return;
      }
      state.ui.loading = false;
      state.ui.error = error.message || "Could not load the workspace.";
      state.ui.recoveringSession = true;
      render();
    } finally {
      workspaceLoadPromise = null;
    }
  })();

  return workspaceLoadPromise;
}

function handleBuildVersionChange() {
  const previousBuildId = localStorage.getItem(APP_BUILD_STORAGE_KEY);
  if (previousBuildId !== APP_BUILD_ID) {
    clearStaleBuildSessions();
    localStorage.setItem(APP_BUILD_STORAGE_KEY, APP_BUILD_ID);
    resetTransientUiState();
  }
}

function clearStaleBuildSessions() {
  Object.keys(localStorage)
    .filter((key) => key.startsWith(`${AUTH_STORAGE_PREFIX}-`))
    .forEach((key) => {
      localStorage.removeItem(key);
    });
}

function resetTransientUiState() {
  state.ui.search = "";
  state.ui.repFilter = "All";
  state.ui.sourceFilter = "All";
  state.ui.statusFilter = "All";
  state.ui.dateFrom = "";
  state.ui.dateTo = "";
  state.ui.activeOpportunityId = null;
  state.ui.bulkAssignUserId = "";
  state.ui.selectedOpportunityIds = [];
}

function stopDragAutoScroll() {
  dragAutoScrollVelocity = 0;
  if (dragAutoScrollFrame) {
    window.cancelAnimationFrame(dragAutoScrollFrame);
    dragAutoScrollFrame = null;
  }
}

function runDragAutoScroll() {
  if (!dragAutoScrollVelocity) {
    dragAutoScrollFrame = null;
    return;
  }
  window.scrollBy(0, dragAutoScrollVelocity);
  dragAutoScrollFrame = window.requestAnimationFrame(runDragAutoScroll);
}

function updateDragAutoScroll(pointerY) {
  const edgeThreshold = 120;
  const maxVelocity = 18;
  const viewportHeight = window.innerHeight;
  let nextVelocity = 0;

  if (pointerY < edgeThreshold) {
    nextVelocity = -Math.max(6, ((edgeThreshold - pointerY) / edgeThreshold) * maxVelocity);
  } else if (pointerY > viewportHeight - edgeThreshold) {
    nextVelocity = Math.max(6, ((pointerY - (viewportHeight - edgeThreshold)) / edgeThreshold) * maxVelocity);
  }

  dragAutoScrollVelocity = nextVelocity;
  if (dragAutoScrollVelocity && !dragAutoScrollFrame) {
    dragAutoScrollFrame = window.requestAnimationFrame(runDragAutoScroll);
  }
  if (!dragAutoScrollVelocity) {
    stopDragAutoScroll();
  }
}

function setupVersionWatchers() {
  if (versionCheckTimer) {
    window.clearInterval(versionCheckTimer);
  }
  versionCheckTimer = window.setInterval(() => {
    if (document.visibilityState === "visible") {
      void checkForNewBuild();
    }
  }, VERSION_CHECK_INTERVAL_MS);
}

async function checkForNewBuild() {
  try {
    const response = await fetch(`${VERSION_CHECK_PATH}?t=${Date.now()}`, {
      cache: "no-store"
    });
    if (!response.ok) {
      return;
    }
    const data = await response.json();
    const remoteBuildId = data?.buildId;
    if (remoteBuildId && remoteBuildId !== APP_BUILD_ID) {
      await refreshForNewBuild();
    }
  } catch (_error) {
    // Silent on purpose. If the check endpoint is briefly unavailable,
    // the current session should keep running normally.
  }
}

async function refreshForNewBuild() {
  if (sessionStorage.getItem(APP_RECOVERY_STORAGE_KEY) === `refresh:${APP_BUILD_ID}`) {
    return;
  }
  sessionStorage.setItem(APP_RECOVERY_STORAGE_KEY, `refresh:${APP_BUILD_ID}`);
  state.ui.notice = "A new version of the app is available. Refresh when convenient to load it.";
  render();
}

function withTimeout(promise, timeoutMs) {
  return Promise.race([
    promise,
    new Promise((_, reject) => {
      window.setTimeout(() => reject(new Error("Workspace request timed out.")), timeoutMs);
    })
  ]);
}

async function attemptAutomaticRecovery(error) {
  if (!state.supabase || !state.session) {
    return false;
  }

  const recoveryAlreadyAttempted = sessionStorage.getItem(APP_RECOVERY_STORAGE_KEY) === APP_BUILD_ID;
  const message = String(error?.message || "");
  const shouldRecover =
    !recoveryAlreadyAttempted &&
    (message.includes("JWT") ||
      message.includes("refresh") ||
      message.includes("session") ||
      message.includes("timed out") ||
      message.includes("Invalid") ||
      message.includes("not found"));

  if (!shouldRecover) {
    return false;
  }

  sessionStorage.setItem(APP_RECOVERY_STORAGE_KEY, APP_BUILD_ID);
  await state.supabase.auth.signOut({ scope: "local" });
  state.session = null;
  state.profile = null;
  state.ui.loading = false;
  state.ui.recoveringSession = false;
  state.ui.error = "";
  state.ui.notice = "We refreshed a stale session after the latest update. Please sign in again.";
  render();
  return true;
}

function isAdmin() {
  return state.profile?.role === "admin";
}

function isInactiveUser() {
  return Boolean(state.profile && !state.profile.active);
}

async function fetchProfile(userId) {
  const { data, error } = await state.supabase
    .from("profiles")
    .select("id, email, full_name, role, active")
    .eq("id", userId)
    .single();
  if (error) throw error;
  return data;
}

async function fetchProfiles() {
  const { data, error } = await state.supabase
    .from("profiles")
    .select("id, email, full_name, role, active")
    .order("full_name", { ascending: true });
  if (error) throw error;
  return data || [];
}

async function fetchSettings() {
  const { data, error } = await state.supabase
    .from("app_settings")
    .select("assumptions, routing_rules, communication_settings, lead_sources, statuses, products, carriers")
    .eq("singleton_key", "default")
    .single();
  if (error) throw error;
  return {
    assumptions: { ...seedSettings.assumptions, ...(data.assumptions || {}) },
    routingRules: {
      ...seedSettings.routingRules,
      ...(data.routing_rules || {}),
      sourceRules: Array.isArray(data.routing_rules?.sourceRules) ? data.routing_rules.sourceRules : seedSettings.routingRules.sourceRules
    },
    communicationSettings: {
      ...seedSettings.communicationSettings,
      ...(data.communication_settings || {})
    },
    leadSources: data.lead_sources?.length ? data.lead_sources : seedSettings.leadSources,
    statuses: data.statuses?.length ? data.statuses : seedSettings.statuses,
    products: data.products?.length ? data.products : seedSettings.products,
    carriers: data.carriers?.length ? data.carriers : seedSettings.carriers
  };
}

async function fetchOpportunities() {
  const { data, error } = await state.supabase
    .from("opportunities")
    .select("*")
    .order("date_received", { ascending: false })
    .order("created_at", { ascending: false });
  if (error) throw error;
  return (data || []).map(mapOpportunityFromDb);
}

async function fetchOpportunityActivities() {
  const { data, error } = await state.supabase
    .from("opportunity_activity")
    .select("*")
    .order("created_at", { ascending: false });
  if (error) {
    if (String(error.message || "").includes("opportunity_activity")) {
      return [];
    }
    throw error;
  }
  return data || [];
}

async function fetchOpportunityAttachments() {
  const { data, error } = await state.supabase
    .from("opportunity_attachments")
    .select("*")
    .order("created_at", { ascending: false });
  if (error) {
    if (String(error.message || "").includes("opportunity_attachments")) {
      return [];
    }
    throw error;
  }
  return data || [];
}

async function fetchCoachingNotes() {
  const { data, error } = await state.supabase
    .from("coaching_notes")
    .select("*")
    .order("week_start", { ascending: false });
  if (error) throw error;
  return data || [];
}

async function fetchCalendarConnectionStatus() {
  if (!state.supabase || !state.session) return null;
  try {
    const payload = await callEdgeFunction(GOOGLE_CALENDAR_STATUS_FUNCTION);
    return payload.connection || null;
  } catch {
    return null;
  }
}

async function fetchCommunicationStatus() {
  if (!state.supabase || !state.session) return null;
  try {
    const payload = await callEdgeFunction(COMMUNICATION_STATUS_FUNCTION);
    return payload || null;
  } catch {
    return null;
  }
}

function mapOpportunityFromDb(row) {
  return {
    id: row.id,
    leadNumber: row.lead_number,
    assignedUserId: row.assigned_user_id,
    assignedRepName: row.assigned_rep_name,
    dateReceived: row.date_received,
    leadSource: row.lead_source,
    businessName: row.business_name,
    targetNiche: row.target_niche,
    productFocus: row.product_focus,
    contactName: row.contact_name || "",
    contactEmail: row.contact_email || "",
    contactPhone: row.contact_phone || "",
    carrier: row.carrier,
    incumbentCarrier: row.incumbent_carrier || "",
    policyType: row.policy_type,
    policyTermMonths: Number(row.policy_term_months || 12),
    renewalStatus: row.renewal_status || "Not Started",
    effectiveDate: row.effective_date || "",
    expirationDate: row.expiration_date || "",
    leadCost: Number(row.lead_cost || 0),
    premiumQuoted: Number(row.premium_quoted || 0),
    premiumBound: Number(row.premium_bound || 0),
    status: row.status,
    firstAttemptDate: row.first_attempt_date || "",
    lastActivityDate: row.last_activity_date || "",
    nextFollowUpDate: row.next_follow_up_date || "",
    nextTask: row.next_task || "",
    taskPriority: row.task_priority || "Medium",
    notes: row.notes || ""
  };
}

function mapOpportunityToDb(formData) {
  const assignedProfile = state.profiles.find((item) => item.id === formData.assignedUserId) || state.profile;
  const policyTermMonths = Number(formData.policyTermMonths || 12);
  const resolvedExpirationDate =
    formData.expirationDate ||
    inferExpirationDate({
      status: formData.status,
      effectiveDate: formData.effectiveDate,
      policyTermMonths
    });
  return {
    id: formData.id || undefined,
    lead_number: formData.leadNumber || generateLeadNumber(formData.dateReceived),
    assigned_user_id: assignedProfile.id,
    assigned_rep_name: assignedProfile.full_name,
    date_received: formData.dateReceived,
    lead_source: formData.leadSource,
    business_name: formData.businessName,
    target_niche: formData.targetNiche,
    product_focus: formData.productFocus,
    contact_name: formData.contactName || "",
    contact_email: formData.contactEmail || "",
    contact_phone: formData.contactPhone || "",
    carrier: formData.carrier,
    incumbent_carrier: formData.incumbentCarrier || "",
    policy_type: formData.policyType,
    policy_term_months: policyTermMonths,
    renewal_status: formData.renewalStatus || "Not Started",
    effective_date: formData.effectiveDate || null,
    expiration_date: resolvedExpirationDate || null,
    lead_cost: Number(formData.leadCost || 0),
    premium_quoted: Number(formData.premiumQuoted || 0),
    premium_bound: Number(formData.premiumBound || 0),
    status: formData.status,
    first_attempt_date: formData.firstAttemptDate || null,
    last_activity_date: formData.lastActivityDate || null,
    next_follow_up_date: formData.nextFollowUpDate || null,
    next_task: formData.nextTask || "",
    task_priority: formData.taskPriority || "Medium",
    notes: formData.notes || ""
  };
}

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

function parseDate(value) {
  if (!value) return null;
  const date = new Date(`${value}T00:00:00`);
  return Number.isNaN(date.valueOf()) ? null : date;
}

function addMonthsToIso(dateValue, months) {
  const date = parseDate(dateValue);
  if (!date || !months) return "";
  const next = new Date(date);
  next.setMonth(next.getMonth() + Number(months));
  next.setDate(next.getDate() - 1);
  return next.toISOString().slice(0, 10);
}

function inferExpirationDate({ status, effectiveDate, policyTermMonths }) {
  if (status !== "Bound" || !effectiveDate) return "";
  return addMonthsToIso(effectiveDate, policyTermMonths || 12);
}

function startOfWeek(dateValue) {
  const date = parseDate(dateValue);
  if (!date) return "";
  const day = date.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + diff);
  return date.toISOString().slice(0, 10);
}

function dayDiff(startValue, endValue) {
  const start = parseDate(startValue);
  const end = parseDate(endValue);
  if (!start || !end) return null;
  return Math.round((end - start) / 86400000);
}

function getCarrierPct(carrierName, policyType) {
  const carrier = state.setup.carriers.find((item) => item.name === carrierName);
  if (!carrier) return Number(state.setup.assumptions.averageCommissionPct || 0);
  return Number(policyType === "Renewal" ? carrier.renewalPct : carrier.newPct);
}

function getRepPayoutPct(leadSource) {
  return leadSource === "Self-Generated" ? 0.5 : 0.4;
}

function getRenewalStatuses() {
  return [
    "Not Started",
    "Reviewing",
    "Quoted",
    "Remarketing",
    "Retained",
    "Lost at Renewal"
  ];
}

function getCommunicationTypes() {
  return [
    "Call",
    "Email",
    "Text",
    "Voicemail",
    "Appointment",
    "Quote Delivered",
    "Note"
  ];
}

function getCommunicationOutcomes() {
  return [
    "No Answer",
    "Left Voicemail",
    "Spoke With Insured",
    "Sent Information",
    "Appointment Set",
    "Quote Delivered",
    "Follow-Up Scheduled",
    "Wrong Number",
    "Other"
  ];
}

function needsFollowUp(status) {
  return [
    "New Lead",
    "Attempted",
    "Contacted",
    "Qualified",
    "Quoted",
    "Pending Decision",
    "Nurture / Recycle"
  ].includes(status);
}

function computeOpportunity(row) {
  const contacted = ["Contacted", "Qualified", "Quoted", "Pending Decision", "Bound"].includes(row.status);
  const quoted =
    ["Quoted", "Pending Decision", "Bound"].includes(row.status) ||
    (row.status === "Lost" && Number(row.premiumQuoted) > 0);
  const bound = row.status === "Bound";
  const lost = row.status === "Lost";
  const closed = bound || lost;
  const agencyCommPct = getCarrierPct(row.carrier, row.policyType);
  const repPayoutPct = getRepPayoutPct(row.leadSource);
  const actualAgencyComm = bound ? Number(row.premiumBound) * agencyCommPct : 0;
  const potentialAgencyComm = quoted && !closed ? Number(row.premiumQuoted) * agencyCommPct : 0;
  const actualRepPayout = actualAgencyComm * repPayoutPct;
  const potentialRepPayout = potentialAgencyComm * repPayoutPct;
  const followUpNeeded = needsFollowUp(row.status);
  const followUpOverdue = Boolean(followUpNeeded && row.nextFollowUpDate && todayIso() > row.nextFollowUpDate);
  const taskOverdue = Boolean(row.nextFollowUpDate && todayIso() > row.nextFollowUpDate && row.status !== "Bound" && row.status !== "Lost");
  const followUpBucket = !followUpNeeded
    ? "Closed"
    : !row.nextFollowUpDate
      ? "No Follow-Up Date"
      : todayIso() > row.nextFollowUpDate
        ? "Overdue"
        : dayDiff(todayIso(), row.nextFollowUpDate) <= 1
          ? "Due Soon"
          : "Future";
  const resolvedExpirationDate =
    row.expirationDate ||
    inferExpirationDate({
      status: row.status,
      effectiveDate: row.effectiveDate,
      policyTermMonths: row.policyTermMonths
    });
  const daysToExpiration = resolvedExpirationDate ? dayDiff(todayIso(), resolvedExpirationDate) : null;
  const renewalTracked = row.bound || row.policyType === "Renewal" || Boolean(resolvedExpirationDate);
  const renewalClosed = ["Retained", "Lost at Renewal"].includes(row.renewalStatus);
  const renewalApproaching = renewalTracked && daysToExpiration !== null && daysToExpiration >= 0 && daysToExpiration <= 90;
  const renewalUrgency =
    daysToExpiration === null
      ? "No Expiration"
      : daysToExpiration < 0
        ? "Expired"
        : daysToExpiration <= 7
          ? "Due This Week"
          : daysToExpiration <= 30
            ? "Due This Month"
            : daysToExpiration <= 90
              ? "Upcoming"
              : "Future";
  const renewalNeedsAttention = renewalTracked && !renewalClosed && daysToExpiration !== null && daysToExpiration <= 60;

  return {
    ...row,
    contacted,
    quoted,
    bound,
    lost,
    closed,
    workedSameDay: Boolean(row.firstAttemptDate && row.firstAttemptDate === row.dateReceived),
    followUpNeeded,
    followUpOverdue,
    taskOverdue,
    agencyCommPct,
    actualAgencyComm,
    potentialAgencyComm,
    repPayoutPct,
    actualRepPayout,
    potentialRepPayout,
    ownerNetAgencyComm: actualAgencyComm - actualRepPayout,
    daysOpen: dayDiff(row.dateReceived, todayIso()),
    weekStart: startOfWeek(row.dateReceived),
    month: row.dateReceived?.slice(0, 7) || "",
    freshLead: dayDiff(row.dateReceived, todayIso()) <= Number(state.setup.assumptions.freshLeadWindowDays || 3),
    followUpBucket,
    resolvedExpirationDate,
    daysToExpiration,
    renewalTracked,
    renewalClosed,
    renewalApproaching,
    renewalUrgency,
    renewalNeedsAttention
  };
}

function getVisibleOpportunities() {
  return state.opportunities.map(computeOpportunity);
}

function filterRowsByTimeframe(rows) {
  if (state.ui.timeframe === "week") {
    const currentWeek = startOfWeek(todayIso());
    return rows.filter((row) => row.weekStart === currentWeek);
  }
  if (state.ui.timeframe === "month") {
    const currentMonth = todayIso().slice(0, 7);
    return rows.filter((row) => row.month === currentMonth);
  }
  return rows;
}

function getFilteredOpportunityList(rows) {
  const search = state.ui.search.trim().toLowerCase();
  return rows.filter((row) => {
    const repMatch = state.ui.repFilter === "All" || row.assignedUserId === state.ui.repFilter;
    const sourceMatch = state.ui.sourceFilter === "All" || row.leadSource === state.ui.sourceFilter;
    const statusMatch = state.ui.statusFilter === "All" || row.status === state.ui.statusFilter;
    const fromMatch = !state.ui.dateFrom || row.dateReceived >= state.ui.dateFrom;
    const toMatch = !state.ui.dateTo || row.dateReceived <= state.ui.dateTo;
    const searchMatch = !search || [
      row.leadNumber,
      row.businessName,
      row.assignedRepName,
      row.leadSource,
      row.productFocus,
      row.carrier
    ].join(" ").toLowerCase().includes(search);
    return repMatch && sourceMatch && statusMatch && fromMatch && toMatch && searchMatch;
  });
}

function getAssignableProfiles() {
  return state.profiles.filter((item) => item.active);
}

function getVisibleManagedProfiles() {
  return state.profiles.filter((item) => item.active);
}

function getRepWorkloadRows(rows) {
  return getAssignableProfiles()
    .map((profile) => {
      const repRows = rows.filter((row) => row.assignedUserId === profile.id && !row.closed);
      return {
        id: profile.id,
        name: profile.full_name,
        openLeads: repRows.length,
        overdue: repRows.filter((row) => row.taskOverdue).length,
        stale: repRows.filter((row) => row.daysOpen >= 7).length,
        quoted: repRows.filter((row) => row.status === "Quoted").length,
        renewalsDue: repRows.filter((row) => row.renewalNeedsAttention).length
      };
    })
    .sort((a, b) => b.openLeads - a.openLeads);
}

function getRoutingSuggestions(workloadRows) {
  if (workloadRows.length < 2) return [];
  const busiest = workloadRows[0];
  const lightest = workloadRows[workloadRows.length - 1];
  const gap = busiest.openLeads - lightest.openLeads;
  if (gap < 5) return [];
  return [
    `${busiest.name} has ${busiest.openLeads} open leads while ${lightest.name} has ${lightest.openLeads}. Consider rebalancing a few active accounts.`,
    busiest.overdue > lightest.overdue
      ? `${busiest.name} also carries more overdue tasks. Auto-assigning to ${lightest.name} could reduce follow-up risk.`
      : `A round-robin or source-based rule would smooth out the current workload gap.`
  ];
}

function getAssignedProfileForForm(formData, existingOpportunity = null) {
  if (!isAdmin()) {
    return state.profile;
  }
  if (formData.assignedUserId && formData.assignedUserId !== "auto") {
    return state.profiles.find((item) => item.id === formData.assignedUserId) || state.profile;
  }
  if (existingOpportunity?.assignedUserId) {
    return state.profiles.find((item) => item.id === existingOpportunity.assignedUserId) || state.profile;
  }
  return resolveAutoAssignedProfile(formData) || state.profile;
}

function resolveAutoAssignedProfile(formData) {
  const activeReps = getAssignableProfiles();
  if (!activeReps.length) return null;
  if (!isAdmin()) return state.profile;

  if (String(formData.leadSource || "").trim().toLowerCase() === "self-generated") {
    return state.profile;
  }

  const { autoAssignEnabled, mode, sourceRules = [], roundRobinCursor = 0 } = state.setup.routingRules || seedSettings.routingRules;
  if (!autoAssignEnabled) {
    return state.profile;
  }

  if (mode === "source_rule") {
    const matchedRule = sourceRules.find((rule) => rule.source === formData.leadSource && rule.userId);
    if (matchedRule) {
      const matchedProfile = activeReps.find((item) => item.id === matchedRule.userId);
      if (matchedProfile) {
        return matchedProfile;
      }
    }
  }

  const normalizedCursor = Number(roundRobinCursor || 0) % activeReps.length;
  return activeReps[normalizedCursor];
}

function normalizeImportHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function matchConfiguredValue(value, options, fallback = "") {
  if (!value && fallback) return fallback;
  const raw = String(value || "").trim();
  if (!raw) return fallback;
  const direct = options.find((item) => item === raw);
  if (direct) return direct;
  const normalized = normalizeImportHeader(raw);
  return options.find((item) => normalizeImportHeader(item) === normalized) || fallback;
}

function coerceImportDate(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return "";
    return `${String(parsed.y).padStart(4, "0")}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  const raw = String(value).trim();
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const normalized = raw.replace(/\//g, "-");
  const date = new Date(normalized);
  if (Number.isNaN(date.valueOf())) return "";
  return date.toISOString().slice(0, 10);
}

function coerceImportNumber(value, fallback = 0) {
  if (value === null || value === undefined || value === "") return fallback;
  const numeric = Number(String(value).replace(/[$,%\s,]/g, ""));
  return Number.isFinite(numeric) ? numeric : fallback;
}

function findProfileByImportValue(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  const normalized = normalizeImportHeader(raw);
  return getAssignableProfiles().find((profile) => {
    const email = String(profile.email || "").trim().toLowerCase();
    return profile.id === raw || email === raw.toLowerCase() || normalizeImportHeader(profile.full_name) === normalized;
  }) || null;
}

function createLeadNumberFactory() {
  const countsByDate = state.opportunities.reduce((acc, row) => {
    const key = row.dateReceived || todayIso();
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  return (dateReceived) => {
    const key = dateReceived || todayIso();
    countsByDate[key] = (countsByDate[key] || 0) + 1;
    const compact = `${key.slice(2, 4)}${key.slice(5, 7)}${key.slice(8, 10)}`;
    return `${compact}-${String(countsByDate[key]).padStart(4, "0")}`;
  };
}

function resolveImportedAssignedProfile(row, roundRobinState) {
  const explicitProfile = findProfileByImportValue(row.assignedRepRaw);
  if (row.assignedRepRaw && !explicitProfile) {
    return { profile: null, strategy: "Assigned rep not recognized", error: `Assigned Rep "${row.assignedRepRaw}" does not match an active rep.` };
  }
  if (explicitProfile) {
    return { profile: explicitProfile, strategy: "Explicit rep from file", error: "" };
  }
  if (!isAdmin()) {
    return { profile: state.profile, strategy: "Imported by rep", error: "" };
  }
  if (String(row.leadSource || "").trim().toLowerCase() === "self-generated") {
    return { profile: state.profile, strategy: "Self-generated kept with uploader", error: "" };
  }

  const activeReps = getAssignableProfiles();
  if (!activeReps.length) {
    return { profile: state.profile, strategy: "No active reps available", error: "" };
  }

  const { autoAssignEnabled, mode, sourceRules = [] } = state.setup.routingRules || seedSettings.routingRules;
  if (!autoAssignEnabled) {
    return { profile: state.profile, strategy: "Auto assign disabled, kept with uploader", error: "" };
  }

  if (mode === "source_rule") {
    const matchedRule = sourceRules.find((rule) => rule.source === row.leadSource && rule.userId);
    if (matchedRule) {
      const matchedProfile = activeReps.find((item) => item.id === matchedRule.userId);
      if (matchedProfile) {
        return { profile: matchedProfile, strategy: `Source rule for ${row.leadSource}`, error: "" };
      }
    }
  }

  const nextIndex = Number(roundRobinState.cursor || 0) % activeReps.length;
  const profile = activeReps[nextIndex];
  roundRobinState.cursor = (nextIndex + 1) % activeReps.length;
  return { profile, strategy: "Round robin fallback", error: "" };
}

function mapImportedLeadRow(rawRow, rowNumber, roundRobinState) {
  const get = (...headers) => {
    const entries = Object.entries(rawRow || {});
    for (const header of headers) {
      const found = entries.find(([key]) => normalizeImportHeader(key) === normalizeImportHeader(header));
      if (found) return found[1];
    }
    return "";
  };

  const dateReceived = coerceImportDate(get("Date Received", "Received Date", "Lead Date")) || todayIso();
  const leadSource = matchConfiguredValue(get("Lead Source", "Source"), state.setup.leadSources, state.setup.leadSources[0] || "");
  const productFocus = matchConfiguredValue(get("Product Focus", "Product"), state.setup.products, state.setup.products[0] || "");
  const carrier = matchConfiguredValue(get("Carrier"), state.setup.carriers.map((item) => item.name), state.setup.carriers[0]?.name || "");
  const status = matchConfiguredValue(get("Status", "Pipeline Stage"), state.setup.statuses, state.setup.statuses[0] || "New Lead");
  const policyType = matchConfiguredValue(get("Policy Type"), ["New", "Renewal"], "New");
  const renewalStatus = matchConfiguredValue(get("Renewal Status"), getRenewalStatuses(), "Not Started");
  const taskPriority = matchConfiguredValue(get("Task Priority", "Priority"), ["High", "Medium", "Low"], "Medium");
  const policyTermMonths = [3, 6, 12].includes(coerceImportNumber(get("Policy Term", "Policy Term (Months)", "Term"), 12))
    ? coerceImportNumber(get("Policy Term", "Policy Term (Months)", "Term"), 12)
    : 12;
  const businessName = String(get("Business Name", "Business", "Account Name")).trim();

  const mappedRow = {
    id: "",
    leadNumber: "",
    assignedUserId: "auto",
    assignedRepName: "",
    dateReceived,
    leadSource,
    businessName,
    targetNiche: String(get("Target Niche", "Niche", "Industry")).trim(),
    productFocus,
    contactName: String(get("Contact Name", "Primary Contact")).trim(),
    contactEmail: String(get("Contact Email", "Email")).trim(),
    contactPhone: String(get("Contact Phone", "Phone")).trim(),
    carrier,
    incumbentCarrier: String(get("Incumbent Carrier")).trim(),
    policyType,
    policyTermMonths,
    renewalStatus,
    effectiveDate: coerceImportDate(get("Effective Date")),
    expirationDate: coerceImportDate(get("Expiration Date")),
    leadCost: coerceImportNumber(get("Lead Cost", "Cost"), 0),
    premiumQuoted: coerceImportNumber(get("Premium Quoted", "Quoted Premium"), 0),
    premiumBound: coerceImportNumber(get("Premium Bound", "Bound Premium"), 0),
    status,
    firstAttemptDate: coerceImportDate(get("First Attempt Date")),
    lastActivityDate: coerceImportDate(get("Last Activity Date")),
    nextFollowUpDate: coerceImportDate(get("Next Follow-Up Date", "Follow-Up Date")),
    nextTask: String(get("Next Task", "Task")).trim(),
    taskPriority,
    notes: String(get("Notes", "Lead Notes")).trim(),
    assignedRepRaw: String(get("Assigned Rep", "Rep", "Producer")).trim()
  };

  const errors = [];
  if (!businessName) {
    errors.push("Business Name is required.");
  }

  const assignment = resolveImportedAssignedProfile(mappedRow, roundRobinState);
  if (assignment.error) {
    errors.push(assignment.error);
  }

  mappedRow.assignedUserId = assignment.profile?.id || "";
  mappedRow.assignedRepName = assignment.profile?.full_name || "";

  return {
    rowNumber,
    strategy: assignment.strategy,
    assigneeName: assignment.profile?.full_name || "Unassigned",
    errors,
    data: mappedRow
  };
}

async function buildLeadImportPreview(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("That file does not contain any importable rows.");
  }
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  if (!rows.length) {
    throw new Error("That file does not contain any lead rows.");
  }
  const roundRobinState = {
    cursor: Number(state.setup.routingRules?.roundRobinCursor || 0)
  };
  return rows.map((row, index) => mapImportedLeadRow(row, index + 2, roundRobinState));
}

function downloadLeadImportTemplate() {
  const templateRows = [
    {
      "Business Name": "Acme Builders",
      "Lead Source": "Purchased Leads",
      "Contact Name": "Jordan Smith",
      "Contact Email": "jordan@acmebuilders.com",
      "Contact Phone": "555-555-0199",
      "Date Received": todayIso(),
      "Product Focus": state.setup.products[0] || "GL / BOP",
      "Target Niche": "Contractor",
      "Assigned Rep": "",
      Status: "New Lead",
      "Lead Cost": 42.5,
      Notes: "Imported from owner lead batch"
    }
  ];
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(templateRows);
  XLSX.utils.book_append_sheet(workbook, sheet, "Lead Import Template");
  XLSX.writeFile(workbook, "golden-leaf-lead-import-template.csv", { bookType: "csv" });
}

async function importLeadPreviewRows() {
  const validRows = state.ui.importPreviewRows.filter((row) => !row.errors.length);
  if (!validRows.length) {
    throw new Error("Fix the import errors before continuing.");
  }

  state.ui.importingCsv = true;
  state.ui.error = "";
  state.ui.notice = "";
  render();

  try {
    const nextLeadNumber = createLeadNumberFactory();
    const payloads = validRows.map((row) =>
      mapOpportunityToDb({
        ...row.data,
        leadNumber: nextLeadNumber(row.data.dateReceived)
      })
    );

    const { data, error } = await state.supabase
      .from("opportunities")
      .insert(payloads)
      .select("*");
    if (error) throw error;

    await Promise.all(
      (data || []).map((item) =>
        logOpportunityActivity({
          opportunityId: item.id,
          title: "Lead imported",
          detail: `${item.business_name || "Lead"} imported from ${state.ui.importFileName || "CSV"} and assigned to ${item.assigned_rep_name}.`
        })
      )
    );

    state.ui.importPreviewRows = [];
    state.ui.importFileName = "";
    state.ui.notice = `${payloads.length} lead${payloads.length === 1 ? "" : "s"} imported successfully.`;
    await loadWorkspace();
  } finally {
    state.ui.importingCsv = false;
  }
}

function getRemovedProfiles() {
  return state.profiles.filter((item) => !item.active);
}

function getSelectedOpportunitySet() {
  return new Set(state.ui.selectedOpportunityIds);
}

function getAssignedLeadCount(profileId) {
  return state.opportunities.filter((item) => item.assignedUserId === profileId).length;
}

function getOpportunityView() {
  return state.ui.opportunityView;
}

function getPipelineGroups(rows) {
  return state.setup.statuses.map((status) => ({
    status,
    rows: rows.filter((row) => row.status === status)
  }));
}

function getPipelinePhaseGroups(rows) {
  const phaseDefinitions = [
    {
      title: "Fresh Lead Intake",
      description: "New assignments and first contact work.",
      statuses: ["New Lead", "Attempted", "Contacted"]
    },
    {
      title: "Active Pipeline",
      description: "Qualified opportunities and open conversations.",
      statuses: ["Qualified", "Quoted", "Pending Decision", "Nurture / Recycle"]
    },
    {
      title: "Closed Outcomes",
      description: "Won and lost deals.",
      statuses: ["Bound", "Lost"]
    }
  ];

  return phaseDefinitions.map((phase) => ({
    ...phase,
    columns: phase.statuses.map((status) => ({
      status,
      rows: rows.filter((row) => row.status === status)
    }))
  }));
}

function getAvailableTabs() {
  return [
    { id: "dashboard", label: "Dashboard" },
    { id: "opportunities", label: isAdmin() ? "Pipeline" : "My Pipeline" },
    { id: "integrations", label: "Integrations" },
    ...(isAdmin() ? [{ id: "reports", label: "Reports" }] : []),
    { id: "scorecards", label: "Scorecards" },
    { id: "coaching", label: "Coaching" },
    ...(isAdmin() ? [{ id: "setup", label: "Setup" }] : [])
  ];
}

function getSetupTabs() {
  return [
    { id: "users", label: "Users" },
    { id: "routing", label: "Routing" },
    { id: "assumptions", label: "Assumptions" },
    { id: "carriers", label: "Carrier Table" }
  ];
}

function getOpportunityTabs() {
  return [
    { id: "stage", label: "Move Stages" },
    { id: "update", label: "Update Leads" },
    { id: "create", label: "New Lead" }
  ];
}

function getDefaultActiveTab() {
  return isAdmin() ? "dashboard" : "opportunities";
}

function ensureActiveTab() {
  const availableTabs = getAvailableTabs().map((tab) => tab.id);
  if (!availableTabs.includes(state.ui.activeTab)) {
    state.ui.activeTab = getDefaultActiveTab();
  }
}

function summarize(rows) {
  return {
    totalLeads: rows.length,
    openLeads: rows.filter((row) => !row.closed).length,
    overdueFollowUp: rows.filter((row) => row.followUpOverdue).length,
    quotesInPipeline: rows.filter((row) => row.quoted && !row.closed).length,
    binds: rows.filter((row) => row.bound).length,
    boundPremium: sum(rows, "premiumBound"),
    pipelineAgencyComm: sum(rows, "potentialAgencyComm"),
    actualAgencyComm: sum(rows, "actualAgencyComm"),
    actualRepPayout: sum(rows, "actualRepPayout"),
    potentialRepPayout: sum(rows, "potentialRepPayout"),
    ownerNetAgencyComm: sum(rows, "ownerNetAgencyComm")
  };
}

function sum(rows, key) {
  return rows.reduce((total, row) => total + Number(row[key] || 0), 0);
}

function getRepScorecards(rows) {
  const reps = isAdmin() ? state.profiles.filter((item) => item.role === "rep" || item.role === "admin") : [state.profile];
  return reps.map((rep) => {
    const repRows = rows.filter((row) => row.assignedUserId === rep.id);
    const leads = repRows.length;
    const quotes = repRows.filter((row) => row.quoted).length;
    return {
      id: rep.id,
      name: rep.full_name,
      leads,
      binds: repRows.filter((row) => row.bound).length,
      sameDayRate: leads ? repRows.filter((row) => row.workedSameDay).length / leads : 0,
      contactRate: leads ? repRows.filter((row) => row.contacted).length / leads : 0,
      quoteRate: leads ? quotes / leads : 0,
      quoteToBindRate: quotes ? repRows.filter((row) => row.bound).length / quotes : 0,
      actualAgencyComm: sum(repRows, "actualAgencyComm"),
      actualRepPayout: sum(repRows, "actualRepPayout")
    };
  });
}

function getRoiRows(rows) {
  return state.setup.leadSources.map((source) => {
    const sourceRows = rows.filter((row) => row.leadSource === source);
    const count = sourceRows.length;
    const quotes = sourceRows.filter((row) => row.quoted).length;
    return {
      source,
      count,
      spend: sum(sourceRows, "leadCost"),
      quoteRate: count ? quotes / count : 0,
      bindRate: count ? sourceRows.filter((row) => row.bound).length / count : 0,
      actualAgencyComm: sum(sourceRows, "actualAgencyComm"),
      actualRepPayout: sum(sourceRows, "actualRepPayout")
    };
  });
}

function getStageCounts(rows) {
  return state.setup.statuses.map((status) => ({
    label: status,
    value: rows.filter((row) => row.status === status).length
  }));
}

function getRepCommissionRows(rows) {
  return rows
    .filter((row) => row.actualRepPayout > 0 || row.potentialRepPayout > 0)
    .map((row) => ({
      businessName: row.businessName,
      status: row.status,
      carrier: row.carrier,
      policyType: row.policyType,
      actualRepPayout: row.actualRepPayout,
      potentialRepPayout: row.potentialRepPayout
    }));
}

function getOpportunityTimeline(opportunityId) {
  return state.opportunityActivities
    .filter((item) => item.opportunity_id === opportunityId)
    .sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
}

function getCommunicationActivities(rows) {
  const visibleIds = new Set(rows.map((row) => row.id));
  return state.opportunityActivities.filter(
    (item) => visibleIds.has(item.opportunity_id) && item.activity_type && item.activity_type !== "system"
  );
}

function getOpportunityAttachments(opportunityId) {
  return state.opportunityAttachments
    .filter((item) => item.opportunity_id === opportunityId)
    .sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
}

function getTaskQueue(rows) {
  return rows
    .filter((row) => !row.closed && (row.nextTask || row.nextFollowUpDate))
    .sort((a, b) => {
      const priorityScore = { High: 0, Medium: 1, Low: 2 };
      const aScore = priorityScore[a.taskPriority || "Medium"] ?? 1;
      const bScore = priorityScore[b.taskPriority || "Medium"] ?? 1;
      if (a.taskOverdue !== b.taskOverdue) return a.taskOverdue ? -1 : 1;
      if (aScore !== bScore) return aScore - bScore;
      return (a.nextFollowUpDate || "9999-12-31").localeCompare(b.nextFollowUpDate || "9999-12-31");
    });
}

function getCommunicationSummary(rows) {
  const activities = getCommunicationActivities(rows);
  return {
    touches: activities.length,
    conversations: activities.filter((item) => item.outcome === "Spoke With Insured").length,
    appointments: activities.filter((item) => item.outcome === "Appointment Set" || item.activity_type === "Appointment").length,
    quoteTouches: activities.filter((item) => item.activity_type === "Quote Delivered" || item.outcome === "Quote Delivered").length
  };
}

function getRepCommunicationRows(rows) {
  const relevantRows = getUserScopedRows(rows);
  const opportunityIdsByRep = new Map();
  relevantRows.forEach((row) => {
    const existing = opportunityIdsByRep.get(row.assignedUserId) || new Set();
    existing.add(row.id);
    opportunityIdsByRep.set(row.assignedUserId, existing);
  });

  const reps = isAdmin() ? getAssignableProfiles() : [state.profile];
  return reps
    .map((rep) => {
      const ids = opportunityIdsByRep.get(rep.id) || new Set();
      const activities = state.opportunityActivities.filter(
        (item) => ids.has(item.opportunity_id) && item.activity_type && item.activity_type !== "system"
      );
      return {
        id: rep.id,
        name: rep.full_name,
        touches: activities.length,
        conversations: activities.filter((item) => item.outcome === "Spoke With Insured").length,
        appointments: activities.filter((item) => item.outcome === "Appointment Set" || item.activity_type === "Appointment").length,
        quoteTouches: activities.filter((item) => item.activity_type === "Quote Delivered" || item.outcome === "Quote Delivered").length
      };
    })
    .filter((row) => row.touches > 0 || row.conversations > 0 || row.appointments > 0 || row.quoteTouches > 0)
    .sort((a, b) => b.touches - a.touches);
}

function getStaleLeads(rows) {
  return rows
    .filter((row) => !row.closed && row.daysOpen >= 7)
    .sort((a, b) => b.daysOpen - a.daysOpen);
}

function getDashboardAlerts(rows) {
  const taskQueue = getTaskQueue(rows);
  const staleLeads = getStaleLeads(rows);
  const overdueTasks = taskQueue.filter((row) => row.taskOverdue);
  const quotesAging = rows.filter((row) => row.status === "Quoted" && row.daysOpen >= 5 && !row.closed);
  const renewalsDue = rows.filter((row) => row.renewalNeedsAttention);
  return {
    overdueTasks,
    staleLeads,
    quotesAging,
    renewalsDue
  };
}

function getRenewalRows(rows) {
  return rows
    .filter((row) => row.renewalTracked)
    .sort((a, b) => {
      const aDays = a.daysToExpiration ?? 9999;
      const bDays = b.daysToExpiration ?? 9999;
      return aDays - bDays;
    });
}

function getRenewalSummary(rows) {
  const renewalRows = getRenewalRows(rows);
  return {
    tracked: renewalRows.length,
    dueThirty: renewalRows.filter((row) => row.daysToExpiration !== null && row.daysToExpiration >= 0 && row.daysToExpiration <= 30).length,
    dueSixty: renewalRows.filter((row) => row.daysToExpiration !== null && row.daysToExpiration >= 0 && row.daysToExpiration <= 60).length,
    retained: renewalRows.filter((row) => row.renewalStatus === "Retained").length,
    lostAtRenewal: renewalRows.filter((row) => row.renewalStatus === "Lost at Renewal").length
  };
}

function getRenewalStageCounts(rows) {
  return getRenewalStatuses().map((status) => ({
    label: status,
    value: rows.filter((row) => row.renewalTracked && row.renewalStatus === status).length
  }));
}

function getRepPerformanceRows(rows) {
  return getRepScorecards(rows)
    .filter((row) => row.leads > 0)
    .sort((a, b) => b.actualAgencyComm - a.actualAgencyComm)
    .map((row) => ({
      ...row,
      overdueCount: rows.filter((item) => item.assignedUserId === row.id && item.taskOverdue).length,
      staleCount: rows.filter((item) => item.assignedUserId === row.id && !item.closed && item.daysOpen >= 7).length
    }));
}

function getSourcePerformanceRows(rows) {
  return getRoiRows(rows)
    .filter((row) => row.count > 0)
    .sort((a, b) => b.actualAgencyComm - a.actualAgencyComm)
    .map((row) => ({
      ...row,
      netRevenue: row.actualAgencyComm - row.spend,
      revenuePerLead: row.count ? row.actualAgencyComm / row.count : 0
    }));
}

function getMonthlyTrendRows(rows, key, monthsBack = 6) {
  const months = [];
  const now = parseDate(todayIso()) || new Date();
  for (let index = monthsBack - 1; index >= 0; index -= 1) {
    const date = new Date(now.getFullYear(), now.getMonth() - index, 1);
    const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    months.push({
      label: date.toLocaleString("en-US", { month: "short" }),
      value: sum(rows.filter((row) => row.month === monthKey), key)
    });
  }
  return months;
}

function getRepLeaderboardRows(rows) {
  return getRepScorecards(rows)
    .filter((row) => row.leads > 0)
    .sort((a, b) => b.actualAgencyComm - a.actualAgencyComm)
    .slice(0, 6)
    .map((row) => ({
      label: row.name,
      value: row.actualAgencyComm
    }));
}

function getRepSourceRows(rows) {
  return getRoiRows(rows)
    .filter((row) => row.count > 0)
    .map((row) => ({
      label: row.source,
      value: row.actualRepPayout
    }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 6);
}

function getUserScopedRows(rows) {
  return isAdmin() ? rows : rows.filter((row) => row.assignedUserId === state.profile?.id);
}

function getCoachingRows(rows) {
  const weekStart = startOfWeek(todayIso());
  const reps = isAdmin() ? state.profiles.filter((item) => item.active) : [state.profile];
  return reps.map((rep) => {
    const repRows = rows.filter((row) => row.assignedUserId === rep.id && row.weekStart === weekStart);
    const note = state.coachingNotes.find((item) => item.rep_user_id === rep.id && item.week_start === weekStart);
    const leads = repRows.length;
    return {
      repId: rep.id,
      repName: rep.full_name,
      weekStart,
      leads,
      contactRate: leads ? repRows.filter((row) => row.contacted).length / leads : 0,
      quotes: repRows.filter((row) => row.quoted).length,
      binds: repRows.filter((row) => row.bound).length,
      biggestGap: note?.biggest_gap || inferBiggestGap(repRows, leads),
      behaviorToImprove: note?.behavior_to_improve || "",
      actionCommitment: note?.action_commitment || "",
      nextReviewNotes: note?.next_review_notes || ""
    };
  });
}

function inferBiggestGap(rows, leads) {
  const assumptions = state.setup.assumptions;
  const metrics = [
    {
      label: "Same-day work rate",
      delta: (leads ? rows.filter((row) => row.workedSameDay).length / leads : 0) - Number(assumptions.sameDayWorkedTargetPct || 0)
    },
    {
      label: "Contact rate",
      delta: (leads ? rows.filter((row) => row.contacted).length / leads : 0) - Number(assumptions.contactRateTargetPct || 0)
    },
    {
      label: "Quote rate",
      delta: (leads ? rows.filter((row) => row.quoted).length / leads : 0) - Number(assumptions.quoteRateTargetPct || 0)
    }
  ];
  metrics.sort((a, b) => a.delta - b.delta);
  return metrics[0].label;
}

function render() {
  document.body.classList.toggle("app-shell-auth", !state.session);
  document.body.classList.toggle("app-shell-logged-in", Boolean(state.session && !isInactiveUser()));
  heroActionsEl.innerHTML = renderHeroActions();
  topNavEl.innerHTML = renderTopNav();
  if (assistantRootEl) {
    assistantRootEl.innerHTML = renderAssistant();
  }

  if (!SUPABASE_READY) {
    appEl.innerHTML = renderSetupRequired();
    bindShellEvents();
    bindAppEvents();
    return;
  }

  if (state.ui.loading) {
    appEl.innerHTML = `<section class="panel"><div class="empty-state"><h3>Loading workspace</h3><p>Connecting to the shared agency database.</p></div></section>`;
    bindShellEvents();
    bindAppEvents();
    return;
  }

  if (!state.session) {
    appEl.innerHTML = renderLogin();
    bindShellEvents();
    bindAppEvents();
    return;
  }

  if (isInactiveUser()) {
    appEl.innerHTML = renderInactiveUser();
    bindShellEvents();
    bindAppEvents();
    return;
  }

  if (state.ui.recoveringSession) {
    appEl.innerHTML = renderRecoveryState();
    bindShellEvents();
    bindAppEvents();
    return;
  }

  ensureActiveTab();

  const allRows = getVisibleOpportunities();
  const timeframeRows = filterRowsByTimeframe(allRows);
  const listRows = getFilteredOpportunityList(allRows);
  const assignableProfiles = getAssignableProfiles();
  const visibleManagedProfiles = getVisibleManagedProfiles();
  const removedProfiles = getRemovedProfiles();
  const workloadRows = isAdmin() ? getRepWorkloadRows(allRows) : [];
  const routingSuggestions = isAdmin() ? getRoutingSuggestions(workloadRows) : [];
  const opportunityView = getOpportunityView();
  const pipelineGroups = getPipelineGroups(listRows);
  const pipelinePhaseGroups = getPipelinePhaseGroups(listRows);
  const summary = summarize(timeframeRows);
  const userSummary = summarize(getUserScopedRows(timeframeRows));
  const scorecards = getRepScorecards(timeframeRows);
  const roiRows = getRoiRows(getUserScopedRows(timeframeRows));
  const coachingRows = getCoachingRows(allRows);
  const dashboardStageRows = getStageCounts(getUserScopedRows(timeframeRows));
  const monthlyAgencyTrendRows = getMonthlyTrendRows(timeframeRows, "actualAgencyComm");
  const monthlyRepTrendRows = getMonthlyTrendRows(getUserScopedRows(timeframeRows), "actualRepPayout");
  const leaderboardRows = getRepLeaderboardRows(timeframeRows);
  const repSourceRows = getRepSourceRows(getUserScopedRows(timeframeRows));
  const repCommissionRows = getRepCommissionRows(getUserScopedRows(timeframeRows));
  const taskQueueRows = getTaskQueue(getUserScopedRows(allRows));
  const dashboardAlerts = getDashboardAlerts(getUserScopedRows(allRows));
  const communicationSummary = getCommunicationSummary(getUserScopedRows(allRows));
  const repCommunicationRows = getRepCommunicationRows(allRows);
  const renewalSummary = getRenewalSummary(getUserScopedRows(timeframeRows));
  const renewalRows = getRenewalRows(getUserScopedRows(allRows));
  const renewalStageRows = getRenewalStageCounts(getUserScopedRows(timeframeRows));
  const ownerRepPerformanceRows = isAdmin() ? getRepPerformanceRows(timeframeRows) : [];
  const ownerSourcePerformanceRows = isAdmin() ? getSourcePerformanceRows(timeframeRows) : [];
  const activeOpportunity =
    state.opportunities.find((item) => item.id === state.ui.activeOpportunityId) ||
    blankOpportunity();
  const activeOpportunityTimeline = activeOpportunity.id ? getOpportunityTimeline(activeOpportunity.id) : [];

  appEl.innerHTML = `
    ${state.ui.error ? `<section class="panel"><p class="error-banner">${escapeHtml(state.ui.error)}</p></section>` : ""}
    ${state.ui.notice ? `<section class="panel"><p class="notice-banner">${escapeHtml(state.ui.notice)}</p></section>` : ""}

    ${state.ui.activeTab === "dashboard" ? `
    <section class="panel workspace-panel" id="dashboard">
      <div class="panel-header">
        <div>
          <h2>${isAdmin() ? "Agency Dashboard" : "My Pipeline Dashboard"}</h2>
          <p>${isAdmin() ? "Owner view across the full agency." : "Producer view scoped to your assigned book of business."}</p>
        </div>
        <label>
          Timeframe
          <select id="timeframeSelect">
            <option value="all" ${state.ui.timeframe === "all" ? "selected" : ""}>All time</option>
            <option value="week" ${state.ui.timeframe === "week" ? "selected" : ""}>Current week</option>
            <option value="month" ${state.ui.timeframe === "month" ? "selected" : ""}>Current month</option>
          </select>
        </label>
      </div>
      <div class="dashboard-grid">
        ${statCard("Total Leads", isAdmin() ? summary.totalLeads : userSummary.totalLeads, "In selected timeframe")}
        ${statCard("Open Leads", isAdmin() ? summary.openLeads : userSummary.openLeads, "Still active")}
        ${statCard("Overdue Follow-Up", isAdmin() ? summary.overdueFollowUp : userSummary.overdueFollowUp, "Needs attention")}
        ${statCard("Quotes in Pipeline", isAdmin() ? summary.quotesInPipeline : userSummary.quotesInPipeline, "Not yet closed")}
        ${statCard("Binds", isAdmin() ? summary.binds : userSummary.binds, "Closed won")}
      </div>
      <div class="metrics-grid">
        ${isAdmin()
          ? `
            ${kpiCard("Bound Premium", formatCurrency(summary.boundPremium), "Actual premium")}
            ${kpiCard("Pipeline Agency Comm", formatCurrency(summary.pipelineAgencyComm), "Projected commission")}
            ${kpiCard("Actual Agency Comm", formatCurrency(summary.actualAgencyComm), "Closed won revenue")}
            ${kpiCard("Owner Net Agency Comm", formatCurrency(summary.ownerNetAgencyComm), "After rep payout")}
          `
          : `
            ${kpiCard("Bound Premium", formatCurrency(userSummary.boundPremium), "Your closed premium")}
            ${kpiCard("Projected Commission", formatCurrency(userSummary.potentialRepPayout), "Open quoted payout")}
            ${kpiCard("Earned Commission", formatCurrency(userSummary.actualRepPayout), "Your take-home on bound business")}
            ${kpiCard("Follow-Ups Due", userSummary.overdueFollowUp, "Leads needing action")}
          `}
      </div>
      <div class="two-column">
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>${isAdmin() ? "Renewal Snapshot" : "My Renewal Snapshot"}</h3>
              <p>${isAdmin() ? "Watch the retention book coming due across the agency." : "Stay ahead of renewals in your assigned book."}</p>
            </div>
          </div>
          <div class="dashboard-grid compact-dashboard-grid">
            ${statCard("Tracked Renewals", renewalSummary.tracked, "Policies with renewal dates")}
            ${statCard("Due in 30 Days", renewalSummary.dueThirty, "Priority retention window")}
            ${statCard("Due in 60 Days", renewalSummary.dueSixty, "Upcoming renewal work")}
            ${statCard("Retained", renewalSummary.retained, "Marked retained")}
            ${statCard("Lost at Renewal", renewalSummary.lostAtRenewal, "Retention losses")}
          </div>
        </article>
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>${isAdmin() ? "Renewal Stage Mix" : "My Renewal Stage Mix"}</h3>
              <p>${isAdmin() ? "See how the renewal book is moving from review to outcome." : "Track the health of your renewal pipeline."}</p>
            </div>
          </div>
          ${renderBarChart(renewalStageRows, { valueFormatter: formatWholeNumber })}
        </article>
      </div>
      <div class="two-column">
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>${isAdmin() ? "Pipeline by Stage" : "My Pipeline by Stage"}</h3>
              <p>${isAdmin() ? "Quick visual of where the agency book is sitting." : "Your current workload across the lead cycle."}</p>
            </div>
          </div>
          ${renderBarChart(dashboardStageRows, { valueFormatter: formatWholeNumber })}
        </article>
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>${isAdmin() ? "Trend Snapshot" : "Commission Trend"}</h3>
              <p>${isAdmin() ? "Recent closed agency commission by month." : "Recent earned payout by month."}</p>
            </div>
          </div>
          ${renderBarChart(isAdmin() ? monthlyAgencyTrendRows : monthlyRepTrendRows, { valueFormatter: formatCompactCurrency })}
        </article>
      </div>
      <div class="two-column">
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>${isAdmin() ? "Urgent Attention Queue" : "My Action Queue"}</h3>
              <p>${isAdmin() ? "Immediate follow-ups and stalled leads across the visible workspace." : "The next tasks that need your attention first."}</p>
            </div>
          </div>
          ${renderTaskQueue(taskQueueRows)}
        </article>
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>${isAdmin() ? "Stale Pipeline Alerts" : "My Stale Lead Alerts"}</h3>
              <p>${isAdmin() ? "Leads and quotes sitting too long without movement." : "Leads in your book that are at risk of going cold."}</p>
            </div>
          </div>
          ${renderAlertsPanel(dashboardAlerts)}
        </article>
      </div>
      <div class="two-column">
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>${isAdmin() ? "Communication Pulse" : "My Communication Pulse"}</h3>
              <p>${isAdmin() ? "Touches, conversations, and appointments logged from the team workspace." : "Your recent lead outreach activity."}</p>
            </div>
          </div>
          <div class="dashboard-grid compact-dashboard-grid">
            ${statCard("Touches Logged", communicationSummary.touches, "Calls, emails, texts, notes")}
            ${statCard("Conversations", communicationSummary.conversations, "Spoke with insured")}
            ${statCard("Appointments Set", communicationSummary.appointments, "Scheduled meetings")}
            ${statCard("Quote Deliveries", communicationSummary.quoteTouches, "Quotes delivered to prospects")}
          </div>
        </article>
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>${isAdmin() ? "Touch Activity by Rep" : "My Outreach Mix"}</h3>
              <p>${isAdmin() ? "Quick coaching view of who is actively working their book." : "See how your touches are stacking up."}</p>
            </div>
          </div>
          ${renderCommunicationLeaderboard(repCommunicationRows)}
        </article>
      </div>
      <div class="two-column">
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>${isAdmin() ? "Top Producers" : "Commission by Lead Source"}</h3>
              <p>${isAdmin() ? "Largest agency revenue contributors in the selected timeframe." : "Where your payout is coming from."}</p>
            </div>
          </div>
          ${renderBarChart(isAdmin() ? leaderboardRows : repSourceRows, { valueFormatter: formatCompactCurrency })}
        </article>
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>${isAdmin() ? "Admin Actions" : "My Commission Detail"}</h3>
              <p>${isAdmin() ? "Quick exports and finance controls for the owner workspace." : "Open and closed commission items in your book."}</p>
            </div>
          </div>
          ${isAdmin()
            ? `
              <div class="action-stack">
                <button class="button button-primary" id="exportWorkbookButton" type="button">Export Workbook</button>
                <p class="mini-note">Downloads a multi-sheet workbook aligned to the original spreadsheet tabs.</p>
              </div>
            `
            : renderCommissionList(repCommissionRows)}
        </article>
      </div>
      <div class="table-card">
        <div class="panel-header">
          <div>
            <h3>${isAdmin() ? "Upcoming Renewals" : "My Upcoming Renewals"}</h3>
            <p>${isAdmin() ? "Accounts nearing expiration so the team can retain before they slip." : "Accounts you should work before renewal deadlines hit."}</p>
          </div>
        </div>
        ${renderRenewalQueue(renewalRows)}
      </div>
    </section>
    ` : ""}

    ${state.ui.activeTab === "integrations" ? renderIntegrations() : ""}

    ${isAdmin() && state.ui.activeTab === "reports" ? `
    <section class="panel workspace-panel" id="reports">
      <div class="panel-header">
        <div>
          <h2>Owner Reports</h2>
          <p>Rep performance and source profitability in one place without crowding the main dashboard.</p>
        </div>
      </div>
      <div class="two-column compact-two-column">
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>Rep Performance Snapshot</h3>
              <p>See who is converting, who is overdue, and where coaching pressure belongs.</p>
            </div>
          </div>
          ${renderOwnerRepPerformance(ownerRepPerformanceRows)}
        </article>
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>Lead Source Profitability</h3>
              <p>Revenue, spend, and return by source so owners can decide where to double down.</p>
            </div>
          </div>
          ${renderOwnerSourcePerformance(ownerSourcePerformanceRows)}
        </article>
      </div>
    </section>
    ` : ""}

    ${state.ui.activeTab === "opportunities" ? `
    <section class="panel workspace-panel" id="opportunities">
      <div class="panel-header">
        <div>
          <h2>${isAdmin() ? "Master Opportunity Log" : "My Opportunity Log"}</h2>
          <p>${isAdmin() ? "Separate create, update, and stage movement into cleaner working modes." : "Work new leads, updates, and stage changes in dedicated tabs instead of one crowded screen."}</p>
        </div>
      </div>
      <div class="section-tabs" role="tablist" aria-label="Pipeline work modes">
        ${getOpportunityTabs()
          .map(
            (tab) => `
              <button
                class="section-tab ${state.ui.opportunityTab === tab.id ? "is-active" : ""}"
                data-opportunity-tab="${tab.id}"
                type="button"
              >
                ${escapeHtml(tab.label)}
              </button>
            `
          )
          .join("")}
      </div>
      <div class="table-card filter-card">
        <div class="toolbar">
          <div class="toolbar-grid">
            <label>
              Search
              <input id="searchInput" value="${escapeHtml(state.ui.search)}" placeholder="Business, rep, lead number" />
            </label>
            ${isAdmin() ? `
              <label>
                Rep
                <select id="repFilterSelect">
                  <option value="All">All reps</option>
                  ${visibleManagedProfiles.map((profile) => `<option value="${profile.id}" ${state.ui.repFilter === profile.id ? "selected" : ""}>${escapeHtml(profile.full_name)}</option>`).join("")}
                </select>
              </label>
            ` : ""}
            <label>
              Lead Source
              <select id="sourceFilterSelect">
                <option value="All">All sources</option>
                ${state.setup.leadSources.map((source) => `<option value="${escapeHtml(source)}" ${state.ui.sourceFilter === source ? "selected" : ""}>${escapeHtml(source)}</option>`).join("")}
              </select>
            </label>
            <label>
              Status
              <select id="statusFilterSelect">
                <option value="All">All statuses</option>
                ${state.setup.statuses.map((status) => `<option value="${escapeHtml(status)}" ${state.ui.statusFilter === status ? "selected" : ""}>${escapeHtml(status)}</option>`).join("")}
              </select>
            </label>
            <label>
              From
              <input id="dateFromInput" type="date" value="${escapeHtml(state.ui.dateFrom)}" />
            </label>
            <label>
              To
              <input id="dateToInput" type="date" value="${escapeHtml(state.ui.dateTo)}" />
            </label>
          </div>
          <div class="toolbar-group">
            <button class="button button-ghost" id="clearFiltersButton" type="button">Clear Filters</button>
            <div class="subtle">${listRows.length} visible</div>
          </div>
        </div>
        ${isAdmin() && state.ui.opportunityTab === "update" ? `
          <div class="bulk-card">
            <div>
              <strong>${state.ui.selectedOpportunityIds.length} selected</strong>
              <div class="subtle">Use filters, select leads, then bulk reassign them.</div>
            </div>
            <div class="toolbar-group">
              <button class="button button-ghost" id="selectVisibleButton" type="button">Select Visible</button>
              <button class="button button-ghost" id="clearSelectionButton" type="button">Clear Selection</button>
              <label class="compact-field">
                Reassign To
                <select id="bulkAssignUserSelect">
                  <option value="">Choose rep</option>
                  ${assignableProfiles.map((profile) => `<option value="${profile.id}" ${state.ui.bulkAssignUserId === profile.id ? "selected" : ""}>${escapeHtml(profile.full_name)}</option>`).join("")}
                </select>
              </label>
              <button class="button button-primary" id="bulkAssignButton" type="button">Bulk Reassign</button>
            </div>
          </div>
        ` : ""}
        ${state.ui.opportunityTab === "stage" ? `
          <div class="stage-strip">
            ${pipelineGroups.map((group) => `
              <article class="stage-pill-card">
                <h4>${escapeHtml(group.status)}</h4>
                <strong>${group.rows.length}</strong>
              </article>
            `).join("")}
          </div>
        ` : ""}
      </div>
      ${state.ui.opportunityTab === "create" ? `
        ${renderCreateLeadWorkspace()}
      ` : ""}
      ${state.ui.opportunityTab === "update" ? `
        <div class="two-column">
          <div class="table-card">
            <div class="table-wrap">
              ${listRows.length ? renderOpportunityTable(listRows, getSelectedOpportunitySet()) : document.getElementById("emptyStateTemplate").innerHTML}
            </div>
          </div>
          ${renderLeadWorkspace(activeOpportunity, activeOpportunityTimeline)}
        </div>
      ` : ""}
      ${state.ui.opportunityTab === "stage" ? `
        <div class="table-card board-host">
          ${renderOpportunityBoard(pipelinePhaseGroups)}
        </div>
      ` : ""}
    </section>
    ` : ""}

    ${state.ui.activeTab === "scorecards" ? `
    <section class="panel workspace-panel" id="scorecards">
      <div class="panel-header">
        <div>
          <h2>${isAdmin() ? "Agency Scorecards and ROI" : "My Production Scorecard"}</h2>
          <p>${isAdmin() ? "Full rollup by rep and lead source." : "Your personal performance against the same tracked metrics."}</p>
        </div>
      </div>
        <div class="two-column">
        <div class="table-card">
          <div class="panel-header"><h3>${isAdmin() ? "Rep Scorecards" : "My Production Scorecard"}</h3></div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>${isAdmin() ? "Rep" : "Producer"}</th>
                  <th>Leads</th>
                  <th>Same Day</th>
                  <th>Contact</th>
                  <th>Quote</th>
                  <th>Q-B</th>
                  <th>${isAdmin() ? "Agency Comm" : "My Comm"}</th>
                </tr>
              </thead>
              <tbody>
                ${scorecards.map((row) => `
                  <tr>
                    <td>${escapeHtml(row.name)}</td>
                    <td>${row.leads}</td>
                    <td>${formatPct(row.sameDayRate)}</td>
                    <td>${formatPct(row.contactRate)}</td>
                    <td>${formatPct(row.quoteRate)}</td>
                    <td>${formatPct(row.quoteToBindRate)}</td>
                    <td>${formatCurrency(isAdmin() ? row.actualAgencyComm : row.actualRepPayout)}</td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
        <div class="table-card">
          <div class="panel-header"><h3>${isAdmin() ? "Lead Source ROI" : "My Lead Source Mix"}</h3></div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Lead Source</th>
                  <th>Count</th>
                  ${isAdmin() ? "<th>Spend</th>" : ""}
                  <th>Quote Rate</th>
                  <th>Bind Rate</th>
                  <th>${isAdmin() ? "Actual Comm" : "My Comm"}</th>
                </tr>
              </thead>
              <tbody>
                ${roiRows.map((row) => `
                  <tr>
                    <td>${escapeHtml(row.source)}</td>
                    <td>${row.count}</td>
                    ${isAdmin() ? `<td>${formatCurrency(row.spend)}</td>` : ""}
                    <td>${formatPct(row.quoteRate)}</td>
                    <td>${formatPct(row.bindRate)}</td>
                    <td>${formatCurrency(isAdmin() ? row.actualAgencyComm : row.actualRepPayout)}</td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </section>
    ` : ""}

    ${state.ui.activeTab === "coaching" ? `
    <section class="panel workspace-panel" id="coaching">
      <div class="panel-header">
        <div>
          <h2>${isAdmin() ? "Weekly Coaching" : "My Coaching Notes"}</h2>
          <p>${isAdmin() ? "Managers can keep coaching commitments next to the production data." : "View current coaching focus and action commitments."}</p>
        </div>
      </div>
      <div class="table-card">
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Week Start</th>
                <th>Rep</th>
                <th>Leads</th>
                <th>Contact Rate</th>
                <th>Quotes</th>
                <th>Binds</th>
                <th>Biggest Gap</th>
                <th>Behavior</th>
                <th>Commitment</th>
                <th>Next Review Notes</th>
              </tr>
            </thead>
            <tbody>
              ${coachingRows.map((row) => `
                <tr>
                  <td>${escapeHtml(row.weekStart)}</td>
                  <td>${escapeHtml(row.repName)}</td>
                  <td>${row.leads}</td>
                  <td>${formatPct(row.contactRate)}</td>
                  <td>${row.quotes}</td>
                  <td>${row.binds}</td>
                  <td>${coachingInput(row, "biggestGap", row.biggestGap)}</td>
                  <td>${coachingInput(row, "behaviorToImprove", row.behaviorToImprove)}</td>
                  <td>${coachingInput(row, "actionCommitment", row.actionCommitment)}</td>
                  <td>${coachingInput(row, "nextReviewNotes", row.nextReviewNotes)}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </div>
    </section>
    ` : ""}

    ${isAdmin() && state.ui.activeTab === "setup" ? `
      <section class="panel workspace-panel" id="setup">
        <div class="panel-header">
          <div>
            <h2>Admin Setup</h2>
            <p>Only the admin can manage global assumptions, producer roster, and dropdown lists.</p>
          </div>
        </div>
        <div class="section-tabs" role="tablist" aria-label="Admin setup sections">
          ${getSetupTabs()
            .map(
              (tab) => `
                <button
                  class="section-tab ${state.ui.setupTab === tab.id ? "is-active" : ""}"
                  data-setup-tab="${tab.id}"
                  type="button"
                >
                  ${escapeHtml(tab.label)}
                </button>
              `
            )
            .join("")}
        </div>

        ${state.ui.setupTab === "users" ? `
          <article class="table-card">
            <div class="panel-header">
              <div>
                <h3>Users</h3>
                <p>Invite reps directly from the app, then manage role and status once they activate.</p>
              </div>
            </div>
            <form id="inviteRepForm" class="invite-form">
              <label>
                Rep Name
                <input name="fullName" placeholder="Producer name" required />
              </label>
              <label>
                Rep Email
                <input name="email" type="email" placeholder="rep@agency.com" required />
              </label>
              <button class="button button-primary" type="submit">Send Invite Email</button>
            </form>
            <div class="table-wrap users-table-wrap">
              <table class="users-table">
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>Email</th>
                    <th>Role</th>
                    <th>Status</th>
                    <th>Assigned Leads</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  ${visibleManagedProfiles.map((profile) => `
                    <tr>
                      <td class="users-table-name"><input data-profile-name="${profile.id}" value="${escapeHtml(profile.full_name)}" /></td>
                      <td class="users-table-email">${escapeHtml(profile.email || "")}</td>
                      <td class="users-table-select">
                        <select data-profile-role="${profile.id}">
                          <option value="rep" ${profile.role === "rep" ? "selected" : ""}>rep</option>
                          <option value="admin" ${profile.role === "admin" ? "selected" : ""}>admin</option>
                        </select>
                      </td>
                      <td class="users-table-select">
                        <select data-profile-active="${profile.id}">
                          <option value="true" ${profile.active ? "selected" : ""}>Active</option>
                          <option value="false" ${!profile.active ? "selected" : ""}>Inactive</option>
                        </select>
                      </td>
                      <td class="users-table-count">${getAssignedLeadCount(profile.id)} leads</td>
                      <td class="users-table-actions">
                        ${profile.id === state.profile.id ? '<span class="subtle">Current admin</span>' : '<button class="button button-ghost" type="button" data-remove-user="' + profile.id + '">Remove</button>'}
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
            <p class="notice">Invite emails create rep accounts using Supabase email sign-in. Deactivate only after their assigned leads have been reassigned.</p>
            ${removedProfiles.length ? `
              <div class="archived-users">
                <h4>Removed Users</h4>
                <div class="table-wrap users-table-wrap">
                  <table class="users-table users-table-removed">
                    <thead>
                      <tr>
                        <th>Name</th>
                        <th>Email</th>
                        <th>Assigned Leads</th>
                        <th>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${removedProfiles.map((profile) => `
                        <tr>
                          <td>${escapeHtml(profile.full_name)}</td>
                          <td>${escapeHtml(profile.email || "")}</td>
                          <td>${getAssignedLeadCount(profile.id)} leads</td>
                          <td><button class="button button-ghost" type="button" data-restore-user="${profile.id}">Restore</button></td>
                        </tr>
                      `).join("")}
                    </tbody>
                  </table>
                </div>
              </div>
            ` : ""}
          </article>
        ` : ""}

        ${state.ui.setupTab === "routing" ? `
          <div class="section-stack">
            <article class="table-card">
              <div class="panel-header">
                <div>
                  <h3>Assignment Automation</h3>
                  <p>Control whether new leads route automatically and how the app picks the assignee.</p>
                </div>
              </div>
              <div class="two-column">
                <label class="mini-card">
                  Auto Assign New Leads
                  <select id="routingAutoAssignSelect">
                    <option value="false" ${!state.setup.routingRules.autoAssignEnabled ? "selected" : ""}>Manual selection</option>
                    <option value="true" ${state.setup.routingRules.autoAssignEnabled ? "selected" : ""}>Enabled</option>
                  </select>
                </label>
                <label class="mini-card">
                  Routing Mode
                  <select id="routingModeSelect">
                    <option value="round_robin" ${state.setup.routingRules.mode === "round_robin" ? "selected" : ""}>Round robin</option>
                    <option value="source_rule" ${state.setup.routingRules.mode === "source_rule" ? "selected" : ""}>Source based</option>
                  </select>
                </label>
              </div>
              <p class="notice">Rep-created leads stay with the rep who entered them. Self-generated leads are never routed away from the creator. Auto assign only applies to owner-entered or imported leads left on <code>Auto Assign</code>.</p>
            </article>
            <article class="table-card">
              <div class="panel-header">
                <div>
                  <h3>Lead Source Routing</h3>
                  <p>Optionally pin owner-fed lead sources to specific reps before round robin is used.</p>
                </div>
              </div>
              <div class="table-wrap">
                <table class="settings-table">
                  <thead>
                    <tr>
                      <th>Lead Source</th>
                      <th>Assigned Rep</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${state.setup.leadSources.map((source) => {
                      const rule = state.setup.routingRules.sourceRules.find((item) => item.source === source);
                      const sourceLocked = source === "Self-Generated";
                      return `
                        <tr>
                          <td>${escapeHtml(source)}</td>
                          <td>
                            ${sourceLocked ? `
                              <div class="routing-static-note">Creator keeps ownership</div>
                            ` : `
                              <select data-routing-source="${escapeHtml(source)}">
                                <option value="">Use routing mode default</option>
                                ${assignableProfiles.map((profile) => `<option value="${profile.id}" ${rule?.userId === profile.id ? "selected" : ""}>${escapeHtml(profile.full_name)}</option>`).join("")}
                              </select>
                            `}
                          </td>
                        </tr>
                      `;
                    }).join("")}
                  </tbody>
                </table>
              </div>
            </article>
            <article class="table-card">
              <div class="panel-header">
                <div>
                  <h3>Bulk Lead Import</h3>
                  <p>Upload a CSV from Excel or Google Sheets, preview assignments, and import the batch without manual re-entry.</p>
                </div>
                <button class="button button-ghost" id="downloadLeadImportTemplateButton" type="button">Download CSV Template</button>
              </div>
              <div class="import-controls">
                <label class="mini-card">
                  Lead File
                  <input id="leadImportFileInput" type="file" accept=".csv,.xlsx,.xls" />
                </label>
                <div class="mini-card import-status-card">
                  <strong>${state.ui.importFileName || "No file selected"}</strong>
                  <div class="subtle">${state.ui.importPreviewRows.length ? `${state.ui.importPreviewRows.filter((row) => !row.errors.length).length} ready · ${state.ui.importPreviewRows.filter((row) => row.errors.length).length} with errors` : "Upload a file to preview routing before import."}</div>
                </div>
                <div class="import-actions">
                  <button class="button button-primary" id="importLeadsButton" type="button" ${state.ui.importPreviewRows.length && !state.ui.importPreviewRows.some((row) => row.errors.length) && !state.ui.importingCsv ? "" : "disabled"}>
                    ${state.ui.importingCsv ? "Importing..." : "Import Leads"}
                  </button>
                  <button class="button button-ghost" id="clearLeadImportButton" type="button" ${state.ui.importPreviewRows.length ? "" : "disabled"}>Clear Preview</button>
                </div>
              </div>
              ${state.ui.importPreviewRows.length ? `
                <div class="table-wrap">
                  <table class="settings-table import-preview-table">
                    <thead>
                      <tr>
                        <th>Row</th>
                        <th>Business Name</th>
                        <th>Source</th>
                        <th>Assigned Rep</th>
                        <th>Routing Result</th>
                        <th>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${state.ui.importPreviewRows.map((row) => `
                        <tr>
                          <td>${row.rowNumber}</td>
                          <td>${escapeHtml(row.data.businessName || "Missing business name")}</td>
                          <td>${escapeHtml(row.data.leadSource || "Unknown")}</td>
                          <td>${escapeHtml(row.assigneeName)}</td>
                          <td>${escapeHtml(row.strategy || "Ready")}</td>
                          <td>
                            ${row.errors.length
                              ? `<div class="import-error-list">${row.errors.map((item) => `<div>${escapeHtml(item)}</div>`).join("")}</div>`
                              : '<span class="tag">Ready</span>'}
                          </td>
                        </tr>
                      `).join("")}
                    </tbody>
                  </table>
                </div>
              ` : ""}
            </article>
            <article class="table-card">
              <div class="panel-header">
                <div>
                  <h3>Rep Workload</h3>
                  <p>See open-load distribution before you rebalance or change routing rules.</p>
                </div>
              </div>
              <div class="workload-grid">
                ${workloadRows.map((row) => `
                  <article class="mini-card workload-card">
                    <strong>${escapeHtml(row.name)}</strong>
                    <div class="workload-metrics">
                      <span>${row.openLeads} open</span>
                      <span>${row.overdue} overdue</span>
                      <span>${row.stale} stale</span>
                      <span>${row.quoted} quoted</span>
                      <span>${row.renewalsDue} renewals due</span>
                    </div>
                  </article>
                `).join("")}
              </div>
              ${routingSuggestions.length ? `
                <div class="routing-hints">
                  ${routingSuggestions.map((tip) => `<p class="mini-note">${escapeHtml(tip)}</p>`).join("")}
                </div>
              ` : `<p class="mini-note">Workload is reasonably balanced right now.</p>`}
            </article>
          </div>
        ` : ""}

        ${state.ui.setupTab === "assumptions" ? `
          <article class="table-card">
            <div class="panel-header">
              <div>
                <h3>Agency Assumptions</h3>
                <p>Locked by default so forecasting and payout assumptions are not changed accidentally.</p>
              </div>
              <button class="button ${state.ui.assumptionEditing ? "button-secondary" : "button-ghost"}" id="toggleAssumptionEditingButton" type="button">
                ${state.ui.assumptionEditing ? "Done Editing" : "Edit Assumptions"}
              </button>
            </div>
            <div class="four-column compact-four-column">
              ${Object.entries(state.setup.assumptions).map(([key, value]) => `
                <label class="mini-card">
                  ${humanize(key)}
                  <input data-assumption="${key}" type="number" step="0.01" value="${value}" ${state.ui.assumptionEditing ? "" : "disabled"} />
                </label>
              `).join("")}
            </div>
          </article>
        ` : ""}

        ${state.ui.setupTab === "carriers" ? `
          <article class="table-card">
            <div class="panel-header">
              <div>
                <h3>Carrier Commission Table</h3>
                <p>Locked by default so the commission model is not changed accidentally.</p>
              </div>
              <button class="button ${state.ui.carrierEditing ? "button-secondary" : "button-ghost"}" id="toggleCarrierEditingButton" type="button">
                ${state.ui.carrierEditing ? "Done Editing" : "Edit Table"}
              </button>
            </div>
            <div class="table-wrap">
              <table class="settings-table">
                <thead>
                  <tr>
                    <th>Carrier</th>
                    <th>New %</th>
                    <th>Renewal %</th>
                  </tr>
                </thead>
                <tbody>
                  ${state.setup.carriers.map((carrier, index) => `
                    <tr>
                      <td><input data-carrier-name="${index}" value="${escapeHtml(carrier.name)}" ${state.ui.carrierEditing ? "" : "disabled"} /></td>
                      <td><input data-carrier-new="${index}" type="number" step="0.01" value="${Number(carrier.newPct || 0)}" ${state.ui.carrierEditing ? "" : "disabled"} /></td>
                      <td><input data-carrier-renewal="${index}" type="number" step="0.01" value="${Number(carrier.renewalPct || 0)}" ${state.ui.carrierEditing ? "" : "disabled"} /></td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </article>
        ` : ""}
      </section>
    ` : ""}
  `;

  bindShellEvents();
  bindAppEvents();
}

function renderHeroActions() {
  if (!SUPABASE_READY) {
    return `<a class="button button-primary" href="#app">Setup Backend</a>`;
  }
  if (!state.session) {
    return "";
  }
  return `
    <div class="auth-summary">
      <div>
        <strong>${escapeHtml(state.profile?.full_name || state.session.user.email || "User")}</strong>
        <div class="subtle">${isInactiveUser() ? "Inactive account" : isAdmin() ? "Admin / owner view" : "Producer view"}</div>
      </div>
      <button id="signOutButton" class="button button-ghost" type="button">Sign Out</button>
    </div>
  `;
}

function renderTopNav() {
  if (!SUPABASE_READY || !state.session) {
    return `<a href="#app">Access</a>`;
  }
  if (isInactiveUser()) {
    return `<a href="#app">Account Status</a>`;
  }
  return getAvailableTabs()
    .map(
      (tab) => `
        <button
          class="tab-button ${state.ui.activeTab === tab.id ? "is-active" : ""}"
          data-app-tab="${tab.id}"
          type="button"
        >
          ${escapeHtml(tab.label)}
        </button>
      `
    )
    .join("");
}

function renderAssistant() {
  if (!SUPABASE_READY || !state.session || isInactiveUser()) {
    return "";
  }

  const activeOpportunity =
    state.opportunities.find((item) => item.id === state.ui.activeOpportunityId) ||
    null;

  return `
    <div class="assistant-shell ${state.ui.assistantOpen ? "is-open" : ""}">
      ${state.ui.assistantOpen ? `
        <section class="assistant-panel">
          <header class="assistant-header">
            <div>
              <strong>Claude Assistant</strong>
              <div class="subtle">${isAdmin() ? "Admin can ask across the agency." : "You can ask about your own leads and metrics."}</div>
            </div>
            <button class="button button-ghost assistant-close" id="assistantCloseButton" type="button">Close</button>
          </header>
          <div class="assistant-context">
            <span class="tag">${isAdmin() ? "Admin scope" : "Rep scope"}</span>
            ${activeOpportunity ? `<span class="tag">Focused lead: ${escapeHtml(activeOpportunity.businessName || activeOpportunity.leadNumber)}</span>` : `<span class="tag">Workspace scope</span>`}
          </div>
          <div class="assistant-messages">
            ${state.ui.assistantMessages.length ? state.ui.assistantMessages.map((message) => `
              <article class="assistant-message ${message.role === "user" ? "is-user" : "is-assistant"}">
                <strong>${message.role === "user" ? "You" : "Claude"}</strong>
                <p>${escapeHtml(message.content)}</p>
              </article>
            `).join("") : `
              <div class="empty-state assistant-empty">
                <h3>Ask Claude about the book</h3>
                <p>${isAdmin() ? "Try: Which reps have the most stale leads, what renewals are due next, or which source is performing best?" : "Try: What should I work first, which leads are overdue, or what renewals are coming due in my book?"}</p>
              </div>
            `}
          </div>
          ${state.ui.assistantError ? `<p class="error-banner">${escapeHtml(state.ui.assistantError)}</p>` : ""}
          <form id="assistantForm" class="assistant-form">
            <textarea id="assistantInput" placeholder="Ask about leads, metrics, renewals, attachments, or appointments...">${escapeHtml(state.ui.assistantInput)}</textarea>
            <div class="assistant-actions">
              <button class="button button-primary" type="submit" ${state.ui.assistantLoading ? "disabled" : ""}>${state.ui.assistantLoading ? "Thinking..." : "Ask Claude"}</button>
            </div>
          </form>
        </section>
      ` : ""}
      <button class="assistant-bubble button button-primary" id="assistantBubbleButton" type="button">
        ${state.ui.assistantOpen ? "Hide Claude" : "Ask Claude"}
      </button>
    </div>
  `;
}

function renderSetupRequired() {
  return `
    <section class="panel">
      <div class="panel-header">
        <div>
          <h2>Hosted Backend Needed</h2>
          <p>This version is wired for shared rep logins and a master agency view, but it needs Supabase keys before it can run.</p>
        </div>
      </div>
      <div class="three-column">
        <article class="table-card">
          <h3>1. Create Supabase Project</h3>
          <p class="mini-note">Create a project, open the SQL editor, and run the schema in <code>supabase-schema.sql</code>.</p>
        </article>
        <article class="table-card">
          <h3>2. Add Config</h3>
          <p class="mini-note">Add <code>VITE_SUPABASE_URL</code> and <code>VITE_SUPABASE_PUBLISHABLE_KEY</code> in your environment variables.</p>
        </article>
        <article class="table-card">
          <h3>3. Invite Reps</h3>
          <p class="mini-note">Create rep and admin accounts in Supabase Auth. The SQL trigger auto-creates their profiles.</p>
        </article>
      </div>
      <p class="notice">Once those two keys are added, this app becomes the shared web app you described: rep login, rep-owned pipelines, and an admin-wide command center.</p>
    </section>
  `;
}

function renderLogin() {
  return `
    <section class="panel">
      <div class="panel-header">
        <div>
          <h2>Agency Login</h2>
          <p>Each producer signs in to their own workspace. The owner sees the full agency operation.</p>
        </div>
      </div>
      ${state.ui.error ? `<p class="error-banner">${escapeHtml(state.ui.error)}</p>` : ""}
      ${state.ui.notice ? `<p class="notice-banner">${escapeHtml(state.ui.notice)}</p>` : ""}
      <div class="auth-panel">
        <form id="loginForm" class="form-card auth-form" novalidate onsubmit="return false;">
          <label>
            Email
            <input type="email" name="email" placeholder="rep@agency.com" required />
          </label>
          <label>
            Password
            <input type="password" name="password" placeholder="••••••••" required />
          </label>
          <button class="button button-primary" id="loginButton" type="button">${state.ui.authLoading ? "Signing In..." : "Sign In"}</button>
          <button class="button button-ghost" id="resetPasswordButton" type="button">Send Password Reset Email</button>
        </form>
        <article class="table-card">
          <h3>How access works</h3>
          <ul>
            <li>Reps can only view and edit their own assigned leads.</li>
            <li>Admins can view every lead, assign reps, and adjust agency-wide settings.</li>
            <li>All producer activity rolls up live into the owner dashboard.</li>
          </ul>
        </article>
      </div>
    </section>
  `;
}

function renderInactiveUser() {
  return `
    <section class="panel">
      <div class="empty-state">
        <h3>Account Inactive</h3>
        <p>Your account has been deactivated by the admin team. Contact the agency owner if you need access restored.</p>
      </div>
    </section>
  `;
}

function renderRecoveryState() {
  return `
    <section class="panel">
      <div class="empty-state">
        <h3>Session Needs Refresh</h3>
        <p>The app updated or your saved browser session fell out of sync. You can reset the local session without manually clearing site data.</p>
        <div class="form-actions recovery-actions">
          <button class="button button-primary" id="resetSessionButton" type="button">Reset Session</button>
          <button class="button button-ghost" id="reloadAppButton" type="button">Reload App</button>
        </div>
      </div>
    </section>
  `;
}

function renderIntegrations() {
  const googleConnected = Boolean(state.calendarConnection?.connected);
  const emailReady = Boolean(state.communicationStatus?.email?.configured);
  const smsReady = Boolean(state.communicationStatus?.sms?.configured);
  return `
    <section class="panel workspace-panel" id="integrations">
      <div class="panel-header">
        <div>
          <h2>Integrations</h2>
          <p>Connect tools that support day-to-day producer work and owner visibility.</p>
        </div>
      </div>
      <div class="two-column compact-two-column">
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>Google Calendar</h3>
              <p>Push follow-ups and appointments to your Google calendar directly from a lead.</p>
            </div>
            <span class="tag">${googleConnected ? "Connected" : "Not Connected"}</span>
          </div>
          <div class="action-stack">
            <div class="mini-note">
              ${googleConnected
                ? `Connected as ${escapeHtml(state.calendarConnection.email || "Google account")}.`
                : "Connect your Google account first, then reps and admins can create calendar events from lead follow-ups and appointments."}
            </div>
            <div class="form-actions">
              <button class="button button-primary" id="connectGoogleCalendarButton" type="button" ${state.ui.calendarSyncLoading || googleConnected ? "disabled" : ""}>
                ${state.ui.calendarSyncLoading && !googleConnected ? "Connecting..." : "Connect Google Calendar"}
              </button>
              <button class="button button-ghost" id="disconnectGoogleCalendarButton" type="button" ${state.ui.calendarSyncLoading || !googleConnected ? "disabled" : ""}>
                ${state.ui.calendarSyncLoading && googleConnected ? "Disconnecting..." : "Disconnect"}
              </button>
            </div>
            <p class="notice">This connection is user-specific. Each rep can connect their own calendar, while admins can connect theirs separately.</p>
          </div>
        </article>
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>Outlook Calendar</h3>
              <p>Microsoft calendar sync is next on deck after Google.</p>
            </div>
            <span class="tag">Coming Next</span>
          </div>
          <div class="empty-state">
            <h3>Google first</h3>
            <p>We started with Google Calendar because it is the most common fit for your current team. Outlook will use the same role-safe pattern after this.</p>
          </div>
        </article>
      </div>
      <div class="two-column compact-two-column">
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>Email and Text Delivery</h3>
              <p>Provider-backed reminder sending for follow-ups and nudges.</p>
            </div>
          </div>
          <div class="dashboard-grid compact-dashboard-grid">
            ${statCard("Email", emailReady ? "Ready" : "Not Ready", emailReady ? "Resend configured" : "Add provider secrets")}
            ${statCard("SMS", smsReady ? "Ready" : "Not Ready", smsReady ? "Twilio configured" : "Add provider secrets")}
          </div>
          <p class="notice">Email reminders use the lead’s contact email. SMS reminders use the lead’s contact phone.</p>
        </article>
        <article class="table-card">
          <div class="panel-header">
            <div>
              <h3>Reminder Templates</h3>
              <p>${isAdmin() ? "Edit the default message templates reps use for one-click reminders." : "These are the agency-approved reminder templates used when you send follow-ups."}</p>
            </div>
            ${isAdmin() ? `<button class="button ${state.ui.reminderEditing ? "button-secondary" : "button-ghost"}" id="toggleReminderEditingButton" type="button">${state.ui.reminderEditing ? "Done Editing" : "Edit Templates"}</button>` : ""}
          </div>
          <div class="section-stack">
            <label class="mini-card">
              Email Subject
              <input data-communication-setting="emailSubjectTemplate" value="${escapeHtml(state.setup.communicationSettings.emailSubjectTemplate || "")}" ${isAdmin() && state.ui.reminderEditing ? "" : "disabled"} />
            </label>
            <label class="mini-card">
              Reply-To Email
              <input data-communication-setting="replyToEmail" value="${escapeHtml(state.setup.communicationSettings.replyToEmail || "")}" ${isAdmin() && state.ui.reminderEditing ? "" : "disabled"} />
            </label>
            <label class="mini-card">
              Email Body
              <textarea data-communication-setting="emailBodyTemplate" ${isAdmin() && state.ui.reminderEditing ? "" : "disabled"}>${escapeHtml(state.setup.communicationSettings.emailBodyTemplate || "")}</textarea>
            </label>
            <label class="mini-card">
              SMS Body
              <textarea data-communication-setting="smsBodyTemplate" ${isAdmin() && state.ui.reminderEditing ? "" : "disabled"}>${escapeHtml(state.setup.communicationSettings.smsBodyTemplate || "")}</textarea>
            </label>
          </div>
          <p class="notice">Available placeholders: <code>{{contactName}}</code>, <code>{{businessName}}</code>, <code>{{repName}}</code>, <code>{{leadNumber}}</code>, <code>{{nextTaskSentence}}</code>, <code>{{nextFollowUpDate}}</code>.</p>
        </article>
      </div>
    </section>
  `;
}

function renderOpportunityForm(row) {
  const autoAssignLabel = state.setup.routingRules.autoAssignEnabled ? "Auto Assign" : "Keep with owner";
  const assigneeOptions = [
    ...(isAdmin() ? [`<option value="auto" ${(!row.id && !row.assignedUserId) || row.assignedUserId === "auto" ? "selected" : ""}>${autoAssignLabel}</option>`] : []),
    ...(isAdmin() ? state.profiles.filter((item) => item.active) : [state.profile])
      .map((profile) => `<option value="${profile.id}" ${row.assignedUserId === profile.id ? "selected" : ""}>${escapeHtml(profile.full_name)}</option>`)
  ].join("");

  return `
    <form id="opportunityForm">
      <input type="hidden" name="id" value="${escapeHtml(row.id || "")}" />
      <input type="hidden" name="leadNumber" value="${escapeHtml(row.leadNumber || "")}" />
      <div class="workspace-form-section">
        <div class="workspace-section-header">
          <h4>Lead Profile</h4>
          <p>Core account details and contact information.</p>
        </div>
        <div class="form-grid">
        <label>
          Date Received
          <input type="date" name="dateReceived" value="${escapeHtml(row.dateReceived || todayIso())}" required />
        </label>
        <label>
          Assigned Rep
          <select name="assignedUserId">${assigneeOptions}</select>
        </label>
        <label>
          Lead Source
          <select name="leadSource">${state.setup.leadSources.map((item) => `<option ${row.leadSource === item ? "selected" : ""}>${escapeHtml(item)}</option>`).join("")}</select>
        </label>
        <label>
          Business Name
          <input name="businessName" value="${escapeHtml(row.businessName || "")}" required />
        </label>
        <label>
          Target Niche
          <input name="targetNiche" value="${escapeHtml(row.targetNiche || "")}" />
        </label>
        <label>
          Product Focus
          <select name="productFocus">${state.setup.products.map((item) => `<option ${row.productFocus === item ? "selected" : ""}>${escapeHtml(item)}</option>`).join("")}</select>
        </label>
        <label>
          Contact Name
          <input name="contactName" value="${escapeHtml(row.contactName || "")}" />
        </label>
        <label>
          Contact Email
          <input type="email" name="contactEmail" value="${escapeHtml(row.contactEmail || "")}" />
        </label>
        <label>
          Contact Phone
          <input name="contactPhone" value="${escapeHtml(row.contactPhone || "")}" />
        </label>
        </div>
      </div>
      <div class="workspace-form-section">
        <div class="workspace-section-header">
          <h4>Workflow</h4>
          <p>Stage, reminders, and next actions that keep the lead moving.</p>
        </div>
        <div class="form-grid">
        <label>
          Carrier
          <select name="carrier">${state.setup.carriers.map((carrier) => `<option ${row.carrier === carrier.name ? "selected" : ""}>${escapeHtml(carrier.name)}</option>`).join("")}</select>
        </label>
        <label>
          Incumbent Carrier
          <input name="incumbentCarrier" value="${escapeHtml(row.incumbentCarrier || "")}" />
        </label>
        <label>
          Policy Type
          <select name="policyType">
            <option value="New" ${row.policyType === "New" ? "selected" : ""}>New</option>
            <option value="Renewal" ${row.policyType === "Renewal" ? "selected" : ""}>Renewal</option>
          </select>
        </label>
        <label>
          Renewal Status
          <select name="renewalStatus">
            ${getRenewalStatuses().map((status) => `<option value="${escapeHtml(status)}" ${row.renewalStatus === status ? "selected" : ""}>${escapeHtml(status)}</option>`).join("")}
          </select>
        </label>
        <label>
          Policy Term
          <select name="policyTermMonths">
            <option value="12" ${Number(row.policyTermMonths || 12) === 12 ? "selected" : ""}>12 months</option>
            <option value="6" ${Number(row.policyTermMonths || 12) === 6 ? "selected" : ""}>6 months</option>
            <option value="3" ${Number(row.policyTermMonths || 12) === 3 ? "selected" : ""}>3 months</option>
          </select>
        </label>
        <label>
          Effective Date
          <input type="date" name="effectiveDate" value="${escapeHtml(row.effectiveDate || "")}" />
        </label>
        <label>
          Expiration Date
          <input type="date" name="expirationDate" value="${escapeHtml(row.expirationDate || "")}" />
        </label>
        <label>
          Next Task
          <input name="nextTask" value="${escapeHtml(row.nextTask || "")}" placeholder="Call back, quote follow-up, collect docs" />
        </label>
        <label>
          Task Priority
          <select name="taskPriority">
            <option value="High" ${row.taskPriority === "High" ? "selected" : ""}>High</option>
            <option value="Medium" ${!row.taskPriority || row.taskPriority === "Medium" ? "selected" : ""}>Medium</option>
            <option value="Low" ${row.taskPriority === "Low" ? "selected" : ""}>Low</option>
          </select>
        </label>
        <label>
          Status
          <select name="status">${state.setup.statuses.map((item) => `<option ${row.status === item ? "selected" : ""}>${escapeHtml(item)}</option>`).join("")}</select>
        </label>
        <label>
          First Attempt Date
          <input type="date" name="firstAttemptDate" value="${escapeHtml(row.firstAttemptDate || "")}" />
        </label>
        <label>
          Last Activity Date
          <input type="date" name="lastActivityDate" value="${escapeHtml(row.lastActivityDate || "")}" />
        </label>
        <label>
          Next Follow-Up Date
          <input type="date" name="nextFollowUpDate" value="${escapeHtml(row.nextFollowUpDate || "")}" />
        </label>
        <label>
          Lead Cost
          <input type="number" step="0.01" name="leadCost" value="${Number(row.leadCost || 0)}" />
        </label>
        <label>
          Premium Quoted
          <input type="number" step="0.01" name="premiumQuoted" value="${Number(row.premiumQuoted || 0)}" />
        </label>
        <label>
          Premium Bound
          <input type="number" step="0.01" name="premiumBound" value="${Number(row.premiumBound || 0)}" />
        </label>
        </div>
      </div>
      <div class="workspace-form-section">
        <div class="workspace-section-header">
          <h4>Notes</h4>
          <p>Context for the next producer touchpoint.</p>
        </div>
        <div class="form-grid">
        <label class="full-span">
          Notes
          <textarea name="notes">${escapeHtml(row.notes || "")}</textarea>
        </label>
        </div>
      </div>
      <div class="form-actions">
        <button class="button button-primary" type="submit">${row.id ? "Save Changes" : "Create Lead"}</button>
        <button class="button button-ghost" type="button" id="newOpportunityButton">Start New</button>
        ${row.id ? '<button class="button button-secondary" type="button" id="deleteOpportunityButton">Delete</button>' : ""}
      </div>
    </form>
  `;
}

function renderLeadWorkspace(row, timeline) {
  const taskTone = row.taskOverdue ? "bad" : (row.taskPriority === "High" ? "warn" : "good");
  const attachments = row.id ? getOpportunityAttachments(row.id) : [];
  return `
    <div class="lead-workspace">
      <article class="form-card lead-workspace-panel">
        <div class="panel-header">
          <div>
            <h3>${row.id ? "Lead Workspace" : "Create Opportunity"}</h3>
            <p>${row.id ? escapeHtml(row.leadNumber) : "New lead record"}</p>
          </div>
          ${row.id ? `<span class="pill">${escapeHtml(row.status)}</span>` : ""}
        </div>
        ${row.id ? `
          <div class="workspace-overview-grid">
            <article class="mini-card">
              <h3>Next Task</h3>
              <strong>${escapeHtml(row.nextTask || "Not set")}</strong>
              <div class="stat-meta">${escapeHtml(row.nextFollowUpDate || "No follow-up date")}</div>
            </article>
            <article class="mini-card">
              <h3>Priority</h3>
              <span class="status-pill" data-tone="${taskTone}">${escapeHtml(row.taskPriority || "Medium")}</span>
              <div class="stat-meta">${row.taskOverdue ? "Overdue now" : "On track"}</div>
            </article>
            <article class="mini-card">
              <h3>Renewal Window</h3>
              <strong>${escapeHtml(row.resolvedExpirationDate || row.expirationDate || "Not set")}</strong>
              <div class="stat-meta">${escapeHtml(row.renewalStatus || "Not Started")} · ${escapeHtml(row.incumbentCarrier || "No incumbent carrier")}</div>
            </article>
            <article class="mini-card">
              <h3>Commission Snapshot</h3>
              <strong>${formatCurrency(isAdmin() ? row.actualAgencyComm : row.actualRepPayout)}</strong>
              <div class="stat-meta">${isAdmin() ? "Actual agency commission" : "Your actual payout"}</div>
            </article>
          </div>
        ` : ""}
        ${renderOpportunityForm(row)}
      </article>
      <article class="table-card lead-workspace-panel">
        <div class="panel-header">
          <div>
            <h3>Log Lead Activity</h3>
            <p>${row.id ? "Capture calls, emails, texts, appointments, and outcomes without leaving the lead." : "Create the lead first, then log rep outreach."}</p>
          </div>
        </div>
        ${renderOpportunityActivityComposer(row)}
      </article>
      <article class="table-card lead-workspace-panel">
        <div class="panel-header">
          <div>
            <h3>Activity Timeline</h3>
            <p>${row.id ? "Every meaningful lead change appears here." : "Timeline starts after the lead is created."}</p>
          </div>
        </div>
        ${renderOpportunityTimeline(row, timeline)}
      </article>
      <article class="table-card lead-workspace-panel">
        <div class="panel-header">
          <div>
            <h3>Attachments</h3>
            <p>${row.id ? "Quotes, proposals, dec pages, and renewal docs live with the lead." : "Create the lead first, then attach files here."}</p>
          </div>
        </div>
        ${renderOpportunityAttachments(row, attachments)}
      </article>
      <article class="table-card lead-workspace-panel">
        <div class="panel-header">
          <div>
            <h3>Send Reminder</h3>
            <p>${row.id ? "Use the agency templates to send a follow-up by email or text without leaving the lead." : "Create the lead first, then send reminders from here."}</p>
          </div>
        </div>
        ${renderOpportunityReminders(row)}
      </article>
      <article class="table-card lead-workspace-panel">
        <div class="panel-header">
          <div>
            <h3>Calendar Sync</h3>
            <p>${row.id ? "Create a Google Calendar event from this lead’s follow-up or appointment." : "Create the lead first, then push a follow-up to calendar."}</p>
          </div>
        </div>
        ${renderOpportunityCalendarSync(row)}
      </article>
    </div>
  `;
}

function renderCreateLeadWorkspace() {
  const row = blankOpportunity();
  return `
    <div class="lead-workspace single-workspace">
      <article class="form-card lead-workspace-panel">
        <div class="panel-header">
          <div>
            <h3>New Lead Intake</h3>
            <p>Enter the lead cleanly here, then move into update and stage management after it is created.</p>
          </div>
        </div>
        ${renderOpportunityForm(row)}
      </article>
    </div>
  `;
}

function renderOpportunityActivityComposer(row) {
  if (!row.id) {
    return `<div class="empty-state"><h3>Create the lead first</h3><p>Once the record exists, reps can log outreach and appointments here.</p></div>`;
  }

  return `
    <form id="activityLogForm" class="activity-log-form">
      <input type="hidden" name="opportunityId" value="${escapeHtml(row.id)}" />
      <label>
        Activity Type
        <select name="activityType">
          ${getCommunicationTypes().map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`).join("")}
        </select>
      </label>
      <label>
        Outcome
        <select name="outcome">
          ${getCommunicationOutcomes().map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`).join("")}
        </select>
      </label>
      <label>
        Next Follow-Up
        <input type="date" name="nextFollowUpDate" value="${escapeHtml(row.nextFollowUpDate || "")}" />
      </label>
      <label>
        Appointment Time
        <input type="datetime-local" name="appointmentAt" />
      </label>
      <label class="full-span">
        Activity Notes
        <textarea name="detail" placeholder="What happened on the touchpoint? What was said, sent, or scheduled?"></textarea>
      </label>
      <div class="form-actions">
        <button class="button button-primary" type="submit">Log Activity</button>
      </div>
    </form>
  `;
}

function renderOpportunityTimeline(row, timeline) {
  if (!row.id) {
    return `<div class="empty-state"><h3>Create the lead first</h3><p>Once the record exists, status changes, notes, and assignments will appear here.</p></div>`;
  }
  if (!timeline.length) {
    return `<div class="empty-state"><h3>No activity yet</h3><p>After you run the new database script, this lead will start recording a visible history automatically.</p></div>`;
  }
  return `
    <div class="timeline-list">
      ${timeline.map((item) => `
        <article class="timeline-entry">
          <div class="timeline-dot"></div>
          <div class="timeline-copy">
            <strong>${escapeHtml(item.title)}</strong>
            <div class="timeline-meta-row">
              ${item.activity_type && item.activity_type !== "system" ? `<span class="tag">${escapeHtml(item.activity_type)}</span>` : ""}
              ${item.outcome ? `<span class="tag">${escapeHtml(item.outcome)}</span>` : ""}
              ${item.next_follow_up_date ? `<span class="tag">Follow-up ${escapeHtml(item.next_follow_up_date)}</span>` : ""}
            </div>
            <p>${escapeHtml(item.detail || "")}</p>
            ${item.appointment_at ? `<span class="subtle">Appointment: ${formatDateTime(item.appointment_at)}</span>` : ""}
            <span class="subtle">${escapeHtml(item.actor_name || "System")} · ${formatDateTime(item.created_at)}</span>
          </div>
        </article>
      `).join("")}
    </div>
  `;
}

function renderCommunicationLeaderboard(rows) {
  if (!rows.length) {
    return `<div class="empty-state"><h3>No outreach logged yet</h3><p>Once reps start recording touches, this turns into a coaching and accountability view.</p></div>`;
  }

  return `
    <div class="owner-report-list">
      ${rows.map((row) => `
        <article class="owner-report-card">
          <div class="owner-report-header">
            <strong>${escapeHtml(row.name)}</strong>
            <span class="tag">${row.touches} touches</span>
          </div>
          <div class="owner-report-metrics">
            <span>${row.conversations} conversations</span>
            <span>${row.appointments} appointments</span>
            <span>${row.quoteTouches} quote deliveries</span>
          </div>
        </article>
      `).join("")}
    </div>
  `;
}

function renderOpportunityAttachments(row, attachments) {
  if (!row.id) {
    return `<div class="empty-state"><h3>Create the lead first</h3><p>Once the record exists, you can upload files tied directly to this account.</p></div>`;
  }

  return `
    <div class="attachment-stack">
      ${state.ui.uploadingAttachmentOpportunityId === row.id ? `
        <div class="upload-progress-card">
          <div class="upload-progress-head">
            <strong>Uploading ${escapeHtml(state.ui.uploadingAttachmentName || "file")}</strong>
            <span>${Math.round(state.ui.uploadProgress)}%</span>
          </div>
          <div class="upload-progress-track">
            <div class="upload-progress-fill" style="width: ${Math.max(4, state.ui.uploadProgress)}%"></div>
          </div>
        </div>
      ` : ""}
      <form id="attachmentUploadForm" class="attachment-form">
        <input type="hidden" name="opportunityId" value="${escapeHtml(row.id)}" />
        <label class="full-span">
          Add File
          <input type="file" name="file" required />
        </label>
        <label>
          File Type
          <select name="fileType">
            <option value="Quote">Quote</option>
            <option value="Proposal">Proposal</option>
            <option value="Dec Page">Dec Page</option>
            <option value="Renewal">Renewal</option>
            <option value="Carrier Doc">Carrier Doc</option>
            <option value="Other">Other</option>
          </select>
        </label>
        <button class="button button-primary" type="submit" ${state.ui.uploadingAttachmentOpportunityId === row.id ? "disabled" : ""}>
          ${state.ui.uploadingAttachmentOpportunityId === row.id ? "Uploading..." : "Upload File"}
        </button>
      </form>
      ${attachments.length
        ? `
          <div class="attachment-list">
            ${attachments.map((file) => `
              <article class="attachment-card">
                <div>
                  <strong>${escapeHtml(file.file_name)}</strong>
                  <div class="subtle">${escapeHtml(file.file_type || "Other")} · ${formatFileSize(file.file_size)} · ${escapeHtml(file.created_by_name || "System")}</div>
                </div>
                <div class="attachment-actions">
                  <button class="button button-ghost" type="button" data-download-attachment="${file.id}">Open</button>
                  <button class="button button-ghost" type="button" data-delete-attachment="${file.id}">Delete</button>
                </div>
              </article>
            `).join("")}
          </div>
        `
        : `<div class="empty-state"><h3>No files yet</h3><p>Upload quotes, proposals, and renewal documents so they stay attached to this lead.</p></div>`}
    </div>
  `;
}

function renderOpportunityCalendarSync(row) {
  if (!row.id) {
    return `<div class="empty-state"><h3>Create the lead first</h3><p>Once the record exists, you can turn follow-ups into real calendar events.</p></div>`;
  }

  if (!state.calendarConnection?.connected) {
    return `
      <div class="empty-state">
        <h3>Connect Google Calendar first</h3>
        <p>Open the Integrations tab to connect your calendar, then come back here to push this lead’s follow-up onto it.</p>
      </div>
    `;
  }

  const defaultDate = row.nextFollowUpDate || row.effectiveDate || todayIso();
  const startAt = `${defaultDate}T09:00`;
  const endAt = `${defaultDate}T09:30`;
  return `
    <form id="calendarEventForm" class="activity-log-form">
      <input type="hidden" name="opportunityId" value="${escapeHtml(row.id)}" />
      <label class="full-span">
        Event Title
        <input name="summary" value="${escapeHtml(`${row.businessName} follow-up`)}" />
      </label>
      <label>
        Start
        <input type="datetime-local" name="startAt" value="${escapeHtml(startAt)}" />
      </label>
      <label>
        End
        <input type="datetime-local" name="endAt" value="${escapeHtml(endAt)}" />
      </label>
      <label class="full-span">
        Event Notes
        <textarea name="description">${escapeHtml(`Lead ${row.leadNumber}\nSource: ${row.leadSource}\nTask: ${row.nextTask || "Follow-up"}\nNotes: ${row.notes || ""}`)}</textarea>
      </label>
      <div class="form-actions">
        <button class="button button-primary" type="submit" ${state.ui.calendarSyncLoading ? "disabled" : ""}>
          ${state.ui.calendarSyncLoading ? "Creating..." : "Create Google Calendar Event"}
        </button>
      </div>
    </form>
  `;
}

function renderOpportunityReminders(row) {
  if (!row.id) {
    return `<div class="empty-state"><h3>Create the lead first</h3><p>Once the record exists, you can send reminder messages from the approved templates.</p></div>`;
  }

  const emailReady = Boolean(state.communicationStatus?.email?.configured);
  const smsReady = Boolean(state.communicationStatus?.sms?.configured);
  const emailPreview = interpolateReminderTemplate(state.setup.communicationSettings.emailBodyTemplate || "", row);
  const smsPreview = interpolateReminderTemplate(state.setup.communicationSettings.smsBodyTemplate || "", row);

  return `
    <div class="section-stack">
      <div class="dashboard-grid compact-dashboard-grid">
        ${statCard("Email", row.contactEmail || "Missing", row.contactEmail ? "Recipient found" : "Add contact email")}
        ${statCard("SMS", row.contactPhone || "Missing", row.contactPhone ? "Recipient found" : "Add contact phone")}
      </div>
      <div class="two-column compact-two-column">
        <article class="mini-card reminder-preview-card">
          <h3>Email Preview</h3>
          <div class="subtle">${escapeHtml(state.setup.communicationSettings.emailSubjectTemplate || "")}</div>
          <p class="reminder-preview-copy">${escapeHtml(emailPreview)}</p>
          <button class="button button-primary" type="button" data-send-reminder="${escapeHtml(row.id)}" data-reminder-channel="email" ${!emailReady || !row.contactEmail || state.ui.reminderSending ? "disabled" : ""}>
            ${state.ui.reminderSending ? "Sending..." : "Send Email Reminder"}
          </button>
        </article>
        <article class="mini-card reminder-preview-card">
          <h3>SMS Preview</h3>
          <p class="reminder-preview-copy">${escapeHtml(smsPreview)}</p>
          <button class="button button-primary" type="button" data-send-reminder="${escapeHtml(row.id)}" data-reminder-channel="sms" ${!smsReady || !row.contactPhone || state.ui.reminderSending ? "disabled" : ""}>
            ${state.ui.reminderSending ? "Sending..." : "Send Text Reminder"}
          </button>
        </article>
      </div>
    </div>
  `;
}

function renderOpportunityTable(rows, selectedRows) {
  return `
    <table>
      <thead>
        <tr>
          ${isAdmin() ? "<th>Select</th>" : ""}
          <th>Lead #</th>
          <th>Business</th>
          <th>Rep</th>
          <th>Status</th>
          <th>Renewal</th>
          <th>Follow-Up</th>
          <th>Quoted</th>
          <th>Bound</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map((row) => `
          <tr data-select-opportunity="${escapeHtml(row.id)}">
            ${isAdmin() ? `
              <td>
                <input
                  type="checkbox"
                  data-select-toggle="${escapeHtml(row.id)}"
                  ${selectedRows.has(row.id) ? "checked" : ""}
                />
              </td>
            ` : ""}
            <td>${escapeHtml(row.leadNumber)}</td>
            <td>
              <strong>${escapeHtml(row.businessName)}</strong>
              <div class="subtle">${escapeHtml(row.leadSource)}</div>
            </td>
            <td>${escapeHtml(row.assignedRepName)}</td>
            <td>${statusPill(row.status, row.followUpOverdue)}</td>
            <td>
              <div>${escapeHtml(row.renewalStatus || "Not Started")}</div>
              <div class="subtle">${escapeHtml(row.resolvedExpirationDate || row.expirationDate || "No expiration date")}</div>
            </td>
            <td><span class="tag">${escapeHtml(row.followUpBucket)}</span></td>
            <td>${formatCurrency(row.premiumQuoted)}</td>
            <td>${formatCurrency(row.premiumBound)}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function renderOpportunityBoard(phases) {
  const hasRows = phases.some((phase) => phase.columns.some((column) => column.rows.length));
  if (!hasRows) {
    return document.getElementById("emptyStateTemplate").innerHTML;
  }

  return `
    <div class="phase-board">
      ${phases.map((phase) => `
        <section class="phase-section">
          <header class="phase-header">
            <div>
              <h3>${escapeHtml(phase.title)}</h3>
              <p>${escapeHtml(phase.description)}</p>
            </div>
          </header>
          <div class="kanban-board">
            ${phase.columns.map((group) => `
              <section class="kanban-column" data-stage-drop="${escapeHtml(group.status)}">
                <header class="kanban-column-header">
                  <div>
                    <h3>${escapeHtml(group.status)}</h3>
                    <p>${group.rows.length} lead${group.rows.length === 1 ? "" : "s"}</p>
                  </div>
                </header>
                <div class="kanban-column-body">
                  ${group.rows.map((row) => `
                    <article
                      class="kanban-card"
                      draggable="true"
                      data-card-drag="${escapeHtml(row.id)}"
                      data-open-opportunity="${escapeHtml(row.id)}"
                    >
                      <div class="kanban-card-top">
                        <strong>${escapeHtml(row.businessName)}</strong>
                        ${row.followUpOverdue ? '<span class="status-pill" data-tone="bad">Overdue</span>' : ""}
                      </div>
                      <div class="subtle">${escapeHtml(row.leadNumber)} · ${escapeHtml(row.leadSource)}</div>
                      <div class="kanban-meta">
                        <span class="tag">${escapeHtml(row.assignedRepName)}</span>
                        <span class="tag">${escapeHtml(row.followUpBucket)}</span>
                      </div>
                      ${row.nextTask ? `<div class="kanban-task-line"><strong>Next:</strong> ${escapeHtml(row.nextTask)}</div>` : ""}
                      <label class="kanban-stage-picker">
                        Stage
                        <select data-quick-status="${escapeHtml(row.id)}">
                          ${state.setup.statuses.map((status) => `<option value="${escapeHtml(status)}" ${row.status === status ? "selected" : ""}>${escapeHtml(status)}</option>`).join("")}
                        </select>
                      </label>
                      <div class="kanban-meta">
                        <span>Quoted ${formatCurrency(row.premiumQuoted)}</span>
                        <span>Bound ${formatCurrency(row.premiumBound)}</span>
                      </div>
                    </article>
                  `).join("")}
                </div>
              </section>
            `).join("")}
          </div>
        `).join("")}
    </div>
  `;
}

function coachingInput(row, field, value) {
  if (!isAdmin()) {
    return `<span>${escapeHtml(value)}</span>`;
  }
  return `<input data-coaching-rep="${row.repId}" data-coaching-field="${field}" value="${escapeHtml(value)}" />`;
}

function statCard(title, value, meta) {
  return `<article class="stat-card"><h3>${escapeHtml(title)}</h3><span class="stat-value">${value}</span><div class="stat-meta">${escapeHtml(meta)}</div></article>`;
}

function kpiCard(title, value, meta) {
  return `<article class="kpi-card"><h3>${escapeHtml(title)}</h3><span class="kpi-value">${escapeHtml(value)}</span><div class="stat-meta">${escapeHtml(meta)}</div></article>`;
}

function renderBarChart(rows, { valueFormatter = formatWholeNumber } = {}) {
  const populatedRows = rows.filter((row) => Number(row.value || 0) > 0);
  if (!populatedRows.length) {
    return `<div class="empty-state"><h3>No chart data yet</h3><p>As activity comes in, this section will visualize it automatically.</p></div>`;
  }
  const maxValue = Math.max(...populatedRows.map((row) => Number(row.value || 0)), 1);
  return `
    <div class="bar-list">
      ${populatedRows.map((row) => `
        <div class="bar-row">
          <strong>${escapeHtml(row.label)}</strong>
          <div class="bar-track">
            <div class="bar-fill" style="width:${Math.max(8, (Number(row.value || 0) / maxValue) * 100).toFixed(1)}%"></div>
          </div>
          <span>${escapeHtml(valueFormatter(row.value))}</span>
        </div>
      `).join("")}
    </div>
  `;
}

function renderCommissionList(rows) {
  if (!rows.length) {
    return `<div class="empty-state"><h3>No earned commission yet</h3><p>Once quotes bind, your commission detail will show up here.</p></div>`;
  }
  return `
    <div class="commission-list">
      ${rows.slice(0, 6).map((row) => `
        <article class="commission-card">
          <div>
            <strong>${escapeHtml(row.businessName)}</strong>
            <div class="subtle">${escapeHtml(row.carrier)} · ${escapeHtml(row.status)}</div>
          </div>
          <div class="commission-values">
            <span>Earned ${formatCurrency(row.actualRepPayout)}</span>
            <span>Projected ${formatCurrency(row.potentialRepPayout)}</span>
          </div>
        </article>
      `).join("")}
    </div>
  `;
}

function renderTaskQueue(rows) {
  if (!rows.length) {
    return `<div class="empty-state"><h3>No open tasks</h3><p>Once reps start setting next tasks and follow-up dates, the action queue will populate here.</p></div>`;
  }
  return `
    <div class="task-queue">
      ${rows.slice(0, 8).map((row) => `
        <article class="task-card" data-select-opportunity="${escapeHtml(row.id)}">
          <div class="task-card-top">
            <strong>${escapeHtml(row.businessName)}</strong>
            <span class="status-pill" data-tone="${row.taskOverdue ? "bad" : row.taskPriority === "High" ? "warn" : "good"}">${escapeHtml(row.taskPriority || "Medium")}</span>
          </div>
          <div class="subtle">${escapeHtml(row.assignedRepName)} · ${escapeHtml(row.status)}</div>
          <p>${escapeHtml(row.nextTask || "No task entered")}</p>
          <div class="task-card-meta">
            <span>${escapeHtml(row.nextFollowUpDate || "No follow-up date")}</span>
            <span>${row.taskOverdue ? "Overdue" : row.followUpBucket}</span>
          </div>
        </article>
      `).join("")}
    </div>
  `;
}

function renderAlertsPanel(alerts) {
  const items = [
    ...alerts.overdueTasks.slice(0, 3).map((row) => ({
      tone: "bad",
      title: `${row.businessName} is overdue`,
      detail: `${row.assignedRepName} has ${row.nextTask || "a follow-up"} due ${row.nextFollowUpDate || "now"}.`
    })),
    ...alerts.quotesAging.slice(0, 2).map((row) => ({
      tone: "warn",
      title: `${row.businessName} quote is aging`,
      detail: `${row.daysOpen} days open in Quoted with ${formatCurrency(row.premiumQuoted)} still in play.`
    })),
    ...alerts.staleLeads.slice(0, 2).map((row) => ({
      tone: "warn",
      title: `${row.businessName} is getting stale`,
      detail: `${row.daysOpen} days open in ${row.status}. Last activity ${row.lastActivityDate || "not logged"}.`
    })),
    ...alerts.renewalsDue.slice(0, 3).map((row) => ({
      tone: row.daysToExpiration !== null && row.daysToExpiration < 0 ? "bad" : "warn",
      title: `${row.businessName} renewal needs attention`,
      detail: `${row.renewalStatus} · ${row.daysToExpiration ?? "?"} days to expiration · ${row.assignedRepName}.`
    }))
  ];

  if (!items.length) {
    return `<div class="empty-state"><h3>No active alerts</h3><p>The visible pipeline is in good shape right now.</p></div>`;
  }

  return `
    <div class="alert-list">
      ${items.map((item) => `
        <article class="alert-card" data-tone="${item.tone}">
          <strong>${escapeHtml(item.title)}</strong>
          <p>${escapeHtml(item.detail)}</p>
        </article>
      `).join("")}
    </div>
  `;
}

function renderRenewalQueue(rows) {
  const items = rows.filter((row) => row.daysToExpiration !== null && row.daysToExpiration <= 90);
  if (!items.length) {
    return `<div class="empty-state"><h3>No renewals coming due</h3><p>As expiration dates get closer, renewal work will surface here automatically.</p></div>`;
  }
  return `
    <div class="task-queue renewal-queue">
      ${items.slice(0, 8).map((row) => `
        <article class="task-card" data-select-opportunity="${escapeHtml(row.id)}">
          <div class="task-card-top">
            <strong>${escapeHtml(row.businessName)}</strong>
            <span class="status-pill" data-tone="${row.daysToExpiration !== null && row.daysToExpiration < 0 ? "bad" : row.daysToExpiration !== null && row.daysToExpiration <= 30 ? "warn" : "good"}">
              ${escapeHtml(row.renewalStatus)}
            </span>
          </div>
          <div class="subtle">${escapeHtml(row.assignedRepName)} · ${escapeHtml(row.carrier || "No carrier")}</div>
          <p>${escapeHtml(row.resolvedExpirationDate || row.expirationDate || "No expiration date")} · ${escapeHtml(row.incumbentCarrier || "No incumbent carrier")}</p>
          <div class="task-card-meta">
            <span>${formatRenewalCountdown(row.daysToExpiration)}</span>
            <span>${escapeHtml(row.nextTask || "No next task")}</span>
          </div>
        </article>
      `).join("")}
    </div>
  `;
}

function renderOwnerRepPerformance(rows) {
  if (!rows.length) {
    return `<div class="empty-state"><h3>No rep data yet</h3><p>As producers start working leads, owner performance reporting will populate here.</p></div>`;
  }
  return `
    <div class="owner-report-list">
      ${rows.map((row) => `
        <article class="owner-report-card">
          <div class="owner-report-header">
            <strong>${escapeHtml(row.name)}</strong>
            <span>${formatCurrency(row.actualAgencyComm)}</span>
          </div>
          <div class="owner-report-metrics">
            <span>${row.leads} leads</span>
            <span>${formatPct(row.contactRate)} contact</span>
            <span>${formatPct(row.quoteRate)} quote</span>
            <span>${formatPct(row.quoteToBindRate)} quote-to-bind</span>
          </div>
          <div class="owner-report-flags">
            <span class="tag">${row.overdueCount} overdue</span>
            <span class="tag">${row.staleCount} stale</span>
          </div>
        </article>
      `).join("")}
    </div>
  `;
}

function renderOwnerSourcePerformance(rows) {
  if (!rows.length) {
    return `<div class="empty-state"><h3>No source data yet</h3><p>Source performance will appear after leads start flowing through the system.</p></div>`;
  }
  return `
    <div class="owner-report-list">
      ${rows.map((row) => `
        <article class="owner-report-card">
          <div class="owner-report-header">
            <strong>${escapeHtml(row.source)}</strong>
            <span>${formatCurrency(row.netRevenue)}</span>
          </div>
          <div class="owner-report-metrics">
            <span>${row.count} leads</span>
            <span>${formatCurrency(row.spend)} spend</span>
            <span>${formatCurrency(row.actualAgencyComm)} revenue</span>
            <span>${formatCurrency(row.revenuePerLead)} per lead</span>
          </div>
          <div class="owner-report-flags">
            <span class="tag">${formatPct(row.quoteRate)} quote rate</span>
            <span class="tag">${formatPct(row.bindRate)} bind rate</span>
          </div>
        </article>
      `).join("")}
    </div>
  `;
}

function humanize(value) {
  return value.replace(/Pct/g, " %").replace(/Days/g, " Days").replace(/([A-Z])/g, " $1").replace(/^./, (char) => char.toUpperCase());
}

function statusPill(status, overdue) {
  let tone = "good";
  if (overdue || status === "Lost") tone = "bad";
  if (["New Lead", "Attempted", "Quoted", "Pending Decision"].includes(status)) tone = "warn";
  return `<span class="status-pill" data-tone="${tone}">${escapeHtml(status)}</span>`;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0
  }).format(Number(value || 0));
}

function formatPct(value) {
  return `${(Number(value || 0) * 100).toFixed(1)}%`;
}

function formatDateTime(value) {
  if (!value) return "";
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit"
  }).format(new Date(value));
}

function formatWholeNumber(value) {
  return new Intl.NumberFormat("en-US", {
    maximumFractionDigits: 0
  }).format(Number(value || 0));
}

function formatCompactCurrency(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    notation: "compact",
    maximumFractionDigits: 1
  }).format(Number(value || 0));
}

function formatRenewalCountdown(daysToExpiration) {
  if (daysToExpiration === null || daysToExpiration === undefined) return "No expiration date";
  if (daysToExpiration < 0) return `${Math.abs(daysToExpiration)} days past expiration`;
  if (daysToExpiration === 0) return "Expires today";
  if (daysToExpiration === 1) return "1 day to renew";
  return `${daysToExpiration} days to renew`;
}

function formatFileSize(bytes) {
  const size = Number(bytes || 0);
  if (!size) return "0 B";
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function interpolateReminderTemplate(template, row) {
  const replacements = {
    contactName: row.contactName || "there",
    businessName: row.businessName || "your account",
    repName: row.assignedRepName || state.profile?.full_name || "our team",
    leadNumber: row.leadNumber || "",
    nextTaskSentence: row.nextTask
      ? `${row.nextTask}${row.nextFollowUpDate ? ` by ${row.nextFollowUpDate}` : ""}.`
      : "Just wanted to check in and keep your quote moving.",
    nextFollowUpDate: row.nextFollowUpDate || ""
  };

  return Object.entries(replacements).reduce(
    (output, [key, value]) => output.replaceAll(`{{${key}}}`, String(value || "")),
    template || ""
  );
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function blankOpportunity() {
  return {
    id: "",
    leadNumber: "",
    assignedUserId: isAdmin() ? "auto" : state.profile?.id || "",
    assignedRepName: state.profile?.full_name || "",
    dateReceived: todayIso(),
    leadSource: state.setup.leadSources[0] || "",
    businessName: "",
    targetNiche: "",
    productFocus: state.setup.products[0] || "",
    carrier: state.setup.carriers[0]?.name || "",
    policyType: "New",
    policyTermMonths: 12,
    renewalStatus: "Not Started",
    leadCost: 0,
    premiumQuoted: 0,
    premiumBound: 0,
    status: state.setup.statuses[0] || "New Lead",
    firstAttemptDate: "",
    lastActivityDate: "",
    nextFollowUpDate: "",
    notes: ""
  };
}

function generateLeadNumber(dateReceived) {
  const compact = `${dateReceived.slice(2, 4)}${dateReceived.slice(5, 7)}${dateReceived.slice(8, 10)}`;
  const matching = state.opportunities.filter((item) => item.dateReceived === dateReceived).length + 1;
  return `${compact}-${String(matching).padStart(4, "0")}`;
}

function bindShellEvents() {
  const signOutButton = document.getElementById("signOutButton");
  if (signOutButton) {
    signOutButton.addEventListener("click", async () => {
      await state.supabase.auth.signOut();
    });
  }

  const resetSessionButton = document.getElementById("resetSessionButton");
  if (resetSessionButton) {
    resetSessionButton.addEventListener("click", async () => {
      await resetLocalSession();
    });
  }

  const reloadAppButton = document.getElementById("reloadAppButton");
  if (reloadAppButton) {
    reloadAppButton.addEventListener("click", () => {
      window.location.reload();
    });
  }
}

function bindAppEvents() {
  const assistantBubbleButton = document.getElementById("assistantBubbleButton");
  if (assistantBubbleButton) {
    assistantBubbleButton.addEventListener("click", () => {
      state.ui.assistantOpen = !state.ui.assistantOpen;
      state.ui.assistantError = "";
      render();
    });
  }

  const assistantCloseButton = document.getElementById("assistantCloseButton");
  if (assistantCloseButton) {
    assistantCloseButton.addEventListener("click", () => {
      state.ui.assistantOpen = false;
      state.ui.assistantError = "";
      render();
    });
  }

  const assistantForm = document.getElementById("assistantForm");
  if (assistantForm) {
    assistantForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const input = document.getElementById("assistantInput");
      const question = String(input?.value || "").trim();
      if (!question) {
        state.ui.assistantError = "Type a question first.";
        render();
        return;
      }
      state.ui.assistantInput = question;
      try {
        await askClaudeAssistant(question);
      } catch (error) {
        state.ui.assistantError = error?.message || "Claude could not answer right now.";
        render();
      }
    });
  }

  const connectGoogleCalendarButton = document.getElementById("connectGoogleCalendarButton");
  if (connectGoogleCalendarButton) {
    connectGoogleCalendarButton.addEventListener("click", async () => {
      try {
        await connectGoogleCalendar();
      } catch (error) {
        state.ui.error = error.message || "Could not start Google Calendar connection.";
        render();
      }
    });
  }

  const disconnectGoogleCalendarButton = document.getElementById("disconnectGoogleCalendarButton");
  if (disconnectGoogleCalendarButton) {
    disconnectGoogleCalendarButton.addEventListener("click", async () => {
      try {
        await disconnectGoogleCalendar();
      } catch (error) {
        state.ui.error = error.message || "Could not disconnect Google Calendar.";
        render();
      }
    });
  }

  const toggleReminderEditingButton = document.getElementById("toggleReminderEditingButton");
  if (toggleReminderEditingButton) {
    toggleReminderEditingButton.addEventListener("click", () => {
      state.ui.reminderEditing = !state.ui.reminderEditing;
      render();
    });
  }

  document.querySelectorAll("[data-communication-setting]").forEach((input) => {
    input.addEventListener("change", async (event) => {
      const key = event.target.dataset.communicationSetting;
      state.setup.communicationSettings[key] = event.target.value;
      await persistSettings();
    });
  });

  const timeframeSelect = document.getElementById("timeframeSelect");
  if (timeframeSelect) {
    timeframeSelect.addEventListener("change", (event) => {
      state.ui.timeframe = event.target.value;
      render();
    });
  }

  document.querySelectorAll("[data-app-tab]").forEach((button) => {
    button.addEventListener("click", () => {
      state.ui.activeTab = button.dataset.appTab;
      render();
    });
  });

  document.querySelectorAll("[data-setup-tab]").forEach((button) => {
    button.addEventListener("click", () => {
      state.ui.setupTab = button.dataset.setupTab;
      render();
    });
  });

  document.querySelectorAll("[data-opportunity-tab]").forEach((button) => {
    button.addEventListener("click", () => {
      state.ui.opportunityTab = button.dataset.opportunityTab;
      render();
    });
  });

  const loginForm = document.getElementById("loginForm");
  if (loginForm) {
    const runLogin = async () => {
      state.ui.error = "";
      state.ui.notice = "";
      const formData = new FormData(loginForm);
      const email = String(formData.get("email"));
      const password = String(formData.get("password"));
      state.ui.authLoading = true;
      render();
      const { error } = await state.supabase.auth.signInWithPassword({
        email,
        password
      });
      state.ui.authLoading = false;
      if (error) {
        state.ui.error = error.message;
        render();
      }
    };

    loginForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      await runLogin();
    });

    const loginButton = document.getElementById("loginButton");
    if (loginButton) {
      loginButton.addEventListener("click", async (event) => {
        event.preventDefault();
        await runLogin();
      });
    }
  }

  const resetPasswordButton = document.getElementById("resetPasswordButton");
  if (resetPasswordButton) {
    resetPasswordButton.addEventListener("click", async () => {
      const emailInput = document.querySelector('#loginForm input[name="email"]');
      const email = emailInput?.value?.trim();
      if (!email) {
        state.ui.error = "Enter an email address first so the reset link knows where to go.";
        state.ui.notice = "";
        render();
        return;
      }
      const { error } = await state.supabase.auth.resetPasswordForEmail(email);
      if (error) {
        state.ui.error = error.message;
        state.ui.notice = "";
      } else {
        state.ui.error = "";
        state.ui.notice = `Password reset email sent to ${email}.`;
      }
      render();
    });
  }

  const searchInput = document.getElementById("searchInput");
  if (searchInput) {
    searchInput.addEventListener("input", (event) => {
      state.ui.search = event.target.value;
      render();
    });
  }

  const statusFilterSelect = document.getElementById("statusFilterSelect");
  if (statusFilterSelect) {
    statusFilterSelect.addEventListener("change", (event) => {
      state.ui.statusFilter = event.target.value;
      render();
    });
  }

  const repFilterSelect = document.getElementById("repFilterSelect");
  if (repFilterSelect) {
    repFilterSelect.addEventListener("change", (event) => {
      state.ui.repFilter = event.target.value;
      render();
    });
  }

  const sourceFilterSelect = document.getElementById("sourceFilterSelect");
  if (sourceFilterSelect) {
    sourceFilterSelect.addEventListener("change", (event) => {
      state.ui.sourceFilter = event.target.value;
      render();
    });
  }

  const dateFromInput = document.getElementById("dateFromInput");
  if (dateFromInput) {
    dateFromInput.addEventListener("change", (event) => {
      state.ui.dateFrom = event.target.value;
      render();
    });
  }

  const dateToInput = document.getElementById("dateToInput");
  if (dateToInput) {
    dateToInput.addEventListener("change", (event) => {
      state.ui.dateTo = event.target.value;
      render();
    });
  }

  const clearFiltersButton = document.getElementById("clearFiltersButton");
  if (clearFiltersButton) {
    clearFiltersButton.addEventListener("click", () => {
      state.ui.search = "";
      state.ui.repFilter = "All";
      state.ui.sourceFilter = "All";
      state.ui.statusFilter = "All";
      state.ui.dateFrom = "";
      state.ui.dateTo = "";
      render();
    });
  }

  document.querySelectorAll("[data-select-opportunity]").forEach((row) => {
    row.addEventListener("click", () => {
      state.ui.activeOpportunityId = row.dataset.selectOpportunity;
      state.ui.activeTab = "opportunities";
      state.ui.opportunityTab = "update";
      render();
    });
  });

  document.querySelectorAll("[data-open-opportunity]").forEach((card) => {
    card.addEventListener("click", () => {
      state.ui.activeOpportunityId = card.dataset.openOpportunity;
      state.ui.activeTab = "opportunities";
      state.ui.opportunityTab = "update";
      render();
    });
  });

  document.querySelectorAll("[data-quick-status]").forEach((select) => {
    select.addEventListener("click", (event) => {
      event.stopPropagation();
    });
    select.addEventListener("change", async (event) => {
      try {
        await quickUpdateOpportunityStatus(event.target.dataset.quickStatus, event.target.value);
      } catch (error) {
        state.ui.error = error.message || "Could not update that lead stage.";
        render();
      }
    });
  });

  document.querySelectorAll("[data-card-drag]").forEach((card) => {
    card.addEventListener("dragstart", (event) => {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", card.dataset.cardDrag);
    });
    card.addEventListener("dragend", () => {
      stopDragAutoScroll();
    });
  });

  document.removeEventListener("dragover", handleDragAutoScroll);
  document.removeEventListener("drop", stopDragAutoScroll);
  document.removeEventListener("dragend", stopDragAutoScroll);
  document.addEventListener("dragover", handleDragAutoScroll);
  document.addEventListener("drop", stopDragAutoScroll);
  document.addEventListener("dragend", stopDragAutoScroll);

  document.querySelectorAll("[data-stage-drop]").forEach((column) => {
    column.addEventListener("dragover", (event) => {
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
      column.classList.add("is-drop-target");
      updateDragAutoScroll(event.clientY);
    });
    column.addEventListener("dragleave", () => {
      column.classList.remove("is-drop-target");
    });
    column.addEventListener("drop", async (event) => {
      event.preventDefault();
      stopDragAutoScroll();
      column.classList.remove("is-drop-target");
      const opportunityId = event.dataTransfer.getData("text/plain");
      const nextStatus = column.dataset.stageDrop;
      if (!opportunityId || !nextStatus) return;
      try {
        await quickUpdateOpportunityStatus(opportunityId, nextStatus);
      } catch (error) {
        state.ui.error = error.message || "Could not move that lead.";
        render();
      }
    });
  });

  document.querySelectorAll("[data-select-toggle]").forEach((checkbox) => {
    checkbox.addEventListener("click", (event) => {
      event.stopPropagation();
    });
    checkbox.addEventListener("change", (event) => {
      const id = event.target.dataset.selectToggle;
      toggleOpportunitySelection(id, event.target.checked);
    });
  });

  const selectVisibleButton = document.getElementById("selectVisibleButton");
  if (selectVisibleButton) {
    selectVisibleButton.addEventListener("click", () => {
      const visibleIds = getFilteredOpportunityList(getVisibleOpportunities()).map((row) => row.id);
      state.ui.selectedOpportunityIds = [...new Set([...state.ui.selectedOpportunityIds, ...visibleIds])];
      render();
    });
  }

  const clearSelectionButton = document.getElementById("clearSelectionButton");
  if (clearSelectionButton) {
    clearSelectionButton.addEventListener("click", () => {
      state.ui.selectedOpportunityIds = [];
      render();
    });
  }

  const bulkAssignUserSelect = document.getElementById("bulkAssignUserSelect");
  if (bulkAssignUserSelect) {
    bulkAssignUserSelect.addEventListener("change", (event) => {
      state.ui.bulkAssignUserId = event.target.value;
    });
  }

  const bulkAssignButton = document.getElementById("bulkAssignButton");
  if (bulkAssignButton) {
    bulkAssignButton.addEventListener("click", async () => {
      try {
        await bulkAssignSelected();
      } catch (error) {
        state.ui.error = error.message || "Could not bulk reassign those leads.";
        render();
      }
    });
  }

  const newOpportunityButton = document.getElementById("newOpportunityButton");
  if (newOpportunityButton) {
    newOpportunityButton.addEventListener("click", () => {
      state.ui.activeOpportunityId = null;
      state.ui.opportunityTab = "create";
      render();
    });
  }

  const opportunityForm = document.getElementById("opportunityForm");
  if (opportunityForm) {
    opportunityForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const formData = Object.fromEntries(new FormData(opportunityForm).entries());
      try {
        await saveOpportunity(formData);
      } catch (error) {
        state.ui.error = error.message || "Could not save the lead.";
        render();
      }
    });
  }

  const deleteOpportunityButton = document.getElementById("deleteOpportunityButton");
  if (deleteOpportunityButton) {
    deleteOpportunityButton.addEventListener("click", async () => {
      const active = state.opportunities.find((item) => item.id === state.ui.activeOpportunityId);
      if (!active || !window.confirm(`Delete ${active.leadNumber}?`)) return;
      try {
        const { error } = await state.supabase.from("opportunities").delete().eq("id", active.id);
        if (error) throw error;
        state.ui.activeOpportunityId = null;
        await loadWorkspace();
      } catch (error) {
        state.ui.error = error.message || "Could not delete the lead.";
        render();
      }
    });
  }

  const attachmentUploadForm = document.getElementById("attachmentUploadForm");
  if (attachmentUploadForm) {
    attachmentUploadForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const formData = new FormData(attachmentUploadForm);
      try {
        await uploadOpportunityAttachment({
          opportunityId: String(formData.get("opportunityId") || ""),
          fileType: String(formData.get("fileType") || "Other"),
          file: formData.get("file")
        });
      } catch (error) {
        state.ui.error = error.message || "Could not upload the file.";
        render();
      }
    });
  }

  const activityLogForm = document.getElementById("activityLogForm");
  if (activityLogForm) {
    activityLogForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const formData = Object.fromEntries(new FormData(activityLogForm).entries());
      try {
        await saveCommunicationActivity({
          opportunityId: String(formData.opportunityId || ""),
          activityType: String(formData.activityType || "Note"),
          outcome: String(formData.outcome || "Other"),
          detail: String(formData.detail || ""),
          nextFollowUpDate: String(formData.nextFollowUpDate || ""),
          appointmentAt: String(formData.appointmentAt || "")
        });
      } catch (error) {
        state.ui.error = error.message || "Could not log that activity.";
        render();
      }
    });
  }

  const calendarEventForm = document.getElementById("calendarEventForm");
  if (calendarEventForm) {
    calendarEventForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const formData = Object.fromEntries(new FormData(calendarEventForm).entries());
      try {
        await createGoogleCalendarEvent({
          opportunityId: String(formData.opportunityId || ""),
          summary: String(formData.summary || ""),
          description: String(formData.description || ""),
          startAt: String(formData.startAt || ""),
          endAt: String(formData.endAt || "")
        });
      } catch (error) {
        state.ui.error = error.message || "Could not create the calendar event.";
        render();
      }
    });
  }

  document.querySelectorAll("[data-send-reminder]").forEach((button) => {
    button.addEventListener("click", async () => {
      try {
        await sendLeadReminder(button.dataset.sendReminder, button.dataset.reminderChannel);
      } catch (error) {
        state.ui.error = error.message || "Could not send the reminder.";
        render();
      }
    });
  });

  document.querySelectorAll("[data-download-attachment]").forEach((button) => {
    button.addEventListener("click", async () => {
      try {
        await openOpportunityAttachment(button.dataset.downloadAttachment);
      } catch (error) {
        state.ui.error = error.message || "Could not open that file.";
        render();
      }
    });
  });

  document.querySelectorAll("[data-delete-attachment]").forEach((button) => {
    button.addEventListener("click", async () => {
      try {
        await deleteOpportunityAttachment(button.dataset.deleteAttachment);
      } catch (error) {
        state.ui.error = error.message || "Could not delete that file.";
        render();
      }
    });
  });

  document.querySelectorAll("[data-coaching-field]").forEach((input) => {
    input.addEventListener("change", async (event) => {
      try {
        await saveCoachingNote(
          event.target.dataset.coachingRep,
          event.target.dataset.coachingField,
          event.target.value
        );
      } catch (error) {
        state.ui.error = error.message || "Could not save coaching notes.";
        render();
      }
    });
  });

  document.querySelectorAll("[data-assumption]").forEach((input) => {
    input.addEventListener("change", async (event) => {
      state.setup.assumptions[event.target.dataset.assumption] = Number(event.target.value || 0);
      await persistSettings();
    });
  });

  document.querySelectorAll("[data-profile-name]").forEach((input) => {
    input.addEventListener("change", async (event) => {
      await persistProfile(event.target.dataset.profileName, {
        full_name: event.target.value
      });
    });
  });

  document.querySelectorAll("[data-profile-role]").forEach((input) => {
    input.addEventListener("change", async (event) => {
      if (event.target.dataset.profileRole === state.profile.id && event.target.value !== "admin") {
        state.ui.error = "You cannot remove your own admin access.";
        render();
        return;
      }
      await persistProfile(event.target.dataset.profileRole, {
        role: event.target.value
      });
    });
  });

  document.querySelectorAll("[data-profile-active]").forEach((input) => {
    input.addEventListener("change", async (event) => {
      const nextActive = event.target.value === "true";
      const profileId = event.target.dataset.profileActive;
      const targetProfile = state.profiles.find((item) => item.id === profileId);
      if (!nextActive && profileId === state.profile.id) {
        state.ui.error = "You cannot deactivate your own admin account.";
        render();
        return;
      }
      if (!nextActive && getAssignedLeadCount(profileId) > 0) {
        state.ui.error = `Reassign ${getAssignedLeadCount(profileId)} lead(s) from ${targetProfile?.full_name || "this rep"} before deactivating them.`;
        render();
        return;
      }
      await persistProfile(event.target.dataset.profileActive, {
        active: nextActive
      });
    });
  });

  document.querySelectorAll("[data-remove-user]").forEach((button) => {
    button.addEventListener("click", async () => {
      try {
        await removeUserFromApp(button.dataset.removeUser);
      } catch (error) {
        state.ui.error = error.message || "Could not remove that user.";
        render();
      }
    });
  });

  document.querySelectorAll("[data-restore-user]").forEach((button) => {
    button.addEventListener("click", async () => {
      try {
        await persistProfile(button.dataset.restoreUser, { active: true });
        state.ui.notice = "User restored.";
        render();
      } catch (error) {
        state.ui.error = error.message || "Could not restore that user.";
        render();
      }
    });
  });

  const inviteRepForm = document.getElementById("inviteRepForm");
  if (inviteRepForm) {
    inviteRepForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const formData = new FormData(inviteRepForm);
      try {
        await sendRepInvite({
          email: String(formData.get("email") || "").trim(),
          fullName: String(formData.get("fullName") || "").trim()
        });
      } catch (error) {
        state.ui.error = error.message || "Could not send the rep invite.";
        render();
      }
    });
  }

  const toggleAssumptionEditingButton = document.getElementById("toggleAssumptionEditingButton");
  if (toggleAssumptionEditingButton) {
    toggleAssumptionEditingButton.addEventListener("click", () => {
      state.ui.assumptionEditing = !state.ui.assumptionEditing;
      state.ui.notice = state.ui.assumptionEditing
        ? "Agency assumptions unlocked for editing."
        : "Agency assumptions locked.";
      state.ui.error = "";
      render();
    });
  }

  const toggleCarrierEditingButton = document.getElementById("toggleCarrierEditingButton");
  if (toggleCarrierEditingButton) {
    toggleCarrierEditingButton.addEventListener("click", () => {
      state.ui.carrierEditing = !state.ui.carrierEditing;
      state.ui.notice = state.ui.carrierEditing
        ? "Carrier commission table unlocked for editing."
        : "Carrier commission table locked.";
      state.ui.error = "";
      render();
    });
  }

  const exportWorkbookButton = document.getElementById("exportWorkbookButton");
  if (exportWorkbookButton) {
    exportWorkbookButton.addEventListener("click", () => {
      try {
        exportAgencyWorkbook();
      } catch (error) {
        state.ui.error = error.message || "Could not export the workbook.";
        render();
      }
    });
  }

  const routingAutoAssignSelect = document.getElementById("routingAutoAssignSelect");
  if (routingAutoAssignSelect) {
    routingAutoAssignSelect.addEventListener("change", async (event) => {
      state.setup.routingRules.autoAssignEnabled = event.target.value === "true";
      await persistSettings();
    });
  }

  const routingModeSelect = document.getElementById("routingModeSelect");
  if (routingModeSelect) {
    routingModeSelect.addEventListener("change", async (event) => {
      state.setup.routingRules.mode = event.target.value;
      await persistSettings();
    });
  }

  document.querySelectorAll("[data-routing-source]").forEach((input) => {
    input.addEventListener("change", async (event) => {
      const source = event.target.dataset.routingSource;
      const userId = event.target.value;
      const otherRules = state.setup.routingRules.sourceRules.filter((rule) => rule.source !== source);
      state.setup.routingRules.sourceRules = userId
        ? [...otherRules, { source, userId }]
        : otherRules;
      await persistSettings();
    });
  });

  const downloadLeadImportTemplateButton = document.getElementById("downloadLeadImportTemplateButton");
  if (downloadLeadImportTemplateButton) {
    downloadLeadImportTemplateButton.addEventListener("click", () => {
      try {
        downloadLeadImportTemplate();
      } catch (error) {
        state.ui.error = error.message || "Could not download the import template.";
        render();
      }
    });
  }

  const leadImportFileInput = document.getElementById("leadImportFileInput");
  if (leadImportFileInput) {
    leadImportFileInput.addEventListener("change", async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      try {
        state.ui.error = "";
        state.ui.notice = "";
        state.ui.importFileName = file.name;
        state.ui.importPreviewRows = await buildLeadImportPreview(file);
        render();
      } catch (error) {
        state.ui.importFileName = "";
        state.ui.importPreviewRows = [];
        state.ui.error = error.message || "Could not preview that import file.";
        render();
      }
    });
  }

  const clearLeadImportButton = document.getElementById("clearLeadImportButton");
  if (clearLeadImportButton) {
    clearLeadImportButton.addEventListener("click", () => {
      state.ui.importPreviewRows = [];
      state.ui.importFileName = "";
      state.ui.notice = "";
      render();
    });
  }

  const importLeadsButton = document.getElementById("importLeadsButton");
  if (importLeadsButton) {
    importLeadsButton.addEventListener("click", async () => {
      try {
        await importLeadPreviewRows();
      } catch (error) {
        state.ui.error = error.message || "Could not import those leads.";
        render();
      }
    });
  }

  document.querySelectorAll("[data-carrier-name]").forEach((input) => {
    input.addEventListener("change", async (event) => {
      state.setup.carriers[Number(event.target.dataset.carrierName)].name = event.target.value;
      await persistSettings();
    });
  });

  document.querySelectorAll("[data-carrier-new]").forEach((input) => {
    input.addEventListener("change", async (event) => {
      state.setup.carriers[Number(event.target.dataset.carrierNew)].newPct = Number(event.target.value || 0);
      await persistSettings();
    });
  });

  document.querySelectorAll("[data-carrier-renewal]").forEach((input) => {
    input.addEventListener("change", async (event) => {
      state.setup.carriers[Number(event.target.dataset.carrierRenewal)].renewalPct = Number(event.target.value || 0);
      await persistSettings();
    });
  });
}

function handleDragAutoScroll(event) {
  updateDragAutoScroll(event.clientY);
}

async function saveOpportunity(formData) {
  const existing = state.opportunities.find((item) => item.id === formData.id);
  const assignedProfile = getAssignedProfileForForm(formData, existing);
  const normalizedFormData = {
    ...formData,
    assignedUserId: assignedProfile?.id || formData.assignedUserId
  };
  const payload = mapOpportunityToDb(normalizedFormData);
  const { data, error } = await state.supabase
    .from("opportunities")
    .upsert(payload)
    .select("*")
    .single();
  if (error) throw error;
  if (
    !existing &&
    isAdmin() &&
    state.setup.routingRules.autoAssignEnabled &&
    (!formData.assignedUserId || formData.assignedUserId === "auto") &&
    String(formData.leadSource || "").trim().toLowerCase() !== "self-generated"
  ) {
    await advanceRoundRobinCursor(assignedProfile?.id);
  }
  await logOpportunityActivity({
    opportunityId: data.id,
    title: existing ? "Lead workspace updated" : "Lead created",
    detail: existing
      ? buildOpportunityChangeSummary(existing, mapOpportunityFromDb(data))
      : `${data.business_name || "Lead"} entered the pipeline in ${data.status}${assignedProfile ? ` and was assigned to ${assignedProfile.full_name}` : ""}.`
  });
  state.ui.activeOpportunityId = data.id;
  state.ui.opportunityTab = existing ? "update" : "stage";
  await loadWorkspace();
}

async function uploadOpportunityAttachment({ opportunityId, fileType, file }) {
  if (!(file instanceof File) || !file.size) {
    throw new Error("Choose a file before uploading.");
  }
  const target = state.opportunities.find((item) => item.id === opportunityId);
  if (!target) {
    throw new Error("Save the lead first before uploading files.");
  }

  const safeName = file.name.replace(/[^\w.-]+/g, "-");
  const filePath = `${opportunityId}/${Date.now()}-${safeName}`;
  state.ui.uploadingAttachmentOpportunityId = opportunityId;
  state.ui.uploadingAttachmentName = file.name;
  state.ui.uploadProgress = 0;
  state.ui.error = "";
  render();

  try {
    await uploadFileWithProgress(filePath, file, (progress) => {
      state.ui.uploadProgress = progress;
      render();
    });

    const { error: metadataError } = await state.supabase.from("opportunity_attachments").insert({
      opportunity_id: opportunityId,
      file_name: file.name,
      file_path: filePath,
      file_type: fileType,
      file_size: file.size,
      mime_type: file.type || "application/octet-stream",
      created_by: state.profile?.id || null,
      created_by_name: state.profile?.full_name || state.session?.user?.email || "System"
    });

    if (metadataError) {
      await state.supabase.storage.from(ATTACHMENTS_BUCKET).remove([filePath]);
      if (String(metadataError.message || "").includes("opportunity_attachments")) {
        throw new Error("Attachment metadata table is missing. Run the latest Supabase schema first.");
      }
      throw metadataError;
    }

    await logOpportunityActivity({
      opportunityId,
      title: "Attachment uploaded",
      detail: `${file.name} added as ${fileType}.`
    });
    state.ui.notice = `${file.name} uploaded successfully.`;
    await loadWorkspace();
  } finally {
    state.ui.uploadProgress = 0;
    state.ui.uploadingAttachmentName = "";
    state.ui.uploadingAttachmentOpportunityId = "";
  }
}

async function saveCommunicationActivity({ opportunityId, activityType, outcome, detail, nextFollowUpDate, appointmentAt }) {
  const opportunity = state.opportunities.find((item) => item.id === opportunityId);
  if (!opportunity) {
    throw new Error("That lead could not be found.");
  }

  const cleanActivityType = matchConfiguredValue(activityType, getCommunicationTypes(), "Note");
  const cleanOutcome = matchConfiguredValue(outcome, getCommunicationOutcomes(), "Other");
  const cleanFollowUpDate = nextFollowUpDate || "";
  const cleanAppointmentAt = appointmentAt ? new Date(appointmentAt).toISOString() : null;
  const detailParts = [];
  if (detail.trim()) detailParts.push(detail.trim());
  if (cleanFollowUpDate) detailParts.push(`Next follow-up ${cleanFollowUpDate}.`);
  if (cleanAppointmentAt) detailParts.push(`Appointment scheduled ${formatDateTime(cleanAppointmentAt)}.`);

  await logOpportunityActivity({
    opportunityId,
    activityType: cleanActivityType,
    outcome: cleanOutcome,
    nextFollowUpDate: cleanFollowUpDate || null,
    appointmentAt: cleanAppointmentAt,
    title: `${cleanActivityType} logged`,
    detail: detailParts.join(" ") || `${cleanActivityType} recorded with outcome: ${cleanOutcome}.`
  });

  const updates = {
    last_activity_date: todayIso()
  };
  if (cleanFollowUpDate) {
    updates.next_follow_up_date = cleanFollowUpDate;
  }
  if (!opportunity.firstAttemptDate && ["Call", "Email", "Text", "Voicemail"].includes(cleanActivityType)) {
    updates.first_attempt_date = todayIso();
  }

  const { error } = await state.supabase
    .from("opportunities")
    .update(updates)
    .eq("id", opportunityId);
  if (error) {
    throw error;
  }

  state.ui.notice = `${cleanActivityType} logged for ${opportunity.businessName}.`;
  await loadWorkspace();
}

async function askClaudeAssistant(question) {
  if (!state.supabase) {
    throw new Error("Assistant backend is not ready.");
  }

  const {
    data: { session }
  } = await state.supabase.auth.getSession();
  if (!session?.access_token) {
    throw new Error("Sign in again before using Claude.");
  }

  state.ui.assistantLoading = true;
  state.ui.assistantError = "";
  state.ui.assistantMessages = [
    ...state.ui.assistantMessages,
    { role: "user", content: question }
  ];
  render();

  try {
    const activeOpportunity = state.opportunities.find((item) => item.id === state.ui.activeOpportunityId);
    const response = await fetch(`${APP_CONFIG.supabaseUrl}/functions/v1/${ASSISTANT_FUNCTION_NAME}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: APP_CONFIG.supabaseAnonKey,
        Authorization: `Bearer ${session.access_token}`
      },
      body: JSON.stringify({
        question,
        activeOpportunityId: activeOpportunity?.id || null,
        history: state.ui.assistantMessages.slice(-8)
      })
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || `Claude request failed (${response.status}).`);
    }

    state.ui.assistantMessages = [
      ...state.ui.assistantMessages,
      { role: "assistant", content: String(payload.answer || "No answer returned.") }
    ];
    state.ui.assistantInput = "";
  } finally {
    state.ui.assistantLoading = false;
    render();
  }
}

async function callEdgeFunction(name, body = {}) {
  if (!state.supabase) {
    throw new Error("Backend is not ready.");
  }
  const {
    data: { session }
  } = await state.supabase.auth.getSession();
  if (!session?.access_token) {
    throw new Error("Sign in again before using this action.");
  }

  const response = await fetch(`${APP_CONFIG.supabaseUrl}/functions/v1/${name}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: APP_CONFIG.supabaseAnonKey,
      Authorization: `Bearer ${session.access_token}`
    },
    body: JSON.stringify(body)
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || `${name} failed (${response.status}).`);
  }
  return payload;
}

async function connectGoogleCalendar() {
  state.ui.calendarSyncLoading = true;
  state.ui.error = "";
  state.ui.notice = "";
  render();
  try {
    const payload = await callEdgeFunction(GOOGLE_CALENDAR_CONNECT_FUNCTION, {
      returnTo: APP_CONFIG.appUrl || window.location.origin
    });
    if (!payload.url) {
      throw new Error("Google connection URL was not returned.");
    }
    window.location.href = payload.url;
  } finally {
    state.ui.calendarSyncLoading = false;
    render();
  }
}

async function disconnectGoogleCalendar() {
  state.ui.calendarSyncLoading = true;
  state.ui.error = "";
  state.ui.notice = "";
  render();
  try {
    await callEdgeFunction(GOOGLE_CALENDAR_DISCONNECT_FUNCTION);
    state.ui.notice = "Google Calendar disconnected.";
    await loadWorkspace({ showLoading: false });
  } finally {
    state.ui.calendarSyncLoading = false;
    render();
  }
}

async function createGoogleCalendarEvent({ opportunityId, summary, description, startAt, endAt }) {
  state.ui.calendarSyncLoading = true;
  state.ui.error = "";
  state.ui.notice = "";
  render();
  try {
    const payload = await callEdgeFunction(GOOGLE_CALENDAR_CREATE_EVENT_FUNCTION, {
      opportunityId,
      summary,
      description,
      startAt,
      endAt
    });
    state.ui.notice = payload.htmlLink
      ? "Calendar event created. Use the link in the timeline to open it."
      : "Calendar event created.";
    await loadWorkspace({ showLoading: false });
  } finally {
    state.ui.calendarSyncLoading = false;
    render();
  }
}

async function sendLeadReminder(opportunityId, channel) {
  const opportunity = state.opportunities.find((item) => item.id === opportunityId);
  if (!opportunity) {
    throw new Error("That lead could not be found.");
  }

  const emailSubject = interpolateReminderTemplate(state.setup.communicationSettings.emailSubjectTemplate || "", opportunity);
  const emailBody = interpolateReminderTemplate(state.setup.communicationSettings.emailBodyTemplate || "", opportunity);
  const smsBody = interpolateReminderTemplate(state.setup.communicationSettings.smsBodyTemplate || "", opportunity);

  state.ui.reminderSending = true;
  state.ui.error = "";
  state.ui.notice = "";
  render();

  try {
    const payload = await callEdgeFunction(SEND_REMINDER_FUNCTION, {
      opportunityId,
      channel,
      subject: emailSubject,
      body: channel === "email" ? emailBody : smsBody,
      replyToEmail: state.setup.communicationSettings.replyToEmail || ""
    });

    await logOpportunityActivity({
      opportunityId,
      activityType: channel === "email" ? "Email" : "Text",
      outcome: "Sent Information",
      title: channel === "email" ? "Reminder email sent" : "Reminder text sent",
      detail: channel === "email"
        ? `Reminder sent to ${opportunity.contactEmail}.${payload.providerId ? ` Provider ID: ${payload.providerId}` : ""}`
        : `Reminder sent to ${opportunity.contactPhone}.${payload.providerId ? ` Provider ID: ${payload.providerId}` : ""}`
    });

    state.ui.notice = channel === "email"
      ? `Reminder email sent to ${opportunity.contactEmail}.`
      : `Reminder text sent to ${opportunity.contactPhone}.`;
    await loadWorkspace({ showLoading: false });
  } finally {
    state.ui.reminderSending = false;
    render();
  }
}

async function openOpportunityAttachment(attachmentId) {
  const attachment = state.opportunityAttachments.find((item) => item.id === attachmentId);
  if (!attachment) {
    throw new Error("That file could not be found.");
  }
  const { data, error } = await state.supabase.storage
    .from(ATTACHMENTS_BUCKET)
    .createSignedUrl(attachment.file_path, 60);
  if (error || !data?.signedUrl) {
    throw new Error("Could not create a secure file link.");
  }
  window.open(data.signedUrl, "_blank", "noopener,noreferrer");
}

async function deleteOpportunityAttachment(attachmentId) {
  const attachment = state.opportunityAttachments.find((item) => item.id === attachmentId);
  if (!attachment) {
    throw new Error("That file could not be found.");
  }
  if (!window.confirm(`Delete ${attachment.file_name}?`)) {
    return;
  }

  const { error: storageError } = await state.supabase.storage
    .from(ATTACHMENTS_BUCKET)
    .remove([attachment.file_path]);
  if (storageError) {
    throw new Error("Could not remove the stored file.");
  }

  const { error: metadataError } = await state.supabase
    .from("opportunity_attachments")
    .delete()
    .eq("id", attachmentId);
  if (metadataError) {
    throw metadataError;
  }

  await logOpportunityActivity({
    opportunityId: attachment.opportunity_id,
    title: "Attachment deleted",
    detail: `${attachment.file_name} was removed from the lead workspace.`
  });
  state.ui.notice = `${attachment.file_name} deleted.`;
  await loadWorkspace();
}

async function uploadFileWithProgress(filePath, file, onProgress) {
  const session = state.session || (await state.supabase.auth.getSession()).data.session;
  if (!session?.access_token) {
    throw new Error("Your session expired. Sign in again before uploading.");
  }

  const encodedPath = filePath
    .split("/")
    .map((part) => encodeURIComponent(part))
    .join("/");
  const uploadUrl = `${APP_CONFIG.supabaseUrl}/storage/v1/object/${ATTACHMENTS_BUCKET}/${encodedPath}`;

  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", uploadUrl, true);
    xhr.setRequestHeader("apikey", APP_CONFIG.supabaseAnonKey);
    xhr.setRequestHeader("Authorization", `Bearer ${session.access_token}`);
    xhr.setRequestHeader("x-upsert", "false");
    if (file.type) {
      xhr.setRequestHeader("Content-Type", file.type);
    }

    xhr.upload.onprogress = (event) => {
      if (event.lengthComputable && typeof onProgress === "function") {
        onProgress((event.loaded / event.total) * 100);
      }
    };

    xhr.onload = () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        if (typeof onProgress === "function") {
          onProgress(100);
        }
        resolve();
        return;
      }
      reject(new Error("Upload failed. Run the latest Supabase schema and confirm the storage bucket exists."));
    };

    xhr.onerror = () => {
      reject(new Error("Upload failed. Check your connection and try again."));
    };

    xhr.send(file);
  });
}

async function quickUpdateOpportunityStatus(opportunityId, nextStatus) {
  const opportunity = state.opportunities.find((item) => item.id === opportunityId);
  if (!opportunity || opportunity.status === nextStatus) {
    return;
  }

  const payload = mapOpportunityToDb({
    ...opportunity,
    status: nextStatus,
    lastActivityDate: todayIso()
  });

  const { error } = await state.supabase
    .from("opportunities")
    .update(payload)
    .eq("id", opportunityId);

  if (error) {
    throw error;
  }

  await logOpportunityActivity({
    opportunityId,
    title: "Stage updated",
    detail: `${opportunity.status} moved to ${nextStatus}.`
  });
  state.ui.opportunityTab = "stage";
  state.ui.notice = `${opportunity.businessName} moved to ${nextStatus}. Admin dashboards will reflect the update automatically.`;
  await loadWorkspace();
}

async function saveCoachingNote(repId, field, value) {
  const weekStart = startOfWeek(todayIso());
  const existing = state.coachingNotes.find((item) => item.rep_user_id === repId && item.week_start === weekStart);
  const payload = {
    id: existing?.id,
    rep_user_id: repId,
    week_start: weekStart,
    biggest_gap: existing?.biggest_gap || "",
    behavior_to_improve: existing?.behavior_to_improve || "",
    action_commitment: existing?.action_commitment || "",
    next_review_notes: existing?.next_review_notes || "",
    created_by: state.profile.id
  };

  const fieldMap = {
    biggestGap: "biggest_gap",
    behaviorToImprove: "behavior_to_improve",
    actionCommitment: "action_commitment",
    nextReviewNotes: "next_review_notes"
  };
  payload[fieldMap[field]] = value;

  const { error } = await state.supabase.from("coaching_notes").upsert(payload);
  if (error) throw error;
  await loadWorkspace();
}

async function persistSettings() {
  try {
    const { error } = await state.supabase.from("app_settings").upsert(
      {
        singleton_key: "default",
        assumptions: state.setup.assumptions,
        routing_rules: state.setup.routingRules,
        communication_settings: state.setup.communicationSettings,
        lead_sources: state.setup.leadSources,
        statuses: state.setup.statuses,
        products: state.setup.products,
        carriers: state.setup.carriers,
        updated_by: state.profile.id
      },
      { onConflict: "singleton_key" }
    );
    if (error) throw error;
    await loadWorkspace();
  } catch (error) {
    state.ui.error = error.message || "Could not save setup settings.";
    render();
  }
}

async function advanceRoundRobinCursor(assignedProfileId) {
  const activeReps = getAssignableProfiles();
  if (!assignedProfileId || activeReps.length <= 1) return;
  const assignedIndex = activeReps.findIndex((item) => item.id === assignedProfileId);
  if (assignedIndex < 0) return;
  state.setup.routingRules.roundRobinCursor = (assignedIndex + 1) % activeReps.length;
  const { error } = await state.supabase.from("app_settings").upsert(
    {
      singleton_key: "default",
      routing_rules: state.setup.routingRules,
      updated_by: state.profile.id
    },
    { onConflict: "singleton_key" }
  );
  if (error) throw error;
}

async function persistProfile(profileId, changes) {
  try {
    state.ui.error = "";
    state.ui.notice = "";
    const { error } = await state.supabase.from("profiles").update(changes).eq("id", profileId);
    if (error) throw error;
    state.ui.notice = "User settings updated.";
    await loadWorkspace();
  } catch (error) {
    state.ui.error = error.message || "Could not update that user.";
    render();
  }
}

function toggleOpportunitySelection(id, checked) {
  if (checked) {
    state.ui.selectedOpportunityIds = [...new Set([...state.ui.selectedOpportunityIds, id])];
  } else {
    state.ui.selectedOpportunityIds = state.ui.selectedOpportunityIds.filter((item) => item !== id);
  }
  render();
}

async function bulkAssignSelected() {
  if (!state.ui.selectedOpportunityIds.length) {
    state.ui.error = "Select at least one lead first.";
    render();
    return;
  }
  if (!state.ui.bulkAssignUserId) {
    state.ui.error = "Choose a rep for the reassignment.";
    render();
    return;
  }

  const targetProfile = state.profiles.find((item) => item.id === state.ui.bulkAssignUserId);
  if (!targetProfile) {
    state.ui.error = "That rep could not be found.";
    render();
    return;
  }

  const { error } = await state.supabase
    .from("opportunities")
    .update({
      assigned_user_id: targetProfile.id,
      assigned_rep_name: targetProfile.full_name
    })
    .in("id", state.ui.selectedOpportunityIds);

  if (error) {
    throw error;
  }

  await Promise.all(state.ui.selectedOpportunityIds.map((opportunityId) =>
    logOpportunityActivity({
      opportunityId,
      title: "Lead reassigned",
      detail: `Assigned to ${targetProfile.full_name}.`
    })
  ));

  state.ui.notice = `${state.ui.selectedOpportunityIds.length} lead(s) reassigned to ${targetProfile.full_name}.`;
  state.ui.bulkAssignUserId = "";
  state.ui.selectedOpportunityIds = [];
  await loadWorkspace();
}

async function sendRepInvite({ email, fullName }) {
  if (!email || !fullName) {
    state.ui.error = "Enter both the rep's name and email.";
    render();
    return;
  }

  const redirectUrl = APP_CONFIG.appUrl || window.location.origin;
  if (!APP_CONFIG.appUrl && window.location.hostname === "localhost") {
    state.ui.error = "Set VITE_APP_URL to your live site URL before sending invites from local development.";
    render();
    return;
  }

  const inviteClient = createClient(APP_CONFIG.supabaseUrl, APP_CONFIG.supabaseAnonKey, {
    auth: {
      persistSession: false,
      autoRefreshToken: false,
      detectSessionInUrl: false
    }
  });

  const { error } = await inviteClient.auth.signInWithOtp({
    email,
    options: {
      shouldCreateUser: true,
      emailRedirectTo: redirectUrl,
      data: {
        full_name: fullName,
        role: "rep"
      }
    }
  });

  if (error) {
    throw error;
  }

  state.ui.error = "";
  state.ui.notice = `Invite sent to ${email}. They should confirm using the link that points to ${redirectUrl}.`;
  render();
}

function buildOpportunityChangeSummary(previous, next) {
  const changes = [];
  if (previous.status !== next.status) changes.push(`Stage ${previous.status} -> ${next.status}`);
  if (previous.assignedUserId !== next.assignedUserId) changes.push(`Reassigned to ${next.assignedRepName}`);
  if (previous.nextTask !== next.nextTask) changes.push(`Next task updated`);
  if (previous.nextFollowUpDate !== next.nextFollowUpDate) changes.push(`Follow-up date changed`);
  if (previous.renewalStatus !== next.renewalStatus) changes.push(`Renewal status ${previous.renewalStatus} -> ${next.renewalStatus}`);
  if (previous.notes !== next.notes) changes.push(`Notes updated`);
  if (previous.premiumQuoted !== next.premiumQuoted) changes.push(`Quoted premium changed`);
  if (previous.premiumBound !== next.premiumBound) changes.push(`Bound premium changed`);
  return changes.length ? changes.join(" · ") : "Lead details refreshed.";
}

async function logOpportunityActivity({ opportunityId, title, detail, activityType = "system", outcome = "", nextFollowUpDate = null, appointmentAt = null }) {
  if (!state.supabase || !opportunityId) return;
  const payload = {
    opportunity_id: opportunityId,
    actor_id: state.profile?.id || null,
    actor_name: state.profile?.full_name || state.session?.user?.email || "System",
    activity_type: activityType,
    outcome,
    next_follow_up_date: nextFollowUpDate || null,
    appointment_at: appointmentAt || null,
    title,
    detail: detail || ""
  };
  const { error } = await state.supabase.from("opportunity_activity").insert(payload);
  if (error && String(error.message || "").match(/activity_type|outcome|appointment_at|next_follow_up_date/i)) {
    throw new Error("Run the latest Supabase schema so communication logging fields are available.");
  }
  if (error && !String(error.message || "").includes("opportunity_activity")) {
    throw error;
  }
}

function exportAgencyWorkbook() {
  const allRows = getVisibleOpportunities();
  const allTimeSummary = summarize(allRows);
  const allTimeRenewalSummary = getRenewalSummary(allRows);
  const allTimeScorecards = getRepScorecards(allRows);
  const allTimeRoiRows = getRoiRows(allRows);
  const allTimeCoachingRows = getCoachingRows(allRows);

  const workbook = XLSX.utils.book_new();
  workbook.Props = {
    Title: "Golden Leaf Agency Workbook",
    Subject: "Agency operating workbook export",
    Author: "Golden Leaf Agency HQ"
  };

  const instructionsSheet = XLSX.utils.aoa_to_sheet([
    ["Golden Leaf Agency Workbook Export"],
    ["Generated", new Date().toLocaleString("en-US")],
    [""],
    ["Tabs included in this export mirror the original workbook structure."],
    ["Instructions"],
    ["Setup"],
    ["Opportunity Log"],
    ["Rep Scorecard"],
    ["Agency Dashboard"],
    ["Lead Source ROI"],
    ["Weekly Coaching"],
    ["Checklists"]
  ]);

  const setupSheet = XLSX.utils.aoa_to_sheet([
    ["Setup"],
    [""],
    ["Assumptions"],
    ["Setting", "Value"],
    ...Object.entries(state.setup.assumptions).map(([key, value]) => [humanize(key), value]),
    [""],
    ["Lead Sources"],
    ["Source"],
    ...state.setup.leadSources.map((value) => [value]),
    [""],
    ["Statuses"],
    ["Status"],
    ...state.setup.statuses.map((value) => [value]),
    [""],
    ["Products"],
    ["Product"],
    ...state.setup.products.map((value) => [value]),
    [""],
    ["Carrier Commission Table"],
    ["Carrier", "New %", "Renewal %", "Notes"],
    ...state.setup.carriers.map((carrier) => [carrier.name, carrier.newPct, carrier.renewalPct, carrier.notes || ""])
  ]);

  const opportunityRows = allRows.map((row) => ({
    "Lead #": row.leadNumber,
    "Date Received": row.dateReceived,
    "Assigned Rep": row.assignedRepName,
    "Lead Source": row.leadSource,
    "Business Name": row.businessName,
    "Target Niche": row.targetNiche,
    "Product Focus": row.productFocus,
    Carrier: row.carrier,
    "Incumbent Carrier": row.incumbentCarrier,
    "Policy Type": row.policyType,
    "Policy Term (Months)": row.policyTermMonths,
    "Renewal Status": row.renewalStatus,
    "Effective Date": row.effectiveDate,
    "Expiration Date": row.resolvedExpirationDate || row.expirationDate,
    Status: row.status,
    "Lead Cost": row.leadCost,
    "Premium Quoted": row.premiumQuoted,
    "Premium Bound": row.premiumBound,
    "Agency Comm %": row.agencyCommPct,
    "Projected Agency Comm": row.potentialAgencyComm,
    "Actual Agency Comm": row.actualAgencyComm,
    "Rep Payout %": row.repPayoutPct,
    "Projected Rep Payout": row.potentialRepPayout,
    "Actual Rep Payout": row.actualRepPayout,
    "Owner Net Agency Comm": row.ownerNetAgencyComm,
    "First Attempt Date": row.firstAttemptDate,
    "Last Activity Date": row.lastActivityDate,
    "Next Follow-Up Date": row.nextFollowUpDate,
    Notes: row.notes
  }));
  const opportunitySheet = XLSX.utils.json_to_sheet(opportunityRows);

  const scorecardSheet = XLSX.utils.json_to_sheet(allTimeScorecards.map((row) => ({
    Rep: row.name,
    Leads: row.leads,
    Binds: row.binds,
    "Same Day Rate": row.sameDayRate,
    "Contact Rate": row.contactRate,
    "Quote Rate": row.quoteRate,
    "Quote to Bind Rate": row.quoteToBindRate,
    "Actual Agency Comm": row.actualAgencyComm,
    "Actual Rep Payout": row.actualRepPayout
  })));

  const dashboardSheet = XLSX.utils.aoa_to_sheet([
    ["Agency Dashboard"],
    ["Metric", "Value"],
    ["Total Leads", allTimeSummary.totalLeads],
    ["Open Leads", allTimeSummary.openLeads],
    ["Overdue Follow-Up", allTimeSummary.overdueFollowUp],
    ["Quotes in Pipeline", allTimeSummary.quotesInPipeline],
    ["Binds", allTimeSummary.binds],
    ["Tracked Renewals", allTimeRenewalSummary.tracked],
    ["Renewals Due in 30 Days", allTimeRenewalSummary.dueThirty],
    ["Renewals Due in 60 Days", allTimeRenewalSummary.dueSixty],
    ["Retained Renewals", allTimeRenewalSummary.retained],
    ["Lost at Renewal", allTimeRenewalSummary.lostAtRenewal],
    ["Bound Premium", allTimeSummary.boundPremium],
    ["Pipeline Agency Comm", allTimeSummary.pipelineAgencyComm],
    ["Actual Agency Comm", allTimeSummary.actualAgencyComm],
    ["Actual Rep Payout", allTimeSummary.actualRepPayout],
    ["Owner Net Agency Comm", allTimeSummary.ownerNetAgencyComm]
  ]);

  const roiSheet = XLSX.utils.json_to_sheet(allTimeRoiRows.map((row) => ({
    "Lead Source": row.source,
    Count: row.count,
    Spend: row.spend,
    "Quote Rate": row.quoteRate,
    "Bind Rate": row.bindRate,
    "Actual Agency Comm": row.actualAgencyComm,
    "Actual Rep Payout": row.actualRepPayout
  })));

  const coachingSheet = XLSX.utils.json_to_sheet(allTimeCoachingRows.map((row) => ({
    "Week Start": row.weekStart,
    Rep: row.repName,
    Leads: row.leads,
    "Contact Rate": row.contactRate,
    Quotes: row.quotes,
    Binds: row.binds,
    "Biggest Gap": row.biggestGap,
    "Behavior to Improve": row.behaviorToImprove,
    "Action Commitment": row.actionCommitment,
    "Next Review Notes": row.nextReviewNotes
  })));

  const checklistSheet = XLSX.utils.aoa_to_sheet([
    ["Checklists"],
    ["Daily Rep Priorities"],
    ["Work fresh leads same day"],
    ["Log first attempt and follow-up date"],
    ["Advance lead to the correct stage"],
    ["Update notes after every conversation"],
    [""],
    ["Admin Priorities"],
    ["Review overdue follow-ups"],
    ["Reassign orphaned leads"],
    ["Review rep scorecards and coaching"],
    ["Export workbook for archive or handoff"]
  ]);

  XLSX.utils.book_append_sheet(workbook, instructionsSheet, "Instructions");
  XLSX.utils.book_append_sheet(workbook, setupSheet, "Setup");
  XLSX.utils.book_append_sheet(workbook, opportunitySheet, "Opportunity Log");
  XLSX.utils.book_append_sheet(workbook, scorecardSheet, "Rep Scorecard");
  XLSX.utils.book_append_sheet(workbook, dashboardSheet, "Agency Dashboard");
  XLSX.utils.book_append_sheet(workbook, roiSheet, "Lead Source ROI");
  XLSX.utils.book_append_sheet(workbook, coachingSheet, "Weekly Coaching");
  XLSX.utils.book_append_sheet(workbook, checklistSheet, "Checklists");

  XLSX.writeFile(workbook, `Golden-Leaf-Agency-Workbook-${todayIso()}.xlsx`);
  state.ui.notice = "Workbook exported.";
  state.ui.error = "";
  render();
}

async function removeUserFromApp(profileId) {
  const targetProfile = state.profiles.find((item) => item.id === profileId);
  if (!targetProfile) {
    state.ui.error = "That user could not be found.";
    render();
    return;
  }
  if (profileId === state.profile.id) {
    state.ui.error = "You cannot remove your own admin account.";
    render();
    return;
  }
  const leadCount = getAssignedLeadCount(profileId);
  if (leadCount > 0) {
    state.ui.error = `Reassign ${leadCount} lead(s) from ${targetProfile.full_name} before removing them.`;
    render();
    return;
  }
  if (!window.confirm(`Remove ${targetProfile.full_name} from the app? They will lose access and move to the removed-users list.`)) {
    return;
  }
  await persistProfile(profileId, { active: false });
  state.ui.notice = `${targetProfile.full_name} was removed from the active app roster.`;
  render();
}

async function resetLocalSession() {
  state.ui.error = "";
  state.ui.notice = "Resetting local session and reloading the app.";
  state.ui.recoveringSession = false;
  resetTransientUiState();
  sessionStorage.removeItem(APP_RECOVERY_STORAGE_KEY);
  if (state.supabase) {
    await state.supabase.auth.signOut({ scope: "local" });
  }
  window.location.reload();
}
