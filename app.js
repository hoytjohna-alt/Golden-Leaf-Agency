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
const AUTH_STORAGE_KEY = `${AUTH_STORAGE_PREFIX}-${APP_BUILD_ID}`;
const VERSION_CHECK_PATH = "/version.json";
const WORKSPACE_TIMEOUT_MS = 10000;
const VERSION_CHECK_INTERVAL_MS = 45000;
const ATTACHMENTS_BUCKET = "opportunity-files";

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
    opportunityView: "board",
    opportunityTab: "stage",
    setupTab: "users",
    carrierEditing: false,
    assumptionEditing: false
  }
};

let versionCheckTimer = null;
let dragAutoScrollVelocity = 0;
let dragAutoScrollFrame = null;

const appEl = document.getElementById("app");
const heroActionsEl = document.getElementById("heroActions");
const topNavEl = document.getElementById("topNav");

init();

async function init() {
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

  state.supabase.auth.onAuthStateChange(async (_event, session) => {
    state.session = session;
    state.profile = null;
    if (session?.user) {
      await loadWorkspace();
    } else {
      state.ui.loading = false;
      render();
    }
  });

  const {
    data: { session }
  } = await state.supabase.auth.getSession();
  state.session = session;
  if (session?.user) {
    await loadWorkspace();
  } else {
    state.ui.loading = false;
    render();
  }
}

async function loadWorkspace() {
  try {
    state.ui.loading = true;
    state.ui.error = "";
    state.ui.notice = "";
    render();

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
      state.ui.loading = false;
      render();
      return;
    }

    const [settings, opportunities, coachingNotes, profiles, opportunityActivities, opportunityAttachments] = await withTimeout(Promise.all([
      fetchSettings(),
      fetchOpportunities(),
      fetchCoachingNotes(),
      isAdmin() ? fetchProfiles() : Promise.resolve([profile]),
      fetchOpportunityActivities(),
      fetchOpportunityAttachments()
    ]), WORKSPACE_TIMEOUT_MS);

    state.setup = settings;
    state.opportunities = opportunities;
    state.coachingNotes = coachingNotes;
    state.profiles = profiles;
    state.opportunityActivities = opportunityActivities;
    state.opportunityAttachments = opportunityAttachments;
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
  }
}

function handleBuildVersionChange() {
  const previousBuildId = localStorage.getItem(APP_BUILD_STORAGE_KEY);
  if (previousBuildId !== APP_BUILD_ID) {
    clearStaleBuildSessions(previousBuildId);
    localStorage.setItem(APP_BUILD_STORAGE_KEY, APP_BUILD_ID);
    resetTransientUiState();
  }
}

function clearStaleBuildSessions(previousBuildId) {
  if (previousBuildId) {
    localStorage.removeItem(`${AUTH_STORAGE_PREFIX}-${previousBuildId}`);
  }
  Object.keys(localStorage)
    .filter((key) => key.startsWith(`${AUTH_STORAGE_PREFIX}-`) && key !== AUTH_STORAGE_KEY)
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
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      void checkForNewBuild();
    }
  });
  window.addEventListener("focus", () => {
    void checkForNewBuild();
  });
  if (versionCheckTimer) {
    window.clearInterval(versionCheckTimer);
  }
  versionCheckTimer = window.setInterval(() => {
    void checkForNewBuild();
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
  resetTransientUiState();
  state.ui.notice = "A new version of the app was deployed. Refreshing your workspace.";
  render();
  window.setTimeout(() => {
    window.location.reload();
  }, 250);
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
    .select("assumptions, lead_sources, statuses, products, carriers")
    .eq("singleton_key", "default")
    .single();
  if (error) throw error;
  return {
    assumptions: { ...seedSettings.assumptions, ...(data.assumptions || {}) },
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
    ...(isAdmin() ? [{ id: "reports", label: "Reports" }] : []),
    { id: "scorecards", label: "Scorecards" },
    { id: "coaching", label: "Coaching" },
    ...(isAdmin() ? [{ id: "setup", label: "Setup" }] : [])
  ];
}

function getSetupTabs() {
  return [
    { id: "users", label: "Users" },
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

  if (!SUPABASE_READY) {
    appEl.innerHTML = renderSetupRequired();
    bindShellEvents();
    return;
  }

  if (state.ui.loading) {
    appEl.innerHTML = `<section class="panel"><div class="empty-state"><h3>Loading workspace</h3><p>Connecting to the shared agency database.</p></div></section>`;
    bindShellEvents();
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

function renderOpportunityForm(row) {
  const assigneeOptions = (isAdmin() ? state.profiles.filter((item) => item.active) : [state.profile])
    .map((profile) => `<option value="${profile.id}" ${row.assignedUserId === profile.id ? "selected" : ""}>${escapeHtml(profile.full_name)}</option>`)
    .join("");

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
            <p>${escapeHtml(item.detail || "")}</p>
            <span class="subtle">${escapeHtml(item.actor_name || "System")} · ${formatDateTime(item.created_at)}</span>
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
        <button class="button button-primary" type="submit">Upload File</button>
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
    assignedUserId: state.profile?.id || "",
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
  const payload = mapOpportunityToDb(formData);
  const { data, error } = await state.supabase
    .from("opportunities")
    .upsert(payload)
    .select("*")
    .single();
  if (error) throw error;
  await logOpportunityActivity({
    opportunityId: data.id,
    title: existing ? "Lead workspace updated" : "Lead created",
    detail: existing
      ? buildOpportunityChangeSummary(existing, mapOpportunityFromDb(data))
      : `${data.business_name || "Lead"} entered the pipeline in ${data.status}.`
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
  const { error: uploadError } = await state.supabase.storage
    .from(ATTACHMENTS_BUCKET)
    .upload(filePath, file, { upsert: false });
  if (uploadError) {
    throw new Error("Upload failed. Run the latest Supabase schema and confirm the storage bucket exists.");
  }

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
    const { error } = await state.supabase.from("app_settings").upsert({
      singleton_key: "default",
      assumptions: state.setup.assumptions,
      lead_sources: state.setup.leadSources,
      statuses: state.setup.statuses,
      products: state.setup.products,
      carriers: state.setup.carriers,
      updated_by: state.profile.id
    });
    if (error) throw error;
    await loadWorkspace();
  } catch (error) {
    state.ui.error = error.message || "Could not save setup settings.";
    render();
  }
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

async function logOpportunityActivity({ opportunityId, title, detail }) {
  if (!state.supabase || !opportunityId) return;
  const payload = {
    opportunity_id: opportunityId,
    actor_id: state.profile?.id || null,
    actor_name: state.profile?.full_name || state.session?.user?.email || "System",
    title,
    detail: detail || ""
  };
  const { error } = await state.supabase.from("opportunity_activity").insert(payload);
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
