const { t } = window.WeHateSpamI18n;

const demoSenders = [
  { id: "super", name: "SuperOffers", address: "promo@superoffers.example", count: 84, category: "spam", future: ["block"], current: "spam", reasonKey: "reasonSuper", color: "coral", initial: "S", safe: true },
  { id: "news", name: "Weekly News", address: "digest@weeklynews.example", count: 42, category: "unsubscribe", future: ["unsubscribe"], current: "archive", reasonKey: "reasonNews", color: "blue", initial: "N", safe: true, canUnsubscribe: true },
  { id: "fit", name: "FitStyle Club", address: "hello@fitstyle.example", count: 67, category: "unsubscribe", future: ["unsubscribe"], current: "archive", reasonKey: "reasonFit", color: "green", initial: "F", safe: true, canUnsubscribe: true },
  { id: "crypto", name: "Crypto Bonus", address: "reward@coin-bonus.example", count: 31, category: "spam", future: ["block"], current: "spam_trash", reasonKey: "reasonCrypto", color: "coral", initial: "C", safe: true },
  { id: "travel", name: "Travel Daily", address: "daily@travel.example", count: 53, category: "unsubscribe", future: ["unsubscribe"], current: "archive", reasonKey: "reasonTravel", color: "blue", initial: "T", safe: true, canUnsubscribe: true },
  { id: "bank", name: "Horizon Bank", address: "alerts@horizonbank.example", count: 18, category: "keep", future: ["keep"], current: "keep", reasonKey: "reasonBank", color: "green", initial: "B", safe: false },
  { id: "studio", name: "North Studio", address: "hello@north-studio.example", count: 12, category: "review", future: ["review"], current: "keep", reasonKey: "reasonStudio", color: "blue", initial: "S", safe: false },
  { id: "maria", name: "Maria Costa", address: "maria@example.com", count: 9, category: "keep", future: ["keep"], current: "keep", reasonKey: "reasonMaria", color: "green", initial: "M", safe: false }
];

let senders = demoSenders;
let gmailAccessToken = null;
let gmailTokenClient = null;
let gmailPreviewResult = null;
let gmailActionResult = null;
let outlookClient = null;
let outlookInitialisation = null;
let outlookAccount = null;
let outlookAccessToken = null;
let outlookPreviewResult = null;
let outlookWriteAccessToken = null;
let outlookRuleAccessToken = null;
let outlookSession = 0;
let outlookPlan = null;
let outlookActionResult = null;
let outlookActionBusy = false;
let mailboxIdentity = "";
const GMAIL_CLIENT_ID = "569340402466-dape774mjn984a1lgnl4u2amrqkkqme0.apps.googleusercontent.com";
const GMAIL_METADATA_SCOPE = "https://www.googleapis.com/auth/gmail.metadata";
const GMAIL_MODIFY_SCOPE = "https://www.googleapis.com/auth/gmail.modify";
const GMAIL_SETTINGS_SCOPE = "https://www.googleapis.com/auth/gmail.settings.basic";
const OUTLOOK_CLIENT_ID = "f2684304-33d1-4f7b-a5eb-601129f53a61";
const OUTLOOK_SCOPES = ["Mail.ReadBasic"];
const OUTLOOK_WRITE_SCOPES = ["Mail.ReadWrite"];
const OUTLOOK_RULE_SCOPES = ["MailboxSettings.ReadWrite"];

const futureLabelKeys = { unsubscribe: "unsubscribe", block: "blockSender", auto_trash: "outlookAutoTrash", keep: "keepProtected", review: "decideLater", treated: "senderAlreadyTreated" };
const currentLabelKeys = { archive: "archive", spam: "markSpam", spam_trash: "spamTrash", trash: "moveTrash", keep: "doNothing" };
const futureLabel = (action) => t(futureLabelKeys[action]);
const currentLabel = (action) => t(currentLabelKeys[action]);
const state = {
  provider: "gmail",
  mode: "demo",
  filter: "all",
  selected: new Set(),
  selectedMembers: new Map(),
  future: Object.fromEntries(senders.map((sender) => [sender.id, new Set(sender.future)])),
  current: Object.fromEntries(senders.map((sender) => [sender.id, "trash"]))
};
const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const providerDialog = $("#provider-dialog");
const reviewDialog = $("#review-dialog");
let autoUnsubscribeContext = null;
let autoUnsubscribeAbort = null;
const openAddressGroups = new Set();

function clearAutoUnsubscribe() {
  autoUnsubscribeAbort?.abort();
  autoUnsubscribeAbort = null;
  autoUnsubscribeContext = null;
  $("#review-unsubscribes").replaceChildren();
  $("#review-action-status").classList.remove("batch-info", "batch-review-note");
}

function planExplanation(id) {
  if (state.mode === "outlook") {
    const current = state.current[id];
    const future = state.future[id];
    if (future.has("auto_trash")) return t(current === "keep" ? "outlookAutoTrashPlan" : "outlookAutoTrashCombinedPlan");
    return t(current === "trash" ? "outlookTrashPlan" : (current === "archive" ? "outlookArchivePlan" : "outlookNoChangesPlan"));
  }
  const future = state.future[id];
  const current = state.current[id];
  let futureText = t("planPending");
  if (future.has("treated")) futureText = t("planAlreadyTreated");
  else if (future.has("unsubscribe") && future.has("block")) futureText = t("planUnsubscribeBlock");
  else if (future.has("unsubscribe")) futureText = t("planUnsubscribe");
  else if (future.has("block")) futureText = t("planBlock");
  else if (future.has("keep")) futureText = t("planProtect");
  const currentText = t({ archive: "planArchive", spam: "planSpam", spam_trash: "planSpamTrash", trash: "planTrash", keep: "planKeep" }[current]);
  return `${futureText} ${currentText}`;
}

function futureSummary(id) {
  const actions = [...state.future[id]];
  return actions.map(futureLabel).join(" + ");
}

function futureValue(id) {
  const choices = state.future[id];
  if (choices.has("unsubscribe") && choices.has("block")) return "unsubscribe+block";
  return [...choices][0];
}

function futureOptions(sender) {
  if (sender.futureTreated) return [["treated", t("senderAlreadyTreated")]];
  if (state.mode === "outlook") return [
    ...(sender.canUnsubscribe ? [["unsubscribe", t("unsubscribe")]] : []),
    ["block", t("blockSender")],
    ["auto_trash", t("outlookAutoTrash")],
    ["review", t("doNothing")]
  ];
  const options = sender.canUnsubscribe
    ? [["unsubscribe", t("unsubscribe")], ["unsubscribe+block", t("unsubscribeBlock")], ["block", t("blockSender")], ["review", t("decideLater")]]
    : [["block", t("blockSender")], ["review", t("decideLater")]];
  return options;
}

const TREATED_SENDERS_STORAGE_KEY = "i-hate-spam-treated-senders-v1";
const TREATED_SENDER_TTL_MS = 90 * 24 * 60 * 60 * 1000;
const TREATED_SENDER_LIMIT = 500;

function treatedSenderKey(sender, provider = state.mode, identity = mailboxIdentity) {
  const address = String(sender?.address || "").trim().toLowerCase();
  const account = String(identity || "").trim().toLowerCase();
  if (provider === "demo" || !account || !address) return "";
  return `${provider}|${account}|${address}`;
}

function readTreatedSenders() {
  let parsed;
  try { parsed = JSON.parse(window.localStorage?.getItem(TREATED_SENDERS_STORAGE_KEY) || "{}"); }
  catch (_) { return {}; }
  const now = Date.now();
  const current = Object.fromEntries(Object.entries(parsed?.items || {}).filter(([key, value]) =>
    typeof key === "string" && Number.isFinite(value?.at) && now - value.at <= TREATED_SENDER_TTL_MS
  ));
  if (Object.keys(current).length !== Object.keys(parsed?.items || {}).length) {
    try { window.localStorage?.setItem(TREATED_SENDERS_STORAGE_KEY, JSON.stringify({ version: 1, items: current })); }
    catch (_) { /* Private browsing or disabled storage: continue without memory. */ }
  }
  return current;
}

function rememberTreatedSender(sender) {
  const key = treatedSenderKey(sender);
  if (!key) return;
  const items = readTreatedSenders();
  items[key] = { at: Date.now() };
  const bounded = Object.fromEntries(Object.entries(items).sort((a, b) => b[1].at - a[1].at).slice(0, TREATED_SENDER_LIMIT));
  try { window.localStorage?.setItem(TREATED_SENDERS_STORAGE_KEY, JSON.stringify({ version: 1, items: bounded })); }
  catch (_) { /* The action still succeeds when storage is unavailable. */ }
  sender.futureTreated = true;
}

function senderWasTreated(sender, provider = state.mode, identity = mailboxIdentity) {
  const key = treatedSenderKey(sender, provider, identity);
  return Boolean(key && readTreatedSenders()[key]);
}

function senderStatus(sender) {
  const status = document.createElement("div");
  const icon = document.createElement("span");
  const copy = document.createElement("span");
  const title = document.createElement("strong");
  const helpButton = document.createElement("button");
  const tooltip = document.createElement("span");
  const tooltipId = `status-tip-${sender.id}`;
  let detail;
  let help;

  status.className = "security-status";
  icon.className = "security-icon";
  icon.setAttribute("aria-hidden", "true");
  helpButton.className = "status-help";
  helpButton.type = "button";
  helpButton.textContent = "i";
  helpButton.setAttribute("aria-describedby", tooltipId);
  tooltip.className = "status-tooltip";
  tooltip.id = tooltipId;
  tooltip.setAttribute("role", "tooltip");

  if (sender.canUnsubscribe) {
    status.classList.add("verified");
    icon.textContent = "✓";
    title.textContent = t(state.mode !== "demo" ? "outlookNewsletterSignal" : "verifiedList");
    detail = state.mode !== "demo" ? "" : ` · ${t("safeUnsubscribe")}`;
    help = t(state.mode !== "demo" ? "unsubSignalsHelp" : "verifiedHelp");
    helpButton.setAttribute("aria-label", t(state.mode !== "demo" ? "recommendationAria" : "verifiedAria"));
  } else if (sender.category === "spam") {
    status.classList.add("suspicious");
    icon.textContent = "!";
    title.textContent = t("suspiciousSender");
    detail = ` · ${t("unavailableUnsubscribe")}`;
    help = t("suspiciousHelp");
    helpButton.setAttribute("aria-label", t("suspiciousAria"));
  } else {
    status.classList.add("neutral");
    icon.textContent = "·";
    const reasonText = t(sender.reasonKey);
    title.textContent = reasonText.split(" · ")[0];
    const detailLine = document.createElement("small");
    detailLine.textContent = sender.category === "review" ? t("uncertainSenderDetail") : reasonText.split(" · ").slice(1).join(" · ");
    copy.append(title);
    if (detailLine.textContent) copy.append(detailLine);
    help = sender.category === "review"
      ? t("uncertainHelp")
      : t("importantHelp");
    helpButton.setAttribute("aria-label", t("recommendationAria"));
  }

  if (!copy.childNodes.length) copy.append(title, document.createTextNode(detail));
  tooltip.textContent = help;
  status.append(icon, copy, helpButton, tooltip);
  return status;
}

function openProvider(provider = "gmail") { showProviderDialog(provider); }
function showProviderDialog(provider) {
  state.provider = provider;
  const title = $("#provider-dialog-title");
  const copy = $("#provider-dialog-copy");
  const status = $("#provider-status");
  const openDemo = $("#open-demo");
  status.textContent = "";
  if (provider === "outlook") {
    title.textContent = t("outlookComingTitle");
    copy.textContent = t("outlookComingCopy");
    openDemo.textContent = t("openDemo");
  } else {
    title.textContent = t("safeModeTitle");
    copy.textContent = t("safeModeCopy");
    openDemo.textContent = t("openDemo");
  }
  providerDialog.showModal();
}
$$('[data-provider]').forEach((button) => button.addEventListener("click", () => {
  if (button.dataset.provider === "gmail") connectGmail();
  else connectOutlook();
}));
$("#demo-button").addEventListener("click", () => openProvider("gmail"));
$("#provider-dialog .dialog-close").addEventListener("click", () => providerDialog.close());
providerDialog.addEventListener("click", (event) => { if (event.target === providerDialog) providerDialog.close(); });

function startDemo(provider = state.provider) {
  if (outlookActionBusy) return;
  state.provider = provider;
  state.mode = "demo";
  gmailActionResult = null;
  outlookPreviewResult = null;
  setSenders(demoSenders);
  providerDialog.close();
  $("#landing-view").hidden = true;
  $(".site-header").hidden = true;
  $("#app-view").hidden = false;
  $("#dashboard").hidden = true;
  $("#success-view").hidden = true;
  $("#gmail-preview").hidden = true;
  $("#analysis-loading").hidden = false;
  const icon = $("#account-provider-icon");
  icon.textContent = provider === "gmail" ? "M" : "O";
  icon.className = provider === "gmail" ? "gmail-icon" : "outlook-icon";
  $("#account-provider").textContent = t("fictionalProvider", { provider: provider === "gmail" ? "Gmail" : "Outlook" });
  $("#account-label").textContent = t("demoInbox");
  simulateAnalysis();
}

$("#open-demo").addEventListener("click", () => startDemo());
function simulateAnalysis() {
  const messageKeys = ["loadingFind", "loadingSecurity", "loadingGrouping", "loadingRecommendations"];
  let step = 0;
  $("#loading-progress").style.width = "4%";
  const timer = setInterval(() => {
    step += 1;
    $("#loading-progress").style.width = `${Math.min(100, 8 + step * 24)}%`;
    $("#loading-status").textContent = t(messageKeys[Math.min(step, messageKeys.length - 1)]);
    if (step >= 4) { clearInterval(timer); setTimeout(showDashboard, 250); }
  }, 430);
}

function updateDashboardCopy() {
  const isGmail = state.mode === "gmail" && gmailPreviewResult;
  const isOutlook = state.mode === "outlook" && outlookPreviewResult;
  const realPreview = isGmail ? gmailPreviewResult : (isOutlook ? outlookPreviewResult : null);
  $("#dashboard").classList.remove("readonly-mode");
  $("#dashboard").classList.toggle("outlook-test-mode", Boolean(isOutlook));
  $("#dashboard-eyebrow").textContent = t(isOutlook ? "outlookReadOnlyEyebrow" : (isGmail ? "realAnalysisComplete" : "analysisComplete"));
  $("#dashboard-title").textContent = t(isOutlook ? "outlookReadOnlyTitle" : (isGmail ? "realInboxTitle" : "inboxCanBreathe"));
  $("#dashboard-copy").textContent = t(isOutlook ? "outlookReadOnlyCopy" : (isGmail ? "realReviewCopy" : "reviewSuggestions"), { count: realPreview?.scanned || 0 });
  if (realPreview) {
    $("#summary-email-count").textContent = String(senders.length);
    $("#summary-email-label").textContent = t("gmailSendersFound");
    $("#summary-newsletter-count").textContent = String(senders.filter((sender) => sender.canUnsubscribe).length);
    $("#summary-newsletter-label").textContent = t(isOutlook ? "outlookNewsletterSignal" : "safeNewsletters");
    $("#summary-review-count").textContent = String(senders.filter((sender) => sender.category === "review").length);
    $("#summary-review-label").textContent = t("uncertainSenders");
    $("#summary-scan-count").textContent = String(realPreview.scanned);
    $("#summary-scan-label").textContent = t("latestMessages");
  } else {
    $("#summary-email-count").textContent = "847";
    $("#summary-email-label").textContent = t("canLeaveSummary");
    $("#summary-newsletter-count").textContent = "126";
    $("#summary-newsletter-label").textContent = t("protectedSenders");
    $("#summary-review-count").textContent = "3";
    $("#summary-review-label").textContent = t("needReview");
    $("#summary-scan-count").textContent = "~12 min";
    $("#summary-scan-label").textContent = t("savedPerWeek");
  }
  $("#disconnect-button").textContent = t(isGmail || isOutlook ? "disconnectAccount" : "exitDemo");
}

function showDashboard() {
  $("#analysis-loading").hidden = true;
  $("#gmail-preview").hidden = true;
  $("#success-view").hidden = true;
  $("#dashboard").hidden = false;
  updateDashboardCopy();
  renderSenders();
  updateBar();
}

function setSenders(nextSenders, provider = state.mode, account = "") {
  clearAutoUnsubscribe();
  senders = nextSenders;
  mailboxIdentity = provider === "demo" ? "" : String(account || mailboxIdentity || "").trim().toLowerCase();
  senders.forEach((sender) => {
    const members = senderMembers(sender);
    members.forEach((member) => { member.futureTreated = provider !== "demo" && senderWasTreated(member, provider, mailboxIdentity); });
    sender.treatedAddressCount = members.filter((member) => member.futureTreated).length;
    sender.futureTreated = provider !== "demo" && members.length > 0 && sender.treatedAddressCount === members.length;
  });
  state.selected.clear();
  state.selectedMembers.clear();
  openAddressGroups.clear();
  state.filter = "all";
  state.future = Object.fromEntries(senders.map((sender) => [sender.id, new Set(sender.futureTreated ? ["treated"] : sender.future.map((action) => action === "keep" ? "review" : action))]));
  state.current = Object.fromEntries(senders.map((sender) => [sender.id, "trash"]));
}

function enterGmailLoading() {
  state.mode = "gmail";
  $("#landing-view").hidden = true;
  $(".site-header").hidden = true;
  $("#app-view").hidden = false;
  $("#analysis-loading").hidden = false;
  $("#dashboard").hidden = true;
  $("#success-view").hidden = true;
  $("#gmail-preview").hidden = true;
  $("#account-provider-icon").textContent = "M";
  $("#account-provider-icon").className = "gmail-icon";
  $("#account-label").textContent = t("connectedAccount");
  $("#account-provider").textContent = t("gmailConnecting");
  $("#loading-status").textContent = t("gmailReadingMetadata");
  $("#loading-progress").style.width = "18%";
  $("#analysis-loading small").textContent = t("gmailMetadataNotice");
}

function enterOutlookLoading() {
  state.mode = "outlook";
  state.provider = "outlook";
  $("#landing-view").hidden = true;
  $(".site-header").hidden = true;
  $("#app-view").hidden = false;
  $("#analysis-loading").hidden = false;
  $("#dashboard").hidden = true;
  $("#success-view").hidden = true;
  $("#gmail-preview").hidden = true;
  $("#account-provider-icon").textContent = "O";
  $("#account-provider-icon").className = "outlook-icon";
  $("#account-label").textContent = t("connectedAccount");
  $("#account-provider").textContent = t("outlookConnecting");
  $("#loading-status").textContent = t("gmailReadingMetadata");
  $("#loading-progress").style.width = "18%";
  $("#analysis-loading small").textContent = t("outlookMetadataNotice");
}

function showProviderError(message) {
  $("#provider-dialog-title").textContent = t("gmailConnectionErrorTitle");
  $("#provider-dialog-copy").textContent = message || t("gmailConnectionErrorCopy");
  $("#provider-status").textContent = "";
  $("#open-demo").textContent = t("openDemo");
  providerDialog.showModal();
}

function safeMailContext() {
  if (window.isSecureContext && window.self === window.top) return true;
  showProviderError(t("secureWindowRequired"));
  return false;
}

function parseFromHeader(value = "") {
  const match = value.match(/^(?:\"?([^\"]*?)\"?\s*)?<([^<>\s]+@[^<>\s]+)>$/) || value.match(/^([^<>\s]+@[^<>\s]+)$/);
  if (!match) return { name: value.trim() || "Unknown sender", address: value.trim().toLowerCase() };
  const address = (match[2] || match[1] || "").trim().toLowerCase();
  const name = (match[1] && match[2] ? match[1].trim() : address.split("@")[0]).replace(/^\"|\"$/g, "") || address;
  return { name, address };
}

const MULTI_LABEL_EMAIL_SUFFIXES = new Set([
  "co.uk", "org.uk", "ac.uk", "com.au", "net.au", "org.au", "co.nz", "com.br", "com.pt",
  "co.jp", "co.in", "com.mx", "com.tr", "com.sg", "com.hk", "co.za"
]);
const SHARED_SENDER_DOMAINS = new Set([
  "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com", "yahoo.com", "icloud.com", "me.com", "aol.com",
  "beehiiv.com", "substack.com", "convertkit.com", "kit.com", "kit-mail.com", "kit-mail3.com", "mailchimp.com",
  "sendgrid.net", "amazonses.com", "mailgun.org", "brevo.com", "stripe.com"
]);
const MULTI_SERVICE_SENDER_DOMAINS = new Set([
  "google.com", "microsoft.com", "apple.com", "amazon.com", "meta.com"
]);
const GENERIC_SENDER_NAMES = new Set(["email", "hello", "info", "mail", "newsletter", "no reply", "noreply", "notifications", "support", "team"]);

function senderMembers(sender) {
  return Array.isArray(sender?.members) && sender.members.length ? sender.members : [sender];
}

function senderAddresses(sender) {
  return [...new Set(senderMembers(sender).map((member) => String(member.address || "").toLowerCase()).filter(Boolean))];
}

function selectedSenderMembers(sender) {
  const selected = state.selectedMembers.get(sender.id);
  if (!selected) return state.selected.has(sender.id) ? senderMembers(sender) : [];
  return senderMembers(sender).filter((member) => selected.has(member.id));
}

function senderEmailCountDisplay(sender) {
  const members = senderMembers(sender);
  const selected = selectedSenderMembers(sender);
  if (selected.length > 0 && selected.length < members.length) {
    return {
      value: t("selectedEmailFraction", {
        count: selected.reduce((total, member) => total + Number(member.count || 0), 0),
        total: sender.count
      }),
      label: t("emailsSelected"),
      partial: true
    };
  }
  return { value: String(sender.count), label: t("emails"), partial: false };
}

function setSenderSelection(sender, checked) {
  if (!checked) {
    state.selected.delete(sender.id);
    state.selectedMembers.delete(sender.id);
    return;
  }
  state.selected.add(sender.id);
  state.selectedMembers.set(sender.id, new Set(senderMembers(sender).map((member) => member.id)));
}

function setMemberSelection(sender, member, checked) {
  const selected = new Set(state.selectedMembers.get(sender.id) || (state.selected.has(sender.id) ? senderMembers(sender).map((item) => item.id) : []));
  checked ? selected.add(member.id) : selected.delete(member.id);
  if (selected.size) {
    state.selected.add(sender.id);
    state.selectedMembers.set(sender.id, selected);
  } else {
    state.selected.delete(sender.id);
    state.selectedMembers.delete(sender.id);
  }
}

function ensureSenderSelection(sender) {
  if (!selectedSenderMembers(sender).length) setSenderSelection(sender, true);
}

function registrableSenderDomain(address = "") {
  const host = String(address).split("@").pop().toLowerCase().replace(/^\.+|\.+$/g, "");
  const labels = host.split(".").filter(Boolean);
  if (labels.length < 2) return host;
  const lastTwo = labels.slice(-2).join(".");
  return MULTI_LABEL_EMAIL_SUFFIXES.has(lastTwo) && labels.length >= 3 ? labels.slice(-3).join(".") : lastTwo;
}

function senderBrandName(name = "") {
  return String(name).normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLowerCase()
    .replace(/\([^)]*\)|\[[^\]]*\]/g, " ")
    .replace(/\b(?:email|newsletter|newsletters|notifications?|no[ -]?reply|support|team|marketing|digest|updates?|offers?|official)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ").trim();
}

function senderDomainBrand(address = "") {
  const domain = registrableSenderDomain(address);
  if (!domain || SHARED_SENDER_DOMAINS.has(domain)) return "";
  let value = domain.split(".")[0].replace(/[^a-z0-9]/g, "");
  const noise = /^(?:email|mail|mailer|news|newsletter|notify|notification|marketing|mkt)|(?:email|mail|mailer|news|newsletter|notify|notification|marketing|mkt)$/;
  let previous;
  do {
    previous = value;
    value = value.replace(noise, "");
  } while (value !== previous);
  return value.length >= 4 ? value : "";
}

function senderGroupingKey(sender) {
  const domain = registrableSenderDomain(sender.address);
  if (!domain) return `address:${sender.address}`;
  const brand = senderBrandName(sender.name);
  const domainBrand = senderDomainBrand(sender.address);
  const compactBrand = brand.replace(/\s+/g, "");
  if (domainBrand && !MULTI_SERVICE_SENDER_DOMAINS.has(domain) && brand.length >= 3 && !GENERIC_SENDER_NAMES.has(brand) && compactBrand === domainBrand) return `organisation:${domainBrand}`;
  if (brand.length >= 3 && !GENERIC_SENDER_NAMES.has(brand)) return `brand:${domain}:${brand}`;
  return `address:${sender.address}`;
}

function groupRelatedSenders(exactSenders, provider) {
  const groups = new Map();
  for (const sender of exactSenders) {
    const key = senderGroupingKey(sender);
    const members = groups.get(key) || [];
    members.push(sender);
    groups.set(key, members);
  }
  return [...groups.entries()].map(([key, members], index) => {
    if (members.length === 1) return members[0];
    const primary = [...members].sort((a, b) => b.count - a.count)[0];
    const display = [...members].sort((a, b) => String(a.name || "").length - String(b.name || "").length || b.count - a.count)[0];
    const allCanUnsubscribe = members.every((member) => member.canUnsubscribe && member.unsubscribeUrl);
    const unsubscribeMember = allCanUnsubscribe ? members[0] : null;
    return {
      ...primary,
      id: `${provider}-group-${index}`,
      name: display.name,
      address: primary.address,
      addresses: senderAddresses({ members }),
      count: members.reduce((sum, member) => sum + member.count, 0),
      category: unsubscribeMember ? "unsubscribe" : "review",
      future: unsubscribeMember ? ["unsubscribe"] : ["review"],
      current: "keep",
      reasonKey: unsubscribeMember ? "gmailNewsletterReason" : "gmailUncertainReason",
      canUnsubscribe: allCanUnsubscribe,
      messageIds: members.flatMap((member) => member.messageIds || []),
      unsubscribeUrl: unsubscribeMember?.unsubscribeUrl || null,
      unsubscribeMessageId: unsubscribeMember?.unsubscribeMessageId || primary.messageIds?.[0],
      oneClick: Boolean(unsubscribeMember?.oneClick),
      members
    };
  }).sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }) || b.count - a.count);
}

async function gmailFetch(path, params = {}) {
  return gmailRequest(path, { params });
}

async function gmailRequest(path, { method = "GET", params = {}, body } = {}) {
  const url = new URL(`https://gmail.googleapis.com/gmail/v1/${path}`);
  Object.entries(params).forEach(([key, value]) => Array.isArray(value) ? value.forEach((item) => url.searchParams.append(key, item)) : url.searchParams.set(key, value));
  const headers = { Authorization: `Bearer ${gmailAccessToken}` };
  if (body !== undefined) headers["Content-Type"] = "application/json";
  const response = await fetch(url, { method, headers, body: body === undefined ? undefined : JSON.stringify(body) });
  if (!response.ok) {
    const detail = await response.json().catch(() => ({}));
    throw new Error(detail.error?.message || `Gmail request failed (${response.status})`);
  }
  if (response.status === 204) return {};
  const text = await response.text();
  return text ? JSON.parse(text) : {};
}

function headerValues(message, name) {
  const lowerName = name.toLowerCase();
  return (message.payload?.headers || []).filter((header) => header.name.toLowerCase() === lowerName).map((header) => header.value);
}

function authenticatedByGmail(message) {
  const results = headerValues(message, "Authentication-Results")
    .filter((value) => /^\s*mx\.google\.com\s*;/i.test(value))
    .join(" ")
    .toLowerCase();
  return /\bdmarc=pass\b/.test(results) || (/\bdkim=pass\b/.test(results) && /\bspf=pass\b/.test(results));
}

function safeHttpsUnsubscribe(message) {
  if (!authenticatedByGmail(message)) return null;
  const raw = headerValues(message, "List-Unsubscribe").join(",");
  const bracketed = [...raw.matchAll(/<([^>]+)>/g)].map((match) => match[1].trim());
  const candidates = bracketed.length ? bracketed : raw.split(",").map((value) => value.trim());
  for (const candidate of candidates) {
    try {
      const url = window.UnsubscribeSafety.publicHttpsUrl(candidate);
      if (url) {
        return {
          url: url.href,
          oneClick: headerValues(message, "List-Unsubscribe-Post").some((value) => /list-unsubscribe\s*=\s*one-click/i.test(value))
        };
      }
    } catch (_) { /* Ignore malformed or non-URL unsubscribe values. */ }
  }
  return null;
}

function outlookHeaderValues(message, name) {
  const lowerName = name.toLowerCase();
  return (message.internetMessageHeaders || [])
    .filter((header) => String(header.name || "").toLowerCase() === lowerName)
    .map((header) => String(header.value || ""));
}

function domainsAlign(candidate = "", senderDomain = "") {
  const cleanValue = candidate.toLowerCase().replace(/[<>]/g, "").trim();
  const cleanCandidate = (cleanValue.includes("@") ? cleanValue.split("@").pop() : cleanValue).replace(/[^a-z0-9.-].*$/, "");
  const cleanSender = senderDomain.toLowerCase();
  return Boolean(cleanCandidate && cleanSender && (cleanCandidate === cleanSender || cleanSender.endsWith(`.${cleanCandidate}`)));
}

function authenticatedByOutlook(message, address) {
  const senderDomain = address.split("@")[1] || "";
  const results = outlookHeaderValues(message, "Authentication-Results").join(" ").toLowerCase();
  if (!/\bdmarc=pass\b/.test(results) || !/\bcompauth=pass\b/.test(results)) return false;
  const dmarcDomains = [...results.matchAll(/\bheader\.from=([^\s;]+)/g)].map((match) => match[1]);
  const dkimDomains = [...results.matchAll(/\bheader\.d=([^\s;]+)/g)].map((match) => match[1]);
  const spfDomains = [...results.matchAll(/\bsmtp\.mailfrom=([^\s;]+)/g)].map((match) => match[1]);
  return dmarcDomains.some((domain) => domainsAlign(domain, senderDomain))
    && ((/\bdkim=pass\b/.test(results) && dkimDomains.some((domain) => domainsAlign(domain, senderDomain)))
      || (/\bspf=pass\b/.test(results) && spfDomains.some((domain) => domainsAlign(domain, senderDomain))));
}

function safeOutlookUnsubscribe(message, address) {
  if (!authenticatedByOutlook(message, address)) return null;
  const raw = outlookHeaderValues(message, "List-Unsubscribe").join(",");
  const bracketed = [...raw.matchAll(/<([^>]+)>/g)].map((match) => match[1].trim());
  const candidates = bracketed.length ? bracketed : raw.split(",").map((value) => value.trim());
  for (const candidate of candidates) {
    try {
      const url = window.UnsubscribeSafety.publicHttpsUrl(candidate);
      if (url) {
        return {
          url: url.href,
          oneClick: outlookHeaderValues(message, "List-Unsubscribe-Post").some((value) => /list-unsubscribe\s*=\s*one-click/i.test(value))
        };
      }
    } catch (_) { /* Ignore malformed, mailto or otherwise unsafe unsubscribe values. */ }
  }
  return null;
}

async function outlookRequest(path, params = {}, { method = "GET", body, discardBody = false, write = false } = {}) {
  const readPath = /^me\/(?:messages\/[^/?#]+|mailFolders\/(?:inbox|archive|deleteditems)(?:\/messages)?)$/;
  const movePath = /^me\/messages\/[^/?#]+\/move$/;
  if (!(method === "GET" && readPath.test(path)) &&
      !(method === "POST" && write && outlookActionBusy && movePath.test(path) && ["archive", "inbox", "deleteditems"].includes(body?.destinationId))) {
    throw new Error("OUTLOOK-REQUEST-NOT-ALLOWED");
  }
  const token = write ? outlookWriteAccessToken : outlookAccessToken;
  if (!token || !outlookAccount) throw new Error("OUTLOOK-NOT-CONNECTED");
  const url = new URL(`https://graph.microsoft.com/v1.0/${path}`);
  Object.entries(params).forEach(([key, value]) => url.searchParams.set(key, value));
  const response = await fetch(url, {
    method,
    headers: { Authorization: `Bearer ${token}`, Prefer: 'IdType="ImmutableId"', ...(body ? { "Content-Type": "application/json" } : {}) },
    body: body ? JSON.stringify(body) : undefined,
    credentials: "omit", redirect: "error", cache: "no-store", referrerPolicy: "no-referrer",
    signal: AbortSignal.timeout(20000)
  });
  if (!response.ok) {
    const detail = await response.json().catch(() => ({}));
    const requestError = new Error(`Microsoft Graph request failed (${response.status})`);
    const graphCode = String(detail.error?.code || "unknown").replace(/[^a-z0-9_-]/gi, "").slice(0, 48) || "unknown";
    requestError.outlookReference = `GRAPH-${response.status}-${graphCode}`;
    throw requestError;
  }
  if (discardBody) {
    // Graph move can return a full message. Never parse/use/store that response.
    // Verify the destination separately using a metadata-only GET and immutable ID.
    try { await response.body?.cancel(); } catch (_) { /* Response is intentionally discarded. */ }
    return {};
  }
  return response.json();
}

function outlookRuleAddressAllowed(address) {
  const normalised = String(address || "").trim().toLowerCase();
  const context = autoUnsubscribeContext;
  return Boolean(normalised && outlookActionBusy && autoCurrent(context) && context.actions.some((action) =>
    action.future.includes("auto_trash") && action.members.some((member) => String(member.address || "").trim().toLowerCase() === normalised)
  ));
}

async function createOutlookDeletedItemsRule(address) {
  const normalised = String(address || "").trim().toLowerCase();
  if (!outlookRuleAddressAllowed(normalised)) throw new Error("OUTLOOK-RULE-NOT-ALLOWED");
  if (!outlookRuleAccessToken || !outlookAccount) throw new Error("OUTLOOK-NOT-CONNECTED");
  const response = await fetch("https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules", {
    method: "POST",
    headers: { Authorization: `Bearer ${outlookRuleAccessToken}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      displayName: `WeHateSpam · ${normalised}`,
      sequence: 1,
      isEnabled: true,
      conditions: { fromAddresses: [{ emailAddress: { address: normalised } }] },
      actions: { delete: true, stopProcessingRules: true }
    }),
    credentials: "omit", redirect: "error", cache: "no-store", referrerPolicy: "no-referrer",
    signal: AbortSignal.timeout(20000)
  });
  if (!response.ok) {
    try { await response.body?.cancel(); } catch (_) { /* Provider response is intentionally discarded. */ }
    throw new Error("OUTLOOK-RULE-CREATE-FAILED");
  }
  const result = await response.json();
  if (!result || typeof result.id !== "string" || !result.id) throw new Error("OUTLOOK-RULE-UNCONFIRMED");
  return { id: result.id, address: normalised };
}

async function readOutlookMetadata() {
  const response = await outlookRequest("me/mailFolders/inbox/messages", {
    "$top": "100",
    "$orderby": "receivedDateTime desc",
    "$select": "id,from,sender"
  });
  const messages = response.value || [];
  const grouped = new Map();

  messages.forEach((message) => {
    const identity = message.from?.emailAddress || message.sender?.emailAddress || {};
    const address = String(identity.address || "").trim().toLowerCase();
    if (!address.includes("@")) return;
    const name = String(identity.name || address.split("@")[0]).trim() || address;
    const existing = grouped.get(address) || { name, address, count: 0, messageIds: [], unsubscribe: null };
    existing.count += 1;
    existing.messageIds.push(message.id);
    grouped.set(address, existing);
  });

  const groupedSenders = [...grouped.values()];
  let cursor = 0;
  const worker = async () => {
    while (cursor < groupedSenders.length) {
      const sender = groupedSenders[cursor++];
      const messageId = sender.messageIds[0];
      if (!messageId) continue;
      try {
        const message = await outlookRequest(`me/messages/${encodeURIComponent(messageId)}`, {
          "$select": "internetMessageHeaders"
        });
        sender.unsubscribe = safeOutlookUnsubscribe(message, sender.address);
      } catch (_) { /* A message can disappear or its headers can be unavailable while the inbox is scanned. */ }
    }
  };
  await Promise.all(Array.from({ length: Math.min(6, Math.max(1, groupedSenders.length)) }, worker));

  const colors = ["coral", "blue", "green"];
  const exactSenders = groupedSenders.sort((a, b) => b.count - a.count).map((sender, index) => ({
    id: `outlook-${index}-${sender.address}`,
    name: sender.name,
    address: sender.address,
    count: sender.count,
    category: sender.unsubscribe ? "unsubscribe" : "review",
    future: ["review"],
    current: "keep",
    reasonKey: sender.unsubscribe ? "gmailNewsletterReason" : "gmailUncertainReason",
    color: colors[index % colors.length],
    initial: sender.name.charAt(0).toUpperCase(),
    safe: false,
    canUnsubscribe: Boolean(sender.unsubscribe),
    messageIds: sender.messageIds,
    unsubscribeUrl: sender.unsubscribe?.url || null,
    unsubscribeMessageId: sender.unsubscribe?.messageId || sender.messageIds[0],
    oneClick: Boolean(sender.unsubscribe?.oneClick)
  }));
  const realSenders = groupRelatedSenders(exactSenders, "outlook");
  return {
    profile: { emailAddress: outlookAccount?.username || "Outlook" },
    scanned: messages.length,
    senders: realSenders
  };
}

async function readGmailMetadata() {
  const profile = await gmailFetch("users/me/profile");
  const listed = await gmailFetch("users/me/messages", { labelIds: "INBOX", maxResults: "100" });
  const ids = (listed.messages || []).map((message) => message.id);
  const headers = ["From", "Subject", "Date", "Authentication-Results", "List-Unsubscribe", "List-Unsubscribe-Post"];
  const details = [];
  let cursor = 0;
  const worker = async () => {
    while (cursor < ids.length) {
      const id = ids[cursor++];
      try {
        details.push(await gmailFetch(`users/me/messages/${id}`, { format: "metadata", metadataHeaders: headers }));
      } catch (_) { /* A message can disappear while the inbox is being scanned. */ }
    }
  };
  await Promise.all(Array.from({ length: Math.min(6, Math.max(1, ids.length)) }, worker));

  const grouped = new Map();
  details.forEach((message) => {
    const parsed = parseFromHeader(headerValues(message, "From")[0]);
    if (!parsed.address || !parsed.address.includes("@")) return;
    const existing = grouped.get(parsed.address) || { ...parsed, count: 0, messageIds: [], unsubscribe: null };
    existing.count += 1;
    existing.messageIds.push(message.id);
    const safeUnsubscribe = safeHttpsUnsubscribe(message);
    if (safeUnsubscribe && (!existing.unsubscribe || Number(message.internalDate || 0) > existing.unsubscribe.internalDate)) {
      existing.unsubscribe = { ...safeUnsubscribe, messageId: message.id, internalDate: Number(message.internalDate || 0) };
    }
    grouped.set(parsed.address, existing);
  });

  const colors = ["coral", "blue", "green"];
  const exactSenders = [...grouped.values()].sort((a, b) => b.count - a.count).map((sender, index) => ({
    id: `gmail-${index}-${sender.address}`,
    name: sender.name,
    address: sender.address,
    count: sender.count,
    category: sender.unsubscribe ? "unsubscribe" : "review",
    future: sender.unsubscribe ? ["unsubscribe"] : ["review"],
    current: "keep",
    reasonKey: sender.unsubscribe ? "gmailNewsletterReason" : "gmailUncertainReason",
    color: colors[index % colors.length],
    initial: sender.name.charAt(0).toUpperCase(),
    safe: false,
    canUnsubscribe: Boolean(sender.unsubscribe),
    messageIds: sender.messageIds,
    unsubscribeUrl: sender.unsubscribe?.url || null,
    unsubscribeMessageId: sender.unsubscribe?.messageId || sender.messageIds[0],
    oneClick: Boolean(sender.unsubscribe?.oneClick)
  }));
  return { profile, scanned: ids.length, senders: groupRelatedSenders(exactSenders, "gmail") };
}

function renderMailboxPreview({ profile, scanned, senders: realSenders }, provider = "gmail") {
  const isOutlook = provider === "outlook";
  const result = { profile, scanned, senders: realSenders };
  if (isOutlook) outlookPreviewResult = result;
  else gmailPreviewResult = result;
  $("#analysis-loading").hidden = true;
  $("#dashboard").hidden = true;
  $("#success-view").hidden = true;
  $("#gmail-preview").hidden = false;
  $("#account-label").textContent = t("connectedAccount");
  $("#account-provider").textContent = `${isOutlook ? "Outlook" : "Gmail"} · ${profile.emailAddress || t("connected")}`;
  $("#disconnect-button").textContent = t("disconnectAccount");
  $("#gmail-preview-title").textContent = t(isOutlook ? "outlookConnectedTitle" : "gmailConnectedTitle");
  $("#gmail-account-copy").textContent = t(isOutlook ? "outlookConnectedCopy" : "gmailConnectedCopy", { email: profile.emailAddress || (isOutlook ? "Outlook" : "Gmail"), count: scanned });
  $(".gmail-readonly-note").textContent = t(isOutlook ? "outlookReadOnlyNotice" : "gmailReadOnlyNotice");
  $("#open-gmail-dashboard").textContent = t(isOutlook ? "outlookReviewSenders" : "reviewAndClean");

  const stats = $("#gmail-preview-stats");
  stats.replaceChildren();
  [[scanned, t("gmailScannedMessages")], [realSenders.length, t("gmailSendersFound")], [realSenders.filter((sender) => sender.canUnsubscribe).length, t("gmailNewslettersFound")]].forEach(([value, label]) => {
    const item = document.createElement("div");
    const number = document.createElement("strong");
    number.textContent = String(value);
    const caption = document.createElement("small");
    caption.textContent = label;
    item.append(number, caption);
    stats.append(item);
  });

  const senderPreview = $("#gmail-sender-preview");
  senderPreview.replaceChildren();
  const title = document.createElement("strong");
  title.textContent = t("gmailTopSenders");
  senderPreview.append(title);
  if (!realSenders.length) {
    const empty = document.createElement("p");
    empty.textContent = t("gmailNoMessages");
    senderPreview.append(empty);
  } else {
    realSenders.slice(0, 8).forEach((sender) => {
      const row = document.createElement("div");
      row.className = "gmail-sender-row";
      const identity = document.createElement("span");
      const addresses = senderAddresses(sender);
      identity.textContent = addresses.length > 1
        ? `${sender.name} · ${t("senderAddressCount", { count: addresses.length })}`
        : `${sender.name} · ${sender.address}`;
      const count = document.createElement("b");
      count.textContent = `${sender.count}`;
      row.append(identity, count);
      senderPreview.append(row);
    });
  }
}

function renderGmailPreview(result) { renderMailboxPreview(result, "gmail"); }
function renderOutlookPreview(result) { renderMailboxPreview(result, "outlook"); }

$("#open-gmail-dashboard").addEventListener("click", showDashboard);

function startGmailOAuth() {
  const status = $("#provider-status");
  if (!window.google?.accounts?.oauth2) {
    showProviderError(t("gmailUnavailableCopy"));
    return;
  }
  state.provider = "gmail";
  status.textContent = t("gmailAuthorisationOpening");
  gmailTokenClient ||= window.google.accounts.oauth2.initTokenClient({
    client_id: GMAIL_CLIENT_ID,
    scope: GMAIL_METADATA_SCOPE,
    include_granted_scopes: false,
    callback: async (response) => {
      if (response.error) {
        showProviderError(t("gmailPermissionDenied"));
        return;
      }
      gmailAccessToken = response.access_token;
      enterGmailLoading();
      try {
        const result = await readGmailMetadata();
        setSenders(result.senders, "gmail", result.profile.emailAddress);
        renderGmailPreview(result);
      } catch (_) {
        gmailAccessToken = null;
        resetDemo();
        showProviderError(t("gmailConnectionErrorCopy"));
      }
    }
  });
  gmailTokenClient.requestAccessToken({ prompt: "consent" });
}

function connectGmail() {
  if (!safeMailContext()) return;
  if (providerDialog.open) providerDialog.close();
  startGmailOAuth();
}

function disconnectGmail() {
  const token = gmailAccessToken;
  gmailAccessToken = null;
  gmailPreviewResult = null;
  gmailActionResult = null;
  if (token && window.google?.accounts?.oauth2?.revoke) window.google.accounts.oauth2.revoke(token, () => {});
  resetDemo();
}

async function getOutlookClient() {
  if (!window.msal?.PublicClientApplication) throw new Error(t("outlookUnavailableCopy"));
  if (!outlookClient) {
    outlookClient = new window.msal.PublicClientApplication({
      auth: {
        clientId: OUTLOOK_CLIENT_ID,
        authority: "https://login.microsoftonline.com/common",
        redirectUri: `${window.location.origin}/auth-redirect.html`,
        navigateToLoginRequestUrl: false
      },
      cache: { cacheLocation: "memoryStorage", storeAuthStateInCookie: false }
    });
    outlookInitialisation = outlookClient.initialize();
  }
  await outlookInitialisation;
  return outlookClient;
}

async function connectOutlook() {
  if (!safeMailContext()) return;
  outlookSession += 1;
  outlookWriteAccessToken = null;
  outlookRuleAccessToken = null;
  outlookPlan = null;
  outlookActionResult = null;
  if (providerDialog.open) providerDialog.close();
  state.provider = "outlook";
  let outlookStage = "AUTH-INIT";
  try {
    const client = await getOutlookClient();
    outlookStage = "SIGN-IN";
    const authentication = await client.loginPopup({ scopes: OUTLOOK_SCOPES, prompt: "select_account" });
    outlookAccount = authentication.account;
    outlookAccessToken = authentication.accessToken;
    if (!outlookAccessToken) {
      outlookStage = "TOKEN";
      const token = await client.acquireTokenSilent({ scopes: OUTLOOK_SCOPES, account: outlookAccount });
      outlookAccessToken = token.accessToken;
    }
    outlookStage = "MAIL-LIST";
    enterOutlookLoading();
    const result = await readOutlookMetadata();
    setSenders(result.senders, "outlook", result.profile.emailAddress);
    renderOutlookPreview(result);
  } catch (error) {
    const failedAccount = outlookAccount;
    try { if (outlookClient && failedAccount) await outlookClient.clearCache({ account: failedAccount }); }
    catch (_) { /* The access token variables are still cleared below. */ }
    outlookAccessToken = null;
    outlookAccount = null;
    resetDemo();
    const rawReference = error?.outlookReference || `${outlookStage}-${error?.errorCode || error?.name || "unknown"}`;
    const reference = String(rawReference).replace(/[^a-z0-9_-]/gi, "").slice(0, 80) || "OUTLOOK-UNKNOWN";
    showProviderError(error?.errorCode === "user_cancelled" ? t("outlookPermissionDenied") : `${t("outlookConnectionErrorCopy")} Ref: ${reference}.`);
  }
}

async function disconnectOutlook() {
  if (outlookActionBusy) return;
  outlookSession += 1;
  outlookWriteAccessToken = null;
  outlookRuleAccessToken = null;
  outlookPlan = null;
  outlookActionResult = null;
  const account = outlookAccount;
  outlookAccessToken = null;
  outlookAccount = null;
  outlookPreviewResult = null;
  try { if (outlookClient && account) await outlookClient.clearCache({ account }); }
  catch (_) { /* The in-memory token is cleared below even if MSAL cleanup fails. */ }
  resetDemo();
}

$("#gmail-disconnect").addEventListener("click", () => state.mode === "outlook" ? disconnectOutlook() : disconnectGmail());

function visibleSenders() {
  return senders.filter((sender) => state.filter === "all" || sender.category === state.filter);
}

function syncSelectAll() {
  const checkbox = $("#select-all");
  const visible = visibleSenders();
  const complete = visible.filter((sender) => selectedSenderMembers(sender).length === senderMembers(sender).length).length;
  const any = visible.some((sender) => selectedSenderMembers(sender).length > 0);
  checkbox.checked = visible.length > 0 && complete === visible.length;
  checkbox.indeterminate = any && complete < visible.length;
}

function renderSenders() {
  const visible = visibleSenders();
  $("#visible-count").textContent = t(visible.length === 1 ? "senderCount" : "senderCountPlural", { count: visible.length });
  const items = visible.map((sender) => {
    const article = document.createElement("article");
    article.className = "sender-item";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    const members = senderMembers(sender);
    const selectedMembers = selectedSenderMembers(sender);
    checkbox.checked = selectedMembers.length > 0 && selectedMembers.length === members.length;
    checkbox.indeterminate = selectedMembers.length > 0 && selectedMembers.length < members.length;
    checkbox.disabled = false;
    checkbox.dataset.select = sender.id;
    checkbox.setAttribute("aria-label", t("selectSender", { name: sender.name }));
    checkbox.addEventListener("change", () => {
      setSenderSelection(sender, checkbox.checked);
      renderSenders();
      updateBar();
    });

    const avatar = document.createElement("span");
    avatar.className = `avatar ${sender.color}`;
    avatar.setAttribute("aria-hidden", "true");
    avatar.textContent = sender.initial;

    const identity = document.createElement("div");
    identity.className = "sender-identity";
    const nameLine = document.createElement("div");
    nameLine.className = "sender-name-line";
    const name = document.createElement("strong");
    name.textContent = sender.name;
    const count = document.createElement("span");
    const countDisplay = senderEmailCountDisplay(sender);
    count.className = `inline-count${countDisplay.partial ? " partial" : ""}`;
    const countNumber = document.createElement("b");
    countNumber.textContent = countDisplay.value;
    count.append(countNumber, document.createTextNode(` ${countDisplay.label}`));
    nameLine.append(name, count);
    const address = document.createElement("small");
    const addresses = senderAddresses(sender);
    if (addresses.length > 1) {
      const addressDetails = document.createElement("details");
      addressDetails.className = "sender-addresses";
      addressDetails.open = openAddressGroups.has(sender.id);
      addressDetails.addEventListener("toggle", () => addressDetails.open ? openAddressGroups.add(sender.id) : openAddressGroups.delete(sender.id));
      const addressSummary = document.createElement("summary");
      addressSummary.textContent = selectedMembers.length && selectedMembers.length < members.length
        ? t("senderAddressesSelected", { count: selectedMembers.length, total: members.length })
        : t("senderAddressCount", { count: addresses.length });
      const addressList = document.createElement("div");
      addressList.className = "sender-address-list";
      const allLabel = document.createElement("label");
      allLabel.className = "sender-address-all";
      const allCheckbox = document.createElement("input");
      allCheckbox.type = "checkbox";
      allCheckbox.checked = selectedMembers.length === members.length;
      allCheckbox.indeterminate = selectedMembers.length > 0 && selectedMembers.length < members.length;
      const allText = document.createElement("span");
      allText.textContent = t("selectAllSenderAddresses");
      allLabel.append(allCheckbox, allText);
      addressList.append(allLabel);
      members.forEach((member) => {
        const item = document.createElement("label");
        const memberCheckbox = document.createElement("input");
        memberCheckbox.type = "checkbox";
        memberCheckbox.checked = selectedMembers.some((selected) => selected.id === member.id);
        memberCheckbox.setAttribute("aria-label", t("selectSenderAddress", { address: member.address }));
        const value = document.createElement("span");
        value.textContent = member.address + (member.futureTreated ? ` · ${t("senderAlreadyTreated")}` : "");
        item.append(memberCheckbox, value);
        memberCheckbox.addEventListener("change", () => {
          setMemberSelection(sender, member, memberCheckbox.checked);
          renderSenders();
          updateBar();
        });
        addressList.append(item);
      });
      allCheckbox.addEventListener("change", () => {
        setSenderSelection(sender, allCheckbox.checked);
        renderSenders();
        updateBar();
      });
      addressDetails.append(addressSummary, addressList);
      identity.append(nameLine, addressDetails);
    } else {
      address.textContent = sender.address;
      identity.append(nameLine, address);
    }
    if (sender.treatedAddressCount > 0 && !sender.futureTreated) {
      const treated = document.createElement("small");
      treated.className = "partially-treated";
      treated.textContent = t("senderAddressesAlreadyTreated", { count: sender.treatedAddressCount, total: addresses.length });
      identity.append(treated);
    }

    const reason = document.createElement("div");
    reason.className = "sender-reason";
    reason.append(senderStatus(sender));

    const makeActionGroup = (kind, labelText, options, selectedValue) => {
      const group = document.createElement("div");
      group.className = `action-group ${kind}-group`;
      const label = document.createElement("label");
      const selectId = `${kind}-${sender.id}`;
      label.htmlFor = selectId;
      label.textContent = labelText;
      const shell = document.createElement("div");
      shell.className = "select-shell";
      const select = document.createElement("select");
      select.id = selectId;
      select.className = kind === "current" ? "action-select current" : "action-select";
      select.dataset[kind] = sender.id;
      select.disabled = outlookActionBusy || (kind === "future" && sender.futureTreated);
      options.forEach(([value, optionLabel]) => {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = optionLabel;
        option.selected = selectedValue === value;
        select.append(option);
      });
      const arrow = document.createElement("span");
      arrow.setAttribute("aria-hidden", "true");
      arrow.textContent = "⌄";
      shell.append(select, arrow);
      group.append(label, shell);
      return group;
    };

    const futureGroup = makeActionGroup("future", t("futureEmails"), futureOptions(sender), futureValue(sender.id));
    const currentChoices = state.mode === "outlook" ? ["keep", "archive", "trash"] : Object.keys(currentLabelKeys);
    const currentGroup = makeActionGroup("current", t("existingEmails"), currentChoices.map((value) => [value, currentLabel(value)]), state.current[sender.id]);
    const note = document.createElement("span");
    note.className = "plan-note";
    note.textContent = planExplanation(sender.id);
    currentGroup.append(note);
    article.append(checkbox, avatar, identity, reason, futureGroup, currentGroup);
    return article;
  });
  $("#sender-list").replaceChildren(...items);
  $("#select-all").disabled = outlookActionBusy;

  $$('[data-future]').forEach((select) => select.addEventListener("change", () => {
    const actions = select.value.split("+");
    state.future[select.dataset.future] = new Set(actions);
    ensureSenderSelection(senders.find((sender) => sender.id === select.dataset.future));
    renderSenders(); updateBar();
  }));
  $$('[data-current]').forEach((select) => select.addEventListener("change", () => {
    state.current[select.dataset.current] = select.value;
    ensureSenderSelection(senders.find((sender) => sender.id === select.dataset.current));
    renderSenders(); updateBar();
  }));
  syncSelectAll();
}

$$('[data-filter]').forEach((button) => button.addEventListener("click", () => {
  state.filter = button.dataset.filter;
  $$('[data-filter]').forEach((item) => item.classList.toggle("active", item === button));
  renderSenders();
}));
$("#select-all").addEventListener("change", (event) => {
  if (outlookActionBusy) return;
  visibleSenders().forEach((sender) => setSenderSelection(sender, event.target.checked));
  renderSenders(); updateBar();
});

function stats() {
  const list = senders.filter((sender) => state.selected.has(sender.id));
  return { list, emails: list.reduce((sum, sender) => sum + selectedSenderMembers(sender).reduce((count, member) => count + member.count, 0), 0) };
}
function updateBar() {
  const { list, emails } = stats();
  const selectedSummaryKey = list.length === 1
    ? (emails === 1 ? "selectedSummaryOneEmail" : "selectedSummaryOne")
    : "selectedSummary";
  $("#selected-summary").textContent = list.length ? t(selectedSummaryKey, { count: list.length, emails }) : t("noAction");
  $("#selected-detail").textContent = t(list.length ? "selectedDetail" : "chooseSenders");
  $("#review-actions").disabled = state.mode === "outlook"
    ? !list.length || list.some((sender) => !["keep", "archive", "trash"].includes(state.current[sender.id]))
    : !list.length;
}

function actionPlan(sender) {
  const plan = document.createElement("div");
  plan.className = "review-plan";
  const name = document.createElement("b");
  const selectedCount = selectedSenderMembers(sender).length;
  const totalAddresses = senderMembers(sender).length;
  name.textContent = totalAddresses > 1 && selectedCount < totalAddresses
    ? `${sender.name} · ${t("senderAddressesSelected", { count: selectedCount, total: totalAddresses })}`
    : sender.name;
  const future = document.createElement("span");
  future.textContent = `${t("futurePrefix")}: ${futureSummary(sender.id)}`;
  const current = document.createElement("span");
  current.textContent = `${t("existingPrefix")}: ${currentLabel(state.current[sender.id])}`;
  plan.append(name, future, current);
  return plan;
}

function resultActionPlan(sender, context, cleanupBySender) {
  const plan = actionPlan(sender);
  const name = plan.childNodes[0];
  const targets = (context.results || context.rows || []).map((result, index) => ({ result, sender: context.senders[index] }))
    .filter((entry) => actionForTarget(context, entry.sender)?.sender === sender);
  const manualTargets = targets.filter(({ result }) => ["manual", "not_sent"].includes(result.status));
  const manual = manualTargets.length > 0;
  const cleanup = cleanupBySender.get(sender);
  const automatic = targets.some(({ result }) => result.status === "requested");
  const futureDone = !targets.length || targets.every(({ result, sender: target }) => result.status === "requested" || (["manual", "not_sent"].includes(result.status) && manualSenderProcessed(context, target)));
  const futureProcessed = targets.length > 0 && futureDone;
  const done = !cleanup?.errors.length && futureDone && (Boolean(targets.length) || Boolean(cleanup));
  if (!manual && !automatic && !cleanup) return plan;
  const icon = document.createElement("span");
  const visiblyComplete = done || futureProcessed;
  icon.className = "result-action-icon " + (visiblyComplete ? "done" : "manual");
  icon.setAttribute("aria-label", t(visiblyComplete ? "actionComplete" : "actionNeedsManual"));
  icon.title = t(visiblyComplete ? "actionComplete" : "actionNeedsManual");
  icon.textContent = visiblyComplete ? "✓" : "✋";
  name.replaceChildren(icon, document.createTextNode(sender.name));
  return plan;
}

function renderReviewSummary() {
  const { list, emails } = stats();
  const total = document.createElement("div");
  total.className = "review-line";
  const totalLabel = document.createElement("span");
  totalLabel.textContent = t("messagesCovered");
  const totalValue = document.createElement("b");
  totalValue.textContent = String(emails);
  total.append(totalLabel, totalValue);
  $("#review-summary").replaceChildren(...list.map(actionPlan), total);
}

function renderSuccessSummary() {
  const { list, emails } = stats();
  $("#success-report").hidden = false;
  $("#undo-outlook-archive").hidden = true;
  $("#result-unsubscribes").replaceChildren();
  $("#result-share").hidden = true;
  $("#result-manual-confirm").hidden = true;
  $("#success-copy").classList.remove("batch-metrics");
  $("#success-status-icon").textContent = "✓";
  $("#success-status-icon").classList.remove("pending");
  if (autoUnsubscribeContext?.phase === "done") { renderBatchResult(); return; }
  $("#undo-outlook-archive").hidden = true;
  $("#undo-outlook-status").textContent = "";
  if (state.mode === "outlook" && outlookActionResult) {
    renderOutlookArchiveResult();
  } else if (state.mode === "gmail" && gmailActionResult) {
    const completed = gmailActionResult.items.filter((item) => !item.errors.length).length;
    const failed = gmailActionResult.items.length - completed;
    $("#success-eyebrow").textContent = t(failed ? "cleanupPartlyComplete" : "cleanupComplete");
    $("#success-title").textContent = t(failed ? "realSuccessPartialTitle" : "realSuccessTitle");
    const realSummaryKey = failed
      ? "realTreatedPartialSummary"
      : (completed === 1
        ? (gmailActionResult.emailCount === 1 ? "realTreatedSummaryOneEmail" : "realTreatedSummaryOne")
        : "realTreatedSummary");
    $("#success-copy").textContent = t(realSummaryKey, { emails: gmailActionResult.emailCount, count: completed, failed });
    const reportLines = gmailActionResult.items.map((item) => {
      const line = actionPlan(item.sender);
      if (item.errors.length) {
        line.classList.add("action-failed");
        const error = document.createElement("small");
        error.className = "action-error";
        error.textContent = t("senderActionFailed");
        line.append(error);
      }
      return line;
    });
    $("#success-report").replaceChildren(...reportLines);
    $("#result-share").hidden = failed > 0 || !gmailActionResult.items.some((item) => !item.errors.length && (state.current[item.sender.id] !== "keep" || state.future[item.sender.id].has("block")));
    $("#restart-demo").textContent = t("scanAgain");
    $("#undo-notice").textContent = t("gmailUndoNotice");
  } else {
    $("#result-share").hidden = false;
    $("#success-eyebrow").textContent = t("demoComplete");
    $("#success-title").textContent = t("lighterInbox");
    const demoSummaryKey = list.length === 1
      ? (emails === 1 ? "treatedSummaryOneEmail" : "treatedSummaryOne")
      : "treatedSummary";
    $("#success-copy").textContent = t(demoSummaryKey, { emails, count: list.length });
    $("#success-report").replaceChildren(...list.map(actionPlan));
    $("#restart-demo").textContent = t("repeatDemo");
    $("#undo-notice").textContent = t("undoNotice");
  }
}

function requiredGmailScopes() {
  const { list } = stats();
  const scopes = [];
  if (list.some((sender) => state.current[sender.id] !== "keep")) scopes.push(GMAIL_MODIFY_SCOPE);
  if (list.some((sender) => state.future[sender.id].has("block"))) scopes.push(GMAIL_SETTINGS_SCOPE);
  return scopes;
}

function requestGmailPermissions(scopes) {
  if (!scopes.length) return Promise.resolve();
  const previousToken = gmailAccessToken;
  const mailboxSenders = senders;
  const expectedEmail = gmailPreviewResult?.profile?.emailAddress?.toLowerCase();
  return new Promise((resolve, reject) => {
    if (!window.google?.accounts?.oauth2) {
      reject(new Error(t("gmailUnavailableCopy")));
      return;
    }
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: GMAIL_CLIENT_ID,
      // Keep the read capability for identity and unsubscribe checks, without
      // silently importing unrelated permissions from an older grant.
      scope: [...new Set([GMAIL_METADATA_SCOPE, ...scopes])].join(" "),
      include_granted_scopes: false,
      async callback(response) {
        if (response.error || !response.access_token) {
          reject(new Error(t("gmailPermissionDenied")));
          return;
        }
        try {
          if (!expectedEmail || state.mode !== "gmail" || senders !== mailboxSenders || gmailAccessToken !== previousToken) throw new Error();
          const identity = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/profile?fields=emailAddress", {
            headers: { Authorization: `Bearer ${response.access_token}` }, credentials: "omit", redirect: "error", cache: "no-store", signal: AbortSignal.timeout(12000)
          });
          if (!identity.ok || (await identity.json()).emailAddress?.toLowerCase() !== expectedEmail || state.mode !== "gmail" || senders !== mailboxSenders || gmailAccessToken !== previousToken) throw new Error();
          gmailAccessToken = response.access_token;
          resolve();
        } catch (_) { reject(new Error(t("gmailPermissionDenied"))); }
      },
      error_callback() { reject(new Error(t("gmailPermissionDenied"))); }
    });
    client.requestAccessToken({ prompt: "" });
  });
}

async function batchModify(messageIds, addLabelIds = [], removeLabelIds = []) {
  if (!messageIds.length) return;
  await gmailRequest("users/me/messages/batchModify", {
    method: "POST",
    body: { ids: messageIds, addLabelIds, removeLabelIds }
  });
}

async function mapWithLimit(values, limit, callback) {
  let cursor = 0;
  const workers = Array.from({ length: Math.min(limit, values.length) }, async () => {
    while (cursor < values.length) {
      const index = cursor++;
      await callback(values[index], index);
    }
  });
  await Promise.all(workers);
}

async function applyExistingAction(sender, action) {
  const ids = sender.messageIds || [];
  if (!ids.length || action === "keep") return;
  if (action === "archive") {
    await batchModify(ids, [], ["INBOX"]);
    return;
  }
  if (action === "spam" || action === "spam_trash") {
    await batchModify(ids, ["SPAM"], ["INBOX"]);
    if (action === "spam") return;
  }
  if (action === "trash" || action === "spam_trash") {
    await mapWithLimit(ids, 4, (id) => gmailRequest(`users/me/messages/${id}/trash`, { method: "POST" }));
  }
}

function normaliseFilterSender(value = "") {
  return value.trim().replace(/^\(|\)$/g, "").toLowerCase();
}

async function applyGmailActions(context) {
  const actions = context.actions;
  const emails = actions.reduce((sum, entry) => sum + entry.messageIds.length, 0);
  const needsBlocking = actions.some((entry) => entry.future.includes("block"));
  const knownBlocked = new Set();
  if (needsBlocking) {
    const response = await gmailRequest("users/me/settings/filters");
    (response.filter || []).forEach((filter) => {
      if ((filter.action?.addLabelIds || []).includes("TRASH") && filter.criteria?.from) {
        knownBlocked.add(normaliseFilterSender(filter.criteria.from));
      }
    });
  }

  const items = [];
  for (const entry of actions) {
    if (!autoCurrent(context)) throw new Error("plan_invalid");
    const sender = entry.sender;
    const item = { sender, errors: [] };
    if (entry.future.includes("block")) {
      for (const address of senderAddresses({ members: entry.futureMembers })) {
        if (knownBlocked.has(normaliseFilterSender(address))) continue;
        try {
          await gmailRequest("users/me/settings/filters", {
            method: "POST",
            body: { criteria: { from: address }, action: { addLabelIds: ["TRASH"] } }
          });
          knownBlocked.add(normaliseFilterSender(address));
        } catch (error) { item.errors.push(error); }
      }
    }
    try { await applyExistingAction({ ...sender, messageIds: entry.messageIds }, entry.current); }
    catch (error) { item.errors.push(error); }
    items.push(item);
  }

  return {
    items,
    emailCount: emails,
    unsubscribes: []
  };
}

const autoToken = () => state.mode === "gmail" ? gmailAccessToken : outlookAccessToken;
const autoCurrent = (context) => context && autoUnsubscribeContext === context && state.mode === context.provider && autoToken() === context.token && context.actions.every((entry) => senders.includes(entry.sender)) && (context.provider !== "outlook" || (context.accountKey === outlookAccountKey() && context.session === outlookSession));
const autoReason = (code) => t({
  automatic_unavailable: "batchManualRequired", outlook_manual: "batchManualAvailable",
  link_missing: "autoNoLink", headers_ambiguous: "autoVerifyFailed", headers_invalid: "autoVerifyFailed",
  signature_unverified: "autoVerifyFailed", destination_unsupported: "autoDestination",
  reconnect_required: "autoReconnect", metadata_permission: "autoReconnect", message_changed: "autoChanged",
  plan_invalid: "autoChanged", plan_expired: "autoChanged", sender_changed: "autoChanged", message_missing: "autoChanged",
  limit_reached: "autoLimited", provider_limited: "autoLimited", sender_limited: "autoLimited",
  already_attempted: "autoAlready", request_rejected: "autoRejected", outcome_unknown: "autoUnknown",
  redirect_received: "autoUnknown", batch_stopped: "autoStopped", request_accepted: "batchAccepted",
  service_not_ready: "autoUnavailable", service_unavailable: "autoUnavailable"
}[code] || "autoVerifyFailed");

function pendingManualRows(context) {
  return (context.results || context.rows || []).map((result, index) => ({ result, sender: context.senders[index] }))
    .filter(({ result }) => ["manual", "not_sent"].includes(result.status));
}

function actionForTarget(context, target) {
  return context?.actions?.find((entry) => entry.sender === target || (entry.members || senderMembers(entry.sender)).includes(target));
}

function manualActionUrl(context, sender) {
  const action = actionForTarget(context, sender);
  if (context?.provider === "outlook" && action?.future.includes("block")) {
    return window.UnsubscribeSafety.publicHttpsUrl("https://outlook.live.com/mail/0/options/mail/junkEmail");
  }
  return window.UnsubscribeSafety.publicHttpsUrl(sender.unsubscribeUrl);
}

function allManualLinksOpened(context) {
  const rows = pendingManualRows(context);
  return rows.length > 0 && rows.every(({ sender }) => manualLinkOpened(context, sender));
}

function manualLinkOpened(context, sender) {
  const url = manualActionUrl(context, sender);
  return Boolean(url && context.openedManual?.get(sender) === url.href);
}

// External pages cannot report completion to us: remember only the user's confirmation.
function manualSenderProcessed(context, sender) {
  return Boolean(context.confirmedManual?.has(sender));
}

function confirmManualSenders(context, targets) {
  if (!autoCurrent(context) || !["review", "done"].includes(context.phase) || outlookActionBusy) return;
  const pending = pendingManualRows(context);
  for (const sender of targets) {
    if (!pending.some((entry) => entry.sender === sender) || !manualLinkOpened(context, sender)) continue;
    context.confirmedManual.add(sender);
    rememberTreatedSender(sender);
  }
  context.manualConfirmed = pending.length > 0 && pending.every(({ sender }) => manualSenderProcessed(context, sender));
}

function updateBatchConfirmation(context) {
  const ready = (context.rows || []).filter((r) => r.status === "ready").length;
  const cleanup = context.actions.some((a) => a.current !== "keep" || (context.provider === "gmail" && a.future.includes("block")) || (context.provider === "outlook" && a.future.includes("auto_trash")));
  const manualOnly = !ready && !cleanup;
  const button = $("#confirm-actions");
  button.hidden = manualOnly && !allManualLinksOpened(context);
  button.disabled = context.phase !== "review" || (manualOnly && !allManualLinksOpened(context));
  button.textContent = t(manualOnly ? "manualFinish" : ready && !cleanup ? "batchConfirm" : "confirmCleanup", { count: ready });
}

function setReviewLoading(key = null) {
  $("#review-loading").hidden = !key;
  if (key) $("#review-loading-copy").textContent = t(key);
}

function manualUnsubscribeLink(sender) {
  const context = autoUnsubscribeContext;
  const url = manualActionUrl(context, sender);
  if (!url) return null;
  const mode = state.mode;
  const blocksInOutlook = context?.provider === "outlook" && actionForTarget(context, sender)?.future.includes("block");
  let opened = context?.openedManual?.get(sender) === url.href;
  const link = document.createElement("a");
  link.className = "secondary-button unsubscribe-button";
  link.textContent = t(blocksInOutlook ? "outlookBlockOpen" : "batchManualOpen");
  link.href = url.href; link.target = "_blank"; link.rel = "noopener noreferrer";
  link.referrerPolicy = "no-referrer";
  const markOpened = () => {
    link.textContent = t("manualLinkOpened");
    link.setAttribute("aria-disabled", "true");
    link.tabIndex = -1;
    link.classList.add("manual-opened");
  };
  if (opened) markOpened();
  link.addEventListener("click", (event) => {
    if (opened || outlookActionBusy || state.mode !== mode || (context && (autoUnsubscribeContext !== context || !actionForTarget(context, sender))) || (!context && !senders.includes(sender)) || manualActionUrl(context, sender)?.href !== url.href) {
      event.preventDefault(); return;
    }
    opened = true; markOpened();
    if (context) {
      context.openedManual.set(sender, url.href);
      // Keep this anchor attached until the browser performs its default navigation.
      if (context.phase === "review") updateBatchConfirmation(context);
      if (context.phase === "done") {
        $("#result-manual-confirm").disabled = false;
        renderBatchMetrics(context);
      }
      // Refresh each sender's indicator after the link's default navigation.
      setTimeout(() => {
        if (autoUnsubscribeContext !== context) return;
        if (context.phase === "review") renderBatchReview();
        if (context.phase === "done") renderBatchResult();
      }, 0);
    }
  });
  return link;
}

function batchGroups(context, results = false) {
  const rows = context.results || context.rows || [];
  const groups = results
    ? [["batchProcessed", (r) => r.status === "requested"], ["batchManualGroup", (r) => ["manual", "not_sent"].includes(r.status)], ["batchUncertain", (r) => r.status === "uncertain"]]
    : [["batchReady", (r) => r.status === "ready"], ["batchManualGroup", (r) => r.status !== "ready"]];
  return groups.flatMap(([key, filter]) => {
    const entries = rows.map((result, index) => ({ result, sender: context.senders[index] })).filter(({ result, sender }) => filter(result) && !(manualSenderProcessed(context, sender) && ["manual", "not_sent"].includes(result.status)));
    if (!entries.length) return [];
    const section = document.createElement("section"); section.className = "batch-group";
    const remaining = entries.filter(({ sender }) => !manualSenderProcessed(context, sender)).length;
    const headingKey = key === "batchManualGroup" ? (remaining === 0 ? "batchProcessed" : remaining === 1 ? "batchManualGroupOne" : key) : key;
    const heading = document.createElement("h3"); heading.textContent = t(headingKey, { count: key === "batchManualGroup" && remaining ? remaining : entries.length }); section.append(heading);
    for (const { result, sender } of entries) {
      const row = document.createElement("div"); row.className = "batch-row";
      const identity = document.createElement("div");
      const name = document.createElement("strong"); name.textContent = sender.name;
      const address = document.createElement("small"); address.textContent = sender.address;
      identity.append(name, address);
      const note = document.createElement("p");
      // A verification limitation does not prove that the sender requires a manual flow.
      const manualStep = result.status === "manual" && ["headers_ambiguous", "headers_invalid", "signature_unverified", "destination_unsupported", "metadata_unavailable", "metadata_too_large"].includes(result.reason) && window.UnsubscribeSafety.publicHttpsUrl(sender.unsubscribeUrl);
      note.textContent = result.status === "ready" ? t("batchReadyNote") : result.reason === "outlook_block" ? t("outlookBlockNote") : manualStep ? t("batchManualAvailable") : autoReason(result.reason);
      identity.append(note); row.append(identity);
      if (["manual", "not_sent"].includes(result.status)) {
        const link = manualUnsubscribeLink(sender); if (link) row.append(link);
        if (manualLinkOpened(context, sender) && !manualSenderProcessed(context, sender)) {
          const confirm = document.createElement("button");
          confirm.type = "button";
          confirm.className = "secondary-button unsubscribe-button";
          confirm.textContent = t("manualConfirmSender");
          confirm.disabled = outlookActionBusy || !["review", "done"].includes(context.phase);
          confirm.addEventListener("click", () => {
            confirmManualSenders(context, [sender]);
            if (context.phase === "review") renderBatchReview();
            if (context.phase === "done") renderBatchResult();
          });
          row.append(confirm);
        }
      }
      section.append(row);
    }
    return [section];
  });
}

function renderBatchReview() {
  const context = autoUnsubscribeContext;
  if (!context) return;
  $("#review-unsubscribes").replaceChildren(...batchGroups(context));
  const loadingKey = ({ checking: "reviewLoading", authorising: "outlookPermissionOpening", sending: "applyingEmailActions" })[context.phase] || null;
  setReviewLoading(loadingKey);
  $("#review-warning").hidden = !context.senders.length;
  $("#review-warning").textContent = t("batchPrivacy");
  const privacyLink = document.createElement("a");
  privacyLink.href = "./privacy.html"; privacyLink.target = "_blank"; privacyLink.rel = "noopener noreferrer";
  privacyLink.textContent = t("batchPrivacyLink");
  $("#review-warning").append(" ", privacyLink);
  const outlookNeedsMailPermission = context.provider === "outlook" && context.actions.some((a) => a.current !== "keep") && !outlookWriteAccessToken;
  const outlookNeedsRulePermission = context.provider === "outlook" && context.actions.some((a) => a.future.includes("auto_trash")) && !outlookRuleAccessToken;
  if (outlookNeedsMailPermission || outlookNeedsRulePermission) {
    $("#review-warning").hidden = false;
    $("#review-warning").textContent = t(outlookNeedsRulePermission ? "outlookRulePermission" : "outlookBatchPermission");
  }
  updateBatchConfirmation(context);
  const reviewStatus = $("#review-action-status");
  reviewStatus.hidden = Boolean(loadingKey);
  reviewStatus.replaceChildren();
  const statusKey = context.error || "batchReviewNote";
  if (!loadingKey && context.phase === "review" && !context.error) {
    const icon = document.createElement("span");
    icon.className = "batch-review-icon";
    icon.setAttribute("aria-hidden", "true");
    icon.textContent = "i";
    reviewStatus.append(icon, document.createTextNode(t(statusKey)));
  } else if (!loadingKey) reviewStatus.textContent = t(statusKey);
  reviewStatus.classList.toggle("batch-info", !context.error);
  reviewStatus.classList.toggle("batch-review-note", context.phase === "review" && !context.error);
}

async function autoRequest(path, payload, context) {
  const controller = new AbortController();
  autoUnsubscribeAbort = controller;
  const timer = setTimeout(() => controller.abort(), 180000);
  try {
    const response = await fetch('/api/unsubscribe/' + path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload), credentials: 'same-origin', redirect: 'error', cache: 'no-store', signal: controller.signal });
    const value = await response.json();
    if (!autoCurrent(context)) throw new Error('plan_invalid');
    if (!response.ok) throw new Error(value.reason || 'service_unavailable');
    return value;
  } finally { clearTimeout(timer); }
}

async function prepareBatchReview(list) {
  clearAutoUnsubscribe();
  const actions = list.map((sender) => {
    const members = Object.freeze([...selectedSenderMembers(sender)]);
    return Object.freeze({
      sender,
      members,
      futureMembers: Object.freeze(members.filter((member) => !senderWasTreated(member))),
      current: state.current[sender.id],
      future: Object.freeze([...state.future[sender.id]]),
      messageIds: Object.freeze(members.flatMap((member) => member.messageIds || []))
    });
  });
  const futureTargets = actions.flatMap((action) => {
    const members = action.futureMembers;
    if (state.mode === "outlook" && action.future.includes("block")) return members;
    if (action.future.includes("unsubscribe")) return members.filter((member) => member.canUnsubscribe);
    return [];
  });
  const context = {
    provider: state.mode, token: autoToken(), accountKey: outlookAccountKey(), session: outlookSession, phase: "checking", openedManual: new Map(), confirmedManual: new Set(), manualConfirmed: false,
    actions,
    senders: [...new Set(futureTargets)], rows: []
  };
  autoUnsubscribeContext = context;
  renderReviewSummary(); renderBatchReview(); reviewDialog.showModal();
  if (context.senders.length > 50) {
    context.phase = "blocked"; context.error = "autoLimitSelection"; renderBatchReview(); return;
  }
  try {
    if (context.provider === "outlook") {
      context.outlookPlans = context.actions.filter((a) => a.current !== "keep").flatMap((a) => a.members.map((member) => ({
        sender: a.sender,
        plan: window.OutlookArchive.createPlan(member, context.accountKey, context.session, a.current)
      })));
    }
    if (context.provider === "gmail") {
      // Provider consent, when needed, happens before preparing the token-bound plan.
      const before = context.token;
      await requestGmailPermissions(requiredGmailScopes());
      if (autoUnsubscribeContext !== context || state.mode !== context.provider || !context.actions.every((entry) => senders.includes(entry.sender))) return;
      context.token = autoToken();
      if (!before || !context.token) throw new Error("reconnect_required");
    }
    context.permissionsReady = true;
    if (context.senders.length && context.provider === "gmail") {
      const value = await autoRequest("prepare", { provider: context.provider, token: context.token, items: context.senders.map((sender) => ({ messageId: sender.unsubscribeMessageId || sender.messageIds[0], sender: sender.address })) }, context);
      if (!Array.isArray(value.rows) || value.rows.length !== context.senders.length) throw new Error("service_unavailable");
      context.plan = value.plan; context.rows = value.rows;
    } else context.rows = context.senders.map((sender) => ({
      status: "manual",
      reason: actionForTarget(context, sender)?.future.includes("block") ? "outlook_block" : "outlook_manual"
    }));
    context.phase = "review";
  } catch (error) {
    if (autoUnsubscribeContext !== context) return;
    context.rows = context.senders.map(() => ({ status: "manual", reason: error.message }));
    context.phase = context.permissionsReady ? "review" : "blocked";
    context.error = error.message === "plan_invalid" ? "autoChanged"
      : ["sign_in_required", "reconnect_required", "metadata_permission"].includes(error.message) ? "autoReconnect"
      : ["limit_reached", "provider_limited"].includes(error.message) ? "autoLimited" : "autoUnavailable";
  }
  renderBatchReview();
}

function batchCleanupItems(context) {
  const outlookItems = [...(context.cleanupItems || []), ...(context.ruleItems || [])];
  const raw = (context.provider === "outlook" ? outlookItems : gmailActionResult?.items || [])
    .filter((item) => context.actions.some((a) => a.sender === item.sender && (a.current !== "keep" || a.future.includes("block") || a.future.includes("auto_trash"))));
  const grouped = new Map();
  for (const item of raw) {
    const combined = grouped.get(item.sender) || { sender: item.sender, errors: [], changedIds: [], ruleIds: [] };
    combined.errors.push(...(item.errors || []));
    combined.changedIds.push(...(item.changedIds || []));
    combined.ruleIds.push(...(item.ruleIds || []));
    grouped.set(item.sender, combined);
  }
  return [...grouped.values()];
}

function recordCompletedFutureTreatments(context) {
  if (!context || context.provider === "demo") return;
  const cleanup = new Map(batchCleanupItems(context).map((item) => [item.sender, item]));
  for (const action of context.actions) {
    const targets = (context.results || []).map((result, index) => ({ result, sender: context.senders[index] }))
      .filter(({ sender }) => actionForTarget(context, sender) === action);
    const cleanupItem = cleanup.get(action.sender);
    const blockDone = action.future.includes("block") && (context.provider === "gmail"
      ? Boolean(cleanupItem && !cleanupItem.errors.length)
      : Boolean(targets.length && targets.every(({ sender }) => manualSenderProcessed(context, sender))));
    if (blockDone) {
      action.futureMembers.forEach(rememberTreatedSender);
      continue;
    }
    if (action.future.includes("auto_trash")) {
      const ruleItem = (context.ruleItems || []).find((item) => item.sender === action.sender);
      const completedAddresses = new Set((ruleItem?.rules || []).map((rule) => rule.address));
      action.futureMembers.filter((member) => completedAddresses.has(String(member.address || "").trim().toLowerCase())).forEach(rememberTreatedSender);
      continue;
    }
    if (!action.future.includes("unsubscribe")) continue;
    targets.forEach(({ result, sender }) => {
      if (result.status === "requested" || (["manual", "not_sent"].includes(result.status) && manualSenderProcessed(context, sender))) {
        rememberTreatedSender(sender);
      }
    });
  }
}

function batchResultMetrics(context) {
  const rows = context.results || [];
  const manual = pendingManualRows(context);
  const cleanup = batchCleanupItems(context);
  const currentCleanup = context.provider === "outlook" ? (context.cleanupItems || []) : cleanup;
  const ruleItems = context.provider === "outlook" ? (context.ruleItems || []) : [];
  const changedIds = new Set();
  let blocked = 0;
  for (const item of currentCleanup) {
    if (item.changedIds) item.changedIds.forEach((id) => changedIds.add(id));
    if (item.errors.length) continue;
    const action = context.actions.find((a) => a.sender === item.sender);
    if (action.current !== "keep") action.messageIds.forEach((id) => changedIds.add(id));
    if (action.future.includes("block")) blocked++;
  }
  const automaticDeletedRules = ruleItems.reduce((count, item) => count + (item.ruleIds?.length || 0), 0);
  const uncertain = rows.filter((r) => r.status === "uncertain").length;
  const manualProcessed = manual.filter(({ sender }) => manualSenderProcessed(context, sender)).length;
  const manualRemaining = manual.length - manualProcessed;
  const cleanupFailures = context.provider === "outlook"
    ? currentCleanup.filter((item) => item.errors.length).length + ruleItems.filter((item) => item.errors.length).length
    : cleanup.filter((item) => item.errors.length).length;
  const remaining = manualRemaining + uncertain + cleanupFailures;
  const metrics = [];
  if (context.actions.some((a) => a.current !== "keep")) metrics.push(["resultEmailsTreated", changedIds.size]);
  if (context.provider === "gmail" && context.actions.some((a) => a.future.includes("block"))) metrics.push(["resultSendersBlocked", blocked]);
  if (context.provider === "outlook" && context.actions.some((a) => a.future.includes("auto_trash"))) metrics.push(["resultAutomaticDeletedRules", automaticDeletedRules]);
  if (rows.length) metrics.push(["resultAutomaticUnsubscribes", rows.filter((r) => r.status === "requested").length]);
  if (manual.length) {
    metrics.push(["resultManualPending", manualRemaining]);
    metrics.push(["resultManualStepsProcessed", manualProcessed]);
  }
  if (uncertain) metrics.push(["resultUncertainActions", uncertain]);
  metrics.push(["resultActionsRemaining", remaining]);
  return metrics;
}

function renderBatchMetrics(context) {
  $("#success-copy").classList.add("batch-metrics");
  $("#success-copy").replaceChildren(...batchResultMetrics(context).map(([key, value]) => {
    const line = document.createElement("span"); line.className = "batch-metric";
    const label = document.createElement("span"); label.textContent = t(key);
    const count = document.createElement("strong"); count.textContent = String(value);
    line.append(label, count); return line;
  }));
}

function renderBatchResult() {
  const context = autoUnsubscribeContext;
  const rows = context.results || [];
  const processed = rows.filter((r) => r.status === "requested").length;
  const manual = rows.filter((r) => ["manual", "not_sent"].includes(r.status)).length;
  $("#success-eyebrow").textContent = t("batchResults");
  $("#success-title").textContent = t("batchResultTitle");
  renderBatchMetrics(context);
  $("#result-unsubscribes").replaceChildren(...batchGroups(context, true));
  const cleanup = batchCleanupItems(context);
  const cleanupBySender = new Map(cleanup.map((item) => [item.sender, item]));
  const details = context.actions.map((action) => {
    const item = cleanupBySender.get(action.sender);
    const row = resultActionPlan(action.sender, context, cleanupBySender);
    if (!item) return row;
    if (item.errors.length) {
      const targets = (context.results || context.rows || []).map((result, index) => ({ result, sender: context.senders[index] }))
        .filter((entry) => actionForTarget(context, entry.sender)?.sender === action.sender);
      const futureProcessed = targets.length > 0 && targets.every(({ result, sender }) => result.status === "requested" || (["manual", "not_sent"].includes(result.status) && manualSenderProcessed(context, sender)));
      const error = document.createElement("small");
      error.className = "result-action-note";
      error.textContent = t(futureProcessed ? "senderActionPartial" : "senderActionFailed");
      row.append(error);
    }
    return row;
  });
  $("#success-report").replaceChildren(...details);
  $("#success-report").hidden = !details.length;
  const metrics = Object.fromEntries(batchResultMetrics(context));
  const pending = Number(metrics.resultActionsRemaining || 0) > 0;
  $("#success-title").textContent = t(pending ? "batchPendingTitle" : "batchResultTitle");
  $("#success-status-icon").textContent = pending ? "…" : "✓";
  $("#success-status-icon").classList.toggle("pending", pending);
  const finish = $("#result-manual-confirm");
  finish.hidden = !manual || context.manualConfirmed;
  finish.disabled = !pendingManualRows(context).some(({ sender }) => manualLinkOpened(context, sender) && !manualSenderProcessed(context, sender));
  finish.textContent = t("manualConfirmOpened");
  $("#result-share").hidden = pending || !(processed || pendingManualRows(context).some(({ sender }) => manualSenderProcessed(context, sender)) || cleanup.some((item) => !item.errors.length));
  $("#restart-demo").textContent = t("scanAgain");
  $("#undo-notice").textContent = cleanup.length ? t(context.provider === "outlook" ? "outlookBatchUndoNotice" : "gmailUndoNotice") : "";
}

$("#result-manual-confirm").addEventListener("click", () => {
  const context = autoUnsubscribeContext;
  if (!autoCurrent(context) || context.phase !== "done" || outlookActionBusy) return;
  confirmManualSenders(context, pendingManualRows(context).map(({ sender }) => sender));
  renderSuccessSummary();
  window.scrollTo({ top: 0, behavior: "smooth" });
});

async function confirmBatch() {
  const context = autoUnsubscribeContext;
  if (!autoCurrent(context) || context.phase !== "review" || outlookActionBusy || !safeMailContext()) return;
  // A review is a snapshot: changing dashboard choices requires a fresh review.
  if (state.selected.size !== context.actions.length || context.actions.some((a) => !state.selected.has(a.sender.id) || state.current[a.sender.id] !== a.current || [...state.future[a.sender.id]].join("+") !== a.future.join("+") || selectedSenderMembers(a.sender).map((member) => member.id).join("|") !== a.members.map((member) => member.id).join("|"))) {
    context.phase = "blocked"; context.error = "autoChanged"; renderBatchReview(); return;
  }
  const needsOutlookMailPermission = context.provider === "outlook" && context.outlookPlans?.length && !outlookWriteAccessToken;
  const needsOutlookRulePermission = context.provider === "outlook" && context.actions.some((action) => action.future.includes("auto_trash")) && !outlookRuleAccessToken;
  if (needsOutlookMailPermission || needsOutlookRulePermission) {
    context.phase = "authorising"; setOutlookBusy(true); renderBatchReview();
    try {
      await requestOutlookBatchPermissions(context);
      if (!autoCurrent(context)) { setOutlookBusy(false); return; }
      context.error = null;
    } catch (_) {
      if (autoCurrent(context)) { context.phase = "review"; context.error = "outlookArchivePreparationFailed"; renderBatchReview(); }
      setOutlookBusy(false);
      return;
    }
  }
  const manualOnly = !context.rows.some((row) => row.status === "ready") && !context.actions.some((a) => a.current !== "keep" || (context.provider === "gmail" && a.future.includes("block")) || (context.provider === "outlook" && a.future.includes("auto_trash")));
  if (manualOnly && !allManualLinksOpened(context)) return;
  if (manualOnly) confirmManualSenders(context, context.senders);
  context.phase = "sending"; setOutlookBusy(true); renderBatchReview();
  try {
    if (context.rows.some((row) => row.status === "ready")) {
      try {
        const value = await autoRequest("execute", { plan: context.plan, token: context.token, confirmed: true }, context);
        if (!Array.isArray(value.outcomes) || value.outcomes.length !== context.rows.length) throw new Error("service_unavailable");
        context.results = value.outcomes;
      } catch (_) {
        if (!autoCurrent(context)) return;
        context.results = context.rows.map((row) => row.status === "ready" ? { status: "uncertain", reason: "outcome_unknown" } : row);
      }
    } else context.results = context.rows;
    // Save accepted unsubscribe results before independent mailbox cleanup can fail.
    context.results.forEach((result, index) => {
      if (result.status === "requested") rememberTreatedSender(context.senders[index]);
    });
    context.plan = null;
    if (context.provider === "gmail") gmailActionResult = await applyGmailActions(context);
    else {
      if (context.outlookPlans?.length) await applyOutlookBatchActions(context);
      if (context.actions.some((action) => action.future.includes("auto_trash"))) await applyOutlookRules(context);
    }
    if (!autoCurrent(context)) return;
    context.phase = "done";
    recordCompletedFutureTreatments(context);
    reviewDialog.close(); $("#dashboard").hidden = true; $("#success-view").hidden = false;
    renderSuccessSummary(); window.scrollTo({ top: 0, behavior: "smooth" });
  } catch (_) {
    context.phase = "done";
    reviewDialog.close(); $("#dashboard").hidden = true; $("#success-view").hidden = false;
    renderSuccessSummary();
  } finally { setOutlookBusy(false); }
}

async function applyOutlookBatchActions(context) {
  // Reuse the verified single-sender mover for each immutable reviewed plan.
  context.cleanupItems = context.outlookPlans.map(({ sender }) => ({ sender, errors: ["not_started"], changedIds: [] }));
  try { await refreshOutlookWriteToken(); }
  catch (_) { return; }
  for (let index = 0; index < context.outlookPlans.length; index++) {
    if (!autoCurrent(context)) return;
    const item = context.cleanupItems[index];
    try {
      const result = await window.OutlookArchive.run(context.outlookPlans[index].plan, outlookArchiveRuntime());
      item.changedIds = result.entries.filter((entry) => ["archived", "trashed"].includes(entry.status)).map((entry) => entry.id);
      item.errors = item.changedIds.length === result.plan.ids.length ? [] : ["move_incomplete"];
      if (result.entries.some((entry) => entry.status === "unknown" || entry.reason === "CHECK-FAILED")) break;
    } catch (_) { item.errors = ["move_unconfirmed"]; break; }
  }
}

async function applyOutlookRules(context) {
  const targets = context.actions.filter((action) => action.future.includes("auto_trash"));
  context.ruleItems = targets.map((action) => ({ sender: action.sender, errors: ["not_started"], ruleIds: [], rules: [] }));
  try { await refreshOutlookRuleToken(); }
  catch (_) { return; }
  for (let index = 0; index < targets.length; index++) {
    if (!autoCurrent(context)) return;
    const action = targets[index];
    const item = context.ruleItems[index];
    item.errors = [];
    for (const member of action.futureMembers) {
      try {
        const rule = await createOutlookDeletedItemsRule(member.address);
        item.ruleIds.push(rule.id);
        item.rules.push(rule);
      } catch (_) {
        item.errors.push("rule_unconfirmed");
        break;
      }
    }
  }
}

function outlookAccountKey(account = outlookAccount) {
  return account?.homeAccountId ? `${account.homeAccountId}|${account.tenantId || ""}|${account.localAccountId || ""}` : "";
}

function currentOutlookSession() {
  return { accountKey: state.mode === "outlook" ? outlookAccountKey() : "", session: outlookSession };
}

function setOutlookBusy(busy) {
  outlookActionBusy = busy;
  ["#confirm-actions", "#cancel-actions", "#review-dialog .dialog-close", "#back-home", "#disconnect-button", "#restart-demo", "#undo-outlook-archive"]
    .forEach((selector) => { $(selector).disabled = busy; });
}

function outlookArchiveRuntime() {
  return {
    currentSession: currentOutlookSession,
    request: (path, { params, ...options } = {}) => outlookRequest(path, params, { ...options, write: true })
  };
}

function acceptOutlookWriteToken(response, expectedKey, expectedSession) {
  if (expectedKey !== outlookAccountKey(response.account) || expectedKey !== outlookAccountKey() || expectedSession !== outlookSession || state.mode !== "outlook") {
    throw new Error("SESSION-CHANGED");
  }
  if (!response.accessToken || !response.scopes?.some((scope) => /^(?:https:\/\/graph\.microsoft\.com\/)?mail\.readwrite$/i.test(scope))) {
    throw new Error("WRITE-PERMISSION-MISSING");
  }
  outlookWriteAccessToken = response.accessToken;
}

function acceptOutlookRuleToken(response, expectedKey, expectedSession) {
  if (expectedKey !== outlookAccountKey(response.account) || expectedKey !== outlookAccountKey() || expectedSession !== outlookSession || state.mode !== "outlook") {
    throw new Error("SESSION-CHANGED");
  }
  if (!response.accessToken || !response.scopes?.some((scope) => /^(?:https:\/\/graph\.microsoft\.com\/)?mailboxsettings\.readwrite$/i.test(scope))) {
    throw new Error("RULE-PERMISSION-MISSING");
  }
  outlookRuleAccessToken = response.accessToken;
}

async function requestOutlookArchivePermission() {
  const expectedKey = outlookAccountKey();
  const expectedSession = outlookSession;
  const client = await getOutlookClient();
  const response = await client.acquireTokenPopup({ scopes: OUTLOOK_WRITE_SCOPES, account: outlookAccount, loginHint: outlookAccount.username });
  acceptOutlookWriteToken(response, expectedKey, expectedSession);
}

async function requestOutlookBatchPermissions(context) {
  const expectedKey = outlookAccountKey();
  const expectedSession = outlookSession;
  const needsMail = Boolean(context.outlookPlans?.length && !outlookWriteAccessToken);
  const needsRules = Boolean(context.actions.some((action) => action.future.includes("auto_trash")) && !outlookRuleAccessToken);
  const scopes = [...new Set([...(needsMail ? OUTLOOK_WRITE_SCOPES : []), ...(needsRules ? OUTLOOK_RULE_SCOPES : [])])];
  if (!scopes.length) return;
  const client = await getOutlookClient();
  const response = await client.acquireTokenPopup({ scopes, account: outlookAccount, loginHint: outlookAccount.username });
  if (needsMail) acceptOutlookWriteToken(response, expectedKey, expectedSession);
  if (needsRules) acceptOutlookRuleToken(response, expectedKey, expectedSession);
}

async function refreshOutlookWriteToken() {
  const expectedKey = outlookAccountKey();
  const expectedSession = outlookSession;
  // No surprise consent popup on execution/undo; renewed only after explicit permission.
  if (!outlookWriteAccessToken) throw new Error("WRITE-PERMISSION-MISSING");
  const response = await outlookClient.acquireTokenSilent({ scopes: OUTLOOK_WRITE_SCOPES, account: outlookAccount });
  acceptOutlookWriteToken(response, expectedKey, expectedSession);
}

async function refreshOutlookRuleToken() {
  const expectedKey = outlookAccountKey();
  const expectedSession = outlookSession;
  if (!outlookRuleAccessToken) throw new Error("RULE-PERMISSION-MISSING");
  const response = await outlookClient.acquireTokenSilent({ scopes: OUTLOOK_RULE_SCOPES, account: outlookAccount });
  acceptOutlookRuleToken(response, expectedKey, expectedSession);
}

function renderOutlookReview() {
  if (!outlookPlan) return;
  const isTrash = outlookPlan.action === "trash";
  const identity = document.createElement("p");
  identity.textContent = `${outlookPlan.name} · ${outlookPlan.address}`;
  const total = document.createElement("strong");
  total.textContent = t(isTrash ? "outlookTrashExactPlan" : "outlookExactPlan", { count: outlookPlan.ids.length });
  $("#review-summary").replaceChildren(identity, total);
  $("#review-warning").textContent = t(outlookWriteAccessToken ? (isTrash ? "outlookTrashFinalWarning" : "outlookFinalWarning") : "outlookWriteWarning");
  $("#confirm-actions").textContent = t("confirmCleanup");
  setReviewLoading();
}

function renderOutlookArchiveResult() {
  const result = outlookActionResult;
  const isTrash = result.plan.action === "trash";
  const archived = result.entries.filter((entry) => ["archived", "trashed"].includes(entry.status)).length;
  const unknown = result.entries.filter((entry) => entry.status === "unknown").length;
  const unchanged = result.plan.ids.length - archived - unknown;
  const restored = result.entries.filter((entry) => entry.restored).length;
  $("#result-share").hidden = !archived;
  $("#success-eyebrow").textContent = t("outlookTestComplete");
  $("#success-title").textContent = t(unknown || unchanged ? "outlookArchivePartial" : (isTrash ? "outlookTrashComplete" : "outlookArchiveComplete"));
  $("#success-copy").textContent = t(isTrash ? "outlookTrashCounts" : "outlookArchiveCounts", { archived, unchanged, unknown });
  const identity = document.createElement("p");
  identity.textContent = `${result.plan.name} · ${result.plan.address}`;
  $("#success-report").replaceChildren(identity);
  const manual = senders.filter((sender) => sender.address.toLowerCase() === result.plan.address.toLowerCase() && state.future[sender.id]?.has("unsubscribe"));
  $("#result-unsubscribes").replaceChildren(...batchGroups({ senders: manual, rows: manual.map(() => ({ status: "manual", reason: "outlook_manual" })) }, true));
  $("#restart-demo").textContent = t("scanAgain");
  $("#undo-notice").textContent = t(isTrash ? "outlookTrashUndoNotice" : "outlookUndoNotice");
  const undo = $("#undo-outlook-archive");
  undo.hidden = !archived || Boolean(result.undoFinished);
  undo.textContent = t("outlookUndoArchive");
  $("#undo-outlook-status").textContent = result.undoFinished ? t("outlookUndoCounts", { restored, remaining: archived - restored }) : "";
}

async function confirmOutlookArchive() {
  if (outlookActionBusy || !outlookPlan || !safeMailContext()) return;
  setOutlookBusy(true);
  try {
    if (!outlookWriteAccessToken) {
      $("#review-action-status").textContent = "";
      setReviewLoading("outlookPermissionOpening");
      await requestOutlookArchivePermission();
    }
    setReviewLoading("applyingEmailActions");
    await refreshOutlookWriteToken();
    outlookActionResult = await window.OutlookArchive.run(outlookPlan, outlookArchiveRuntime());
    outlookPlan = null;
    reviewDialog.close();
    $("#dashboard").hidden = true;
    $("#success-view").hidden = false;
    renderSuccessSummary();
    window.scrollTo({ top: 0, behavior: "smooth" });
  } catch (_) {
    // Do not expose OAuth responses, provider messages, tokens or message IDs.
    outlookPlan = null;
    setReviewLoading();
    $("#review-action-status").textContent = t("outlookArchivePreparationFailed");
  } finally {
    setOutlookBusy(false);
    if (!outlookPlan) $("#confirm-actions").disabled = true;
  }
}

$("#undo-outlook-archive").addEventListener("click", async () => {
  if (outlookActionBusy || !outlookActionResult || !safeMailContext()) return;
  setOutlookBusy(true);
  $("#undo-outlook-status").textContent = t("applyingActions");
  try {
    await refreshOutlookWriteToken();
    await window.OutlookArchive.undo(outlookActionResult, outlookArchiveRuntime());
    outlookActionResult.undoFinished = true;
    renderOutlookArchiveResult();
  } catch (_) { $("#undo-outlook-status").textContent = t(outlookActionResult?.plan.action === "trash" ? "outlookTrashUndoFailed" : "outlookUndoFailed"); }
  finally { setOutlookBusy(false); }
});

reviewDialog.addEventListener("cancel", (event) => { if (outlookActionBusy) event.preventDefault(); });
window.addEventListener("beforeunload", (event) => { if (outlookActionBusy) { event.preventDefault(); event.returnValue = ""; } });

$("#review-actions").addEventListener("click", async () => {
  if (outlookActionBusy) return;
  const { list } = stats();
  if (!list.length) return;
  $("#confirm-actions").hidden = false;
  $("#review-warning").hidden = false;
  $("#review-unsubscribes").replaceChildren();
  if (state.mode !== "demo" && !safeMailContext()) return;
  if (state.mode === "outlook" && list.some((sender) => state.current[sender.id] !== "keep")) {
    if (list.some((sender) => !["keep", "archive", "trash"].includes(state.current[sender.id]))) return;
    if (list.length > 1 || list.some((sender) => state.future[sender.id].has("block") || state.future[sender.id].has("unsubscribe") || state.future[sender.id].has("auto_trash"))) return prepareBatchReview(list);
    clearAutoUnsubscribe();
    outlookPlan = window.OutlookArchive.createPlan(list[0], outlookAccountKey(), outlookSession, state.current[list[0].id]);
    renderOutlookReview();
    // Outlook cleanup keeps its existing permission/confirmation contract.
    const manual = list.filter((sender) => state.future[sender.id].has("unsubscribe"));
    $("#review-unsubscribes").replaceChildren(...batchGroups({ senders: manual, rows: manual.map(() => ({ status: "manual", reason: "outlook_manual" })) }));
    $("#review-action-status").textContent = "";
    $("#confirm-actions").disabled = false; $("#cancel-actions").disabled = false;
    reviewDialog.showModal(); return;
  }
  if (state.mode !== "demo") return prepareBatchReview(list);
  renderReviewSummary();
  $("#review-warning").textContent = t("reviewWarning");
  $("#confirm-actions").textContent = t("confirmCleanup");
  $("#confirm-actions").disabled = false; $("#cancel-actions").disabled = false;
  $("#review-action-status").textContent = "";
  reviewDialog.showModal();
});
function cancelReview() {
  if (outlookActionBusy) return;
  clearAutoUnsubscribe(); reviewDialog.close();
}
$("#review-dialog .dialog-close").addEventListener("click", cancelReview);
$("#cancel-actions").addEventListener("click", cancelReview);
reviewDialog.addEventListener("close", () => {
  if (autoUnsubscribeContext && !["sending", "done"].includes(autoUnsubscribeContext.phase)) clearAutoUnsubscribe();
});
$("#confirm-actions").addEventListener("click", async () => {
  if (autoUnsubscribeContext) return confirmBatch();
  if (state.mode === "outlook") return confirmOutlookArchive();
  if (state.mode !== "demo") return;
  reviewDialog.close(); $("#dashboard").hidden = true; $("#success-view").hidden = false;
  renderSuccessSummary(); window.scrollTo({ top: 0, behavior: "smooth" });
});

document.addEventListener("wehatespam:languagechange", () => {
  const livePreview = state.mode === "gmail" ? gmailPreviewResult : (state.mode === "outlook" ? outlookPreviewResult : null);
  if (livePreview) {
    $("#account-label").textContent = t("connectedAccount");
    $("#account-provider").textContent = `${state.mode === "outlook" ? "Outlook" : "Gmail"} · ${livePreview.profile.emailAddress || t("connected")}`;
    if (!$("#gmail-preview").hidden) renderMailboxPreview(livePreview, state.mode);
  } else {
    $("#account-label").textContent = t("demoInbox");
    $("#account-provider").textContent = t("fictionalProvider", { provider: state.provider === "gmail" ? "Gmail" : "Outlook" });
  }
  if (!$("#dashboard").hidden) { updateDashboardCopy(); renderSenders(); updateBar(); }
  if (!$("#success-view").hidden) renderSuccessSummary();
  setShareLinks();
  $("#share-status").textContent = "";
  if (reviewDialog.open) {
    if (autoUnsubscribeContext) { renderReviewSummary(); renderBatchReview(); return; }
    if (state.mode === "outlook") { renderOutlookReview(); return; }
    renderReviewSummary();
    const scopes = state.mode === "gmail" ? requiredGmailScopes() : [];
    $("#review-warning").textContent = t(state.mode === "gmail" ? (scopes.length ? "gmailReviewWarning" : "gmailUnsubscribeWarning") : "reviewWarning");
    $("#confirm-actions").textContent = t(state.mode === "gmail" && scopes.length ? "authoriseAndClean" : "confirmCleanup");
  }
});

function resetDemo() {
  if (outlookActionBusy) return;
  outlookSession += 1;
  outlookWriteAccessToken = null;
  outlookRuleAccessToken = null;
  outlookPlan = null;
  outlookActionResult = null;
  gmailAccessToken = null;
  gmailPreviewResult = null;
  gmailActionResult = null;
  outlookAccessToken = null;
  outlookAccount = null;
  outlookPreviewResult = null;
  state.mode = "demo";
  state.provider = "gmail";
  setSenders(demoSenders, "demo", "");
  $("#app-view").hidden = true; $("#landing-view").hidden = false; $(".site-header").hidden = false;
  $("#gmail-preview").hidden = true;
  $("#dashboard").classList.remove("readonly-mode");
  $("#dashboard").classList.remove("outlook-test-mode");
  $("#provider-status").textContent = "";
  window.scrollTo({ top: 0, behavior: "smooth" });
}
$("#back-home").addEventListener("click", () => state.mode === "gmail" ? disconnectGmail() : (state.mode === "outlook" ? disconnectOutlook() : resetDemo()));
$("#disconnect-button").addEventListener("click", () => state.mode === "gmail" ? disconnectGmail() : (state.mode === "outlook" ? disconnectOutlook() : resetDemo()));
$("#restart-demo").addEventListener("click", async () => {
  if (outlookActionBusy) return;
  outlookActionResult = null;
  outlookPlan = null;
  if (state.mode === "demo") {
    startDemo(state.provider);
    return;
  }
  state.mode === "outlook" ? enterOutlookLoading() : enterGmailLoading();
  try {
    const result = state.mode === "outlook" ? await readOutlookMetadata() : await readGmailMetadata();
    setSenders(result.senders, state.mode, result.profile.emailAddress);
    gmailActionResult = null;
    state.mode === "outlook" ? renderOutlookPreview(result) : renderGmailPreview(result);
  } catch (_) {
    if (state.mode === "outlook") {
      await disconnectOutlook();
      showProviderError(t("outlookConnectionErrorCopy"));
    } else {
      disconnectGmail();
      showProviderError(t("gmailConnectionErrorCopy"));
    }
  }
});

function shareUrl() { return `${window.location.origin}${window.location.pathname}`; }
function shareMessage() { return t("shareMessage"); }
function setShareLinks() {
  const url = shareUrl();
  const message = shareMessage();
  const links = {
    whatsapp: `https://wa.me/?text=${encodeURIComponent(`${message} ${url}`)}`,
    email: `mailto:?subject=${encodeURIComponent(t("shareEmailSubject"))}&body=${encodeURIComponent(`${message}\n\n${url}`)}`,
    linkedin: `https://www.linkedin.com/sharing/share-offsite/?url=${encodeURIComponent(url)}`,
    x: `https://twitter.com/intent/tweet?text=${encodeURIComponent(message)}&url=${encodeURIComponent(url)}`
  };
  $$('[data-share]').forEach((link) => { link.href = links[link.dataset.share]; });
}
async function copyShareLink() {
  try { await navigator.clipboard.writeText(shareUrl()); $("#share-status").textContent = t("shareCopied"); }
  catch (_) { $("#share-status").textContent = t("shareCopyUnavailable"); }
}
$("#copy-share-link").addEventListener("click", copyShareLink);
setShareLinks();

function registerWebMcp() {
  const context = document.modelContext;
  if (!context?.registerTool) return;
  const controller = new AbortController();
  try {
    void Promise.resolve(context.registerTool({
      name: "start_demo_analysis",
      title: "Start demo analysis",
      description: "Runs the local demo with fictional Gmail or Outlook data.",
      inputSchema: { type: "object", properties: { provider: { type: "string", enum: ["gmail", "outlook"] } }, required: ["provider"], additionalProperties: false },
      annotations: { readOnlyHint: false, untrustedContentHint: false },
      execute(input) {
        if (!input || !["gmail", "outlook"].includes(input.provider)) throw new Error("provider deve ser gmail ou outlook");
        startDemo(input.provider);
        return { status: "started", provider: input.provider, data: "fictional" };
      }
    }, { signal: controller.signal })).catch(() => {});
  } catch (_) { /* Browser sem suporte completo a WebMCP. */ }
}
registerWebMcp();
void getOutlookClient().catch(() => { /* The connect button will show a translated error if initialization remains unavailable. */ });
