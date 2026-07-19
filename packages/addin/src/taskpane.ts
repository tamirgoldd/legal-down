import {
  createRepairPlan,
  docxToFlatOpc,
  flatOpcToDocx,
  WordOrderError,
  repairDocument,
  type RepairPlan
} from "@word-order/core";
import "./taskpane.css";

const scanButton = document.querySelector<HTMLButtonElement>("#scan")!;
const applyButton = document.querySelector<HTMLButtonElement>("#apply")!;
const status = document.querySelector<HTMLDivElement>("#status")!;
const results = document.querySelector<HTMLDivElement>("#results")!;
const backup = document.querySelector<HTMLElement>("#backup")!;
const backupCheck = document.querySelector<HTMLInputElement>("#backup-check")!;
const intro = document.querySelector<HTMLElement>("#intro")!;
const successScreen = document.querySelector<HTMLElement>("#success-screen")!;
const successHeading = document.querySelector<HTMLHeadingElement>("#success-heading")!;
const successSummary = document.querySelector<HTMLDivElement>("#success-summary")!;
const scanAgainButton = document.querySelector<HTMLButtonElement>("#scan-again")!;

let sourceBytes: Uint8Array | null = null;
let activePlan: RepairPlan | null = null;

function setBusy(busy: boolean, message: string): void {
  scanButton.disabled = busy;
  applyButton.disabled = busy || !backupCheck.checked;
  status.textContent = message;
  status.className = busy ? "notice working" : "notice";
}

function element<K extends keyof HTMLElementTagNameMap>(tag: K, className?: string, text?: string): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text) node.textContent = text;
  return node;
}

function renderPlan(plan: RepairPlan): void {
  results.replaceChildren();
  results.hidden = false;
  const grid = element("div", "summary-grid");
  const stats = [
    [plan.summary.numberingConversions, "numbers"],
    [plan.formatting.length, "formatting"],
    [plan.summary.crossReferencesConverted, "references"],
    [plan.summary.brokenReferences, "broken"],
    [plan.summary.anomalies, "review"]
  ] as const;
  for (const [value, label] of stats) {
    const stat = element("div", "stat");
    stat.append(element("strong", "", String(value)), element("span", "", label));
    grid.append(stat);
  }
  results.append(grid);

  if (plan.formatting.length) {
    const section = element("section", "changes");
    section.append(element("h2", "", "Formatting plan"));
    const groups = new Map<string, number>();
    for (const change of plan.formatting) groups.set(change.category, (groups.get(change.category) ?? 0) + 1);
    for (const [category, count] of groups) {
      const row = element("div", "change");
      row.append(element("strong", "", category.replace("-", " ")), element("span", "", `${count} ${count === 1 ? "repair" : "repairs"}`));
      section.append(row);
    }
    results.append(section);
  }

  if (plan.anomalies.length || plan.warnings.length) {
    const section = element("section", "findings");
    section.append(element("h2", "", "Findings"));
    for (const finding of [...plan.anomalies, ...plan.warnings]) {
      const row = element("div", "finding");
      row.append(element("span", "", `¶ ${finding.paragraphIndex + 1}`), element("p", "", finding.message));
      section.append(row);
    }
    results.append(section);
  }

  if (plan.numbering.length) {
    const section = element("section", "changes");
    section.append(element("h2", "", "Numbering plan"));
    for (const change of plan.numbering.slice(0, 20)) {
      const row = element("div", "change");
      row.append(element("code", "old", change.oldNumber), element("span", "", "→"), element("code", "new", change.newNumber));
      section.append(row);
    }
    results.append(section);
  }
}

function repairLabel(count: number, singular: string, plural: string): string {
  return `${count} ${count === 1 ? singular : plural}`;
}

function renderSuccess(plan: RepairPlan): void {
  successSummary.replaceChildren();
  const applied = element("div", "success-stats");
  const stats = [
    [plan.summary.numberingConversions, "numbering repair", "numbering repairs"],
    [plan.formatting.length, "formatting repair", "formatting repairs"],
    [plan.summary.crossReferencesConverted, "reference linked", "references linked"]
  ] as const;
  for (const [count, singular, plural] of stats) {
    const stat = element("div", "success-stat");
    stat.append(element("strong", "", String(count)), element("span", "", count === 1 ? singular : plural));
    applied.append(stat);
  }
  successSummary.append(applied);

  const details = element("div", "applied-list");
  details.append(element("h2", "", "Applied from your plan"));
  if (plan.summary.numberingConversions > 0) {
    details.append(element("p", "", repairLabel(plan.summary.numberingConversions, "numbering repair applied", "numbering repairs applied")));
  }
  const formattingGroups = new Map<string, number>();
  for (const change of plan.formatting) {
    const label = change.category.replaceAll("-", " ");
    formattingGroups.set(label, (formattingGroups.get(label) ?? 0) + 1);
  }
  for (const [category, count] of formattingGroups) {
    details.append(element("p", "", repairLabel(count, `${category} fix applied`, `${category} fixes applied`)));
  }
  if (plan.summary.crossReferencesConverted > 0) {
    details.append(element("p", "", repairLabel(plan.summary.crossReferencesConverted, "cross-reference linked", "cross-references linked")));
  }
  if (plan.warnings.length > 0 || plan.summary.brokenReferences > 0) {
    const untouched = plan.warnings.length + plan.summary.brokenReferences;
    details.append(element("p", "preserved-note", repairLabel(untouched, "review item left unchanged", "review items left unchanged")));
  }
  successSummary.append(details);

  intro.hidden = true;
  status.hidden = true;
  results.hidden = true;
  backup.hidden = true;
  scanButton.hidden = true;
  applyButton.hidden = true;
  scanAgainButton.hidden = false;
  successScreen.hidden = false;
  window.scrollTo({ top: 0 });
  successHeading.focus();
}

function resetTaskPane(): void {
  sourceBytes = null;
  activePlan = null;
  backupCheck.checked = false;
  successSummary.replaceChildren();
  successScreen.hidden = true;
  intro.hidden = false;
  status.hidden = false;
  status.textContent = "Ready to scan the open document.";
  status.className = "notice";
  results.hidden = true;
  results.replaceChildren();
  backup.hidden = true;
  scanButton.hidden = false;
  scanButton.disabled = false;
  applyButton.hidden = true;
  applyButton.disabled = true;
  scanAgainButton.hidden = true;
  window.scrollTo({ top: 0 });
  scanButton.focus();
}

async function currentDocumentPackage(): Promise<string> {
  return Word.run(async (context) => {
    const ooxml = context.document.body.getOoxml();
    await context.sync();
    return ooxml.value;
  });
}

async function scanDocument(): Promise<void> {
  setBusy(true, "Reading the open document locally…");
  try {
    const flatOpc = await currentDocumentPackage();
    sourceBytes = await flatOpcToDocx(flatOpc);
    activePlan = await createRepairPlan(sourceBytes);
    renderPlan(activePlan);
    if (activePlan.status === "blocked") {
      status.textContent = activePlan.blockedReason ?? "Repair is blocked.";
      status.className = "notice error";
      applyButton.hidden = true;
      backup.hidden = true;
    } else if (activePlan.status === "clean") {
      status.textContent = "No repair is needed.";
      status.className = "notice success";
      applyButton.hidden = true;
      backup.hidden = true;
    } else {
      status.textContent = "Scan complete. Review the plan before applying.";
      status.className = "notice success";
      applyButton.hidden = false;
      backup.hidden = false;
      backupCheck.checked = false;
      applyButton.disabled = true;
    }
  } catch (cause) {
    status.textContent = cause instanceof Error ? cause.message : "Word did not return a readable OOXML package.";
    status.className = "notice error";
  } finally {
    scanButton.disabled = false;
  }
}

async function applyRepair(): Promise<void> {
  if (!sourceBytes || !activePlan || !backupCheck.checked) return;
  setBusy(true, "Applying one document replacement…");
  try {
    const completedPlan = activePlan;
    const result = await repairDocument(sourceBytes, completedPlan, { allowAnomalies: completedPlan.anomalies.length > 0 });
    const repairedFlatOpc = await docxToFlatOpc(result.bytes);
    await Word.run(async (context) => {
      context.document.body.insertOoxml(repairedFlatOpc, Word.InsertLocation.replace);
      await context.sync();
    });
    renderSuccess(completedPlan);
    sourceBytes = null;
    activePlan = null;
  } catch (cause) {
    status.textContent = cause instanceof WordOrderError ? cause.message : cause instanceof Error ? cause.message : "The repair could not be applied safely.";
    status.className = "notice error";
  } finally {
    scanButton.disabled = false;
    applyButton.disabled = !backupCheck.checked;
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    status.textContent = "Open this add-in inside Microsoft Word.";
    status.className = "notice error";
    scanButton.disabled = true;
    return;
  }
  scanButton.addEventListener("click", () => void scanDocument());
  applyButton.addEventListener("click", () => void applyRepair());
  scanAgainButton.addEventListener("click", resetTaskPane);
  backupCheck.addEventListener("change", () => { applyButton.disabled = !backupCheck.checked; });
});
