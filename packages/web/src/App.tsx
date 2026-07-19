import { useCallback, useRef, useState } from "react";
import {
  createRepairPlan,
  WordOrderError,
  repairDocument,
  type RepairPlan
} from "@word-order/core";

type Stage = "idle" | "reading" | "review" | "repairing";

function fileSize(bytes: number): string {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function fixedName(name: string): string {
  return name.replace(/\.docx$/i, "") + ".fixed.docx";
}

function download(bytes: Uint8Array, name: string): void {
  const blob = new Blob([new Uint8Array(bytes)], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fixedName(name);
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1_000);
}

function PrivacyProof(): React.JSX.Element {
  return (
    <div className="privacy-proof" aria-label="Privacy guarantees">
      <span><i aria-hidden="true">01</i> Local processing</span>
      <span><i aria-hidden="true">02</i> No document upload</span>
      <span><i aria-hidden="true">03</i> No generative AI</span>
    </div>
  );
}

function BeforeAfter(): React.JSX.Element {
  return (
    <section className="example-section" id="example">
      <div className="section-intro">
        <span className="eyebrow">A real repair</span>
        <h2>Same words.<br />Back in order.</h2>
        <p>This synthetic agreement was deliberately damaged with mixed fonts, broken numbering, random emphasis, bad alignment, deep indents, uneven margins, and a highlighted placeholder. Word Order repaired the document without rewriting its language.</p>
      </div>
      <div className="comparison" aria-label="Before and after document comparison">
        <figure className="document-shot before-shot">
          <figcaption><strong>Before</strong><span>Broken structure + formatting</span></figcaption>
          <div className="image-frame"><img src="examples/before-agreement.png" alt="A badly formatted services agreement with inconsistent numbering, typography, alignment, spacing, and a highlighted amount placeholder" /></div>
        </figure>
        <div className="comparison-arrow" aria-hidden="true"><span>→</span><small>one local pass</small></div>
        <figure className="document-shot after-shot">
          <figcaption><strong>After</strong><span>Native lists + restrained styles</span></figcaption>
          <div className="image-frame"><img src="examples/after-agreement.png" alt="The same services agreement after repair, with ordered numbering, consistent typography, alignment, margins, and spacing" /></div>
        </figure>
      </div>
      <p className="example-note"><span>✓</span> Rendered from the actual input and output DOCX files. The sample is synthetic and contains no client information.</p>
    </section>
  );
}

function PrivacyArchitecture(): React.JSX.Element {
  return (
    <section className="privacy-architecture" id="privacy">
      <div className="section-intro privacy-intro">
        <span className="eyebrow">Privacy, enforced by the architecture</span>
        <h2>Your contract never becomes a prompt.</h2>
        <p>A DOCX file is really a ZIP package of small XML files. Word Order opens that package in your browser's memory, checks the document structure with fixed rules, and writes a repaired copy back to your device. There is no document server in the middle.</p>
      </div>
      <div className="privacy-boundary" aria-label="Word Order privacy boundary">
        <div className="boundary-device">
          <div className="boundary-label"><span>Your device</span><small>Document stays inside this boundary</small></div>
          <ol className="privacy-flow">
            <li><i>01</i><strong>Read the copy</strong><p>The browser reads your selected file into local memory.</p></li>
            <li><i>02</i><strong>Open the XML</strong><p>Its Word XML, lists, styles, fields, and margins are unpacked locally.</p></li>
            <li><i>03</i><strong>Apply fixed rules</strong><p>Deterministic checks create an inspectable repair plan—no model guesses at your meaning.</p></li>
            <li><i>04</i><strong>Download locally</strong><p>Only targeted XML is repaired; untouched package parts are carried into a new DOCX.</p></li>
          </ol>
        </div>
        <aside className="boundary-network">
          <span className="network-line" aria-hidden="true">×</span>
          <p className="eyebrow">Not sent anywhere</p>
          <ul>
            <li>Document bytes</li>
            <li>Contract text</li>
            <li>Repair plan</li>
            <li>Repaired output</li>
          </ul>
          <p className="boundary-note">The static site has no upload endpoint, account, document database, analytics, or AI provider.</p>
        </aside>
      </div>
      <div className="language-lock">
        <span className="lock-mark" aria-hidden="true">Aa</span>
        <div><span className="eyebrow">Language locked</span><h3>Structure changes. Meaning does not.</h3></div>
        <p>Generative AI can paraphrase, omit, or invent text. Word Order does not generate prose at all. It makes targeted changes to OOXML instructions—the machinery Word uses for numbering, references, styles, spacing, and layout—and tests that the document's visible text is preserved.</p>
      </div>
    </section>
  );
}

function AddinNotice(): React.JSX.Element {
  return (
    <section className="addin-notice" id="word-addin">
      <div><span className="eyebrow">Word add-in</span><h2>Repair inside Word.</h2><p>Bring the same local, deterministic engine into the document you already have open. Scan first, review the proposed fixes, then apply them without sending the document to Word Order.</p></div>
      <div className="addin-actions"><a className="button dark" href="addin/manifest.xml" download>Download current manifest</a><a className="text-link" href="https://github.com/tamirgoldd/word-order/blob/main/docs/addin.md" target="_blank" rel="noreferrer">Installation guide ↗</a></div>
    </section>
  );
}

function ProductStory(): React.JSX.Element {
  return (
    <>
      <section className="trust-strip" aria-label="Product guarantees">
        <article><span>01</span><strong>No document upload</strong><p>Your file is processed in local memory, never sent to Word Order.</p></article>
        <article><span>02</span><strong>No generative AI</strong><p>No model paraphrases, reframes, omits, or invents your language.</p></article>
        <article><span>03</span><strong>Inspectable repairs</strong><p>Fixed rules show every proposed structural change first.</p></article>
        <article><span>04</span><strong>Open source</strong><p>Verify the privacy boundary, run offline, or contribute.</p></article>
      </section>

      <PrivacyArchitecture />

      <BeforeAfter />

      <section className="how" id="how-it-works">
        <div className="section-intro compact-intro"><span className="eyebrow">How it works</span><h2>A repair plan<br />you can inspect.</h2><p>No generative AI, no language model, and no silent rewriting. The same deterministic engine powers the website, CLI, and Word add-in.</p></div>
        <div className="principles">
          <article><span>01</span><h3>Inventory</h3><p>Reads typed numbers, Word lists, indentation, styles, tables, fields, and references.</p></article>
          <article><span>02</span><h3>Plan</h3><p>Shows each proposed repair and stops when the document is unsafe or ambiguous.</p></article>
          <article><span>03</span><h3>Rebuild</h3><p>Writes native multilevel lists, live references, and a coherent Word style system.</p></article>
        </div>
      </section>

      <section className="coverage" id="what-it-fixes">
        <div className="section-intro"><span className="eyebrow">One pass, full cleanup</span><h2>More than numbering.</h2><p>The formatting engine activates only when it sees strong evidence of damage, then removes the drift without flattening the document.</p></div>
        <div className="coverage-grid">
          <article><span>1 → 2 → 3</span><h3>Broken numbering</h3><p>Duplicates, skips, pasted clauses, manual labels, and damaged restarts become native Word lists.</p></article>
          <article><span>Aa · Aa · Aa</span><h3>Font chaos</h3><p>Comic Sans titles and Times, Arial, Calibri, Georgia, Courier, or Verdana swaps collapse into a restrained style system.</p></article>
          <article><span><b>B</b> <i>I</i> <u>U</u></span><h3>Random emphasis</h3><p>Accidental bold, italic, underline, and size changes are removed while meaningful emphasis is preserved.</p></article>
          <article><span>↔ ↤ ↦</span><h3>Layout drift</h3><p>Centered clauses, right-aligned signatures, deep indents, spacing, and uneven margins are normalized.</p></article>
          <article><span>§ 4.02 ↗</span><h3>Dead references</h3><p>Text references are audited and converted into Word REF fields when a safe target exists.</p></article>
          <article><span>[AMOUNT]</span><h3>Placeholder noise</h3><p>Stray highlighting is cleared while bracketed placeholder text remains intact and reviewable.</p></article>
        </div>
      </section>

      <section className="open-source">
        <div><span className="eyebrow">Built in public</span><h2>The engine is the product.</h2></div>
        <div className="source-copy"><p>Word Order is a small TypeScript monorepo with one DOM-free OOXML core shared by the web app, CLI, and Word add-in. The public code shows exactly which Word XML parts are inspected and changed. Unknown XML and untouched package parts are preserved.</p><div className="source-actions"><a className="button dark" href="https://github.com/tamirgoldd/word-order" target="_blank" rel="noreferrer">Explore the GitHub repo ↗</a><a className="text-link" href="https://github.com/tamirgoldd/word-order/blob/main/docs/architecture.md" target="_blank" rel="noreferrer">Read the architecture</a></div></div>
      </section>

      <AddinNotice />
    </>
  );
}

function Summary({ plan }: { plan: RepairPlan }): React.JSX.Element {
  const stats = [
    [plan.summary.numberingConversions, "numbers repaired"],
    [plan.formatting.length, "formatting repairs"],
    [plan.summary.crossReferencesConverted, "references linked"],
    [plan.summary.brokenReferences, "broken references"],
    [plan.summary.anomalies + plan.warnings.length, "items to review"]
  ] as const;
  return (
    <div className="summary-grid">
      {stats.map(([value, label]) => (
        <div className={label.includes("broken") && value ? "stat warning-stat" : "stat"} key={label}>
          <strong>{value}</strong>
          <span>{label}</span>
        </div>
      ))}
    </div>
  );
}

function PlanReview({ plan }: { plan: RepairPlan }): React.JSX.Element {
  const formattingGroups = [...plan.formatting.reduce((groups, change) => {
    const list = groups.get(change.category) ?? [];
    list.push(change);
    groups.set(change.category, list);
    return groups;
  }, new Map<string, typeof plan.formatting>()).entries()];
  return (
    <div className="plan-sections">
      {plan.numbering.length > 0 && (
        <section className="review-card">
          <div className="section-heading">
            <div><span className="eyebrow">Numbering plan</span><h3>Typed labels become native Word lists</h3></div>
            <span className="count">{plan.numbering.length}</span>
          </div>
          <div className="change-list">
            {plan.numbering.slice(0, 12).map((change) => (
              <div className="change-row" key={change.paragraphIndex}>
                <code className="old-number">{change.oldNumber}</code>
                <span className="arrow" aria-hidden="true">→</span>
                <code className="new-number">{change.newNumber}</code>
                <span className="preview">{change.textPreview || "Untitled paragraph"}</span>
                <span className={change.confidence === "high" ? "confidence" : "confidence review"}>{change.confidence}</span>
              </div>
            ))}
          </div>
          {plan.numbering.length > 12 && <p className="more">+ {plan.numbering.length - 12} more changes in the downloaded report</p>}
        </section>
      )}

      {plan.formatting.length > 0 && (
        <section className="review-card">
          <div className="section-heading">
            <div><span className="eyebrow">Formatting plan</span><h3>Inconsistent styling becomes a restrained legal-document system</h3></div>
            <span className="count">{plan.formatting.length}</span>
          </div>
          <div className="formatting-list">
            {formattingGroups.map(([category, changes]) => {
              const example = changes[0];
              return (
                <div className="formatting-row" key={category}>
                  <strong>{category.replace("-", " ")}</strong>
                  <span>{example?.oldValue} <i aria-hidden="true">→</i> {example?.newValue}</span>
                  <small>{changes.length} {changes.length === 1 ? "repair" : "repairs"}</small>
                </div>
              );
            })}
          </div>
        </section>
      )}

      {plan.crossReferences.length > 0 && (
        <section className="review-card">
          <div className="section-heading">
            <div><span className="eyebrow">Cross-reference audit</span><h3>Live links and dead ends</h3></div>
            <span className="count">{plan.crossReferences.length}</span>
          </div>
          <div className="reference-list">
            {plan.crossReferences.map((reference, index) => (
              <div className="reference-row" key={`${reference.paragraphIndex}-${reference.start}-${index}`}>
                <span className={reference.status === "resolved" ? "status-dot resolved" : "status-dot unresolved"} aria-hidden="true" />
                <span><code>{reference.display}</code> {reference.status === "resolved" ? `links to ${reference.targetNumber}` : reference.reason}</span>
                <small>¶ {reference.paragraphIndex + 1}</small>
              </div>
            ))}
          </div>
        </section>
      )}

      {plan.anomalies.length > 0 && (
        <section className="review-card anomaly-card">
          <div className="section-heading">
            <div><span className="eyebrow">Your decision</span><h3>We found numbering that is not safe to guess</h3></div>
          </div>
          {plan.anomalies.map((anomaly) => (
            <div className="anomaly" key={anomaly.paragraphIndex}>
              <span>¶ {anomaly.paragraphIndex + 1}</span>
              <p>{anomaly.message}</p>
              {anomaly.proposedNumber && <code>{anomaly.oldNumber} → {anomaly.proposedNumber}</code>}
            </div>
          ))}
        </section>
      )}


      {plan.warnings.length > 0 && (
        <section className="review-card warning-card">
          <div className="section-heading">
            <div><span className="eyebrow">Editorial review</span><h3>Wording we deliberately did not rewrite</h3></div>
            <span className="count">{plan.warnings.length}</span>
          </div>
          {plan.warnings.map((warning, index) => (
            <div className="anomaly" key={`${warning.paragraphIndex}-${warning.code}-${index}`}>
              <span>¶ {warning.paragraphIndex + 1}</span>
              <p>{warning.message}</p>
            </div>
          ))}
        </section>
      )}
    </div>
  );
}

export function App(): React.JSX.Element {
  const inputRef = useRef<HTMLInputElement>(null);
  const [stage, setStage] = useState<Stage>("idle");
  const [file, setFile] = useState<File | null>(null);
  const [bytes, setBytes] = useState<Uint8Array | null>(null);
  const [plan, setPlan] = useState<RepairPlan | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [confirmed, setConfirmed] = useState(false);
  const [dragging, setDragging] = useState(false);

  const scan = useCallback(async (nextFile: File) => {
    setError(null);
    setConfirmed(false);
    if (!nextFile.name.toLowerCase().endsWith(".docx")) {
      setError("Choose a modern Word .docx file. Legacy .doc files are not supported.");
      return;
    }
    if (nextFile.size > 200 * 1024 * 1024) {
      setError("This document is over 200 MB and is too large to process safely in a browser tab.");
      return;
    }
    setStage("reading");
    try {
      const nextBytes = new Uint8Array(await nextFile.arrayBuffer());
      const nextPlan = await createRepairPlan(nextBytes);
      setFile(nextFile);
      setBytes(nextBytes);
      setPlan(nextPlan);
      setStage("review");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "The document could not be read.");
      setStage("idle");
    }
  }, []);

  const apply = async () => {
    if (!file || !bytes || !plan) return;
    setError(null);
    setStage("repairing");
    try {
      const result = await repairDocument(bytes, plan, { allowAnomalies: confirmed });
      if (result.changed) download(result.bytes, file.name);
      else setError("This document is already structurally clean. No new file was needed.");
    } catch (cause) {
      setError(cause instanceof WordOrderError ? cause.message : "The repair could not be completed safely.");
    } finally {
      setStage("review");
    }
  };

  const reset = () => {
    setStage("idle");
    setFile(null);
    setBytes(null);
    setPlan(null);
    setError(null);
    setConfirmed(false);
    if (inputRef.current) inputRef.current.value = "";
  };

  const busy = stage === "reading" || stage === "repairing";
  const blocked = plan?.status === "blocked";
  const confirmationNeeded = Boolean(plan?.anomalies.length);

  return (
    <div className="page-shell">
      <header className="site-header">
        <a className="brand" href="." aria-label="Word Order home"><span className="brand-mark"><i>W</i><i>O</i></span><span>Word Order</span></a>
        <nav>
          <a href="#example">Before &amp; after</a>
          <a href="#privacy">Privacy</a>
          <a href="#how-it-works">How it works</a>
          <a href="#what-it-fixes">What it fixes</a>
          <a className="nav-source" href="https://github.com/tamirgoldd/word-order" target="_blank" rel="noreferrer">GitHub ↗</a>
        </nav>
      </header>

      <main>
        <section className={plan ? "hero compact" : "hero"}>
          <div className="hero-copy">
            <p className="kicker"><span /> Private DOCX repair for legal teams</p>
            <h1>Put broken Word documents <em>back in order.</em></h1>
            <p className="lede">Repair numbering, cross-references, fonts, spacing, alignment, and margins in one inspectable pass. No cloud upload. No AI rewrite. Your document and its language stay on your device.</p>
            {!plan && <div className="hero-actions"><a className="button" href="#repair">Repair a document</a><a className="text-link" href="#example">See a real before &amp; after ↓</a></div>}
            <PrivacyProof />
          </div>
          {!plan && <div className="repair-card" aria-label="Example repair summary"><div className="repair-card-head"><span>REPAIR PLAN</span><small>word-order.local</small></div><div className="repair-score"><strong>27</strong><span>safe repairs found</span></div><div className="repair-list"><p><i>01</i><span>Numbering</span><s>1, 3, 3, 2</s><b>1, 2, 3, 4</b></p><p><i>02</i><span>Typography</span><s>6 font families</s><b>1 style system</b></p><p><i>03</i><span>Layout</span><s>drifted</s><b>normalized</b></p><p><i>04</i><span>Content</span><s>—</s><b>preserved</b></p></div><div className="repair-card-foot"><span>LOCAL ONLY</span><span>READY TO REVIEW</span></div></div>}
        </section>

        <section className="workspace" id="repair" aria-live="polite">
          {!plan && <div className="workspace-heading"><span className="eyebrow">Try Word Order</span><h2>Repair a copy. Keep the original.</h2><p>No sign-up, upload endpoint, AI provider, or document telemetry.</p></div>}
          {!plan ? (
            <div
              className={dragging ? "dropzone dragging" : "dropzone"}
              onDragEnter={(event) => { event.preventDefault(); setDragging(true); }}
              onDragOver={(event) => event.preventDefault()}
              onDragLeave={(event) => { if (event.currentTarget === event.target) setDragging(false); }}
              onDrop={(event) => {
                event.preventDefault();
                setDragging(false);
                const dropped = event.dataTransfer.files[0];
                if (dropped) void scan(dropped);
              }}
            >
              <input ref={inputRef} type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onChange={(event) => {
                const selected = event.target.files?.[0];
                if (selected) void scan(selected);
              }} />
              <div className="upload-glyph" aria-hidden="true">DOCX<span>↑</span></div>
              <h2>{stage === "reading" ? "Reading the document locally…" : "Drop the broken agreement here"}</h2>
              <p>{stage === "reading" ? "Inventorying paragraphs, list definitions, and references." : "or choose a .docx file — nothing is uploaded"}</p>
              <button className="primary" type="button" disabled={busy} onClick={() => inputRef.current?.click()}>{busy ? "Scanning…" : "Choose Word document"}</button>
              <small>Your original file is never overwritten.</small>
            </div>
          ) : (
            <div className="review-workspace">
              <div className="file-bar">
                <div className="file-icon">W</div>
                <div><strong>{file?.name}</strong><span>{file ? fileSize(file.size) : ""} · scanned on this device</span></div>
                <button className="text-button" type="button" onClick={reset}>Choose another</button>
              </div>

              {blocked && <div className="blocking-message"><strong>Repair paused</strong><p>{plan.blockedReason}</p></div>}
              {!blocked && <Summary plan={plan} />}
              <PlanReview plan={plan} />

              {!blocked && plan.status !== "clean" && (
                <div className="apply-bar">
                  <div>
                    {confirmationNeeded ? (
                      <label className="confirmation"><input type="checkbox" checked={confirmed} onChange={(event) => setConfirmed(event.target.checked)} /><span>I reviewed the flagged changes and approve the proposed numbering.</span></label>
                    ) : <p><strong>Ready to repair.</strong><br />Your original stays untouched.</p>}
                  </div>
                  <button className="primary download-button" type="button" disabled={busy || (confirmationNeeded && !confirmed)} onClick={() => void apply()}>{stage === "repairing" ? "Building document…" : "Download fixed .docx"}</button>
                </div>
              )}
              {plan.status === "clean" && <div className="clean-message"><strong>No automatic repairs remain.</strong><span>{plan.warnings.length ? `${plan.warnings.length} editorial ${plan.warnings.length === 1 ? "warning remains" : "warnings remain"} for human review.` : "The document already uses consistent native numbering, formatting, and references."}</span></div>}
            </div>
          )}
          {error && <div className="error-message" role="alert">{error}</div>}
        </section>

        {!plan && (
          <>
            <ProductStory />
            <section className="faq">
              <span className="eyebrow">Plain answers</span>
              <details><summary>Does my contract leave the computer?</summary><p>No. The static page has no document API or upload endpoint. Your browser reads the DOCX into local memory, repairs it there, and downloads a new copy. The app can keep working after you disconnect from the internet.</p></details>
              <details><summary>Is any contract text sent to an AI model?</summary><p>No. Word Order does not use generative AI or a language model. It never turns your document into a prompt. Fixed, open-source rules inspect Word's XML instructions and preserve the visible wording.</p></details>
              <details><summary>Will the numbers still work after someone edits the file?</summary><p>Yes. Word Order creates Word's native multilevel numbering rather than hardcoded text, then uses bookmarks and REF fields for cross-references.</p></details>
              <details><summary>Does it rewrite or “improve” contract language?</summary><p>No. It changes formatting and Word structure, not prose. Suspiciously long run-on paragraphs are flagged for human review without inserting breaks or changing their wording.</p></details>
              <details><summary>What happens to tracked changes?</summary><p>Word Order stops and asks you to accept or reject them first. It does not attempt a risky repair through revision markup.</p></details>
              <details><summary>Can I run it inside Word?</summary><p>Yes. <a href="addin/manifest.xml" download>Download the Word Order manifest</a> and follow the installation guide. The add-in scans the open document locally, shows its repair plan, and confirms the fixes after they are applied.</p></details>
            </section>
          </>
        )}
      </main>

      <footer><div><a className="brand footer-brand" href="." aria-label="Word Order home"><span className="brand-mark"><i>W</i><i>O</i></span><span>Word Order</span></a><p>Privacy-first, open-source document repair. Not legal advice.<br />Not affiliated with or endorsed by Microsoft.</p></div><div className="footer-links"><a href="#repair">Repair a document</a><a href="addin/manifest.xml" download>Word add-in manifest</a><a href="https://github.com/tamirgoldd/word-order">GitHub ↗</a><span>MIT licensed · 2026</span></div></footer>
    </div>
  );
}
