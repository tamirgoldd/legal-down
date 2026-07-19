# Architecture

Word Order uses one deterministic engine across three interfaces. The core
opens the OOXML ZIP, inventories paragraphs, infers hierarchy, audits severe
formatting drift, creates a repair plan, and only then rebuilds the changed
package parts.

## Privacy boundary

The hosted web app is static. It has no application server, upload route,
account, analytics, document storage, telemetry collector, or AI integration.
The selected file is read with the browser File API and passed directly to the
core as an in-memory `Uint8Array`. JSZip opens and rebuilds the OOXML package in
that same browser context; the output is returned through a local Blob download.
The production page adds a Content Security Policy with `connect-src 'none'`,
which blocks browser connection requests even if one were added accidentally.

The Word add-in is hosted as static HTML and JavaScript and obtains OOXML from
the open document through Office.js. It passes that package to the same local
core. Word Order has no endpoint that accepts the document, extracted text,
repair plan, or result. The CLI likewise operates on local files.

## Why this is not generative AI

The engine does not call a language model and has no prose-generation step.
It parses explicit OOXML patterns, computes a deterministic plan, and performs
targeted transformations. A given document and engine version produce the same
plan. Ambiguous structures become warnings or blocking errors rather than model
guesses.

The engine does not use browser DOM APIs. Reads are tolerant, while writes are
targeted string transformations so unknown XML and untouched ZIP members are
not normalized accidentally. A plan records every numbering and reference
decision and has a digest tied to the source document.

Formatting normalization is deliberately thresholded. Ordinary documents are
left alone; normalization activates when the engine sees strong evidence such
as three or more direct font families, four or more direct sizes, novelty fonts,
highlighting, deep indents, or materially uneven margins. The repair maps title,
subtitle, heading, body, and signature paragraphs to `LD*` Word styles and
removes only conflicting direct formatting. Meaningful all-caps emphasis and
bracketed placeholders remain bold, but placeholder highlighting is removed.
Long run-on paragraphs are warnings because changing their wording would violate
the text-preservation invariant.

The `LD*` style IDs and `_LDRef_*` bookmark names are legacy internal OOXML
identifiers retained for compatibility with documents repaired before the Word
Order rename. Their visible Word style names use the current product name.

## Failure boundaries

- Tracked changes block planning and repair.
- Missing or malformed mandatory OOXML parts produce a typed error.
- Ambiguous sequences produce confirmation-required anomalies.
- A clean document returns its original bytes without rezipping.
- Cross-references that cannot be resolved are reported and left as text.
- Formatting cleanup never changes sentence text or silently invents paragraph
  breaks.

The corpus generator turns declarative agreements into `.docx` fixtures and
applies known mutation operators. Tests assert plan snapshots, preserved plain
text, untouched package parts, idempotence, and valid package relationships.
