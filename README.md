# Word Order

**Put broken Word documents back in order — locally.**

[![CI](https://github.com/tamirgoldd/word-order/actions/workflows/ci.yml/badge.svg)](https://github.com/tamirgoldd/word-order/actions/workflows/ci.yml)
[![GitHub Pages](https://github.com/tamirgoldd/word-order/actions/workflows/pages.yml/badge.svg)](https://github.com/tamirgoldd/word-order/actions/workflows/pages.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-123d31.svg)](LICENSE)

[Try the web app](https://tamirgoldd.github.io/word-order/) · [See the real before and after](#before--after) · [Read the architecture](docs/architecture.md)

Word Order is a privacy-first, open-source DOCX repair engine for legal documents. It fixes broken numbering, dead cross-references, mixed fonts and sizes, accidental emphasis, alignment, indents, spacing, highlighting, and uneven margins without uploading the file or sending contract text to an AI provider.

The web app runs entirely in the browser and works offline. There is no document backend, account, analytics, content telemetry, or generative AI. The CLI supports batch checks, and the Word add-in is a thin interface over the same deterministic repair engine.

> Early alpha: work on a copy and review every proposed change. Word Order refuses to repair documents with tracked changes or unresolved structural ambiguity.

## Before & after

These images were rendered from the actual synthetic input and repaired DOCX files used to test the formatting engine.

| Before | After |
| --- | --- |
| Broken numbering, six font families, random sizing and emphasis, erratic alignment and indents, uneven margins, and a highlighted placeholder. | Ordered native lists, consistent reusable styles, normalized layout and margins, and preserved wording. |
| ![Badly formatted synthetic services agreement before Word Order repair](packages/web/public/examples/before-agreement.png) | ![The same synthetic services agreement after Word Order repair](packages/web/public/examples/after-agreement.png) |

## What it repairs

- Broken or manually typed clause numbering, including duplicates, skips, and restarts
- Textual cross-references, converted to live Word `REF` fields when a safe target exists
- Random font-family and size swaps, including novelty fonts
- Accidental bold, italic, underline, alignment, indentation, and spacing drift
- Uneven section margins and stray placeholder highlighting
- Long run-on paragraphs, flagged for human review without rewriting the words

The output uses native Word multilevel lists, bookmarks, fields, and named styles. It keeps working when someone edits the repaired document in Word.

## Privacy-first by architecture

Word Order does not ask you to trust a privacy promise wrapped around a cloud service. The web app is a static program that downloads to your browser. It has no upload endpoint, document database, account system, analytics, or connection to an AI provider. Its production Content Security Policy also blocks browser connection requests.

A `.docx` file is a ZIP package containing XML instructions for text, numbering, styles, fields, margins, and relationships. The repair happens like this:

1. Your browser reads the selected DOCX into local memory. The file input is not submitted anywhere.
2. JSZip opens the package in that tab. The engine inventories Word's XML structure and visible text.
3. Fixed TypeScript rules build an inspectable repair plan. No language model predicts what the contract “should” say.
4. The engine changes only targeted OOXML parts, preserves unknown XML and untouched ZIP members, and downloads a new DOCX to your device.

| The network receives | The network does not receive |
| --- | --- |
| Static Word Order HTML, CSS, and JavaScript when you open the site | Your document bytes, extracted contract text, repair plan, or repaired output |

### Why there are no AI hallucinations

Generative AI can paraphrase, omit, invent, or reframe language. Word Order never generates prose. It repairs Word's structural instructions—such as list definitions, bookmarks, `REF` fields, named styles, spacing, and margins—while treating visible document text as an invariant. Corpus tests compare text before and after repair and fail if wording changes outside recognized typed-number tokens. Ambiguous numbering and long run-on clauses are surfaced for human review instead of being guessed or rewritten.

## Try it locally

Requires Node.js 22+ and pnpm 11+.

```bash
pnpm install
pnpm check
pnpm dev
```

Vite prints the local web-app URL. The CLI separates scanning from repair:

```bash
pnpm --filter @word-order/cli start -- scan agreement.docx
pnpm --filter @word-order/cli start -- fix agreement.docx \
  -o agreement.fixed.docx --report plan.json
```

## Safety model

- Document bytes and extracted text stay on the device. There is no backend, upload API, account, analytics, document logging, AI provider, or content telemetry.
- `scan` produces an inspectable plan. `fix` refuses tracked changes and unresolved anomalies.
- Wording is never silently edited. Editorial concerns are warnings, not automatic rewrites.
- Untouched ZIP members and out-of-scope OOXML parts are preserved.
- The input file is never overwritten.

## Word add-in

Download the [current hosted manifest](https://tamirgoldd.github.io/word-order/addin/manifest.xml) to use the same local repair engine inside Word.

The task pane scans the open document, shows the proposed repair plan, asks for backup confirmation, and then displays a completion screen summarizing the fixes applied. See the [add-in installation guide](docs/addin.md).

## Project map

| Package | Purpose |
| --- | --- |
| `@word-order/core` | DOM-free OOXML inventory, inference, planning, and rebuild |
| `@word-order/cli` | Node.js `scan` and `fix` commands |
| `@word-order/web` | Offline-capable drag-and-drop PWA |
| `@word-order/addin` | Office.js Word task pane |

See [Architecture](docs/architecture.md), [Word add-in sideloading](docs/addin.md), [Contributing](CONTRIBUTING.md), and [Security](SECURITY.md).

## License and name

MIT licensed. Word Order is open-source software, not a law firm and not legal advice. It is not affiliated with or endorsed by Microsoft.
