# PR6 local implementation and review response

PR: https://github.com/p10node/office-engine/pull/6
Review: https://github.com/p10node/office-engine/pull/6#pullrequestreview-5953670126

Local branch: `t3/pr6-nov18-local`. Reconstructed only the localization delta from published source `acd8b168bfcf1491946aebba5c39e5e30923ab7f` onto exact main `fb718881fd8fc6fc0a5ee805f602507f306d033e`. Main's security, source-save, flush, frozen-pane and Visio behavior is preserved. No core implementation changed. No preserved checkout was modified; no push, GitHub mutation, merge, comment or issue was made.

Read both complete reviews, all discussion and inline comments through paginated `gh api` reads. There were two reviews, one discussion comment and zero inline comments. Raw evidence is retained under `artifacts/pr6-validation/pr6-*.json`. No applicable repository/ancestor `AGENTS.md` was present. Applied the requested React best-practices skill to context inheritance, hook dependencies, stable object examples and bundle separation.

| Nov18 finding | Fix / evidence or rationale |
| --- | --- |
| 1. Default `en-US` changes browser date/count formatting | English has an undefined formatting tag. Browser defaults remain intact. Unit regression failed before the fix; production and packed-package browsers verify actual Document details dates and grouped counts in `en-GB` and `vi-VN`. |
| 2. Rebase; add Visio labels and supported-format text | Single localization reconstruction starts at exact main. Both tables include `from.visio` and Visio in `file.unsupported`. Running built/packed viewers open `sample.vsdx`, show English/German details and preserve view-only behavior. Main security browser tests and save/reopen tests pass. |
| 3. Separate print/PDF/HTML roots lose translations | The active resolved table is explicitly passed through File actions and mounted with `StringsProvider` in the shared export root. Real PPTX fixture with an EMF picture shows `EMF-Bild`; downloaded HTML is reopened and checked. Both print and PDF preparations are inspected. A PDF generated from the prepared HTML contains `EMF-Bild` under `pdftotext`. Restoring the former unwrapped root in an emitted bundle makes the regression fail. |
| 4. Malformed locale crashes rendering | Tags are canonicalized/validated when resolving the provider value. Malformed tags fall back to undefined/browser formatting. Unit tests plus running ESM/standalone viewers exercise `de_DE`, malformed prototype-like names and malformed `strings.locale`; actual Details remains usable. |
| 5. Redundant locale enum / attribute | `locale` now means a BCP 47 formatting tag across React, mount and the custom element. Removed `LOCALES` and `LOCALE_NAMES`. Text selection remains explicit through `strings`; German stays opt-in. Explicit prop overrides table/inherited formatting. Docs explain precedence and invalid/unsupported fallback. |
| 6. Preformatted count arguments prevent pluralization | DOCX/XLSX/PPTX status, loading status, hidden sheets, used range, limits, print limits and slide-number callbacks receive numeric counts. Grouped-count callbacks also receive the active formatting tag as their final optional argument. Singular German/English forms are fixed. Real one-word DOCX/one-cell XLSX statuses and a numeric host callback are checked in the running production package. |
| 7. DocViewer ignores outer provider | Providers merge over the enclosing resolved table. SSR tests exercise nested providers and DocViewer; the running built/packed viewer inherits German, applies partial overrides and switches live without reopening. |
| N1. Strings identity resets comment hosts / goto counters | Hints read the current controller table lazily. Removed strings from all three host-registration effects. Changing tables invalidates translated placed-thread cache without replacing hosts. Browser regression navigates twice to B3 across fresh string identities; restoring the old effect dependency causes the second navigation to fail. Existing live hint/location tests pass. |
| N2. Provider memo identity and inline docs objects | Documentation hoists override objects and recommends hoisting/memoizing dynamic overrides. Kept identity memoization: function-valued overrides can capture different state even when their source text serializes identically, so a content-string key would be unsafe. This follows the review's offered hoisted-object alternative. |
| N3. Main-thread English errors omitted from exceptions | API docs explicitly list engine/transport errors, main-thread fetch errors and collaboration syncing errors alongside parsing-worker errors. These remain diagnostic exceptions, avoiding changes to core/transport APIs outside localization scope. |
| N4. Helper defaults silently hide missed strings arguments | Removed English defaults from `styleChoices`, `sheetPrintable` and `importFile`'s error callback. All call sites provide context; typecheck/build verify them. |
| N5. Stale PR description | Corrected description prepared locally in `artifacts/pr6-description.md`; not posted. Covers complete views, opt-in standard German and current formatting API. |
| N6. Misaligned docs locale row | Replaced the enum row with the current string/formatting contract, keeping all five table columns aligned with neighboring rows. |

Earlier Nov17 review points remain addressed: all three view-localization commits are included in PR6; docs use existing keys and match implemented scope; German is an optional package entry with standard spelling and `de-DE` formatting; rendered German/partial-override tests cover controls. The former unsupported-locale enum problem is superseded by the documented formatting-tag contract.

## Executed validation

All commands use `npx --yes bun@1.4.4`; dependencies installed with `install --frozen-lockfile`.

- Full unit suite: 707 tests pass; typecheck passes.
- Production build (ESM, standalone embed, worker, types, CSS) and all five companion package builds pass.
- 49 selected source-browser tests pass: localization, embeds, security probes and File menu save/reopen/export flows.
- Eight production-browser tests pass, then repeat against the extracted actual `npm pack` artifact: browser defaults/fallback, standalone embed, provider inheritance, live switching, numeric/plural counts, comment navigation, Visio and export roots.
- Default built viewer/embed import graph contains no German table; emitted German entry loads explicitly with `de-DE`. Actual package import/export smoke and package dry runs pass.
- Regression proof: temporarily restoring the former comment effect dependency and unwrapped export root in the emitted bundle produces two failing browser regressions. Restored the fixed bundle afterward and reran checks. The initial focused source tests also reproduced the forced default, plural and provider-inheritance failures before implementation.
- `git diff --check` passes. No exact HTML snapshots were added.

Logs, package archive/import evidence, browser observations, downloaded/reopened HTML, prepared print/PDF HTML, generated PDF and extracted PDF text are retained locally in `artifacts/pr6-validation/` (gitignored). `browser-evidence.json` contains concrete observed values. The PDF check stubs the native dialog and generates a PDF from the prepared print HTML; manual OS print-dialog interaction was not tested. The entire unrelated browser suite was not run.

## Risks and integration for the parent

Independent validation/review remains required before publication; these checks are the implementation agent's own validation. GitHub still describes the published `acd8b16` head, whose CI passed, with `CHANGES_REQUESTED` remaining from Nov17 and a Nov18 `COMMENTED` review. The new local head has no remote CI/review result.

Use this local PR6 head as the base of the remaining linear 6→7→8→9 stack. Do not replay the three view-localization commits again from #8. Later layers must use numeric status/count callbacks, the optional final formatting tag, required helper context and formatting-only `locale`; they must not reintroduce `LOCALES`, auto-import German, or replace main's save/security/Visio code. Locale-aware numeric input remains a separate dependency (existing #17); this PR changes display formatting, not parsing of numeric input. Engine/transport diagnostic errors remain English as explicitly documented.
