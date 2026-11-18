# feat(viewer): complete typed UI localization with opt-in German

Hosts can translate the viewer or override individual labels without forking its components. A typed `Strings` table covers the shared chrome and dialogs, DOCX/XLSX/PPTX views, comments, drawing placeholders, collaboration labels and Visio import labels. Partial overrides inherit an enclosing `StringsProvider` or English. Print, PDF and HTML export carry the same table into their separate React roots.

English is the only default language bundled. Standard German is an explicit package import using `de-DE` formatting:

```tsx
import { DocViewer } from '@p10node/office-engine'
import { de } from '@p10node/office-engine/locales/de'

<DocViewer src={file} strings={de} />
```

`locale` is an optional BCP 47 number/date formatting tag; it does not choose or load UI translations. Without an override, formatting follows the browser. The prop takes precedence over `strings.locale`, and malformed tags fall back safely. Count callbacks receive numbers for pluralization and an optional final formatting tag. The API docs show hoisted partial overrides, custom tables and US/British English variants.

Reconstructed on main `fb718881`, preserving its security, save and Visio changes. German UI labels survive export; changing strings no longer resets spreadsheet comment navigation.

Validation: frozen Bun 1.4.4 install; 707 unit tests; typecheck; production/companion builds; package checks; 49 source-browser tests; eight production-browser checks repeated against the packed npm artifact. Browser checks include default/invalid locales, provider inheritance, live switching, singular/numeric callbacks, Visio and separate print/export roots; downloaded HTML was reopened and PDF text inspected. Native OS print-dialog interaction and the entire unrelated browser suite were not tested. Core/transport diagnostic errors and locale-aware numeric-input parsing remain documented exceptions/separate work.

Prepared locally for independent validation before publication; not posted.
