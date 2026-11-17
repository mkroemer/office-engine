# API

## React

```tsx
import { DocViewer } from '@p10node/office-engine'

<DocViewer src={file} style={{ height: 600 }} />                 // File | Blob | ArrayBuffer | Uint8Array
<DocViewer src="https://cdn.example.com/report.xlsx" />          // URL → Range requests
<DocViewer src={url} headers={{ Authorization: `Bearer ${t}` }} credentials="include" />
<DocViewer src={file} editable onSave={(blob, name) => upload(blob, name)} />  // editing
<DocViewer src={file} allowSave={false} allowDownload={false} />  // no Save/Download, no image download
<DocViewer src={file} fileMenu={false} />                         // no File menu
<DocViewer src={file} theme="dark" />                             // dark chrome; "auto" follows the OS setting
<DocViewer src={file} strings={{ file: { menu: 'Files' } }} />      // override individual UI texts
```

### File menu

| Item                      | DOCX                       | XLSX                           | PPTX                       |
|---------------------------|----------------------------|--------------------------------|----------------------------|
| New                       | blank document             | blank spreadsheet              | blank presentation         |
| Open… (`Ctrl+O`)          | any supported file         | any supported file             | any supported file         |
| Import CSV into new sheet | –                          | `.csv` / `.tsv`                | –                          |
| Save (`Ctrl+S`)           | ✓                          | ✓                              | ✓                          |
| Rename…                   | ✓                          | ✓                              | ✓                          |
| Download as (native)      | `.docx`                    | `.xlsx`                        | `.pptx`                    |
| PDF                       | pages as shown             | current sheet, fit to width    | one slide per page         |
| Web page                  | `.html`, images inlined    | `.html`                        | `.html`, images inlined    |
| Text                      | `.txt`, `.md`              | `.csv`, `.tsv` (current sheet) | `.txt` (incl. notes)       |
| Document details          | pages, words, characters   | sheets, used range             | slides, slide size         |
| Print (`Ctrl+P`)          | ✓                          | ✓                              | ✓                          |
| Host items, Close         | `fileMenuItems`, `onClose` | `fileMenuItems`, `onClose`     | `fileMenuItems`, `onClose` |

Open converts other formats: `.csv` / `.tsv` → workbook, `.txt` / `.md` / `.html` → document (headings, lists, bold / italic, links), `.mmd` / `.puml` / `.dot` / `.gv` → document with the drawing. Converted files start unsaved; Save writes the Office file.
New / Open replace the shown document until the host passes another `src` (unsaved changes are confirmed first; both are hidden during collaboration).
Rename changes the shown name and the name used by Save / Download (`onRename` receives it); it does not reload the document.
Document details also lists the file's core properties (title, author, created, modified…).
PDF goes through the browser print dialog ("Save as PDF"), so text stays selectable. `allowDownload={false}` also hides Download as / Print and blocks `Ctrl+P`.

```tsx
<DocViewer
  src={file}
  fileMenuItems={[{ label: 'Share…', icon: 'open', onClick: share }, { label: 'Version history', onClick: history }]}
  onRename={(name) => api.rename(id, name)}
  onClose={() => router.back()}
/>
```

CSS is auto-injected, no import needed (or import `@p10node/office-engine/styles.css` with `injectStyles={false}`). Container must have a height. Colors, type and shape: [Theming](theming.md).

| Prop               | Type                                           | Default   | Notes                                                                                                                    |
|--------------------|------------------------------------------------|-----------|--------------------------------------------------------------------------------------------------------------------------|
| `src`              | `File \| Blob \| ArrayBuffer \| string \| URL` | –         | Office (also 97-2003); OpenDocument; RTF; Pages / Keynote / Numbers; Visio; CSV / TSV, text, Markdown, HTML, diagrams    |
| `name`             | `string`                                       | auto      | Display name; its extension also tells the format of a `Blob` / `ArrayBuffer`                                            |
| `toolbar`          | `boolean`                                      | `true`    | Top bar (name, status, progress)                                                                                         |
| `toolbarExtra`     | `ReactNode`                                    | –         | Extra buttons on the right of the toolbar                                                                                |
| `zoom`             | `number \| 'fit'`                              | `'fit'`   | Initial zoom (docx/pptx)                                                                                                 |
| `thumbnails`       | `boolean`                                      | `true`    | Thumbnail column for pptx                                                                                                |
| `headers`          | `Record<string,string>`                        | –         | Headers for URL requests                                                                                                 |
| `credentials`      | `RequestCredentials`                           | –         | Cross-origin cookies                                                                                                     |
| `worker`           | `() => Worker \| false`                        | inline    | Custom worker (CSP), `false` = main thread                                                                               |
| `styleTarget`      | `Document \| ShadowRoot`                       | document  | Where CSS is injected                                                                                                    |
| `injectStyles`     | `boolean`                                      | `true`    | Inject the stylesheet; false when the page loads `@p10node/office-engine/styles.css` itself                              |
| `theme`            | `'light' \| 'dark' \| 'auto' \| ThemeOptions`  | `'light'` | Color scheme of the chrome; `'auto'` follows `prefers-color-scheme`; an object adds token overrides ([Theming](theming.md)). Documents keep their own colors |
| `locale`           | `string`                                       | inherited | BCP 47 number/date formatting tag; does not select UI texts. See [UI texts](#ui-texts)                                   |
| `strings`          | `StringsOverride`                              | –         | Texts merged over the enclosing provider or English, key by key; may include a formatting `locale`                       |
| `editable`         | `boolean`                                      | `false`   | Open in edit mode (Pages, Keynote, Numbers and Visio files stay view only)                                               |
| `allowEdit`        | `boolean`                                      | `true`    | Show Edit/View button                                                                                                    |
| `allowSave`        | `boolean`                                      | `true`    | Show Save/Download button, enable Ctrl+S                                                                                 |
| `allowDownload`    | `boolean`                                      | `true`    | Allow downloads to device (save without `onSave`, images)                                                                |
| `fileMenu`         | `boolean`                                      | `true`    | File menu: open, download as PDF / HTML / text, print                                                                    |
| `allowOpen`        | `boolean`                                      | `true`    | File → New / Open (off during collaboration)                                                                             |
| `captureSelectAll` | `boolean`                                      | `true`    | Ctrl/⌘+A from anywhere on the page selects the document; `false`: from the viewer only                                   |
| `fileMenuItems`    | `FileMenuItem[]`                               | –         | Extra File menu entries (`label`, `onClick`, `icon`, `hint`)                                                             |
| `icons`            | `Icons`                                        | built-in  | Your own icons: a map by `IconName` or a resolver `(name, { size }) => node`; see [Icons](#icons)                        |
| `components`       | `Partial<UIComponents>`                        | built-in  | Your own `Button`, `Select`, `Separator`, `Dialog`, `Toast` in the toolbars and dialogs; see [Theming](theming.md#your-own-components) |
| `onRename`         | `(name) => void`                               | –         | File → Rename result (extension included)                                                                                |
| `onClose`          | `() => void`                                   | –         | Adds File → Close                                                                                                        |
| `onSave`           | `(blob, name) => void`                         | download  | Receives saved file (downloads by default): the Office file, or a Markdown / HTML / SVG file as itself                   |
| `onDirtyChange`    | `(dirty) => void`                              | –         | Whether there are unsaved changes                                                                                        |
| `collab`           | `CollabOptions`                                | –         | Real-time collaboration, see [collaboration](collaboration.md)                                                           |
| `comments`         | `CommentOptions \| false`                      | `{}`      | Comments (see below); `false` hides them                                                                                 |
| `onLoad`           | `(client) => void`                             | –         | Once per opened document                                                                                                 |
| `onError`          | `(error) => void`                              | –         |                                                                                                                          |
| `controller`       | `ViewerController`                             | its own   | The viewer's state and commands for the host, see [headless mode](#headless-mode)                                        |
| `headless`         | `boolean`                                      | `false`   | The document alone: no top bar, zoom control, outline panel, sheet tabs or slide list                                    |

### Headless mode

For a host that draws its own controls around the document: `headless` leaves the viewer with the document alone, and a
`controller` (`createViewerController()`) carries its state and the commands that drive it. While editing, the formatting
toolbar keeps its row above the document; dialogs, the find bar and the comments panel are the viewer's own. `toolbar`
set explicitly brings back the top bar (File menu, comments, Save) and leaves the rest headless.

```tsx
import { createViewerController, DocViewer, useViewerState } from '@p10node/office-engine'
import { useMemo } from 'react'

function Viewer({ file }: { file: File }) {
  const ctl = useMemo(() => createViewerController(), [])
  const s = useViewerState(ctl)
  return (
    <>
      <nav>
        {s.outline?.items.map((i) => (
          <button key={i.id} aria-current={i.current} style={{ paddingLeft: i.level * 12 }} onMouseDown={(e) => e.preventDefault()} onClick={() => ctl.goto(i.id)}>
            {i.title}
          </button>
        ))}
      </nav>
      {s.zoom && <button onClick={() => ctl.zoomIn()}>{s.zoom.value === 'fit' ? 'Fit' : `${Math.round(s.zoom.effective * 100)}%`}</button>}
      {s.canEdit && <button onClick={() => ctl.setEditing(!s.editing)}>{s.editing ? 'View' : 'Edit'}</button>}
      {s.dirty && <button onClick={() => ctl.save()}>Save</button>}
      <DocViewer src={file} controller={ctl} headless style={{ height: 600 }} />
    </>
  )
}
```

| `ViewerState` | Meaning |
|---------------|---------|
| `rev` | counts the changes (the state object is replaced on each) |
| `kind` | `'docx'`, `'xlsx'`, `'pptx'`, or `null` before a document is open |
| `loading` | loading or downloading |
| `editing`, `canEdit` | in edit mode; whether editing is possible (allowed, not view only) |
| `dirty`, `saving` | unsaved changes; a save running |
| `canUndo`, `canRedo` | whether undo / redo can run (spreadsheets count the edits made here; presentations offer redo whenever editing) |
| `zoom` | `{ value: 'fit' \| number, effective }`, `null` for spreadsheets |
| `page` | `{ current, count }` from 1: the page (documents; `count` 0 while laid out) or slide in view; `null` for spreadsheets |
| `outline` | `{ kind, items: [{ id, title, level, current }] }`: a document's headings, a workbook's visible sheets, a presentation's slides (titled by their title placeholder, else "Slide n") |
| `status` | the status line ("Page 2 of 7 · 1,204 words", "Slide 3 of 12", "200 rows × 8 cols") |
| `format` | while editing, the formatting at the selection (`TextFormat`, below), retained while using native host inputs; `null` without a valid target, while viewing, and in a presentation unless text is being edited |

Commands: `setEditing(on)`, `save()`, `undo()`, `redo()`, `zoomIn()`, `zoomOut()`, `setZoom('fit' | number)`,
`goto(outlineId)`, `gotoPage(n)`, `present()` (a slideshow from the slide in view), `print()`, `fileAction(id)` (a File menu
action: `pdf`, `html`, `text:md`, `details`, `rename`…), `toggleComments(open?)`, `format(command)` (below). A command the
open document has no use for does nothing. `subscribe(fn)` returns its unsubscribe; `snapshot()` the current state;
`useViewerState(ctl)` is the state in React. Without `react`, `mount()` returns the controller (`mounted.controller`), and
`<office-viewer headless>` has it as `element.controller`.

#### Formatting

One formatting API for documents, cells and slide text, for a toolbar of your own:

```tsx
const f = useViewerState(ctl).format
<button aria-pressed={!!f?.bold} disabled={!f} onMouseDown={(e) => e.preventDefault()}
  onClick={() => ctl.format({ op: 'toggle', mark: 'bold' })}>B</button>
```

`TextFormat`: `bold`, `italic`, `underline`, `strike`, `font`, `size` (points), `color`, `highlight` (a cell's fill in a
spreadsheet), `align` (`'left' | 'center' | 'right' | 'justify'`), `list` (`'bullet' | 'number' | null`; undefined for
cells). `FormatCommand`: `{ op: 'toggle', mark }`, `{ op: 'font', name }`, `{ op: 'size', pt }`, `{ op: 'color', color }`,
`{ op: 'highlight', color }` (`null` removes either), `{ op: 'align', align }`, `{ op: 'list', list }`, `{ op: 'clear' }`.
Sizes must be finite and greater than zero; `NaN`, infinities, zero and negative sizes do nothing, without changing the
document or its history. Existing finite positive sizes keep the editor's usual precision.

`list` commands **toggle**: requesting the same kind again removes it. DOCX removes the list only when **all** selected
paragraphs already have that kind; otherwise it applies that kind to all of them. PPTX decides from the **first** selected
paragraph, including a mixed selection, and applies the result to all selected paragraphs. The snapshot does not summarize
mixed formatting: both report the first selected paragraph's list. `list: null` means no list;
`list: undefined` means lists are unsupported (cells). Other optional `TextFormat` fields may be undefined when unspecified
or unsupported; `format: null` means there is no formatting target.

Cells support marks, font, size, text color, fill, left/center/right alignment and clear; lists and justified alignment do
nothing. DOCX and PPTX support every command above. Tables, pictures, tracked changes and cell borders stay with the built-in
toolbars.

Buttons should prevent mouse down to keep the document selection. Native inputs, textareas and selects may take focus:
the valid editor target remains available, and formatting preserves the control's focus and text selection, including live
`onChange` commands. A true selection loss, leaving edit mode, ending slide text editing, or replacing the source clears
the target. For example, a host size input can use `disabled={!f}` and send a size on change:

```tsx
<input aria-label="Font size" type="number" disabled={!f} defaultValue={f?.size}
  onChange={(e) => ctl.format({ op: 'size', pt: e.currentTarget.valueAsNumber })} />
```

Controller actions obey `allowSave`, `allowDownload` and `allowOpen` even when the File menu is hidden.
`fileMenu={false}` hides that menu and its keyboard shortcuts; it does not disable a host's controller controls.
`allowEdit` controls the Edit/View toggle and CSV import; `editable` and `setEditing()` control edit mode, including when the toggle is hidden.
Prevent the default mouse-down action on host outline buttons (as above) to preserve the editor focus and an open XLSX formula while switching sheets.
New/Open replace the document and remain available for a read-only file when `allowOpen` permits them.
Import requires `allowEdit` and an editable workbook; it remains available with `allowOpen={false}`. Controllers are created with
`createViewerController()`; publishing and attachment methods are internal to the engine.

Use one controller per mounted viewer. Give a modal or fullscreen viewer its own controller, even when it shows the same file.
During a temporary overlap, the latest mounted viewer owns a shared controller. Releasing it clears the controller;
an older viewer that remains mounted does not automatically regain ownership.

### Comments

```tsx
<DocViewer
  src={file}
  editable
  comments={{
    user: { id: me.id, name: me.name },                              // author of the comments written here
    mentions: (q) => api.members(nodeId, q),                         // or a list: [{ id, name }, …]
    onEvent: (e) => e.mentioned.length && api.notify(e.mentioned, e), // notify people mentioned
    open: true,                                                      // panel open on load
  }}
/>
```

| `CommentOptions` | Type                                                      | Default                      |
|------------------|-----------------------------------------------------------|------------------------------|
| `user`           | `{ name, id?, initials? }`                                | the `collab` user, else "Me" |
| `mentions`       | `CommentUser[] \| (query) => CommentUser[] \| Promise<…>` | none (no `@` suggestions)    |
| `canComment`     | `boolean`                                                 | while editing                |
| `onEvent`        | `(e: CommentEvent) => void`                               | –                            |
| `open`           | `boolean`                                                 | `false`                      |
| `focus`          | `string` (a comment id)                                   | –                            |

`focus` shows a thread when the document opens (a link from a notification): the panel opens on it and the document scrolls to its place.

`onEvent` is called for the changes made in this viewer (not for collaborators' — each of them reports their own):

| `CommentEvent` | Meaning                                                                                     |
|----------------|---------------------------------------------------------------------------------------------|
| `type`         | `'add'`, `'reply'`, `'edit'`, `'delete'`, `'resolve'`, `'reopen'`                           |
| `comment`      | the comment (`DocComment`: `id`, `parent?`, `author`, `time`, `text`, `mentions?`, `done?`) |
| `thread`       | the thread's first comment                                                                  |
| `mentioned`    | people mentioned by this change: all of a new comment's, those an edit added                |
| `participants` | the people who wrote the thread's other comments (not the one who made the change)          |
| `quote`        | what the thread points at: the commented text, `Sheet1!B3`, `Slide 2`                       |

A comment's author and mentions carry the `id` the host gave (stored in the file's people / persons / authors list), so a host can tell its users apart after the file was saved and reopened. Ownership (edit / delete) is by `user.id` when there is one, else by name.

Use individual views / build your own UI:

```tsx
import { useDocument, DocxView, XlsxView, PptxView } from '@p10node/office-engine'

const { client, error, progress } = useDocument(src)
if (client?.kind === 'xlsx') return <div className="de-root"><XlsxView client={client} /></div>
```

The views style themselves from design tokens on their root: `useThemeRoot(theme)` gives its class name and style (see [Theming](theming.md)).

### Icons

Every icon the viewer draws, toolbars, menus, zoom, outline, Present and the add-sheet tab included, has a name (`IconName`, all of them in `ICON_NAMES`). `icons` replaces any of them, either as a map or as a resolver that is called for every icon:

```tsx
import { DocViewer, type IconName, type Icons } from '@p10node/office-engine'

// Define outside the host component so the map keeps its identity across renders.
const icons: Icons = { undo: <MyUndo />, redo: <MyRedo /> }

<DocViewer src={file} icons={icons} />
```

A name the map leaves out, or a resolver returning `undefined` / `null`, keeps the built-in icon, so a set can cover just part of the names. Your icon is rendered inside `<span class="de-icon">` whose font-size is the icon's size in pixels: draw it at `1em` (a bare `<svg>` child is sized to `1em` for you) and with `currentColor`, so it follows the theme. The resolver also gets `{ size }` in pixels.

The names exported in `ICON_NAMES` are stable public keys, including names originally used only inside the viewer.
New names may be added; existing keys must not be renamed as an internal refactor.

With [Lucide](https://lucide.dev) (`lucide-react`):

```tsx
import { Bold, Italic, ListTree, Plus, Presentation, Redo2, Replace, Search, Undo2, X, ZoomIn, ZoomOut, type LucideIcon } from 'lucide-react'

const LUCIDE: Partial<Record<IconName, LucideIcon>> = {
  undo: Undo2, redo: Redo2, bold: Bold, italic: Italic, search: Search, replace: Replace,
  zoomIn: ZoomIn, zoomOut: ZoomOut, outline: ListTree, close: X, present: Presentation, sheetAdd: Plus,
}
const lucide = (name: IconName) => {
  const C = LUCIDE[name]
  return C && <C size="1em" strokeWidth={1.8} />
}

<DocViewer src={file} icons={lucide} />
```

With [Phosphor](https://phosphoricons.com) (`@phosphor-icons/react`):

```tsx
import { ArrowArcLeft, ArrowArcRight, ArrowsLeftRight, ListBullets, MagnifyingGlass, MagnifyingGlassMinus, MagnifyingGlassPlus, Plus, PresentationChart, TextB, TextItalic, X, type Icon } from '@phosphor-icons/react'

const PHOSPHOR: Partial<Record<IconName, Icon>> = {
  undo: ArrowArcLeft, redo: ArrowArcRight, bold: TextB, italic: TextItalic, search: MagnifyingGlass, replace: ArrowsLeftRight,
  zoomIn: MagnifyingGlassPlus, zoomOut: MagnifyingGlassMinus, outline: ListBullets, close: X, present: PresentationChart, sheetAdd: Plus,
}
const phosphor = (name: IconName) => {
  const C = PHOSPHOR[name]
  return C && <C size="1em" />
}
```

The standalone views get the same icons from `IconsProvider`, which `DocViewer` uses internally; a nested provider falls back to the enclosing one for names it leaves out:

```tsx
import { IconsProvider, XlsxView } from '@p10node/office-engine'

<IconsProvider icons={lucide}>
  <div className="de-root"><XlsxView client={client} /></div>
</IconsProvider>
```

### UI texts

Viewer UI texts (titles, labels, menu entries, placeholders, dialogs, toasts, status lines such as "Page 1 of 5") come from one typed table, `Strings`, with the exceptions listed below. Only English UI texts with US spelling are included in the default viewer and standalone embed bundle. German is a separate optional import with standard German spelling and `de-DE` number/date formatting (for example `1.234.567,89`):

```tsx
import { DocViewer } from '@p10node/office-engine'
import { de } from '@p10node/office-engine/locales/de'

<DocViewer src={file} strings={de} />
```

The locale module is included only when your app imports it. The German table uses standard German spelling and `de-DE` formatting. `locale="de-DE"` alone changes number/date formatting while keeping English UI texts; it does not load German. With no table or formatting override, dates and grouped counts retain the browser's locale. Page/slide numbers, measurements, file sizes and zoom labels retain ASCII digits and a decimal point, as before. An explicit formatting tag applies to all of these displays.

`strings` merges over the enclosing `StringsProvider` (or English) key by key, so a group can be given in part. Nested providers and `DocViewer` inherit the outer texts and formatting. Undefined values do not replace inherited values: `strings={en}` changes the texts while retaining an enclosing formatting tag; also pass `locale="en-US"` to select US formatting explicitly. A value is a string, or a function where the text takes numbers or names. Count callbacks receive raw numbers for pluralization; callbacks that format grouped counts also receive the active formatting tag as their final optional argument. Hoist override objects outside render, or memoize them when they depend on props, to keep the provider value stable. To customize an imported language, spread that table and any groups you change:

```tsx
import { DocViewer } from '@p10node/office-engine'
import { de } from '@p10node/office-engine/locales/de'

const germanTexts = {
  ...de,
  file: { ...de.file, menu: 'Ablage' },
  docx: {
    ...de.docx,
    status: (page: number, pages: number, words: number, locale?: string) =>
      `${page.toLocaleString(locale)}/${pages.toLocaleString(locale)} · ${words.toLocaleString(locale)} ${words === 1 ? 'Wort' : 'Wörter'}`,
  },
}

<DocViewer src={file} strings={germanTexts} />
```

Views used on their own get their texts from a `StringsProvider` (English without one); `useStrings()` reads them in your own components:

```tsx
import { StringsProvider, useStrings, XlsxView } from '@p10node/office-engine'
import { de } from '@p10node/office-engine/locales/de'

const sheetTexts = { ...de, xlsx: { ...de.xlsx, addSheet: 'Neues Blatt' } }

<StringsProvider strings={sheetTexts}>
  <div className="de-root"><XlsxView client={client} /></div>
</StringsProvider>
```

For a regional variant, override only the texts that differ. The `locale` prop is a BCP 47 formatting tag and takes precedence over `strings.locale`; otherwise the table or enclosing provider supplies the tag, then the browser default. Malformed tags (for example `de_DE`) are ignored: a valid table or inherited tag still applies, otherwise the default formatting above is used. Well-formed but unsupported tags follow `Intl` fallback. For example, a host can use US English as the base and supply British formatting and selected British spellings:

```tsx
import { DocViewer, StringsProvider, en, type StringsOverride } from '@p10node/office-engine'

const usEnglish = { ...en, locale: 'en-US' }
const britishOverrides: StringsOverride = {
  edit: { textColor: 'Text colour' },
  format: { borderColor: 'Border colour' },
}

<StringsProvider strings={usEnglish}>
  <DocViewer src={file} locale="en-GB" strings={britishOverrides} />
</StringsProvider>
```

The same overrides work with `StringsProvider`, `mount()` and the web component's `strings` property. The example changes only those two labels; other text keeps the US English defaults. Override any other regional differences you need.

For a complete language of your own, pass a `Strings` table as `strings`. Starting from `en` (or `de`) keeps untranslated entries as fallbacks; the following example only translates Cancel. Replace the remaining entries to complete the translation:

```ts
import { en, type Strings } from '@p10node/office-engine'

export const fr: Strings = { ...en, locale: 'fr-FR', common: { ...en.common, cancel: 'Annuler' } }
// <DocViewer src={file} strings={fr} />
```

A host can keep its own language tables in its app and pass them through `strings`; no engine registration is needed. A new optional package locale follows the German entry: add its typed table, a `src/locales/<code>.ts` entry and the corresponding build/package export. Locale parity tests check complete keys and nonempty texts.

For vanilla embeds, pass the same optional table to `mount(el, { src, strings: de })` or assign it to the web component:

```html
<script src="office-engine.embed.js"></script>
<office-viewer id="viewer" src="/file.xlsx"></office-viewer>
<script type="module">
  import { de } from './locales/de.js' // optional module next to the distributed embed bundle
  document.getElementById('viewer').strings = de
</script>
```

Document content stays in its original language. Default new sheet names, the author "Me" and new text box content also stay unchanged; the “Page 1 of 5” header/footer preset inserts text in the UI language. Font names, formula names and help, number format previews, host-provided labels and engine/transport errors (including main-thread fetch errors such as `HTTP 404 while fetching …`, collaboration errors such as `Document is still syncing`, and parsing-worker messages) are not translated by the table. Print, PDF and HTML export inherit the active UI table for viewer-generated labels; document content remains unchanged.

## Without a React build

```html
<script src="office-engine.embed.js"></script>   <!-- dist/, bundles React + worker -->

<office-viewer src="/files/report.docx" style="height:600px"></office-viewer>

<div id="v" style="height:600px"></div>
<script>
  const viewer = OfficeEngine.mount(document.getElementById('v'), { src: '/files/deck.pptx' })
  input.onchange = (e) => viewer.update({ src: e.target.files[0] })
  // viewer.destroy()
</script>
```

`<office-viewer>`: attributes `src`, `name`, `editable`, `toolbar="false"`, `thumbnails="false"`, `save="false"`, `download="false"`, `file-menu="false"`, `open="false"`, `capture-select-all="false"`, `theme="dark"` (or `"auto"`), `locale="en-GB"` (number/date formatting), `collab` / `room` / `user` / `base` (real-time collaboration; `user` also writes the comments), `comments="false"`, `headless`; properties `.file = File`, `.icons` (map or resolver, as `icons` above), `.strings = { … }` (see [UI texts](#ui-texts)), `.controller` (see [headless mode](#headless-mode)); events `load`, `error`, `comment` (`detail`: a `CommentEvent`).

## Core API (no React)

```ts
import { DocumentClient } from '@p10node/office-engine/core'

const doc = await DocumentClient.open(fileOrUrl, { onProgress: console.log })
doc.kind                                   // 'docx' | 'xlsx' | 'pptx'
doc.info.from                              // 'pages' / 'keynote' / 'numbers' / 'visio': an Apple or Visio file, converted to docx / pptx / xlsx / pptx (view only)
doc.info.imported                          // 'csv' / 'tsv' / 'text' / 'markdown' / 'html' / 'svg' / 'mermaid' / 'plantuml' / 'graphviz': a plain file, built as xlsx / docx (editable)
doc.info.converted                         // 'doc' / 'xls' / 'ppt' / 'rtf' / 'odt' / 'ods' / 'odp': converted to docx / xlsx / pptx (editable, saved as such)
doc.info.ext                               // 'xlsm' / 'xltx' / 'docm' / 'dotx' / 'pptm' / 'potx' / 'ppsx' …: the file is that variant, and is saved as it (undefined: plain docx / xlsx / pptx)
doc.on('docx:blocks', (blocks) => {})      // docx: streamed blocks
await doc.call('docx.start')
doc.call('xlsx.load', 0)                   // xlsx: stream sheet
doc.on('xlsx:progress', (p) => {})
await doc.call('xlsx.block', 0, 1000, 1100) // formatted cells, rows [1000, 1100)
await doc.call('pptx.slide', 12)           // slide model with layout/master/theme resolved
await doc.call('xlsx.set', 0, [{ r: 4, c: 1, input: '=SUM(A1:A4)' }])
const clip = await doc.call('xlsx.copy', 0, { r0: 0, c0: 0, r1: 9, c1: 2 })   // cells as they are (values, formats, formulas) + `text` for the clipboard
await doc.call('xlsx.sheetOp', 0, { op: 'paste', range: { r0: 20, c0: 0, r1: 29, c1: 2 }, cells: clip.cells, from: { sheet: clip.sheet, r0: 0, c0: 0 } })
await doc.call('xlsx.sheetOp', 0, { op: 'clear', range: { r0: 0, c0: 3, r1: 1048575, c1: 3 } })   // a whole column, every row of it
await doc.call('xlsx.import', 0, csvFile, { name: 'cities' }) // a new sheet after sheet 0 with the rows of a CSV / TSV, streamed in when loaded
await doc.call('xlsx.text', 0, ',') // the sheet as CSV (a Blob; '\t' for TSV), the values as shown
// open(src, { residentCells: 12_000_000 }): cells of a CSV / TSV kept in memory; the other rows are read from the file as they are shown
const comments = await doc.call('comments.list')  // DocComment[]: threads are roots + replies (threadsOf)
doc.on('comment', ({ id, c }) => {})       // a comment set / deleted (c null), here or by a collaborator
await doc.call('comments.set', c.id, { ...c, done: true }) // XLSX / PPTX threads carry their place (`at`)
const blob = await doc.call('save')        // edited file
await doc.call('plain.source')             // text of a Markdown / HTML / diagram file (null for other files)
doc.destroy()
```

## Standalone packages

Diagrams and Markdown without the viewer (no DOM needed, no dependency): [`@p10node/diagram`](../packages/diagram), [`@p10node/mermaid`](../packages/mermaid), [`@p10node/plantuml`](../packages/plantuml), [`@p10node/graphviz`](../packages/graphviz), [`@p10node/markdown`](../packages/markdown).

```ts
import { render, renderPage } from '@p10node/diagram'
const { svg, width, height } = render('sequenceDiagram\n  Ann->>Ben: hi') // language detected
renderPage() // <pre class="mermaid">, <code class="language-dot">… of the page become drawings

import { render as markdown } from '@p10node/markdown'
el.innerHTML = markdown(text) // safe HTML; ```mermaid / ```plantuml / ```dot drawn
```

## CSP

The default worker is an inline blob → requires `worker-src blob:`. If that's not allowed, host `dist/office-engine.worker.js` (exported as `@p10node/office-engine/worker`) on the same origin:

```tsx
<DocViewer src={src} worker={() => new Worker('/static/office-engine.worker.js', { type: 'module' })} />
```

## Links in documents

Links from a document (DOCX text and pictures, PPTX text) are clickable in view mode when they point to `http(s)`, `mailto`, `tel`, `sms`, `ftp(s)`, a relative path or a place in the document (`#bookmark`). Other targets (`javascript:`, `data:`, `vbscript:`, `file:`…) show as plain text, also in the web page and Markdown exports, so a crafted file cannot run script in the host page. The file itself keeps its links unchanged. External links open in a new tab with `rel="noreferrer noopener"`.
