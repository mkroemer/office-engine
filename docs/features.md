# Features

## Loading and rendering

- **Lazy ZIP**: reads only the central directory, inflates each part on demand (native `DecompressionStream`, 0 dependencies)
- **HTTP Range**: open a URL → fetch only the bytes needed (256 KB block cache); server without Range support → automatic full-download fallback
- **Streaming XML parser** runs in a **Web Worker** → UI never blocks
- **Progressive render**:
  - DOCX: blocks (paragraph/table) stream in → offscreen height measurement in time slices → pagination → only visible pages render
  - XLSX: rows stream in, 2D virtualized grid (only cells in viewport render), freeze panes, merge, number format, sheet tabs
  - PPTX: slides parsed on demand when scrolled into view, virtualized thumbnails

## Format support

| Format | Supported                                                                                                                                                                                                      | Not yet                                                                    |
|--------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|----------------------------------------------------------------------------|
| DOCX   | inherited styles, run formatting, lists/numbering, tables (merge, border, shading, table style), images, hyperlinks, header/footer, PAGE field, text boxes, page/section breaks, text columns, tracked changes | footnote body, charts, exact float wrapping, paragraphs split across pages |
| XLSX   | shared strings, number/date formats, font/fill/border/align, merge, freeze, hidden columns/rows, wrap, formula calculation (135 functions, recalculated on edit), multiple sheets, images (over the cells)     | charts (placeholder), shapes, conditional formatting                       |
| PPTX   | master/layout/theme, placeholder inheritance, ~50 preset shapes, custGeom, images (crop), tables + table style, groups, bullet/autonum, autofit                                                                | charts (placeholder), SmartArt, animations, video playback                 |

Password-protected (encrypted) files are not supported.

### Word, Excel and PowerPoint 97-2003, RTF and OpenDocument

These files open **editable**: the worker converts them in memory to the Office Open XML file the views show, and **Save** writes that file (`report.doc` is saved as `report.docx`, `budget.ods` as `budget.xlsx`; the original is never overwritten). The bar shows a `DOC` / `XLS` / `PPT` / `RTF` / `ODT` / `ODS` / `ODP` badge, *Document details* says "Word 97-2003 document, opened as Word document". Nothing leaves the browser and no library is added: every format is read by the library's own code in the worker.

| File                    | Opens as                                                                                   |
|-------------------------|--------------------------------------------------------------------------------------------|
| `.doc`, `.dot`          | document (Word 97-2003)                                                                    |
| `.xls`, `.xlt`          | workbook (Excel 97-2003, also Excel 5.0 / 95 and workbook streams outside a compound file) |
| `.ppt`, `.pot`, `.pps`  | presentation (PowerPoint 97-2003)                                                          |
| `.rtf`                  | document                                                                                   |
| `.odt`, `.ott`, `.fodt` | document (OpenDocument text, template, flat XML)                                           |
| `.ods`, `.ots`, `.fods` | workbook (OpenDocument spreadsheet, template, flat XML)                                    |
| `.odp`, `.otp`, `.fodp` | presentation (OpenDocument presentation, template, flat XML)                               |

The format comes from the file's bytes (compound file, `{\rtf`, OpenDocument `mimetype`, flat XML), so a `.doc` that is really RTF, or a file without a name, opens too.

**Word 97-2003 (`.doc`)**

- **Converted**: paragraph and character styles (built-in ones get Word's names: Heading 1, Title…), fonts, sizes, bold / italic / underline kinds / strikethrough, caps, colors, highlight, shading, borders, super/subscript, spacing, alignment, indents, line spacing, tabs, keep with next, page breaks; bullet and numbered lists (9 levels); tables (merged cells, borders, shading, padding, row heights, header rows, nested tables); sections (page size, orientation, margins, columns, section breaks), headers and footers (first page, even pages), footnotes and endnotes, comments, bookmarks, tracked insertions / deletions; fields (links, `PAGE` / `NUMPAGES`, the rest with their result), inline and floating pictures, text boxes and shapes, groups, WordArt text
- **Not yet**: Word 6.0 / 95 and older files (an error says so), table styles, drop caps, formatting revisions; EMF / WMF pictures are kept in the file (Word shows them) and shown as a placeholder

**Excel 97-2003 (`.xls`)**

- **Converted**: every sheet (hidden ones stay hidden), values of every type, the shared string table, fonts, fills, borders, alignment, wrap, indent, rotation, number formats (built-in and custom), Excel 2007+ colors and themes, merged cells, column widths and row heights, hidden rows / columns, outline levels, frozen panes, gridlines, zoom, the active sheet; formulas (every token, shared and array formulas, names, 3-D references) with their last values; defined names, print areas, page setup, headers and footers, hyperlinks, cell comments, tab colors; Excel 5.0 / 95 files
- **Not yet**: pictures, charts, conditional formats, data validation; chart sheets open empty; formulas pointing at other workbooks keep their values only

**PowerPoint 97-2003 (`.ppt`)**

- **Converted**: slide size, masters and title masters (backgrounds, shapes, placeholders, text styles, color schemes), slides with their own backgrounds and color schemes, hidden slides, footers / slide numbers / dates; about 190 shape types, freeforms, fills (solid, gradient, picture, texture, pattern), lines (dash, arrowheads), shadows, rotation and flips, groups, tables (as grouped cells); text with bullets (also numbered), fonts, sizes, colors, alignment, spacing, links, slide number and date fields; pictures; speaker notes and comments (saved in the file)
- **Not yet**: PowerPoint 95 and older files (an error says so), animations and transitions, sounds and movies, charts and embedded objects (their preview picture is shown); EMF / WMF pictures are kept in the file and shown as a placeholder

**RTF**

- **Converted**: every code page (also Shift-JIS, GBK, Big5, Korean), `\u` Unicode text, Symbol / Wingdings characters; paragraph and character styles (built-in ones get Word's names), fonts, sizes, bold / italic / underline kinds / strikethrough, caps, colors, highlight, shading, borders, super/subscript, spacing, alignment, indents, line spacing, tabs, frames; bullet and numbered lists (also the older WordPad / Word 95 kind); tables (merged cells, borders, shading, padding, header rows, row heights, nested tables); sections (page size, orientation, margins, columns, breaks, page numbering, first page), headers and footers (first page, left / right pages), footnotes and endnotes, bookmarks, links, `PAGE` / `NUMPAGES` fields (the rest with their result), tracked insertions / deletions; PNG / JPEG / GIF / BMP pictures (size, scale, crop), text boxes and shapes in groups
- **Not yet**: comments, equations, drawing objects and lines, drop caps, line numbering, page borders; EMF / WMF pictures and OLE objects (their preview picture is shown when it is a bitmap)

**OpenDocument text (`.odt`)**

- **Converted**: named styles (LibreOffice's built-in ones get Word's names), automatic styles as direct formatting, every character and paragraph property the DOCX view shows (fonts for each script, sizes, colors, underline kinds, strike, highlight, super/subscript, caps, letter spacing, alignment, margins, line height, keep together / with next, breaks, borders, backgrounds, tab stops), numbered headings, bullet / numbered / picture lists (9 levels, continued numbering), tables (spans, header rows, borders, backgrounds, padding, nested tables), pictures (inline and anchored) and text boxes, links, bookmarks, footnotes and endnotes, comments, page styles as sections (size, orientation, margins, columns, page numbering), headers and footers (left / first page), page number and count fields, tables of contents (as their text)
- **Not yet**: tracked changes (shown accepted), lines and connectors, objects with only a StarView preview, drop caps, line numbering, picture cropping

**OpenDocument spreadsheet (`.ods`)**

- **Converted**: every sheet (hidden, tab color, right to left), values (numbers, percentages, currencies, dates, times, booleans, errors), formulas translated to Excel's syntax with their last values, number formats (also conditional), cell styles (font, fill, borders, alignment, wrap, rotation, indent), merged cells, column widths and row heights, hidden rows / columns, outline groups, frozen panes, gridlines, zoom, the active sheet, named ranges, print ranges, page setup with headers and footers, hyperlinks, comments. The empty rows and columns LibreOffice writes to the end of a sheet (a million rows) are not expanded
- **Not yet**: pictures and charts, conditional formatting, data validation, rich text inside a cell

**OpenDocument presentation (`.odp`)**

- **Converted**: slide size, master pages as slide masters (theme colors and fonts, backgrounds, drawings, placeholders, title / outline text styles), slides with their own backgrounds, hidden slides, date / footer / slide number; text with lists, bullets and numbering, fonts, sizes, colors, alignment, spacing, links, fields; preset and custom shapes (`draw:enhanced-geometry` drawn as custom geometry), lines and connectors with arrowheads, fills (solid, gradient, bitmap, hatch), shadows, rotation, skew and flips; pictures (cropped), groups, tables (cell styles, merged cells), speaker notes and comments (saved in the file), text that shrinks to fit
- **Not yet**: charts and embedded objects without a picture (a labeled box), 3D scenes, Fontwork effects, animations and transitions

**Not supported**: password-protected files, Word 6.0 / 95, Excel 2.x-4.0 and PowerPoint 95 files, OpenOffice.org 1.x files (`.sxw`, `.sxc`, `.sxi`): each gives an error saying so.

### Apple Pages, Keynote and Numbers (view only)

Files from Pages 5, Keynote 6 and Numbers 3 (2013) and later open read-only: the worker converts a `.pages` document to DOCX, a `.key` presentation to PPTX and a `.numbers` spreadsheet to XLSX, in memory, and the matching view shows it. Find (Pages and Numbers; a Keynote presentation, like any presentation here, has no find), print, word count, slideshow, sheet tabs and *Download as* work as usual; *Download as* **Word** / **PowerPoint** / **Excel** saves the converted file. The bar shows a `PAGES` / `KEYNOTE` / `NUMBERS` badge and no Edit button, even with `editable`. Single files and zipped package folders (`Report.pages/…` inside a ZIP) both open.

**Pages**

- **Converted**: paragraph and character styles (named styles become Word styles), fonts, sizes, bold/italic/underline/strikethrough, colors, highlights, super/subscript, caps, alignment, indents, spacing, page breaks, bullet and numbered lists (levels), links, inline images, text boxes and shapes with a fill, tables (merged cells, header rows, cell fills, banded rows, lines, text/number/date cells, rich text cells), header and footer (left, center and right parts, page number), footnotes (listed after the text), page size and margins
- **Not yet**: charts, equations, comments, tracked changes, exact positions of floating objects (placed in the text flow: those on the first page before the body, the others after it; on page layout documents, page by page), image cropping, several sections with different headers

**Keynote**

- **Converted**: slide size and order, skipped slides (hidden), master slide objects (drawn under the slide's), slide backgrounds (color, gradient, image), title / body / slide number placeholders, text boxes (text styles, bullet and numbered lists, links), shapes (rectangles, rounded rectangles, polygons, arrows, stars, custom paths and curves, fill and line), images cropped by their mask, movie poster frames, groups, rotation and flips, tables (merged cells, header rows, cell fills, banded rows, lines)
- **Not yet**: charts (a placeholder box), builds and transitions, speaker notes, shadows and reflections, text that shrinks to fit

**Numbers**

- **Converted**: every table becomes a worksheet (named after its sheet, plus the table's name when a sheet holds several), hidden sheets stay hidden; values keep their type: numbers, currencies, percentages, scientific and dates with the matching Excel number format (Numbers date patterns translated), booleans; durations, fractions, numbers in other bases, checkboxes, star ratings and pop-up menu choices as the text Numbers shows; rich text cells as their text; cell look (font, size, bold/italic/underline/strikethrough, color, alignment, vertical alignment, wrap, fill, banded rows, lines), column widths, row heights (grown to fit wrapped text), merged cells, hidden rows / columns, frozen header rows / columns
- **Not yet**: formulas (cells hold the values Numbers computed), charts, images and text boxes placed on sheets, custom formats (shown with the General format), accounting-style currency, the document locale's decimal / thousands separators, conditional highlighting, comments

Cell values in Pages and Keynote tables use the same formats.

**Not supported**: password-protected files and the iWork '09 and older XML format: each gives an error saying so.

### Microsoft Visio (view only)

Visio drawings open read-only: the worker converts a `.vsdx` file to PPTX in memory, a slide per page, and the presentation view shows it (thumbnails of the pages, zoom, slideshow, print, *Download as* **PowerPoint**). The bar shows a `VISIO` badge and no Edit button, even with `editable`. Macro-enabled drawings (`.vsdm`) and templates (`.vstx`, `.vstm`) open the same way; a stencil (`.vssx`, `.vssm`) shows each of its masters as a slide. The package is told by its content, whatever its name.

Each shape is drawn where it is on its page: its geometry becomes a custom path already turned, flipped and placed (so a drawing of nested, rotated groups needs no frames of its own), its text a text box over it, its picture a picture. Text is laid out by the browser in the shape's text block, and can be selected and copied.

| Part of the drawing | Converted                                                                                                                                            |
|---------------------|------------------------------------------------------------------------------------------------------------------------------------------------------|
| ShapeSheet          | cells inherited from the master's shape, then from the line, fill and text styles and their parents; rows and sections a shape overrides or deletes  |
| Geometry            | `MoveTo`, `LineTo`, `ArcTo`, `EllipticalArcTo`, `Ellipse`, `RelCubBezTo`, `RelQuadBezTo`, `PolylineTo`, `NURBSTo`, splines, the relative row types   |
| Geometry sections   | several per shape, filled together (holes), `NoFill` / `NoLine` / `NoShow`, rounded corners (`Rounding`)                                             |
| Placement           | pin, size, angle, flips, through any depth of groups; 1-D shapes (lines, connectors) along their line                                                |
| Masters             | instances with their own size, colors and text; members of an instance, listed or not                                                                |
| Fill                | solid, transparency, gradient stops and angle, the gradient patterns of older drawings, hatches (as a mixed color)                                   |
| Line                | color, weight, transparency, dash patterns, line ends (begin / end, three sizes)                                                                     |
| Themes              | themed fills, lines, fonts and text colors: scheme and variation colors put through the theme's styles, connector styles, readable text over a fill  |
| Text                | the text block (place, size, angle, margins, vertical alignment), character rows (font, size, color, bold, italic, underline, strikethrough, case,   |
|                     | super / subscript), paragraph rows (alignment, indents, spacing, bullets), fields (their shown value), line breaks, a background behind the text     |
| Pictures            | PNG / JPEG / GIF / BMP / SVG, cropped to the part the shape shows, turned and flipped                                                                |
| Pages               | page size and drawing scale, background pages under the pages that use them, layers that are hidden, a group's own drawing behind or in front        |

- **Page sizes**: a presentation has one slide size, so it is the largest page's; a smaller page is drawn as a sheet at the center of its slide. A page over PowerPoint's 56 in is scaled down, text and line weights with it
- **Not yet**: EMF / WMF pictures (the icons of many network stencils: a labeled box), shadows, line caps and compound lines, exact line end shapes (drawn as triangles, with PowerPoint's kinds in the saved file), tabs, text along a vertical direction, hyperlinks, shape data and comments, connectors routed again (they are drawn as saved), formulas (cells hold the values Visio computed)
- **Not supported**: Visio 2003-2010 binary drawings (`.vsd`), which give an error saying to save the drawing as `.vsdx`
- **Limits**: a drawing of more than 400,000 shapes (counting each member of every master instance), or of more than 256 MB of slide content, is refused

### CSV, text, Markdown, HTML and SVG

A `src` (file, Blob, buffer or URL) that is not an Office package opens too, converted in the worker and **editable**:

| File                          | Opens as                                                                                                                                                                                                                                                                  |
|-------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `.csv`, `.tsv` / `.tab`       | workbook, one sheet named after the file; the rows stream in as the file is read (a million rows open like a large workbook; past 12 million cells the rows are read from the file by pages as they are shown); numbers and booleans typed, codes like `007` kept as text |
| `.txt`, `.text`, `.log`       | document, one paragraph per line                                                                                                                                                                                                                                          |
| `.md`, `.markdown`            | document: headings, lists, bold / italic / strike, links, code, quotes, tables (as tab-separated lines)                                                                                                                                                                   |
| `.html`, `.htm`, `.xhtml`     | document: headings, paragraphs, lists, inline formatting and links; scripts and styles are dropped                                                                                                                                                                        |
| `.svg`                        | the picture alone (no page), scaled down to fit the view, at its own size (`width` / `height` in any unit, else the `viewBox`) otherwise; the file is kept as it is, given the SVG namespace when it lacks one; *Download as* Word gives a document holding the picture   |
| `.svgz`, `.csv.gz`, `.md.gz`… | the file inside: a gzip-compressed plain file is unpacked whole in the worker (native `DecompressionStream`), the name inside telling the format                                                                                                                          |

- **Format**: the extension (of `name`, the File or the URL), else the MIME type (Blob type, `Content-Type`), else the content (SVG, HTML, delimited rows, or text). A gzip file (`.svgz`, `.csv.gz`) is unpacked first and the name inside decides. A file named `.docx` / `.xlsx` / `.pptx` / … is read as a package, so an HTML error page served for a missing file still fails (with a message saying it is a web page); the old names `.doc` / `.xls` / `.xlt` are told by their bytes: a web page, MHTML ("Save as Web Page"), CSV, TSV or text saved under such a name opens as what it is. The bytes also decide when the name is wrong the other way: a real `.doc` / `.xls` named `.docx` / `.xlsx` is converted. A file that is none of these says what it is ("This is a PDF file…", an image, an empty file, an Excel 2003 XML spreadsheet), a password-protected one says to save it without a password, a damaged one (a cut-short download, a part that does not unpack or fails its checksum, a main part that is not XML) says it is damaged
- **CSV**: comma, semicolon (then `12,5` is a number) or tab, detected from the first line; quoted fields, newlines inside quotes
- **Streamed**: the text goes from the file straight into the sheet, in chunks, in the worker (never a string of the whole file, never a workbook built first): the first rows show at once and the rest streams in like a large workbook's. Column widths come from the first 5 000 rows. Saving writes every row from the sheet
- **Paged**: past 12 million cells (about 300 MB in Chrome) the rows stay in the file: the loader notes where each page of 1 024 rows starts and goes on counting, and a page is read when its rows are shown, edited, searched, exported or saved, the least recently shown pages making room (pages with edits stay). So there is no row limit. In such a sheet rows and columns cannot be inserted or deleted, and an edit or a formula may not need more rows than fit in memory (`=SUM(A:A)` over 100 million rows is refused with a message). `residentCells` in the open parameters sets the budget
- **Export**: *Download as* CSV / TSV starts with a byte order mark (UTF-8, so Excel on Windows shows accents). A text that starts with `=`, `+`, `-` or `@` (or a tab / return) is written with a `'` in front, so that a spreadsheet does not run it as a formula (CSV injection); numbers (`-5`), dates and booleans are not texts and are written as they are. A number in a cell with no number format (General) is written with its digits (up to 15 significant, no exponent while the plain form is short), not rounded as the grid shows it; a number with a format is written as shown
- **Excel's limit**: a sheet of more than 1,048,576 rows cannot be saved as a workbook: the bar says so, saving fails with a message, and *Download as* CSV / TSV gives the rows (written in the worker, page by page)
- **Encoding**: the `encoding` open parameter when given (a `TextDecoder` label, e.g. `windows-1258`), else the byte order mark (UTF-8, UTF-16), the declared `charset`, UTF-8, else Windows-1258 when the bytes read as Vietnamese in it (tone marks after vowels: Excel's CSV on Vietnamese Windows; the text is put in NFC), else Windows-1252 (Excel's CSV on other Windows)
- **Columns**: a sheet holds 16,384 columns (up to XFD); the columns of a wider CSV after it are not loaded and the viewer says so (a message; `warning` event from the worker)
- The bar shows a `CSV` / `TSV` / `TXT` / `MD` / `HTML` / `SVG` badge
- **Saving**: a Markdown, HTML or SVG file is saved **as itself** (`.md` / `.html` / `.svg`, also what `onSave` receives). Markdown and HTML are built again from the document when it was edited there: headings, lists (nested), bold / italic / underline / strike, links, code, colors, tables, pictures (embedded as data URLs; a diagram block becomes its drawing); the text reads with its tracked changes accepted. An unedited file is saved as it is. CSV / TSV and text are saved as a workbook / document (`.xlsx` / `.docx`); *Download as* Word / CSV / text / Markdown gives the other formats

### Diagrams: Mermaid, PlantUML, Graphviz

Diagram sources open as a document holding the drawing (an SVG picture, editable around it, saved as `.docx`). The drawing is made in the worker by a built-in renderer: no server, no network request, no third-party library.

| File                                         | Draws                                                                                                                |
|----------------------------------------------|----------------------------------------------------------------------------------------------------------------------|
| `.mmd`, `.mermaid`                           | Mermaid: flowchart / graph, sequence (with `box`), class, state, ER, pie, Gantt, mind map                            |
| `.puml`, `.plantuml`, `.pu`, `.iuml`, `.wsd` | PlantUML: sequence, class / object, use case, component / deployment, state, activity, mind map, WBS, `@startdot`    |
| `.dot`, `.gv`                                | Graphviz DOT: graph / digraph, clusters, default attributes, edge chains, `rank=same`, record and HTML labels (text) |

- **In other files**: ` ```mermaid `, ` ```plantuml ` / ` ```puml `, ` ```dot ` / ` ```graphviz ` code blocks in Markdown, and `<pre class="mermaid">`, `<div class="mermaid">`, `<pre><code class="language-…">` blocks in HTML are drawn in place; other code blocks stay code
- **Format**: the extension, else the MIME type (`text/vnd.graphviz`…), else the content: `@startuml`, `digraph {`, a Mermaid header (`graph TD`, `sequenceDiagram`…), also for a Blob or URL served as `text/plain`. A `.dot` file that is a Word 97 template (binary) is not read as a graph
- **Syntax**: PlantUML names and aliases as in the official examples (`actor Guest as g`, `actor "Guest User" as g`, `actor g as "Guest User"`), one-line deployment elements (`agent`, `artifact`, `card`, `cloud`, `file`, `folder`, `frame`, `stack`, `storage`…), state descriptions (`S : text`), activity swimlanes and notes, preprocessing (`!define` with arguments, `!$var`, `!if` / `!ifdef` / `!else` / `!endif`); Mermaid markdown strings (read as their text), quoted edge labels, edge ids (`A e1@--> B`); DOT `strict`
- **Errors**: a Mermaid flowchart / sequence line, or a DOT text, that cannot be read is reported with its line number instead of a partial drawing (comments and supported-but-ignored directives are not errors)
- **Layout**: layered, like Mermaid's and Graphviz's `dot` (ranks, fewer crossings, straight long edges, clusters kept apart); `TB` / `LR` / `BT` / `RL`; node shapes, edge labels, arrowheads (UML, crow's foot), colors from `classDef` / `style` / DOT attributes
- **Page**: the drawing is scaled down to fit the page; a wide one gets a landscape page. Several diagrams in one file (PlantUML blocks, DOT graphs) are drawn one after another
- **Safe**: the text is escaped, colors are kept only when they are plain colors, and the picture has no scripts, links or external references
- **Code side by side**: the **Code** button (Markdown, HTML, SVG and diagram files) shows the source next to the document, with line numbers and syntax colors; editing it rebuilds the document as you type (the page keeps its scroll position). Meanwhile the document is a preview (the source is what is edited). For Markdown and HTML the pane shows the source built from the document when that was edited last, so the two never disagree; a diagram's document can be edited around the drawing, but that is not in its source (the pane says so), and Save writes the Word file. The source pane copies or downloads the source. In a narrow viewer the source goes above the document; the divider can be dragged (or moved with the arrow keys)
- **Not drawn**: other Mermaid kinds (git graph, journey, timeline, quadrant, C4, Sankey…), other PlantUML kinds (Gantt, JSON / YAML, salt, timing, network), the Graphviz engines other than `dot` (`neato`, `circo`… are laid out as `dot`), HTML tables in DOT labels (their text is shown). A diagram that cannot be drawn shows why, followed by its source

## Editing

### DOCX

- **Typing**: type/delete/Enter, paste (keeps formatting when pasting from Word/Google Docs/web), IME input (e.g. Vietnamese)
- **Text formatting**: B/I/U/S, superscript/subscript, font, size, color, highlight, paint format
- **Styles**: Normal, Title, Subtitle, Heading 1–6 always available (created in the file when missing; Ctrl/⌘+Alt+0…6), update a style to match the selection (right-click → Update style to match: font, size, color, B/I/U/S, alignment and spacing go into the style, every paragraph of it follows; written to that file's own styles only), document outline panel
- **Paragraphs**: alignment, line spacing, bullet/number list, indent, page break
- **Images**: insert (file picker/paste/drag & drop), click to select with 8 resize handles, original size, fit width, replace, alt text, download
- **Tables**: insert, add/delete rows & columns, delete table, Tab between cells; select cells by dragging across them, merge / unmerge (across and down), cell color, cell borders (all / outer / inner / none / one side, color, line style), vertical alignment, distribute columns, drag a column edge to resize
- **Page setup**: paper size (A4, Letter, Legal, A3, A5, B5), portrait / landscape, margins (cm, presets), 1–3 text columns with spacing; applies to the whole document
- **Headers & footers**: header, footer and first-page parts (double-click the page's header / footer area, or the toolbar), alignment, bold, size, page number / page count fields with presets ("Page 1 of 5", "Trang 1/5"), different first page
- **Tracked changes**: insertions and deletions from Word shown in their author's color and kept on save; next / previous change, accept / reject (the change at the caret, the selection, or all); *Track changes* records your own edits as suggestions (typing, deleting, pasting, replacing)
- **Other**: select all (Ctrl/⌘+A), link (dialog), special characters, right-click context menu (clipboard, link, table, image, insert, styles), find & replace, word count, undo/redo

### XLSX

- **Cells**: edit, line break inside a cell or the formula bar (`Alt+Enter` or `Shift+Enter`: the cell wraps and its row grows to show every line), range selection, select all (Ctrl/⌘+A), copy/cut/paste (within a workbook: exact values, formats and formulas, a cut moving the cells and what reads them; from other programs: tab separated text, quoted cells with line breaks included; a whole column or sheet is cleared and copied whole, the rows not shown yet too), fill down/right and the fill handle (drag the corner of the selection down or to the right: one row / column is copied, several are repeated), one copied value pasted into a selection fills it, a paste of HTML only (a table) pastes its cells, a paste that would drop data over merged cells is refused with a message, merging cells that hold data says what was dropped, sort (the rows of the table around the selection, its header left on top; in Excel's order: numbers, text, FALSE, TRUE, errors, cells without anything last in either direction; comments go with their rows; a range holding merged cells is refused with a message), find & replace (match case, entire cell, in formulas). A date or a percentage typed into a cell without a format shows as one; a cell formatted as Text keeps what is typed. `+A1+B1`, `-A1` and `@SUM(A1:A3)` are formulas (what reads a cell or calls a function: `-5` is a number, `- item` is text); `#N/A` and the other error values typed are errors; `1,234,567`, `$50`, `€1,200.50` and `50 ₫` are numbers shown as typed; a number of more than 15 digits (a card number) stays text. A formula typed into a cell without a format shows as what it computes (`=TODAY()`, `=DATE(…)`, `=A1+30` as dates, `=A1` as A1 is shown)
- **Keyboard**: arrows, `Tab` / `Enter` (after Tabs, Enter returns to the column the row was started in), `Home` / `End` / `Ctrl+Home` / `Ctrl+End`, `Ctrl+Arrow` (to the edge of the block of filled cells), `Shift+Space` / `Ctrl+Space` (select the row / column), `Alt+=` (AutoSum), `Ctrl+;` (today; with Shift the time), `Ctrl+Enter` (the entry into every selected cell), `F4` (cycles `$` on the reference at the caret), `F2`, `Ctrl+S` (also in a cell being edited). Text that comes from an IME (Vietnamese Telex included), a virtual keyboard or an emoji picker starts the edit of the selected cell: the grid keeps a hidden text field for the keyboard, which becomes the cell editor. Its on-screen keyboard is not opened until a cell is edited
- **Sizes**: `Wrap text` and a text rotation grow the rows to show the text (tilted text goes out over the cells to its right, as in Excel); double click on the edge of a column / row fits it; a column left at the default width widens when a number format (date, currency…) would show `####`. A number without a format is shown in what fits its column, measured in the font the browser has (fewer decimals, then the scientific form, `####` last). Pattern fills show as the mix of their colors, diagonal borders are drawn, shrink to fit scales the text; rich text (parts of a cell in fonts of their own) is shown, and typing over it leaves plain text. Not done: zoom; rows of a file that say no height are not fitted to wrapped or turned text when it is opened
- **Formulas**: auto-calculated (135 functions; shared and array formulas; `INDIRECT` / `OFFSET` / `NOW` / `RAND` at every change, `TODAY` / `NOW` / `RAND` also when the file is opened; a workbook that says to compute on load is computed when opened; a formula using a function the engine does not have, or a defined name, keeps the value the file holds), function name & argument hints while typing, click/drag cells (incl. other sheets, arrow keys) to insert references, colored references, auto-close parentheses, functions menu (SUM/AVERAGE/COUNT/MAX/MIN)
- **Formatting**: paint format, currency/percent/decimal buttons + 20 number formats and a format code typed by hand (*Custom format…*: `#,##0 "đ"`, `ddd dd/mm`), font, size, B/I/U/S, text/fill color, borders (8 kinds, color, line style), merge all/horizontally/vertically, horizontal & vertical alignment, wrap, text rotation, clear formatting. A column or a row selected by its header (or `Ctrl+Space` / `Shift+Space`), or the whole sheet, takes the format as its own (`<col style>`, `<row s>`): its empty cells show it and what is typed there later has it, without a cell per place; a plain range of more cells than can be formatted one by one is refused with a message
- **Rows & columns**: insert/delete/hide/unhide (defined names, print areas, filters, hyperlinks, data validation, conditional formats, page breaks, tables, the anchors of pictures and charts and the cells charts draw follow), drag to resize column width/row height, freeze, gridlines
- **Sheets**: add/rename/delete/reorder, duplicate (cells, formats, widths and merges; formulas naming the sheet itself name the copy; drawings, comments and tables of the original are not copied). A name that is empty or taken is refused with a message; one that is cut to 31 characters or has `\ / ? * [ ] :` replaced says so
- **Protection**: a protected sheet refuses the edits of its locked cells and the format / structure edits its protection does not allow, a workbook with a locked structure the changes of its sheet list, each with a message; locked cells are not cut either (the password is not asked for: protection is not removed here)
- **Other**: context menu, undo/redo

### PPTX

- **Text**: font, size, B/I/U/S, color, highlight, superscript/subscript, alignment, vertical alignment, line spacing, bullet/numbered list, indent level (also Tab / Shift+Tab in a list item), clear formatting; the text of table cells is edited like a shape's (click a cell; Tab / Shift+Tab move between cells; Delete in a cell edits the text, with the table selected as an object it deletes the table)
- **Shapes**: select/drag/resize/delete (arrow keys to nudge), fill, border color/weight/dash, arrange (front/forward/backward/back), rotate 90°/flip, align on slide, duplicate (⌘D), undo/redo
- **Insert**: 30 shapes, text box, image
- **Slides**: background, add/duplicate/delete/move/hide (all undoable, also a deleted slide with its notes and comments), slideshow from current slide or beginning
- **Speaker notes**: read from the notes slide (also of converted `.ppt` / `.odp`) and included in *Download as plain text*; not shown or edited in the view
- **Other**: context menus on shapes/slides/thumbnails

### Toolbar

Single row, buttons that don't fit collapse into `⋮`; custom-drawn dropdowns (font/style preview, number formats with examples, ↑↓/Enter/Esc/type-ahead keys, never steals focus from the document); font size box `− 11 +` accepts direct input; cell borders picked from an icon grid; color pickers have a palette plus a **Custom** row (`+` opens the system color picker, recent custom colors are remembered across sessions). Every icon can be replaced with the host's own set (Lucide, Phosphor or any other) through the `icons` prop, see [API](api.md#icons).

### Dark mode

`theme="dark"`, or `"auto"` to follow the OS / browser setting (`prefers-color-scheme`), darkens the chrome: bars, toolbar, menus, dialogs, sheet headers and tabs, thumbnails, popovers, scrollbars. Pages, slides and cells keep the colors the document gives them, like paper in a dark PDF viewer. Everything is driven by design tokens (CSS variables on `.de-root`), which a host sets through a theme object or its own CSS: see [Theming](theming.md).

## Comments

- **Where**: DOCX — the selected text (or the word at the caret); XLSX — a cell (one thread per cell, as in Excel: commenting on a cell that has one answers it); PPTX — next to the selected shape, or where the slide is right-clicked
- **Start one**: `Ctrl/⌘+Alt+M`, right-click → Comment, or Comments → New comment
- **Threads**: replies, resolve / reopen (answering a resolved thread reopens it), edit and delete your own comments, deleting a thread's first comment deletes the thread
- **@mentions**: type `@` to pick someone from the host's list (`comments.mentions`); the host hears who was mentioned (`comments.onEvent`) to notify them
- **Panel**: threads in document order with what they point at (the commented text, `Sheet1!B3`, `Slide 2`), resolved ones on demand; picking a thread shows its place, clicking commented text / a marked cell / a pin on a slide shows its thread
- **In the document**: commented text is highlighted (stronger for the current thread), cells get a mark in their top right corner, slides a pin with the author's initials; resolved threads are not shown
- **DOCX ranges follow the text**: typing inside a commented range extends it, typing right after it does not; splitting a paragraph keeps both parts covered; undo takes a new comment away (redo brings it back); a thread whose text is deleted leaves the file
- **XLSX**: threads move with inserted / deleted rows and columns and go with a deleted row, column or sheet (undo brings them back)
- **PPTX**: a slide's threads stay with it (not copied by Duplicate slide, deleted with the slide)
- **Shared rooms**: comments are shared operations like edits — every editor sees them live, newcomers get them from the file or the history
- **Read-only** while viewing (default: comments can be written while editing, `comments.canComment` overrides)

Stored the Office way, so Word, Excel and PowerPoint show them:

| Format | Parts                                                                                                                 | Notes                                                                                                                                                  |
|--------|-----------------------------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------|
| DOCX   | `comments.xml`, `commentsExtended.xml` (threads, resolved), `commentsIds.xml`, `commentsExtensible.xml`, `people.xml` | ranges as `commentRangeStart` / `commentRangeEnd` + reference; untouched comments keep their XML (rich text); comments in headers / footnotes are kept |
| XLSX   | `threadedComments/*.xml`, `persons/person.xml`, plus the legacy `comments*.xml` + VML Excel keeps next to them        | notes (legacy comments) are read as one-comment threads; a note nobody answered or changed stays a note                                                |
| PPTX   | modern comments `comments/modernComment_*.xml`, `authors.xml`                                                         | legacy comments are read as one-comment threads and written back as modern comments                                                                    |

Comment text is plain (lines kept); formatting inside comments read from a file is kept while the comment is not edited. Ids are durable: a comment keeps its id after saving and reopening (DOCX durable ids, XLSX / PPTX GUIDs).

## File menu

- **New**: blank document, spreadsheet or presentation (16:9, Office theme, title / content / blank layouts)
- **Open** (`Ctrl/⌘+O`): `.docx` `.xlsx` `.pptx` open as they are; `.doc` `.xls` `.ppt` `.rtf` and OpenDocument `.odt` `.ods` `.odp` are converted (editable, saved as `.docx` / `.xlsx` / `.pptx`); `.csv` / `.tsv` become a workbook (numbers and booleans typed, codes like `007` kept as text); `.txt` / `.md` / `.html` become a document (headings, lists, bold / italic / strike, links, code); `.svg` / `.mmd` / `.puml` / `.dot` become a document with the picture
- **Download as**: the native format, PDF, web page (`.html`, images embedded), plain text / Markdown (DOCX), CSV / TSV of the current sheet (XLSX), slide text with notes (PPTX)
- **Import CSV into new sheet** (XLSX): adds a sheet after the current one with the file's rows, streamed in the same way (in a shared session the rows go to everyone as cell edits)
- **Save** (`Ctrl/⌘+S`), **Rename** (shown name and the name used for saving), **Document details** (file, core properties, page / word / slide counts)
- **Host entries** (`fileMenuItems`: Share, Version history…) and **Close** (`onClose`)
- **PDF / Print** (`Ctrl/⌘+P`): rendered off-screen and sent to the browser print dialog (choose "Save as PDF"): DOCX pages as paginated on screen, PPTX one slide per page, XLSX current sheet fitted to the paper width

## Saving

- DOCX: edited paragraphs are regenerated from the original XML (keeps `pPr`, `rPr`, images, bookmarks, comments); untouched paragraphs keep original bytes
- XLSX: only `<row>`s containing edited cells are rewritten (streamed, never loads the whole sheet into a DOM); new strings written as inline strings; the values of the formulas an edit recomputed are written (in every sheet), and `fullCalcOnLoad` set so Excel recalculates all of them; functions newer than the format get their `_xlfn.` prefix; a workbook with macros, or a template, is saved as what it is (`OpenResult.ext`: `xlsm`, `xltx`…; the macros are kept)
- PPTX: only `a:p` of edited shapes are regenerated
- Unchanged ZIP entries are copied as-is (no recompression)

| Operation                              | File                                  | Time           |
|----------------------------------------|---------------------------------------|----------------|
| Type 1 char (browser)                  | DOCX 2073 pages                       | ~7 ms          |
| Save after edit                        | DOCX 20k paragraphs                   | ~0.5 s         |
| Save after editing 2 cells             | XLSX 300k rows (25 MB)                | ~3 s           |
| Open (rows streamed in)                | CSV 1M rows × 10 cols (99 MB)         | ~2.4 s         |
| Open (rows paged, ~1 GB of memory)     | CSV 25M rows × 10 cols (2.7 GB)       | ~14 s          |
| Open (rows paged, ~1 GB of memory)     | CSV 100M rows × 10 cols (10.6 GB)     | ~50 s          |
| Open (rows paged, <1 GB of memory)     | CSV 500M rows × 10 cols (55 GB)       | ~4.5 min       |
| Save (every row written)               | CSV 1M rows → XLSX (56 MB)            | ~13 s          |

## Limitations

- Formulas: no defined names, structured refs (`Table1[Col]`), 3-D refs
- DOCX: text boxes and nested tables are read-only; paragraphs with fields spanning multiple paragraphs (TOC) are read-only; headers / footers are edited as text lines (pictures, tables and shapes in them are replaced when edited, with a warning); a header taller than the top margin overlaps the text; tracked changes don't cover formatting changes, paragraph marks or pictures; page setup applies to the whole document (not per section)
- XLSX: no chart/pivot/conditional format/data validation/filter editing; a table whose cells are all deleted stays in the file; cells cut and pasted do not take their merges, comments or conditional formats along; a sheet read by pages (a CSV past 12 million cells) takes no row / column insert or delete, and its formulas only the rows that fit in memory
- PPTX: can't rotate shapes with the mouse yet, no table insert / row or column edit (Tab in the last cell adds no row), no animations/transitions, speaker notes are not editable; undo of slide list changes (add/delete/move) is not available in a shared (collaborative) session
- Comments: plain text only; DOCX comments can't cover text in more than one table cell or across a table; comments in headers, footers and footnotes are kept but not shown; PPTX comments anchored to a shape are shown at the shape's position when the file is opened and keep that anchor until moved
- Files from someone else: nothing a file says is used as an address to load when it is opened (a color that is not a color, a link that is not a web / mail / phone link and a name the format does not know are ignored; pictures linked from outside the file are not fetched). A package is refused, with a message, when it names a part twice (also by case), holds more than 200,000 files, has a part that unpacks to more than 1 GB (16 GB for sheet data and the document body, which are streamed) or nests XML deeper than 1,024 elements
