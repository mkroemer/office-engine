# Theming

The viewer's chrome (bars, toolbars, menus, dialogs, panels, sheet headers, tabs) takes every color, font, radius and
shadow from design tokens: CSS custom properties named `--de-*` on the viewer root (`.de-root`). The document itself
(pages, slides, cells) keeps its own colors in every theme.

## The `theme` prop

```tsx
<DocViewer src={file} theme="auto" />                          // light, dark, or following the OS

<DocViewer
  src={file}
  theme={{
    mode: 'auto',
    tokens: { accent: '#7c3aed', accentHover: '#6d28d9', font: 'Inter, sans-serif', radius: '4px' }, // both modes
    light: { bg: '#f4f4f5', chrome: '#ffffff' },
    dark: { bg: '#18181b', chrome: '#27272a', accent: '#a78bfa' },
  }}
/>
```

`tokens` apply in both modes, `light` and `dark` go over them in their mode, and anything left out keeps the built-in
value of that mode. Names are camelCase (`accentHover` sets `--de-accent-hover`); values are any CSS value, so a token can
point at the host's own variables: `{ accent: 'var(--brand-500)' }`. `createTheme()` types a theme object written apart
from the prop, and `THEME_TOKENS` lists every token.

`mount()` takes the same `theme` option; `<office-viewer theme="dark">` takes a mode.

## Plain CSS

The tokens are ordinary custom properties, so a stylesheet can set them as well:

```css
.my-app .de-root { --de-accent: #7c3aed; --de-font: Inter, sans-serif; }
.my-app .de-root.de-dark { --de-accent: #a78bfa; }
```

Class names starting with `de-` are stable and can be styled further; the tokens are the supported way to change
colors, type and shape.

## Loading the stylesheet yourself

By default the viewer injects its stylesheet once into the document (or the shadow root in `styleTarget`). To load it
yourself, for a Content Security Policy without inline styles, a bundler that extracts CSS, or a cascade layer of your own,
turn injection off and import the file:

```tsx
import '@p10node/office-engine/styles.css'
<DocViewer src={file} injectStyles={false} />
```

```css
/* the viewer's rules below the host's own, whatever their specificity */
@layer office-engine, app;
@import '@p10node/office-engine/styles.css' layer(office-engine);
```

The stylesheet is not put in a layer of its own: rules in a layer lose to every unlayered rule, so a host's element
reset (`button { … }`) would restyle the viewer's controls. Importing it into a layer is the host's choice.

## Your own root

The views (`DocxView`, `XlsxView`, `PptxView`) can be mounted without `DocViewer` (see [API](api.md)); `useThemeRoot()`
gives the class name and style their root needs:

```tsx
const root = useThemeRoot(theme)
return (
  <OfficeUIProvider icons={icons} components={components} locale="de">
    <div className={root.className} style={root.style}><XlsxView client={client} /></div>
  </OfficeUIProvider>
)
```

`OfficeUIProvider` carries what `DocViewer` takes as `icons`, `components`, `locale` and `strings` to views mounted on
their own.

## Your own components

Where tokens are not enough, the toolbars and dialogs render the host's own components: a design system's buttons,
selects and dialogs instead of the built-in ones. Each one left out keeps the built-in one.

```tsx
<DocViewer
  src={file}
  editable
  components={{
    Button: ({ title, icon, label, children, active, expanded, disabled, onClick, onMouseDown }) => (
      <MyIconButton tooltip={title} aria-label={title} pressed={active} aria-expanded={expanded}
        disabled={disabled} onClick={onClick} onMouseDown={onMouseDown}>
        {icon}{label}{children}
      </MyIconButton>
    ),
    Separator: () => <MyDivider vertical />,
  }}
/>
```

| Component | Used for | Props |
|---|---|---|
| `Button` | every toolbar button, the toggles that open color, table and border pickers, the overflow button | `ButtonProps`: `title`, `icon` (resolved through `icons`), `iconName`, `label`, `children`, `active`, `expanded`, `disabled`, `wide`, `onClick`, `onMouseDown`, `className` (the built-in classes) |
| `Select` | font, size list, style, format and menu dropdowns | `SelectProps`: `title`, `options` (`DropOption`: `value`, `label`, `hint`, `icon`, `disabled`, `sep`, `head`, `style`), `value` (undefined for a command menu), `onChange`, `label`, `placeholder`, `width`, `menuWidth`, `disabled`, `alignRight` |
| `Separator` | between toolbar groups | none |
| `Dialog` | prompts (link, rename, …) and confirmations | `DialogProps`: `title`, `kind` (`'alert'` for a question), `children` (the body), `actions` (`label`, `onClick`, `primary`, `submit`, `autoFocus`), `onSubmit` (the body is a form), `onClose` (Escape, backdrop) |
| `Toast` | short messages at the bottom | `text` (the viewer removes it after a few seconds) |

The document must keep the focus while a toolbar control is used, or the command loses its selection:

- `Button`: pass `onMouseDown` to the element. It prevents the default, which keeps the focus.
- `Select`: open on mouse down with the default prevented, and handle its keys without moving the focus into the list.
  The built-in one captures the keys at the window while open.

The overflow menu measures each toolbar item from a wrapper of its own, so a host component needs no ref.

## Tokens

| Group | Tokens |
|---|---|
| Surfaces and text | `bg` (around the document), `chrome` (bars), `pop` (menus, dialogs), `ebar` (formatting toolbar), `text`, `muted`, `ink`, `inkStrong`, `inkSoft`, `border`, `sep`, `tbSep` |
| Accent and states | `accent`, `accentHover`, `accentInk` (text on the accent), `accentSoft`, `on`, `onInk`, `onInkStrong` (pressed buttons), `hover`, `hoverSolid`, `hoverSoft`, `disabled`, `danger`, `dangerBorder`, `ok` |
| Inputs and tabs | `inputBg`, `inputBorder`, `swatchBorder`, `tabHover`, `tabOn`, `tabInk` |
| Overlays | `popShadow`, `shadow`, `shadowSm`, `shadowMd`, `shadowLg`, `backdrop`, `toastBg`, `toastInk`, `tipBg`, `tipInk` |
| Type and shape | `font`, `fontMono`, `fontSize`, `radiusSm`, `radius`, `radiusLg`, `radiusXl` |
| Sheets | `grid`, `head`, `headInk`, `headBorder`, `cornerBorder`, `freeze` |
| Comments | `cmBg`, `cmOn`, `cmLine`, `cmLineOn`, `cmCard`, `cmQuote` |
| Find | `find`, `findOn` (text matches), `findCell`, `findCellOn`, `findCellLine` (cell matches) |
| Files, people, connection | `kindDocx`, `kindXlsx`, `kindPptx`, `kindInk` (badges), `avatarInk` (text on a person's color), `statusIdle`, `statusLive`, `statusBusy`, `statusError` |
| Slideshow | `presentBg`, `presentInk` |
| Paper | `page`, `pageInk`, `sheetInk`, `skeleton` (loading pages), `phBorder`, `phA`, `phB`, `phInk` (placeholders), `phOutline`, `phHint` (empty text boxes) |
| Source editor | `srcBg`, `srcGutter`, `srcSel`, `srcFont`, `srcComment`, `srcKeyword`, `srcString`, `srcOperator`, `srcNumber`, `srcHeading`, `srcAttr`, `srcTag`, `srcMeta` |

The paper tokens are the same in both modes: a document is drawn on white paper with black ink unless it says
otherwise. Built-in values: `src/react/styles.css`, the `.de-root` and `.de-root.de-dark` blocks.
