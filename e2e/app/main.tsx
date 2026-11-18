/**
 * Fixture for the browser tests: DocViewer configured from the query string, callbacks recorded on window.
 *   ?src=samples/sample.docx&editable=1&allowSave=0&onSave=1&items=1&close=1&collab=ws://…&room=r
 *   comments: &me=Ann&meId=u1 (author), &people=1 (mentions: Ann, Bob Stone, Cid), &comments=0 (hidden), &cmOpen=1
 *   &icons=1: every icon from a host resolver, as `<svg data-host-icon={name}>`
 *   texts: &locale=de-DE opts into the explicitly imported German table, &fileLabel=… overrides the File label
 *   &components=1: the host's own Button, Separator, Dialog and Toast (`data-host-*` markers)
 */
import { useState } from 'react'
import { createRoot } from 'react-dom/client'
import { DocViewer, type DocViewerProps, type FileMenuItem, type UIComponents } from '../../src'
import { de } from '../../src/locales/de'

type Ev = { type: string; [k: string]: unknown }
declare global {
  interface Window {
    __events: Ev[]
    __setProps: (p: Partial<DocViewerProps>) => void
    __setLanguage: (language: 'en' | 'de-DE') => void
    __german: typeof de
  }
}
window.__events = []
const log = (e: Ev) => window.__events.push(e)

const q = new URLSearchParams(location.search)
const flag = (k: string) => (q.has(k) ? q.get(k) !== '0' : undefined)
const texts = q.get('locale') === 'de-DE' || q.get('locale') === 'de' ? de : undefined

async function b64(blob: Blob) {
  const bytes = new Uint8Array(await blob.arrayBuffer())
  let s = ''
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  return btoa(s)
}

const items: FileMenuItem[] = [
  { label: 'Share…', icon: 'open', onClick: () => log({ type: 'item', label: 'Share' }) },
  { label: 'Version history', hint: 'beta', onClick: () => log({ type: 'item', label: 'Version history' }) },
]

// plain host components: no engine classes, so nothing of the engine's CSS styles them
const components: Partial<UIComponents> = {
  Button: (p) => (
    <button type="button" data-host-button="" data-active={p.active || undefined} title={p.title} aria-label={p.title} aria-expanded={p.expanded} disabled={p.disabled} onMouseDown={p.onMouseDown} onClick={p.onClick} style={p.iconName === 'more' ? { width: 140 } : undefined}>
      {p.icon}
      {p.label}
      {p.children}
    </button>
  ),
  Separator: () => <i data-host-sep="" />,
  Dialog: (p) => {
    const actions = p.actions.map((a, i) => (
      <button key={i} type={a.submit ? 'submit' : 'button'} data-primary={a.primary || undefined} autoFocus={a.autoFocus} onClick={a.submit && p.onSubmit ? undefined : a.onClick}>
        {a.label}
      </button>
    ))
    const body = (
      <>
        <h2>{p.title}</h2>
        {p.children}
        {actions}
      </>
    )
    return (
      <div data-host-dialog="" role={p.kind === 'alert' ? 'alertdialog' : 'dialog'} aria-label={p.title} onKeyDown={(e) => e.key === 'Escape' && p.onClose()}>
        {p.onSubmit ? (
          <form
            onSubmit={(e) => {
              e.preventDefault()
              p.onSubmit!()
            }}
          >
            {body}
          </form>
        ) : (
          body
        )}
      </div>
    )
  },
  Toast: ({ text }) => <output data-host-toast="">{text}</output>,
}

const initial: DocViewerProps = {
  src: q.get('src'),
  name: q.get('name') ?? undefined,
  style: { height: '100vh' },
  toolbar: flag('toolbar'),
  editable: flag('editable'),
  allowEdit: flag('allowEdit'),
  allowSave: flag('allowSave'),
  allowDownload: flag('allowDownload'),
  fileMenu: flag('fileMenu'),
  allowOpen: flag('allowOpen'),
  thumbnails: flag('thumbnails'),
  captureSelectAll: flag('captureSelectAll'),
  theme: (q.get('theme') as DocViewerProps['theme']) ?? undefined,
  injectStyles: flag('injectStyles'),
  locale: (texts ? undefined : q.get('locale') as DocViewerProps['locale']) ?? undefined,
  strings: q.has('fileLabel') ? { ...texts, file: { ...(texts?.file ?? {}), menu: q.get('fileLabel')! } } : texts,
  onSave: q.has('onSave') ? async (blob, name) => void log({ type: 'save', name, b64: await b64(blob) }) : undefined,
  onDirtyChange: (dirty) => log({ type: 'dirty', dirty }),
  onRename: (name) => log({ type: 'rename', name }),
  onClose: q.has('close') ? () => log({ type: 'close' }) : undefined,
  fileMenuItems: q.has('items') ? items : undefined,
  onLoad: (client) => log({ type: 'load', kind: client.kind, name: client.name, imported: client.info?.imported }),
  onError: (error) => log({ type: 'error', message: error.message }),
  comments:
    q.get('comments') === '0'
      ? false
      : {
          user: q.get('me') ? { name: q.get('me')!, id: q.get('meId') ?? undefined } : undefined,
          mentions: q.has('people') ? [{ name: 'Ann', id: 'u1' }, { name: 'Bob Stone', id: 'u2' }, { name: 'Cid', id: 'u3' }] : undefined,
          open: flag('cmOpen'),
          onEvent: (e) => log({ type: 'comment', action: e.type, id: e.comment.id, text: e.comment.text, thread: e.thread.id, quote: e.quote, mentioned: e.mentioned.map((m) => m.id ?? m.name), participants: e.participants.map((m) => m.id ?? m.name) }),
        },
  components: q.has('components') ? components : undefined,
  icons: q.has('icons')
    ? (name) => (
        <svg data-host-icon={name} viewBox="0 0 24 24">
          <circle cx={12} cy={12} r={8} fill="currentColor" />
        </svg>
      )
    : undefined,
  collab: q.get('collab') ? { transport: q.get('collab')!, room: q.get('room') ?? 'e2e', base: Number(q.get('base')) || 0, user: { name: q.get('user') ?? 'Tester' } } : undefined,
}

function App() {
  const [props, setProps] = useState(initial)
  window.__setProps = (p) => setProps((cur) => ({ ...cur, ...p }))
  window.__setLanguage = (language) => setProps((cur) => ({ ...cur, strings: language === 'de-DE' ? de : undefined }))
  return <DocViewer {...props} />
}

createRoot(document.getElementById('root')!).render(<App />)
