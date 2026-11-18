import { useEffect, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { createPortal } from 'react-dom'
import { PromptDialog } from '../../src/react/common/Dialog'
import { Btn, Dropdown, Pop, ToolbarRow } from '../../src/react/edit/Toolbar'
import { OfficeUIProvider, type ButtonProps, type DialogProps, type SelectProps } from '../../src/react/ui'
import '../../src/react/styles.css'

window.__events = []
const noop = () => {}
function Button(p: ButtonProps) {
  return <button type="button" aria-label={p.title} onClick={p.onClick} onMouseDown={p.onMouseDown} style={p.iconName === 'more' ? undefined : { width: 34 }}> {p.iconName === 'more' ? 'More tools' : p.title} </button>
}
function WideButton(p: ButtonProps) {
  return <span style={{ display: 'inline-flex', padding: '0 20px' }}><Button {...p} /></span>
}
function DelayedDialog(p: DialogProps) {
  const [ready, setReady] = useState(false)
  useEffect(() => setReady(true), [])
  return <section role="dialog" aria-label={p.title}>
    <form onSubmit={(e) => { e.preventDefault(); p.onSubmit?.() }}>{ready && p.children}<button hidden type="submit" /></form>
    <footer>{p.actions.map((a) => <button key={a.label} type="button" onClick={a.onClick}>{a.label}</button>)}</footer>
  </section>
}
function PortalSelect(p: SelectProps) {
  const [open, setOpen] = useState(false)
  return <>
    <button type="button" aria-label={p.title} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen((v) => !v)}>{p.title}</button>
    {open && createPortal(<div role="listbox" aria-label={p.title} style={{ position: 'fixed', top: 300, left: 10 }}>
      {p.options.map((o) => <button key={o.value} role="option" type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => { p.onChange(o.value); setOpen(false) }}>{o.label}</button>)}
    </div>, document.body)}
  </>
}
const picker = (title: string) => <Dropdown title={title} options={[{ value: 'chosen', label: 'Choose option' }]} onChange={(value) => window.__events.push({ type: 'pick', value, title })} />
function App() {
  const [wide, setWide] = useState(false)
  const [tick, setTick] = useState(0)
  return <OfficeUIProvider components={{ Button: wide ? WideButton : Button, Dialog: new URLSearchParams(location.search).has('defaultDialog') ? undefined : DelayedDialog, Select: PortalSelect }}>
    <div className="de-root" style={{ width: 616 }}>
      <button onClick={() => setWide((v) => !v)}>Replace components</button>
      <button onClick={() => setTick((n) => n + 1)}>Rerender {tick}</button>
      <button onMouseDown={(e) => e.stopPropagation()}>Outside with stopped propagation</button>
      <ToolbarRow>{Array.from({ length: 17 }, (_, i) => <Btn key={i} title={String(i)} label={String(i)} onClick={noop} />)}</ToolbarRow>
      <div data-portal-overflow="" style={{ width: 180 }}><ToolbarRow>
        {Array.from({ length: 8 }, (_, i) => <Btn key={i} title={String(i)} label={String(i)} onClick={noop} />)}
        {picker('Overflow picker')}
      </ToolbarRow></div>
      <Pop title="Open popover" icon="table">{() => picker('Popover picker')}</Pop>
      <PromptDialog spec={{ title: 'Rename', initial: 'Report', onOk: (value) => window.__events.push({ type: 'submit', value }) }} onClose={() => window.__events.push({ type: 'close' })} />
    </div>
  </OfficeUIProvider>
}
createRoot(document.getElementById('root')!).render(<App />)
