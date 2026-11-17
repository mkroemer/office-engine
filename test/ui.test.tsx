import { describe, expect, test } from 'bun:test'
import type { ReactNode } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import type { Icon as PublicIcon } from '../src'
import { ConfirmDialog, PromptDialog, Toast } from '../src/react/common/Dialog'
import { Btn, ColorBtn, Dropdown, Icon, Pop, Sep } from '../src/react/edit/Toolbar'
import { OfficeUIProvider, UIComponentsProvider, type ButtonProps, type DialogProps, type SelectProps, type UIComponents } from '../src/react/ui'

const seen: { button: ButtonProps[]; select: SelectProps[]; dialog: DialogProps[] } = { button: [], select: [], dialog: [] }
const host: Partial<UIComponents> = {
  Button: (p) => {
    seen.button.push(p)
    return (
      <button data-host="button" className={p.className} title={p.title}>
        {p.icon}
        {p.label}
        {p.children}
      </button>
    )
  },
  Select: (p) => {
    seen.select.push(p)
    return <span data-host="select">{p.options.length}</span>
  },
  Separator: () => <hr data-host="sep" />,
  Dialog: (p) => {
    seen.dialog.push(p)
    return (
      <section data-host="dialog" aria-label={p.title}>
        {p.children}
        {p.actions.map((a) => a.label).join('|')}
      </section>
    )
  },
  Toast: ({ text }) => <output data-host="toast">{text}</output>,
}
/** renders with the host components, or with `null` with the built-in ones */
const html = (node: ReactNode, components: Partial<UIComponents> | null = host) => renderToStaticMarkup(<UIComponentsProvider components={components ?? undefined}>{node}</UIComponentsProvider>)
const noop = () => {}

describe('UI components', () => {
  test('without host components the built-in markup is unchanged', () => {
    expect(html(<Btn icon="bold" title="Bold" active onClick={noop} />, null)).toStartWith('<button type="button" class="de-tb de-tb-on" title="Bold" aria-label="Bold" aria-pressed="true">')
    expect(html(<Sep />, null)).toBe('<span class="de-tb-sep"></span>')
    expect(html(<Dropdown title="Style" options={[{ value: 'a', label: 'A' }]} onChange={noop} value="a" />, null)).toContain('class="de-dd"')
    expect(html(<Toast text="Saved" onDone={noop} />, null)).toBe('<div class="de-toast" role="status">Saved</div>')
    expect(html(<PromptDialog spec={{ title: 'Link', label: 'URL', initial: 'https://x', placeholder: 'https://', note: 'A link', onOk: noop, extra: { label: 'Remove link', run: noop } }} onClose={noop} />, null)).toBe(
      '<div class="de-dialog-back"><form class="de-dialog" role="dialog" aria-label="Link"><div class="de-dialog-title">Link</div><label class="de-dialog-label">URL</label><input placeholder="https://" value="https://x"/><div class="de-dialog-note">A link</div><div class="de-dialog-btns"><button type="button" class="de-btn de-dialog-extra">Remove link</button><button type="button" class="de-btn">Cancel</button><button type="submit" class="de-btn de-btn-primary">OK</button></div></form></div>',
    )
    expect(html(<PromptDialog spec={{ title: 'Notes', area: true, initial: 'Text', onOk: noop }} onClose={noop} />, null)).toBe(
      '<div class="de-dialog-back"><form class="de-dialog" role="dialog" aria-label="Notes"><div class="de-dialog-title">Notes</div><textarea rows="3">Text</textarea><div class="de-dialog-btns"><button type="button" class="de-btn">Cancel</button><button type="submit" class="de-btn de-btn-primary">OK</button></div></form></div>',
    )
    expect(html(<ConfirmDialog title="Delete?" message="Gone for good" ok="Delete" onOk={noop} onClose={noop} />, null)).toBe(
      '<div class="de-dialog-back"><div class="de-dialog" role="alertdialog" aria-label="Delete?"><div class="de-dialog-title">Delete?</div><div class="de-dialog-note">Gone for good</div><div class="de-dialog-btns"><button type="button" class="de-btn">Cancel</button><button type="button" class="de-btn de-btn-primary" autofocus="">Delete</button></div></div></div>',
    )
  })

  test('Button gets title, resolved icon, state, the engine classes and the focus-keeping mouse down', () => {
    seen.button.length = 0
    const out = html(<Btn icon="bold" title="Bold" active onClick={noop} />)
    expect(out).toContain('data-host="button"')
    expect(out).toContain('<svg')
    const p = seen.button[0]
    expect(p).toMatchObject({ title: 'Bold', iconName: 'bold', active: true, className: 'de-tb de-tb-on' })
    let prevented = false
    p.onMouseDown({ preventDefault: () => (prevented = true) } as never)
    expect(prevented).toBe(true)
  })

  test('popup buttons pass expanded and their extra content', () => {
    seen.button.length = 0
    html(<ColorBtn icon="color" title="Text color" value="#ff0000" onPick={noop} />)
    html(<Pop icon="table" title="Insert table" chevron>{() => null}</Pop>)
    expect(seen.button.map((p) => [p.title, p.expanded, !!p.children])).toEqual([
      ['Text color', false, true],
      ['Insert table', false, true],
    ])
    expect(seen.button[1].className).toBe('de-tb de-tb-chev')
  })

  test('Separator and Select are replaced; Select gets the options and value', () => {
    seen.select.length = 0
    expect(html(<Sep />)).toBe('<hr data-host="sep"/>')
    expect(html(<Dropdown title="Style" options={[{ value: 'a', label: 'A' }, { value: 'b', label: 'B' }]} onChange={noop} value="b" />)).toBe('<span data-host="select">2</span>')
    expect(seen.select[0]).toMatchObject({ title: 'Style', value: 'b' })
  })

  test('dialogs: prompt submits through its primary action, confirm is an alert', () => {
    seen.dialog.length = 0
    const out = html(<PromptDialog spec={{ title: 'Link', initial: 'https://x', onOk: noop, extra: { label: 'Remove link', run: noop } }} onClose={noop} />)
    expect(out).toContain('aria-label="Link"')
    expect(out).toContain('value="https://x"')
    expect(out).toContain('Remove link|Cancel|OK')
    html(<ConfirmDialog title="Delete?" message="Gone for good" ok="Delete" onOk={noop} onClose={noop} />)
    const [prompt, confirm] = seen.dialog
    expect(prompt.onSubmit).toBeFunction()
    expect(prompt.actions[0]).toMatchObject({ label: 'Remove link', extra: true })
    expect(prompt.actions.at(-1)).toMatchObject({ primary: true, submit: true })
    expect(confirm.kind).toBe('alert')
    expect(confirm.onSubmit).toBeUndefined()
    expect(confirm.actions.at(-1)).toMatchObject({ label: 'Delete', primary: true, autoFocus: true })
  })

  test('host Select can render option icons through the public Icon export and host icon provider', () => {
    // The root entry includes Vite worker imports; check its public type without loading it in Bun.
    const OptionIcon: typeof PublicIcon = Icon
    const Select = (p: SelectProps) => <>{p.options.map((o) => o.icon && <OptionIcon key={o.value} name={o.icon} size={16} />)}</>
    const out = renderToStaticMarkup(
      <OfficeUIProvider components={{ Select }} icons={{ bold: <b>host bold</b> }}>
        <Dropdown title="Format" options={[{ value: 'b', label: 'Bold', icon: 'bold' }, { value: 'i', label: 'Italic', icon: 'italic' }]} onChange={noop} />
      </OfficeUIProvider>,
    )
    expect(out).toContain('<span class="de-icon" style="font-size:16px;width:16px;height:16px" aria-hidden="true"><b>host bold</b></span>')
    expect(out).toContain('<svg width="16" height="16"')
  })

  test('Toast', () => {
    expect(html(<Toast text="Saved" onDone={noop} />)).toBe('<output data-host="toast">Saved</output>')
  })

  test('nested providers: inner components win, the outer ones stay for the rest', () => {
    const out = renderToStaticMarkup(
      <UIComponentsProvider components={host}>
        <UIComponentsProvider components={{ Separator: () => <i>inner</i> }}>
          <Sep />
          <Toast text="t" onDone={noop} />
        </UIComponentsProvider>
      </UIComponentsProvider>,
    )
    expect(out).toBe('<i>inner</i><output data-host="toast">t</output>')
  })
})
