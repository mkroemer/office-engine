import { describe, expect, test } from 'bun:test'
import { renderToStaticMarkup } from 'react-dom/server'
import { hostIcon, Icon, ICON_NAMES, IconsProvider, type IconName, type Icons } from '../src/react/edit/Toolbar'

const html = (icons: Icons | undefined, name: IconName, size?: number) =>
  renderToStaticMarkup(
    <IconsProvider icons={icons}>
      <Icon name={name} size={size} />
    </IconsProvider>,
  )

describe('icons', () => {
  test('names cover the controls that had text glyphs', () => {
    for (const n of ['zoomIn', 'zoomOut', 'outline', 'present', 'sheetAdd', 'replace', 'close'] as const) expect(ICON_NAMES).toContain(n)
  })

  test('built-in svg without host icons', () => {
    const out = html(undefined, 'bold')
    expect(out).toStartWith('<svg')
    expect(out).not.toContain('de-icon')
  })

  test('map: listed names replace, missing names stay built-in', () => {
    const icons = { bold: <i>B</i> }
    expect(html(icons, 'bold', 20)).toBe('<span class="de-icon" style="font-size:20px;width:20px;height:20px" aria-hidden="true"><i>B</i></span>')
    expect(html(icons, 'italic')).toStartWith('<svg')
  })

  test('resolver: gets name and size; undefined, null and false keep the built-in icon', () => {
    const seen: [IconName, number][] = []
    const icons: Icons = (name, { size }) => {
      seen.push([name, size])
      if (name === 'undo') return <b>{size}</b>
      if (name === 'redo') return null
      if (name === 'zoomIn') return false
      return undefined
    }
    expect(html(icons, 'undo', 16)).toContain('<b>16</b>')
    expect(html(icons, 'redo')).toStartWith('<svg')
    expect(html(icons, 'zoomIn')).toStartWith('<svg')
    expect(html(icons, 'cut')).toStartWith('<svg')
    expect(seen).toEqual([
      ['undo', 16],
      ['redo', 18],
      ['zoomIn', 18],
      ['cut', 18],
    ])
  })

  test('an empty string or 0 is a host icon, not a fallback', () => {
    expect(hostIcon({ bold: '' }, 'bold', 18)).toBe('')
    expect(hostIcon(() => 0, 'bold', 18)).toBe(0)
  })

  test('nested providers: inner first, then the enclosing one, then built-in', () => {
    const out = (name: IconName) =>
      renderToStaticMarkup(
        <IconsProvider icons={() => <u>outer</u>}>
          <IconsProvider icons={{ bold: <u>inner</u> }}>
            <Icon name={name} />
          </IconsProvider>
        </IconsProvider>,
      )
    expect(out('bold')).toContain('inner')
    expect(out('italic')).toContain('outer')
    const none = renderToStaticMarkup(
      <IconsProvider icons={{ bold: <u>outer</u> }}>
        <IconsProvider>
          <Icon name="italic" />
        </IconsProvider>
      </IconsProvider>,
    )
    expect(none).toStartWith('<svg')
  })
})
