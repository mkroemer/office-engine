import { expect, test } from 'bun:test'
import { useBrowser, viewer } from './harness'

useBrowser()

test('More keeps its natural width in a 600px row and replacement host components are remeasured', async () => {
  const v = await viewer({}, { page: 'ui.html', wait: false })
  const more = v.page.locator('.de-ebar').first().locator('.de-ebar-morewrap button')
  await more.waitFor()
  const bounds = () => v.page.locator('.de-ebar').first().evaluate((row) => {
    const wrapper = row.querySelector<HTMLElement>('.de-ebar-morewrap')!
    const clone = wrapper.cloneNode(true) as HTMLElement
    clone.style.cssText = 'position:absolute;width:max-content;visibility:hidden'
    row.append(clone)
    const natural = clone.getBoundingClientRect().width
    clone.remove()
    return {
      natural, width: wrapper.getBoundingClientRect().width,
      right: wrapper.getBoundingClientRect().right, edge: row.getBoundingClientRect().right - 8,
      cut: row.querySelectorAll(':scope > [data-i]').length,
    }
  })
  const before = await bounds()
  console.info('initial More bounds', JSON.stringify(before))
  await v.page.getByText('Replace components', { exact: true }).click()
  const after = await bounds()
  console.info('replacement More bounds', JSON.stringify(after))
  expect(before.width).toBeCloseTo(before.natural, 1)
  expect(before.right).toBeLessThanOrEqual(before.edge + 1)
  expect(after.cut).toBeLessThan(before.cut)
  expect(after.width).toBeCloseTo(after.natural, 1)
  expect(after.right).toBeLessThanOrEqual(after.edge + 1)
  // Unrelated host rerenders must not read toolbar item geometry again.
  await v.page.evaluate(() => {
    const original = Element.prototype.getBoundingClientRect
    Element.prototype.getBoundingClientRect = function () {
      if (this.matches('.de-ebar > [data-i], .de-ebar-morewrap')) window.__events.push({ type: 'measure' })
      return original.call(this)
    }
  })
  await v.page.getByText('Rerender 0', { exact: true }).click()
  console.info('geometry reads on unrelated rerender', (await v.events('measure')).length)
  expect(await v.events('measure')).toHaveLength(0)
  await more.click()
  expect(await v.page.locator('.de-ebar-more button').count()).toBeGreaterThan(0)
  expect(v.errors).toEqual([])
})

test('a delayed host dialog focuses and selects the prompt input', async () => {
  const v = await viewer({}, { page: 'ui.html', wait: false })
  const input = v.page.getByRole('dialog').locator('input')
  await input.waitFor()
  const focus = await input.evaluate((el: HTMLInputElement) => ({ focused: document.activeElement === el, start: el.selectionStart, end: el.selectionEnd }))
  console.info('delayed prompt focus/selection', JSON.stringify(focus))
  expect(focus).toEqual({ focused: true, start: 0, end: 6 })
  expect(v.errors).toEqual([])
})

test('a prompt host footer outside the form and Enter each submit once', async () => {
  const v = await viewer({}, { page: 'ui.html', wait: false })
  const input = v.page.getByRole('dialog').locator('input')
  await input.waitFor()
  expect(await v.page.getByRole('button', { name: 'OK', exact: true }).evaluate((el) => !!el.closest('form'))).toBe(false)
  await input.fill('From footer')
  await v.page.getByRole('button', { name: 'OK', exact: true }).click()
  console.info('external footer submit events', JSON.stringify(await v.events('submit')))
  expect(await v.events('submit')).toEqual([{ type: 'submit', value: 'From footer' }])
  expect(await v.events('close')).toHaveLength(1)
  await input.fill('From form')
  await input.press('Enter')
  console.info('footer then Enter submit events', JSON.stringify(await v.events('submit')))
  expect(await v.events('submit')).toEqual([{ type: 'submit', value: 'From footer' }, { type: 'submit', value: 'From form' }])
  expect(await v.events('close')).toHaveLength(2)
  expect(v.errors).toEqual([])
})

test('the built-in prompt submits once per primary click or Enter', async () => {
  const v = await viewer({ defaultDialog: true }, { page: 'ui.html', wait: false })
  const input = v.page.getByRole('dialog').locator('input')
  await input.fill('Click')
  await v.page.getByRole('button', { name: 'OK', exact: true }).click()
  expect(await v.events('submit')).toEqual([{ type: 'submit', value: 'Click' }])
  await input.fill('Enter')
  await input.press('Enter')
  expect(await v.events('submit')).toEqual([{ type: 'submit', value: 'Click' }, { type: 'submit', value: 'Enter' }])
  expect(await v.events('close')).toHaveLength(2)
  expect(v.errors).toEqual([])
})

for (const placement of ['overflow', 'popover'] as const) {
  test(`a React-portaled host Select in ${placement} survives mousedown and delivers onChange`, async () => {
    const v = await viewer({}, { page: 'ui.html', wait: false })
    const title = placement === 'overflow' ? 'Overflow picker' : 'Popover picker'
    const parent = placement === 'overflow' ? v.page.locator('[data-portal-overflow] .de-ebar-more') : v.page.locator('.de-tb-pop').last()
    if (placement === 'overflow') await v.page.locator('[data-portal-overflow] .de-ebar-morewrap > button').click()
    else await v.page.getByRole('button', { name: 'Open popover', exact: true }).click()
    await v.page.getByRole('button', { name: title, exact: true }).click()
    const option = v.page.getByRole('option', { name: 'Choose option' })
    await option.waitFor()
    expect(await option.evaluate((el) => !!el.closest('.de-ebar-more, .de-tb-pop'))).toBe(false)
    await option.click()
    console.info(`${placement} portal after option click`, JSON.stringify({ parentVisible: await parent.isVisible(), events: await v.events('pick') }))
    expect(await v.events('pick')).toEqual([{ type: 'pick', value: 'chosen', title }])
    expect(await parent.isVisible()).toBe(true)
    await v.page.getByText('Outside with stopped propagation', { exact: true }).click()
    await parent.waitFor({ state: 'detached', timeout: 5000 })
    expect(await parent.count()).toBe(0)
    expect(v.errors).toEqual([])
  })
}
