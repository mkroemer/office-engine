/**
 * Which functions of Excel and Google Sheets the formula engine computes: writes `docs/formula-coverage.md`.
 *
 *   bun scripts/formula-coverage.ts             from the lists kept in scripts/formula-coverage/
 *   bun scripts/formula-coverage.ts --refresh   downloads the two lists again first
 *   bun scripts/formula-coverage.ts --out FILE  writes the report there
 *
 * The lists are the ones the two help centers publish. Run it again after adding a function.
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { supportedFunctions } from '../src/core/xlsx/formula'

const EXCEL_URL = 'https://support.microsoft.com/en-us/excel/excel-functions-alphabetical'
const SHEETS_URL = 'https://support.google.com/docs/table/25273?hl=en'
const DIR = join(import.meta.dir, 'formula-coverage')
const UA = 'Mozilla/5.0 (compatible; formula-coverage)'

interface Entry {
  name: string
  category: string
  url: string
  /** the version of Excel that added it, when the list says */
  since?: string
  /** listed under Compatibility: an old name of a function */
  old?: boolean
}

interface List {
  read: string
  functions: Entry[]
}

/** the categories of Excel, shortened; those of Google Sheets that have no other name stay */
const CATEGORY: Record<string, string> = {
  'Math and trigonometry': 'Math',
  'Date and time': 'Date',
  'Lookup and reference': 'Lookup',
  'Look and reference': 'Lookup',
  Information: 'Info',
  'Add-in and Automation': 'Add-in',
  Filter: 'Lookup',
}
const ORDER = ['Math', 'Statistical', 'Text', 'Date', 'Logical', 'Lookup', 'Info', 'Array', 'Financial', 'Engineering', 'Database', 'Operator', 'Parser', 'Web', 'Google', 'AI', 'Cube', 'Add-in']

/** the old names Excel lists under Compatibility that are not statistical */
const OLD: Record<string, string> = { CEILING: 'Math', FLOOR: 'Math', CONCATENATE: 'Text' }

/** everyday functions, picked by hand: the first to write when one is missing */
const CORE = new Set(
  (
    'ACOS AGGREGATE ASIN ATAN ATAN2 COS DEGREES MROUND QUOTIENT RADIANS SEQUENCE SIN SUBTOTAL SUMSQ TAN AVERAGEA CORREL FORECAST FORECAST.LINEAR FREQUENCY INTERCEPT ' +
    'PERCENTILE PERCENTILE.EXC PERCENTILE.INC QUARTILE QUARTILE.EXC QUARTILE.INC RANK.AVG SLOPE CLEAN DOLLAR FIXED NUMBERVALUE REGEXEXTRACT REGEXREPLACE REGEXTEST TEXTAFTER ' +
    'TEXTBEFORE TEXTSPLIT UNICHAR UNICODE DATEVALUE DAYS360 ISOWEEKNUM NETWORKDAYS.INTL TIMEVALUE WEEKNUM WORKDAY WORKDAY.INTL YEARFRAC LAMBDA LET ADDRESS FILTER FORMULATEXT ' +
    'SORT SORTBY TRANSPOSE UNIQUE XMATCH CELL ISFORMULA FV IPMT IRR NPER NPV PMT PPMT PV RATE SLN XIRR XNPV'
  ).split(' '),
)

/** the value comes from a service or from the application itself: out of reach of a library that runs in the browser alone */
const OUT_OF_REACH = new Set(
  (
    'DETECTLANGUAGE TRANSLATE IMAGE RTD INFO STOCKHISTORY IMPORTDATA IMPORTFEED IMPORTHTML IMPORTRANGE IMPORTXML WEBSERVICE GOOGLEFINANCE GOOGLETRANSLATE SPARKLINE AI ' +
    'CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE CALL REGISTER.ID'
  ).split(' '),
)

/** functions of the engine that neither list names, and why they are there */
const UNLISTED: Record<string, string> = {
  SINGLE: 'how a file writes the `@` operator',
  ANCHORARRAY: 'how a file writes the `#` operator: `A1#`',
}

/** what a function computed here does not do as its application does */
const LIMITS: [string, string][] = [
  [
    'FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT',
    'the model Excel names (exponential triple smoothing, additive), fitted by least squares: Excel does not publish how it fits its own, so a forecast may differ from Excel\'s in its last digits',
  ],
  ['GROUPBY PIVOTBY', 'one function at a time (not an array of functions side by side); PIVOTBY takes one column of values'],
  ['QUERY', 'select, where, group by, order by, limit, offset, label; no `pivot`; `format` and `options` are read and left aside'],
  ['FILTERXML', 'paths with `/`, `//`, `*`, `@attribute`, `text()` and predicates on position, attributes and children: not the whole of XPath'],
  ['REGEXTEST REGEXEXTRACT REGEXREPLACE REGEXMATCH', 'the regular expressions of JavaScript, close to the PCRE2 of Excel and the RE2 of Google Sheets, not the same in every corner'],
  ['ASC DBCS JIS FINDB LEFTB LENB MIDB REPLACEB RIGHTB SEARCHB', 'as Excel without a double-byte language set: the byte functions count characters'],
  ['PHONETIC', 'the text itself: the phonetic reading a file keeps for Japanese text is not read'],
  ['TO_DATE TO_DOLLARS TO_PERCENT TO_PURE_NUMBER', 'give the number as it is: how a cell shows it is the cell\'s format'],
  ['CELL', '`address`, `col`, `row`, `contents`, `type`, `format`, `color`, `parentheses`, `prefix`, `protect`, `width`; `filename` is empty'],
  ['GETPIVOTDATA', 'reads the values a pivot table of the file shows (the engine does not compute pivot tables)'],
  ['SUBTOTAL AGGREGATE', 'a row hidden inside a filter or a table counts as filtered out; hiding rows does not by itself compute the formulas again'],
  ['BESSELI BESSELJ BESSELK BESSELY', 'computed to full precision: the examples of the help center show the fewer digits of older versions of Excel'],
]

const text = (s: string) =>
  s
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&nbsp;|&#160;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()

const rowsOf = (html: string) => [...html.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g)].map((m) => [...m[1].matchAll(/<td[^>]*>([\s\S]*?)<\/td>/g)].map((c) => c[1]))

function parseExcel(html: string): Entry[] {
  const out: Entry[] = []
  for (const cells of rowsOf(html)) {
    if (cells.length < 2) continue
    const href = /href="([^"]+)"/.exec(cells[0])?.[1]
    const cat = /<strong>\s*([^<:]+?)\s*:?\s*<\/strong>/.exec(cells[1])?.[1]
    if (!href || !cat) continue
    const label = text(cells[0])
    const since = /\(([^)]+)\)\s*$/.exec(label)?.[1]
    const url = href.startsWith('http') ? href : href.startsWith('/') ? `https://support.microsoft.com${href}` : `https://support.microsoft.com/en-us/excel/${href}`
    for (const part of label.replace(/\([^)]*\)\s*$/, '').split(',')) {
      // (the list has a stray letter after a name or two: BETA.INVn)
      const name = part.replace(/\bfunctions?\b/i, '').trim().replace(/[a-z]+$/, '')
      if (!/^[A-Z][A-Z0-9_.]*$/.test(name)) continue
      const old = cat === 'Compatibility'
      out.push({ name, category: old ? (OLD[name] ?? 'Statistical') : (CATEGORY[cat] ?? cat), url, ...(since ? { since } : {}), ...(old ? { old } : {}) })
    }
  }
  return out
}

function parseSheets(html: string): Entry[] {
  const out = new Map<string, Entry>()
  for (const cells of rowsOf(html)) {
    if (cells.length < 4) continue
    const name = text(cells[1])
    const cat = text(cells[0])
    const href = /href="([^"]+)"/.exec(cells[3])?.[1]
    if (!/^[A-Z][A-Z0-9_.]*$/.test(name) || !href || out.has(name)) continue
    out.set(name, { name, category: CATEGORY[cat] ?? cat, url: href.startsWith('http') ? href : `https://support.google.com${href}` })
  }
  return [...out.values()]
}

async function refresh() {
  mkdirSync(DIR, { recursive: true })
  const today = new Date().toISOString().slice(0, 10)
  for (const [file, url, parse] of [
    ['excel.json', EXCEL_URL, parseExcel],
    ['sheets.json', SHEETS_URL, parseSheets],
  ] as const) {
    const res = await fetch(url, { headers: { 'user-agent': UA, 'accept-language': 'en' } })
    if (!res.ok) throw new Error(`${url}: ${res.status}`)
    const functions = parse(await res.text())
    if (functions.length < 300) throw new Error(`${url}: only ${functions.length} functions read: the page is not the list it was`)
    writeFileSync(join(DIR, file), JSON.stringify({ read: today, functions } satisfies List, null, 1) + '\n')
    console.log(`${file}: ${functions.length} functions`)
  }
}

const load = (file: string): List => JSON.parse(readFileSync(join(DIR, file), 'utf8'))

const pct = (a: number, b: number) => (b ? `${Math.round((100 * a) / b)}%` : '-')
const code = (names: string[]) => names.map((n) => `\`${n}\``).join(' ')

function report(): string {
  const excel = load('excel.json')
  const sheets = load('sheets.json')
  const have = new Set(supportedFunctions())
  const inExcel = new Map(excel.functions.map((f) => [f.name, f]))
  const inSheets = new Map(sheets.functions.map((f) => [f.name, f]))
  const all = [...new Set([...inExcel.keys(), ...inSheets.keys()])].sort()
  // the category of Excel; of Google Sheets for what it alone has, and for the old names of Excel it has too
  const category = (n: string) => {
    const e = inExcel.get(n)
    const g = inSheets.get(n)
    return e && !(e.old && g && ORDER.includes(g.category)) ? e.category : (g ?? e)!.category
  }
  const missing = all.filter((n) => !have.has(n))
  const reachable = missing.filter((n) => !OUT_OF_REACH.has(n))
  const both = all.filter((n) => inExcel.has(n) && inSheets.has(n))
  const count = (names: string[]) => names.filter((n) => have.has(n)).length
  const line = (label: string, names: string[]) => `| ${label} | ${names.length} | ${count(names)} | ${names.length - count(names)} | ${pct(count(names), names.length)} |`
  const link = (n: string, of: Map<string, Entry>) => {
    const f = of.get(n)
    return f ? `[${f.since ?? 'yes'}](${f.url})` : ''
  }
  const table = (names: string[], withCategory: boolean) =>
    [
      `| Function |${withCategory ? ' Category |' : ''} Excel | Google Sheets | Notes |`,
      `| --- |${withCategory ? ' --- |' : ''} --- | --- | --- |`,
      ...names.map((n) => `| \`${n}\` |${withCategory ? ` ${category(n)} |` : ''} ${link(n, inExcel)} | ${link(n, inSheets)} | ${inExcel.get(n)?.old ? 'old name' : ''} |`),
    ].join('\n')
  const byCategory = (names: string[]) => ORDER.map((c) => [c, names.filter((n) => category(n) === c)] as const).filter(([, list]) => list.length)

  const out: string[] = []
  out.push('# Formula function coverage', '')
  out.push(
    `> Written by \`bun scripts/formula-coverage.ts\` (do not edit by hand: run it again after adding a function). It compares the engine's function registry (\`src/core/xlsx/formula/functions.ts\`) with the function lists of [Excel](${EXCEL_URL}) (read ${excel.read}) and [Google Sheets](${SHEETS_URL}) (read ${sheets.read}). \`--refresh\` downloads the two lists again.`,
    '',
  )
  out.push(
    'The lists are taken as published: a function a help center has not listed is counted as not being there (Google Sheets computes `XMATCH`, its list does not have it), and neither list tells how much of a function the engine computes: only whether it has one of that name. What is known to differ is under [Limits](#limits).',
    '',
  )
  out.push('## Summary', '', '| Reference | Functions | Computed | Missing | Coverage |', '| --- | ---: | ---: | ---: | ---: |')
  out.push(line('Excel', [...inExcel.keys()]), line('Google Sheets', [...inSheets.keys()]), line('In both', both), line('In either', all), '')
  const extra = [...have].filter((n) => !inExcel.has(n) && !inSheets.has(n)).sort()
  out.push(
    `The engine computes **${count(all)}** of them${extra.length ? ` (and ${code(extra)}, which neither list names)` : ''}; **${missing.length}** are missing${
      missing.length === missing.length - reachable.length ? ', all of them out of reach (see below)' : ''
    }. A formula that uses a missing one is not computed: it keeps the value the file holds, which no longer follows the cells it reads, and shows \`#NAME?\` when typed in the editor.`,
    '',
  )
  if (reachable.length) {
    const core = reachable.filter((n) => CORE.has(n))
    const shared = reachable.filter((n) => !CORE.has(n) && inExcel.has(n) && inSheets.has(n))
    const one = reachable.filter((n) => !CORE.has(n) && !(inExcel.has(n) && inSheets.has(n)))
    out.push('| Priority | Missing | What |', '| --- | ---: | --- |')
    out.push(`| 1 | ${core.length} | everyday functions, picked by hand (\`CORE\` in the script) |`)
    out.push(`| 2 | ${shared.length} | the other functions both applications have: a workbook of either may use them |`)
    out.push(`| 3 | ${one.length} | functions only one of the two has |`)
    out.push(`| [none](#out-of-reach) | ${missing.length - reachable.length} | the value comes from a service or the application itself: out of reach of a library that runs in the browser alone |`, '')
  }
  out.push('## By category', '', '| Category | Functions | Computed | Missing | Coverage |', '| --- | ---: | ---: | ---: | ---: |')
  for (const [c, names] of byCategory(all)) out.push(line(c, names))
  out.push(
    '',
    'Categories are those of Excel, shortened; a function only Google Sheets has keeps the category it has there (*Array*, *Operator*, *Parser*, *Google*, *AI*). The old names Excel lists under *Compatibility* are counted with what they compute.',
    '',
  )
  out.push('## What the engine evaluates besides functions', '')
  out.push(
    '- **Names inside a formula**: `LET` and `LAMBDA` (with optional parameters, `ISOMITTED`, a LAMBDA called at once or passed to `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY`, a function named without being called: `BYROW(rows, SUM)`).',
    '- **Dynamic arrays**: a formula that gives several values fills the cells next to it; `#SPILL!` when a cell is in the way; `A1#` for the cells a formula fills, `@` for the one value of a range that is the formula\'s own. Saved as Excel saves them (an array formula marked dynamic), so they stay dynamic there and here.',
    '- **Defined names**, of the workbook and local to a sheet, **structured references** (`Table1[Qty]`, `Table1[@Qty]`, `[@Qty]`, `Table1[[#Headers],[Qty]]`, `Table1[#All]`) and **3-D references** (`First:Last!A1`).',
    '',
  )
  if (reachable.length) {
    const core = reachable.filter((n) => CORE.has(n))
    if (core.length) out.push('## Priority 1: everyday functions', '', `${core.length} functions. In the columns, the version of Excel that added the function (or *yes*) links to its documentation.`, '', table(core, true), '')
    const shared = reachable.filter((n) => !CORE.has(n) && inExcel.has(n) && inSheets.has(n))
    if (shared.length) {
      out.push('## Priority 2: in both applications', '')
      for (const [c, names] of byCategory(shared)) out.push(`### ${c} (${names.length})`, '', table(names, false), '')
    }
    const onlyExcel = reachable.filter((n) => !CORE.has(n) && inExcel.has(n) && !inSheets.has(n))
    const onlySheets = reachable.filter((n) => !CORE.has(n) && !inExcel.has(n) && inSheets.has(n))
    if (onlyExcel.length || onlySheets.length) out.push('## Priority 3: in one application only', '')
    if (onlyExcel.length) out.push(`### Excel only (${onlyExcel.length})`, '', table(onlyExcel, true), '')
    if (onlySheets.length) out.push(`### Google Sheets only (${onlySheets.length})`, '', table(onlySheets, true), '')
  }
  const out_ = missing.filter((n) => OUT_OF_REACH.has(n))
  if (out_.length) {
    out.push('## Out of reach', '', 'Not to be written: the value the file holds is the best there is.', '')
    out.push(table(ORDER.flatMap((c) => out_.filter((n) => category(n) === c)), true), '')
  }
  out.push('## Limits', '', 'Functions the engine computes, and what it does not do as their application does:', '')
  out.push('| Functions | Limit |', '| --- | --- |')
  for (const [names, note] of LIMITS) {
    const list = names.split(' ').filter((n) => have.has(n))
    if (list.length) out.push(`| ${code(list)} | ${note} |`)
  }
  out.push(
    '',
    'The engine opens the formats of Excel, and Google Sheets saves a function Excel does not have as `IFERROR(__xludf.DUMMYFUNCTION("…"), value)`: such a file opens with its values. The functions of Google Sheets count for formulas typed in the editor.',
    '',
  )
  out.push('## Computed today', '')
  for (const [c, names] of byCategory(all.filter((n) => have.has(n)))) out.push(`- **${c}** (${names.length}): ${code(names)}`, '')
  if (extra.length) out.push(`- **In neither list** (${extra.length}): ${extra.map((n) => `\`${n}\`${UNLISTED[n] ? ` (${UNLISTED[n]})` : ''}`).join(', ')}`, '')
  return out.join('\n')
}

const args = process.argv.slice(2)
if (args.includes('--refresh')) await refresh()
const at = args.indexOf('--out')
const target = at >= 0 && args[at + 1] ? args[at + 1] : join(import.meta.dir, '..', 'docs', 'formula-coverage.md')
mkdirSync(dirname(target), { recursive: true })
writeFileSync(target, report())
console.log(`written: ${target}`)
