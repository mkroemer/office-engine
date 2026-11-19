/** LET, LAMBDA and its helpers, the dynamic array functions, `@` and `#`, and how a file writes them. */
import { describe, expect, test } from 'bun:test'
import { colIndex, evaluateDynamic, evaluateFormula, FormulaError, parseFormula, spills, stripFunctionPrefixes, withFunctionPrefixes, type CellProvider, type ErrorCode, type Value } from '../src/core/xlsx/formula'

const cells: Record<string, Value> = {
  A1: 3, B1: 'c',
  A2: 1, B2: 'a',
  A3: 2, B3: 'b',
  A4: 1, B4: 'a',
  D1: 'East', E1: 10,
  D2: 'West', E2: 5,
  D3: 'East', E3: 7,
  D4: 'North', E4: 1,
}
const at = (a: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a)!
  return `${Number(m[2]) - 1},${colIndex(m[1])}`
}
const store = new Map(Object.entries(cells).map(([a, v]) => [at(a), v]))
const P: CellProvider = {
  value: (_s, r, c) => store.get(`${r},${c}`) ?? null,
  bounds: () => ({ rows: 4, cols: 5 }),
  hasSheet: (n) => n === 'S',
  spill: (_s, r, c) => (r === 0 && c === 0 ? { r1: 3, c1: 0 } : undefined),
  formula: (_s, r, c) => (r === 0 && c === 0 ? 'SEQUENCE(4)' : undefined),
}
const E = (code: ErrorCode) => new FormulaError(code)
const ev = (f: string, r = 9, c = 9) => evaluateFormula(f, { provider: P, sheet: 'S', r, c })
const arr = (f: string) => evaluateDynamic(parseFormula(f), { provider: P, sheet: 'S', r: 9, c: 9 })

describe('LET and LAMBDA', () => {
  test('names, functions, recursion limits', () => {
    expect(ev('LET(x,2,y,3,x*y)')).toBe(6)
    expect(ev('LET(x,2,y,x+1,x*y)')).toBe(6)
    expect(ev('LET(r,A1:A4,SUM(r)/ROWS(r))')).toBe(1.75)
    expect(ev('LET(f,LAMBDA(a,a+1),f(4))')).toBe(5)
    expect(ev('LET(X,2,x*3)')).toBe(6)
    expect(ev('LET(_xlpm.x,2,_xlpm.x*3)')).toBe(6)
    expect(ev('LAMBDA(x,x+1)(5)')).toBe(6)
    expect(ev('LAMBDA(a,b,a*b)(4,5)')).toBe(20)
    expect(ev('LAMBDA(a,b,a*b)(4)')).toEqual(E('#VALUE!'))
    expect(ev('LAMBDA(x,x+1)')).toEqual(E('#CALC!'))
    expect(ev('LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(5)')).toBe(5)
    expect(ev('LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(5,2)')).toBe(7)
    expect(ev('LET(x,1,x,2,x)')).toBe(2)
    expect(ev('LET(x,1)')).toEqual(E('#VALUE!'))
    expect(ev('LET(A1,1,2)')).toEqual(E('#VALUE!'))
    // a closure keeps the names it was made with
    expect(ev('LET(k,10,add,LAMBDA(x,x+k),k,1,add(5))')).toBe(15)
    // a name of LET is not known inside its own value: no recursion through LET
    expect(ev('LET(f,LAMBDA(n,IF(n<2,1,n*f(n-1))),f(5))')).toEqual(E('#NAME?'))
    // a function that calls itself without end stops with #NUM!
    expect(ev('LET(y,LAMBDA(f,x,f(f,x+1)),y(y,1))')).toEqual(E('#NUM!'))
    expect(ev('LET(y,LAMBDA(f,n,IF(n<2,1,n*f(f,n-1))),y(y,5))')).toBe(120)
  })

  test('MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY', () => {
    expect(arr('MAP({1,2,3},LAMBDA(x,x*2))')).toEqual([[2, 4, 6]])
    expect(arr('MAP(A1:A3,{10;20;30},LAMBDA(a,b,a+b))')).toEqual([[13], [21], [32]])
    expect(arr('MAP(A1:A2,LAMBDA(c,ROW(c)))')).toEqual([[1], [2]])
    expect(ev('REDUCE(0,{1,2,3},LAMBDA(a,b,a+b))')).toBe(6)
    expect(ev('REDUCE(,A1:A4,LAMBDA(a,b,MAX(a,b)))')).toBe(3)
    expect(arr('SCAN(0,{1,2,3},LAMBDA(a,b,a+b))')).toEqual([[1, 3, 6]])
    expect(arr('BYROW({1,2;3,4},LAMBDA(r,SUM(r)))')).toEqual([[3], [7]])
    expect(arr('BYCOL({1,2;3,4},SUM)')).toEqual([[4, 6]])
    expect(arr('BYROW(A1:B2,LAMBDA(r,COLUMNS(r)))')).toEqual([[2], [2]])
    expect(arr('MAKEARRAY(2,2,LAMBDA(r,c,r*c))')).toEqual([
      [1, 2],
      [2, 4],
    ])
    expect(arr('MAP({1,2},LAMBDA(x,SEQUENCE(x)))')).toEqual([[1, E('#CALC!')]])
    expect(ev('MAP({1,2},3)')).toEqual(E('#VALUE!'))
  })
})

describe('dynamic arrays', () => {
  test('SEQUENCE, TRANSPOSE, FILTER, SORT, SORTBY, UNIQUE', () => {
    expect(arr('SEQUENCE(2,3)')).toEqual([
      [1, 2, 3],
      [4, 5, 6],
    ])
    expect(arr('SEQUENCE(3,1,10,-2)')).toEqual([[10], [8], [6]])
    expect(arr('SEQUENCE(0)')).toEqual([[E('#CALC!')]])
    expect(arr('TRANSPOSE({1,2;3,4})')).toEqual([
      [1, 3],
      [2, 4],
    ])
    expect(arr('FILTER(A1:B4,A1:A4>1)')).toEqual([
      [3, 'c'],
      [2, 'b'],
    ])
    expect(arr('FILTER(A1:B4,A1:A4>9)')).toEqual([[E('#CALC!')]])
    expect(arr('FILTER(A1:B4,A1:A4>9,"none")')).toEqual([['none']])
    expect(arr('FILTER(A1:B4,{TRUE,FALSE})')).toEqual([[3], [1], [2], [1]])
    expect(arr('FILTER(A1:B4,{1,2,3})')).toEqual([[E('#VALUE!')]])
    expect(arr('SORT(A1:A4)')).toEqual([[1], [1], [2], [3]])
    expect(arr('SORT(A1:B4,2,-1)')).toEqual([
      [3, 'c'],
      [2, 'b'],
      [1, 'a'],
      [1, 'a'],
    ])
    expect(arr('SORT({3,1,2},,,TRUE)')).toEqual([[1, 2, 3]])
    expect(arr('SORT({"b";1;TRUE;"A"})')).toEqual([[1], ['A'], ['b'], [true]])
    expect(arr('SORTBY(B1:B4,A1:A4,-1)')).toEqual([['c'], ['b'], ['a'], ['a']])
    expect(arr('SORTBY(A1:B4,B1:B4,1,A1:A4,-1)')).toEqual([
      [1, 'a'],
      [1, 'a'],
      [2, 'b'],
      [3, 'c'],
    ])
    expect(arr('UNIQUE(A1:A4)')).toEqual([[3], [1], [2]])
    expect(arr('UNIQUE(A1:B4,,TRUE)')).toEqual([
      [3, 'c'],
      [2, 'b'],
    ])
    expect(arr('UNIQUE({1,1,2},TRUE)')).toEqual([[1, 2]])
    expect(arr('UNIQUE({"a";"A";"b"})')).toEqual([['a'], ['b']])
  })

  test('reshaping and stacking', () => {
    expect(arr('CHOOSECOLS({1,2,3;4,5,6},3,1)')).toEqual([
      [3, 1],
      [6, 4],
    ])
    expect(arr('CHOOSEROWS({1,2,3;4,5,6},-1)')).toEqual([[4, 5, 6]])
    expect(arr('CHOOSECOLS({1,2,3},4)')).toEqual([[E('#VALUE!')]])
    expect(arr('HSTACK({1;2},{3;4;5})')).toEqual([
      [1, 3],
      [2, 4],
      [E('#N/A'), 5],
    ])
    expect(arr('VSTACK({1,2},{3})')).toEqual([
      [1, 2],
      [3, E('#N/A')],
    ])
    expect(arr('TOCOL({1,2;3,4})')).toEqual([[1], [2], [3], [4]])
    expect(arr('TOCOL({1,2;3,4},0,TRUE)')).toEqual([[1], [3], [2], [4]])
    expect(arr('TOROW(A1:B2)')).toEqual([[3, 'c', 1, 'a']])
    expect(arr('TOCOL({1,#N/A,2},2)')).toEqual([[1], [2]])
    expect(arr('WRAPROWS({1,2,3,4,5},2)')).toEqual([
      [1, 2],
      [3, 4],
      [5, E('#N/A')],
    ])
    expect(arr('WRAPCOLS({1,2,3,4,5},2,0)')).toEqual([
      [1, 3, 5],
      [2, 4, 0],
    ])
    expect(arr('TAKE({1,2,3;4,5,6},1)')).toEqual([[1, 2, 3]])
    expect(arr('TAKE({1,2,3;4,5,6},-1,2)')).toEqual([[4, 5]])
    expect(arr('TAKE({1,2,3;4,5,6},,-1)')).toEqual([[3], [6]])
    expect(arr('DROP({1,2,3;4,5,6},1)')).toEqual([[4, 5, 6]])
    expect(arr('DROP({1,2,3;4,5,6},,-2)')).toEqual([[1], [4]])
    expect(arr('DROP({1,2,3},1)')).toEqual([[E('#CALC!')]])
    expect(arr('EXPAND({1,2},2,3,0)')).toEqual([
      [1, 2, 0],
      [0, 0, 0],
    ])
    expect(arr('EXPAND({1,2;3,4},1)')).toEqual([[E('#VALUE!')]])
    expect(ev('ROWS(TAKE(A1:A4,2))')).toBe(2)
    expect(ev('SUM(TRIMRANGE(A1:C9))')).toBe(7)
    expect(ev('ROWS(TRIMRANGE(A1:C9))')).toBe(4)
    expect(ev('COLUMNS(TRIMRANGE(A1:C9))')).toBe(2)
  })

  test('matrices', () => {
    expect(arr('MMULT({1,2;3,4},{5,6;7,8})')).toEqual([
      [19, 22],
      [43, 50],
    ])
    expect(arr('MMULT({1,2;3,4},{1,2,3})')).toEqual([[E('#VALUE!')]])
    expect(ev('MDETERM({1,2;3,4})')).toBeCloseTo(-2, 12)
    expect(ev('MDETERM({1,3,8,5;1,3,6,1;1,1,1,0;7,3,10,2})')).toBeCloseTo(88, 9)
    expect(ev('MDETERM({1,2,3;4,5,6})')).toEqual(E('#VALUE!'))
    const inv = arr('MINVERSE({4,-1;2,0})') as number[][]
    expect(inv[0][0]).toBeCloseTo(0, 12)
    expect(inv[0][1]).toBeCloseTo(0.5, 12)
    expect(inv[1][0]).toBeCloseTo(-1, 12)
    expect(inv[1][1]).toBeCloseTo(2, 12)
    expect(arr('MINVERSE({1,2;2,4})')).toEqual([[E('#NUM!')]])
    expect(arr('MUNIT(2)')).toEqual([
      [1, 0],
      [0, 1],
    ])
  })

  test('GROUPBY and PIVOTBY', () => {
    expect(arr('GROUPBY(D1:D4,E1:E4,SUM)')).toEqual([
      ['East', 17],
      ['North', 1],
      ['West', 5],
      ['Total', 23],
    ])
    expect(arr('GROUPBY(D1:D4,E1:E4,LAMBDA(x,MAX(x)),0,0,-2)')).toEqual([
      ['East', 10],
      ['West', 5],
      ['North', 1],
    ])
    expect(arr('GROUPBY(D1:D4,E1:E4,PERCENTOF,,0)')).toEqual([
      ['East', 17 / 23],
      ['North', 1 / 23],
      ['West', 5 / 23],
    ])
    expect(arr('GROUPBY(D1:D4,E1:E4,SUM,,0,,E1:E4>1)')).toEqual([
      ['East', 17],
      ['West', 5],
    ])
    expect(arr('PIVOTBY(D1:D4,B1:B4,E1:E4,SUM)')).toEqual([
      ['', 'a', 'b', 'c', 'Total'],
      ['East', '', 7, 10, 17],
      ['North', 1, '', '', 1],
      ['West', 5, '', '', 5],
      ['Total', 6, 7, 10, 23],
    ])
  })
})

describe('the operators @ and #', () => {
  test('@ takes the one value that is the formula\'s own', () => {
    expect(ev('@A1:A4', 2, 5)).toBe(2)
    expect(ev('@A1:A4*10', 0, 5)).toBe(30)
    expect(ev('@SEQUENCE(3)')).toBe(1)
    expect(arr('@A1:A4')).toEqual([[E('#VALUE!')]])
  })

  test('A1# is what the formula of A1 fills', () => {
    expect(ev('SUM(A1#)')).toBe(7)
    expect(ev('ROWS(A1#)')).toBe(4)
    expect(ev('SUM(A2#)')).toEqual(E('#REF!'))
    expect(arr('A1#*2')).toEqual([[6], [2], [4], [2]])
  })
})

describe('what spills', () => {
  const yes = [
    'A1:A3', 'A:A*2', 'SEQUENCE(3)', '{1,2}', 'IF(A1:A3>1,1,0)', 'LEN(A1:A3)', 'VLOOKUP(A1:A3,D:E,2,0)', 'SORT(A1:A3)&"x"', 'Name', 'A1#', 'LET(x,A1:A3,x)', 'LET(x,A1,y,B1:B3,x*y)',
    'LAMBDA(x,x+1)(A1:A3)', 'Unknown(1)', 'Table1[Qty]',
    'INDEX(A1:C3,0,1)', 'INDEX(A1:C3,2)', 'INDEX(A1:C3,,2)', 'OFFSET(A1,1,1,3)', 'OFFSET(A1:A3,1,0)', 'XLOOKUP(1,A1:A9,B1:C9)', 'XLOOKUP(A1:A3,D:D,E:E)', 'ROW(A1:A3)', 'INDIRECT("A1:A3")',
  ]
  const no = [
    'A1+B1', 'SUM(A1:A3)', 'VLOOKUP(A1,D:E,2,0)', 'LET(x,A1,x*2)', 'LET(f,LAMBDA(a,a+1),f(4))', 'LAMBDA(x,x+1)(5)', 'LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(5)', 'COUNTIF(A:A,B1)', '@A1:A3', 'IF(A1>1,B1,C1)', 'SUMPRODUCT(A1:A3,B1:B3)', 'TEXTJOIN(",",TRUE,A1:A3)', '{1}', 'MAX(LEN(A1:A3))',
    'INDEX(A1:C3,2,2)', 'INDEX(A1:A9,MATCH(1,B1:B9,0))', 'OFFSET(A1,1,1)', 'OFFSET(A1:A3,1,0,1,1)', 'XLOOKUP(1,A1:A9,B1:B9)', 'XLOOKUP(1,A1:A1,B1:B1)', 'ROW()', 'ROW(A5)', 'A1:A1',
  ]
  test('formulas that may give several values are dynamic arrays', () => {
    for (const f of yes) expect([f, spills(parseFormula(f))]).toEqual([f, true])
    for (const f of no) expect([f, spills(parseFormula(f))]).toEqual([f, false])
  })
})

describe('what spills, with the names of the workbook', () => {
  test('a defined name is what it stands for', () => {
    const names = (name: string) => ({ rate: parseFormula('0.5'), cell: parseFormula('S!$A$1'), list: parseFormula('S!$A$1:$A$9'), twice: parseFormula('LAMBDA(x,x*2)') })[name.toLowerCase()]
    expect(spills(parseFormula('Rate*2'), names)).toBe(false)
    expect(spills(parseFormula('Cell+1'), names)).toBe(false)
    expect(spills(parseFormula('List*2'), names)).toBe(true)
    expect(spills(parseFormula('SUM(List)'), names)).toBe(false)
    expect(spills(parseFormula('Twice(21)'), names)).toBe(false)
    expect(spills(parseFormula('Twice(A1:A3)'), names)).toBe(true)
    expect(spills(parseFormula('Nope+1'), names)).toBe(true)
  })
})

describe('as a file writes them', () => {
  test('prefixes of the new functions, of the names of LET / LAMBDA, SINGLE and ANCHORARRAY', () => {
    expect(withFunctionPrefixes('LET(x,1,x+1)')).toBe('_xlfn.LET(_xlpm.x,1,_xlpm.x+1)')
    expect(withFunctionPrefixes('=LAMBDA(a,[b],a+b)(1,2)')).toBe('=_xlfn.LAMBDA(_xlpm.a,[_xlpm.b],_xlpm.a+_xlpm.b)(1,2)')
    expect(withFunctionPrefixes('LET(f,LAMBDA(v,v*2),f(A1))')).toBe('_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2),_xlpm.f(A1))')
    expect(withFunctionPrefixes('@A1:A3*2')).toBe('_xlfn.SINGLE(A1:A3)*2')
    expect(withFunctionPrefixes('=@SUM(A1:A3)+@B1')).toBe('=_xlfn.SINGLE(SUM(A1:A3))+_xlfn.SINGLE(B1)')
    expect(withFunctionPrefixes('SUM(A1#)')).toBe('SUM(_xlfn.ANCHORARRAY(A1))')
    expect(withFunctionPrefixes('=Sheet2!B2#+1')).toBe('=_xlfn.ANCHORARRAY(Sheet2!B2)+1')
    expect(withFunctionPrefixes('FILTER(A1:A3,A1:A3>1)')).toBe('_xlfn._xlws.FILTER(A1:A3,A1:A3>1)')
    expect(withFunctionPrefixes('GROUPBY(A:A,B:B,SUM)')).toBe('_xlfn.GROUPBY(A:A,B:B,SUM)')
    expect(withFunctionPrefixes('Table1[@Qty]*2')).toBe('Table1[[#This Row],[Qty]]*2')
    expect(withFunctionPrefixes('SUM(Table1[@[Q1]:[Q4]])')).toBe('SUM(Table1[[#This Row],[Q1]:[Q4]])')
    expect(withFunctionPrefixes('Table1[@]')).toBe('Table1[#This Row]')
    expect(withFunctionPrefixes('[@Qty]*2', 'Sales')).toBe('Sales[[#This Row],[Qty]]*2')
    expect(withFunctionPrefixes('SUM([Qty])', 'Sales')).toBe('SUM(Sales[Qty])')
    expect(withFunctionPrefixes('SUM(Table1[Qty])', 'Sales')).toBe('SUM(Table1[Qty])')
    expect(withFunctionPrefixes('[@Qty]*2')).toBe('[@Qty]*2')
    expect(stripFunctionPrefixes('_xlfn.LET(_xlpm.x,1,_xlpm.x+1)')).toBe('LET(x,1,x+1)')
    expect(stripFunctionPrefixes('=_xlfn.LAMBDA(_xlpm.a,[_xlpm.b],_xlpm.a+_xlpm.b)(1,2)')).toBe('=LAMBDA(a,[b],a+b)(1,2)')
    expect(stripFunctionPrefixes('_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2),_xlpm.f(A1))')).toBe('LET(f,LAMBDA(v,v*2),f(A1))')
    expect(stripFunctionPrefixes('SUM(_xlfn.ANCHORARRAY(A1))')).toBe('SUM(A1#)')
    expect(stripFunctionPrefixes('=_xlfn.SINGLE(SUM(A1:A3))+_xlfn.SINGLE(B1)')).toBe('=@SUM(A1:A3)+@B1')
    expect(stripFunctionPrefixes('_xlfn.SINGLE(_xlfn.ANCHORARRAY(Sheet2!B2))')).toBe('@Sheet2!B2#')
    expect(stripFunctionPrefixes('_xlfn.SINGLE(A1:A3*2)')).toBe('SINGLE(A1:A3*2)')
    for (const f of ['=@SUM(A1:A3)+@B1', 'SUM(A1#)', '@Sheet2!B2#', 'LET(x,1,x+1)']) expect(stripFunctionPrefixes(withFunctionPrefixes(f))).toBe(f)
    // what was read from a file evaluates as it is
    expect(ev('_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2),_xlpm.f(A1))')).toBe(6)
    expect(ev('SUM(_xlfn.ANCHORARRAY(A1))')).toBe(7)
    expect(ev('_xlfn.SINGLE(A1:A4)', 1, 5)).toBe(1)
  })
})
