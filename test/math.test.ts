/** TeX math → MathML (and OMML): the constructs people write in Markdown, and what it does with what it cannot read */
import { describe, expect, test } from 'bun:test'
import { mathOmml, mathOf } from '../src/core/math/omml'
import { MAX_TEX, mathmlString, texMathml, texToMathml, type MNode } from '../src/core/math/tex'
import { ommlText } from '../src/core/docx/omml'
import { parseXml } from '../src/core/xml'

/** the drawing of a formula as markup (without the math / semantics wrapper and the annotation) */
const body = (tex: string, display = false) => {
  const r = texToMathml(tex, display)
  const sem = r.node.k![0]
  return mathmlString(sem.k![0])
}
const errorOf = (tex: string, display = false) => texToMathml(tex, display).error
/** the invisible function application after a function name */
const FA = '\u2061'
const NBSP = '\u00a0'

describe('TeX → MathML: tokens', () => {
  test('letters, numbers, operators', () => {
    expect(body('x+3.14-y')).toBe('<mrow><mi>x</mi><mo>+</mo><mn>3.14</mn><mo>−</mo><mi>y</mi></mrow>')
    expect(body('a*b/c')).toBe('<mrow><mi>a</mi><mo>∗</mo><mi>b</mi><mo>/</mo><mi>c</mi></mrow>')
    // spaces do not matter, comments are skipped
    expect(body('a   =\tb % a comment\n')).toBe('<mrow><mi>a</mi><mo>=</mo><mi>b</mi></mrow>')
    // brackets typed as themselves do not grow
    expect(body('(x)')).toBe('<mrow><mo stretchy="false">(</mo><mi>x</mi><mo stretchy="false">)</mo></mrow>')
  })

  test('Greek letters and symbols; capital Greek upright', () => {
    expect(body('\\alpha\\beta\\varepsilon\\varphi')).toBe('<mrow><mi>α</mi><mi>β</mi><mi>ε</mi><mi>φ</mi></mrow>')
    expect(body('\\Gamma')).toBe('<mi mathvariant="normal">Γ</mi>')
    expect(body('\\varGamma')).toBe('<mi>𝛤</mi>')
    expect(body('\\infty\\partial\\nabla\\hbar\\ell')).toBe('<mrow><mi>∞</mi><mi>∂</mi><mi>∇</mi><mi>ℏ</mi><mi>ℓ</mi></mrow>')
  })

  test('operators, relations, arrows', () => {
    expect(body('a \\le b \\ne c \\approx d \\equiv e')).toBe('<mrow><mi>a</mi><mo>≤</mo><mi>b</mi><mo>≠</mo><mi>c</mi><mo>≈</mo><mi>d</mi><mo>≡</mo><mi>e</mi></mrow>')
    expect(body('\\pm\\times\\cdot\\cup\\in\\subseteq')).toBe('<mrow><mo>±</mo><mo>×</mo><mo>⋅</mo><mo>∪</mo><mo>∈</mo><mo>⊆</mo></mrow>')
    expect(body('\\to\\Rightarrow\\iff\\mapsto\\longrightarrow')).toBe('<mrow><mo>→</mo><mo>⇒</mo><mo>⟺</mo><mo>↦</mo><mo>⟶</mo></mrow>')
    expect(body('\\ldots\\cdots\\vdots\\ddots')).toBe('<mrow><mo>…</mo><mo>⋯</mo><mo>⋮</mo><mo>⋱</mo></mrow>')
  })

  test('\\not and the negated relations', () => {
    expect(body('a \\not= b')).toBe('<mrow><mi>a</mi><mo>≠</mo><mi>b</mi></mrow>')
    expect(body('\\not\\in')).toBe('<mo>∉</mo>')
    expect(body('\\not\\subset')).toBe('<mo>⊄</mo>')
    // no negated form: a long solidus over it
    expect(body('\\not\\perp')).toBe('<mo>⊥\u0338</mo>')
  })

  test('escaped characters', () => {
    expect(body('\\{x\\}')).toBe('<mrow><mo stretchy="false">{</mo><mi>x</mi><mo stretchy="false">}</mo></mrow>')
    expect(body('100\\%')).toBe('<mrow><mn>100</mn><mi>%</mi></mrow>')
    expect(body('\\$5')).toBe('<mrow><mo stretchy="false">$</mo><mn>5</mn></mrow>')
  })
})

describe('TeX → MathML: scripts, primes, big operators', () => {
  test('sub / superscripts', () => {
    expect(body('x^2')).toBe('<msup><mi>x</mi><mn>2</mn></msup>')
    expect(body('x_i^2')).toBe('<msubsup><mi>x</mi><mi>i</mi><mn>2</mn></msubsup>')
    expect(body('x^{2n}_{i,j}')).toBe('<msubsup><mi>x</mi><mrow><mi>i</mi><mo>,</mo><mi>j</mi></mrow><mrow><mn>2</mn><mi>n</mi></mrow></msubsup>')
    expect(body('e^{i\\pi}')).toBe('<msup><mi>e</mi><mrow><mi>i</mi><mi>π</mi></mrow></msup>')
    // no base
    expect(body('^2')).toBe('<msup><mrow></mrow><mn>2</mn></msup>')
  })

  test('primes', () => {
    expect(body("f'")).toBe('<msup><mi>f</mi><mo>′</mo></msup>')
    expect(body("f''(x)")).toBe('<mrow><msup><mi>f</mi><mo>″</mo></msup><mo stretchy="false">(</mo><mi>x</mi><mo stretchy="false">)</mo></mrow>')
    expect(body("f'^2")).toBe('<msup><mi>f</mi><mrow><mo>′</mo><mn>2</mn></mrow></msup>')
  })

  test('big operators: limits under / over in display style, beside inline; integrals beside', () => {
    expect(body('\\sum_{i=1}^n', true)).toBe('<munderover><mo>∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover>')
    expect(body('\\sum_{i=1}^n')).toBe('<msubsup><mo>∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></msubsup>')
    expect(body('\\prod_k', true)).toBe('<munder><mo>∏</mo><mi>k</mi></munder>')
    expect(body('\\int_0^1', true)).toBe('<msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup>')
    expect(body('\\oint_C', true)).toBe('<msub><mo>∮</mo><mi>C</mi></msub>')
    // \limits / \nolimits
    expect(body('\\sum\\limits_{k}')).toBe('<munder><mo movablelimits="false">∑</mo><mi>k</mi></munder>')
    expect(body('\\sum\\nolimits_{k}', true)).toBe('<msub><mo>∑</mo><mi>k</mi></msub>')
    expect(body('\\int\\limits_a^b', true)).toBe('<munderover><mo movablelimits="false">∫</mo><mi>a</mi><mi>b</mi></munderover>')
    // inside a fraction the style is not display: limits beside
    expect(body('\\frac{\\sum_k a_k}{n}', true)).toContain('<msub><mo>∑</mo><mi>k</mi></msub>')
  })

  test('named functions: upright, the function application, limits for lim / max', () => {
    expect(body('\\sin x')).toBe(`<mrow><mrow><mi>sin</mi><mo rspace="0.1667em">${FA}</mo></mrow><mi>x</mi></mrow>`)
    // no space before a bracket
    expect(body('\\log(x)')).toBe(`<mrow><mrow><mi>log</mi><mo>${FA}</mo></mrow><mo stretchy="false">(</mo><mi>x</mi><mo stretchy="false">)</mo></mrow>`)
    expect(body('\\sin^2 x')).toBe(`<mrow><mrow><msup><mi>sin</mi><mn>2</mn></msup><mo rspace="0.1667em">${FA}</mo></mrow><mi>x</mi></mrow>`)
    expect(body('\\lim_{x\\to 0} f', true)).toBe(`<mrow><mrow><munder><mi>lim</mi><mrow><mi>x</mi><mo>→</mo><mn>0</mn></mrow></munder><mo rspace="0.1667em">${FA}</mo></mrow><mi>f</mi></mrow>`)
    expect(body('\\lim_{n} a')).toContain('<msub><mi>lim</mi><mi>n</mi></msub>')
    expect(body('\\operatorname{argmax}_x f', true)).toContain('<msub><mi>argmax</mi><mi>x</mi></msub>')
    expect(body('\\operatorname*{argmax}_x f', true)).toContain('<munder><mi>argmax</mi><mi>x</mi></munder>')
    expect(body('\\limsup')).toContain('<mi>lim sup</mi>')
  })

  test('double scripts are an error', () => {
    expect(errorOf('x^2^3')).toBe('Double superscript')
    expect(errorOf('x_1_2')).toBe('Double subscript')
  })
})

describe('TeX → MathML: fractions, roots, delimiters, accents', () => {
  test('fractions and binomials', () => {
    expect(body('\\frac{a}{b}')).toBe('<mfrac><mi>a</mi><mi>b</mi></mfrac>')
    expect(body('\\frac12')).toBe('<mfrac><mn>1</mn><mn>2</mn></mfrac>')
    expect(body('\\dfrac{a}{b}')).toBe('<mstyle displaystyle="true" scriptlevel="0"><mfrac><mi>a</mi><mi>b</mi></mfrac></mstyle>')
    expect(body('\\tfrac{a}{b}')).toBe('<mstyle displaystyle="false" scriptlevel="0"><mfrac><mi>a</mi><mi>b</mi></mfrac></mstyle>')
    expect(body('\\binom{n}{k}')).toBe('<mrow><mo>(</mo><mfrac linethickness="0"><mi>n</mi><mi>k</mi></mfrac><mo>)</mo></mrow>')
    expect(body('{a \\over b}')).toBe('<mfrac><mi>a</mi><mi>b</mi></mfrac>')
    expect(body('n \\choose k')).toBe('<mrow><mo>(</mo><mfrac linethickness="0"><mi>n</mi><mi>k</mi></mfrac><mo>)</mo></mrow>')
  })

  test('roots', () => {
    expect(body('\\sqrt{x}')).toBe('<msqrt><mi>x</mi></msqrt>')
    expect(body('\\sqrt[3]{x+1}')).toBe('<mroot><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><mn>3</mn></mroot>')
    expect(body('\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}')).toBe(
      '<mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac>',
    )
  })

  test('\\left … \\right, \\middle and sized delimiters', () => {
    const f = '<mo fence="true" stretchy="true" symmetric="true">'
    expect(body('\\left( \\frac{a}{b} \\right)')).toBe(`<mrow>${f}(</mo><mfrac><mi>a</mi><mi>b</mi></mfrac>${f})</mo></mrow>`)
    expect(body('\\left\\{ x \\right.')).toBe(`<mrow>${f}{</mo><mi>x</mi></mrow>`)
    expect(body('\\left\\langle x \\middle| y \\right\\rangle')).toBe(`<mrow>${f}⟨</mo><mi>x</mi><mo stretchy="true" symmetric="true">|</mo><mi>y</mi>${f}⟩</mo></mrow>`)
    expect(body('\\left\\lfloor x \\right\\rfloor')).toContain('⌊')
    expect(body('\\big( \\Bigg]')).toBe('<mrow><mo minsize="1.2em" maxsize="1.2em" stretchy="true" symmetric="true">(</mo><mo minsize="2.470em" maxsize="2.470em" stretchy="true" symmetric="true">]</mo></mrow>')
    expect(body('\\bigl| x \\bigr|')).toContain('minsize="1.2em"')
    expect(errorOf('\\left( x')).toBe('Missing \\right')
    expect(errorOf('x \\right)')).toBe('\\right without \\left')
    expect(errorOf('\\left x \\right)')).toBe('Not a delimiter: x')
  })

  test('accents, over / under lines, braces', () => {
    expect(body('\\hat{x}')).toBe('<mover accent="true"><mi>x</mi><mo stretchy="false">^</mo></mover>')
    expect(body('\\widehat{xy}')).toBe('<mover accent="true"><mrow><mi>x</mi><mi>y</mi></mrow><mo stretchy="true">^</mo></mover>')
    expect(body('\\vec v')).toBe('<mover accent="true"><mi>v</mi><mo stretchy="false">→</mo></mover>')
    expect(body('\\bar{x}\\tilde{y}\\dot{z}\\ddot{w}')).toBe(
      '<mrow><mover accent="true"><mi>x</mi><mo stretchy="false">¯</mo></mover><mover accent="true"><mi>y</mi><mo stretchy="false">~</mo></mover><mover accent="true"><mi>z</mi><mo stretchy="false">˙</mo></mover><mover accent="true"><mi>w</mi><mo stretchy="false">¨</mo></mover></mrow>',
    )
    expect(body('\\overline{AB}')).toBe('<mover accent="true"><mrow><mi>A</mi><mi>B</mi></mrow><mo stretchy="true">‾</mo></mover>')
    expect(body('\\underline{x}')).toBe('<munder><mi>x</mi><mo stretchy="true">_</mo></munder>')
    // a brace takes its label under / over it, in any style
    expect(body('\\overbrace{a+b}^{n}')).toBe('<mover><mover><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mo stretchy="true">⏞</mo></mover><mi>n</mi></mover>')
    expect(body('\\underbrace{x}_{k}')).toBe('<munder><munder><mi>x</mi><mo stretchy="true">⏟</mo></munder><mi>k</mi></munder>')
    expect(body('\\overset{!}{=}')).toBe('<mover><mo>=</mo><mo>!</mo></mover>')
    expect(body('\\underset{x}{\\min}')).toBe(`<munder><mrow><mi>min</mi><mo>${FA}</mo></mrow><mi>x</mi></munder>`)
    expect(body('\\xrightarrow{f}')).toBe('<mover><mo stretchy="true" minsize="2em">→</mo><mi>f</mi></mover>')
  })
})

describe('TeX → MathML: text, fonts, spaces, color, environments', () => {
  test('\\text: words, spaces at its ends kept, math inside', () => {
    expect(body('\\text{if } x')).toBe(`<mrow><mtext>if${NBSP}</mtext><mi>x</mi></mrow>`)
    expect(body('\\text{a $x^2$ b}')).toBe(`<mrow><mtext>a${NBSP}</mtext><msup><mi>x</mi><mn>2</mn></msup><mtext>${NBSP}b</mtext></mrow>`)
    expect(body('\\textbf{bold}')).toBe('<mtext mathvariant="bold">bold</mtext>')
    expect(body('\\text{50\\% off}')).toBe('<mtext>50% off</mtext>')
  })

  test('font commands: the letters of the math alphabets', () => {
    expect(body('\\mathbb{R}')).toBe('<mi>ℝ</mi>')
    expect(body('\\mathbb{N}\\mathbb{Z}\\mathbb{Q}\\mathbb{C}')).toBe('<mrow><mi>ℕ</mi><mi>ℤ</mi><mi>ℚ</mi><mi>ℂ</mi></mrow>')
    expect(body('\\mathbf{x}')).toBe('<mi>𝐱</mi>')
    expect(body('\\mathbf{v}_1')).toBe('<msub><mi>𝐯</mi><mn>1</mn></msub>')
    expect(body('\\mathcal{L}\\mathcal{A}')).toBe('<mrow><mi>ℒ</mi><mi>𝒜</mi></mrow>')
    expect(body('\\mathfrak{g}')).toBe('<mi>𝔤</mi>')
    expect(body('\\mathsf{A}\\mathtt{a}\\mathit{x}')).toBe('<mrow><mi>𝖠</mi><mi>𝚊</mi><mi>𝑥</mi></mrow>')
    expect(body('\\boldsymbol{\\alpha}\\bm{x}')).toBe('<mrow><mi>𝜶</mi><mi>𝒙</mi></mrow>')
    expect(body('\\mathbf{\\Gamma}')).toBe('<mi>𝚪</mi>')
    expect(body('\\mathbb{1}')).toBe('<mn>𝟙</mn>')
    expect(body('\\mathrm{d}x')).toBe('<mrow><mi mathvariant="normal">d</mi><mi>x</mi></mrow>')
    expect(body('\\mathrm{max}')).toBe('<mi>max</mi>')
    expect(body('{\\rm d}x')).toBe('<mrow><mi mathvariant="normal">d</mi><mi>x</mi></mrow>')
  })

  test('spaces', () => {
    expect(body('a\\,b')).toBe('<mrow><mi>a</mi><mspace width="0.1667em"/><mi>b</mi></mrow>')
    expect(body('a\\quad b\\qquad c')).toBe('<mrow><mi>a</mi><mspace width="1em"/><mi>b</mi><mspace width="2em"/><mi>c</mi></mrow>')
    expect(body('a\\!b\\;c\\:d\\ e')).toContain('<mspace width="-0.1667em"/>')
    expect(body('\\hspace{3mu}')).toBe('<mspace width="0.1667em"/>')
    expect(body('\\hspace{1cm}')).toBe('<mspace width="2.85em"/>')
    expect(body('a~b')).toBe(`<mrow><mi>a</mi><mtext>${NBSP}</mtext><mi>b</mi></mrow>`)
  })

  test('color, when it is a color', () => {
    expect(body('\\color{red} x y')).toBe('<mstyle mathcolor="red"><mi>x</mi><mi>y</mi></mstyle>')
    expect(body('{\\color{#0a0} x} y')).toBe('<mrow><mstyle mathcolor="#0a0"><mi>x</mi></mstyle><mi>y</mi></mrow>')
    expect(body('\\textcolor{blue}{x}')).toBe('<mstyle mathcolor="blue"><mi>x</mi></mstyle>')
    expect(body('\\color[HTML]{FF8800} x')).toBe('<mstyle mathcolor="#ff8800"><mi>x</mi></mstyle>')
    // not a color: ignored, nothing of it reaches an attribute
    expect(body('\\color{red;background:url(x)} x')).toBe('<mi>x</mi>')
    expect(body('\\textcolor{expression(1)}{x}')).toBe('<mi>x</mi>')
  })

  test('styles, phantoms, boxes, modulo', () => {
    expect(body('\\displaystyle \\sum_k')).toBe('<mstyle displaystyle="true" scriptlevel="0"><munder><mo>∑</mo><mi>k</mi></munder></mstyle>')
    expect(body('\\phantom{x}')).toBe('<mphantom><mi>x</mi></mphantom>')
    expect(body('\\boxed{x}')).toBe('<menclose notation="box"><mi>x</mi></menclose>')
    expect(body('a \\bmod b')).toContain('<mi>mod</mi>')
    expect(body('a \\equiv b \\pmod{n}')).toContain('<mo>(</mo><mi>mod</mi><mspace width="0.3333em"/><mi>n</mi><mo>)</mo>')
  })

  test('matrices', () => {
    expect(body('\\begin{matrix} a & b \\\\ c & d \\end{matrix}')).toBe('<mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable>')
    const f = (c: string) => `<mo fence="true" stretchy="true" symmetric="true">${c}</mo>`
    const m = '<mtable><mtr><mtd><mn>1</mn></mtd><mtd><mn>0</mn></mtd></mtr><mtr><mtd><mn>0</mn></mtd><mtd><mn>1</mn></mtd></mtr></mtable>'
    for (const [env, open, close] of [['pmatrix', '(', ')'], ['bmatrix', '[', ']'], ['Bmatrix', '{', '}'], ['vmatrix', '|', '|'], ['Vmatrix', '‖', '‖']])
      expect(body(`\\begin{${env}} 1 & 0 \\\\ 0 & 1 \\end{${env}}`)).toBe(`<mrow>${f(open)}${m}${f(close)}</mrow>`)
    // a last \\ does not make an empty row; a short row is filled
    expect(body('\\begin{matrix} a & b \\\\ c \\\\ \\end{matrix}')).toBe('<mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mrow></mrow></mtd></mtr></mtable>')
  })

  test('cases, aligned, gathered, array', () => {
    expect(body('\\begin{cases} 1 & x > 0 \\\\ 0 & \\text{else} \\end{cases}')).toBe(
      '<mrow><mo fence="true" stretchy="true" symmetric="true">{</mo><mtable columnalign="left left" columnspacing="1em">' +
        '<mtr><mtd columnalign="left"><mn>1</mn></mtd><mtd columnalign="left"><mrow><mi>x</mi><mo>&gt;</mo><mn>0</mn></mrow></mtd></mtr>' +
        '<mtr><mtd columnalign="left"><mn>0</mn></mtd><mtd columnalign="left"><mtext>else</mtext></mtd></mtr></mtable></mrow>',
    )
    // aligned: right / left pairs, the relation after & spaced as a relation, display style inside
    expect(body('\\begin{aligned} a &= b \\\\ &= c \\end{aligned}')).toBe(
      '<mtable displaystyle="true" columnalign="right left" columnspacing="0em 2em">' +
        '<mtr><mtd columnalign="right"><mi>a</mi></mtd><mtd columnalign="left"><mrow><mi></mi><mo>=</mo><mi>b</mi></mrow></mtd></mtr>' +
        '<mtr><mtd columnalign="right"><mrow></mrow></mtd><mtd columnalign="left"><mrow><mi></mi><mo>=</mo><mi>c</mi></mrow></mtd></mtr></mtable>',
    )
    expect(body('\\begin{align*} x &= 1 \\end{align*}')).toContain('columnalign="right left"')
    expect(body('\\begin{gathered} a \\\\ b \\end{gathered}')).toBe('<mtable displaystyle="true"><mtr><mtd><mi>a</mi></mtd></mtr><mtr><mtd><mi>b</mi></mtd></mtr></mtable>')
    expect(body('\\begin{array}{l|cr} a & b & c \\end{array}')).toBe(
      '<mtable columnalign="left center right"><mtr><mtd columnalign="left"><mi>a</mi></mtd><mtd><mi>b</mi></mtd><mtd columnalign="right"><mi>c</mi></mtd></mtr></mtable>',
    )
    // \\ at the top of a display formula: lines, centered
    expect(body('a \\\\ b', true)).toBe('<mtable displaystyle="true"><mtr><mtd><mi>a</mi></mtd></mtr><mtr><mtd><mi>b</mi></mtd></mtr></mtable>')
    expect(body('\\sum_{\\substack{i<n \\\\ j}} x', true)).toContain('<mtable scriptlevel="1" rowspacing="0.1em">')
    expect(errorOf('\\begin{foo} x \\end{foo}')).toBe('Unknown environment foo')
    expect(errorOf('\\begin{matrix} x \\end{pmatrix}')).toBe('\\begin{matrix} ended by \\end{pmatrix}')
    expect(errorOf('\\begin{matrix} x')).toBe('Missing \\end')
  })

  test('\\tag shows its label at the right', () => {
    expect(body('E = mc^2 \\tag{1}', true)).toContain('<mspace width="2em"/><mtext>(1)</mtext>')
  })
})

describe('TeX → MathML: the math element, errors, bounds', () => {
  test('a math element: block for display, the TeX as written in an annotation', () => {
    const tex = '\n  \\frac{a}{b} < c\n'
    expect(texMathml(tex, true, false)).toBe(
      '<math display="block"><semantics><mrow><mfrac><mi>a</mi><mi>b</mi></mfrac><mo>&lt;</mo><mi>c</mi></mrow>' +
        '<annotation encoding="application/x-tex">\n  \\frac{a}{b} &lt; c\n</annotation></semantics></math>',
    )
  })

  test('a command it does not know: shown as written, the rest drawn, said in the error', () => {
    const r = texToMathml('\\foo{x} + \\bar{y} + \\baz', false)
    expect(r.error).toBe('Not supported: \\foo, \\baz')
    const s = mathmlString(r.node)
    expect(s).toContain('<merror><mtext>\\foo</mtext></merror><mi>x</mi><mo>+</mo>')
    expect(s).toContain('<mover accent="true"><mi>y</mi>')
    // a command name from the file is never a lookup of an object's own properties
    expect(texToMathml('\\constructor \\toString \\hasOwnProperty \\valueOf', false).error).toBe('Not supported: \\constructor, \\toString, \\hasOwnProperty, \\valueOf')
  })

  test('input it cannot read: the whole source in an merror, the reason given', () => {
    for (const [tex, why] of [
      ['\\frac{a}{', 'Missing }'],
      ['{x', 'Missing }'],
      ['x}', 'Unexpected }'],
      ['\\frac{a}', 'Missing argument'],
      ['a & b \\end{x}', '\\end without \\begin'],
      ['\\sqrt[3', 'Missing ]'],
      ['x^', 'Missing argument'],
      ['\\', 'A backslash ends the text'],
      ['#1', 'Unexpected #'],
      ['\\text{$x}', 'Unclosed $ in \\text'],
      ['\\hspace{wide}', 'Not a length: wide'],
    ] as [string, string][]) {
      const r = texToMathml(tex, false)
      expect([tex, r.error]).toEqual([tex, why])
      expect(mathmlString(r.node)).toBe(`<math><semantics><merror><mtext>${tex.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</mtext></merror><annotation encoding="application/x-tex">${tex.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</annotation></semantics></math>`)
    }
  })

  test('bounded: nesting depth, length, elements', () => {
    expect(texToMathml('{'.repeat(200) + 'x' + '}'.repeat(200), false).error).toBe('Nested too deeply')
    expect(texToMathml('\\frac{'.repeat(100) + 'x' + '}{y}'.repeat(100), false).error).toBe('Nested too deeply')
    expect(texToMathml('\\sqrt{'.repeat(100) + 'x' + '}'.repeat(100), false).error).toBe('Nested too deeply')
    expect(texToMathml('x'.repeat(MAX_TEX + 1), false).error).toBe('Longer than 20,000 characters')
    const t0 = performance.now()
    expect(texToMathml('x^{y}_{z}'.repeat(2200), false).error).toBe(undefined)
    expect(performance.now() - t0).toBeLessThan(2000)
  })

  test('never throws: broken and random input', () => {
    const broken = [
      '', ' ', '}', '{', '{}', '^', '_', "'", '&', '\\\\', '\\', '\\left', '\\right', '\\left(', '\\right)', '\\middle|', '\\begin', '\\begin{', '\\begin{}', '\\end{matrix}',
      '\\frac', '\\frac{', '\\frac{}{}', '\\sqrt[', '\\sqrt[]{}', '\\sqrt{}', '\\text', '\\text{', '\\color', '\\color{', '\\color{}{}', '\\not', '\\not}', '\\big', '\\big}', '\\left.\\right.',
      'x^{', 'x_', 'x^^', "x''^'", '\\over', 'a\\over', '\\over b', '\\choose', '\\operatorname', '\\operatorname*{}', '\\mathbb', '\\mathbb{', '\\hspace', '\\hspace{', '\\kern',
      '\\begin{array}', '\\begin{array}{', '\\begin{array}{zz} a \\end{array}', '\\begin{cases}', '\\substack', '\\substack{', '\\tag', '\\tag{', '\\label', '\\xrightarrow', '\\xrightarrow[',
      '$', '$$', '#', '%', '~', '\\%', '\\$', '\\#', '\\&', '\\_', '\\{', '\\}', '\\|', '\\,', '\\ ', '\\\n', '\u0000', '\uffff', '\ud800', '😀^😀', 'α_β', '\\left\\{\\middle\\|\\right\\}',
      '\\begin{matrix}&&&\\\\&&\\end{matrix}', '\\frac{\\frac{\\frac{}{}}{}}{}', '\\sqrt[\\sqrt[\\sqrt]]', '\\text{\\text{\\text{$\\text{$x$}$}}}', '\\color{red}', '{\\displaystyle}', '\\rm',
    ]
    for (const tex of broken) {
      for (const display of [false, true]) {
        const r = texToMathml(tex, display)
        expect(r.node.t).toBe('math')
        expect(() => mathOmml({ tex, delim: display ? '$$' : '$' })).not.toThrow()
      }
    }
    // random strings of TeX's own characters and commands
    const parts = ['\\frac', '\\sqrt', '{', '}', '^', '_', '&', '\\\\', '\\left(', '\\right)', '\\begin{matrix}', '\\end{matrix}', 'x', '1', '+', '\\alpha', '\\text{a}', "'", '[', ']', '\\', ' ', '\\over', '\\not', '\\color{red}', '\\sum', '\\lim', '$']
    let seed = 7
    const rnd = (n: number) => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff), seed % n)
    for (let k = 0; k < 3000; k++) {
      let tex = ''
      for (let j = rnd(16); j >= 0; j--) tex += parts[rnd(parts.length)]
      const r = texToMathml(tex, rnd(2) === 1)
      expect(r.node.t).toBe('math')
      expect(textOfAnnotation(r.node)).toBe(tex)
    }
  })
})

const textOfAnnotation = (n: MNode) => n.k![0].k![1].x

describe('TeX → OMML (the Word file a Markdown document is)', () => {
  const omml = (tex: string, display = false) => mathOmml({ tex, delim: display ? '$$' : '$' })
  /** the OMML without the run properties, for reading */
  const bare = (x: string) => x.replace(/<w:rPr>.*?<\/w:rPr>/g, '').replace(/ xml:space="preserve"/g, '')

  test('inline: m:oMath; display: m:oMathPara; the TeX in attributes, line breaks kept', () => {
    expect(bare(omml('x^2'))).toBe('<m:oMath oe:tex="x^2" oe:delim="$"><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>')
    const x = omml('\n a < b & "c"\n', true)
    expect(x).toStartWith('<m:oMathPara oe:tex="&#10; a &lt; b &amp; &quot;c&quot;&#10;" oe:delim="$$"><m:oMath>')
    const el = parseXml(`<w:p xmlns:w="w" xmlns:m="m" xmlns:oe="oe">${x}</w:p>`).c[0]
    expect(mathOf(el.a)).toEqual({ tex: '\n a < b & "c"\n', delim: '$$' })
  })

  test('structures: fraction, radical, n-ary with its operand, delimiters, matrix, accents, functions', () => {
    expect(bare(omml('\\frac{a}{b}'))).toContain('<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>')
    expect(bare(omml('\\binom{n}{k}'))).toContain('<m:f><m:fPr><m:type m:val="noBar"/></m:fPr>')
    expect(bare(omml('\\sqrt{x}'))).toContain('<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>')
    expect(bare(omml('\\sqrt[3]{x}'))).toContain('<m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e>')
    // the sum takes what follows it, up to the relation
    expect(bare(omml('\\sum_{i=1}^{n} i^2 = s', true))).toContain(
      '<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr><m:sub><m:r><m:t>i</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:sSup><m:e><m:r><m:t>i</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:e></m:nary><m:r><m:t>=</m:t></m:r>',
    )
    expect(bare(omml('\\int_0^1 x\\,dx'))).toContain('<m:limLoc m:val="subSup"/>')
    expect(bare(omml('\\int f'))).toContain('<m:subHide m:val="1"/><m:supHide m:val="1"/>')
    expect(bare(omml('\\left( x \\middle| y \\right]'))).toContain('<m:d><m:dPr><m:begChr m:val="("/><m:sepChr m:val="|"/><m:endChr m:val="]"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e><m:e><m:r><m:t>y</m:t></m:r></m:e></m:d>')
    expect(bare(omml('\\begin{pmatrix} a & b \\end{pmatrix}'))).toContain('<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="2"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr><m:e>')
    expect(bare(omml('\\begin{aligned} a &= b \\end{aligned}'))).toContain('<m:mcJc m:val="right"/></m:mcPr></m:mc><m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="left"/>')
    expect(bare(omml('\\hat{x}'))).toContain('<m:acc><m:accPr><m:chr m:val="\u0302"/></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>')
    expect(bare(omml('\\overline{x}'))).toContain('<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>')
    expect(bare(omml('\\overbrace{x}^{n}'))).toContain('<m:limUpp><m:e><m:groupChr><m:groupChrPr><m:chr m:val="⏞"/>')
    expect(bare(omml('\\sin x'))).toContain('<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>sin</m:t></m:r>')
    expect(bare(omml('\\text{if}'))).toContain('<m:r><m:rPr><m:nor/></m:rPr><m:t>if</m:t></m:r>')
    expect(omml('\\color{red}{x}')).toContain('<w:color w:val="FF0000"/>')
  })

  test('the linear text the document shows for search and plain text', () => {
    const text = (tex: string, display = false) => ommlText(parseXml(`<w:p xmlns:w="w" xmlns:m="m" xmlns:oe="oe">${omml(tex, display)}</w:p>`).c[0])
    expect(text('e^{i\\pi} + 1 = 0')).toBe('e^iπ+1=0')
    expect(text('\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}', true)).toBe('(−b±√(b²−4ac))/2a')
    expect(text('\\sum_{i=1}^n i')).toBe('∑_(i=1)^n i')
  })

  test('a formula that cannot be read is its source, as text', () => {
    expect(bare(omml('\\frac{a}{'))).toBe('<m:oMath oe:tex="\\frac{a}{" oe:delim="$"><m:r><m:rPr><m:nor/></m:rPr><m:t>\\frac{a}{</m:t></m:r></m:oMath>')
  })
})
