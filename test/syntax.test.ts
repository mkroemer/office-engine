import { describe, expect, test } from 'bun:test'
import { highlight, highlightLines, languageOf, languageOfFile, type Token } from '../src/core/syntax'

/** the class of the token that holds `piece` (its `nth` occurrence in the text); the piece must lie in one token */
function cls(lang: string, text: string, piece: string, nth = 0): string {
  const toks = highlight(lang, text)
  expect(toks.map((t) => t[1]).join('')).toBe(text)
  let at = -1
  for (let k = 0; k <= nth; k++) at = text.indexOf(piece, at + 1)
  expect(at).toBeGreaterThanOrEqual(0)
  let o = 0
  for (const [c, t] of toks) {
    if (at >= o && at < o + t.length) {
      expect(at + piece.length).toBeLessThanOrEqual(o + t.length)
      return c
    }
    o += t.length
  }
  throw new Error(`${piece} not found`)
}

const joined = (toks: Token[]) => toks.map((t) => t[1]).join('')

describe('syntax: languages and names', () => {
  test('aliases and file names', () => {
    expect(languageOf('js')).toBe('javascript')
    expect(languageOf('JavaScript')).toBe('javascript')
    expect(languageOf(' TS ')).toBe('typescript')
    expect(languageOf('py')).toBe('python')
    expect(languageOf('sh')).toBe('shell')
    expect(languageOf('bash')).toBe('shell')
    expect(languageOf('yml')).toBe('yaml')
    expect(languageOf('c++')).toBe('cpp')
    expect(languageOf('C#')).toBe('csharp')
    expect(languageOf('rs')).toBe('rust')
    expect(languageOf('golang')).toBe('go')
    expect(languageOf('tsx')).toBe('tsx')
    expect(languageOf('jsx')).toBe('javascript')
    expect(languageOf('svg')).toBe('xml')
    expect(languageOf('cobol')).toBeUndefined()
    expect(languageOf('constructor')).toBeUndefined()
    expect(languageOf('')).toBeUndefined()
    expect(languageOfFile('main.ts')).toBe('typescript')
    expect(languageOfFile('script.py')).toBe('python')
    expect(languageOfFile('config.json')).toBe('json')
    expect(languageOfFile('ci.YML')).toBe('yaml')
    expect(languageOfFile('site.css')).toBe('css')
    expect(languageOfFile('lib.rs')).toBe('rust')
    expect(languageOfFile('notes.txt')).toBeUndefined()
    expect(languageOfFile('README.md')).toBeUndefined()
    expect(languageOfFile(undefined, 'application/json; charset=utf-8')).toBe('json')
    expect(languageOfFile('blob', 'text/x-python')).toBe('python')
  })

  test('an unknown language is plain text', () => {
    expect(highlight('cobol', 'MOVE A TO B.\nSTOP RUN.')).toEqual([['', 'MOVE A TO B.'], ['', '\n'], ['', 'STOP RUN.']])
    expect(highlight(undefined, 'x = 1')).toEqual([['', 'x = 1']])
    expect(highlight('text', 'if (a) "b"')).toEqual([['', 'if (a) "b"']])
  })
})

describe('syntax: JavaScript / TypeScript / JSX', () => {
  const JS = 'const greeting = "Hello";\nconsole.log(greeting) // say it\nfunction add(a, b) { return a + 0x1F + 1_000n }'
  test('keywords, strings, comments, numbers, functions', () => {
    expect(cls('javascript', JS, 'const')).toBe('k')
    expect(cls('javascript', JS, '"Hello"')).toBe('s')
    expect(cls('javascript', JS, '// say it')).toBe('c')
    expect(cls('javascript', JS, 'log')).toBe('f')
    expect(cls('javascript', JS, 'add')).toBe('f')
    expect(cls('javascript', JS, 'return')).toBe('k')
    expect(cls('javascript', JS, '0x1F')).toBe('n')
    expect(cls('javascript', JS, '1_000n')).toBe('n')
    expect(cls('javascript', JS, 'greeting', 1)).toBe('')
    // a member named like a keyword is not one
    expect(cls('javascript', 'promise.catch(e)', 'catch')).toBe('f')
    expect(cls('javascript', 'x.default', 'default')).toBe('')
  })

  test('template literals, nested, across lines', () => {
    const t = 'const s = `a ${b ? `in ${c + "}"}` : d} z`\nconst u = `line one\nline ${two}\nend` + 1'
    expect(cls('javascript', t, '`a ')).toBe('s')
    expect(cls('javascript', t, '${')).toBe('o')
    expect(cls('javascript', t, '`in ')).toBe('s')
    expect(cls('javascript', t, '"}"')).toBe('s')
    expect(cls('javascript', t, ' z`')).toBe('s')
    expect(cls('javascript', t, 'line one')).toBe('s')
    expect(cls('javascript', t, 'two')).toBe('')
    expect(cls('javascript', t, 'end`')).toBe('s')
    expect(cls('javascript', t, '1', 0)).toBe('n')
  })

  test('regular expressions and division', () => {
    const t = 'const r = /a[/]b\\/c/gi; const q = x / y / 2; s.replace(/x/g, "")'
    expect(cls('javascript', t, '/a[/]b\\/c/gi')).toBe('s')
    expect(cls('javascript', t, 'x / y')).toBe('')
    expect(cls('javascript', t, '/x/g')).toBe('s')
  })

  test('unterminated strings and comments', () => {
    const t = 'let a = "open\nlet b = 2 /* open\nstill comment\n*/ let c'
    expect(cls('javascript', t, '"open')).toBe('s')
    // a "…" string ends with its line
    expect(cls('javascript', t, 'let', 1)).toBe('k')
    expect(cls('javascript', t, 'still comment')).toBe('c')
    expect(cls('javascript', t, 'let', 2)).toBe('k')
  })

  test('JSX elements, attributes and expressions; TS generics are not elements', () => {
    const t = 'return (\n  <div className="box" onClick={() => go(1)}>\n    Hi {name} <b>there</b>\n    <Item key={k} />\n  </div>\n)\nconst x = a < b'
    expect(cls('javascript', t, '<div')).toBe('t')
    expect(cls('javascript', t, 'className')).toBe('a')
    expect(cls('javascript', t, '"box"')).toBe('s')
    expect(cls('javascript', t, 'go')).toBe('f')
    expect(cls('javascript', t, 'Hi ')).toBe('')
    expect(cls('javascript', t, 'there')).toBe('')
    expect(cls('javascript', t, 'Item')).toBe('t')
    expect(cls('javascript', t, '</')).toBe('t')
    expect(cls('javascript', t, 'const')).toBe('k')
    expect(cls('javascript', t, ' < ')).toBe('')
    const ts = 'const m = new Map<string, number>()\nfunction id<T>(x: T): T { return x }\nconst f = <T,>(v: T) => v\ninterface Point { x: number }'
    expect(cls('typescript', ts, 'string')).toBe('y')
    expect(cls('typescript', ts, 'number')).toBe('y')
    expect(cls('typescript', ts, 'Point')).toBe('y')
    expect(cls('typescript', ts, 'interface')).toBe('k')
    // after <T,> the code goes on as code, not as an element's children
    expect(cls('tsx', ts, 'interface')).toBe('k')
    expect(cls('tsx', ts, 'number', 1)).toBe('y')
    expect(cls('tsx', ts + '\nconst el = <App />', '<')).toBe('')
    expect(cls('tsx', ts + '\nconst el = <App />', 'App')).toBe('t')
  })

  test('decorators', () => {
    expect(cls('typescript', '@Component({ selector: "x" })\nclass A {}', '@Component')).toBe('m')
  })
})

describe('syntax: other languages', () => {
  test('JSON: keys, strings, numbers, constants, comments', () => {
    const t = '{\n  "name": "demo", // c\n  "n": -1.5e3,\n  "ok": true,\n  "list": [null, "a"]\n}'
    expect(cls('json', t, '"name"')).toBe('a')
    expect(cls('json', t, '"demo"')).toBe('s')
    expect(cls('json', t, '-1.5e3')).toBe('n')
    expect(cls('json', t, 'true')).toBe('n')
    expect(cls('json', t, 'null')).toBe('n')
    expect(cls('json', t, '"a"')).toBe('s')
    expect(cls('json', t, '// c')).toBe('c')
  })

  test('Python: def, class, decorators, strings with prefixes, triple quotes', () => {
    const t = '@app.route("/")\ndef home(req):\n    """Doc\n    more"""\n    x = f"{req}" + r\'\\d\'  # note\n    return None\nclass Shop(Base):\n    pass'
    expect(cls('python', t, '@app.route')).toBe('m')
    expect(cls('python', t, 'def')).toBe('k')
    expect(cls('python', t, 'home')).toBe('f')
    expect(cls('python', t, '"""Doc')).toBe('s')
    expect(cls('python', t, 'more"""')).toBe('s')
    expect(cls('python', t, 'f"{req}"')).toBe('s')
    expect(cls('python', t, "r'\\d'")).toBe('s')
    expect(cls('python', t, '# note')).toBe('c')
    expect(cls('python', t, 'None')).toBe('n')
    expect(cls('python', t, 'Shop')).toBe('y')
    // @ as an operator is not a decorator
    expect(cls('python', 'c = a @ b', '@')).toBe('')
  })

  test('CSS and SCSS: selectors, properties, values, at-rules, comments', () => {
    const t = '/* theme\n   colors */\n@media (max-width: 600px) {\n  .card > a:hover, #main {\n    color: #0b57d0 !important;\n    margin: -4px 1.5em;\n    background: url(x.png) var(--bg);\n  }\n}'
    expect(cls('css', t, 'colors */')).toBe('c')
    expect(cls('css', t, '@media')).toBe('k')
    expect(cls('css', t, 'max-width')).toBe('a')
    expect(cls('css', t, '600px')).toBe('n')
    expect(cls('css', t, '.card')).toBe('a')
    expect(cls('css', t, 'a', 3)).toBe('t')
    expect(cls('css', t, ':hover')).toBe('k')
    expect(cls('css', t, '#main')).toBe('a')
    expect(cls('css', t, 'color', 1)).toBe('a')
    expect(cls('css', t, '#0b57d0')).toBe('n')
    expect(cls('css', t, '!important')).toBe('k')
    expect(cls('css', t, '-4px')).toBe('n')
    expect(cls('css', t, 'url')).toBe('f')
    expect(cls('css', t, '--bg')).toBe('a')
    const s = '$pad: 4px;\n.nav {\n  // nested\n  &:hover { padding: $pad; }\n  li { color: red }\n}'
    expect(cls('scss', s, '$pad')).toBe('a')
    expect(cls('scss', s, '// nested')).toBe('c')
    expect(cls('scss', s, '&')).toBe('k')
    expect(cls('scss', s, 'padding')).toBe('a')
    expect(cls('scss', s, 'li')).toBe('t')
  })

  test('HTML / XML: tags, attributes, entities, comments; scripts and styles in their language', () => {
    const t = '<!doctype html>\n<p class="x" hidden data-n=3>a &amp; b</p>\n<!-- one\ntwo -->\n<script>\nconst a = "s" // c\n</script>\n<style>p { color: red }</style>'
    expect(cls('html', t, '<!doctype html>')).toBe('k')
    expect(cls('html', t, '<p')).toBe('t')
    expect(cls('html', t, 'class')).toBe('a')
    expect(cls('html', t, '"x"')).toBe('s')
    expect(cls('html', t, 'hidden')).toBe('a')
    expect(cls('html', t, '3')).toBe('s')
    expect(cls('html', t, '&amp;')).toBe('n')
    expect(cls('html', t, 'two -->')).toBe('c')
    expect(cls('html', t, 'const')).toBe('k')
    expect(cls('html', t, '// c')).toBe('c')
    expect(cls('html', t, '</script')).toBe('t')
    expect(cls('html', t, 'color')).toBe('a')
    const x = '<?xml version="1.0"?>\n<a:b xmlns:a="u"><![CDATA[<raw>]]></a:b>'
    expect(cls('xml', x, '<?xml version="1.0"?>')).toBe('m')
    expect(cls('xml', x, '<a:b')).toBe('t')
    expect(cls('xml', x, 'xmlns:a')).toBe('a')
    expect(cls('xml', x, '<raw>')).toBe('c')
  })

  test('YAML: keys, values, comments, items, anchors, block scalars', () => {
    const t = '---\nname: demo # the name\nport: 8080\non: [push, pull_request]\nsteps:\n  - run: |\n      echo hi\n      echo "x: y"\n  - "quoted key": &a yes\nafter: *a'
    expect(cls('yaml', t, '---')).toBe('k')
    expect(cls('yaml', t, 'name')).toBe('a')
    expect(cls('yaml', t, 'demo')).toBe('s')
    expect(cls('yaml', t, '# the name')).toBe('c')
    expect(cls('yaml', t, '8080')).toBe('n')
    expect(cls('yaml', t, 'push')).toBe('s')
    expect(cls('yaml', t, '-', 3)).toBe('o')
    expect(cls('yaml', t, 'run')).toBe('a')
    expect(cls('yaml', t, '|')).toBe('o')
    expect(cls('yaml', t, '      echo "x: y"')).toBe('s')
    expect(cls('yaml', t, '"quoted key"')).toBe('a')
    expect(cls('yaml', t, '&a')).toBe('m')
    expect(cls('yaml', t, 'yes')).toBe('n')
    expect(cls('yaml', t, 'after')).toBe('a')
    expect(cls('yaml', t, '*a')).toBe('m')
  })

  test('shell: commands, variables, strings, comments, here-documents', () => {
    const t = '#!/bin/bash\n# build\nNAME="w $USER"\nif [ -n "$1" ]; then\n  echo ${NAME} | tr a b > out.txt\nfi\ncat <<EOF\nraw $text\nEOF\nls -la'
    expect(cls('shell', t, '#!/bin/bash')).toBe('m')
    expect(cls('shell', t, '# build')).toBe('c')
    expect(cls('shell', t, 'NAME')).toBe('a')
    expect(cls('shell', t, '"w ')).toBe('s')
    expect(cls('shell', t, '$USER')).toBe('a')
    expect(cls('shell', t, 'if')).toBe('k')
    expect(cls('shell', t, 'then')).toBe('k')
    expect(cls('shell', t, 'echo')).toBe('f')
    expect(cls('shell', t, '${NAME}')).toBe('a')
    expect(cls('shell', t, '|')).toBe('o')
    expect(cls('shell', t, 'tr')).toBe('f')
    expect(cls('shell', t, '<<EOF')).toBe('o')
    expect(cls('shell', t, 'raw $text')).toBe('s')
    expect(cls('shell', t, 'ls')).toBe('f')
  })

  test('SQL: keywords in any case, strings, names, comments', () => {
    const t = "SELECT id, \"Name\" FROM users u -- all\nWHERE u.note = 'it''s' AND count(*) > 10;\n/* multi\nline */ select 1"
    expect(cls('sql', t, 'SELECT')).toBe('k')
    expect(cls('sql', t, 'select')).toBe('k')
    expect(cls('sql', t, '"Name"')).toBe('a')
    expect(cls('sql', t, '-- all')).toBe('c')
    expect(cls('sql', t, "'it''s'")).toBe('s')
    expect(cls('sql', t, 'count')).toBe('f')
    expect(cls('sql', t, '10')).toBe('n')
    expect(cls('sql', t, 'line */')).toBe('c')
  })

  test('Go, Rust, Java, C, C++, C#', () => {
    const go = 'package main\nfunc main() {\n\ts := `raw\nstring`\n\tvar n int = 0\n}'
    expect(cls('go', go, 'func')).toBe('k')
    expect(cls('go', go, 'main', 1)).toBe('f')
    expect(cls('go', go, 'string`')).toBe('s')
    expect(cls('go', go, 'int')).toBe('y')
    const rs = "#[derive(Debug)]\nfn f<'a>(x: &'a str) -> Option<u8> { /* a /* b */ c */ println!(\"{}\", r#\"q\"#); 'c' }"
    expect(cls('rust', rs, '#[derive(Debug)]')).toBe('m')
    expect(cls('rust', rs, "'a", 0)).toBe('k')
    expect(cls('rust', rs, 'str')).toBe('y')
    expect(cls('rust', rs, 'Option')).toBe('y')
    expect(cls('rust', rs, ' c */')).toBe('c')
    expect(cls('rust', rs, 'println')).toBe('f')
    expect(cls('rust', rs, 'r#"q"#')).toBe('s')
    expect(cls('rust', rs, "'c'")).toBe('s')
    const java = '@Override\npublic String text() {\n  return """\n    block\n    """;\n}'
    expect(cls('java', java, '@Override')).toBe('m')
    expect(cls('java', java, 'String')).toBe('y')
    expect(cls('java', java, 'block')).toBe('s')
    const cc = '#include <stdio.h>\n#define MAX 10\nint main(void) { char *s = "x"; return MAX; }'
    expect(cls('c', cc, '#include')).toBe('m')
    expect(cls('c', cc, '<stdio.h>')).toBe('s')
    expect(cls('c', cc, 'int')).toBe('y')
    expect(cls('c', cc, 'main')).toBe('f')
    expect(cls('c', cc, 'MAX', 1)).toBe('n')
    const cpp = 'auto s = R"x(a "quoted" )\nstring)x";\nclass Widget : public Base {};\nstd::vector<int> v{nullptr};'
    expect(cls('cpp', cpp, 'string)x"')).toBe('s')
    expect(cls('cpp', cpp, 'Widget')).toBe('y')
    expect(cls('cpp', cpp, 'nullptr')).toBe('n')
    const cs = 'var p = @"C:\\dir\n""quoted""";\nvar q = $"{a} b";\n#region X\npublic class C { }'
    expect(cls('csharp', cs, '""quoted"""')).toBe('s')
    expect(cls('csharp', cs, '$"{a} b"')).toBe('s')
    expect(cls('csharp', cs, '#region')).toBe('m')
    expect(cls('csharp', cs, 'class')).toBe('k')
  })

  test('Markdown: code blocks in their own language; diff; TOML', () => {
    const md = '# Title\n\n```js\nconst a = "b" // c\n```\n\n````md\n```js\nnot closed by this\n```\n````\n\ntext **bold**'
    expect(cls('markdown', md, '# Title')).toBe('h')
    expect(cls('markdown', md, 'const')).toBe('k')
    expect(cls('markdown', md, '"b"')).toBe('s')
    expect(cls('markdown', md, '// c')).toBe('c')
    expect(cls('markdown', md, '**bold**')).toBe('b')
    // a fence with a language does not close a block
    expect(cls('markdown', md, '```js', 1)).toBe('k')
    // still JavaScript inside the outer block
    expect(cls('markdown', md, 'this')).toBe('k')
    expect(cls('diff', '--- a/x\n+++ b/x\n@@ -1 +1 @@\n-old\n+new\n same', '-old')).toBe('t')
    expect(cls('diff', '--- a/x\n+++ b/x\n@@ -1 +1 @@\n-old\n+new\n same', '+new')).toBe('s')
    const toml = '[server]\nhost = "a" # c\nport = 80\ndebug = true\nnote = """\nmulti\n"""'
    expect(cls('toml', toml, '[server]')).toBe('y')
    expect(cls('toml', toml, 'host ')).toBe('a')
    expect(cls('toml', toml, '"a"')).toBe('s')
    expect(cls('toml', toml, '80')).toBe('n')
    expect(cls('toml', toml, 'multi')).toBe('s')
  })
})

describe('syntax: bounded', () => {
  const LANGS = ['javascript', 'typescript', 'tsx', 'json', 'python', 'css', 'scss', 'html', 'xml', 'yaml', 'shell', 'sql', 'go', 'rust', 'java', 'c', 'cpp', 'csharp', 'markdown', 'mermaid', 'plantuml', 'graphviz', 'diff', 'toml', 'ini']

  test('the tokens always put the text back together, whatever the text', () => {
    // pieces that open and close what the scanners follow, in random order
    const pieces = ['"', "'", '`', '${', '}', '{', '/*', '*/', '//', '#', '<', '>', '</', '/>', '<!--', '-->', '\n', ' ', 'a', '1', '.', '\\', 'r#"', '"""', "'''", '[', ']', '(', ')', ':', '- ', '|', '@x', '$x', '```', '~~~', '<script>', '</script>', 'R"x(', ')x"', '@"', '\t', '\r\n', 'é', '😀']
    let seed = 7
    const rnd = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff)
    for (let round = 0; round < 40; round++) {
      let text = ''
      for (let k = 0; k < 400; k++) text += pieces[Math.floor(rnd() * pieces.length)]
      for (const lang of LANGS) {
        const toks = highlight(lang, text)
        expect(joined(toks)).toBe(text)
        expect(toks.every((t) => t[1].length > 0)).toBe(true)
        const lines = highlightLines(lang, text)
        expect(lines.length).toBe(text.split('\n').length)
      }
    }
  })

  test('a text over the limit is plain; a line over the limit is plain and keeps the state around it', () => {
    const big = 'const x = 1\n'.repeat(1000)
    expect(highlight('javascript', big, 100)).toEqual(highlight(undefined, big))
    const long = '/* open\n' + 'x'.repeat(20_000) + '\nstill */ const'
    const lines = highlightLines('javascript', long)
    expect(lines[1]).toEqual([['', 'x'.repeat(20_000)]])
    expect(lines[2][0]).toEqual(['c', 'still */'])
    expect(lines[2].at(-1)).toEqual(['k', 'const'])
  })

  test('nesting is capped', () => {
    const deep = 'const s = ' + '`${'.repeat(5_000) + '\nnext()'
    const toks = highlight('javascript', deep)
    expect(joined(toks)).toBe(deep)
  })

  test('pathological inputs stay fast', () => {
    // long lines of what opens a string, a comment, a tag, a link, an emphasis… and never closes it: a scanner that
    // looks ahead to the end of the line from each of them is quadratic (seconds here), a linear one takes milliseconds
    const text = (s: string) => Array.from({ length: 2 }, () => s.repeat(Math.floor(9_000 / s.length))).join('\n')
    const inputs = ['"', '"\\', '`', '/', '/*', '<a', '<!--', '[', '**x ', '_x ', '[x](', '<<', '-'.repeat(50), '${', 'r#', 'R"', '"""', '<http://', '#', 'a:b ']
    let worst = 0
    let what = ''
    for (const s of inputs) {
      const t = text(s)
      for (const lang of LANGS) {
        const t0 = performance.now()
        const toks = highlight(lang, t + '\n' + lang + s)
        const ms = performance.now() - t0
        expect(joined(toks)).toBe(t + '\n' + lang + s)
        if (ms > worst) [worst, what] = [ms, `${lang} ${JSON.stringify(s)}`]
      }
    }
    expect(`${what}: ${worst < 1000 ? 'fast' : `${Math.round(worst)} ms`}`).toBe(`${what}: fast`)
  }, 120_000)

  test('an edit reads again only the lines it changed (the others come from the cache)', () => {
    const text = Array.from({ length: 200 }, (_, k) => `const v${k} = ${k} // line ${k}`).join('\n')
    const a = highlightLines('javascript', text)
    const b = highlightLines('javascript', text.replace('v100 =', 'v100 +='))
    expect(b[99]).toBe(a[99])
    expect(b[101]).toBe(a[101])
    expect(b[100]).not.toBe(a[100])
  })
})
