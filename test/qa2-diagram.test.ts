import { describe, expect, test } from 'bun:test'
import { markdownToHtml } from '../src/core/convert'
import { DiagramError, renderDiagrams } from '../src/core/diagram'
import { parseDot } from '../src/core/diagram/dot'
import type { Graph } from '../src/core/diagram/graph'
import { parseMermaid } from '../src/core/diagram/mermaid'
import type { Model } from '../src/core/diagram/model'
import { parsePlantUml } from '../src/core/diagram/plantuml'

const graphOf = (m: Model): Graph => {
  if (m.t !== 'graph') throw new Error(`expected a graph, got ${m.t}`)
  return m.graph
}
const puml = (body: string) => parsePlantUml(`@startuml\n${body}\n@enduml`)[0]
const ids = (g: Graph) => g.nodes.map((n) => n.id)
const labels = (g: Graph) => g.nodes.map((n) => n.label)
const edgeList = (g: Graph) => g.edges.map((e) => `${e.from}>${e.to}`)
const failsWith = (lang: 'mermaid' | 'plantuml' | 'graphviz', src: string, re: RegExp) => {
  let err: unknown
  try {
    renderDiagrams(lang, src)
  } catch (e) {
    err = e
  }
  expect(err).toBeInstanceOf(DiagramError)
  expect((err as Error).message).toMatch(re)
}

describe('P3-023 PlantUML name / alias', () => {
  const KINDS = ['actor', 'usecase', 'component', 'database', 'node', 'cloud', 'rectangle', 'agent', 'folder', 'storage']
  for (const kw of KINDS) {
    test(`${kw}: the three forms name the element, edges go through the alias`, () => {
      for (const decl of [`${kw} Guest as g`, `${kw} "Guest User" as g`, `${kw} g as "Guest User"`]) {
        const g = graphOf(puml(`${decl}\nusecase UC1\ng --> UC1`))
        expect(ids(g)).toEqual(['g', 'UC1'])
        expect(g.nodes[0].label).toBe(decl.includes('Guest User') ? 'Guest User' : 'Guest')
        expect(edgeList(g)).toEqual(['g>UC1'])
      }
    })
  }
  test('class, object and state declarations', () => {
    const c = graphOf(puml('class Cat as C\nclass "Big Owner" as O\nC --> O'))
    expect(ids(c)).toEqual(['C', 'O'])
    expect(labels(c)).toEqual(['Cat', 'Big Owner'])
    const s = graphOf(puml('state Idle as I\nstate "Busy now" as B\n[*] --> I\nI --> B'))
    expect(s.nodes.find((n) => n.id === 'I')!.label).toBe('Idle')
    expect(s.nodes.find((n) => n.id === 'B')!.label).toBe('Busy now')
    expect(s.nodes.filter((n) => n.id === 'I' || n.id === 'B').length).toBe(2)
  })
  test('sequence participants', () => {
    const m = puml('actor Guest as g\nparticipant "Long name" as L\nparticipant Bob as b\ng -> L : hi\nL -> b : yo')
    if (m.t !== 'sequence') throw new Error(m.t)
    expect(m.seq.participants.map((p) => `${p.id}=${p.label}`)).toEqual(['g=Guest', 'L=Long name', 'b=Bob'])
  })
  test('bracket shorthands keep working', () => {
    const g = graphOf(puml('[First component] as C1\n(Use the application) as UC\nC1 --> UC'))
    expect(ids(g).sort()).toEqual(['C1', 'UC'])
  })
})

describe('P3-024 PlantUML deployment elements', () => {
  test('every element keyword of the official deployment example is drawn', () => {
    const kws = ['actor', 'agent', 'artifact', 'boundary', 'card', 'circle', 'cloud', 'collections', 'component', 'control', 'database', 'entity', 'file', 'folder', 'frame', 'hexagon', 'interface', 'label', 'node', 'package', 'person', 'queue', 'rectangle', 'stack', 'storage', 'usecase']
    const g = graphOf(puml(kws.map((k) => `${k} ${k}`).join('\n')))
    expect(ids(g).sort()).toEqual([...kws].sort())
  })
  test('actor + database + queue with arrows and no participant stay a description diagram', () => {
    const g = graphOf(puml('actor A\ndatabase D\nqueue Q\ncloud K\nA --> D\nD --> Q\nQ --> K'))
    expect(edgeList(g)).toEqual(['A>D', 'D>Q', 'Q>K'])
  })
  test('a plain sequence diagram with boundary / control stays a sequence', () => {
    expect(puml('boundary B\ncontrol C\nB -> C : go (now)').t).toBe('sequence')
  })
})

describe('P3-025 Mermaid markdown strings', () => {
  const SRC = `flowchart LR
subgraph "One"
  a("\`The **cat**
  in the hat\`") -- "edge label" --> b{{"\`The **dog**
  in the hog\`"}}
end
subgraph 'Two'
  c("\`The **cat**
  in the hat\`") -- "\`Bold **edge label**\`" --> d("The dog
  in the hog")
end`
  test('labels, edges and groups', () => {
    const g = graphOf(parseMermaid(SRC))
    expect(ids(g)).toEqual(['a', 'b', 'c', 'd'])
    expect(g.nodes.map((n) => n.label.replace(/\s+/g, ' '))).toEqual(['The cat in the hat', 'The dog in the hog', 'The cat in the hat', 'The dog in the hog'])
    expect(edgeList(g)).toEqual(['a>b', 'c>d'])
    expect(g.edges.map((e) => e.label)).toEqual(['edge label', 'Bold edge label'])
    expect(g.clusters.map((c) => c.label)).toEqual(['One', 'Two'])
    expect(renderDiagrams('mermaid', SRC).length).toBe(1)
  })
  test('single-line markdown labels', () => {
    const g = graphOf(parseMermaid('graph TD\nA["`x **y** _z_`"] --> B'))
    expect(labels(g)).toEqual(['x y z', 'B'])
  })
  test('an unterminated string is an error naming the line', () => {
    failsWith('mermaid', 'graph TD\nA["`oops --> B\nC --> D', /line 2/i)
  })
})

describe('P3-026 Mermaid / DOT', () => {
  test('edge labels with | or -- inside quotes', () => {
    const g = graphOf(parseMermaid('graph LR\nA -->|"a | b"| B\nA -- "text with -- dashes" --> C'))
    expect(ids(g)).toEqual(['A', 'B', 'C'])
    expect(g.edges.map((e) => e.label)).toEqual(['a | b', 'text with -- dashes'])
  })
  test('edge ids are read', () => {
    const g = graphOf(parseMermaid('flowchart LR\nA e1@--> B\nB e2@==> C\ne1@{ animate: true }\ne2@{ curve: linear }'))
    expect(ids(g)).toEqual(['A', 'B', 'C'])
    expect(edgeList(g)).toEqual(['A>B', 'B>C'])
    expect(g.edges[1].line).toBe('thick')
  })
  test('sequence box groups participants', () => {
    const m = parseMermaid('sequenceDiagram\nbox Purple Alice & John\nparticipant A as Alice\nparticipant J as John\nend\nparticipant B as Bob\nA->>J: hi\nJ->>B: yo')
    if (m.t !== 'sequence') throw new Error(m.t)
    expect(m.seq.boxes?.length).toBe(1)
    expect(m.seq.boxes![0].label).toBe('Alice & John')
    expect(m.seq.participants.map((p) => p.box)).toEqual([0, 0, undefined])
    const svg = renderDiagrams('mermaid', 'sequenceDiagram\nbox Purple Alice & John\nparticipant A\nend\nA->>A: x')[0].svg
    expect(svg).toContain('Alice &amp; John')
  })
  test('sequence box with a color and no title', () => {
    const m = parseMermaid('sequenceDiagram\nbox rgb(33,66,99)\nparticipant A\nend\nbox transparent Aqua\nparticipant B\nend\nA->>B: x')
    if (m.t !== 'sequence') throw new Error(m.t)
    expect(m.seq.boxes!.map((b) => b.label)).toEqual(['', 'Aqua'])
  })
  test('DOT strict merges duplicate edges', () => {
    expect(parseDot('strict digraph { a -> b; a -> b; b -> a }')[0].edges.map((e) => `${e.from}>${e.to}`)).toEqual(['a>b', 'b>a'])
    expect(parseDot('strict graph { a -- b; b -- a; a -- b [color=red] }')[0].edges.length).toBe(1)
    expect(parseDot('digraph { a -> b; a -> b }')[0].edges.length).toBe(2)
  })
})

describe('P3-027 PlantUML silent drops', () => {
  test('state descriptions', () => {
    const g = graphOf(puml('[*] --> State1\nState1 --> [*]\nState1 : this is a string\nState1 : this is another string\nState1 -> State2\nState2 : Just a test'))
    const flat = (id: string) => (g.nodes.find((n) => n.id === id)!.sections ?? []).flatMap((s) => s.rows.flat())
    expect(flat('State1')).toEqual(['State1', 'this is a string', 'this is another string'])
    expect(flat('State2')).toEqual(['State2', 'Just a test'])
    const svg = renderDiagrams('plantuml', '@startuml\nState1 : this is a string\n[*] --> State1\n@enduml')[0].svg
    expect(svg).toContain('this is a string')
  })
  test('activity swimlanes', () => {
    const g = graphOf(puml('|Swimlane1|\nstart\n:foo1;\n|#AntiqueWhite|Swimlane2|\n:foo2;\n|Swimlane1|\n:foo3;\nstop'))
    expect(g.clusters.map((c) => c.label)).toEqual(['Swimlane1', 'Swimlane2'])
    const lane = (t: string) => g.clusters.find((c) => c.label === t)!.id
    expect(g.nodes.filter((n) => n.label === 'foo1' || n.label === 'foo3').map((n) => n.cluster)).toEqual([lane('Swimlane1'), lane('Swimlane1')])
    expect(g.nodes.find((n) => n.label === 'foo2')!.cluster).toBe(lane('Swimlane2'))
    expect(g.edges.length).toBe(4)
  })
  test('activity notes', () => {
    const g = graphOf(puml('start\n:foo1;\nnote right\n  This note is on\n  several lines\nend note\n:foo2;\nnote left: one line\nfloating note left: free\nstop'))
    const notes = g.nodes.filter((n) => n.shape === 'note')
    expect(notes.map((n) => n.label)).toEqual(['This note is on\nseveral lines', 'one line', 'free'])
    expect(g.nodes.some((n) => n.label === 'foo2')).toBe(true)
    // the two attached notes hang off the action before them
    expect(g.edges.filter((e) => e.line === 'dashed').length).toBe(2)
  })
  test('preprocessing: !define and variables', () => {
    const m = puml('!define WHO Charlie\n!$msg = "from var"\n!$n ?= 5\nAlice -> WHO : $msg\nWHO -> Bob : $n')
    if (m.t !== 'sequence') throw new Error(m.t)
    expect(m.seq.participants.map((p) => p.id)).toEqual(['Alice', 'Charlie', 'Bob'])
    expect(m.seq.events.map((e) => (e.t === 'msg' ? e.text : ''))).toEqual(['from var', '5'])
  })
  test('preprocessing: macros with arguments, !undef, conditionals', () => {
    const m = puml('!define greet(x) Hello x\n!define FLAG\n!ifdef FLAG\nA -> B : greet(world)\n!else\nA -> B : no\n!endif\n!ifndef FLAG\nA -> B : never\n!endif')
    if (m.t !== 'sequence') throw new Error(m.t)
    expect(m.seq.events.map((e) => (e.t === 'msg' ? e.text : ''))).toEqual(['Hello world'])
  })
})

describe('N3-008 syntax errors are reported', () => {
  test('Mermaid flowchart', () => {
    failsWith('mermaid', 'graph TD\nA --> B\nA[unclosed --> B', /line 3/)
    failsWith('mermaid', 'graph TD\nA -->|unterminated B', /line 2/)
    failsWith('mermaid', 'graph TD\nA -->', /line 2/)
    failsWith('mermaid', 'graph TD\nA --> B\n%% a comment\n???', /line 4/)
  })
  test('Mermaid flowchart: valid syntax stays valid', () => {
    const ok = [
      'graph TD\n%% comment\nA[Start] --> B{Q?}\nB -->|yes| C((x)) & D\nclassDef c fill:#f9f\nclass A c\nstyle B fill:#fff\nlinkStyle 0 stroke:red\nclick A callback "tip"\nsubgraph s1 [Title]\ndirection LR\nC --> D\nend\nE@{ shape: rect, label: "x" } --> F:::c\nA -.-> B\nA ==> B\nA ~~~ B\nA <--> B\nA --x B\nA --o B\naccTitle: t\naccDescr: d',
      'flowchart LR\nA-->B;B-->C;\nA --- B\nA-- text --> B\nA -- text --- B\nA[/x/] --> B[\\y\\]\nid1(["a"]) --> id2[("b")] --> id3[[c]]',
      '---\ntitle: T\n---\n%%{init: {"theme":"dark"}}%%\ngraph TB\nA --> B',
    ]
    for (const s of ok) expect(renderDiagrams('mermaid', s).length).toBe(1)
  })
  test('Mermaid sequence', () => {
    failsWith('mermaid', 'sequenceDiagram\nA->>B: ok\nA->>', /line 3/)
    failsWith('mermaid', 'sequenceDiagram\nparticipant A\nwhat is this', /line 3/)
  })
  test('Mermaid sequence: valid syntax stays valid', () => {
    const s = 'sequenceDiagram\n%% c\nautonumber\ntitle: T\nparticipant A as Alice\nactor B\nA->>+B: hi\nB-->>-A: yo\nNote over A,B: n\nNote right of A: r\nloop every\nA->>B: x\nend\nalt a\nA-)B: z\nelse b\nA-xB: q\nend\nrect rgb(1,2,3)\nA->>B: w\nend\nactivate A\ndeactivate A\ncreate participant C\nA->>C: n\ndestroy C\nlinks A: {"x": "http://a"}\nopt o\nB->>A: m\nend\npar p\nA->>B: 1\nand q\nA->>B: 2\nend\ncritical c\nA->>B: 1\noption o\nA->>B: 2\nend\nbreak b\nA->>B: 1\nend'
    expect(renderDiagrams('mermaid', s).length).toBe(1)
  })
  test('DOT', () => {
    failsWith('graphviz', 'digraph {\n a -> b [\n}', /line/)
    failsWith('graphviz', 'digraph G {{{\n a -> b', /line|\}/)
    failsWith('graphviz', 'digraph {\n a ->\n}', /line 2/)
    failsWith('graphviz', 'digraph {\n a -> b @ c\n}', /line 2/)
    failsWith('graphviz', 'digraph {\n a -- b\n}', /line 2/)
    failsWith('graphviz', 'digraph {\n a [label="open]\n}', /line 2/)
  })
  test('DOT: valid syntax stays valid', () => {
    const s = '/* c */ strict digraph "G" {\n# pre\n// c\nrankdir=LR; node [shape=box]; edge [color=red]\na:p1 -> b:n -> { c d } [label=<<b>x</b>>]\nsubgraph cluster_0 { label="L"; e -> f }\n{ rank=same; a; b }\n"q" -> a [label="x" + "y"]\n-1 -> 2.5\n}'
    expect(renderDiagrams('graphviz', s).length).toBe(1)
  })
})

describe('P3-019 Markdown emphasis is linear', () => {
  const time = (md: string) => {
    const t = performance.now()
    const html = markdownToHtml(md)
    return [performance.now() - t, html] as const
  }
  for (const unit of ['*a **a ', '_a __a ', '~~a *a _a ', '[a](b ', '[a ', '**a ']) {
    test(`pattern ${JSON.stringify(unit)}`, () => {
      const [t1] = time(unit.repeat(10000))
      const [t2] = time(unit.repeat(80000))
      // the old code took ~10 s for 50k; linear takes well under a second even in the slow VM
      expect(t2).toBeLessThan(5000)
      expect(t1).toBeLessThan(2000)
    })
  }
  test('a long run of spaces in a heading', () => {
    expect(time('# a' + ' '.repeat(100000) + 'b')[0]).toBeLessThan(3000)
  })
  test('output of matched emphasis is unchanged', () => {
    expect(markdownToHtml('a **b** *c* _d_ ~~e~~ __f__ `g` [h](http://x) ![i](http://y)')).toBe(
      '<p>a <b>b</b> <i>c</i> <i>d</i> <s>e</s> <b>f</b> <code>g</code> <a href="http://x">h</a> i</p>',
    )
    expect(markdownToHtml('**a __b** c__ and *x **y** z* snake_case_name 2*3*4')).toBe('<p><b>a <i>_b</b> c_</i> and <i>x <b>y</b> z</i> snake_case_name 2*3*4</p>')
    expect(markdownToHtml('## T ##  \n\n# #')).toBe('<h2>T</h2>\n<h1></h1>')
  })
})
