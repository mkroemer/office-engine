#!/usr/bin/env node
/** office-convert: converts files on the command line, or runs the HTTP service (`office-convert serve`). */
import { existsSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { availableParallelism } from 'node:os'
import { basename, dirname, extname, join, resolve } from 'node:path'
import { parseArgs } from 'node:util'
import pkg from '../package.json' with { type: 'json' }
import { targetOf, type Target } from './convert'
import { Converter } from './pool'
import { createHandler } from './server'
import { listen } from './serve'

const USAGE = `office-convert ${pkg.version}: Office files converted without LibreOffice

Usage:
  office-convert [options] <file…>     each file next to itself, or into --out
  office-convert <in> <out>            the format from the output's extension
  office-convert serve [options]       the HTTP service (POST /convert, Gotenberg's routes)

Options:
  -t, --to <format>     docx, xlsx, pptx, txt, md, html, csv, tsv
                        (default: the Office format the file opens as: .doc → .docx, .csv → .xlsx)
  -o, --out <path>      the output file (one input; - for stdout) or an existing directory
      --name <name>     the name of the file read from stdin (-): its extension tells text formats
      --sheet <n|name>  CSV / TSV: the sheet written (default: the one the file was saved on)
      --encoding <enc>  text files: the encoding they are read in (default: a guess)
  -f, --force           overwrite files that exist
  -j, --jobs <n>        files converted at once (default: the number of CPUs)
      --timeout <s>     seconds a file may take (default 120; 0: no limit)
      --max-size <MB>   largest file read (default 100)

serve:
  -p, --port <n>        port (default $PORT or 3000)
      --host <addr>     address (default $HOST or 0.0.0.0)
  -j, --jobs <n>        conversions at once (default $WORKERS or the number of CPUs)
      --timeout <s>     seconds a file may take (default $TIMEOUT or 120)
      --max-size <MB>   largest request (default $MAX_SIZE or 100)
      --max-files <n>   files in one request (default $MAX_FILES or 20)

  -h, --help, -v, --version
`

const fail = (message: string, code = 2): never => {
  process.stderr.write(`office-convert: ${message}\n`)
  process.exit(code)
}

const { values: o, positionals } = (() => {
  try {
    return parseArgs({
      allowPositionals: true,
      options: {
        to: { type: 'string', short: 't' },
        out: { type: 'string', short: 'o' },
        name: { type: 'string' },
        sheet: { type: 'string' },
        encoding: { type: 'string' },
        force: { type: 'boolean', short: 'f' },
        jobs: { type: 'string', short: 'j' },
        timeout: { type: 'string' },
        'max-size': { type: 'string' },
        'max-files': { type: 'string' },
        port: { type: 'string', short: 'p' },
        host: { type: 'string' },
        help: { type: 'boolean', short: 'h' },
        version: { type: 'boolean', short: 'v' },
      },
    })
  } catch (e) {
    return fail(`${(e as Error).message}\n\n${USAGE}`)
  }
})()

/** a number option (a flag, else an environment variable, else the default) */
const num = (flag: string | undefined, env: string | undefined, dflt: number, what: string): number => {
  const v = flag ?? env
  if (v === undefined || v === '') return dflt
  const n = Number(v)
  if (!Number.isFinite(n) || n < 0) fail(`${what} must be a number, not “${v}”`)
  return n
}

if (o.help) {
  process.stdout.write(USAGE)
  process.exit(0)
}
if (o.version) {
  process.stdout.write(`${pkg.version}\n`)
  process.exit(0)
}

const env = process.env
if (positionals[0] === 'serve') {
  if (positionals.length > 1) fail(`serve takes no files\n\n${USAGE}`)
  const maxBytes = num(o['max-size'], env.MAX_SIZE, 100, '--max-size') * 1024 * 1024
  const converter = new Converter({ workers: num(o.jobs, env.WORKERS, NaN, '--jobs') || undefined, timeout: num(o.timeout, env.TIMEOUT, 120, '--timeout') * 1000 })
  const handler = createHandler(converter, { maxBytes, maxFiles: num(o['max-files'], env.MAX_FILES, 20, '--max-files'), version: pkg.version })
  const server = await listen(handler, { port: num(o.port, env.PORT, 3000, '--port'), host: o.host ?? env.HOST ?? '0.0.0.0', maxBytes })
  process.stderr.write(`office-convert ${pkg.version} listening on ${server.url} (${converter.workers} threads)\n`)
  const stop = () => {
    void server.close().then(() => converter.close()).then(() => process.exit(0))
  }
  process.on('SIGINT', stop)
  process.on('SIGTERM', stop)
} else await convertFiles()

async function convertFiles() {
  let inputs = positionals
  let out = o.out
  let to: Target | undefined
  if (o.to) to = targetOf(o.to) ?? fail(`Unknown format “${o.to}”: docx, xlsx, pptx, txt, md, html, csv, tsv`)
  // office-convert in.doc out.docx: a second name that is no file yet, with a format's extension, is the output
  if (!out && !to && inputs.length === 2 && !existsSync(inputs[1]) && targetOf(extname(inputs[1]))) {
    out = inputs[1]
    inputs = [inputs[0]]
  }
  if (!inputs.length) fail(`No file given\n\n${USAGE}`)
  const outDir = out && out !== '-' && existsSync(out) && statSync(out).isDirectory() ? out : undefined
  if (out && !outDir && inputs.length > 1) fail('--out names a file, but there are several inputs: give a directory')
  // the output's extension is the format, when --to does not say
  if (out && !outDir && out !== '-' && !to) to = targetOf(extname(out)) ?? undefined
  const maxBytes = num(o['max-size'], undefined, 100, '--max-size') * 1024 * 1024
  const converter = new Converter({ workers: Math.min(inputs.length, num(o.jobs, undefined, 0, '--jobs') || availableParallelism()), timeout: num(o.timeout, undefined, 120, '--timeout') * 1000 })

  let failed = 0
  // outputs this run writes: two inputs of the same name (report.doc, report.odt) would write one file
  const claimed = new Map<string, string>()
  await Promise.all(
    inputs.map(async (input) => {
      try {
        const stdin = input === '-'
        if (!stdin && statSync(input).size > maxBytes) throw new Error(`larger than ${maxBytes / 1024 / 1024} MB (--max-size)`)
        const data = new Uint8Array(readFileSync(stdin ? 0 : input))
        if (data.byteLength > maxBytes) throw new Error(`larger than ${maxBytes / 1024 / 1024} MB (--max-size)`)
        const r = await converter.convert(data, { ...(to ? { to } : {}), name: stdin ? (o.name ?? 'document') : basename(input), ...(o.sheet !== undefined ? { sheet: o.sheet } : {}), ...(o.encoding ? { encoding: o.encoding } : {}) })
        for (const w of r.warnings) process.stderr.write(`${input}: ${w}\n`)
        if (out === '-') return void process.stdout.write(r.data)
        const dest = out && !outDir ? out : join(outDir ?? (stdin ? '.' : dirname(input)), r.name)
        if (!stdin && resolve(dest) === resolve(input)) throw new Error(`would overwrite itself: give --out or --to`)
        const other = claimed.get(resolve(dest))
        if (other) throw new Error(`${dest} is written from ${other} already`)
        claimed.set(resolve(dest), input)
        if (!o.force && existsSync(dest)) throw new Error(`${dest} exists (--force overwrites it)`)
        writeFileSync(dest, r.data)
        process.stderr.write(`${input} → ${dest}\n`)
      } catch (e) {
        failed++
        process.stderr.write(`${input}: ${(e as Error).message}\n`)
      }
    }),
  )
  await converter.close()
  process.exit(failed ? 1 : 0)
}
