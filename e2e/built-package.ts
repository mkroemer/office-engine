/** Emitted-package suites require a complete, fresh production build. */
import { beforeAll } from 'bun:test'
import { existsSync, readdirSync, statSync } from 'node:fs'
import { resolve } from 'node:path'

export const packageRoot = resolve(process.env.BUILT_PACKAGE_ROOT ?? process.cwd())
export const builtModule = (file: string) => `/@fs${packageRoot}/dist/${file}`

function newestSource(path: string): number {
  if (!existsSync(path)) return 0
  if (!statSync(path).isDirectory()) return statSync(path).mtimeMs
  return Math.max(0, ...readdirSync(path).map((name) => newestSource(`${path}/${name}`)))
}

/** An external package without sources is checked for completeness only. */
export function useBuiltPackage() {
  beforeAll(() => {
    const outputs = ['index.js', 'locales/de.js', 'office-engine.worker.js', 'office-engine.embed.js', 'styles.css']
    for (const file of outputs) {
      if (!existsSync(`${packageRoot}/dist/${file}`)) throw new Error(`Missing dist/${file} in ${packageRoot}; run bun run build before browser tests`)
    }
    const inputs = ['src', 'package.json', 'bun.lock', 'vite.config.ts', 'vite.embed.config.ts', 'vite.worker.config.ts', 'tsconfig.build.json']
    // Published packages do not include source/build inputs; a checkout must always be fresh,
    // including when explicitly selected through BUILT_PACKAGE_ROOT.
    if (existsSync(`${packageRoot}/src`)) {
      const newest = Math.max(...inputs.map((file) => newestSource(`${packageRoot}/${file}`)))
      for (const file of outputs) {
        if (statSync(`${packageRoot}/dist/${file}`).mtimeMs < newest) throw new Error(`Stale dist/${file} in ${packageRoot}; run bun run build before browser tests`)
      }
    }
  })
}
