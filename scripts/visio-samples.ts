/** Writes the Visio demo file: playground/public/samples/sample.vsdx (built by test/visio-builder.ts). Usage: bun scripts/visio-samples.ts */
import { sampleVisio } from '../test/visio-builder'

const dir = `${import.meta.dir}/../playground/public/samples`
await Bun.write(`${dir}/sample.vsdx`, await sampleVisio())
console.log('wrote sample.vsdx in', dir)
