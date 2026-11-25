/**
 * Conversions on worker threads: a file that takes too long, or a caller that gives up, ends its thread (and with it
 * whatever the engine was doing), and the server's own thread is never blocked by a large file.
 */
import { availableParallelism } from 'node:os'
import { Worker } from 'node:worker_threads'
import { convert, ConvertError, type ConvertOptions, type Converted } from './convert'
import type { Job, JobResult } from './worker'

export interface ConverterOptions {
  /** threads converting at once (default: the number of CPUs); 0 converts on this thread, without time limits */
  workers?: number
  /** conversions waiting for a thread before more are refused with a `BusyError` (default 100) */
  queue?: number
  /** milliseconds a conversion may take before its thread is ended (default 120000; 0: no limit) */
  timeout?: number
  /** megabytes of JavaScript heap a thread may use (Node only: Bun has no such limit) */
  memory?: number
}

/** the conversion took longer than `timeout` */
export class TimeoutError extends Error {
  constructor(ms: number) {
    super(`The conversion took longer than ${ms / 1000} s`)
    this.name = 'TimeoutError'
  }
}

/** the thread converting the file stopped (out of memory, a crash) */
export class ThreadError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'ThreadError'
  }
}

/** every thread is busy and the queue is full */
export class BusyError extends Error {
  constructor() {
    super('The converter is busy: try again later')
    this.name = 'BusyError'
  }
}

interface Task {
  data: Uint8Array
  opts: Omit<ConvertOptions, 'signal'>
  signal?: AbortSignal
  resolve: (r: Converted) => void
  reject: (e: Error) => void
}

interface Thread {
  worker: Worker
  task?: Task
  timer?: ReturnType<typeof setTimeout>
  abort?: () => void
}

// the source (`bun packages/convert/src/cli.ts`) or the built package (dist/index.js next to dist/worker.js)
const WORKER_URL = new URL(import.meta.url.endsWith('.ts') ? './worker.ts' : './worker.js', import.meta.url)

/** Converts files on a pool of worker threads, one file per thread at a time. */
export class Converter {
  readonly workers: number
  readonly timeout: number
  private readonly queueMax: number
  private readonly memory?: number
  private readonly idle: Thread[] = []
  private readonly busy = new Set<Thread>()
  private readonly waiting: Task[] = []
  private seq = 0
  private closed = false

  constructor(opts: ConverterOptions = {}) {
    this.workers = Math.max(0, Math.floor(opts.workers ?? availableParallelism()))
    this.timeout = opts.timeout ?? 120_000
    this.queueMax = opts.queue ?? 100
    this.memory = opts.memory
  }

  /** conversions running and waiting */
  get load(): { running: number; queued: number } {
    return { running: this.busy.size, queued: this.waiting.length }
  }

  convert(data: Uint8Array, opts: ConvertOptions = {}): Promise<Converted> {
    if (this.closed) return Promise.reject(new Error('The converter is closed'))
    if (!this.workers) return convert(data, opts)
    const { signal, ...rest } = opts
    if (signal?.aborted) return Promise.reject(signal.reason)
    return new Promise<Converted>((resolve, reject) => {
      const task: Task = { data, opts: rest, signal, resolve, reject }
      const thread = this.idle.pop() ?? (this.busy.size < this.workers ? this.spawn() : undefined)
      if (thread) return this.run(thread, task)
      if (this.waiting.length >= this.queueMax) return reject(new BusyError())
      this.waiting.push(task)
      signal?.addEventListener(
        'abort',
        () => {
          const at = this.waiting.indexOf(task)
          if (at < 0) return
          this.waiting.splice(at, 1)
          reject(signal.reason)
        },
        { once: true },
      )
    })
  }

  /** ends every thread; conversions under way and waiting are rejected */
  async close(): Promise<void> {
    this.closed = true
    for (const t of this.waiting.splice(0)) t.reject(new Error('The converter is closed'))
    const all = [...this.idle.splice(0), ...this.busy]
    this.busy.clear()
    for (const t of all) this.end(t, new Error('The converter is closed'))
    await Promise.all(all.map((t) => t.worker.terminate()))
  }

  private spawn(): Thread {
    const worker = new Worker(WORKER_URL, this.memory ? { resourceLimits: { maxOldGenerationSizeMb: this.memory } } : {})
    const thread: Thread = { worker }
    worker.on('message', (msg: JobResult) => {
      const task = thread.task
      if (!task) return
      this.release(thread)
      if (msg.ok) task.resolve({ ...msg.result, data: new Uint8Array(msg.result.data) })
      else task.reject(msg.code ? new ConvertError(msg.error, msg.code) : new Error(msg.error))
    })
    // a thread that dies (out of memory, a crash) takes its conversion with it; the next one gets a new thread
    worker.on('error', (e: Error) => this.discard(thread, new ThreadError(`The conversion thread failed: ${e.message}`)))
    worker.on('exit', () => this.discard(thread, new ThreadError('The conversion thread stopped')))
    return thread
  }

  private run(thread: Thread, task: Task): void {
    if (task.signal?.aborted) {
      task.reject(task.signal.reason)
      return this.next(thread)
    }
    thread.task = task
    this.busy.add(thread)
    // (a thread at work keeps the process alive, an idle one does not)
    thread.worker.ref()
    if (this.timeout) thread.timer = setTimeout(() => this.kill(thread, new TimeoutError(this.timeout)), this.timeout)
    if (task.signal) {
      const signal = task.signal
      thread.abort = () => this.kill(thread, signal.reason instanceof Error ? signal.reason : new Error('The conversion was aborted'))
      signal.addEventListener('abort', thread.abort, { once: true })
    }
    const id = ++this.seq
    // the bytes move to the thread: the caller's copy is left empty when it owns the whole buffer
    const data = task.data.byteOffset === 0 && task.data.byteLength === task.data.buffer.byteLength ? task.data : task.data.slice()
    thread.worker.postMessage({ id, data, opts: task.opts } satisfies Job, [data.buffer as ArrayBuffer])
  }

  /** the thread is done with its task: it takes the next one, or waits */
  private release(thread: Thread) {
    clearTimeout(thread.timer)
    if (thread.abort) thread.task?.signal?.removeEventListener('abort', thread.abort)
    thread.task = thread.timer = thread.abort = undefined
    this.busy.delete(thread)
    this.next(thread)
  }

  private next(thread: Thread): void {
    if (this.closed) return
    const task = this.waiting.shift()
    if (task) return this.run(thread, task)
    thread.worker.unref()
    this.idle.push(thread)
  }

  /** ends a thread in the middle of its conversion; a new one takes the queue on */
  private kill(thread: Thread, err: Error) {
    this.discard(thread, err)
    void thread.worker.terminate()
  }

  private discard(thread: Thread, err: Error) {
    const known = this.busy.delete(thread) || this.idle.includes(thread)
    const at = this.idle.indexOf(thread)
    if (at >= 0) this.idle.splice(at, 1)
    this.end(thread, err)
    if (!known || this.closed) return
    const task = this.waiting.shift()
    if (task) this.run(this.spawn(), task)
  }

  private end(thread: Thread, err: Error) {
    const task = thread.task
    clearTimeout(thread.timer)
    if (thread.abort) task?.signal?.removeEventListener('abort', thread.abort)
    thread.task = thread.timer = thread.abort = undefined
    task?.reject(err)
  }
}
