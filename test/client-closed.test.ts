import { describe, expect, test } from 'bun:test'
import { DocumentClient } from '../src/core/worker/client'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`

describe('a closed document', () => {
  test('says it is closed, so work that outlives it (a column fit after an edit) can stop', async () => {
    const client = await DocumentClient.open(Bun.file(S('sample.xlsx')) as unknown as Blob, { worker: false })
    expect(client.closed).toBe(false)
    await client.call('xlsx.block', 0, 0, 10)
    client.destroy()
    expect(client.closed).toBe(true)
    await expect(client.call('xlsx.block', 0, 0, 10)).rejects.toThrow('Document closed')
  })

  test('an open that is given up rejects at once (its worker ended); one that finished is not touched by a late signal', async () => {
    const file = Bun.file(S('sample.xlsx')) as unknown as Blob
    const before = new AbortController()
    before.abort()
    await expect(DocumentClient.open(file, { worker: false, signal: before.signal })).rejects.toThrow('Document closed')
    const during = new AbortController()
    const opening = DocumentClient.open(file, { worker: false, signal: during.signal })
    during.abort()
    await expect(opening).rejects.toThrow('Document closed')
    const after = new AbortController()
    const client = await DocumentClient.open(file, { worker: false, signal: after.signal })
    after.abort()
    expect(client.closed).toBe(false)
    expect(Array.isArray((await client.call('xlsx.block', 0, 0, 10)).rows)).toBe(true)
    client.destroy()
  })
})
