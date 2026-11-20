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
})
