import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

describe('Universal Object Resize API Integration Tests (ppt.resize)', () => {
  const runTests = existsSync(FIXTURE_FILE)

  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  let ppt

  beforeAll(async () => {
    ppt = await PPTXTemplater.load(FIXTURE_FILE)
  })

  it('1. should support absolute resizing (width, height, width+height)', async () => {
    ppt.useSlide(1)

    // Absolute width on first element of slide 1
    ppt.resize('first', { width: 6.5 })

    // Absolute height on first element
    ppt.resize('first', { height: 1.5 })

    // Absolute width + height
    ppt.resize('first', { width: 7.0, height: 2.0 })

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
    expect(buffer.length).toBeGreaterThan(0)
  })

  it('2. should support relative resizing (increase and decrease)', async () => {
    ppt.useSlide(2)

    // Increase width
    ppt.resize('first', { increase: { width: 1.0 } })

    // Decrease height
    ppt.resize('first', { decrease: { height: 0.5 } })

    // Increase both
    ppt.resize('first', { increase: { width: 1.5, height: 0.8 } })

    // Decrease both
    ppt.resize('first', { decrease: { width: 0.3, height: 0.3 } })

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
  })

  it('3. should support percentage scaling (uniform scale and independent X/Y scale)', async () => {
    ppt.useSlide(3)

    // Uniform scale up by 20%
    ppt.resize('first', { scale: 1.2 })

    // Uniform scale down by 20%
    ppt.resize('first', { scale: 0.8 })

    // Independent scale (X * 1.1, Y * 0.9)
    ppt.resize('first', { scale: { width: 1.1, height: 0.9 } })

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
  })

  it('4. should correctly handle table resizing via universal resize API', async () => {
    ppt.useSlide(3) // Slide 3 contains table 'Table'

    // Resize table using increase
    ppt.resize('Table', { increase: { width: 2.0, height: 1.0 } })

    // Resize table using percentage scale
    ppt.resize('Table', { scale: 1.1 })

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
  })
})
