import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

describe('Table Row/Column Auto-Adjustment Integration Tests', () => {
  const runTests = existsSync(FIXTURE_FILE)

  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  let ppt

  beforeAll(async () => {
    ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3) // Slide 3 contains table 'Table'
  })

  it('1. should adjust table row height with direct height value and auto mode', async () => {
    // Test fixed row height (e.g. height: 35 pt = 444500 EMUs)
    ppt.adjustTableRow('Table', {
      row: 1,
      height: 35,
    })

    // Test auto height calculation for row 2
    ppt.adjustTableRow('Table', {
      row: 2,
      auto: true,
    })

    // Test array of row adjustments
    ppt.adjustTableRow('Table', [
      { row: 0, height: 40 },
      { row: 1, auto: true },
    ])

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
    expect(buffer.length).toBeGreaterThan(0)
  })

  it('2. should adjust table column width with fixed width and auto mode', async () => {
    // Test fixed column width (e.g. width: 2.4 inches)
    ppt.adjustTableColumn('Table', {
      column: 0,
      width: 2.4,
    })

    // Test auto column width calculation
    ppt.adjustTableColumn('Table', {
      column: 1,
      auto: true,
    })

    // Test array of column adjustments
    ppt.adjustTableColumn('Table', [
      { column: 0, width: 3.1 },
      { column: 1, auto: true },
    ])

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
    expect(buffer.length).toBeGreaterThan(0)
  })

  it('3. should resize table by adjusting both rows and columns with adjustTable', async () => {
    // Format A: { auto: true }
    ppt.adjustTable('Table', {
      auto: true,
    })

    // Format B: { rows: 'auto', columns: 'auto' }
    ppt.adjustTable('Table', {
      rows: 'auto',
      columns: 'auto',
    })

    // Format C: { rows: { auto: true }, columns: { auto: true } }
    ppt.adjustTable('Table', {
      rows: {
        auto: true,
      },
      columns: {
        auto: true,
      },
    })

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
  })

  it('4. should handle tables with mixed text, wrapped text, icons/shapes, and merged cells', async () => {
    // Update table with data and cell shapes
    ppt.updateTable('Table', {
      rows: [
        { A: 'Executive Leader', V: 'Active Status', B: 100 },
        { A: 'Software Architect with long title that wraps lines', V: 'Pending Review', B: 250 },
      ],
      cellShapes: {
        V: () => ({
          type: 'circle',
          fill: '#10B981',
          width: 16,
          height: 16,
          position: 'left',
        }),
      },
    })

    // Merge cells to test merged region safety during adjustTable
    ppt.mergeCells('Table', 0, 0, 0, 1)

    // Adjust full table
    ppt.adjustTable('Table', { auto: true })

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
  })
})
