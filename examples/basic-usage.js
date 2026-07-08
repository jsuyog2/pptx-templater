/**
 * @fileoverview Basic usage example for pptx-templater.
 *
 * This example demonstrates the core workflow:
 *  1. Load a PPTX template
 *  2. Select specific slides
 *  3. Replace text placeholders
 *  4. Add a new slide
 *  5. Save to file and buffer
 *
 * To run: node examples/basic-usage.js
 */

const { PPTXTemplater } = require('../src/index.js')
const { existsSync } = require('fs')
const { resolve } = require('path')

async function PPTXTemplaterDemo() {
  const TEMPLATE_PATH_PPTX = resolve(__dirname, '../templates/sample.pptx')
  const OUTPUT_PATH_EXTRACTION = resolve(__dirname, '../templates/sample')
  const OUTPUT_PATH = resolve(__dirname, '../examples/output/basic-output.pptx')

  if (existsSync(TEMPLATE_PATH_PPTX)) {
    await PPTXTemplater.extractPptx(TEMPLATE_PATH_PPTX, OUTPUT_PATH_EXTRACTION, { overwrite: true })
  }

  await main(TEMPLATE_PATH_PPTX, OUTPUT_PATH)
}

async function PPTXTemplaterCorporateDemo() {
  const TEMPLATE_PATH_CORPORATE_PPTX = resolve(__dirname, '../templates/officePPT.pptx')
  const OUTPUT_PATH_CORPORATE = resolve(__dirname, '../examples/output/basic-corporate-output.pptx')
  const OUTPUT_PATH_CORPORATE_EXTRACTION = resolve(__dirname, '../templates/officePPT')

  if (existsSync(TEMPLATE_PATH_CORPORATE_PPTX)) {
    await PPTXTemplater.extractPptx(
      TEMPLATE_PATH_CORPORATE_PPTX,
      OUTPUT_PATH_CORPORATE_EXTRACTION,
      { overwrite: true }
    )
  }

  await main(TEMPLATE_PATH_CORPORATE_PPTX, OUTPUT_PATH_CORPORATE)
}

async function main(TEMPLATE_PATH_PPTX, OUTPUT_PATH) {
  try {
    // Check if template exists
    if (!existsSync(TEMPLATE_PATH_PPTX)) {
      console.log('ℹ Template file not found at:', TEMPLATE_PATH_PPTX)
      console.log('  Place a PPTX file at templates/sample.pptx to run this example.')
      console.log('\nRunning API demonstration without file I/O...\n')
      await demonstrateAPI()
      return
    }

    console.log('📂 Loading template:', TEMPLATE_PATH_PPTX)
    const ppt = await PPTXTemplater.load(TEMPLATE_PATH_PPTX)

    console.log(`📊 Loaded ${ppt.slideCount} slides`)

    // Step 1: Work on specific slide
    ppt.useSlide(1)

    // Step 2: Replace text placeholders
    ppt.replaceText({
      '{{title}}': 'Quarterly Business Review',
      '{{subtitle}}': 'Q1 2026 Results',
      '{{date}}': new Date().toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric',
      }),
      '{{company}}': 'Acme Corporation',
      '{{presenter}}': 'Jane Smith',
    })

    console.log('✅ Text replaced')

    ppt.updateChart('Chart', {
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [
        { name: 'Product A', values: [145, 210, 190, 250] },
        { name: 'Product B', values: [90, 130, 160, 200] },
        { name: 'Product C', values: [70, 85, 100, 120] },
      ],
    })
    ppt.updateChartTitle('Chart', 'Global SaaS\nRevenue (2026)')

    ppt.applyZOrder(1, [{ id: 'Shape', zIndex: 2 }])

    ppt.removeSlide(2)

    ppt.useSlide(2)
    ppt.replaceText({
      '{{title}}': 'New Title',
    })
    ppt.addHyperlink({
      text: 'Google',
      url: 'https://www.google.com',
    })

    ppt.addSlideLink({
      element: 'Hello New Title',
      sourceSlide: 2,
      targetSlide: 1,
    })
    const rows = await ppt.getTableRows('Table')
    console.log(rows)

    ppt.updateTable('Table', [
      ['Name', '', 'Role', 'Dept'],
      [
        'Alice',
        '',
        'The final implementation should allow users to build highly visual dashboards and reports.',
        'Platform',
      ],
      ['Bob', '', '     Designer', 'Product'],
    ])
    ppt.addTableRow('Table', [
      '',
      '',
      'Designer',
      'Value',
      'Product',
      {
        type: 'circle',
        color: '#0021dbff',
        width: 12,
        height: 12,
        position: 'center',
      },
    ])
    ppt.addTableRow('Table', [
      'Bob',
      '',
      'Designer',
      'Value',
      'Product',
      { type: 'circle', color: '#4CAF50', width: 12, height: 12, position: 'center' },
    ])

    ppt.removeTableRow('Table', 1)

    // merge all cells of first column with all cells of second col
    ppt.mergeCells('Table', 1, 0, 2, 0)
    ppt.mergeCells('Table', 1, 1, 2, 1)
    // ppt.mergeCells('Table', 1, 0, 1, 1)
    // ppt.mergeCells('Table', 2, 0, 2, 1)

    ppt.updateTable('Table2', [
      ['Name', '', 'Role', 'Dept'],
      ['Alice', '', 'Engineer', 'Platform'],
      ['Bob', { value: 'Alice', align: 'ctr' }, 'Designer', 'Product'],
    ])

    ppt.addCellShape('Table2', 1, 2, {
      type: 'rectangle',
      fill: '#10B981',
      x: 10,
      y: 15,
      width: 25,
      height: 15,
    })

    ppt.mergeCells('Table2', 0, 0, 0, 1)
    ppt.mergeCells('Table2', 1, 0, 1, 1)

    ppt.useSlide(3)

    ppt.updateChart('Label Chart', {
      categories: [''],
      series: [
        { name: 'Product A', values: [{ data: 145, label: '145\n(18.24%)' }] },
        { name: 'Product B', values: [{ data: 210, label: '210\n(26.42%)' }] },
        { name: 'Product C', values: [{ data: 190, label: '190\n(23.9%)' }] },
        { name: 'Product D', values: [] },
      ],
    })

    const positions = await ppt.getChartLabelPositions('Label Chart')
    positions.forEach((d, i) => {
      ppt.updateTextBoxPosition(`Series ${i + 1}`, {
        x: d.x - 3500000,
        y: d.y,
      })
    })

    ppt.useSlide(4)

    ppt.updateText('List', {
      list: ['Fast PPTX generation', 'OpenXML based', 'Chart updates', 'Table updates'],
    })

    ppt.useSlide(5)

    const TeamData = [
      [
        '',
        'Team A',
        'Bob 1',
        'Designer',
        'The final implementation should allow users to build highly visual dashboards and reports.',
        '40',
        'N',
        'A',
      ],
      [
        '',
        '',
        'Charlie 1',
        'Admin',
        'The final implementation should allow users to build highly visual dashboards and reports.',
        '30',
        'P',
        'B',
      ],
      [
        '',
        '',
        'David 1',
        'Developer',
        'The final implementation should allow users to build highly visual dashboards and reports.',
        '100',
        'P',
        'A',
      ],
      [
        '',
        'Team B',
        'User 1',
        'Tester',
        'The final implementation should allow users to build highly visual dashboards and reports.',
        '60',
        'N',
        'A',
      ],
      [
        '',
        '',
        'Alice 1',
        'Manager',
        'The final implementation should allow users to build highly visual dashboards and reports.',
        '85',
        'P',
        'B',
      ],
      [
        '',
        'Team A',
        'Bob',
        'Designer',
        'The final implementation should allow users to build highly visual dashboards and reports.',
        '40',
        'N',
        'A',
      ],
      [
        '',
        '',
        'Charlie',
        'Admin',
        'The final implementation should allow users to build highly visual dashboards and reports.',
        '30',
        'P',
        'B',
      ],
      [
        '',
        '',
        'David',
        'Developer',
        'The final implementation should allow users to build highly visual dashboards and reports.',
        '100',
        'P',
        'A',
      ],
      [
        '',
        'Team B',
        'User',
        'Tester',
        'The final implementation should allow users to build highly visual dashboards and reports.',
        '60',
        'N',
        'A',
      ],
      [
        '',
        '',
        'Alice',
        'Manager',
        'The final implementation should allow users to build highly visual dashboards and reports.',
        '85',
        'P',
        'B',
      ],
    ]
    const totalSlide = TeamData.length / 5
    let slideNumber = 5
    for (let index = 1; index <= totalSlide - 1; index++) {
      await ppt.duplicateSlide(slideNumber, slideNumber + index)
    }
    TeamData.forEach((element, i) => {
      if (i % 5 === 0 && i !== 0) {
        ppt.removeTableRow('Table', 1)
        ppt.useSlide(slideNumber + 1)
        slideNumber++
      }
      ppt.addTableRow('Table', element)
    })
    ppt.removeTableRow('Table', 1)

    ppt.useSlide(5)
    let row = 0
    TeamData.forEach((element, rowIndex) => {
      element.forEach((cell, colIndex) => {
        if (rowIndex % 5 === 0 && rowIndex !== 0) {
          row = 0
          ppt.useSlide(6)
        }
        if (colIndex === 7 && (cell === 'A' || cell === 'B')) {
          ppt.addCellShape('Table', row + 1, colIndex, {
            type: 'circle',
            fill: cell === 'A' ? '#10B981' : '#EF4444',
            width: 15,
            height: 15,
            position: 'center',
          })

          ppt.updateCell('Table', row + 1, colIndex, '')
          row++
        }
      })
    })

    ppt.mergeCells('Table', 1, 0, 3, 0)
    ppt.mergeCells('Table', 1, 1, 3, 1)
    ppt.mergeCells('Table', 4, 0, 5, 0)
    ppt.mergeCells('Table', 4, 1, 5, 1)

    ppt.alignShapeToCell('Team A Logo', 'Table', 1, 0)
    ppt.alignShapeToCell('Team B Logo', 'Table', 4, 0)

    ppt.useSlide(5)
    ppt.mergeCells('Table', 1, 0, 3, 0)
    ppt.mergeCells('Table', 1, 1, 3, 1)
    ppt.mergeCells('Table', 4, 0, 5, 0)
    ppt.mergeCells('Table', 4, 1, 5, 1)

    ppt.alignShapeToCell('Team A Logo', 'Table', 1, 0)
    ppt.alignShapeToCell('Team B Logo', 'Table', 4, 0)

    // Step 3: Save to file
    await ppt.saveToFile(OUTPUT_PATH)
    console.log('💾 Saved to:', OUTPUT_PATH)

    // Step 4: Also get as buffer (for API responses)
    const buffer = await ppt.toBuffer()
    console.log(`📦 Buffer size: ${(buffer.length / 1024).toFixed(1)} KB`)

    // Step 5: Validate
    const validation = ppt.validate()
    if (validation.valid) {
      console.log('✅ Validation passed')
    } else {
      console.warn('⚠ Validation warnings:', validation.errors)
    }

    console.log('\n🎉 Example completed successfully!')
  } catch (err) {
    console.error('❌ Error:', err.message)
    if (process.env.DEBUG) console.error(err.stack)
    process.exit(1)
  }
}

/**
 * Demonstrates the API structure without needing an actual PPTX file.
 */
async function demonstrateAPI() {
  console.log('=== API Usage Patterns ===\n')

  // Show the API shape
  console.log(`
// Load a template
const ppt = await PPTXTemplater.load('template.pptx');

// Select slides to operate on
ppt.useSlide(1);                     // Single slide
ppt.useSlide(1, 3, 5);              // Multiple slides
ppt.useAllSlides();                  // All slides

// Replace {{placeholder}} text
ppt.replaceText({
  '{{title}}': 'Quarterly Report',
  '{{year}}':  '2026',
});

// Update chart data
ppt.updateChart('sales-chart', {
  categories: ['Jan', 'Feb', 'Mar'],
  series: [{ name: 'Revenue', values: [120, 150, 180] }],
});

// Update table rows
ppt.updateTable('data-table', [
  ['Name',  'Role',       'Dept'],
  ['Alice', 'Engineer',   'Platform'],
  ['Bob',   'Designer',   'Product'],
]);

// Add hyperlinks
ppt.addHyperlink({ text: 'Visit Us', url: 'https://example.com' });

// Link slide numbers
ppt.linkSlideNumber({ slide: 1, targetSlide: 5 });

// Add new slides
ppt.addSlide({
  title: 'New Slide',
  elements: [{ type: 'text', value: 'Hello World' }],
});

// Clone and remove slides
ppt.cloneSlide(1);          // Duplicate slide 1
ppt.removeSlide(3);         // Remove slide 3
ppt.reorderSlides([3,1,2]); // Reorder

// Export
await ppt.saveToFile('./output/report.pptx');
const buffer = await ppt.toBuffer();
const stream = await ppt.toStream();
  `)

  console.log('=== End of API Demo ===')
}

PPTXTemplaterDemo()
PPTXTemplaterCorporateDemo()
