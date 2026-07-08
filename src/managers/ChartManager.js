/**
 * @fileoverview ChartManager - Updates chart data in PPTX slides.
 *
 * Charts in OpenXML PPTX:
 * ─────────────────────────────────────────────────────────────────
 * Charts are embedded as separate XML files referenced via relationships.
 * A chart on a slide appears as:
 *
 *   <p:graphicFrame>
 *     <p:nvGraphicFramePr>
 *       <p:cNvPr id="5" name="sales-chart"/>    ← chart name/ID
 *     </p:nvGraphicFramePr>
 *     <a:graphic>
 *       <a:graphicData uri="...chart">
 *         <c:chart r:id="rId5"/>                ← references chart XML
 *       </a:graphicData>
 *     </a:graphic>
 *   </p:graphicFrame>
 *
 * The chart XML at ppt/charts/chartN.xml contains:
 *  - c:chartSpace → root element
 *    - c:chart → chart metadata
 *      - c:plotArea → chart data
 *        - c:barChart / c:lineChart / c:pieChart / etc.
 *          - c:ser → each data series
 *            - c:idx / c:order → series index/order
 *            - c:tx → series name
 *            - c:cat → categories (X-axis)
 *            - c:val → values (Y-axis)
 *
 * Category Reference Types:
 *  - c:strRef  → string categories (labels)
 *  - c:numRef  → numeric categories
 *
 * Value Reference Types:
 *  - c:numRef  → numeric values (typical)
 *
 * Data is stored in both a formula (c:f) and a cache (c:strCache/c:numCache).
 * We update both to ensure compatibility with both cached and live data.
 */

const { createLogger } = require('../utils/logger.js')
const { ChartNotFoundError } = require('../utils/errors.js')
const { REL_TYPES } = require('./RelationshipManager.js')
const { ChartWorkbookUpdater } = require('./charts/ChartWorkbookUpdater.js')
const { ChartCacheGenerator } = require('./charts/ChartCacheGenerator.js')

const JSZip = require('jszip')

const logger = createLogger('ChartManager')

/**
 * Supported chart types and their XML element names.
 */
const CHART_TYPE_MAP = {
  bar: 'c:barChart',
  line: 'c:lineChart',
  pie: 'c:pieChart',
  area: 'c:areaChart',
  scatter: 'c:scatterChart',
  doughnut: 'c:doughnutChart',
  radar: 'c:radarChart',
  bubble: 'c:bubbleChart',
  stock: 'c:stockChart',
}

/**
 * @class ChartManager
 * @description Handles chart data updates by directly manipulating chart XML files.
 *
 * Unlike high-level charting libraries, this manager edits the raw OpenXML
 * chart structure, allowing full control while preserving styles and themes.
 */
class ChartManager {
  /** @private @type {XMLParser} */
  #xmlParser
  /** @private @type {ZipManager} */
  #zipManager

  /**
   * Cache of chart ZIP paths: maps chartName → { zipPath, slideIndex }
   * @private @type {Map<string, { zipPath: string, slideIndex: number }>}
   */
  #chartRegistry = new Map()

  /**
   * Promise queue for sequential execution per chart ZIP path.
   * @private @type {Map<string, Promise>}
   */
  #chartQueues = new Map()

  /**
   * Promise queue for sequential execution per slide ZIP path.
   * @private @type {Map<string, Promise>}
   */
  #slideQueues = new Map()

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser
  }

  async initialize(zipManager) {
    this.#zipManager = zipManager
    const chartFiles = zipManager
      .listFiles('ppt/charts/')
      .filter(f => f.endsWith('.xml') && !f.includes('_rels'))

    for (const chartPath of chartFiles) {
      // Chart name is inferred from file name
      const chartName = chartPath.split('/').pop().replace('.xml', '')
      this.#chartRegistry.set(chartName, { zipPath: chartPath, slideIndex: null })
    }

    logger.debug(`Found ${chartFiles.length} chart file(s)`)
  }

  /**
   * Updates chart data for a named chart within a slide.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string} chartId - Chart name (from shape's cNvPr name attribute) or chart file name.
   * @param {ChartData} data - New chart data.
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @throws {ChartNotFoundError} If the chart cannot be found.
   */
  updateChart(slideIndex, chartId, data, slideManager, relationshipManager) {
    this.#validateChartData(data)
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)

    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }

    logger.debug(`Updating chart "${chartId}" at ${chartInfo.zipPath}`)
    this.#updateChartXml(
      chartInfo.zipPath,
      data,
      relationshipManager,
      slideIndex,
      chartId,
      slideManager
    )
  }

  /**
   * Returns all chart info objects for a given slide.
   *
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Array<{name: string, zipPath: string}>}
   */
  getChartsInSlide(slideIndex, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const rels = relationshipManager.getRelationshipsByType(slideInfo.zipPath, REL_TYPES.CHART)

    return rels.map(rel => {
      const chartPath = relationshipManager.resolveTarget(slideInfo.zipPath, rel.target)
      return { rId: rel.id, zipPath: chartPath }
    })
  }

  /**
   * Finds a chart in a slide by name/ID.
   * @private
   *
   * @param {number} slideIndex
   * @param {string} chartId - Shape name or rId.
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {{ zipPath: string }|null}
   */
  findChartInSlide(slideIndex, chartId, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideXml = slideManager.getSlideXml(slideIndex)

    // Strategy 1: Look for shape with matching name (cNvPr name attribute)
    const shapeNamePattern = new RegExp(
      `<p:cNvPr[^>]*name="${chartId}"[^>]*>(?:.*?)<c:chart[^>]*r:id="(rId\\d+)"`,
      's'
    )
    const rIdMatch = shapeNamePattern.exec(slideXml)

    // Strategy 2: Find graphicFrame shapes and match chart rIds
    if (!rIdMatch) {
      const chartRIdPattern = /<c:chart[^>]*r:id="(rId\d+)"/g
      let chartMatch
      while ((chartMatch = chartRIdPattern.exec(slideXml)) !== null) {
        const rId = chartMatch[1]
        const rel = relationshipManager.getRelationshipById(slideInfo.zipPath, rId)
        if (rel) {
          const chartPath = relationshipManager.resolveTarget(slideInfo.zipPath, rel.target)
          // Check if chart file name matches chartId
          if (chartPath.includes(chartId) || rel.id === chartId) {
            return { zipPath: chartPath }
          }
          // For the first chart found, if chartId looks like a chart file name
          if (chartId.startsWith('chart')) {
            return { zipPath: chartPath }
          }
        }
      }
    }

    if (rIdMatch) {
      const rId = rIdMatch[1]
      const rel = relationshipManager.getRelationshipById(slideInfo.zipPath, rId)
      if (rel) {
        const chartPath = relationshipManager.resolveTarget(slideInfo.zipPath, rel.target)
        return { zipPath: chartPath }
      }
    }

    // Strategy 3: Direct chart registry lookup
    if (this.#chartRegistry.has(chartId)) {
      return this.#chartRegistry.get(chartId)
    }

    // Strategy 4: Try chartN naming convention
    const chartPath = `ppt/charts/${chartId}.xml`
    if (this.#zipManager.hasFile(chartPath)) {
      return { zipPath: chartPath }
    }

    return null
  }

  /**
   * Updates the chart XML file with new data.
   * Preserves all styling, themes, and chart configuration.
   *
   * @private
   * @param {string} chartZipPath - ZIP path to the chart XML file.
   * @param {ChartData} data - New chart data.
   * @param {RelationshipManager} relationshipManager
   */
  #enqueueChartTask(chartZipPath, taskFn) {
    if (!this.#zipManager.hasFile(chartZipPath)) {
      throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)
    }
    const queue = this.#chartQueues.get(chartZipPath) || Promise.resolve()
    const nextTask = queue.then(() => taskFn())
    this.#chartQueues.set(chartZipPath, nextTask)
    this.#zipManager.addPendingPromise(nextTask)
  }

  #updateChartXml(chartZipPath, data, relationshipManager, slideIndex, chartId, slideManager) {
    this.#enqueueChartTask(chartZipPath, () =>
      this.updateChartAsync(
        chartZipPath,
        data,
        relationshipManager,
        slideIndex,
        chartId,
        slideManager
      )
    )
  }

  /**
   * Async version of chart update — updates XML and embedded workbook.
   *
   * @param {string} chartZipPath
   * @param {ChartData} data
   * @param {RelationshipManager} relationshipManager
   * @returns {Promise<void>}
   */
  async updateChartAsync(
    chartZipPath,
    data,
    relationshipManager,
    slideIndex,
    chartId,
    slideManager
  ) {
    // 1. Read Chart XML
    const xml = await this.#zipManager.readFile(chartZipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)

    // 2. Normalize and split the chart data
    const normalized = this.#normalizeChartData(data)
    const cleanNumericData = normalized.cleanData
    const seriesLabels = normalized.labels

    // 3. Apply Chart XML Updates
    let updatedXml = this.#applyChartData(xml, cleanNumericData, chartZipPath)
    if (data.title !== undefined) {
      updatedXml = require('./charts/ChartCacheGenerator.js').ChartCacheGenerator.updateTitle(
        updatedXml,
        data.title
      )
    }
    this.#zipManager.writeFile(chartZipPath, updatedXml)

    // 4. Find and Update Embedded Workbook
    if (relationshipManager) {
      const rels = relationshipManager.getRelationshipsByType(chartZipPath, REL_TYPES.PACKAGE)
      for (const rel of rels) {
        const xlsxPath = relationshipManager.resolveTarget(chartZipPath, rel.target)
        const buffer = await this.#zipManager.readBinaryFile(xlsxPath)
        if (buffer) {
          logger.debug(`Found embedded workbook: ${xlsxPath}`)
          const updatedXlsx = await ChartWorkbookUpdater.updateWorkbook(buffer, cleanNumericData)
          if (updatedXlsx) {
            logger.debug(`Writing updated workbook to: ${xlsxPath}, size: ${updatedXlsx.length}`)
            this.#zipManager.writeBinaryFile(xlsxPath, updatedXlsx)
          }
        } else {
          logger.debug(`Could not find workbook at: ${xlsxPath}`)
        }
      }
    }

    // 5. Apply custom data labels if present or if showSeriesNameInBar is enabled
    for (let i = 0; i < seriesLabels.length; i++) {
      const labels = seriesLabels[i]
      const ser = cleanNumericData.series[i]
      const showSerName =
        ser.showSeriesNameInBar !== undefined ? ser.showSeriesNameInBar : data.showSeriesNameInBar

      if ((labels && labels.some(l => l !== undefined)) || showSerName) {
        const labelOptions = {
          series: i,
          ...(labels ? { labels: labels.map(l => (l === undefined ? '' : String(l))) } : {}),
          showSeriesNameInBar: !!showSerName,
        }
        await this.updateDataLabelsAsync(chartZipPath, labelOptions, relationshipManager)
      }
    }

    // 6. Apply series name labels if enabled
    if (
      data.seriesNameLabels &&
      slideIndex !== undefined &&
      chartId !== undefined &&
      slideManager !== undefined
    ) {
      const finalXml = await this.#zipManager.readFile(chartZipPath)
      await this.#applySeriesNameLabels(
        slideIndex,
        chartId,
        data,
        finalXml,
        slideManager,
        relationshipManager
      )
    }
  }

  /**
   * Applies new chart data to the chart XML string.
   *
   * @private
   * @param {string} xml - Original chart XML.
   * @param {ChartData} data - New data to apply.
   * @param {string} context - For error messages.
   * @returns {string} Updated XML.
   */
  #applyChartData(xml, data, context) {
    const { categories, series } = data

    // Detect chart type
    const chartType = this.#detectChartType(xml)
    logger.debug(`Updating ${chartType} chart at ${context}`)

    let updatedXml = xml

    if (series && series.length > 0) {
      updatedXml = ChartCacheGenerator.appendDynamicSeries(updatedXml, series.length)
    }

    if (categories && categories.length > 0) {
      updatedXml = ChartCacheGenerator.updateCategories(updatedXml, categories)
    }

    if (series && series.length > 0) {
      updatedXml = ChartCacheGenerator.updateSeries(
        updatedXml,
        series,
        categories ? categories.length : null
      )
    }

    return updatedXml
  }

  /**
   * Updates only chart categories.
   */
  updateChartCategories(slideIndex, chartId, categories, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    this.#enqueueChartTask(chartInfo.zipPath, () =>
      this.updateChartCategoriesAsync(chartInfo.zipPath, categories, relationshipManager)
    )
  }

  async updateChartCategoriesAsync(chartZipPath, categories, relationshipManager) {
    const xml = await this.#zipManager.readFile(chartZipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)
    const data = this.#extractChartData(xml)
    data.categories = categories
    await this.updateChartAsync(chartZipPath, data, relationshipManager)
  }

  /**
   * Replaces a specific chart series.
   */
  replaceChartSeries(
    slideIndex,
    chartId,
    seriesIndex,
    newSeriesData,
    slideManager,
    relationshipManager
  ) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    this.#enqueueChartTask(chartInfo.zipPath, () =>
      this.replaceChartSeriesAsync(
        chartInfo.zipPath,
        seriesIndex,
        newSeriesData,
        relationshipManager
      )
    )
  }

  async replaceChartSeriesAsync(chartZipPath, seriesIndex, newSeriesData, relationshipManager) {
    const xml = await this.#zipManager.readFile(chartZipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)
    const data = this.#extractChartData(xml)
    data.series[seriesIndex] = newSeriesData
    await this.updateChartAsync(chartZipPath, data, relationshipManager)
  }

  /**
   * Updates the chart title.
   */
  updateChartTitle(slideIndex, chartId, title, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    this.#enqueueChartTask(chartInfo.zipPath, () =>
      this.updateChartTitleAsync(chartInfo.zipPath, title)
    )
  }

  async updateChartTitleAsync(chartZipPath, title) {
    const xml = await this.#zipManager.readFile(chartZipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)
    const updatedXml = ChartCacheGenerator.updateTitle(xml, title)
    this.#zipManager.writeFile(chartZipPath, updatedXml)
  }

  /**
   * Helper to extract data from chart XML.
   */
  #extractChartData(xml) {
    const categories = []
    const series = []

    const catMatch = /<c:cat>([\s\S]*?)<\/c:cat>/.exec(xml)
    if (catMatch) {
      const catXml = catMatch[1]
      const ptPattern = /<c:pt idx="\d+">\s*<c:v>([^<]*)<\/c:v>/g
      let match
      while ((match = ptPattern.exec(catXml)) !== null) {
        categories.push(match[1])
      }
    }

    const serPattern = /<c:ser>([\s\S]*?)<\/c:ser>/g
    let serMatch
    let idx = 0
    while ((serMatch = serPattern.exec(xml)) !== null) {
      const serXml = serMatch[1]

      let name = `Series ${idx + 1}`
      const txMatch = /<c:tx>([\s\S]*?)<\/c:tx>/.exec(serXml)
      if (txMatch) {
        const nameValMatch = /<c:v>([^<]*)<\/c:v>/.exec(txMatch[1])
        if (nameValMatch) name = nameValMatch[1]
      }

      const values = []
      const valMatch = /<c:val>([\s\S]*?)<\/c:val>/.exec(serXml)
      if (valMatch) {
        const ptPattern = /<c:pt idx="\d+">\s*<c:v>([^<]*)<\/c:v>/g
        let match
        while ((match = ptPattern.exec(valMatch[1])) !== null) {
          values.push(Number(match[1]) || 0)
        }
      }

      series.push({ name, values })
      idx++
    }

    return { categories, series }
  }

  /**
   * Detects the chart type from its XML.
   * @private
   * @param {string} xml
   * @returns {string} Chart type name.
   */
  #detectChartType(xml) {
    for (const [name, element] of Object.entries(CHART_TYPE_MAP)) {
      if (xml.includes(element)) return name
    }
    return 'unknown'
  }

  updateDataLabels(slideIndex, chartId, options, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    this.#enqueueChartTask(chartInfo.zipPath, () =>
      this.updateDataLabelsAsync(chartInfo.zipPath, options, relationshipManager)
    )
  }

  async updateDataLabelsAsync(chartZipPath, options, relationshipManager) {
    const xml = await this.#zipManager.readFile(chartZipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)

    const chartData = this.#extractChartData(xml)
    const { categories, series } = chartData

    const seriesIndex = options.series !== undefined ? options.series : 0
    if (seriesIndex >= series.length) {
      throw new Error(
        `Series index ${seriesIndex} out of bounds (chart has ${series.length} series)`
      )
    }

    const seriesData = series[seriesIndex]

    let resolvedLabels = null
    let labelsFromCells = options.labelsFromCells

    const hasCustomLabels =
      options.labels || options.labelMap || options.template || options.labelsFromCells

    if (hasCustomLabels) {
      const values = seriesData.values || []
      const sumValues = values.reduce((sum, v) => sum + (Number(v) || 0), 0)
      const seriesName = seriesData.name || ''

      const pointsCount = Math.max(
        categories.length,
        values.length,
        options.labels ? options.labels.length : 0
      )

      if (options.labelsFromCells && !options.labels && !options.template && !options.labelMap) {
        if (relationshipManager) {
          const rels = relationshipManager.getRelationshipsByType(chartZipPath, REL_TYPES.PACKAGE)
          for (const rel of rels) {
            const xlsxPath = relationshipManager.resolveTarget(chartZipPath, rel.target)
            const buffer = await this.#zipManager.readBinaryFile(xlsxPath)
            if (buffer) {
              const zip = await JSZip.loadAsync(buffer)
              const sheetFile = zip.file('xl/worksheets/sheet1.xml')
              if (sheetFile) {
                const sheetXml = await sheetFile.async('text')
                const sharedStrings = await ChartWorkbookUpdater.getSharedStrings(zip)
                const cells = ChartWorkbookUpdater.parseWorksheetCells(sheetXml, sharedStrings)

                const range = ChartWorkbookUpdater.parseCellRange(options.labelsFromCells)
                const startColNum = ChartWorkbookUpdater.colLetterToNum(range.startCol)

                resolvedLabels = []
                for (let i = 0; i < pointsCount; i++) {
                  let cellRef
                  if (range.startRow === range.endRow) {
                    cellRef = `${ChartWorkbookUpdater.numToColLetter(startColNum + i)}${range.startRow}`
                  } else {
                    cellRef = `${range.startCol}${range.startRow + i}`
                  }
                  resolvedLabels.push(cells[cellRef] !== undefined ? String(cells[cellRef]) : '')
                }
              }
            }
          }
        }
      } else {
        resolvedLabels = []
        for (let i = 0; i < pointsCount; i++) {
          const cat = categories[i] !== undefined ? String(categories[i]) : ''
          const val = values[i] !== undefined ? values[i] : ''

          let pct = 0
          if (sumValues > 0 && val !== '') {
            pct = Math.round((Number(val) / sumValues) * 100)
          }

          let customLabel = ''
          if (options.labels && options.labels[i] !== undefined) {
            customLabel = String(options.labels[i])
          } else if (options.labelMap && cat && options.labelMap[cat] !== undefined) {
            customLabel = String(options.labelMap[cat])
          }

          let textContent = customLabel
          if (options.template) {
            textContent = options.template
              .replace(/{category}/g, cat)
              .replace(/{value}/g, String(val))
              .replace(/{percentage}/g, String(pct))
              .replace(/{series}/g, seriesName)
              .replace(/{customLabel}/g, customLabel)
          }
          resolvedLabels.push(textContent)
        }
      }
    }

    if (resolvedLabels && !labelsFromCells) {
      const colLetter = ChartWorkbookUpdater.getColumnLetter(1 + series.length + seriesIndex)
      labelsFromCells = `Sheet1!$${colLetter}$2:$${colLetter}$${1 + resolvedLabels.length}`
    }

    const resolvedOptions = {
      ...options,
      labels: resolvedLabels || options.labels,
      labelsFromCells,
    }

    const updatedXml = ChartCacheGenerator.updateDataLabelsInXml(
      xml,
      seriesIndex,
      resolvedOptions,
      categories,
      seriesData
    )
    this.#zipManager.writeFile(chartZipPath, updatedXml)

    if (relationshipManager && (resolvedOptions.labels || resolvedOptions.labelsFromCells)) {
      const rels = relationshipManager.getRelationshipsByType(chartZipPath, REL_TYPES.PACKAGE)
      for (const rel of rels) {
        const xlsxPath = relationshipManager.resolveTarget(chartZipPath, rel.target)
        const buffer = await this.#zipManager.readBinaryFile(xlsxPath)
        if (buffer) {
          const workbookData = {
            categories,
            series: series.map((ser, idx) => {
              if (idx === seriesIndex) {
                return {
                  ...ser,
                  labels: resolvedOptions.labels,
                  labelsFromCells: resolvedOptions.labelsFromCells,
                }
              }
              return ser
            }),
          }
          const updatedXlsx = await ChartWorkbookUpdater.updateWorkbook(buffer, workbookData)
          if (updatedXlsx) {
            this.#zipManager.writeBinaryFile(xlsxPath, updatedXlsx)
          }
        }
      }
    }
  }

  async getDataLabels(slideIndex, chartId, options, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    const queue = this.#chartQueues.get(chartInfo.zipPath) || Promise.resolve()
    await queue
    const xml = await this.#zipManager.readFile(chartInfo.zipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartInfo.zipPath}`)
    const seriesIndex = options && options.series !== undefined ? options.series : 0
    return ChartCacheGenerator.getDataLabelsFromXml(xml, seriesIndex)
  }

  getChartType(slideIndex, chartId, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) return 'unknown'
    const cachedXml = this.#zipManager.readCachedFile(chartInfo.zipPath)
    if (cachedXml) {
      return this.#detectChartType(cachedXml)
    }
    return 'unknown'
  }

  async getChartTypeAsync(slideIndex, chartId, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) return 'unknown'
    const queue = this.#chartQueues.get(chartInfo.zipPath) || Promise.resolve()
    await queue
    const xml = await this.#zipManager.readFile(chartInfo.zipPath)
    if (!xml) return 'unknown'
    return this.#detectChartType(xml)
  }

  #validateChartData(data) {
    const { categories, series } = data
    if (data.title !== undefined && typeof data.title !== 'string') {
      throw new Error('Chart title must be a string')
    }
    if (!series || series.length === 0) return

    // Series lengths remain consistent (if categories exist, check against length of categories)
    const expectedLen = categories
      ? categories.length
      : series[0].values
        ? series[0].values.length
        : 0

    for (const ser of series) {
      const name = ser.name || 'Unknown'
      if (!ser.values) {
        ser.values = []
      }

      // Check if there is any label in the series to determine padding style
      let seriesHasLabels = false
      for (const val of ser.values) {
        if (typeof val === 'object' && val !== null && val.label !== undefined) {
          seriesHasLabels = true
          break
        }
      }

      // Pad or truncate values to match expectedLen
      if (ser.values.length < expectedLen) {
        while (ser.values.length < expectedLen) {
          if (seriesHasLabels) {
            ser.values.push({ value: null, label: '' })
          } else {
            ser.values.push(null)
          }
        }
      } else if (ser.values.length > expectedLen) {
        ser.values = ser.values.slice(0, expectedLen)
      }

      const len = ser.values.length

      // Check values inside the series
      let hasLabels = false
      let labelCount = 0
      for (const val of ser.values) {
        if (typeof val === 'object' && val !== null) {
          const numVal = val.value !== undefined ? val.value : val.data
          // Data values remain numeric or null
          if (numVal !== null && (typeof numVal !== 'number' || isNaN(numVal))) {
            throw new Error(`Data value must be numeric in series ${name}`)
          }
          if (val.label !== undefined) {
            hasLabels = true
            labelCount++
            // Labels are strings
            if (typeof val.label !== 'string') {
              throw new Error(`Label must be a string in series ${name}`)
            }
          }
        } else {
          // Data values remain numeric (primitive value) or null
          if (val !== null && (typeof val !== 'number' || isNaN(val))) {
            throw new Error(`Data value must be numeric in series ${name}`)
          }
        }
      }

      // Label count matches value count
      if (hasLabels && labelCount !== len) {
        throw new Error(`Label count mismatch for series ${name}`)
      }
    }
  }

  #normalizeChartData(data) {
    const cleanSeries = []
    const seriesLabels = []

    if (data.series) {
      data.series.forEach(ser => {
        const cleanValues = []
        const labels = []
        let hasLabel = false

        if (ser.values) {
          ser.values.forEach(v => {
            if (typeof v === 'object' && v !== null) {
              const val = v.value !== undefined ? v.value : v.data !== undefined ? v.data : null
              cleanValues.push(val)
              labels.push(v.label)
              if (v.label !== undefined) hasLabel = true
            } else {
              cleanValues.push(v === null || v === undefined ? null : Number(v) || 0)
              labels.push(undefined)
            }
          })
        }

        cleanSeries.push({
          ...ser,
          values: cleanValues,
        })
        seriesLabels.push(hasLabel ? labels : null)
      })
    }

    return {
      cleanData: {
        ...data,
        series: cleanSeries,
      },
      labels: seriesLabels,
    }
  }

  async validateChartLabels(slideIndex, chartId, options, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    const queue = this.#chartQueues.get(chartInfo.zipPath) || Promise.resolve()
    await queue
    const xml = await this.#zipManager.readFile(chartInfo.zipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartInfo.zipPath}`)

    const { ValidationEngine } = require('../core/ValidationEngine')
    return ValidationEngine.validateChartLabels(xml, options)
  }

  async validateSeriesNameLabels(slideIndex, chartId, options, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    const queue = this.#chartQueues.get(chartInfo.zipPath) || Promise.resolve()
    await queue
    const xml = await this.#zipManager.readFile(chartInfo.zipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartInfo.zipPath}`)

    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideXml = await this.#zipManager.readFile(slideInfo.zipPath)
    const xfrm = this.#findChartCoordinates(
      slideXml,
      chartId,
      relationshipManager,
      slideInfo.zipPath
    )

    const { ValidationEngine } = require('../core/ValidationEngine')
    return ValidationEngine.validateSeriesNameLabels(xml, xfrm, options)
  }

  #findChartCoordinates(slideXml, chartId, relationshipManager, slideZipPath) {
    const gfPattern = /<p:graphicFrame>([\s\S]*?)<\/p:graphicFrame>/g
    let match
    while ((match = gfPattern.exec(slideXml)) !== null) {
      const gfContent = match[0]
      const nameMatch = /<p:cNvPr[^>]*name="([^"]+)"/.exec(gfContent)
      const idMatch = /<p:cNvPr[^>]*id="([^"]+)"/.exec(gfContent)

      let isMatch = false
      if (nameMatch && nameMatch[1] === chartId) {
        isMatch = true
      } else if (idMatch && idMatch[1] === chartId) {
        isMatch = true
      } else {
        const chartRIdMatch = /<c:chart[^>]*r:id="([^"]+)"/.exec(gfContent)
        if (chartRIdMatch) {
          const rId = chartRIdMatch[1]
          const rel = relationshipManager.getRelationshipById(slideZipPath, rId)
          if (rel) {
            const chartPath = relationshipManager.resolveTarget(slideZipPath, rel.target)
            if (chartPath.includes(chartId) || rel.id === chartId) {
              isMatch = true
            }
          }
        }
      }

      if (isMatch) {
        const xfrmMatch = /<p:xfrm[^>]*>([\s\S]*?)<\/p:xfrm>/.exec(gfContent)
        if (xfrmMatch) {
          const offMatch = /<a:off\s+x="(\d+)"\s+y="(\d+)"\/>/.exec(xfrmMatch[1])
          const extMatch = /<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\/>/.exec(xfrmMatch[1])
          if (offMatch && extMatch) {
            return {
              left: parseInt(offMatch[1], 10),
              top: parseInt(offMatch[2], 10),
              width: parseInt(extMatch[1], 10),
              height: parseInt(extMatch[2], 10),
            }
          }
        }
      }
    }
    return null
  }

  #parsePlotAreaLayout(chartXml) {
    const plotAreaMatch = /<c:plotArea>([\s\S]*?)<\/c:plotArea>/.exec(chartXml)
    if (plotAreaMatch) {
      const layoutMatch = /<c:layout>([\s\S]*?)<\/c:layout>/.exec(plotAreaMatch[1])
      if (layoutMatch) {
        const xMatch = /<c:x\s+val="([^"]+)"\/>/.exec(layoutMatch[1])
        const yMatch = /<c:y\s+val="([^"]+)"\/>/.exec(layoutMatch[1])
        const wMatch = /<c:w\s+val="([^"]+)"\/>/.exec(layoutMatch[1])
        const hMatch = /<c:h\s+val="([^"]+)"\/>/.exec(layoutMatch[1])

        return {
          x: xMatch ? parseFloat(xMatch[1]) : 0.1,
          y: yMatch ? parseFloat(yMatch[1]) : 0.1,
          w: wMatch ? parseFloat(wMatch[1]) : 0.8,
          h: hMatch ? parseFloat(hMatch[1]) : 0.8,
        }
      }
    }
    return { x: 0.1, y: 0.1, w: 0.8, h: 0.8 }
  }

  #parseChartTypeAndGrouping(chartXml) {
    let dir = 'col'
    let grouping = 'clustered'
    let chartType = 'unknown'

    if (chartXml.includes('<c:barChart>') || chartXml.includes('c:barChart')) {
      chartType = 'bar'
      const barDirMatch = /<c:barDir\s+val="([^"]+)"\/>/.exec(chartXml)
      if (barDirMatch) dir = barDirMatch[1]

      const groupingMatch = /<c:grouping\s+val="([^"]+)"\/>/.exec(chartXml)
      if (groupingMatch) grouping = groupingMatch[1]
    }

    return { chartType, dir, grouping }
  }

  #resolveAxisMax(grouping, seriesValues, categoriesCount, seriesCount, chartXml) {
    const maxMatch = /<c:max\s+val="([^"]+)"\/>/.exec(chartXml)
    if (maxMatch) {
      return parseFloat(maxMatch[1])
    }

    if (grouping === 'stacked' || grouping === 'percentStacked') {
      const sums = []
      for (let c = 0; c < categoriesCount; c++) {
        let sum = 0
        for (let s = 0; s < seriesCount; s++) {
          sum += Math.abs(Number(seriesValues[s]?.[c]) || 0)
        }
        sums.push(sum)
      }
      const maxSum = Math.max(...sums, 1)
      return this.#getAxisMax(maxSum)
    } else {
      let maxVal = 1
      for (let s = 0; s < seriesCount; s++) {
        for (let c = 0; c < categoriesCount; c++) {
          maxVal = Math.max(maxVal, Math.abs(Number(seriesValues[s]?.[c]) || 0))
        }
      }
      return this.#getAxisMax(maxVal)
    }
  }

  #resolveSegmentsGeometry(geom) {
    const segments = []
    const fontSize = 10
    const emuPerPt = 12700
    const {
      plotLeft,
      plotTop,
      plotWidth,
      plotHeight,
      dir,
      grouping,
      categories,
      seriesCount,
      categoriesCount,
      seriesValues,
      seriesNames,
      axisMax,
    } = geom

    for (let c = 0; c < categoriesCount; c++) {
      const categoryName = categories[c] !== undefined ? String(categories[c]) : `Category ${c + 1}`

      if (dir === 'col') {
        const colWidth = plotWidth / Math.max(1, categoriesCount)
        const colLeftX = plotLeft + c * colWidth
        const colCenterX = colLeftX + colWidth / 2

        if (grouping === 'stacked' || grouping === 'percentStacked') {
          let categorySum = 0
          for (let s = 0; s < seriesCount; s++) {
            categorySum += Math.abs(Number(seriesValues[s]?.[c]) || 0)
          }
          if (categorySum === 0) categorySum = 1

          const barWidth = colWidth * 0.6
          const barLeft = colCenterX - barWidth / 2

          let cumulative = 0
          for (let s = 0; s < seriesCount; s++) {
            const val = Number(seriesValues[s]?.[c]) || 0
            const absVal = Math.abs(val)
            const nextCumulative = cumulative + absVal

            let segBottomY, segTopY
            if (grouping === 'percentStacked') {
              segBottomY = plotTop + plotHeight - plotHeight * (cumulative / categorySum)
              segTopY = plotTop + plotHeight - plotHeight * (nextCumulative / categorySum)
            } else {
              segBottomY = plotTop + plotHeight - plotHeight * (cumulative / axisMax)
              segTopY = plotTop + plotHeight - plotHeight * (nextCumulative / axisMax)
            }

            const segHeight = Math.max(0, segBottomY - segTopY)
            const segCenterY = (segBottomY + segTopY) / 2
            const seriesName = seriesNames[s] || `Series ${s + 1}`

            const labelText = String(val)
            const lblWidth = labelText.length * fontSize * 0.55 * emuPerPt
            const lblHeight = fontSize * 1.2 * emuPerPt

            segments.push({
              series: seriesName,
              category: categoryName,
              seriesIndex: s,
              categoryIndex: c,
              value: val,
              bar: {
                x: Math.round(barLeft),
                y: Math.round(segTopY),
                width: Math.round(barWidth),
                height: Math.round(segHeight),
              },
              label: {
                x: Math.round(colCenterX - lblWidth / 2),
                y: Math.round(segCenterY - lblHeight / 2),
                width: Math.round(lblWidth),
                height: Math.round(lblHeight),
              },
            })

            cumulative = nextCumulative
          }
        } else {
          const slotWidth = colWidth / Math.max(1, seriesCount)
          const barWidth = slotWidth * 0.8

          for (let s = 0; s < seriesCount; s++) {
            const val = Number(seriesValues[s]?.[c]) || 0
            const slotLeftX = colLeftX + s * slotWidth
            const slotCenterX = slotLeftX + slotWidth / 2
            const barLeft = slotCenterX - barWidth / 2

            const barHeight = plotHeight * (Math.abs(val) / axisMax)
            let barTopY, barBottomY
            if (val >= 0) {
              barTopY = plotTop + plotHeight - barHeight
              barBottomY = plotTop + plotHeight
            } else {
              barTopY = plotTop + plotHeight
              barBottomY = plotTop + plotHeight + barHeight
            }

            const seriesName = seriesNames[s] || `Series ${s + 1}`
            const labelText = String(val)
            const lblWidth = labelText.length * fontSize * 0.55 * emuPerPt
            const lblHeight = fontSize * 1.2 * emuPerPt
            const segCenterY = (barTopY + barBottomY) / 2

            segments.push({
              series: seriesName,
              category: categoryName,
              seriesIndex: s,
              categoryIndex: c,
              value: val,
              bar: {
                x: Math.round(barLeft),
                y: Math.round(barTopY),
                width: Math.round(barWidth),
                height: Math.round(barBottomY - barTopY),
              },
              label: {
                x: Math.round(slotCenterX - lblWidth / 2),
                y: Math.round(segCenterY - lblHeight / 2),
                width: Math.round(lblWidth),
                height: Math.round(lblHeight),
              },
            })
          }
        }
      } else {
        const rowHeight = plotHeight / Math.max(1, categoriesCount)
        const catTopY = plotTop + c * rowHeight
        const catCenterY = catTopY + rowHeight / 2

        if (grouping === 'stacked' || grouping === 'percentStacked') {
          let categorySum = 0
          for (let s = 0; s < seriesCount; s++) {
            categorySum += Math.abs(Number(seriesValues[s]?.[c]) || 0)
          }
          if (categorySum === 0) categorySum = 1

          const barHeight = rowHeight * 0.6
          const barTop = catCenterY - barHeight / 2

          let cumulative = 0
          for (let s = 0; s < seriesCount; s++) {
            const val = Number(seriesValues[s]?.[c]) || 0
            const absVal = Math.abs(val)
            const nextCumulative = cumulative + absVal

            let segLeftX, segRightX
            if (grouping === 'percentStacked') {
              segLeftX = plotLeft + plotWidth * (cumulative / categorySum)
              segRightX = plotLeft + plotWidth * (nextCumulative / categorySum)
            } else {
              segLeftX = plotLeft + plotWidth * (cumulative / axisMax)
              segRightX = plotLeft + plotWidth * (nextCumulative / axisMax)
            }

            const segWidth = Math.max(0, segRightX - segLeftX)
            const segCenterX = (segLeftX + segRightX) / 2
            const seriesName = seriesNames[s] || `Series ${s + 1}`

            const labelText = String(val)
            const lblWidth = labelText.length * fontSize * 0.55 * emuPerPt
            const lblHeight = fontSize * 1.2 * emuPerPt

            segments.push({
              series: seriesName,
              category: categoryName,
              seriesIndex: s,
              categoryIndex: c,
              value: val,
              bar: {
                x: Math.round(segLeftX),
                y: Math.round(barTop),
                width: Math.round(segWidth),
                height: Math.round(barHeight),
              },
              label: {
                x: Math.round(segCenterX - lblWidth / 2),
                y: Math.round(catCenterY - lblHeight / 2),
                width: Math.round(lblWidth),
                height: Math.round(lblHeight),
              },
            })

            cumulative = nextCumulative
          }
        } else {
          const slotHeight = rowHeight / Math.max(1, seriesCount)
          const barHeight = slotHeight * 0.8

          for (let s = 0; s < seriesCount; s++) {
            const val = Number(seriesValues[s]?.[c]) || 0
            const slotTopY = catTopY + s * slotHeight
            const slotCenterY = slotTopY + slotHeight / 2
            const barTop = slotCenterY - barHeight / 2

            const barWidth = plotWidth * (Math.abs(val) / axisMax)
            let barLeftX, barRightX
            if (val >= 0) {
              barLeftX = plotLeft
              barRightX = plotLeft + barWidth
            } else {
              barLeftX = plotLeft - barWidth
              barRightX = plotLeft
            }

            const seriesName = seriesNames[s] || `Series ${s + 1}`
            const labelText = String(val)
            const lblWidth = labelText.length * fontSize * 0.55 * emuPerPt
            const lblHeight = fontSize * 1.2 * emuPerPt
            const segCenterX = (barLeftX + barRightX) / 2

            segments.push({
              series: seriesName,
              category: categoryName,
              seriesIndex: s,
              categoryIndex: c,
              value: val,
              bar: {
                x: Math.round(barLeftX),
                y: Math.round(barTop),
                width: Math.round(barRightX - barLeftX),
                height: Math.round(barHeight),
              },
              label: {
                x: Math.round(segCenterX - lblWidth / 2),
                y: Math.round(slotCenterY - lblHeight / 2),
                width: Math.round(lblWidth),
                height: Math.round(lblHeight),
              },
            })
          }
        }
      }
    }

    return segments
  }

  #resolveBarSegmentCoordinates(
    plotLeft,
    plotTop,
    plotWidth,
    plotHeight,
    dir,
    grouping,
    categoriesCount,
    seriesCount,
    seriesValues,
    axisMax
  ) {
    const categories = Array.from({ length: categoriesCount }, (_, i) => `Category ${i + 1}`)
    const seriesNames = Array.from({ length: seriesCount }, (_, i) => `Series ${i + 1}`)
    const geom = {
      plotLeft,
      plotTop,
      plotWidth,
      plotHeight,
      dir,
      grouping,
      categories,
      seriesCount,
      categoriesCount,
      seriesValues,
      seriesNames,
      axisMax,
    }
    const segments = this.#resolveSegmentsGeometry(geom)
    return segments.map(seg => ({
      seriesIndex: seg.seriesIndex,
      categoryIndex: seg.categoryIndex,
      x: seg.label.x + seg.label.width / 2,
      y: seg.label.y + seg.label.height / 2,
    }))
  }

  #cleanRPrAttributes(attrs) {
    if (!attrs) return 'lang="en-US"'

    const attrMap = {}
    const pattern = /([a-zA-Z0-9:]+)="([^"]*)"/g
    let match
    while ((match = pattern.exec(attrs)) !== null) {
      attrMap[match[1]] = match[2]
    }

    delete attrMap['u']
    delete attrMap['strike']
    delete attrMap['kern']
    delete attrMap['baseline']
    delete attrMap['spc']
    delete attrMap['dirty']
    delete attrMap['smtClean']

    if (attrMap['b'] === '0' || attrMap['b'] === 'false') delete attrMap['b']
    if (attrMap['i'] === '0' || attrMap['i'] === 'false') delete attrMap['i']

    if (!attrMap['lang']) {
      attrMap['lang'] = 'en-US'
    }

    return Object.entries(attrMap)
      .map(([k, v]) => `${k}="${v}"`)
      .join(' ')
  }

  #resolveVerticalCollisions(labels, plotTop, plotBottom) {
    labels.sort((a, b) => a.targetY - b.targetY)
    const N = labels.length
    for (let iter = 0; iter < 100; iter++) {
      let moved = false
      for (let i = 0; i < N - 1; i++) {
        const curr = labels[i]
        const next = labels[i + 1]

        const currBottom = curr.y + curr.height / 2
        const nextTop = next.y - next.height / 2

        if (currBottom > nextTop) {
          const overlap = currBottom - nextTop
          curr.y -= overlap / 2
          next.y += overlap / 2
          moved = true
        }
      }

      for (let i = 0; i < N; i++) {
        const lbl = labels[i]
        const halfH = lbl.height / 2
        if (lbl.y - halfH < plotTop) {
          lbl.y = plotTop + halfH
        }
        if (lbl.y + halfH > plotBottom) {
          lbl.y = plotBottom - halfH
        }
      }

      if (!moved) break
    }
  }

  #escapeXml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;')
  }

  async #applySeriesNameLabels(
    slideIndex,
    chartId,
    data,
    chartXml,
    slideManager,
    relationshipManager
  ) {
    const options = data.seriesNameLabels
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideZipPath = slideInfo.zipPath

    const queue = this.#slideQueues.get(slideZipPath) || Promise.resolve()
    const nextTask = queue.then(async () => {
      let slideXml = await this.#zipManager.readFile(slideZipPath)

      const escapedChartId = chartId.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')
      const shapePattern = new RegExp(
        `<p:sp>(?:(?!<p:sp>)[\\s\\S])*?name="SeriesNameLabel-${escapedChartId}-\\d+(?:-\\d+)?"(?:(?!<\\/p:sp>)[\\s\\S])*?<\\/p:sp>`,
        'g'
      )
      slideXml = slideXml.replace(shapePattern, '')

      if (!options.enabled) {
        this.#zipManager.writeFile(slideZipPath, slideXml)
        return
      }

      const chartXfrm = this.#findChartCoordinates(
        slideXml,
        chartId,
        relationshipManager,
        slideZipPath
      )
      if (!chartXfrm) {
        logger.warn(`Could not find coordinates for chart "${chartId}" on slide ${slideIndex}`)
        this.#zipManager.writeFile(slideZipPath, slideXml)
        return
      }

      const plotLayout = this.#parsePlotAreaLayout(chartXml)
      const { dir, grouping } = this.#parseChartTypeAndGrouping(chartXml)

      const normalized = this.#normalizeChartData(data)
      const cleanNumericData = normalized.cleanData

      const categoriesCount = cleanNumericData.categories ? cleanNumericData.categories.length : 1
      const seriesCount = cleanNumericData.series ? cleanNumericData.series.length : 0
      const seriesValues = cleanNumericData.series.map(s => s.values || [])
      const seriesNames = cleanNumericData.series.map(s => s.name || '')

      const axisMax = this.#resolveAxisMax(
        grouping,
        seriesValues,
        categoriesCount,
        seriesCount,
        chartXml
      )

      const txPrMatch = /<c:txPr>([\s\S]*?)<\/c:txPr>/.exec(chartXml)
      const existingTxPr = txPrMatch ? txPrMatch[1] : ''
      const { ChartCacheGenerator } = require('./charts/ChartCacheGenerator.js')
      const { bodyPr, pPrXml, rPrXml } = ChartCacheGenerator.extractTxPrParts(existingTxPr)

      const szMatch = /sz="(\d+)"/.exec(rPrXml)
      let fontSize = szMatch ? parseInt(szMatch[1], 10) / 100 : 10
      if (options.style && options.style.fontSize !== undefined) {
        fontSize = Number(options.style.fontSize)
      }

      const position = options.position || 'left'
      const autoFit = options.autoFit !== false
      const emuPerPt = 12700
      const charAspect = 0.55
      const spacing = 200000

      const labelItems = []
      let maxLabelWidth = 0

      for (let s = 0; s < seriesCount; s++) {
        const seriesName = seriesNames[s] || `Series ${s + 1}`
        const textWidth = seriesName.length * fontSize * charAspect * emuPerPt
        maxLabelWidth = Math.max(maxLabelWidth, textWidth)
        labelItems.push({
          seriesIndex: s,
          seriesName,
          textWidth,
        })
      }

      const slideWidth = 12192000
      const fontEmuHeight = fontSize * 1.2 * emuPerPt

      const labelsForCollision = labelItems.map(item => {
        let maxWidth = maxLabelWidth
        if (position === 'left') {
          maxWidth = chartXfrm.left - spacing
        } else {
          maxWidth = slideWidth - (chartXfrm.left + chartXfrm.width) - spacing
        }
        maxWidth = Math.max(100000, maxWidth)

        let boxWidth = item.textWidth
        let boxHeight = fontEmuHeight
        if (autoFit) {
          if (boxWidth > maxWidth) {
            boxWidth = maxWidth
            const lines = Math.ceil(item.textWidth / maxWidth)
            boxHeight = lines * fontEmuHeight
          }
        } else {
          const standardWidth = 1371600
          boxWidth = Math.min(standardWidth, maxWidth)
          if (item.textWidth > boxWidth) {
            const lines = Math.ceil(item.textWidth / boxWidth)
            boxHeight = lines * fontEmuHeight
          }
        }

        const segments = this.#resolveSegmentsGeometry({
          plotLeft: chartXfrm.left + chartXfrm.width * plotLayout.x,
          plotTop: chartXfrm.top + chartXfrm.height * plotLayout.y,
          plotWidth: chartXfrm.width * plotLayout.w,
          plotHeight: chartXfrm.height * plotLayout.h,
          dir,
          grouping,
          categories: cleanNumericData.categories,
          seriesCount,
          categoriesCount,
          seriesValues,
          seriesNames,
          axisMax,
        })

        const matchedSegs = segments.filter(seg => seg.seriesIndex === item.seriesIndex)
        let sumY = 0
        matchedSegs.forEach(seg => {
          sumY += seg.bar.y + seg.bar.height / 2
        })
        const targetY =
          matchedSegs.length > 0 ? sumY / matchedSegs.length : chartXfrm.top + chartXfrm.height / 2

        return {
          seriesIndex: item.seriesIndex,
          name: item.seriesName,
          targetY,
          y: targetY,
          width: boxWidth,
          height: boxHeight,
        }
      })

      this.#resolveVerticalCollisions(
        labelsForCollision,
        chartXfrm.top,
        chartXfrm.top + chartXfrm.height
      )

      const finalBodyPr =
        bodyPr ||
        '<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:spAutoFit/></a:bodyPr>'
      let finalPPrXml = pPrXml || '<a:pPr/>'
      let templateRPr =
        rPrXml ||
        `<a:rPr lang="en-US" sz="${Math.round(fontSize * 100)}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr>`

      if (!templateRPr || !templateRPr.trim().startsWith('<a:rPr')) {
        templateRPr = `<a:rPr lang="en-US" sz="${Math.round(fontSize * 100)}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr>`
      }
      if (!finalPPrXml || !finalPPrXml.trim().startsWith('<a:pPr')) {
        finalPPrXml = '<a:pPr/>'
      }

      let existingSolidFill = ''
      let existingLatin = ''
      let existingEa = ''
      let existingCs = ''

      const solidFillMatch = /(<a:solidFill>[\s\S]*?<\/a:solidFill>|<a:solidFill[^>]*\/>)/.exec(
        templateRPr
      )
      if (solidFillMatch) {
        existingSolidFill = solidFillMatch[1]
      }

      const latinMatch = /(<a:latin[^>]*\/>)/.exec(templateRPr)
      if (latinMatch) {
        existingLatin = latinMatch[1]
      }

      const eaMatch = /(<a:ea[^>]*\/>)/.exec(templateRPr)
      if (eaMatch) {
        existingEa = eaMatch[1]
      }

      const csMatch = /(<a:cs[^>]*\/>)/.exec(templateRPr)
      if (csMatch) {
        existingCs = csMatch[1]
      }

      const rPrStartMatch = /^<a:rPr([^>]*)>/.exec(templateRPr)
      let rPrAttrs = rPrStartMatch ? rPrStartMatch[1] : ''
      if (rPrAttrs.endsWith('/')) {
        rPrAttrs = rPrAttrs.slice(0, -1)
      }
      rPrAttrs = rPrAttrs.trim()

      const cleanedAttrs = this.#cleanRPrAttributes(rPrAttrs)

      const attrMap = {}
      const pattern = /([a-zA-Z0-9:]+)="([^"]*)"/g
      let match
      while ((match = pattern.exec(cleanedAttrs)) !== null) {
        attrMap[match[1]] = match[2]
      }

      attrMap['sz'] = String(Math.round(fontSize * 100))
      if (options.style) {
        if (options.style.bold !== undefined) {
          if (options.style.bold) attrMap['b'] = '1'
          else delete attrMap['b']
        }
        if (options.style.italic !== undefined) {
          if (options.style.italic) attrMap['i'] = '1'
          else delete attrMap['i']
        }
        if (options.style.color !== undefined) {
          const hexColor = String(options.style.color).replace('#', '').trim()
          existingSolidFill = `<a:solidFill><a:srgbClr val="${hexColor}"/></a:solidFill>`
        }
        if (options.style.fontFamily !== undefined) {
          const typeface = options.style.fontFamily
          existingLatin = `<a:latin typeface="${typeface}"/>`
          existingEa = `<a:ea typeface="${typeface}"/>`
          existingCs = `<a:cs typeface="${typeface}"/>`
        }
      }

      if (
        existingSolidFill &&
        (existingSolidFill.includes('val="bg1"') || existingSolidFill.includes('val="bg2"'))
      ) {
        existingSolidFill = '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
      }

      const finalAttrsStr = Object.entries(attrMap)
        .map(([k, v]) => `${k}="${v}"`)
        .join(' ')

      let finalRPrXml = `<a:rPr ${finalAttrsStr}>`
      if (existingSolidFill) finalRPrXml += existingSolidFill
      if (existingLatin) finalRPrXml += existingLatin
      if (existingEa) finalRPrXml += existingEa
      if (existingCs) finalRPrXml += existingCs
      finalRPrXml += '</a:rPr>'

      let alignVal = options.style ? options.style.align : undefined
      if (!alignVal) {
        alignVal = position === 'left' ? 'r' : 'l'
      } else {
        if (alignVal === 'left') alignVal = 'l'
        else if (alignVal === 'right') alignVal = 'r'
        else if (alignVal === 'center') alignVal = 'ctr'
      }

      if (finalPPrXml.includes('algn=')) {
        finalPPrXml = finalPPrXml.replace(/algn="[^"]+"/, `algn="${alignVal}"`)
      } else if (finalPPrXml.includes('<a:pPr')) {
        finalPPrXml = finalPPrXml.replace('<a:pPr', `<a:pPr algn="${alignVal}"`)
      } else {
        finalPPrXml = `<a:pPr algn="${alignVal}"/>`
      }

      const idMatchPattern = /id="(\d+)"/g
      let idMatch
      const existingIds = []
      while ((idMatch = idMatchPattern.exec(slideXml)) !== null) {
        existingIds.push(parseInt(idMatch[1], 10))
      }

      let shapesXml = ''
      const { generateUniqueId } = require('../utils/idUtils.js')
      for (let i = 0; i < labelsForCollision.length; i++) {
        const lbl = labelsForCollision[i]
        const newId = generateUniqueId(existingIds)
        existingIds.push(newId)

        let boxLeft = 0
        if (position === 'left') {
          boxLeft = chartXfrm.left - lbl.width - spacing
        } else {
          boxLeft = chartXfrm.left + chartXfrm.width + spacing
        }
        boxLeft = Math.max(0, boxLeft)
        const boxTop = lbl.y - lbl.height / 2

        shapesXml += `<p:sp>
          <p:nvSpPr>
            <p:cNvPr id="${newId}" name="SeriesNameLabel-${chartId}-${lbl.seriesIndex}"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="${Math.round(boxLeft)}" y="${Math.round(boxTop)}"/>
              <a:ext cx="${Math.round(lbl.width)}" cy="${Math.round(lbl.height)}"/>
            </a:xfrm>
            <a:prstGeom prst="rect">
              <a:avLst/>
            </a:prstGeom>
            <a:noFill/>
          </p:spPr>
          <p:txBody>
            ${finalBodyPr}
            <a:lstStyle/>
            <a:p>
              ${finalPPrXml}
              <a:r>
                ${finalRPrXml}
                <a:t>${this.#escapeXml(lbl.name)}</a:t>
              </a:r>
              <a:endParaRPr lang="en-US"/>
            </a:p>
          </p:txBody>
        </p:sp>`
      }

      const spTreeEnd = '</p:spTree>'
      if (slideXml.includes(spTreeEnd)) {
        slideXml = slideXml.replace(spTreeEnd, `${shapesXml}${spTreeEnd}`)
      }

      this.#zipManager.writeFile(slideZipPath, slideXml)
    })

    this.#slideQueues.set(slideZipPath, nextTask)
    return nextTask
  }

  #getAxisMax(maxVal) {
    if (maxVal <= 0) return 1
    const padded = maxVal * 1.05
    const power = Math.floor(Math.log10(padded))
    const temp = padded / Math.pow(10, power)
    let niceMax
    if (temp <= 1.0) niceMax = 1.0
    else if (temp <= 1.2) niceMax = 1.2
    else if (temp <= 1.5) niceMax = 1.5
    else if (temp <= 2.0) niceMax = 2.0
    else if (temp <= 2.5) niceMax = 2.5
    else if (temp <= 3.0) niceMax = 3.0
    else if (temp <= 4.0) niceMax = 4.0
    else if (temp <= 5.0) niceMax = 5.0
    else if (temp <= 6.0) niceMax = 6.0
    else if (temp <= 8.0) niceMax = 8.0
    else niceMax = 10.0
    return niceMax * Math.pow(10, power)
  }

  async #resolveChartGeometry(slideIndex, chartId, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new Error(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    const queue = this.#chartQueues.get(chartInfo.zipPath) || Promise.resolve()
    await queue
    const chartXml = await this.#zipManager.readFile(chartInfo.zipPath)
    if (!chartXml) {
      throw new Error(`Chart file not found: ${chartInfo.zipPath}`)
    }

    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideXml = await this.#zipManager.readFile(slideInfo.zipPath)
    const chartXfrm = this.#findChartCoordinates(
      slideXml,
      chartId,
      relationshipManager,
      slideInfo.zipPath
    )
    if (!chartXfrm) {
      throw new Error(`Coordinates not found for chart "${chartId}" on slide ${slideIndex}`)
    }

    const plotLayout = this.#parsePlotAreaLayout(chartXml)
    const { dir, grouping } = this.#parseChartTypeAndGrouping(chartXml)

    const chartData = this.#extractChartData(chartXml)
    const categories = chartData.categories
    const series = chartData.series

    const categoriesCount = categories.length || 1
    const seriesCount = series.length
    const seriesValues = series.map(s => s.values || [])
    const seriesNames = series.map(s => s.name || '')

    const axisMax = this.#resolveAxisMax(
      grouping,
      seriesValues,
      categoriesCount,
      seriesCount,
      chartXml
    )

    const chartLeft = chartXfrm.left
    const chartTop = chartXfrm.top
    const chartWidth = chartXfrm.width
    const chartHeight = chartXfrm.height

    const px = plotLayout.x
    const py = plotLayout.y
    const pw = plotLayout.w
    const ph = plotLayout.h
    const plotLeft = chartLeft + chartWidth * px
    const plotTop = chartTop + chartHeight * py
    const plotWidth = chartWidth * pw
    const plotHeight = chartHeight * ph

    return {
      chartLeft,
      chartTop,
      chartWidth,
      chartHeight,
      plotLeft,
      plotTop,
      plotWidth,
      plotHeight,
      dir,
      grouping,
      categories,
      series,
      seriesCount,
      categoriesCount,
      seriesValues,
      seriesNames,
      axisMax,
    }
  }

  async getChartLabelPositions(slideIndex, chartId, slideManager, relationshipManager) {
    const geom = await this.#resolveChartGeometry(
      slideIndex,
      chartId,
      slideManager,
      relationshipManager
    )
    const segments = this.#resolveSegmentsGeometry(geom)
    return segments.map(seg => ({
      series: seg.series,
      category: seg.category,
      seriesIndex: seg.seriesIndex,
      categoryIndex: seg.categoryIndex,
      value: seg.value,
      x: seg.label.x,
      y: seg.label.y,
      width: seg.label.width,
      height: seg.label.height,
    }))
  }

  async getChartBarPositions(slideIndex, chartId, slideManager, relationshipManager) {
    const geom = await this.#resolveChartGeometry(
      slideIndex,
      chartId,
      slideManager,
      relationshipManager
    )
    const segments = this.#resolveSegmentsGeometry(geom)
    return segments.map(seg => ({
      series: seg.series,
      category: seg.category,
      seriesIndex: seg.seriesIndex,
      categoryIndex: seg.categoryIndex,
      value: seg.value,
      x: seg.bar.x,
      y: seg.bar.y,
      width: seg.bar.width,
      height: seg.bar.height,
    }))
  }

  async addTextAtPosition(slideIndex, options, slideManager) {
    const { text, x, y, width = 1200000, height = 300000, style = {} } = options
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideZipPath = slideInfo.zipPath

    const queue = this.#slideQueues.get(slideZipPath) || Promise.resolve()
    const nextTask = queue.then(async () => {
      let slideXml = await this.#zipManager.readFile(slideZipPath)

      const fontSize = style.fontSize || 10
      const fontFamily = style.fontFamily || 'Arial'
      const alignVal = style.align === 'center' ? 'ctr' : style.align === 'right' ? 'r' : 'l'

      let colorXml = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
      if (style.color) {
        const hexColor = String(style.color).replace('#', '').trim()
        colorXml = `<a:solidFill><a:srgbClr val="${hexColor}"/></a:solidFill>`
      }

      const finalBodyPr =
        '<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:spAutoFit/></a:bodyPr>'
      const finalPPrXml = `<a:pPr algn="${alignVal}"/>`

      let rPrAttrs = `lang="en-US" sz="${Math.round(fontSize * 100)}"`
      if (style.bold) rPrAttrs += ' b="1"'
      if (style.italic) rPrAttrs += ' i="1"'

      const finalRPrXml = `<a:rPr ${rPrAttrs}>${colorXml}<a:latin typeface="${fontFamily}"/><a:ea typeface="${fontFamily}"/><a:cs typeface="${fontFamily}"/></a:rPr>`

      const idMatchPattern = /id="(\d+)"/g
      let idMatch
      const existingIds = []
      while ((idMatch = idMatchPattern.exec(slideXml)) !== null) {
        existingIds.push(parseInt(idMatch[1], 10))
      }
      const { generateUniqueId } = require('../utils/idUtils.js')
      const newId = generateUniqueId(existingIds)

      const shapeXml = `<p:sp>
        <p:nvSpPr>
          <p:cNvPr id="${newId}" name="TextBoxAtPosition-${newId}"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="${Math.round(x)}" y="${Math.round(y)}"/>
            <a:ext cx="${Math.round(width)}" cy="${Math.round(height)}"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          ${finalBodyPr}
          <a:lstStyle/>
          <a:p>
            ${finalPPrXml}
            <a:r>
              ${finalRPrXml}
              <a:t>${this.#escapeXml(text)}</a:t>
            </a:r>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </p:txBody>
      </p:sp>`

      const spTreeEnd = '</p:spTree>'
      if (slideXml.includes(spTreeEnd)) {
        slideXml = slideXml.replace(spTreeEnd, `${shapeXml}${spTreeEnd}`)
      }

      this.#zipManager.writeFile(slideZipPath, slideXml)
    })

    this.#slideQueues.set(slideZipPath, nextTask)
    return nextTask
  }

  async addTextNearChartLabel(slideIndex, options, slideManager, relationshipManager) {
    const { chart, text, position = 'left', style = {} } = options

    const allowedPositions = ['left', 'right']
    if (!allowedPositions.includes(position)) {
      throw new Error(`Invalid position "${position}". Only "left" and "right" are supported.`)
    }

    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideZipPath = slideInfo.zipPath

    const queue = this.#slideQueues.get(slideZipPath) || Promise.resolve()
    const nextTask = queue.then(async () => {
      const labels = await this.getChartLabelPositions(
        slideIndex,
        chart,
        slideManager,
        relationshipManager
      )

      let slideXml = await this.#zipManager.readFile(slideZipPath)
      const escapedChartId = chart.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')
      const shapePattern = new RegExp(
        `<p:sp>(?:(?!<p:sp>)[\\s\\S])*?name="SeriesNameLabel-${escapedChartId}-\\d+"(?:(?!<\\/p:sp>)[\\s\\S])*?<\\/p:sp>`,
        'g'
      )
      slideXml = slideXml.replace(shapePattern, '')
      this.#zipManager.writeFile(slideZipPath, slideXml)

      const fontSize = style.fontSize || 10
      const emuPerPt = 12700
      const charAspect = 0.55
      const spacing = 200000

      const labelItems = []
      let maxLabelWidth = 0

      for (let i = 0; i < labels.length; i++) {
        const lbl = labels[i]
        let labelText = ''
        if (typeof text === 'function') {
          labelText = text({
            series: lbl.series,
            category: lbl.category,
            value: lbl.value,
          })
        } else {
          labelText = String(text)
        }

        const textWidth = labelText.length * fontSize * charAspect * emuPerPt
        maxLabelWidth = Math.max(maxLabelWidth, textWidth)
        labelItems.push({
          seriesIndex: i,
          seriesName: lbl.series,
          categoryName: lbl.category,
          labelText,
          textWidth,
          lbl,
        })
      }

      const chartXfrm = this.#findChartCoordinates(
        slideXml,
        chart,
        relationshipManager,
        slideZipPath
      )

      const autoFit = style.autoFit !== false
      const slideWidth = 12192000
      const fontEmuHeight = fontSize * 1.2 * emuPerPt

      const labelsForCollision = labelItems.map(item => {
        let maxWidth = maxLabelWidth
        if (chartXfrm) {
          if (position === 'left') {
            maxWidth = chartXfrm.left - spacing
          } else {
            maxWidth = slideWidth - (chartXfrm.left + chartXfrm.width) - spacing
          }
        }
        maxWidth = Math.max(100000, maxWidth)

        let boxWidth = item.textWidth
        let boxHeight = fontEmuHeight
        if (autoFit) {
          if (boxWidth > maxWidth) {
            boxWidth = maxWidth
            const lines = Math.ceil(item.textWidth / maxWidth)
            boxHeight = lines * fontEmuHeight
          }
        } else {
          const standardWidth = 1371600
          boxWidth = Math.min(standardWidth, maxWidth)
          if (item.textWidth > boxWidth) {
            const lines = Math.ceil(item.textWidth / boxWidth)
            boxHeight = lines * fontEmuHeight
          }
        }

        return {
          seriesIndex: item.seriesIndex,
          name: item.labelText,
          targetY: item.lbl.y + item.lbl.height / 2,
          y: item.lbl.y + item.lbl.height / 2,
          width: boxWidth,
          height: boxHeight,
          lbl: item.lbl,
        }
      })

      if (chartXfrm) {
        this.#resolveVerticalCollisions(
          labelsForCollision,
          chartXfrm.top,
          chartXfrm.top + chartXfrm.height
        )
      }

      const chartInfo = this.findChartInSlide(slideIndex, chart, slideManager, relationshipManager)
      const chartXml = await this.#zipManager.readFile(chartInfo.zipPath)
      const txPrMatch = /<c:txPr>([\s\S]*?)<\/c:txPr>/.exec(chartXml)
      const existingTxPr = txPrMatch ? txPrMatch[1] : ''
      const { ChartCacheGenerator } = require('./charts/ChartCacheGenerator.js')
      const { pPrXml, rPrXml } = ChartCacheGenerator.extractTxPrParts(existingTxPr)

      const finalBodyPr =
        '<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:spAutoFit/></a:bodyPr>'
      let finalPPrXml = pPrXml || '<a:pPr/>'
      let templateRPr =
        rPrXml ||
        `<a:rPr lang="en-US" sz="${Math.round(fontSize * 100)}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr>`

      if (!templateRPr || !templateRPr.trim().startsWith('<a:rPr')) {
        templateRPr = `<a:rPr lang="en-US" sz="${Math.round(fontSize * 100)}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr>`
      }
      if (!finalPPrXml || !finalPPrXml.trim().startsWith('<a:pPr')) {
        finalPPrXml = '<a:pPr/>'
      }

      let existingSolidFill = ''
      let existingLatin = ''
      let existingEa = ''
      let existingCs = ''

      const solidFillMatch = /(<a:solidFill>[\s\S]*?<\/a:solidFill>|<a:solidFill[^>]*\/>)/.exec(
        templateRPr
      )
      if (solidFillMatch) {
        existingSolidFill = solidFillMatch[1]
      }

      const latinMatch = /(<a:latin[^>]*\/>)/.exec(templateRPr)
      if (latinMatch) {
        existingLatin = latinMatch[1]
      }

      const eaMatch = /(<a:ea[^>]*\/>)/.exec(templateRPr)
      if (eaMatch) {
        existingEa = eaMatch[1]
      }

      const csMatch = /(<a:cs[^>]*\/>)/.exec(templateRPr)
      if (csMatch) {
        existingCs = csMatch[1]
      }

      const rPrStartMatch = /^<a:rPr([^>]*)>/.exec(templateRPr)
      let rPrAttrs = rPrStartMatch ? rPrStartMatch[1] : ''
      if (rPrAttrs.endsWith('/')) {
        rPrAttrs = rPrAttrs.slice(0, -1)
      }
      rPrAttrs = rPrAttrs.trim()

      const cleanedAttrs = this.#cleanRPrAttributes(rPrAttrs)

      const attrMap = {}
      const pattern = /([a-zA-Z0-9:]+)="([^"]*)"/g
      let match
      while ((match = pattern.exec(cleanedAttrs)) !== null) {
        attrMap[match[1]] = match[2]
      }

      attrMap['sz'] = String(Math.round(fontSize * 100))
      if (style.bold !== undefined) {
        if (style.bold) attrMap['b'] = '1'
        else delete attrMap['b']
      }
      if (style.italic !== undefined) {
        if (style.italic) attrMap['i'] = '1'
        else delete attrMap['i']
      }

      if (style.color !== undefined) {
        const hexColor = String(style.color).replace('#', '').trim()
        existingSolidFill = `<a:solidFill><a:srgbClr val="${hexColor}"/></a:solidFill>`
      } else {
        if (
          existingSolidFill &&
          (existingSolidFill.includes('val="bg1"') || existingSolidFill.includes('val="bg2"'))
        ) {
          existingSolidFill = '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        }
      }

      if (style.fontFamily !== undefined) {
        const typeface = style.fontFamily
        existingLatin = `<a:latin typeface="${typeface}"/>`
        existingEa = `<a:ea typeface="${typeface}"/>`
        existingCs = `<a:cs typeface="${typeface}"/>`
      }

      const finalAttrsStr = Object.entries(attrMap)
        .map(([k, v]) => `${k}="${v}"`)
        .join(' ')

      let finalRPrXml = `<a:rPr ${finalAttrsStr}>`
      if (existingSolidFill) finalRPrXml += existingSolidFill
      if (existingLatin) finalRPrXml += existingLatin
      if (existingEa) finalRPrXml += existingEa
      if (existingCs) finalRPrXml += existingCs
      finalRPrXml += '</a:rPr>'

      let alignVal = style.align
      if (!alignVal) {
        alignVal = position === 'left' ? 'r' : 'l'
      } else {
        if (alignVal === 'left') alignVal = 'l'
        else if (alignVal === 'right') alignVal = 'r'
        else if (alignVal === 'center') alignVal = 'ctr'
      }

      if (finalPPrXml.includes('algn=')) {
        finalPPrXml = finalPPrXml.replace(/algn="[^"]+"/, `algn="${alignVal}"`)
      } else if (finalPPrXml.includes('<a:pPr')) {
        finalPPrXml = finalPPrXml.replace('<a:pPr', `<a:pPr algn="${alignVal}"`)
      } else {
        finalPPrXml = `<a:pPr algn="${alignVal}"/>`
      }

      slideXml = await this.#zipManager.readFile(slideZipPath)
      const idMatchPattern = /id="(\d+)"/g
      let idMatch
      const existingIds = []
      while ((idMatch = idMatchPattern.exec(slideXml)) !== null) {
        existingIds.push(parseInt(idMatch[1], 10))
      }

      let shapesXml = ''
      const { generateUniqueId } = require('../utils/idUtils.js')
      for (let i = 0; i < labelsForCollision.length; i++) {
        const lbl = labelsForCollision[i]
        const newId = generateUniqueId(existingIds)
        existingIds.push(newId)

        let boxLeft = 0
        if (chartXfrm) {
          if (position === 'left') {
            boxLeft = chartXfrm.left - lbl.width - spacing
          } else {
            boxLeft = chartXfrm.left + chartXfrm.width + spacing
          }
        } else {
          boxLeft = lbl.lbl.x - lbl.width - spacing
        }
        boxLeft = Math.max(0, boxLeft)
        const boxTop = lbl.y - lbl.height / 2

        shapesXml += `<p:sp>
          <p:nvSpPr>
            <p:cNvPr id="${newId}" name="SeriesNameLabel-${chart}-${lbl.lbl.seriesIndex}"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="${Math.round(boxLeft)}" y="${Math.round(boxTop)}"/>
              <a:ext cx="${Math.round(lbl.width)}" cy="${Math.round(lbl.height)}"/>
            </a:xfrm>
            <a:prstGeom prst="rect">
              <a:avLst/>
            </a:prstGeom>
            <a:noFill/>
          </p:spPr>
          <p:txBody>
            ${finalBodyPr}
            <a:lstStyle/>
            <a:p>
              ${finalPPrXml}
              <a:r>
                ${finalRPrXml}
                <a:t>${this.#escapeXml(lbl.name)}</a:t>
              </a:r>
              <a:endParaRPr lang="en-US"/>
            </a:p>
          </p:txBody>
        </p:sp>`
      }

      const spTreeEnd = '</p:spTree>'
      if (slideXml.includes(spTreeEnd)) {
        slideXml = slideXml.replace(spTreeEnd, `${shapesXml}${spTreeEnd}`)
      }
      this.#zipManager.writeFile(slideZipPath, slideXml)
    })

    this.#slideQueues.set(slideZipPath, nextTask)
    return nextTask
  }
}

module.exports = { ChartManager }
