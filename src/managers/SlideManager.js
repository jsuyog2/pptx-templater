/**
 * @fileoverview SlideManager - Manages individual slide operations.
 *
 * Slides in OpenXML PPTX:
 * ─────────────────────────────────────────────────────────────────
 * Each slide is an XML file stored at ppt/slides/slideN.xml.
 * The slide order is defined in ppt/presentation.xml under:
 *   <p:sldIdLst>
 *     <p:sldId id="256" r:id="rId2"/>   ← slide 1
 *     <p:sldId id="257" r:id="rId3"/>   ← slide 2
 *   </p:sldIdLst>
 *
 * Each slide has:
 *  - A content tree: p:sld > p:cSld > p:spTree > [shapes/text/images]
 *  - A layout reference via its .rels file (points to slideLayouts/slideLayoutN.xml)
 *  - Optional notes slide: ppt/notesSlides/notesSlideN.xml
 *  - Optional animation data: embedded in the slide XML itself
 *
 * Shape types in spTree:
 *  - p:sp       → Text/shape placeholders
 *  - p:pic      → Images
 *  - p:graphicFrame → Charts, tables, SmartArt
 *  - p:grpSp    → Grouped shapes
 *  - p:cxnSp    → Connectors
 */

const { createLogger } = require('../utils/logger.js')
const { PPTXError, SlideNotFoundError } = require('../utils/errors.js')
const { REL_TYPES } = require('./RelationshipManager.js')
const { buildNewSlideXml } = require('../templates/slideTemplate.js')
const { remapRelationshipIds, generateRelationshipId } = require('../utils/relationshipUtils.js')
const { generateSlideId } = require('../utils/idUtils.js')

const logger = createLogger('SlideManager')

/** MIME type for PPTX slide parts. */
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'

const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
const CHART_STYLE_CONTENT_TYPE = 'application/vnd.ms-office.chartstyle+xml'
const CHART_COLORS_CONTENT_TYPE = 'application/vnd.ms-office.chartcolorstyle+xml'
const WORKBOOK_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const ACTIVEX_CONTENT_TYPE = 'application/vnd.ms-office.activeX'

/**
 * @typedef {Object} SlideInfo
 * @property {number} index - 1-based slide number.
 * @property {string} zipPath - Path within the ZIP.
 * @property {string} relationshipId - rId in presentation.xml.rels.
 * @property {string} slideId - Unique slide ID from presentation.xml.
 * @property {string[]} tags - Custom tags assigned via tagSlide().
 * @property {string} title - Slide title (extracted from title placeholder).
 */

/**
 * @class SlideManager
 * @description Manages slide loading, ordering, modification, and creation.
 */
class SlideManager {
  /** @private @type {XMLParser} */
  #xmlParser
  /** @private @type {RelationshipManager} */
  #relationshipManager
  /** @private @type {ContentTypesManager} */
  #contentTypesManager
  /** @private @type {ZipManager} */
  #zipManager

  /**
   * Slide registry: maps 1-based index → SlideInfo.
   * @private @type {Map<number, SlideInfo>}
   */
  #slides = new Map()

  /**
   * Raw XML cache: maps zipPath → XML string.
   * @private @type {Map<string, string>}
   */
  #slideXmlCache = new Map()

  /**
   * Slide states: maps 1-based index → state object.
   * @private @type {Map<number, Object>}
   */
  #slideStates = new Map()

  /**
   * Custom tags: maps tag → array of 1-based indices.
   * @private @type {Map<string, number[]>}
   */
  #tags = new Map()

  /**
   * Parsed presentation.xml object (cached for modification).
   * @private @type {Object}
   */
  #presentationObj = null

  /**
   * Map of old presentation relationship IDs to new sequential IDs.
   * @private @type {Map<string, string>}
   */
  #presentationIdMap = null

  /**
   * @param {XMLParser} xmlParser
   * @param {RelationshipManager} relationshipManager
   * @param {ContentTypesManager} contentTypesManager
   */
  constructor(xmlParser, relationshipManager, contentTypesManager) {
    this.#xmlParser = xmlParser
    this.#relationshipManager = relationshipManager
    this.#contentTypesManager = contentTypesManager
  }

  /**
   * Initializes by reading presentation.xml and discovering all slides.
   *
   * @param {ZipManager} zipManager
   * @returns {Promise<void>}
   */
  async initialize(zipManager) {
    this.#zipManager = zipManager
    const presentationXml = await zipManager.readFile('ppt/presentation.xml')

    if (!presentationXml) {
      // If no presentation.xml, create a blank one
      await this.#initializeBlankPresentation()
      return
    }

    this.#presentationObj = this.#xmlParser.parse(presentationXml, 'presentation.xml')
    await this.#discoverSlides(zipManager)
  }

  /**
   * Discovers all slides from presentation.xml and caches their info.
   * @private
   */
  async #discoverSlides(zipManager) {
    const rels = this.#relationshipManager.getRelationships('ppt/presentation.xml')
    const slideRels = rels.filter(r => r.type === REL_TYPES.SLIDE)

    // Get slide order from presentation.xml sldIdLst
    const sldIdList = this.#xmlParser.findAll(
      this.#presentationObj,
      'p:presentation.p:sldIdLst.p:sldId'
    )

    // Map rId → slide info from sldIdLst
    const rIdToSlideId = new Map()
    for (const sldId of sldIdList) {
      const rId = sldId['@_r:id']
      const slideId = sldId['@_id']
      if (rId) rIdToSlideId.set(rId, slideId)
    }

    // Attempt to read slide titles from docProps/app.xml to preserve them
    let slideTitles = []
    try {
      const appXml = await zipManager.readFile('docProps/app.xml')
      if (appXml) {
        const appObj = this.#xmlParser.parse(appXml, 'app.xml')
        const lpstrs = appObj?.Properties?.TitlesOfParts?.['vt:vector']?.['vt:lpstr']
        if (lpstrs) {
          const allLpstrs = Array.isArray(lpstrs) ? lpstrs : [lpstrs]
          // Slide titles are usually the last N items where N = slide count
          // We take the last slideRels.length items
          slideTitles = allLpstrs.slice(-slideRels.length)
        }
      }
    } catch (e) {
      logger.warn('Failed to parse app.xml for slide titles', e)
    }

    // Build ordered slide list
    let slideIndex = 1
    for (const sldId of sldIdList) {
      const rId = sldId['@_r:id']
      const slideRel = slideRels.find(r => r.id === rId)
      if (!slideRel) continue

      // Resolve absolute path from relative target
      const zipPath = this.#relationshipManager.resolveTarget(
        'ppt/presentation.xml',
        slideRel.target
      )

      const slideInfo = {
        index: slideIndex,
        zipPath,
        relationshipId: rId,
        slideId: rIdToSlideId.get(rId) || String(256 + slideIndex),
        tags: [],
        title: slideTitles[slideIndex - 1] || '',
      }

      this.#slides.set(slideIndex, slideInfo)
      this.#slideStates.set(slideIndex, {
        xmlStr: null,
        xmlObj: null,
        dirty: false,
        indexBuilt: false,
        shapeMap: new Map(),
        picMap: new Map(),
        tableMap: new Map(),
        chartMap: new Map(),
      })
      slideIndex++
    }

    logger.debug(`Discovered ${this.#slides.size} slides`)
  }

  /**
   * Returns the total number of slides.
   * @returns {number}
   */
  get slideCount() {
    return this.#slides.size
  }

  /**
   * Returns all 1-based slide indices.
   * @returns {number[]}
   */
  getAllSlideIndices() {
    return Array.from(this.#slides.keys()).sort((a, b) => a - b)
  }

  /**
   * Returns info objects for all slides.
   * @returns {SlideInfo[]}
   */
  getAllSlideInfo() {
    return this.getAllSlideIndices().map(i => this.#slides.get(i))
  }

  /**
   * Resolves a string/number ref to an array of 1-based slide indices.
   * - If number: returns [number]
   * - If string: looks up tag registry
   *
   * @param {number|string} ref
   * @returns {number[]}
   */
  resolveSlideRef(ref) {
    if (typeof ref === 'number') {
      this.#assertSlideExists(ref)
      return [ref]
    }
    const taggedIndices = this.#tags.get(ref)
    if (!taggedIndices || taggedIndices.length === 0) {
      throw new SlideNotFoundError(`No slides found with tag: "${ref}"`)
    }
    return taggedIndices
  }

  /**
   * Gets the raw XML string for a slide.
   * Loads from ZIP on first access, then returns from cache.
   *
   * @param {number} slideIndex - 1-based index.
   * @returns {string} Slide XML content.
   */
  getSlideXml(slideIndex) {
    this.#assertSlideExists(slideIndex)
    const info = this.#slides.get(slideIndex)
    const state = this.#slideStates.get(slideIndex)

    if (state.dirty && state.xmlObj) {
      const decl = this.#xmlParser.extractDeclaration(
        state.xmlStr || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      )
      state.xmlStr = this.#xmlParser.build(state.xmlObj, decl)
      this.#zipManager.writeFile(info.zipPath, state.xmlStr)
      this.#slideXmlCache.set(info.zipPath, state.xmlStr)
      state.dirty = false
    }

    const dirtyContent = this.#zipManager.readCachedFile(info.zipPath)
    if (dirtyContent && dirtyContent !== state.xmlStr) {
      state.xmlStr = dirtyContent
      this.#slideXmlCache.set(info.zipPath, dirtyContent)
      state.xmlObj = null
      state.indexBuilt = false
    }

    if (state.xmlStr) {
      return state.xmlStr
    }

    const cached = this.#zipManager.readCachedFile(info.zipPath)
    if (cached) {
      state.xmlStr = cached
      this.#slideXmlCache.set(info.zipPath, cached)
      return cached
    }

    // This is sync because we pre-load; async callers should use getSlideXmlAsync
    throw new PPTXError(`Slide ${slideIndex} XML not pre-loaded. Use getSlideXmlAsync().`)
  }

  /**
   * Async version of getSlideXml — loads from ZIP if not cached.
   *
   * @param {number} slideIndex - 1-based index.
   * @returns {Promise<string>}
   */
  async getSlideXmlAsync(slideIndex) {
    this.#assertSlideExists(slideIndex)
    const info = this.#slides.get(slideIndex)
    const state = this.#slideStates.get(slideIndex)

    const dirtyContent = this.#zipManager.readCachedFile(info.zipPath)
    if (dirtyContent && dirtyContent !== state.xmlStr) {
      state.xmlStr = dirtyContent
      this.#slideXmlCache.set(info.zipPath, dirtyContent)
      state.xmlObj = null
      state.indexBuilt = false
    }

    if (!state.xmlStr) {
      const xml = await this.#zipManager.readFile(info.zipPath)
      if (!xml) throw new SlideNotFoundError(`Slide ${slideIndex} XML not found at ${info.zipPath}`)
      state.xmlStr = xml
      this.#slideXmlCache.set(info.zipPath, xml)
    }

    return state.xmlStr
  }

  /**
   * Sets (replaces) the XML for a slide and marks it as dirty.
   *
   * @param {number} slideIndex - 1-based index.
   * @param {string} xml - New XML content.
   */
  setSlideXml(slideIndex, xml) {
    this.#assertSlideExists(slideIndex)
    const info = this.#slides.get(slideIndex)
    const state = this.#slideStates.get(slideIndex)

    state.xmlStr = xml
    state.xmlObj = null
    state.dirty = false
    state.indexBuilt = false

    this.#slideXmlCache.set(info.zipPath, xml)
    this.#zipManager.writeFile(info.zipPath, xml)
  }

  /**
   * Tags a slide with a custom string identifier.
   *
   * @param {number} slideIndex - 1-based index.
   * @param {string} tag - Tag string.
   */
  tagSlide(slideIndex, tag) {
    this.#assertSlideExists(slideIndex)
    const info = this.#slides.get(slideIndex)
    if (!info.tags.includes(tag)) info.tags.push(tag)

    if (!this.#tags.has(tag)) this.#tags.set(tag, [])
    const tagList = this.#tags.get(tag)
    if (!tagList.includes(slideIndex)) tagList.push(slideIndex)
  }

  /**
   * Gets the SlideInfo for a slide.
   *
   * @param {number} slideIndex - 1-based index.
   * @returns {SlideInfo}
   */
  getSlideInfo(slideIndex) {
    this.#assertSlideExists(slideIndex)
    return this.#slides.get(slideIndex)
  }

  /**
   * Adds a completely new slide to the presentation.
   * Creates the XML file, adds it to presentation.xml, and registers relationships.
   *
   * @param {NewSlideOptions} options
   * @param {RelationshipManager} relationshipManager
   * @param {MediaManager} mediaManager
   */
  addNewSlide(options, relationshipManager, _mediaManager) {
    const newIndex = this.#slides.size + 1
    let nextFileIndex = 1
    while (this.#zipManager.hasFile(`ppt/slides/slide${nextFileIndex}.xml`)) {
      nextFileIndex++
    }
    const slideFileName = `slide${nextFileIndex}.xml`
    const slideZipPath = `ppt/slides/${slideFileName}`

    // Generate the slide XML
    const slideXml = buildNewSlideXml(options, newIndex)

    // Write the slide XML to the ZIP
    this.#zipManager.writeFile(slideZipPath, slideXml)
    this.#slideXmlCache.set(slideZipPath, slideXml)

    // Add slide relationship to presentation.xml.rels
    const rId = relationshipManager.addRelationship(
      'ppt/presentation.xml',
      REL_TYPES.SLIDE,
      `slides/${slideFileName}`
    )

    // Add slide layout relationship to the new slide's .rels
    // Reference the first available layout
    const layoutRelsAll = this.#zipManager
      .listFiles('ppt/slideLayouts/')
      .filter(f => f.endsWith('.xml'))
    const firstLayout = layoutRelsAll[0] || 'ppt/slideLayouts/slideLayout1.xml'
    const relativeLayoutPath = `../slideLayouts/${firstLayout.split('/').pop()}`
    relationshipManager.addRelationship(slideZipPath, REL_TYPES.SLIDE_LAYOUT, relativeLayoutPath)

    // Generate a unique slide ID
    const existingSlideIds = Array.from(this.#slides.values()).map(s => s.slideId)
    const newSlideId = generateSlideId(existingSlideIds)

    const slideInfo = {
      index: newIndex,
      zipPath: slideZipPath,
      relationshipId: rId,
      slideId: newSlideId,
      tags: [],
      title: options.title || '',
    }

    this.#slides.set(newIndex, slideInfo)
    this.#slideStates.set(newIndex, {
      xmlStr: slideXml,
      xmlObj: null,
      dirty: false,
      indexBuilt: false,
      shapeMap: new Map(),
      picMap: new Map(),
      tableMap: new Map(),
      chartMap: new Map(),
    })

    // Update presentation.xml sldIdLst
    this.#addSlideToPresentation(rId, newSlideId)

    // Update [Content_Types].xml
    this.#registerSlideContentType(slideFileName)

    logger.debug(`Added new slide ${newIndex} at ${slideZipPath}`)
  }

  /**
   * Clones an existing slide, duplicating its XML and relationships.
   *
   * @param {number} sourceIndex - 1-based source slide number.
   * @param {number} [atPosition] - Insert position (1-based). Default: append.
   * @param {RelationshipManager} relationshipManager
   * @param {MediaManager} mediaManager
   * @returns {Promise<void>}
   */
  async cloneSlide(sourceIndex, atPosition, relationshipManager, mediaManager) {
    const promise = this.#cloneSlideInternal(
      sourceIndex,
      atPosition,
      relationshipManager,
      mediaManager
    )
    this.#zipManager.addPendingPromise(promise)
    return promise
  }

  /**
   * @private
   */
  async #cloneSlideInternal(sourceIndex, atPosition, relationshipManager, mediaManager) {
    this.#assertSlideExists(sourceIndex)
    const sourceInfo = this.#slides.get(sourceIndex)
    logger.debug('Source Slide Info:', sourceInfo)

    const newIndex = this.#slides.size + 1
    let nextFileIndex = 1
    while (this.#zipManager.hasFile(`ppt/slides/slide${nextFileIndex}.xml`)) {
      nextFileIndex++
    }
    const slideFileName = `slide${nextFileIndex}.xml`
    const slideZipPath = `ppt/slides/${slideFileName}`

    // Snapshot source XML — never mutate the original slide content
    const sourceXmlSnapshot = this.getSlideXml(sourceIndex)
    logger.debug('Source XML length:', sourceXmlSnapshot ? sourceXmlSnapshot.length : 0)

    const sourceRels = relationshipManager.getRelationships(sourceInfo.zipPath)
    logger.debug('Source Rels Path searched:', relationshipManager.getRelsPath(sourceInfo.zipPath))
    logger.debug('Source Rels found:', sourceRels)

    const idMap = await this.#deepCloneSlideRelationships(
      sourceInfo.zipPath,
      slideZipPath,
      relationshipManager,
      mediaManager,
      sourceRels,
      [REL_TYPES.NOTES_SLIDE]
    )
    logger.debug('Deep-cloned relationship ID map:', Array.from(idMap.entries()))

    let cloneXml = remapRelationshipIds(sourceXmlSnapshot, idMap)
    cloneXml = this.#regenerateTableRowIds(cloneXml)

    this.#zipManager.writeFile(slideZipPath, cloneXml)
    this.#slideXmlCache.set(slideZipPath, cloneXml)

    // Clone notes slide if source slide has one
    const notesRel = sourceRels.find(r => r.type === REL_TYPES.NOTES_SLIDE)
    if (notesRel) {
      const sourceNotesPath = relationshipManager.resolveTarget(sourceInfo.zipPath, notesRel.target)
      let notesXml = this.#zipManager.readCachedFile(sourceNotesPath)
      if (!notesXml) {
        notesXml = await this.#zipManager.readFile(sourceNotesPath)
      }

      if (notesXml) {
        let nextNotesIndex = 1
        while (this.#zipManager.hasFile(`ppt/notesSlides/notesSlide${nextNotesIndex}.xml`)) {
          nextNotesIndex++
        }
        const notesFileName = `notesSlide${nextNotesIndex}.xml`
        const notesZipPath = `ppt/notesSlides/${notesFileName}`

        // Deep-clone relationships from source notes slide (so slide-specific resources like tags are also copied)
        const sourceNotesRels = relationshipManager.getRelationships(sourceNotesPath)
        const notesIdMap = await this.#deepCloneSlideRelationships(
          sourceNotesPath,
          notesZipPath,
          relationshipManager,
          mediaManager,
          sourceNotesRels
        )

        // Update target in notes relationships to point back to the new slide
        // IMPORTANT: patch the back-reference BEFORE flushing to ZIP so the
        // .rels file on disk contains the correct slide path.
        const notesRels = relationshipManager.getRelationships(notesZipPath)
        const slideRel = notesRels.find(
          r => r.type === REL_TYPES.SLIDE || r.type.endsWith('/slide')
        )
        if (slideRel) {
          slideRel.target = `../slides/${slideFileName}`
        }
        // Flush AFTER the target is patched so the written .rels is correct
        relationshipManager.flushRelationships(notesZipPath)

        // Remap relationship IDs in notes XML
        notesXml = remapRelationshipIds(notesXml, notesIdMap)

        // Write the notes slide XML to ZIP
        this.#zipManager.writeFile(notesZipPath, notesXml)

        // Add content type override
        this.#contentTypesManager.addOverride(
          notesZipPath,
          'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
        )

        // Add notes slide relationship to the new slide
        relationshipManager.addRelationship(
          slideZipPath,
          REL_TYPES.NOTES_SLIDE,
          `../notesSlides/${notesFileName}`
        )
      }
    }

    // Add to presentation.xml
    const rId = relationshipManager.addRelationship(
      'ppt/presentation.xml',
      REL_TYPES.SLIDE,
      `slides/${slideFileName}`
    )

    const existingSlideIds = Array.from(this.#slides.values()).map(s => s.slideId)
    const newSlideId = generateSlideId(existingSlideIds)

    const slideInfo = {
      index: newIndex,
      zipPath: slideZipPath,
      relationshipId: rId,
      slideId: newSlideId,
      tags: [...sourceInfo.tags],
      title: sourceInfo.title,
    }

    this.#slides.set(newIndex, slideInfo)
    this.#slideStates.set(newIndex, {
      xmlStr: cloneXml,
      xmlObj: null,
      dirty: false,
      indexBuilt: false,
      shapeMap: new Map(),
      picMap: new Map(),
      tableMap: new Map(),
      chartMap: new Map(),
    })
    this.#addSlideToPresentation(rId, newSlideId, sourceInfo.slideId)
    this.#registerSlideContentType(slideFileName)

    logger.debug(`Cloned slide ${sourceIndex} to new slide ${newIndex}`)
  }

  /**
   * Deep-clones slide relationships so every embedded part (chart, image, VML drawing,
   * SmartArt, custom XML, audio/video, etc.) gets its own independent copy.
   * Shared presentation-level resources (layout, master, theme, tableStyles) are reused.
   *
   * IMPORTANT: Every internal part that is unique to a slide MUST be deep-copied here.
   * Sharing mutable internal parts between slides causes PowerPoint corruption, particularly
   * in enterprise environments with strict validators (e.g. Boldon James Classifier).
   * @private
   */
  async #deepCloneSlideRelationships(
    sourcePath,
    destPath,
    relationshipManager,
    mediaManager,
    sourceRels,
    excludeTypes = []
  ) {
    const idMap = new Map()
    const destRels = []

    for (const rel of sourceRels) {
      if (excludeTypes.includes(rel.type)) continue

      let target = rel.target
      const resolvedTarget =
        rel.targetMode === 'External'
          ? null
          : relationshipManager.resolveTarget(sourcePath, rel.target)
      const typeEnd = rel.type.split('/').pop().toLowerCase()

      if (rel.targetMode === 'External') {
        // External hyperlinks, audio/video links — reuse as-is
      } else if (rel.type === REL_TYPES.CHART) {
        // Charts → deep copy with independent workbook/style/color
        target = await this.#copyChartPart(resolvedTarget, relationshipManager)
      } else if (rel.type === REL_TYPES.IMAGE) {
        // Raster/vector images — copy as new media part
        const newMediaPath = await mediaManager.copyMediaAsNewPart(resolvedTarget)
        target = `../media/${newMediaPath.split('/').pop()}`
      } else if (
        rel.type === REL_TYPES.SLIDE_LAYOUT ||
        rel.type === REL_TYPES.SLIDE_MASTER ||
        rel.type === REL_TYPES.THEME ||
        rel.type === REL_TYPES.TABLE_STYLES ||
        rel.type ===
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster' ||
        rel.type ===
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide' ||
        rel.type === REL_TYPES.SLIDE
      ) {
        // Shared presentation-level resources — reuse same target
      } else if (typeEnd === 'audio' || typeEnd === 'video' || typeEnd === 'media') {
        // Embedded audio/video binary — copy to new media file
        if (this.#zipManager.hasFile(resolvedTarget)) {
          const newPath = await this.#copyBinaryPart(resolvedTarget)
          if (newPath) target = this.#makeRelativeTarget(destPath, newPath)
        }
      } else if (typeEnd === 'vmldrawing') {
        // VML drawings (legacy shapes, comment boxes) — must be independent per slide
        if (this.#zipManager.hasFile(resolvedTarget)) {
          const newTarget = await this.#copyGenericXmlPart(
            resolvedTarget,
            destPath,
            relationshipManager
          )
          if (newTarget) target = newTarget
        }
      } else if (
        typeEnd === 'diagramdata' ||
        typeEnd === 'diagramlayout' ||
        typeEnd === 'diagramquickstyle' ||
        typeEnd === 'diagramcolors' ||
        typeEnd === 'diagram'
      ) {
        // SmartArt/diagram parts — each slide needs its own copy
        if (this.#zipManager.hasFile(resolvedTarget)) {
          const newTarget = await this.#copyGenericXmlPart(
            resolvedTarget,
            destPath,
            relationshipManager
          )
          if (newTarget) target = newTarget
        }
      } else if (typeEnd === 'oleobject' || typeEnd === 'activex') {
        // Non-chart OLE / ActiveX objects — copy binary
        if (this.#zipManager.hasFile(resolvedTarget)) {
          const newPath = await this.#copyBinaryPart(resolvedTarget)
          if (newPath) {
            target = this.#makeRelativeTarget(destPath, newPath)
          }
        }
      } else if (rel.target && !rel.target.startsWith('http') && resolvedTarget) {
        // Catch-all for any other internal part: copy XML if the file extension is .xml,
        // otherwise copy as binary. This handles custom XML (e.g. Boldon James classification
        // parts) and any future unknown relationship types.
        if (this.#zipManager.hasFile(resolvedTarget)) {
          if (resolvedTarget.toLowerCase().endsWith('.xml')) {
            // Text/XML content
            const newTarget = await this.#copyGenericXmlPart(
              resolvedTarget,
              destPath,
              relationshipManager
            )
            if (newTarget) target = newTarget
          } else {
            // Binary content
            const newPath = await this.#copyBinaryPart(resolvedTarget)
            if (newPath) {
              target = this.#makeRelativeTarget(destPath, newPath)
            }
          }
        } else {
          // Part not in ZIP (may be external or already purged) — keep original target
          logger.warn(`Clone: part not found in ZIP, reusing original target: ${resolvedTarget}`)
        }
      }

      const newId = generateRelationshipId(destRels.map(r => r.id))
      const newRel = { id: newId, type: rel.type, target }
      if (rel.targetMode) newRel.targetMode = rel.targetMode
      destRels.push(newRel)
      idMap.set(rel.id, newId)
    }

    relationshipManager.setRelationships(destPath, destRels)
    relationshipManager.flushRelationships(destPath)
    return idMap
  }

  /**
   * Copies a chart part and all of its dependent chart style/color/workbook parts.
   * @private
   * @returns {Promise<string>} Relative target for the cloned slide relationship.
   */
  async #copyChartPart(sourceChartPath, relationshipManager) {
    let chartXml = this.#zipManager.readCachedFile(sourceChartPath)
    if (!chartXml) {
      chartXml = await this.#zipManager.readFile(sourceChartPath)
    }
    if (!chartXml) {
      throw new PPTXError(`Cannot clone slide: chart not found at ${sourceChartPath}`)
    }

    const sourceChartRels = relationshipManager.getRelationships(sourceChartPath)
    let nextChartNum = 1
    while (this.#zipManager.hasFile(`ppt/charts/chart${nextChartNum}.xml`)) {
      nextChartNum++
    }
    const destChartFileName = `chart${nextChartNum}.xml`
    const destChartPath = `ppt/charts/${destChartFileName}`

    const destChartRels = []
    const chartRelIdMap = new Map()

    for (const rel of sourceChartRels) {
      const resolved = relationshipManager.resolveTarget(sourceChartPath, rel.target)
      let newTarget = rel.target

      if (rel.type === REL_TYPES.PACKAGE || rel.target.includes('../embeddings/')) {
        const bytes = await this.#zipManager.readBinaryFile(resolved)
        if (bytes) {
          const fileName = resolved.split('/').pop()
          let nextEmbed = 1
          let destWorkbookPath = `ppt/embeddings/Microsoft_Excel_Worksheet${nextEmbed}.xlsx`
          if (fileName.endsWith('.bin')) {
            destWorkbookPath = `ppt/embeddings/oleObject${nextEmbed}.bin`
          }
          while (this.#zipManager.hasFile(destWorkbookPath)) {
            nextEmbed++
            destWorkbookPath = fileName.endsWith('.bin')
              ? `ppt/embeddings/oleObject${nextEmbed}.bin`
              : `ppt/embeddings/Microsoft_Excel_Worksheet${nextEmbed}.xlsx`
          }
          this.#zipManager.writeBinaryFile(destWorkbookPath, bytes)
          this.#contentTypesManager.addOverride(
            destWorkbookPath,
            fileName.endsWith('.bin') ? ACTIVEX_CONTENT_TYPE : WORKBOOK_CONTENT_TYPE
          )
          newTarget = `../embeddings/${destWorkbookPath.split('/').pop()}`
        }
      } else if (/colors\d+\.xml$/i.test(rel.target)) {
        newTarget = await this.#copyChartSupportXml(resolved, 'colors', CHART_COLORS_CONTENT_TYPE)
      } else if (/style\d+\.xml$/i.test(rel.target)) {
        newTarget = await this.#copyChartSupportXml(resolved, 'style', CHART_STYLE_CONTENT_TYPE)
      }

      const newId = generateRelationshipId(destChartRels.map(r => r.id))
      const newRel = { id: newId, type: rel.type, target: newTarget }
      if (rel.targetMode) newRel.targetMode = rel.targetMode
      destChartRels.push(newRel)
      chartRelIdMap.set(rel.id, newId)
    }

    chartXml = remapRelationshipIds(chartXml, chartRelIdMap)
    this.#zipManager.writeFile(destChartPath, chartXml)
    relationshipManager.setRelationships(destChartPath, destChartRels)
    relationshipManager.flushRelationships(destChartPath)
    this.#contentTypesManager.addOverride(destChartPath, CHART_CONTENT_TYPE)

    return `../charts/${destChartFileName}`
  }

  /**
   * Copies a chart colors/style XML part to a new sequentially numbered file.
   * @private
   */
  async #copyChartSupportXml(sourcePath, prefix, contentType) {
    let content = this.#zipManager.readCachedFile(sourcePath)
    if (!content) {
      content = await this.#zipManager.readFile(sourcePath)
    }
    if (!content) return sourcePath.split('/').pop()

    let nextNum = 1
    let destFileName = `${prefix}${nextNum}.xml`
    while (this.#zipManager.hasFile(`ppt/charts/${destFileName}`)) {
      nextNum++
      destFileName = `${prefix}${nextNum}.xml`
    }
    const destPath = `ppt/charts/${destFileName}`
    this.#zipManager.writeFile(destPath, content)
    this.#contentTypesManager.addOverride(destPath, contentType)
    return destFileName
  }

  /**
   * Copies a generic XML text part to a new sequentially numbered path in the same
   * directory as the source. Returns a relative target path from the destination slide,
   * or null if the copy failed.
   * @private
   * @param {string} sourcePath - Absolute ZIP path of the source XML part.
   * @returns {Promise<string|null>} Relative target for the cloned slide relationship.
   */
  async #copyGenericXmlPart(sourcePath, destPath, relationshipManager) {
    let content = this.#zipManager.readCachedFile(sourcePath)
    if (!content) {
      content = await this.#zipManager.readFile(sourcePath)
    }
    if (!content) return null

    const parts = sourcePath.split('/')
    const dir = parts.slice(0, -1).join('/')
    const fileName = parts[parts.length - 1]
    // Extract base name and extension: e.g. 'vmlDrawing1.xml' → base='vmlDrawing', ext='xml'
    const lastDot = fileName.lastIndexOf('.')
    const ext = lastDot >= 0 ? fileName.slice(lastDot) : ''
    const base = lastDot >= 0 ? fileName.slice(0, lastDot).replace(/\d+$/, '') : fileName

    let num = 1
    let destPartPath = `${dir}/${base}${num}${ext}`
    while (this.#zipManager.hasFile(destPartPath)) {
      num++
      destPartPath = `${dir}/${base}${num}${ext}`
    }

    this.#zipManager.writeFile(destPartPath, content)

    // Mirror content type if one exists for the source
    const srcContentType = this.#contentTypesManager.getOverrideContentType(sourcePath)
    if (srcContentType) {
      this.#contentTypesManager.addOverride(destPartPath, srcContentType)
    }

    // Clone any relationships/dependencies of this XML part recursively
    await this.#copyPartDependencies(sourcePath, destPartPath, relationshipManager)

    logger.debug(`Cloned XML part: ${sourcePath} → ${destPartPath}`)
    return this.#makeRelativeTarget(destPath, destPartPath)
  }

  /**
   * Recursively copies relationships and dependency parts for a cloned XML part.
   * @private
   * @param {string} sourcePartPath - Absolute ZIP path of the source XML part.
   * @param {string} destPartPath - Absolute ZIP path of the destination XML part.
   * @param {RelationshipManager} relationshipManager
   * @returns {Promise<void>}
   */
  async #copyPartDependencies(sourcePartPath, destPartPath, relationshipManager) {
    const rels = relationshipManager.getRelationships(sourcePartPath)
    if (!rels || rels.length === 0) return

    // 1. Copy relationships file first (sets up destPath rels list)
    relationshipManager.copyRelationships(sourcePartPath, destPartPath)

    // 2. Iterate and clone each target part
    const destRels = relationshipManager.getRelationships(destPartPath)
    let dirty = false

    for (const rel of destRels) {
      if (rel.targetMode === 'External' || !rel.target || rel.target.startsWith('http')) continue

      const resolved = relationshipManager.resolveTarget(sourcePartPath, rel.target)
      if (!this.#zipManager.hasFile(resolved)) continue

      let newTarget = rel.target
      if (resolved.toLowerCase().endsWith('.xml')) {
        // XML part - copy recursively
        newTarget = await this.#copyGenericXmlPart(resolved, destPartPath, relationshipManager)
      } else {
        // Binary part
        const newBinaryPath = await this.#copyBinaryPart(resolved)
        if (newBinaryPath) {
          newTarget = this.#makeRelativeTarget(destPartPath, newBinaryPath)
        }
      }

      if (newTarget && newTarget !== rel.target) {
        rel.target = newTarget
        dirty = true
      }
    }

    if (dirty) {
      relationshipManager.flushRelationships(destPartPath)
    }
  }

  /**
   * Copies a binary part (audio, video, OLE object, etc.) to a new path in the same
   * directory as the source. Returns the new absolute ZIP path, or null on failure.
   * @private
   * @param {string} sourcePath - Absolute ZIP path of the source binary part.
   * @returns {Promise<string|null>} New absolute ZIP path.
   */
  async #copyBinaryPart(sourcePath) {
    const data = await this.#zipManager.readBinaryFile(sourcePath)
    if (!data) return null

    const parts = sourcePath.split('/')
    const dir = parts.slice(0, -1).join('/')
    const fileName = parts[parts.length - 1]
    const lastDot = fileName.lastIndexOf('.')
    const ext = lastDot >= 0 ? fileName.slice(lastDot) : ''
    const base = lastDot >= 0 ? fileName.slice(0, lastDot).replace(/\d+$/, '') : fileName

    let num = 1
    let destPath = `${dir}/${base}${num}${ext}`
    while (this.#zipManager.hasFile(destPath)) {
      num++
      destPath = `${dir}/${base}${num}${ext}`
    }

    this.#zipManager.writeBinaryFile(destPath, data)

    const srcContentType = this.#contentTypesManager.getOverrideContentType(sourcePath)
    if (srcContentType) {
      this.#contentTypesManager.addOverride(destPath, srcContentType)
    }

    logger.debug(`Cloned binary part: ${sourcePath} → ${destPath}`)
    return destPath
  }

  /**
   * Builds a relative target path from a source ZIP part to a destination ZIP path.
   * @private
   * @param {string} fromPath - Absolute path of the part containing the relationship.
   * @param {string} toPath - Absolute ZIP path of the target.
   * @returns {string} Relative path from fromPath's directory to toPath.
   */
  #makeRelativeTarget(fromPath, toPath) {
    const fromDir = fromPath.split('/').slice(0, -1)
    const toParts = toPath.split('/')
    // Find common prefix length
    let common = 0
    while (
      common < fromDir.length &&
      common < toParts.length - 1 &&
      fromDir[common] === toParts[common]
    ) {
      common++
    }
    const ups = fromDir.length - common
    const downs = toParts.slice(common)
    return (ups > 0 ? '../'.repeat(ups) : './') + downs.join('/')
  }

  /**
   * Assigns fresh a16:rowId values in cloned slide XML so table rows are independent.
   * @private
   */
  #regenerateTableRowIds(xml) {
    const used = new Set()
    return xml.replace(/(<a16:rowId[^>]*\sval=")([^"]+)(")/g, (_match, prefix, _oldVal, suffix) => {
      let newVal
      do {
        newVal = String(Math.floor(Math.random() * 0xffffffff))
      } while (used.has(newVal))
      used.add(newVal)
      return `${prefix}${newVal}${suffix}`
    })
  }

  /**
   * Removes a slide from the presentation.
   *
   * @param {number} slideIndex - 1-based index.
   */
  removeSlide(slideIndex) {
    this.#assertSlideExists(slideIndex)
    const info = this.#slides.get(slideIndex)

    // 1. Check for notes slide relationship and remove it first
    const slideRels = this.#relationshipManager.getRelationships(info.zipPath)
    const notesRel = slideRels.find(r => r.type === REL_TYPES.NOTES_SLIDE)
    if (notesRel) {
      const notesPath = this.#relationshipManager.resolveTarget(info.zipPath, notesRel.target)

      // Remove from ZIP
      this.#zipManager.removeFile(notesPath)

      // Remove its relationships file — use getRelsPath for robustness
      const notesRelsPath = this.#relationshipManager.getRelsPath(notesPath)
      this.#zipManager.removeFile(notesRelsPath)

      // Remove relationships from cache
      this.#relationshipManager.deleteRelationships(notesPath)

      // Remove content type override
      this.#contentTypesManager.removeOverride(notesPath)
    }

    // 2. Remove slide XML from ZIP
    this.#zipManager.removeFile(info.zipPath)

    // 3. Remove slide .rels file — use getRelsPath so non-standard paths work correctly
    const slideRelsPath = this.#relationshipManager.getRelsPath(info.zipPath)
    this.#zipManager.removeFile(slideRelsPath)

    // 4. Remove slide relationship cache entry
    this.#relationshipManager.deleteRelationships(info.zipPath)

    // 5. Remove slide XML cache entries
    this.#slideXmlCache.delete(info.zipPath)
    this.#slideStates.delete(slideIndex)

    // 6. Remove content type from [Content_Types].xml
    this.#contentTypesManager.removeOverride(info.zipPath)

    // 7. Remove relationship from presentation.xml.rels
    this.#relationshipManager.removeRelationship('ppt/presentation.xml', info.relationshipId)

    // 8. Remove from presentation.xml sldIdLst and sections BEFORE reindexing.
    //    This ensures the sldIdLst is consistent with the relationship list.
    this.#removeSlideFromPresentation(info.slideId)

    // 9. Remove from slides map and reindex remaining slides
    this.#slides.delete(slideIndex)
    this.#reindexSlides()

    // 10. Rebuild sldIdLst and synchronize sections to reflect the updated slide order.
    //     This is essential: after reindexing, the relationship IDs and slide IDs are
    //     still valid but the sldIdLst order must be rebuilt so sections match.
    this.rebuildPresentationSlideOrder()

    logger.debug(`Removed slide ${slideIndex}`)
  }

  /**
   * Reorders slides to match the given order array.
   *
   * @param {number[]} order - Array of 1-based slide numbers in desired order.
   */
  reorderSlides(order) {
    const current = this.getAllSlideIndices()
    if (order.length !== current.length) {
      throw new PPTXError(
        `reorderSlides: order array length (${order.length}) must match slide count (${current.length})`
      )
    }

    const slidesCopy = new Map(this.#slides)
    const statesCopy = new Map(this.#slideStates)
    this.#slides.clear()
    this.#slideStates.clear()

    order.forEach((oldIndex, newPos) => {
      const info = slidesCopy.get(oldIndex)
      if (!info) throw new SlideNotFoundError(`Slide ${oldIndex} not found`)
      info.index = newPos + 1
      this.#slides.set(newPos + 1, info)

      const state = statesCopy.get(oldIndex)
      this.#slideStates.set(newPos + 1, state)
    })

    // Rebuild presentation sldIdLst
    this.rebuildPresentationSlideOrder()
    logger.debug(`Reordered slides: [${order.join(', ')}]`)
  }

  /**
   * Resolves a slide reference (index, slideId string, or tag string) to SlideInfo.
   *
   * @param {number|string} slideRef
   * @returns {SlideInfo|null}
   */
  resolveSlideInfo(slideRef) {
    let index
    if (typeof slideRef === 'number') {
      index = slideRef
    } else {
      // 1. Try finding by slideId string
      for (const info of this.#slides.values()) {
        if (info.slideId === String(slideRef)) {
          return info
        }
      }
      // 2. Try finding by tag
      try {
        const indices = this.resolveSlideRef(slideRef)
        if (indices && indices.length > 0) {
          index = indices[0]
        }
      } catch (e) {
        // Fallback: parse as slide index
        const parsedNum = parseInt(slideRef, 10)
        if (!isNaN(parsedNum)) {
          index = parsedNum
        }
      }
    }

    if (index !== undefined) {
      return this.#slides.get(index) || null
    }
    return null
  }

  /**
   * Imports a slide from another PPTX template (PPTXTemplater instance).
   * Preserves all relationships: layouts, media, charts, workbooks, etc.
   *
   * @param {PPTXTemplater} sourceEngine - Source presentation.
   * @param {number|string} slideRef - Slide index (1-based), slide ID, or custom tag.
   * @param {MediaManager} mediaManager - Destination media manager.
   * @returns {Promise<number>} Index of the imported slide.
   */
  async importSlide(sourceEngine, slideRef, mediaManager) {
    const sourceSlideManager = sourceEngine.slideManager
    const sourceRelManager = sourceEngine.relationshipManager
    const sourceZip = sourceEngine.zipManager

    const sourceSlideInfo = sourceSlideManager.resolveSlideInfo(slideRef)
    if (!sourceSlideInfo) {
      throw new SlideNotFoundError(`Source slide "${slideRef}" not found`)
    }

    const newIndex = this.#slides.size + 1
    let nextFileIndex = 1
    while (this.#zipManager.hasFile(`ppt/slides/slide${nextFileIndex}.xml`)) {
      nextFileIndex++
    }
    const slideFileName = `slide${nextFileIndex}.xml`
    const slideZipPath = `ppt/slides/${slideFileName}`

    // Read the source slide's XML
    let slideXml = await sourceSlideManager.getSlideXmlAsync(sourceSlideInfo.index)

    // Get relationships from the source slide
    const sourceRels = sourceRelManager.getRelationships(sourceSlideInfo.zipPath)

    // Map to track old rId -> new rId in the destination slide's .rels file
    const idMap = new Map()

    const EXT_TO_MIME_LOCAL = {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      webp: 'image/webp',
      bmp: 'image/bmp',
      xml: 'application/xml',
      rels: 'application/vnd.openxmlformats-package.relationships+xml',
    }

    for (const rel of sourceRels) {
      const resolvedTarget = sourceRelManager.resolveTarget(sourceSlideInfo.zipPath, rel.target)

      if (rel.type === REL_TYPES.SLIDE_LAYOUT) {
        // Map to destination's slide layout.
        const layoutFileName = rel.target.split('/').pop()
        const destLayoutPath = `ppt/slideLayouts/${layoutFileName}`

        let targetLayout = `../slideLayouts/${layoutFileName}`
        if (!this.#zipManager.hasFile(destLayoutPath)) {
          // Find first available layout
          const layoutFiles = this.#zipManager
            .listFiles('ppt/slideLayouts/')
            .filter(f => f.endsWith('.xml'))
          if (layoutFiles.length > 0) {
            targetLayout = `../slideLayouts/${layoutFiles[0].split('/').pop()}`
          } else {
            targetLayout = '../slideLayouts/slideLayout1.xml'
          }
        }

        const newRId = this.#relationshipManager.addRelationship(
          slideZipPath,
          rel.type,
          targetLayout
        )
        idMap.set(rel.id, newRId)
      } else if (rel.type === REL_TYPES.IMAGE) {
        // Copy media file
        const mediaBytes = await sourceZip.readBinaryFile(resolvedTarget)
        if (mediaBytes) {
          const destMediaZipPath = await mediaManager.embedImage(mediaBytes)
          const relativeMediaTarget = `../media/${destMediaZipPath.split('/').pop()}`
          const newRId = this.#relationshipManager.addRelationship(
            slideZipPath,
            rel.type,
            relativeMediaTarget
          )
          idMap.set(rel.id, newRId)
        }
      } else if (rel.type === REL_TYPES.CHART) {
        // Copy chart XML and its relationships
        const chartXml = await sourceZip.readFile(resolvedTarget)
        if (chartXml) {
          const chartRels = sourceRelManager.getRelationships(resolvedTarget)

          let nextChartId = 1
          while (this.#zipManager.hasFile(`ppt/charts/chart${nextChartId}.xml`)) {
            nextChartId++
          }
          const destChartZipPath = `ppt/charts/chart${nextChartId}.xml`
          const chartFileName = `chart${nextChartId}.xml`

          // Handle workbook packages within charts
          for (const chartRel of chartRels) {
            const resolvedChartTarget = sourceRelManager.resolveTarget(
              resolvedTarget,
              chartRel.target
            )
            const workbookBytes = await sourceZip.readBinaryFile(resolvedChartTarget)

            if (workbookBytes) {
              const workbookFileName = resolvedChartTarget.split('/').pop()
              let nextEmbedId = 1
              let destWorkbookZipPath = `ppt/embeddings/Microsoft_Excel_Worksheet${nextEmbedId}.xlsx`
              if (workbookFileName.endsWith('.bin')) {
                destWorkbookZipPath = `ppt/embeddings/oleObject${nextEmbedId}.bin`
              }
              while (this.#zipManager.hasFile(destWorkbookZipPath)) {
                nextEmbedId++
                destWorkbookZipPath = workbookFileName.endsWith('.bin')
                  ? `ppt/embeddings/oleObject${nextEmbedId}.bin`
                  : `ppt/embeddings/Microsoft_Excel_Worksheet${nextEmbedId}.xlsx`
              }

              this.#zipManager.writeBinaryFile(destWorkbookZipPath, workbookBytes)

              const workbookContentType = workbookFileName.endsWith('.bin')
                ? 'application/vnd.ms-office.activeX'
                : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
              this.#contentTypesManager.addOverride(destWorkbookZipPath, workbookContentType)

              const relativeWorkbookPath = `../embeddings/${destWorkbookZipPath.split('/').pop()}`
              this.#relationshipManager.addRelationship(
                destChartZipPath,
                chartRel.type,
                relativeWorkbookPath
              )
            }
          }

          this.#zipManager.writeFile(destChartZipPath, chartXml)
          this.#contentTypesManager.addOverride(
            destChartZipPath,
            'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
          )

          const relativeChartPath = `../charts/${chartFileName}`
          const newRId = this.#relationshipManager.addRelationship(
            slideZipPath,
            rel.type,
            relativeChartPath
          )
          idMap.set(rel.id, newRId)
        }
      } else if (rel.type === REL_TYPES.HYPERLINK) {
        const newRId = this.#relationshipManager.addRelationship(
          slideZipPath,
          rel.type,
          rel.target,
          rel.targetMode
        )
        idMap.set(rel.id, newRId)
      } else {
        // Fallback for notes, themes, styles or custom XML
        if (rel.target && !rel.target.startsWith('http')) {
          const targetBytes = await sourceZip.readBinaryFile(resolvedTarget)
          if (targetBytes && !this.#zipManager.hasFile(resolvedTarget)) {
            this.#zipManager.writeBinaryFile(resolvedTarget, targetBytes)
            const ext = resolvedTarget.split('.').pop().toLowerCase()
            const mime = EXT_TO_MIME_LOCAL[ext] || 'application/octet-stream'
            this.#contentTypesManager.addDefault(ext, mime)
          }
        }
        const newRId = this.#relationshipManager.addRelationship(
          slideZipPath,
          rel.type,
          rel.target,
          rel.targetMode
        )
        idMap.set(rel.id, newRId)
      }
    }

    // Remap all relationship IDs inside the imported slide XML
    slideXml = remapRelationshipIds(slideXml, idMap)

    // Save the remapped slide XML to ZIP
    this.#zipManager.writeFile(slideZipPath, slideXml)
    this.#slideXmlCache.set(slideZipPath, slideXml)

    // Generate unique Slide ID
    const existingIds = Array.from(this.#slides.values()).map(s => parseInt(s.slideId, 10))
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 255
    const newSlideId = String(maxId + 1)

    // Add relationship from presentation.xml
    const rId = this.#relationshipManager.addRelationship(
      'ppt/presentation.xml',
      REL_TYPES.SLIDE,
      `slides/${slideFileName}`
    )

    const slideInfo = {
      index: newIndex,
      zipPath: slideZipPath,
      relationshipId: rId,
      slideId: newSlideId,
      tags: [...sourceSlideInfo.tags],
      title: sourceSlideInfo.title || '',
    }

    this.#slides.set(newIndex, slideInfo)
    this.#slideStates.set(newIndex, {
      xmlStr: slideXml,
      xmlObj: null,
      dirty: false,
      indexBuilt: false,
      shapeMap: new Map(),
      picMap: new Map(),
      tableMap: new Map(),
      chartMap: new Map(),
    })

    // Add entry in presentation.xml sldIdLst
    this.#addSlideToPresentation(rId, newSlideId)

    // Register slide in content types
    this.#registerSlideContentType(slideFileName)

    logger.debug(`Successfully imported slide "${slideRef}" to index ${newIndex}`)
    return newIndex
  }

  /**
   * Exports a subset of slides to a new PPTXTemplater.
   *
   * @param {number[]} slideIndices - 1-based slide indices to export.
   * @param {PPTXTemplater} sourceEngine - Current PPTXTemplater instance.
   * @returns {Promise<PPTXTemplater>}
   */
  async exportSlides(slideIndices, sourceEngine) {
    // Lazy import to avoid circular dep
    const { PPTXTemplater } = require('../core/PPTXTemplater.js')

    // Create a blank new PPTX
    const newEngine = await PPTXTemplater.create()

    // Remove the default slides from the blank template to avoid orphans
    const defaultSlides = newEngine.slideManager.getAllSlideIndices()
    for (const dIdx of defaultSlides.reverse()) {
      newEngine.slideManager.removeSlide(dIdx)
    }

    // Copy selected slides into the new engine
    for (const idx of slideIndices) {
      this.#assertSlideExists(idx)
      await newEngine.slideManager.importSlide(sourceEngine, idx, newEngine.mediaManager)
    }

    return newEngine
  }

  /**
   * Validates the slide structure.
   *
   * @param {RelationshipManager} relationshipManager
   * @param {ZipManager} zipManager
   * @returns {ValidationResult}
   */
  validateStructure(relationshipManager, zipManager) {
    const errors = []
    const warnings = []

    for (const [index, info] of this.#slides) {
      if (!zipManager.hasFile(info.zipPath)) {
        errors.push(`Slide ${index}: XML file missing at ${info.zipPath}`)
      }

      const rels = relationshipManager.getRelationships(info.zipPath)
      const layoutRel = rels.find(r => r.type === REL_TYPES.SLIDE_LAYOUT)
      if (!layoutRel) {
        warnings.push(`Slide ${index}: No slide layout relationship found`)
      }
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings,
    }
  }

  /**
   * Pre-loads all slide XML into cache (for bulk operations).
   * @returns {Promise<void>}
   */
  async preloadAll() {
    await Promise.all(
      this.getAllSlideIndices().map(async i => {
        await this.getSlideXmlAsync(i)

        // Also check and preload notes slide XML
        const info = this.#slides.get(i)
        const rels = this.#relationshipManager.getRelationships(info.zipPath)
        const notesRel = rels.find(r => r.type === REL_TYPES.NOTES_SLIDE)
        if (notesRel) {
          const notesPath = this.#relationshipManager.resolveTarget(info.zipPath, notesRel.target)
          await this.#zipManager.readFile(notesPath)
        }
      })
    )
  }

  /**
   * Adds a slide ID to sections in presentation.xml.
   * @private
   */
  #addSlideToSections(slideId, sourceSlideId = null) {
    if (!this.#presentationObj) return
    const extLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:extLst')
    if (!extLst?.['p:ext']) return

    const exts = Array.isArray(extLst['p:ext']) ? extLst['p:ext'] : [extLst['p:ext']]
    for (const ext of exts) {
      const sectionLst = ext['p14:sectionLst']
      if (!sectionLst?.['p14:section']) continue

      const sections = sectionLst['p14:section']
      const targetIdStr = String(slideId)
      const sourceIdStr = sourceSlideId ? String(sourceSlideId) : null

      if (sourceIdStr) {
        for (const section of sections) {
          const sldIdLst = section['p14:sldIdLst']
          if (!sldIdLst?.['p14:sldId']) continue

          const sldIds = sldIdLst['p14:sldId']
          const idx = sldIds.findIndex(s => String(s['@_id']) === sourceIdStr)
          if (idx !== -1) {
            logger.debug(
              `Inserting slide ${targetIdStr} after slide ${sourceIdStr} in section "${section['@_name']}"`
            )
            sldIds.splice(idx + 1, 0, { '@_id': targetIdStr })
            return
          }
        }
      }

      if (sections.length > 0) {
        const lastSection = sections[sections.length - 1]
        if (!lastSection['p14:sldIdLst']) {
          lastSection['p14:sldIdLst'] = { 'p14:sldId': [] }
        }
        if (!lastSection['p14:sldIdLst']['p14:sldId']) {
          lastSection['p14:sldIdLst']['p14:sldId'] = []
        }
        const sldIds = lastSection['p14:sldIdLst']['p14:sldId']
        logger.debug(`Appending slide ${targetIdStr} to last section "${lastSection['@_name']}"`)
        sldIds.push({ '@_id': targetIdStr })
      }
    }
  }

  /**
   * Updates the presentation.xml sldIdLst with a new slide entry.
   * @private
   */
  #addSlideToPresentation(rId, slideId, sourceSlideId = null) {
    if (!this.#presentationObj) return

    let sldIdLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:sldIdLst')
    if (!sldIdLst) {
      this.#xmlParser.setNode(this.#presentationObj, 'p:presentation.p:sldIdLst', { 'p:sldId': [] })
      sldIdLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:sldIdLst')
    }

    if (!sldIdLst['p:sldId']) sldIdLst['p:sldId'] = []
    if (!Array.isArray(sldIdLst['p:sldId'])) {
      sldIdLst['p:sldId'] = [sldIdLst['p:sldId']]
    }

    sldIdLst['p:sldId'].push({ '@_id': slideId, '@_r:id': rId })
    this.#addSlideToSections(slideId, sourceSlideId)
    this.#flushPresentation()
  }

  /**
   * Removes a slide from presentation.xml sldIdLst.
   * @private
   */
  #removeSlideFromPresentation(slideId) {
    if (!this.#presentationObj) return
    const sldIdLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:sldIdLst')
    if (!sldIdLst?.['p:sldId']) return

    const targetIdStr = String(slideId)
    sldIdLst['p:sldId'] = (
      Array.isArray(sldIdLst['p:sldId']) ? sldIdLst['p:sldId'] : [sldIdLst['p:sldId']]
    ).filter(s => String(s['@_id']) !== targetIdStr)

    // Also remove from any PowerPoint sections
    this.#removeSlideFromSections(slideId)

    this.#flushPresentation()
  }

  /**
   * Removes a slide ID from all sections in presentation.xml.
   * @private
   * @param {string} slideId - Unique slide ID.
   */
  #removeSlideFromSections(slideId) {
    if (!this.#presentationObj) return
    const extLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:extLst')
    if (!extLst?.['p:ext']) return

    const exts = Array.isArray(extLst['p:ext']) ? extLst['p:ext'] : [extLst['p:ext']]
    for (const ext of exts) {
      const sectionLst = ext['p14:sectionLst']
      if (!sectionLst?.['p14:section']) continue

      const sections = sectionLst['p14:section'] // Guaranteed to be array by XMLParser config

      for (const section of sections) {
        const sldIdLst = section['p14:sldIdLst']
        if (!sldIdLst?.['p14:sldId']) continue

        const sldIds = sldIdLst['p14:sldId'] // Guaranteed to be array by XMLParser config
        const targetIdStr = String(slideId)
        const filtered = sldIds.filter(s => String(s['@_id']) !== targetIdStr)

        if (filtered.length !== sldIds.length) {
          logger.debug(`Removing slide ${targetIdStr} from section "${section['@_name']}"`)
          section['p14:sldIdLst']['p14:sldId'] = filtered
        }
      }
    }
  }

  /**
   * Rebuilds presentation.xml sldIdLst in the current slide order.
   */
  rebuildPresentationSlideOrder() {
    if (!this.#presentationObj) return
    const sldIdLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:sldIdLst')
    if (!sldIdLst) return

    const ordered = this.getAllSlideIndices().map(i => {
      const info = this.#slides.get(i)
      return { '@_id': info.slideId, '@_r:id': info.relationshipId }
    })

    sldIdLst['p:sldId'] = ordered

    // Synchronize sections order with the new slide order
    const orderedSlideIds = ordered.map(o => String(o['@_id']))
    const extLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:extLst')
    if (extLst?.['p:ext']) {
      const exts = Array.isArray(extLst['p:ext']) ? extLst['p:ext'] : [extLst['p:ext']]
      for (const ext of exts) {
        const sectionLst = ext['p14:sectionLst']
        if (sectionLst?.['p14:section']) {
          const sections = Array.isArray(sectionLst['p14:section'])
            ? sectionLst['p14:section']
            : [sectionLst['p14:section']]

          // 1. Identify section anchors (first slide of each section)
          const sectionAnchors = []
          for (const section of sections) {
            const sldIdLstObj = section['p14:sldIdLst']
            if (sldIdLstObj?.['p14:sldId']) {
              const sldIds = Array.isArray(sldIdLstObj['p14:sldId'])
                ? sldIdLstObj['p14:sldId']
                : [sldIdLstObj['p14:sldId']]
              // Find the first valid slide ID in this section that is still present in the ordered list
              let anchorId = null
              for (const sldId of sldIds) {
                if (sldId && sldId['@_id']) {
                  const idStr = String(sldId['@_id'])
                  if (orderedSlideIds.includes(idStr)) {
                    anchorId = idStr
                    break
                  }
                }
              }
              sectionAnchors.push({ section, anchorId })
            } else {
              sectionAnchors.push({ section, anchorId: null })
            }
          }

          // 2. Clear the slide lists of all sections
          for (const section of sections) {
            if (!section['p14:sldIdLst']) {
              section['p14:sldIdLst'] = { 'p14:sldId': [] }
            } else {
              section['p14:sldIdLst']['p14:sldId'] = []
            }
          }

          // 3. Populate sections by tracing ordered slide IDs and switching sections when an anchor is reached
          let currentSection = sections[0] || null

          for (const slideId of orderedSlideIds) {
            // Check if this slideId is the anchor of another section
            const matchingAnchor = sectionAnchors.find(sa => sa.anchorId === slideId)
            if (matchingAnchor) {
              currentSection = matchingAnchor.section
            }

            if (currentSection) {
              if (!currentSection['p14:sldIdLst']) {
                currentSection['p14:sldIdLst'] = { 'p14:sldId': [] }
              }
              currentSection['p14:sldIdLst']['p14:sldId'].push({ '@_id': slideId })
            }
          }
        }
      }
    }

    this.#flushPresentation()
  }

  /**
   * Re-indexes slide map after a removal.
   * @private
   */
  #reindexSlides() {
    const sorted = Array.from(this.#slides.entries()).sort(([a], [b]) => a - b)
    this.#slides.clear()

    const sortedStates = Array.from(this.#slideStates.entries()).sort(([a], [b]) => a - b)
    this.#slideStates.clear()

    sorted.forEach(([, info], i) => {
      info.index = i + 1
      this.#slides.set(i + 1, info)
    })

    sortedStates.forEach(([, state], i) => {
      this.#slideStates.set(i + 1, state)
    })
  }

  /**
   * Registers a new slide in [Content_Types].xml.
   * @private
   */
  #registerSlideContentType(slideFileName) {
    this.#contentTypesManager.addOverride(`ppt/slides/${slideFileName}`, SLIDE_CONTENT_TYPE)
  }

  /**
   * Writes the updated presentation.xml back to the ZIP.
   * @private
   */
  #flushPresentation() {
    if (!this.#presentationObj || !this.#zipManager) return
    const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    const xml = this.#xmlParser.build(this.#presentationObj, declaration)
    this.#zipManager.writeFile('ppt/presentation.xml', xml)
  }

  /**
   * Initializes a blank presentation.xml structure.
   * @private
   */
  async #initializeBlankPresentation() {
    // Used when creating from scratch
    this.#presentationObj = {
      'p:presentation': {
        '@_xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@_xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        '@_xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'p:sldMasterIdLst': {},
        'p:sldIdLst': {},
        'p:sldSz': { '@_cx': '9144000', '@_cy': '5143500' },
        'p:notesSz': { '@_cx': '6858000', '@_cy': '9144000' },
      },
    }
    this.#flushPresentation()
  }

  /**
   * Duplicates a slide.
   *
   * @param {number} slideIndex
   * @param {number} [atPosition]
   * @param {RelationshipManager} relationshipManager
   * @param {MediaManager} mediaManager
   * @returns {Promise<number>}
   */
  async duplicateSlide(slideIndex, atPosition, relationshipManager, mediaManager) {
    const promise = this.#duplicateSlideInternal(
      slideIndex,
      atPosition,
      relationshipManager,
      mediaManager
    )
    this.#zipManager.addPendingPromise(promise)
    return promise
  }

  /**
   * Internal duplicate implementation.
   * @private
   */
  async #duplicateSlideInternal(slideIndex, atPosition, relationshipManager, mediaManager) {
    await this.#cloneSlideInternal(slideIndex, null, relationshipManager, mediaManager)
    const count = this.slideCount
    if (atPosition !== undefined && atPosition !== count) {
      const order = []
      for (let i = 1; i < count; i++) {
        order.push(i)
      }
      order.splice(atPosition - 1, 0, count)
      this.reorderSlides(order)
      return atPosition
    }
    return count
  }

  /**
   * Moves a slide to a new position.
   *
   * @param {number} fromIndex
   * @param {number} toIndex
   */
  moveSlide(fromIndex, toIndex) {
    this.#assertSlideExists(fromIndex)
    if (toIndex < 1 || toIndex > this.slideCount) {
      throw new PPTXError(`Destination index ${toIndex} out of bounds`)
    }
    const order = this.getAllSlideIndices()
    const [removed] = order.splice(fromIndex - 1, 1)
    order.splice(toIndex - 1, 0, removed)
    this.reorderSlides(order)
  }

  /**
   * Inserts a new slide at a specific index.
   *
   * @param {number} slideIndex
   * @param {Object} options
   * @param {RelationshipManager} relationshipManager
   * @param {MediaManager} mediaManager
   */
  insertSlide(slideIndex, options, relationshipManager, mediaManager) {
    this.addNewSlide(options, relationshipManager, mediaManager)
    const count = this.slideCount
    if (slideIndex !== undefined && slideIndex !== count) {
      const order = []
      for (let i = 1; i < count; i++) {
        order.push(i)
      }
      order.splice(slideIndex - 1, 0, count)
      this.reorderSlides(order)
    }
  }

  /**
   * Gets all slides.
   *
   * @returns {SlideInfo[]}
   */
  getSlides() {
    return this.getAllSlideInfo()
  }

  /**
   * Asserts a slide index is valid.
   * @private
   * @param {number} index - 1-based slide index.
   */
  #assertSlideExists(index) {
    if (!this.#slides.has(index)) {
      throw new SlideNotFoundError(
        `Slide ${index} does not exist. Total slides: ${this.#slides.size}`
      )
    }
  }

  getSlideObj(slideIndex) {
    this.#assertSlideExists(slideIndex)
    const state = this.#slideStates.get(slideIndex)

    if (!state.xmlObj) {
      const xml = this.getSlideXml(slideIndex)
      const info = this.#slides.get(slideIndex)
      state.xmlObj = this.#xmlParser.parse(xml, info.zipPath.split('/').pop())
      state.indexBuilt = false
    }

    if (!state.indexBuilt) {
      state.shapeMap.clear()
      state.picMap.clear()
      state.tableMap.clear()
      state.chartMap.clear()

      const spTree =
        state.xmlObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
        state.xmlObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
        state.xmlObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

      this.#buildSlideIndexRecursive(
        spTree,
        state.shapeMap,
        state.picMap,
        state.tableMap,
        state.chartMap
      )
      state.indexBuilt = true
    }

    return state.xmlObj
  }

  #buildSlideIndexRecursive(container, shapeMap, picMap, tableMap, chartMap) {
    if (!container) return

    // Shapes
    let shapes = container['p:sp'] || []
    if (!Array.isArray(shapes)) shapes = [shapes]
    for (const shape of shapes) {
      const cNvPr = shape?.['p:nvSpPr']?.['p:cNvPr']
      if (cNvPr) {
        const name = cNvPr['@_name']
        const id = String(cNvPr['@_id'])
        const entry = { shape, parent: container, type: 'sp' }
        if (name) shapeMap.set(name, entry)
        if (id) shapeMap.set(id, entry)
      }
    }

    // Pictures
    let pics = container['p:pic'] || []
    if (!Array.isArray(pics)) pics = [pics]
    for (const pic of pics) {
      const cNvPr = pic?.['p:nvPicPr']?.['p:cNvPr']
      if (cNvPr) {
        const name = cNvPr['@_name']
        const id = String(cNvPr['@_id'])
        const embedId = pic?.['p:blipFill']?.['a:blip']?.['@_r:embed']
        const entry = { pic, parent: container, type: 'pic' }
        if (name) picMap.set(name, entry)
        if (id) picMap.set(id, entry)
        if (embedId) picMap.set(embedId, entry)
      }
    }

    // Graphic frames
    let frames = container['p:graphicFrame'] || []
    if (!Array.isArray(frames)) frames = [frames]
    for (const frame of frames) {
      const cNvPr = frame?.['p:nvGraphicFramePr']?.['p:cNvPr']
      const name = cNvPr ? cNvPr['@_name'] : null
      const id = cNvPr ? String(cNvPr['@_id']) : null

      const tbl = frame?.['a:graphic']?.['a:graphicData']?.['a:tbl']
      if (tbl) {
        const entry = { table: tbl, frame, parent: container, type: 'table' }
        if (name) tableMap.set(name, entry)
        if (id) tableMap.set(id, entry)
      }

      const chart = frame?.['a:graphic']?.['a:graphicData']?.['c:chart']
      if (chart) {
        const embedId = chart['@_r:id']
        const entry = { chart, frame, parent: container, type: 'chart' }
        if (name) chartMap.set(name, entry)
        if (id) chartMap.set(id, entry)
        if (embedId) chartMap.set(embedId, entry)
      }
    }

    // Groups
    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const group of groups) {
      const cNvPr = group?.['p:nvGrpSpPr']?.['p:cNvPr']
      if (cNvPr) {
        const name = cNvPr['@_name']
        const id = String(cNvPr['@_id'])
        const entry = { shape: group, parent: container, type: 'grpSp' }
        if (name) shapeMap.set(name, entry)
        if (id) shapeMap.set(id, entry)
      }
      this.#buildSlideIndexRecursive(group, shapeMap, picMap, tableMap, chartMap)
    }
  }

  flush() {
    for (const [index, state] of this.#slideStates) {
      if (state.dirty && state.xmlObj) {
        this.getSlideXml(index)
      }
    }
  }

  /**
   * Normalizes the slide filenames, slide IDs, and relationship IDs on export.
   * Ensures slide filenames are strictly sequential (slide1.xml, slide2.xml, ...)
   * and match their visual order, and updates all relationships and content types.
   * Also normalizes notes slide filenames to match slide order.
   *
   * @param {RelationshipManager} relationshipManager
   * @param {ContentTypesManager} contentTypesManager
   */
  async normalizeStructure(relationshipManager, contentTypesManager) {
    const slides = this.getAllSlideInfo()
    if (slides.length === 0) return

    // 1. Map old path -> new path, old filename -> new filename
    const pathMap = new Map()
    const nameMap = new Map()

    slides.forEach((info, idx) => {
      const newPath = `ppt/slides/slide${idx + 1}.xml`
      pathMap.set(info.zipPath, newPath)

      const oldName = info.zipPath.split('/').pop()
      const newName = `slide${idx + 1}.xml`
      nameMap.set(oldName, newName)
    })

    // Read all slide XML contents and relationships into memory first
    const slideData = []
    for (const info of slides) {
      const xml = this.getSlideXml(info.index)
      const rels = relationshipManager.getRelationships(info.zipPath)
      slideData.push({
        info,
        xml,
        rels: [...rels],
      })
    }

    // Phase 1: Remove all old files from ZIP and clear their cache entries
    for (const data of slideData) {
      const oldPath = data.info.zipPath
      const newPath = pathMap.get(oldPath)

      if (oldPath !== newPath) {
        // Remove XML file from ZIP
        this.#zipManager.removeFile(oldPath)

        // Remove .rels file from ZIP and clear RelationshipManager cache
        const oldRelsKey = relationshipManager.getRelsPath(oldPath)
        this.#zipManager.removeFile(oldRelsKey)
        relationshipManager.deleteRelationships(oldPath)

        // Remove content type override
        contentTypesManager.removeOverride(oldPath)
      }
    }

    // Phase 2: Write all new files and update cache entries
    for (let i = 0; i < slideData.length; i++) {
      const { info, xml, rels } = slideData[i]
      const oldPath = info.zipPath
      const newPath = pathMap.get(oldPath)

      // Update slide XML cache key
      this.#slideXmlCache.delete(oldPath)
      this.#slideXmlCache.set(newPath, xml)

      // Update SlideInfo path in memory
      info.zipPath = newPath

      // Write slide XML to ZIP
      this.#zipManager.writeFile(newPath, xml)

      // Add content type override
      contentTypesManager.addOverride(newPath, SLIDE_CONTENT_TYPE)

      // Set relationships for the new path in memory and flush to ZIP
      relationshipManager.setRelationships(newPath, rels)
      relationshipManager.flushRelationships(newPath)
    }

    // 2. Update slide-to-slide hyperlink targets and notes slide targets in all slide relationships
    for (const info of slides) {
      const rels = relationshipManager.getRelationships(info.zipPath)
      let dirty = false
      for (const rel of rels) {
        if (rel.target && !rel.targetMode) {
          const targetParts = rel.target.split('/')
          const targetName = targetParts[targetParts.length - 1]
          if (nameMap.has(targetName)) {
            targetParts[targetParts.length - 1] = nameMap.get(targetName)
            rel.target = targetParts.join('/')
            dirty = true
          }
        }
      }
      if (dirty) {
        relationshipManager.flushRelationships(info.zipPath)
      }
    }

    // 3. Normalize notes slides: renumber them to match the slide order
    await this.#normalizeNotesSlides(slides, relationshipManager, contentTypesManager, nameMap)

    // 4. Update slide targets in presentation.xml relationships if slide filenames changed
    const presRels = relationshipManager.getRelationships('ppt/presentation.xml')

    presRels.forEach(rel => {
      const target = rel.target
      if (target) {
        const targetParts = target.split('/')
        const targetName = targetParts[targetParts.length - 1]
        if (nameMap.has(targetName)) {
          targetParts[targetParts.length - 1] = nameMap.get(targetName)
          rel.target = targetParts.join('/')
        }
      }
    })

    // Flush relationships for ppt/presentation.xml to ZIP
    relationshipManager.flushRelationships('ppt/presentation.xml')

    // Clear any stale presentationIdMap (no longer used)
    this.#presentationIdMap = null

    // 7. Rebuild presentation.xml sldIdLst and synchronize sections
    this.rebuildPresentationSlideOrder()

    // 8. Remove orphan slide parts left over from add/remove/duplicate operations
    this.#purgeOrphanSlideParts(slides, relationshipManager, contentTypesManager)
  }

  /**
   * Deletes slide XML/.rels parts that are no longer referenced by the presentation.
   * @private
   */
  #purgeOrphanSlideParts(slides, relationshipManager, contentTypesManager) {
    const referenced = new Set(slides.map(s => s.zipPath))
    const slideFiles = this.#zipManager
      .listFiles('ppt/slides/')
      .filter(f => /\/slide\d+\.xml$/.test(f))

    for (const zipPath of slideFiles) {
      if (referenced.has(zipPath)) continue

      this.#zipManager.removeFile(zipPath)
      const relsPath = relationshipManager.getRelsPath(zipPath)
      this.#zipManager.removeFile(relsPath)
      relationshipManager.deleteRelationships(zipPath)
      contentTypesManager.removeOverride(zipPath)
      this.#slideXmlCache.delete(zipPath)
      logger.debug(`Purged orphan slide part: ${zipPath}`)
    }
  }

  /**
   * Normalizes notes slide filenames so they are sequentially numbered
   * matching their associated slides and updates all relevant relationships.
   * @private
   */
  async #normalizeNotesSlides(slides, relationshipManager, contentTypesManager, slideNameMap) {
    const NOTES_CONTENT_TYPE =
      'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'

    // Collect current notes slides from slide relationships
    const notesData = []
    for (let i = 0; i < slides.length; i++) {
      const info = slides[i]
      const rels = relationshipManager.getRelationships(info.zipPath)
      const notesRel = rels.find(r => r.type === REL_TYPES.NOTES_SLIDE)
      if (!notesRel) continue

      const oldNotesPath = relationshipManager.resolveTarget(info.zipPath, notesRel.target)
      const newNotesFileName = `notesSlide${i + 1}.xml`
      const newNotesPath = `ppt/notesSlides/${newNotesFileName}`
      notesData.push({
        slideInfo: info,
        notesRel,
        oldNotesPath,
        newNotesPath,
        newNotesFileName,
        slideNewFileName: info.zipPath.split('/').pop(),
      })
    }

    if (notesData.length === 0) return

    // Build a map of old notes filename → new notes filename (for cross-references)
    const notesNameMap = new Map()
    for (const nd of notesData) {
      const oldName = nd.oldNotesPath.split('/').pop()
      notesNameMap.set(oldName, nd.newNotesFileName)
    }

    // Phase A: Read all notes XML and relationships into memory, then remove old files if path changes
    for (const nd of notesData) {
      if (nd.oldNotesPath !== nd.newNotesPath) {
        // Load cached XML (already preloaded by preloadAll)
        let notesXml = this.#zipManager.readCachedFile(nd.oldNotesPath)
        if (!notesXml) {
          // Fallback: async read from ZIP
          notesXml = await this.#zipManager.readFile(nd.oldNotesPath)
        }
        nd._notesXml = notesXml

        const notesRels = relationshipManager.getRelationships(nd.oldNotesPath)
        nd._notesRels = [...notesRels]

        // Remove old notes XML, .rels, and content type
        this.#zipManager.removeFile(nd.oldNotesPath)
        const oldRelsKey = relationshipManager.getRelsPath(nd.oldNotesPath)
        this.#zipManager.removeFile(oldRelsKey)
        relationshipManager.deleteRelationships(nd.oldNotesPath)
        contentTypesManager.removeOverride(nd.oldNotesPath)
      }
    }

    // Phase B: Write new notes files and update relationships
    for (const nd of notesData) {
      if (nd.oldNotesPath === nd.newNotesPath) {
        // Path unchanged — still ensure the back-reference to the slide is correct
        const notesRels = relationshipManager.getRelationships(nd.newNotesPath)
        const backRef = notesRels.find(r => r.type === REL_TYPES.SLIDE || r.type.endsWith('/slide'))
        if (backRef) {
          const expectedTarget = `../slides/${nd.slideNewFileName}`
          if (backRef.target !== expectedTarget) {
            backRef.target = expectedTarget
            relationshipManager.flushRelationships(nd.newNotesPath)
          }
        }
        continue
      }

      const notesXml = nd._notesXml
      const notesRels = nd._notesRels || []

      // Update the notes rels: fix back-reference to the (possibly renamed) slide
      for (const nr of notesRels) {
        if (nr.type === REL_TYPES.SLIDE || nr.type.endsWith('/slide')) {
          nr.target = `../slides/${nd.slideNewFileName}`
        } else if (nr.target && !nr.targetMode) {
          const parts = nr.target.split('/')
          const fname = parts[parts.length - 1]
          if (slideNameMap.has(fname)) {
            parts[parts.length - 1] = slideNameMap.get(fname)
            nr.target = parts.join('/')
          }
        }
      }

      // Write new notes XML to ZIP
      if (notesXml) {
        this.#zipManager.writeFile(nd.newNotesPath, notesXml)
      }

      // Write new notes relationships
      relationshipManager.setRelationships(nd.newNotesPath, notesRels)
      relationshipManager.flushRelationships(nd.newNotesPath)

      // Register content type
      contentTypesManager.addOverride(nd.newNotesPath, NOTES_CONTENT_TYPE)

      // Update the slide's notes relationship target to point to the new notes path
      const slideRels = relationshipManager.getRelationships(nd.slideInfo.zipPath)
      const slideNotesRel = slideRels.find(r => r.type === REL_TYPES.NOTES_SLIDE)
      if (slideNotesRel) {
        slideNotesRel.target = `../notesSlides/${nd.newNotesFileName}`
        relationshipManager.flushRelationships(nd.slideInfo.zipPath)
      }
    }

    // Also update any notes-to-notes cross-references in unchanged notes
    for (const nd of notesData) {
      const notesRels = relationshipManager.getRelationships(nd.newNotesPath)
      let dirty = false
      for (const nr of notesRels) {
        if (nr.target && !nr.targetMode) {
          const parts = nr.target.split('/')
          const fname = parts[parts.length - 1]
          if (notesNameMap.has(fname)) {
            parts[parts.length - 1] = notesNameMap.get(fname)
            nr.target = parts.join('/')
            dirty = true
          }
        }
      }
      if (dirty) {
        relationshipManager.flushRelationships(nd.newNotesPath)
      }
    }
  }

  /**
   * Recursively walks an in-memory parsed XML object and remaps @_r:id attribute
   * values according to the provided ID map. This is more reliable than regex
   * replacement on the XML string because it only targets actual relationship ID
   * attributes, not arbitrary attribute values that might match the pattern.
   * @private
   * @param {Object} obj
   * @param {Map<string,string>} idMap
   */
  #remapPresentationRIds(obj, idMap) {
    if (!obj || typeof obj !== 'object') return
    for (const key of Object.keys(obj)) {
      if (key === '@_r:id') {
        const currentVal = obj[key]
        if (idMap.has(currentVal)) {
          obj[key] = idMap.get(currentVal)
        }
      } else {
        const child = obj[key]
        if (Array.isArray(child)) {
          for (const item of child) {
            this.#remapPresentationRIds(item, idMap)
          }
        } else if (child && typeof child === 'object') {
          this.#remapPresentationRIds(child, idMap)
        }
      }
    }
  }

  markSlideObjDirty(slideIndex) {
    this.#assertSlideExists(slideIndex)
    const state = this.#slideStates.get(slideIndex)
    state.dirty = true
    state.indexBuilt = false
  }

  getSlideShape(slideIndex, shapeId) {
    this.getSlideObj(slideIndex)
    return this.#slideStates.get(slideIndex).shapeMap.get(String(shapeId)) || null
  }

  getSlidePic(slideIndex, picId) {
    this.getSlideObj(slideIndex)
    const state = this.#slideStates.get(slideIndex)
    if (picId === 'first') {
      return Array.from(state.picMap.values())[0] || null
    }
    return state.picMap.get(String(picId)) || null
  }

  getSlideTable(slideIndex, tableId) {
    this.getSlideObj(slideIndex)
    const state = this.#slideStates.get(slideIndex)
    if (tableId === 'first') {
      return Array.from(state.tableMap.values())[0] || null
    }
    return state.tableMap.get(String(tableId)) || null
  }

  getSlideChart(slideIndex, chartId) {
    this.getSlideObj(slideIndex)
    const state = this.#slideStates.get(slideIndex)
    if (chartId === 'first') {
      return Array.from(state.chartMap.values())[0] || null
    }
    return state.chartMap.get(String(chartId)) || null
  }

  /**
   * Finds any drawable object (shape, picture, table, chart, group, graphicFrame) on a slide by name or ID.
   *
   * @param {number} slideIndex
   * @param {string} objectIdOrName
   * @returns {{ element: Object, type: string, frame: Object|null, parent: Object }|null}
   */
  findDrawableObject(slideIndex, objectIdOrName) {
    this.getSlideObj(slideIndex)
    const state = this.#slideStates.get(slideIndex)
    if (!state) return null

    const key = String(objectIdOrName)

    // 1. Check shapeMap (shapes, textboxes, group shapes)
    let res = state.shapeMap.get(key)
    if (!res && objectIdOrName === 'first') {
      res = Array.from(state.shapeMap.values())[0] || null
    }
    if (res) {
      return { element: res.shape, type: res.type, frame: null, parent: res.parent }
    }

    // 2. Check picMap (pictures, icons, SVGs)
    res = state.picMap.get(key)
    if (!res && objectIdOrName === 'first') {
      res = Array.from(state.picMap.values())[0] || null
    }
    if (res) {
      return { element: res.pic, type: 'pic', frame: null, parent: res.parent }
    }

    // 3. Check tableMap (tables)
    res = state.tableMap.get(key)
    if (!res && objectIdOrName === 'first') {
      res = Array.from(state.tableMap.values())[0] || null
    }
    if (res) {
      return { element: res.table, type: 'table', frame: res.frame, parent: res.parent }
    }

    // 4. Check chartMap (charts)
    res = state.chartMap.get(key)
    if (!res && objectIdOrName === 'first') {
      res = Array.from(state.chartMap.values())[0] || null
    }
    if (res) {
      return { element: res.chart, type: 'chart', frame: res.frame, parent: res.parent }
    }

    // 5. Inspect spTree for graphicFrames or placeholders
    const slideObj = this.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
    if (!spTree) return null

    let frames = spTree['p:graphicFrame'] || []
    if (!Array.isArray(frames)) frames = [frames]
    for (const frame of frames) {
      const cNvPr = frame?.['p:nvGraphicFramePr']?.['p:cNvPr']
      if (cNvPr) {
        const name = cNvPr['@_name']
        const id = String(cNvPr['@_id'])
        if (objectIdOrName === 'first' || name === objectIdOrName || id === key) {
          const tbl = frame?.['a:graphic']?.['a:graphicData']?.['a:tbl']
          if (tbl) return { element: tbl, type: 'table', frame, parent: spTree }
          const chart = frame?.['a:graphic']?.['a:graphicData']?.['c:chart']
          if (chart) return { element: chart, type: 'chart', parent: spTree }
          return { element: frame, type: 'graphicFrame', frame, parent: spTree }
        }
      }
    }

    return null
  }
}

module.exports = { SlideManager }
