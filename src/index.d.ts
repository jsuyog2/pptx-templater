export class PPTXTemplater {
  static load(source: string | Buffer, options?: any): Promise<PPTXTemplater>;
  static setLogLevel(level: string): void;
  static preload(source: string | Buffer): Promise<any>;
  static cache(source: string | Buffer): Promise<any>;
  static fromCache(source: string | Buffer): Promise<PPTXTemplater>;
  static clearCache(): void;
  static fromPresentationXml(options: any): Promise<PPTXTemplater>;
  static create(): Promise<PPTXTemplater>;
  static extractPptx(pptxPath: string, outputPath: string, options?: any): Promise<void>;
  static buildPptx(folderPath: string, pptxPath: string): Promise<void>;

  slideCount: number;
  useSlide(slideNumber: number): this;
  replaceText(replacements: Record<string, string | number>): this;
  updateChart(chartName: string, data: any): this;
  updateChartTitle(chartName: string, title: string): this;
  applyZOrder(slideNumber: number, zOrderConfig: any[]): this;
  removeSlide(slideNumber: number): this;
  addHyperlink(options: any): this;
  addSlideLink(options: any): this;
  getTableRows(tableName: string): Promise<any[]>;
  updateTable(tableName: string, rows: any[][]): this;
  addTableRow(tableName: string, row: any[]): this;
  removeTableRow(tableName: string, rowIndex: number): this;
  adjustTableRow(tableName: string, options: any): this;
  adjustTableColumn(tableName: string, options: any): this;
  adjustTable(tableName: string, options?: any): this;
  mergeCells(tableName: string, startRow: number, startCol: number, endRow: number, endCol: number): this;
  saveToFile(filePath: string, options?: any): Promise<void>;
  save(filePath: string, options?: any): Promise<void>;
  saveXml(folderPath: string): Promise<void>;
  saveToFolder(folderPath: string): Promise<void>;
  toBuffer(options?: any): Promise<Buffer>;
  toStream(options?: any): Promise<any>;
  saveToStream(writableOrOptions: any, options?: any): Promise<void>;
  exportSlides(...slideNumbers: number[]): Promise<any>;
  importSlideFrom(sourceEngine: PPTXTemplater, slideRef: any): Promise<any>;
  repair(): Promise<this>;
  inspectXML(xmlPath: string): Promise<any>;
  validateCharts(): Promise<any>;
  repairCharts(): Promise<any>;
  inspectChartXML(chartFileName: string): Promise<any>;
  getDataLabels(chartId: string, options?: any): Promise<any>;
  validateDataLabels(chartId: string, options?: any): Promise<any>;
  validateChartLabels(chartId: string, options?: any): Promise<any>;
  validateSeriesNameLabels(chartId: string, options?: any): Promise<any>;
  getChartLabelPositions(chartId: string): Promise<any>;
  getChartBarPositions(chartId: string): Promise<any>;
  addShape(typeOrOptions: any, options?: any): Promise<any>;
  updateShape(shapeId: string, options: any): Promise<any>;
  removeShape(shapeId: string): Promise<any>;
  addCellShape(tableId: string, rowIndex: number, colIndex: number, options: any): Promise<any>;
  updateCellShape(tableId: string, rowIndex: number, colIndex: number, shapeIndex: number, options: any): Promise<any>;
  removeCellShape(tableId: string, rowIndex: number, colIndex: number, shapeIndex: number): Promise<any>;
  replaceImage(imageIdOrName: string, sourcePathOrBuffer: string | Buffer): Promise<any>;
  addImage(sourcePathOrBuffer: string | Buffer, options?: any): Promise<any>;
  validatePresentation(): Promise<any>;
  validatePresentationXml(): Promise<any>;
  validateSlide(slideIndex: number): Promise<any>;
  validateTable(tableId: string): Promise<any>;
  validateArchive(): Promise<any>;
}

export const PPTXTemplate: typeof PPTXTemplater;

export class ZipManager {
  [key: string]: any;
}

export class XMLParser {
  [key: string]: any;
}

export const Z_ORDER_SYMBOL: unique symbol;

export class SlideManager {
  [key: string]: any;
}

export class ChartManager {
  [key: string]: any;
}

export class TableManager {
  [key: string]: any;
}

export class ShapeManager {
  [key: string]: any;
}

export class ImageManager {
  [key: string]: any;
}

export class TextManager {
  [key: string]: any;
}

export class HyperlinkManager {
  [key: string]: any;
}

export class MediaManager {
  [key: string]: any;
}

export class RelationshipManager {
  [key: string]: any;
}

export class OutputWriter {
  [key: string]: any;
}

export class TemplateEngine {
  [key: string]: any;
}

export class ValidationEngine {
  [key: string]: any;
}

export function generateRelationshipId(slideId: any): string;
export function parseRelationshipId(relId: string): any;
export function validateXml(xmlStr: string): any;
export function validateXML(xmlStr: string): any;
export function safeParseXml(xmlStr: string): any;
export function repairXML(xmlStr: string): string;
export function scanForEntities(xmlStr: string): any;
export function analyzeXmlFile(xmlStr: string): any;
export function reportXmlComplexity(xmlStr: string): any;
export function createLogger(name: string): any;
export function setGlobalLogLevel(level: string): void;
export function resetLogLevel(): void;

export class PPTXError extends Error {}
export class SlideNotFoundError extends PPTXError {}
export class ChartNotFoundError extends PPTXError {}
export class TableNotFoundError extends PPTXError {}
