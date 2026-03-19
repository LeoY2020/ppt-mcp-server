/**
 * PPTX Parser Service
 * Parses PowerPoint (PPTX) files to extract text, styles, and structure.
 * 
 * PPTX is a ZIP archive containing XML files:
 * - ppt/presentation.xml - Presentation metadata
 * - ppt/slides/slide1.xml, slide2.xml, ... - Individual slides
 * - ppt/slideLayouts/ - Layout templates
 * - ppt/slideMasters/ - Master slide templates
 */

import fs from 'fs';
import path from 'path';
import unzipper from 'unzipper';
import { parseString, Builder } from 'xml2js';

export interface TextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontName?: string;
  color?: string;
}

export interface Paragraph {
  text: string;
  runs: TextRun[];
  alignment?: 'left' | 'center' | 'right' | 'justify';
  bullet?: boolean;
  level?: number;
}

export interface Shape {
  id: string;
  name?: string;
  type: 'text' | 'image' | 'chart' | 'table' | 'unknown';
  position: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
  paragraphs?: Paragraph[];
  imageRef?: string;
}

export interface Slide {
  index: number;
  id: string;
  shapes: Shape[];
  layout?: string;
  notes?: string;
}

export interface PresentationInfo {
  title?: string;
  author?: string;
  subject?: string;
  creator?: string;
  created?: Date;
  modified?: Date;
  slideCount: number;
}

export interface PPTXContent {
  info: PresentationInfo;
  slides: Slide[];
}

// Parse XML to JS object
async function parseXml(xml: string): Promise<any> {
  return new Promise((resolve, reject) => {
    parseString(xml, {
      explicitArray: false,
      mergeAttrs: false,
      attrNameProcessors: [(name: string) => name.replace(/^a:/, '').replace(/^p:/, '').replace(/^r:/, '').replace(/^p14:/, '').replace(/^p15:/, '')]
    }, (err, result) => {
      if (err) reject(err);
      else resolve(result);
    });
  });
}

// Convert EMUs (English Metric Units) to points (1 inch = 914400 EMUs, 1 inch = 72 points)
function emuToPoints(emu: string): number {
  return Math.round(parseInt(emu) / 914400 * 72);
}

// Extract text runs from a text body
function extractTextRuns(textBody: any): TextRun[] {
  if (!textBody) return [];
  
  const runs: TextRun[] = [];
  const paragraphs = Array.isArray(textBody['a:p']) ? textBody['a:p'] : [textBody['a:p']].filter(Boolean);
  
  for (const para of paragraphs) {
    if (!para) continue;
    
    const textRuns = Array.isArray(para['a:r']) ? para['a:r'] : [para['a:r']].filter(Boolean);
    
    for (const run of textRuns) {
      if (!run) continue;
      
      const text = run['a:t'] || '';
      if (!text && text !== '') continue;
      
      const runProps = run['a:rPr']?.$ || {};
      
      runs.push({
        text: text,
        bold: runProps.b === '1',
        italic: runProps.i === '1',
        underline: runProps.u === 'sng' || runProps.u === '1',
        fontSize: runProps.sz ? parseInt(runProps.sz) / 100 : undefined,
        fontName: runProps['latin']?.typeface,
        color: runProps.solidFill?.['srgbClr']?.$?.val
      });
    }
  }
  
  return runs;
}

// Extract paragraphs from a text body
function extractParagraphs(textBody: any): Paragraph[] {
  if (!textBody) return [];
  
  const paragraphs: Paragraph[] = [];
  const paras = Array.isArray(textBody['a:p']) ? textBody['a:p'] : [textBody['a:p']].filter(Boolean);
  
  for (const para of paras) {
    if (!para) continue;
    
    const runs: TextRun[] = [];
    const textRuns = Array.isArray(para['a:r']) ? para['a:r'] : [para['a:r']].filter(Boolean);
    
    let fullText = '';
    
    for (const run of textRuns) {
      if (!run) continue;
      
      const text = run['a:t'] || '';
      fullText += text;
      
      const runProps = run['a:rPr']?.$ || {};
      
      runs.push({
        text: text,
        bold: runProps.b === '1',
        italic: runProps.i === '1',
        underline: runProps.u === 'sng' || runProps.u === '1',
        fontSize: runProps.sz ? parseInt(runProps.sz) / 100 : undefined,
        fontName: runProps['latin']?.typeface,
        color: runProps.solidFill?.['srgbClr']?.$?.val
      });
    }
    
    const paraProps = para['a:pPr']?.$ || {};
    const alignment = paraProps.algn as 'l' | 'ctr' | 'r' | 'just' | undefined;
    
    paragraphs.push({
      text: fullText,
      runs: runs,
      alignment: alignment === 'l' ? 'left' : alignment === 'ctr' ? 'center' : alignment === 'r' ? 'right' : alignment === 'just' ? 'justify' : undefined,
      bullet: !!para['a:pPr']?.['a:buFont'] || !!para['a:pPr']?.['a:buChar'],
      level: paraProps.lvl ? parseInt(paraProps.lvl) : undefined
    });
  }
  
  return paragraphs;
}

// Extract shape info
function extractShape(shape: any, index: number): Shape {
  const spPr = shape['p:spPr'] || shape['p:picPr'] || {};
  const nvSpPr = shape['p:nvSpPr'] || shape['p:nvPicPr'] || {};
  
  // Get position and size
  const xfrm = spPr['a:xfrm'];
  const position = {
    x: xfrm?.['a:off']?.$?.x ? emuToPoints(xfrm['a:off'].$.x) : 0,
    y: xfrm?.['a:off']?.$?.y ? emuToPoints(xfrm['a:off'].$.y) : 0,
    width: xfrm?.['a:ext']?.$?.cx ? emuToPoints(xfrm['a:ext'].$.cx) : 0,
    height: xfrm?.['a:ext']?.$?.cy ? emuToPoints(xfrm['a:ext'].$.cy) : 0
  };
  
  // Get shape ID and name
  const cNvPr = nvSpPr['p:cNvPr']?.$ || {};
  const shapeId = cNvPr.id || String(index);
  const shapeName = cNvPr.name;
  
  // Determine shape type
  let type: Shape['type'] = 'unknown';
  let paragraphs: Paragraph[] = [];
  let imageRef: string | undefined;
  
  if (shape['p:txBody']) {
    type = 'text';
    paragraphs = extractParagraphs(shape['p:txBody']);
  } else if (shape['p:picPr'] || shape['p:nvPicPr']) {
    type = 'image';
    // Get image reference
    const blip = shape['p:blipFill']?.['a:blip'];
    if (blip?.$?.embed) {
      imageRef = blip.$.embed;
    }
  } else if (spPr['a:txBody']) {
    type = 'text';
    paragraphs = extractParagraphs(spPr['a:txBody']);
  } else if (shape['p:graphicFrame']) {
    type = 'chart';
  }
  
  return {
    id: shapeId,
    name: shapeName,
    type,
    position,
    paragraphs,
    imageRef
  };
}

// Parse a single slide
function parseSlide(slideXml: string, slideIndex: number): Promise<Slide> {
  return new Promise(async (resolve, reject) => {
    try {
      const parsed = await parseXml(slideXml);
      
      const slideData = parsed['p:sld'];
      if (!slideData) {
        reject(new Error('Invalid slide XML'));
        return;
      }
      
      const cSld = slideData['p:cSld'] || {};
      const spTree = cSld['p:spTree'] || {};
      
      // Get all shapes
      const shapes: Shape[] = [];
      
      // Process shapes (can be array or single object)
      const sps = spTree['p:sp'];
      if (sps) {
        const shapeArray = Array.isArray(sps) ? sps : [sps];
        shapeArray.forEach((sp: any, i: number) => {
          try {
            shapes.push(extractShape(sp, i));
          } catch (e) {
            // Skip invalid shapes
          }
        });
      }
      
      // Process pictures
      const pics = spTree['p:pic'];
      if (pics) {
        const picArray = Array.isArray(pics) ? pics : [pics];
        picArray.forEach((pic: any, i: number) => {
          try {
            shapes.push(extractShape(pic, shapes.length + i));
          } catch (e) {
            // Skip invalid pictures
          }
        });
      }
      
      // Get slide notes if available
      let notes: string | undefined;
      if (slideData['p:notes']?.['p:cSld']?.['p:spTree']?.['p:sp']) {
        const notesSp = slideData['p:notes']['p:cSld']['p:spTree']['p:sp'];
        const notesTexts: string[] = [];
        const notesShapes = Array.isArray(notesSp) ? notesSp : [notesSp];
        for (const ns of notesShapes) {
          if (ns['p:txBody']) {
            const paras = extractParagraphs(ns['p:txBody']);
            notesTexts.push(...paras.map(p => p.text));
          }
        }
        notes = notesTexts.join('\n');
      }
      
      resolve({
        index: slideIndex,
        id: `slide${slideIndex}`,
        shapes,
        layout: slideData.$?.layout,
        notes
      });
    } catch (error) {
      reject(error);
    }
  });
}

// Parse presentation.xml to get core info
async function parsePresentationInfo(xml: string): Promise<PresentationInfo> {
  const parsed = await parseXml(xml);
  const presData = parsed['p:presentation'];
  
  let slideCount = 0;
  if (presData?.['p:sldIdLst']?.['p:sldId']) {
    const sldIds = presData['p:sldIdLst']['p:sldId'];
    slideCount = Array.isArray(sldIds) ? sldIds.length : 1;
  }
  
  return {
    slideCount,
    title: undefined,
    author: undefined,
    subject: undefined,
    creator: undefined
  };
}

// Parse core.xml for document properties
async function parseCoreProperties(xml: string): Promise<Partial<PresentationInfo>> {
  const parsed = await parseXml(xml);
  const props = parsed['cp:coreProperties'] || parsed;
  
  return {
    title: props['dc:title'] || undefined,
    author: props['dc:creator'] || props['cp:lastModifiedBy'] || undefined,
    subject: props['dc:subject'] || undefined,
    creator: props['dc:creator'] || undefined,
    created: props['dcterms:created']?._ || props['dcterms:created'] ? new Date(props['dcterms:created']?._ || props['dcterms:created']) : undefined,
    modified: props['dcterms:modified']?._ || props['dcterms:modified'] ? new Date(props['dcterms:modified']?._ || props['dcterms:modified']) : undefined
  };
}

/**
 * Parse a PPTX file and return its content
 */
export async function parsePPTX(filePath: string): Promise<PPTXContent> {
  return new Promise((resolve, reject) => {
    if (!fs.existsSync(filePath)) {
      reject(new Error(`File not found: ${filePath}`));
      return;
    }
    
    const slides: Slide[] = [];
    let presentationInfo: PresentationInfo = { slideCount: 0 };
    let coreProps: Partial<PresentationInfo> = {};
    
    fs.createReadStream(filePath)
      .pipe(unzipper.Parse())
      .on('entry', async (entry: any) => {
        const fileName = entry.path;
        
        try {
          // Parse presentation.xml
          if (fileName === 'ppt/presentation.xml') {
            const xml = await entry.buffer();
            presentationInfo = await parsePresentationInfo(xml.toString());
          }
          // Parse core properties
          else if (fileName === 'docProps/core.xml') {
            const xml = await entry.buffer();
            coreProps = await parseCoreProperties(xml.toString());
          }
          // Parse slides
          else if (fileName.match(/^ppt\/slides\/slide(\d+)\.xml$/)) {
            const match = fileName.match(/slide(\d+)\.xml$/);
            if (match) {
              const slideIndex = parseInt(match[1]);
              const xml = await entry.buffer();
              const slide = await parseSlide(xml.toString(), slideIndex);
              slides.push(slide);
            }
          } else {
            entry.autodrain();
          }
        } catch (error) {
          entry.autodrain();
        }
      })
      .on('close', () => {
        // Sort slides by index
        slides.sort((a, b) => a.index - b.index);
        
        resolve({
          info: {
            ...presentationInfo,
            ...coreProps,
            slideCount: slides.length
          },
          slides
        });
      })
      .on('error', (error: Error) => {
        reject(new Error(`Failed to parse PPTX: ${error.message}`));
      });
  });
}

/**
 * Get a summary of the presentation
 */
export function getPresentationSummary(content: PPTXContent): string {
  const lines: string[] = [];
  
  lines.push(`# Presentation Summary`);
  lines.push('');
  
  if (content.info.title) lines.push(`**Title:** ${content.info.title}`);
  if (content.info.author) lines.push(`**Author:** ${content.info.author}`);
  if (content.info.subject) lines.push(`**Subject:** ${content.info.subject}`);
  lines.push(`**Total Slides:** ${content.slides.length}`);
  
  if (content.slides.length > 0) {
    lines.push('');
    lines.push('## Slide Overview');
    for (const slide of content.slides) {
      const textContent = slide.shapes
        .filter(s => s.type === 'text' && s.paragraphs)
        .flatMap(s => s.paragraphs!)
        .map(p => p.text)
        .join(' ')
        .substring(0, 100);
      
      lines.push(`- **Slide ${slide.index}:** ${textContent}${textContent.length >= 100 ? '...' : ''}`);
    }
  }
  
  return lines.join('\n');
}
