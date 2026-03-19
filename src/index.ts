#!/usr/bin/env node
/**
 * PPT MCP Server
 * 
 * MCP server for reading PowerPoint (PPTX) files including text content,
 * styles, and document structure.
 * 
 * Tools provided:
 * - ppt_read_presentation: Read all slides from a PPTX file
 * - ppt_get_slide: Get detailed content of a specific slide
 * - ppt_get_info: Get presentation metadata
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  parsePPTX,
  getPresentationSummary,
  Slide
} from "./services/pptParser.js";

// Constants
const CHARACTER_LIMIT = 50000;

// Response format enum
enum ResponseFormat {
  MARKDOWN = "markdown",
  JSON = "json"
}

// Zod Schemas (as raw shapes for SDK)
const ReadPresentationSchema = {
  file_path: z.string()
    .min(1, "File path is required")
    .describe("Absolute path to the PPTX file"),
  include_styles: z.boolean()
    .default(true)
    .describe("Whether to include text styling information (bold, italic, font size, etc.)"),
  include_positions: z.boolean()
    .default(false)
    .describe("Whether to include shape positions and dimensions"),
  response_format: z.nativeEnum(ResponseFormat)
    .default(ResponseFormat.MARKDOWN)
    .describe("Output format: 'markdown' for human-readable or 'json' for machine-readable")
};

const GetSlideSchema = {
  file_path: z.string()
    .min(1, "File path is required")
    .describe("Absolute path to the PPTX file"),
  slide_index: z.number()
    .int()
    .min(1)
    .describe("Slide number (1-based index)"),
  include_styles: z.boolean()
    .default(true)
    .describe("Whether to include text styling information"),
  include_positions: z.boolean()
    .default(true)
    .describe("Whether to include shape positions and dimensions"),
  response_format: z.nativeEnum(ResponseFormat)
    .default(ResponseFormat.MARKDOWN)
    .describe("Output format: 'markdown' for human-readable or 'json' for machine-readable")
};

const GetInfoSchema = {
  file_path: z.string()
    .min(1, "File path is required")
    .describe("Absolute path to the PPTX file"),
  response_format: z.nativeEnum(ResponseFormat)
    .default(ResponseFormat.MARKDOWN)
    .describe("Output format: 'markdown' for human-readable or 'json' for machine-readable")
};

// Format slide content as markdown
function formatSlideMarkdown(
  slide: Slide,
  includeStyles: boolean,
  includePositions: boolean
): string {
  const lines: string[] = [];
  
  lines.push(`## Slide ${slide.index}`);
  lines.push('');
  
  if (slide.notes) {
    lines.push(`> **Notes:** ${slide.notes}`);
    lines.push('');
  }
  
  for (const shape of slide.shapes) {
    if (shape.type === 'text' && shape.paragraphs && shape.paragraphs.length > 0) {
      if (shape.name) {
        lines.push(`### ${shape.name}`);
      }
      
      if (includePositions) {
        lines.push(`*[Position: (${shape.position.x}, ${shape.position.y}), Size: ${shape.position.width}x${shape.position.height}]*`);
        lines.push('');
      }
      
      for (const para of shape.paragraphs) {
        if (!para.text.trim()) {
          lines.push('');
          continue;
        }
        
        let text = para.text;
        
        if (includeStyles && para.runs.length > 0) {
          let formattedText = '';
          for (const run of para.runs) {
            let runText = run.text;
            if (run.bold) runText = `**${runText}**`;
            if (run.italic) runText = `*${runText}*`;
            formattedText += runText;
          }
          text = formattedText || para.text;
        }
        
        if (para.bullet) {
          const indent = '  '.repeat(para.level || 0);
          lines.push(`${indent}- ${text}`);
        } else {
          lines.push(text);
        }
      }
      lines.push('');
    } else if (shape.type === 'image') {
      const posInfo = includePositions ? ` at (${shape.position.x}, ${shape.position.y})` : '';
      lines.push(`*[Image${posInfo}]*`);
      lines.push('');
    }
  }
  
  return lines.join('\n');
}

// Format slide as JSON
function formatSlideJson(
  slide: Slide,
  includeStyles: boolean,
  includePositions: boolean
): object {
  return {
    index: slide.index,
    id: slide.id,
    layout: slide.layout,
    notes: slide.notes,
    shapes: slide.shapes.map(shape => {
      const result: Record<string, unknown> = {
        id: shape.id,
        name: shape.name,
        type: shape.type
      };
      
      if (includePositions) {
        result.position = shape.position;
      }
      
      if (shape.type === 'text' && shape.paragraphs) {
        result.paragraphs = shape.paragraphs.map(para => {
          const p: Record<string, unknown> = {
            text: para.text,
            alignment: para.alignment,
            bullet: para.bullet,
            level: para.level
          };
          
          if (includeStyles) {
            p.runs = para.runs.map(run => ({
              text: run.text,
              bold: run.bold,
              italic: run.italic,
              underline: run.underline,
              fontSize: run.fontSize,
              fontName: run.fontName,
              color: run.color
            }));
          }
          
          return p;
        });
      }
      
      if (shape.type === 'image') {
        result.imageRef = shape.imageRef;
      }
      
      return result;
    })
  };
}

// Truncate response if needed
function truncateResponse(text: string, originalLength: number): string {
  if (text.length <= CHARACTER_LIMIT) {
    return text;
  }
  
  return text.substring(0, CHARACTER_LIMIT) +
    `\n\n... [Response truncated from ${originalLength} to ${CHARACTER_LIMIT} characters. Use ppt_get_slide to view individual slides.]`;
}

// Error handling
function handleError(error: unknown): string {
  if (error instanceof Error) {
    if (error.message.includes('File not found')) {
      return `Error: ${error.message}. Please check the file path is correct and the file exists.`;
    }
    if (error.message.includes('Failed to parse')) {
      return `Error: Failed to parse PPTX file. The file may be corrupted or not a valid PPTX format.`;
    }
    return `Error: ${error.message}`;
  }
  return `Error: Unexpected error occurred`;
}

// Create MCP server
const server = new McpServer({
  name: "ppt-mcp-server",
  version: "1.0.0"
});

// Tool: Read entire presentation
server.tool(
  "ppt_read_presentation",
  `Read all slides from a PPTX file, extracting text content, styles, and structure.

This tool reads the entire presentation and returns content from all slides.
For large presentations, consider using ppt_get_slide to view individual slides.

Args:
  - file_path (string): Absolute path to the PPTX file
  - include_styles (boolean): Include text styling (bold, italic, font size). Default: true
  - include_positions (boolean): Include shape positions and dimensions. Default: false
  - response_format ('markdown' | 'json'): Output format. Default: 'markdown'

Returns:
  Presentation content with all slides. For JSON format:
  {
    "info": { "title", "author", "slideCount", ... },
    "slides": [
      {
        "index": number,
        "shapes": [
          { "type", "paragraphs": [...], "position": {...} }
        ]
      }
    ]
  }

Examples:
  - Read presentation with styles: { "file_path": "/path/to/presentation.pptx" }
  - JSON output: { "file_path": "/path/to/presentation.pptx", "response_format": "json" }`,
  ReadPresentationSchema,
  async (params) => {
    try {
      const content = await parsePPTX(params.file_path);
      
      if (params.response_format === ResponseFormat.JSON) {
        const output = {
          info: content.info,
          slides: content.slides.map(slide => 
            formatSlideJson(slide, params.include_styles ?? true, params.include_positions ?? false)
          )
        };
        const text = JSON.stringify(output, null, 2);
        return {
          content: [{ type: "text" as const, text: truncateResponse(text, text.length) }]
        };
      }
      
      // Markdown format
      const lines: string[] = [];
      
      lines.push('# Presentation Content');
      lines.push('');
      if (content.info.title) {
        lines.push(`**Title:** ${content.info.title}`);
      }
      if (content.info.author) {
        lines.push(`**Author:** ${content.info.author}`);
      }
      lines.push(`**Total Slides:** ${content.slides.length}`);
      lines.push('');
      lines.push('---');
      lines.push('');
      
      for (const slide of content.slides) {
        lines.push(formatSlideMarkdown(slide, params.include_styles ?? true, params.include_positions ?? false));
        lines.push('---');
        lines.push('');
      }
      
      const text = lines.join('\n');
      return {
        content: [{ type: "text" as const, text: truncateResponse(text, text.length) }]
      };
    } catch (error) {
      return {
        content: [{ type: "text" as const, text: handleError(error) }],
        isError: true
      };
    }
  }
);

// Tool: Get specific slide
server.tool(
  "ppt_get_slide",
  `Get detailed content of a specific slide from a PPTX file.

Returns all text content, styles, and shape information for a single slide.

Args:
  - file_path (string): Absolute path to the PPTX file
  - slide_index (number): Slide number (1-based, first slide is 1)
  - include_styles (boolean): Include text styling information. Default: true
  - include_positions (boolean): Include shape positions and dimensions. Default: true
  - response_format ('markdown' | 'json'): Output format. Default: 'markdown'

Returns:
  Detailed slide content with all shapes and text.

Examples:
  - Get first slide: { "file_path": "/path/to/presentation.pptx", "slide_index": 1 }
  - Get slide 3 as JSON: { "file_path": "/path/to/presentation.pptx", "slide_index": 3, "response_format": "json" }`,
  GetSlideSchema,
  async (params) => {
    try {
      const content = await parsePPTX(params.file_path);
      
      const slide = content.slides.find(s => s.index === params.slide_index);
      if (!slide) {
        return {
          content: [{
            type: "text" as const,
            text: `Error: Slide ${params.slide_index} not found. Presentation has ${content.slides.length} slides.`
          }],
          isError: true
        };
      }
      
      if (params.response_format === ResponseFormat.JSON) {
        const output = formatSlideJson(slide, params.include_styles ?? true, params.include_positions ?? true);
        return {
          content: [{ type: "text" as const, text: JSON.stringify(output, null, 2) }]
        };
      }
      
      const text = formatSlideMarkdown(slide, params.include_styles ?? true, params.include_positions ?? true);
      return {
        content: [{ type: "text" as const, text }]
      };
    } catch (error) {
      return {
        content: [{ type: "text" as const, text: handleError(error) }],
        isError: true
      };
    }
  }
);

// Tool: Get presentation info
server.tool(
  "ppt_get_info",
  `Get metadata and summary of a PPTX file without reading all content.

Returns title, author, slide count, and a brief summary of each slide.

Args:
  - file_path (string): Absolute path to the PPTX file
  - response_format ('markdown' | 'json'): Output format. Default: 'markdown'

Returns:
  Presentation metadata and slide summary.

Examples:
  - Get info: { "file_path": "/path/to/presentation.pptx" }`,
  GetInfoSchema,
  async (params) => {
    try {
      const content = await parsePPTX(params.file_path);
      
      if (params.response_format === ResponseFormat.JSON) {
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({
              info: content.info,
              slides: content.slides.map(s => ({
                index: s.index,
                shapeCount: s.shapes.length,
                textPreview: s.shapes
                  .filter(shape => shape.type === 'text' && shape.paragraphs)
                  .flatMap(shape => shape.paragraphs!)
                  .map(p => p.text)
                  .join(' ')
                  .substring(0, 100)
              }))
            }, null, 2)
          }]
        };
      }
      
      const summary = getPresentationSummary(content);
      return {
        content: [{ type: "text" as const, text: summary }]
      };
    } catch (error) {
      return {
        content: [{ type: "text" as const, text: handleError(error) }],
        isError: true
      };
    }
  }
);

// Start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("PPT MCP Server running via stdio");
}

main().catch(error => {
  console.error("Server error:", error);
  process.exit(1);
});
