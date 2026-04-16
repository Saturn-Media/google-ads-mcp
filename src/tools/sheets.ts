/**
 * Google Sheets Tools
 *
 * MCP tools for creating, reading, and writing Google Sheets.
 * Uses the Sheets API v4 via googleapis.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { getSheetsClient, getDriveClient } from '../lib/sheets-client.js';

// --- Tool Schemas ---

const CreateSpreadsheetSchema = z.object({
  title: z.string().describe('Title of the new spreadsheet'),
  folder_id: z.string().optional().describe('Google Drive folder ID to create in'),
});

const CreateSheetSchema = z.object({
  spreadsheet_id: z.string().describe('The spreadsheet ID'),
  title: z.string().describe('Title for the new sheet tab'),
});

const ListSheetsSchema = z.object({
  spreadsheet_id: z.string().describe('The spreadsheet ID'),
});

const GetSheetDataSchema = z.object({
  spreadsheet_id: z.string().describe('The spreadsheet ID'),
  sheet: z.string().describe('Sheet tab name'),
  range: z.string().optional().describe('A1 notation range (e.g. "A1:C10"). Omit for all data.'),
});

const UpdateCellsSchema = z.object({
  spreadsheet_id: z.string().describe('The spreadsheet ID'),
  sheet: z.string().describe('Sheet tab name'),
  range: z.string().describe('A1 notation range (e.g. "A1:N100")'),
  data: z.array(z.array(z.any())).describe('2D array of values'),
});

const BatchUpdateCellsSchema = z.object({
  spreadsheet_id: z.string().describe('The spreadsheet ID'),
  sheet: z.string().describe('Sheet tab name'),
  ranges: z.record(z.string(), z.array(z.array(z.any()))).describe('Map of range strings to 2D arrays'),
});

const BatchUpdateSchema = z.object({
  spreadsheet_id: z.string().describe('The spreadsheet ID'),
  requests: z.array(z.any()).describe('Array of Sheets API batchUpdate request objects'),
});

// --- Tool Implementations ---

export async function createSpreadsheet(args: unknown): Promise<string> {
  const { title, folder_id } = CreateSpreadsheetSchema.parse(args);
  const sheets = getSheetsClient();
  const res = await sheets.spreadsheets.create({
    requestBody: { properties: { title } },
    fields: 'spreadsheetId',
  });
  const id = res.data.spreadsheetId!;

  // Move to folder if specified
  if (folder_id) {
    const drive = getDriveClient();
    await drive.files.update({
      fileId: id,
      addParents: folder_id,
      fields: 'id',
    });
  }

  return JSON.stringify({
    spreadsheet_id: id,
    url: `https://docs.google.com/spreadsheets/d/${id}`,
    title,
  });
}

export async function createSheet(args: unknown): Promise<string> {
  const { spreadsheet_id, title } = CreateSheetSchema.parse(args);
  const sheets = getSheetsClient();
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: spreadsheet_id,
    requestBody: {
      requests: [{ addSheet: { properties: { title } } }],
    },
  });
  const sheetId = res.data.replies?.[0]?.addSheet?.properties?.sheetId;
  return JSON.stringify({ sheet_id: sheetId, title });
}

export async function listSheets(args: unknown): Promise<string> {
  const { spreadsheet_id } = ListSheetsSchema.parse(args);
  const sheets = getSheetsClient();
  const res = await sheets.spreadsheets.get({ spreadsheetId: spreadsheet_id });
  const sheetList = res.data.sheets?.map(s => ({
    title: s.properties?.title,
    sheetId: s.properties?.sheetId,
  }));
  return JSON.stringify(sheetList);
}

export async function getSheetData(args: unknown): Promise<string> {
  const { spreadsheet_id, sheet, range } = GetSheetDataSchema.parse(args);
  const sheets = getSheetsClient();
  const fullRange = range ? `${sheet}!${range}` : sheet;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: spreadsheet_id,
    range: fullRange,
  });
  return JSON.stringify({ values: res.data.values || [], range: fullRange });
}

export async function updateCells(args: unknown): Promise<string> {
  const { spreadsheet_id, sheet, range, data } = UpdateCellsSchema.parse(args);
  const sheets = getSheetsClient();
  const fullRange = `${sheet}!${range}`;
  const res = await sheets.spreadsheets.values.update({
    spreadsheetId: spreadsheet_id,
    range: fullRange,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: data },
  });
  return JSON.stringify({
    updated_range: res.data.updatedRange,
    updated_rows: res.data.updatedRows,
    updated_cells: res.data.updatedCells,
  });
}

export async function batchUpdateCells(args: unknown): Promise<string> {
  const { spreadsheet_id, sheet, ranges } = BatchUpdateCellsSchema.parse(args);
  const sheets = getSheetsClient();
  const data = Object.entries(ranges).map(([range, values]) => ({
    range: `${sheet}!${range}`,
    values: values as any[][],
  }));
  const res = await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: spreadsheet_id,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data,
    },
  } as any);
  return JSON.stringify({
    total_updated_cells: (res as any).data?.totalUpdatedCells,
    total_updated_rows: (res as any).data?.totalUpdatedRows,
  });
}

export async function batchUpdate(args: unknown): Promise<string> {
  const { spreadsheet_id, requests } = BatchUpdateSchema.parse(args);
  const sheets = getSheetsClient();
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: spreadsheet_id,
    requestBody: { requests },
  });
  return JSON.stringify({
    replies_count: res.data.replies?.length || 0,
  });
}

// --- Tool Definitions ---

export const createSpreadsheetTool: Tool = {
  name: 'sheets-create-spreadsheet',
  description: 'Create a new Google Spreadsheet. Returns the spreadsheet ID and URL.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      title: { type: 'string' as const, description: 'Title of the new spreadsheet' },
    },
    required: ['title'],
  },
};

export const createSheetTool: Tool = {
  name: 'sheets-create-sheet',
  description: 'Create a new sheet tab in an existing Google Spreadsheet.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      spreadsheet_id: { type: 'string' as const, description: 'The spreadsheet ID' },
      title: { type: 'string' as const, description: 'Title for the new tab' },
    },
    required: ['spreadsheet_id', 'title'],
  },
};

export const listSheetsTool: Tool = {
  name: 'sheets-list-sheets',
  description: 'List all sheet tabs in a Google Spreadsheet.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      spreadsheet_id: { type: 'string' as const, description: 'The spreadsheet ID' },
    },
    required: ['spreadsheet_id'],
  },
};

export const getSheetDataTool: Tool = {
  name: 'sheets-get-data',
  description: 'Read data from a Google Spreadsheet range.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      spreadsheet_id: { type: 'string' as const, description: 'The spreadsheet ID' },
      sheet: { type: 'string' as const, description: 'Sheet tab name' },
      range: { type: 'string' as const, description: 'A1 notation range (optional)' },
    },
    required: ['spreadsheet_id', 'sheet'],
  },
};

export const updateCellsTool: Tool = {
  name: 'sheets-update-cells',
  description: 'Write a 2D array of values to a Google Spreadsheet range. Use for bulk data writes.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      spreadsheet_id: { type: 'string' as const, description: 'The spreadsheet ID' },
      sheet: { type: 'string' as const, description: 'Sheet tab name' },
      range: { type: 'string' as const, description: 'A1 notation range (e.g. "A1:N100")' },
      data: {
        type: 'array' as const,
        items: { type: 'array' as const, items: {} },
        description: '2D array of values',
      },
    },
    required: ['spreadsheet_id', 'sheet', 'range', 'data'],
  },
};

export const batchUpdateCellsTool: Tool = {
  name: 'sheets-batch-update-cells',
  description: 'Write data to multiple ranges in one call.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      spreadsheet_id: { type: 'string' as const, description: 'The spreadsheet ID' },
      sheet: { type: 'string' as const, description: 'Sheet tab name' },
      ranges: {
        type: 'object' as const,
        description: 'Map of range strings to 2D arrays of values',
      },
    },
    required: ['spreadsheet_id', 'sheet', 'ranges'],
  },
};

export const batchUpdateTool: Tool = {
  name: 'sheets-batch-update',
  description: 'Execute a batch update on a Google Spreadsheet (formatting, freeze rows, conditional formatting, etc.). Takes an array of Sheets API batchUpdate request objects.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      spreadsheet_id: { type: 'string' as const, description: 'The spreadsheet ID' },
      requests: {
        type: 'array' as const,
        items: { type: 'object' as const },
        description: 'Array of Sheets API batchUpdate request objects',
      },
    },
    required: ['spreadsheet_id', 'requests'],
  },
};

export const sheetsTools: Tool[] = [
  createSpreadsheetTool,
  createSheetTool,
  listSheetsTool,
  getSheetDataTool,
  updateCellsTool,
  batchUpdateCellsTool,
  batchUpdateTool,
];
