/**
 * Audit Sheets REST Endpoint
 *
 * POST /api/sheets/write-audit
 * Accepts classified audit data as JSON and creates a formatted Google Sheet
 * with all 4 tabs (Executive Summary, Full Review, Implementation Plan, Negative Gaps).
 *
 * This bypasses MCP tool call size limits by handling the data-heavy
 * Sheets write server-side.
 */

import { Router, Request, Response } from 'express';
import { getSheetsClient } from '../lib/sheets-client.js';

const router = Router();

interface AuditData {
  client_name: string;
  date_range: string;
  currency: string;
  executive_summary: {
    total_spend: number;
    estimated_waste: number;
    waste_pct: number;
    terms_analysed: number;
    terms_above_threshold: number;
    existing_negatives: number;
    recommended_new_negatives: number;
    action_summary: { action: string; count: number; spend: number }[];
    theme_summary: { theme: string; spend: number; clicks: number; conversions: number; priority: string }[];
  };
  full_review: {
    action: string;
    search_term: string;
    theme: string;
    cost: number;
    clicks: number;
    impressions: number;
    conversions: number;
    ctr: number;
    avg_cpc: number;
    campaign: string;
    ad_group: string;
    keyword: string;
    match_type: string;
    reason: string;
  }[];
  implementation_plan: {
    priority: string;
    negative_keyword: string;
    match_type: string;
    campaign: string;
    theme: string;
    est_spend_blocked: number;
    rationale: string;
  }[];
  negative_gaps: {
    term: string;
    currently_negative: string;
    weekly_spend_leaking: number;
    terms_matched: number;
    status: string;
  }[];
}

router.post('/api/sheets/write-audit', async (req: Request, res: Response) => {
  try {
    const data: AuditData = req.body;

    if (!data.client_name || !data.full_review) {
      res.status(400).json({ error: 'Missing required fields: client_name, full_review' });
      return;
    }

    const sheets = getSheetsClient();
    const title = `${data.client_name} — Negative Keyword Audit ${data.date_range}`;

    // 1. Create spreadsheet
    const createRes = await sheets.spreadsheets.create({
      requestBody: {
        properties: { title },
        sheets: [
          { properties: { title: 'Executive Summary', index: 0 } },
          { properties: { title: 'Full Review', index: 1 } },
          { properties: { title: 'Implementation Plan', index: 2 } },
          { properties: { title: 'Negative Gaps', index: 3 } },
        ],
      },
      fields: 'spreadsheetId,sheets.properties.sheetId',
    });

    const spreadsheetId = createRes.data.spreadsheetId!;
    const sheetIds = createRes.data.sheets!.map(s => s.properties!.sheetId!);

    // 2. Build all tab data
    const currency = data.currency || 'USD';
    const es = data.executive_summary;

    // Executive Summary tab
    const esData: any[][] = [
      [title],
      [`Date Range: ${data.date_range}`, '', `Currency: ${currency}`],
      [],
      ['Metric', 'Value'],
      ['Total Spend', es?.total_spend || 0],
      ['Estimated Waste', es?.estimated_waste || 0],
      ['Waste %', es?.waste_pct ? `${es.waste_pct}%` : '0%'],
      ['Terms Analysed', es?.terms_analysed || 0],
      ['Terms Above Threshold', es?.terms_above_threshold || 0],
      ['Existing Negatives', es?.existing_negatives || 0],
      ['Recommended New Negatives', es?.recommended_new_negatives || 0],
      [],
      ['Action Summary'],
      ['Action', 'Count', 'Spend'],
      ...(es?.action_summary || []).map(a => [a.action, a.count, a.spend]),
      [],
      ['Theme Summary'],
      ['Theme', 'Spend', 'Clicks', 'Conversions', 'Priority'],
      ...(es?.theme_summary || []).map(t => [t.theme, t.spend, t.clicks, t.conversions, t.priority]),
    ];

    // Full Review tab
    const frHeaders = ['Action', 'Search Term', 'Theme', 'Cost', 'Clicks', 'Impr', 'Conv', 'CTR', 'Avg CPC', 'Campaign', 'Ad Group', 'Keyword', 'Match Type', 'Reason'];
    const frData: any[][] = [
      [`${data.client_name} — Full Review`],
      [`${data.date_range}`],
      [],
      [],
      [],
      frHeaders,
      ...data.full_review.map(r => [
        r.action, r.search_term, r.theme, r.cost, r.clicks,
        r.impressions, r.conversions, `${r.ctr}%`, r.avg_cpc,
        r.campaign, r.ad_group, r.keyword, r.match_type, r.reason,
      ]),
    ];

    // Implementation Plan tab
    const ipHeaders = ['Priority', 'Negative Keyword', 'Match Type', 'Campaign', 'Theme', 'Est. Weekly Spend Blocked', 'Rationale'];
    const ipData: any[][] = [
      [`${data.client_name} — Implementation Plan`],
      [],
      ipHeaders,
      ...(data.implementation_plan || []).map(r => [
        r.priority, r.negative_keyword, r.match_type, r.campaign,
        r.theme, r.est_spend_blocked, r.rationale,
      ]),
    ];

    // Negative Gaps tab
    const ngHeaders = ['Term', 'Currently Negative?', 'Weekly Spend Leaking', 'Terms Matched', 'Status'];
    const ngData: any[][] = [
      [`${data.client_name} — Negative Gaps`],
      [],
      ngHeaders,
      ...(data.negative_gaps || []).map(r => [
        r.term, r.currently_negative, r.weekly_spend_leaking,
        r.terms_matched, r.status,
      ]),
    ];

    // 3. Write all tabs in one batchUpdate
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [
          { range: 'Executive Summary!A1', values: esData },
          { range: 'Full Review!A1', values: frData },
          { range: 'Implementation Plan!A1', values: ipData },
          { range: 'Negative Gaps!A1', values: ngData },
        ],
      },
    } as any);

    // 4. Apply formatting
    const actionColors: Record<string, any> = {
      REMOVE: { red: 1, green: 0.9, blue: 0.9 },
      KEEP: { red: 0.9, green: 1, blue: 0.9 },
      WATCH: { red: 1, green: 0.97, blue: 0.88 },
      REVIEW: { red: 0.94, green: 0.94, blue: 0.94 },
    };

    const formatRequests: any[] = [];

    // Freeze header row on Full Review (row 6)
    formatRequests.push({
      updateSheetProperties: {
        properties: { sheetId: sheetIds[1], gridProperties: { frozenRowCount: 6 } },
        fields: 'gridProperties.frozenRowCount',
      },
    });

    // Freeze header row on Implementation Plan (row 3)
    formatRequests.push({
      updateSheetProperties: {
        properties: { sheetId: sheetIds[2], gridProperties: { frozenRowCount: 3 } },
        fields: 'gridProperties.frozenRowCount',
      },
    });

    // Freeze header row on Negative Gaps (row 3)
    formatRequests.push({
      updateSheetProperties: {
        properties: { sheetId: sheetIds[3], gridProperties: { frozenRowCount: 3 } },
        fields: 'gridProperties.frozenRowCount',
      },
    });

    // Header formatting for Full Review (row 6)
    formatRequests.push({
      repeatCell: {
        range: { sheetId: sheetIds[1], startRowIndex: 5, endRowIndex: 6 },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.1, green: 0.1, blue: 0.18 },
            textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });

    // Header formatting for Implementation Plan (row 3)
    formatRequests.push({
      repeatCell: {
        range: { sheetId: sheetIds[2], startRowIndex: 2, endRowIndex: 3 },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.1, green: 0.1, blue: 0.18 },
            textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });

    // Header formatting for Negative Gaps (row 3)
    formatRequests.push({
      repeatCell: {
        range: { sheetId: sheetIds[3], startRowIndex: 2, endRowIndex: 3 },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.1, green: 0.1, blue: 0.18 },
            textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });

    // Conditional formatting for Action column on Full Review
    for (const [action, bgColor] of Object.entries(actionColors)) {
      formatRequests.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [{ sheetId: sheetIds[1], startRowIndex: 6, startColumnIndex: 0, endColumnIndex: 14 }],
            booleanRule: {
              condition: {
                type: 'CUSTOM_FORMULA',
                values: [{ userEnteredValue: `=$A7="${action}"` }],
              },
              format: { backgroundColor: bgColor },
            },
          },
          index: 0,
        },
      });
    }

    // Auto-resize columns
    for (let i = 0; i < 4; i++) {
      formatRequests.push({
        autoResizeDimensions: {
          dimensions: { sheetId: sheetIds[i], dimension: 'COLUMNS', startIndex: 0, endIndex: 20 },
        },
      });
    }

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: formatRequests },
    });

    const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;

    res.json({
      spreadsheet_id: spreadsheetId,
      url,
      title,
      tabs_written: 4,
      full_review_rows: data.full_review.length,
      implementation_plan_rows: data.implementation_plan?.length || 0,
      negative_gaps_rows: data.negative_gaps?.length || 0,
    });

  } catch (error) {
    console.error('[Audit Sheets] Error:', error);
    const msg = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: msg });
  }
});

export default router;
