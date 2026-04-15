/**
 * Search Tool (Raw GAQL)
 *
 * Executes raw Google Ads Query Language (GAQL) queries against any
 * customer account accessible via the MCC login. This provides full
 * flexibility for complex queries that pre-built tools don't cover.
 *
 * Compatible with the Python google-ads-mcp `search` tool interface.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { executeQueryForCustomer, executeQuery } from '../lib/google-ads-client.js';

const SearchSchema = z.object({
  customer_id: z
    .string()
    .describe('The Google Ads customer ID (no dashes). If not provided, uses the default from environment.')
    .optional(),
  fields: z
    .array(z.string())
    .describe('The fields to fetch (e.g. ["campaign.name", "metrics.clicks"])'),
  resource: z
    .string()
    .describe('The resource to query (e.g. "campaign", "search_term_view", "campaign_criterion")'),
  conditions: z
    .array(z.string())
    .describe('WHERE conditions (e.g. ["segments.date BETWEEN \'2026-04-01\' AND \'2026-04-07\'"])')
    .optional(),
  orderings: z
    .array(z.string())
    .describe('ORDER BY clauses (e.g. ["metrics.cost_micros DESC"])')
    .optional(),
  limit: z
    .number()
    .describe('Maximum number of rows to return')
    .optional(),
});

export async function search(args: unknown): Promise<string> {
  const input = SearchSchema.parse(args);

  // Build GAQL query
  const selectClause = input.fields.join(', ');
  let query = `SELECT ${selectClause} FROM ${input.resource}`;

  if (input.conditions && input.conditions.length > 0) {
    query += ` WHERE ${input.conditions.join(' AND ')}`;
  }

  if (input.orderings && input.orderings.length > 0) {
    query += ` ORDER BY ${input.orderings.join(', ')}`;
  }

  if (input.limit) {
    query += ` LIMIT ${input.limit}`;
  }

  try {
    let results: any[];

    if (input.customer_id) {
      results = await executeQueryForCustomer(query, input.customer_id);
    } else {
      results = await executeQuery(query);
    }

    return JSON.stringify({
      rowCount: results.length,
      rows: results,
    }, null, 2);
  } catch (error) {
    if (error instanceof Error) return `Error: ${error.message}`;
    return 'Unknown error executing GAQL query';
  }
}

export const searchTool: Tool = {
  name: 'search',
  description: 'Execute a raw GAQL (Google Ads Query Language) query. Supports any resource, field, and condition. Use for complex queries not covered by pre-built tools.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      customer_id: {
        type: 'string' as const,
        description: 'Google Ads customer ID (no dashes). Omit to use the default account.',
      },
      fields: {
        type: 'array' as const,
        items: { type: 'string' as const },
        description: 'Fields to select (e.g. ["campaign.name", "metrics.clicks"])',
      },
      resource: {
        type: 'string' as const,
        description: 'Resource to query (e.g. "campaign", "search_term_view")',
      },
      conditions: {
        type: 'array' as const,
        items: { type: 'string' as const },
        description: 'WHERE conditions',
      },
      orderings: {
        type: 'array' as const,
        items: { type: 'string' as const },
        description: 'ORDER BY clauses',
      },
      limit: {
        type: 'number' as const,
        description: 'Max rows to return',
      },
    },
    required: ['fields', 'resource'],
  },
};
