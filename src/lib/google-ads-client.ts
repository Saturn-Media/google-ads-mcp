/**
 * Google Ads API Client
 *
 * Wrapper around google-ads-api that handles authentication,
 * GAQL query execution, and result formatting.
 */

import { GoogleAdsApi, Customer } from 'google-ads-api';
import { env } from '../config/index.js';

let clientInstance: GoogleAdsApi | null = null;
let customerInstance: Customer | null = null;

/**
 * Get or create the Google Ads API client singleton
 */
function getClient(): GoogleAdsApi {
  if (!clientInstance) {
    clientInstance = new GoogleAdsApi({
      client_id: env.GOOGLE_ADS_CLIENT_ID,
      client_secret: env.GOOGLE_ADS_CLIENT_SECRET,
      developer_token: env.GOOGLE_ADS_DEVELOPER_TOKEN,
    });
  }
  return clientInstance;
}

/**
 * Get or create the customer instance singleton
 */
function getCustomer(): Customer {
  if (!customerInstance) {
    const client = getClient();
    customerInstance = client.Customer({
      customer_id: env.GOOGLE_ADS_CUSTOMER_ID,
      login_customer_id: env.GOOGLE_ADS_LOGIN_CUSTOMER_ID,
      refresh_token: env.GOOGLE_ADS_REFRESH_TOKEN,
    });
  }
  return customerInstance;
}

/**
 * Execute a GAQL query and return parsed results
 *
 * @param query - Google Ads Query Language (GAQL) query string
 * @returns Array of result rows
 */
export async function executeQuery(query: string): Promise<any[]> {
  const customer = getCustomer();
  try {
    const results = await customer.query(query);
    return results;
  } catch (error) {
    if (error instanceof Error) {
      throw new Error(`Google Ads API error: ${error.message}`);
    }
    throw new Error('Unknown error querying Google Ads API');
  }
}

/**
 * Execute a GAQL query for a specific customer account
 *
 * Creates a customer instance for the given customer_id using the
 * same MCC credentials. Allows querying any account accessible
 * via the login_customer_id.
 *
 * @param query - Google Ads Query Language (GAQL) query string
 * @param customerId - The customer account ID to query (no dashes)
 * @returns Array of result rows
 */
export async function executeQueryForCustomer(query: string, customerId: string): Promise<any[]> {
  const client = getClient();
  const customer = client.Customer({
    customer_id: customerId,
    login_customer_id: env.GOOGLE_ADS_LOGIN_CUSTOMER_ID,
    refresh_token: env.GOOGLE_ADS_REFRESH_TOKEN,
  });
  try {
    const results = await customer.query(query);
    return results;
  } catch (error) {
    if (error instanceof Error) {
      throw new Error(`Google Ads API error: ${error.message}`);
    }
    throw new Error('Unknown error querying Google Ads API');
  }
}

/**
 * List all customer accounts accessible via the authenticated credentials
 *
 * @returns Array of accessible customer resource names
 */
export async function listAccessibleCustomers(): Promise<any> {
  const client = getClient();
  try {
    const response = await client.listAccessibleCustomers(env.GOOGLE_ADS_REFRESH_TOKEN);
    return response;
  } catch (error) {
    if (error instanceof Error) {
      throw new Error(`Google Ads API error: ${error.message}`);
    }
    throw new Error('Unknown error listing accessible customers');
  }
}

/**
 * Execute a GAQL query and return results via SearchStream (for large datasets)
 *
 * @param query - Google Ads Query Language (GAQL) query string
 * @returns Array of result rows
 */
export async function executeStreamQuery(query: string): Promise<any[]> {
  const customer = getCustomer();
  try {
    const stream = customer.queryStream(query);
    const results: any[] = [];
    for await (const row of stream) {
      results.push(row);
    }
    return results;
  } catch (error) {
    if (error instanceof Error) {
      throw new Error(`Google Ads API error: ${error.message}`);
    }
    throw new Error('Unknown error querying Google Ads API');
  }
}

/**
 * Convert cost_micros (millionths of account currency) to dollars
 *
 * Google Ads API returns costs in micros (1,000,000 micros = $1.00)
 *
 * @param micros - Cost in micros
 * @returns Cost in dollars, rounded to 2 decimal places
 */
export function microsToDollars(micros: number | string | Long): number {
  const value = typeof micros === 'string' ? parseInt(micros, 10) : Number(micros);
  return parseFloat((value / 1_000_000).toFixed(2));
}

/**
 * Build a date range clause for GAQL queries
 *
 * @param dateRange - Date range preset
 * @returns GAQL WHERE clause for date range
 */
export function buildDateRangeClause(dateRange: string): string {
  switch (dateRange) {
    case 'TODAY':
      return "segments.date = TODAY";
    case 'YESTERDAY':
      return "segments.date = YESTERDAY";
    case 'LAST_7_DAYS':
      return "segments.date DURING LAST_7_DAYS";
    case 'LAST_14_DAYS':
      return "segments.date DURING LAST_14_DAYS";
    case 'LAST_30_DAYS':
      return "segments.date DURING LAST_30_DAYS";
    case 'THIS_MONTH':
      return "segments.date DURING THIS_MONTH";
    case 'LAST_MONTH':
      return "segments.date DURING LAST_MONTH";
    case 'LAST_90_DAYS': {
      const end = new Date();
      const start = new Date();
      start.setDate(start.getDate() - 90);
      const startStr = start.toISOString().split('T')[0];
      const endStr = end.toISOString().split('T')[0];
      return `segments.date BETWEEN '${startStr}' AND '${endStr}'`;
    }
    default:
      return "segments.date DURING LAST_30_DAYS";
  }
}

/**
 * Format a GAQL result row's metrics into a clean object
 */
export function formatMetrics(row: any): Record<string, number> {
  const metrics: Record<string, number> = {};
  const m = row.metrics;

  if (!m) return metrics;

  if (m.impressions !== undefined) metrics.impressions = Number(m.impressions);
  if (m.clicks !== undefined) metrics.clicks = Number(m.clicks);
  if (m.cost_micros !== undefined) metrics.cost = microsToDollars(m.cost_micros);
  if (m.ctr !== undefined) metrics.ctr = parseFloat((Number(m.ctr) * 100).toFixed(2));
  if (m.average_cpc !== undefined) metrics.avg_cpc = microsToDollars(m.average_cpc);
  if (m.average_cpm !== undefined) metrics.avg_cpm = microsToDollars(m.average_cpm);
  if (m.conversions !== undefined) metrics.conversions = parseFloat(Number(m.conversions).toFixed(2));
  if (m.conversions_value !== undefined) metrics.conversions_value = parseFloat(Number(m.conversions_value).toFixed(2));
  if (m.cost_per_conversion !== undefined) metrics.cost_per_conversion = microsToDollars(m.cost_per_conversion);
  if (m.interaction_rate !== undefined) metrics.interaction_rate = parseFloat((Number(m.interaction_rate) * 100).toFixed(2));

  return metrics;
}

// Type alias for google-ads-api Long values
type Long = { low: number; high: number; unsigned: boolean };
