/**
 * Tool Registry
 *
 * Central registry of all MCP tools available in this server.
 * Tools are registered here to be discovered by clients via tools/list.
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { accountsTool } from './accounts.js';
import { campaignsTool } from './campaigns.js';
import { adGroupsTool } from './ad-groups.js';
import { adsTool } from './ads.js';
import { keywordsTool } from './keywords.js';
import { searchTermsTool } from './search-terms.js';
import { assetGroupsTool } from './asset-groups.js';
import { conversionsTool } from './conversions.js';
import { demographicPerformanceTool } from './demographic-performance.js';
import { budgetAnalysisTool } from './budget-analysis.js';
import { dataExportTool } from './data-export.js';
import { searchTool } from './search.js';

/**
 * All available MCP tools
 */
export const tools: Tool[] = [
  searchTool,
  accountsTool,
  campaignsTool,
  adGroupsTool,
  adsTool,
  keywordsTool,
  searchTermsTool,
  assetGroupsTool,
  conversionsTool,
  demographicPerformanceTool,
  budgetAnalysisTool,
  dataExportTool,
];
