/**
 * Google Ads MCP Server
 *
 * Entry point for the Model Context Protocol server that provides
 * conversational access to Google Ads API for campaign analytics.
 *
 * Remote HTTP server with Google OAuth authentication.
 */

import 'dotenv/config';
import express from 'express';
import session from 'express-session';
import cors from 'cors';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { env } from './config/index.js';
import { requireAuthForToolCall } from './middleware/auth.js';
import authRoutes from './routes/auth.js';
import oauthRoutes from './routes/oauth.js';
import { AccessTokenStore } from './auth/device-code.js';
import { tools } from './tools/index.js';
import { getAccounts } from './tools/accounts.js';
import { getCampaignPerformance } from './tools/campaigns.js';
import { getAdGroupPerformance } from './tools/ad-groups.js';
import { getAdPerformance } from './tools/ads.js';
import { getKeywordPerformance } from './tools/keywords.js';
import { getSearchTerms } from './tools/search-terms.js';
import { getAssetGroupPerformance } from './tools/asset-groups.js';
import { getConversions } from './tools/conversions.js';
import { getDemographicPerformance } from './tools/demographic-performance.js';
import { getBudgetAnalysis } from './tools/budget-analysis.js';
import { exportData } from './tools/data-export.js';
import { search } from './tools/search.js';

// Extend express-session types
declare module 'express-session' {
  interface SessionData {
    authRequest?: {
      clientId: string;
      redirectUri: string;
      state?: string;
      codeChallenge?: string;
      codeChallengeMethod?: string;
    };
  }
}

/**
 * Initialize MCP server
 */
const server = new Server(
  {
    name: 'google-ads-mcp',
    version: '0.1.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

/**
 * Register tools/list handler
 */
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return { tools };
});

/**
 * Register tools/call handler
 */
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    let result: string;

    switch (name) {
      case 'get-account':
        result = await getAccounts();
        break;
      case 'get-campaign-performance':
        result = await getCampaignPerformance(args);
        break;
      case 'get-ad-group-performance':
        result = await getAdGroupPerformance(args);
        break;
      case 'get-ad-performance':
        result = await getAdPerformance(args);
        break;
      case 'get-keyword-performance':
        result = await getKeywordPerformance(args);
        break;
      case 'get-search-terms':
        result = await getSearchTerms(args);
        break;
      case 'get-asset-group-performance':
        result = await getAssetGroupPerformance(args);
        break;
      case 'get-conversions':
        result = await getConversions(args);
        break;
      case 'get-demographic-performance':
        result = await getDemographicPerformance(args);
        break;
      case 'get-budget-analysis':
        result = await getBudgetAnalysis(args);
        break;
      case 'export-data':
        result = await exportData(args);
        break;
      case 'search':
        result = await search(args);
        break;
      default:
        throw new Error(`Unknown tool: ${name}`);
    }

    return {
      content: [{ type: 'text', text: result }],
    };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    return {
      content: [{ type: 'text', text: `Error: ${errorMessage}` }],
      isError: true,
    };
  }
});

/**
 * Initialize access token store
 */
const accessTokenStore = new AccessTokenStore();
global.accessTokenStore = accessTokenStore;

/**
 * Start MCP server with HTTP transport and OAuth authentication
 */
async function main() {
  const app = express();

  // Trust proxy in production
  if (env.NODE_ENV === 'production') {
    app.set('trust proxy', 1);
  }

  // CORS
  app.use(cors({
    origin: true,
    credentials: true,
    methods: ['GET', 'POST', 'OPTIONS', 'DELETE'],
    allowedHeaders: [
      'Content-Type',
      'Authorization',
      'Cookie',
      'Accept',
      'Mcp-Session-Id',
      'MCP-Protocol-Version',
      'Last-Event-ID',
    ],
    exposedHeaders: ['Content-Type', 'Mcp-Session-Id'],
  }));

  // Middleware
  app.use(express.json());
  app.use(express.urlencoded({ extended: true }));
  app.use(session({
    secret: env.SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      secure: env.NODE_ENV === 'production',
      sameSite: 'lax',
      maxAge: env.SESSION_TTL,
    },
    name: 'connect.sid',
  }));

  // MCP Server metadata endpoint
  app.get('/mcp-metadata', (_req, res) => {
    res.json({
      name: 'google-ads',
      displayName: 'Google Ads Analytics',
      description: 'Access Google Ads API for comprehensive campaign analytics',
      version: '0.1.0',
      transport: 'http',
      authentication: {
        type: 'oauth',
        provider: 'google',
        automatic: true,
      },
      capabilities: [
        'Campaign performance analytics',
        'Ad group and ad-level metrics',
        'Keyword performance analysis',
        'Search term reports',
        'PMAX asset group performance',
        'Conversion tracking',
        'Budget pacing analysis',
        'Demographic performance segmentation',
      ],
    });
  });

  // Health check
  app.get('/health', (_req, res) => {
    res.json({
      status: 'ok',
      version: '0.1.0',
      timestamp: new Date().toISOString(),
    });
  });

  // OAuth routes
  app.use(oauthRoutes);

  // Auth routes
  app.use('/auth', authRoutes);

  // MCP transport — per-request transports with session management
  const transports = new Map<string, StreamableHTTPServerTransport>();

  // MCP endpoints
  app.get('/mcp', requireAuthForToolCall, async (req, res) => {
    const sessionId = req.headers['mcp-session-id'] as string | undefined;
    if (!sessionId || !transports.has(sessionId)) {
      res.status(400).json({ error: 'Invalid or missing session ID' });
      return;
    }
    const transport = transports.get(sessionId)!;
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.setHeader('X-Accel-Buffering', 'no');

    try {
      await transport.handleRequest(req, res);
    } catch (error) {
      console.error('[MCP] GET /mcp error:', error);
      if (!res.headersSent) {
        res.status(500).json({ error: 'Internal server error' });
      }
    }
  });

  const mcpBodyParser = express.text({ type: '*/*' });
  app.post('/mcp', mcpBodyParser, requireAuthForToolCall, async (req, res) => {
    const sessionId = req.headers['mcp-session-id'] as string | undefined;

    // Existing session
    if (sessionId && transports.has(sessionId)) {
      const transport = transports.get(sessionId)!;
      try {
        await transport.handleRequest(req, res, req.body);
      } catch (error) {
        console.error('[MCP] POST /mcp error:', error);
        if (!res.headersSent) {
          res.status(500).json({ error: 'Internal server error' });
        }
      }
      return;
    }

    // New session — create transport with session ID generator
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => crypto.randomUUID(),
      onsessioninitialized: (id) => {
        transports.set(id, transport);
        console.error(`[MCP] Session created: ${id}`);
      },
    });

    transport.onclose = () => {
      const id = transport.sessionId;
      if (id) {
        transports.delete(id);
        console.error(`[MCP] Session closed: ${id}`);
      }
    };

    await server.connect(transport);

    try {
      await transport.handleRequest(req, res, req.body);
    } catch (error) {
      console.error('[MCP] POST /mcp error:', error);
      if (!res.headersSent) {
        res.status(500).json({ error: 'Internal server error' });
      }
    }
  });

  // Handle session cleanup via DELETE
  app.delete('/mcp', async (req, res) => {
    const sessionId = req.headers['mcp-session-id'] as string | undefined;
    if (sessionId && transports.has(sessionId)) {
      const transport = transports.get(sessionId)!;
      await transport.handleRequest(req, res);
      transports.delete(sessionId);
    } else {
      res.status(404).json({ error: 'Session not found' });
    }
  });

  // Start HTTP server
  const PORT = env.PORT;
  const HOST = env.HOST;

  app.listen(PORT, HOST, () => {
    console.error(`Google Ads MCP server running on http://${HOST}:${PORT}`);
    console.error(`Health check: http://${HOST}:${PORT}/health`);
    console.error(`Environment: ${env.NODE_ENV}`);
  });
}

main().catch((error) => {
  console.error('Failed to start MCP server:', error);
  process.exit(1);
});
