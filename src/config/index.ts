/**
 * Environment Configuration
 *
 * Validates and exports environment variables required for Google Ads API.
 * Uses zod for runtime validation with clear error messages.
 */

import { z } from 'zod';

const envSchema = z.object({
  // Google Ads API credentials
  GOOGLE_ADS_CLIENT_ID: z
    .string()
    .min(1, 'GOOGLE_ADS_CLIENT_ID is required. See .env.example for setup.'),
  GOOGLE_ADS_CLIENT_SECRET: z
    .string()
    .min(1, 'GOOGLE_ADS_CLIENT_SECRET is required. See .env.example for setup.'),
  GOOGLE_ADS_REFRESH_TOKEN: z
    .string()
    .min(1, 'GOOGLE_ADS_REFRESH_TOKEN is required. Run scripts/get_refresh_token.py to obtain one.'),
  GOOGLE_ADS_DEVELOPER_TOKEN: z
    .string()
    .min(1, 'GOOGLE_ADS_DEVELOPER_TOKEN is required. Apply at https://ads.google.com/aw/apicenter'),
  GOOGLE_ADS_LOGIN_CUSTOMER_ID: z
    .string()
    .min(1, 'GOOGLE_ADS_LOGIN_CUSTOMER_ID is required (MCC account ID, no dashes).'),
  GOOGLE_ADS_CUSTOMER_ID: z
    .string()
    .min(1, 'GOOGLE_ADS_CUSTOMER_ID is required (ad account ID, no dashes).'),

  // Google Sheets API credentials (optional — uses same client ID/secret as Ads)
  GOOGLE_SHEETS_REFRESH_TOKEN: z
    .string()
    .optional(),
  GOOGLE_SHEETS_CLIENT_ID: z
    .string()
    .optional(),
  GOOGLE_SHEETS_CLIENT_SECRET: z
    .string()
    .optional(),

  // Server configuration
  PORT: z
    .string()
    .optional()
    .default('3000')
    .transform((val) => parseInt(val, 10)),
  HOST: z
    .string()
    .optional()
    .default('0.0.0.0'),
  NODE_ENV: z
    .enum(['development', 'production', 'test'])
    .optional()
    .default('development'),

  // Google OAuth Configuration (for team auth)
  GOOGLE_OAUTH_CLIENT_ID: z
    .string()
    .optional(),
  GOOGLE_OAUTH_CLIENT_SECRET: z
    .string()
    .optional(),
  GOOGLE_OAUTH_CALLBACK_URL: z
    .string()
    .url()
    .optional()
    .default('http://localhost:3000/auth/callback'),

  // Session
  SESSION_SECRET: z
    .string()
    .min(32, 'SESSION_SECRET must be at least 32 characters. Generate with: openssl rand -base64 32'),
  SESSION_TTL: z
    .string()
    .optional()
    .default('86400000')
    .transform((val) => parseInt(val, 10)),
});

// Parse and validate environment variables
const parseEnv = () => {
  try {
    return envSchema.parse({
      GOOGLE_ADS_CLIENT_ID: process.env.GOOGLE_ADS_CLIENT_ID,
      GOOGLE_ADS_CLIENT_SECRET: process.env.GOOGLE_ADS_CLIENT_SECRET,
      GOOGLE_ADS_REFRESH_TOKEN: process.env.GOOGLE_ADS_REFRESH_TOKEN,
      GOOGLE_ADS_DEVELOPER_TOKEN: process.env.GOOGLE_ADS_DEVELOPER_TOKEN,
      GOOGLE_ADS_LOGIN_CUSTOMER_ID: process.env.GOOGLE_ADS_LOGIN_CUSTOMER_ID,
      GOOGLE_ADS_CUSTOMER_ID: process.env.GOOGLE_ADS_CUSTOMER_ID,
      GOOGLE_SHEETS_REFRESH_TOKEN: process.env.GOOGLE_SHEETS_REFRESH_TOKEN,
      GOOGLE_SHEETS_CLIENT_ID: process.env.GOOGLE_SHEETS_CLIENT_ID,
      GOOGLE_SHEETS_CLIENT_SECRET: process.env.GOOGLE_SHEETS_CLIENT_SECRET,
      PORT: process.env.PORT,
      HOST: process.env.HOST,
      NODE_ENV: process.env.NODE_ENV,
      GOOGLE_OAUTH_CLIENT_ID: process.env.GOOGLE_OAUTH_CLIENT_ID,
      GOOGLE_OAUTH_CLIENT_SECRET: process.env.GOOGLE_OAUTH_CLIENT_SECRET,
      GOOGLE_OAUTH_CALLBACK_URL: process.env.GOOGLE_OAUTH_CALLBACK_URL,
      SESSION_SECRET: process.env.SESSION_SECRET,
      SESSION_TTL: process.env.SESSION_TTL,
    });
  } catch (error) {
    if (error instanceof z.ZodError) {
      const messages = error.issues.map((err) => `  - ${err.path.join('.')}: ${err.message}`);
      throw new Error(
        `Environment validation failed:\n${messages.join('\n')}\n\nPlease check your .env file.`
      );
    }
    throw error;
  }
};

export const env = parseEnv();
