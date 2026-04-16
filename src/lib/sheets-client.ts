/**
 * Google Sheets API Client
 *
 * Provides authenticated access to Google Sheets API using OAuth2 refresh token.
 * Uses the same Google Cloud project as the Ads client but with Sheets-scoped token.
 */

import { google, sheets_v4 } from 'googleapis';

let sheetsInstance: sheets_v4.Sheets | null = null;
let driveInstance: ReturnType<typeof google.drive> | null = null;

function getAuth() {
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_SHEETS_CLIENT_ID || process.env.GOOGLE_ADS_CLIENT_ID,
    process.env.GOOGLE_SHEETS_CLIENT_SECRET || process.env.GOOGLE_ADS_CLIENT_SECRET,
  );
  oauth2Client.setCredentials({
    refresh_token: process.env.GOOGLE_SHEETS_REFRESH_TOKEN,
  });
  return oauth2Client;
}

export function getSheetsClient(): sheets_v4.Sheets {
  if (!sheetsInstance) {
    sheetsInstance = google.sheets({ version: 'v4', auth: getAuth() });
  }
  return sheetsInstance;
}

export function getDriveClient() {
  if (!driveInstance) {
    driveInstance = google.drive({ version: 'v3', auth: getAuth() });
  }
  return driveInstance;
}
