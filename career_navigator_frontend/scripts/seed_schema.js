#!/usr/bin/env node
/**
 * Seed Supabase schema and temporarily disable RLS for MVP.
 * This script attempts to apply SQL using a Postgres connection string (SUPABASE_DB_URL or PG* variables).
 * If DB connection info is not provided, it will print next steps and exit 0 (no error).
 */
const fs = require('fs');
const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '../.env') }); // local .env in this folder if present
require('dotenv').config({ path: path.resolve(__dirname, '../.env.local') }); // optional alternate
require('dotenv').config({ path: path.resolve(__dirname, '../.env.development') }); // optional

const { Client } = require('pg');

function buildPgConfig() {
  const connStr = process.env.SUPABASE_DB_URL || process.env.DATABASE_URL;
  if (connStr) {
    return {
      connectionString: connStr,
      ssl: { rejectUnauthorized: false }
    };
  }
  const host = process.env.PGHOST;
  const database = process.env.PGDATABASE;
  const user = process.env.PGUSER;
  const password = process.env.PGPASSWORD;
  const port = process.env.PGPORT ? parseInt(process.env.PGPORT, 10) : undefined;

  if (host && database && user && password) {
    return {
      host, database, user, password, port,
      ssl: { rejectUnauthorized: false }
    };
  }
  return null;
}

async function run() {
  const schemaPath = path.resolve(__dirname, '../supabase/schema.sql');
  if (!fs.existsSync(schemaPath)) {
    console.error(`[seed:schema] ERROR: schema file not found at ${schemaPath}`);
    process.exit(1);
  }
  const sql = fs.readFileSync(schemaPath, 'utf8');

  const pgConfig = buildPgConfig();
  if (!pgConfig) {
    console.warn('[seed:schema] No database connection info provided. Skipping schema application.');
    console.warn('  To auto-apply schema, set one of: SUPABASE_DB_URL or PGHOST/PGDATABASE/PGUSER/PGPASSWORD/PGPORT.');
    console.warn('  You can also apply schema manually via Supabase SQL editor using the file at supabase/schema.sql.');
    process.exit(0); // Exit successfully to allow data seeding step to proceed if RLS is already disabled and tables exist.
  }

  const client = new Client(pgConfig);
  try {
    console.log('[seed:schema] Connecting to database...');
    await client.connect();
    console.log('[seed:schema] Applying schema and disabling RLS...');
    await client.query(sql);
    console.log('[seed:schema] Schema applied successfully.');
  } catch (err) {
    console.error('[seed:schema] ERROR applying schema:', err.message);
    process.exit(1);
  } finally {
    try { await client.end(); } catch {}
  }
}

run();
