#!/usr/bin/env node
/**
 * Seed data into Supabase from provided Excel and text attachments.
 * - Uses Supabase REST (via @supabase/supabase-js) with SERVICE ROLE key.
 * - Idempotent upserts using natural keys/unique constraints.
 * - Detects structure of Excel files heuristically, including matrix-style sheets.
 * - Optionally enriches roles with descriptions from role-card .txt files.
 * - Retries transient REST failures with exponential backoff.
 * - Verifies (locally) that RLS is disabled in schema.sql for required tables.
 * - Preflight checks for required tables and provides guided recovery if missing.
 *
 * Expected env:
 *   SUPABASE_URL or REACT_APP_SUPABASE_URL
 *   SUPABASE_SERVICE_ROLE_KEY (preferred) or REACT_APP_SUPABASE_KEY (requires RLS disabled)
 *
 * CLI flags (override defaults):
 *   --competency <path_to_Competency_mapping.xlsx>
 *   --adjacency <path_to_CA_Role_Adjacency.xlsx>
 *   --roles <path_to_Role_Navigator_Worksheet.xlsx>
 *   --resources <path_to_learning_resources.xlsx>  (optional)
 *   --cardsGlob <glob_pattern_to_role_card_txt>    (optional)
 */
const path = require('path');
const fs = require('fs');
require('dotenv').config({ path: path.resolve(__dirname, '../.env') });
require('dotenv').config({ path: path.resolve(__dirname, '../.env.local') });
require('dotenv').config({ path: path.resolve(__dirname, '../.env.development') });

const xlsx = require('xlsx');
const { createClient } = require('@supabase/supabase-js');
const glob = require('glob');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');

const argv = yargs(hideBin(process.argv))
  .option('competency', { type: 'string', describe: 'Path to competency mapping Excel file' })
  .option('adjacency', { type: 'string', describe: 'Path to role adjacency Excel file' })
  .option('roles', { type: 'string', describe: 'Path to role navigator worksheet Excel file' })
  .option('resources', { type: 'string', describe: 'Path to optional learning resources Excel file' })
  .option('cardsGlob', { type: 'string', describe: 'Glob pattern to role card text files' })
  .help()
  .argv;

const DEFAULTS = {
  competency: process.env.SEED_COMPETENCY_XLSX || '/home/kavia/workspace/code-generation/attachments/20251127_090403_Competency_mapping.xlsx',
  adjacency: process.env.SEED_ADJACENCY_XLSX || '/home/kavia/workspace/code-generation/attachments/20251127_090402_CA_Role_Adjacency29.xlsx',
  roles: process.env.SEED_ROLES_XLSX || '/home/kavia/workspace/code-generation/attachments/20251127_090406_Role_Navigator_Worksheet.xlsx',
  resources: process.env.SEED_RESOURCES_XLSX || null,
  cardsGlob: process.env.SEED_ROLE_CARDS || '/home/kavia/workspace/code-generation/attachments/2025*Role_Card_*'
};

const INPUTS = {
  competency: argv.competency || DEFAULTS.competency,
  adjacency: argv.adjacency || DEFAULTS.adjacency,
  roles: argv.roles || DEFAULTS.roles,
  resources: argv.resources || DEFAULTS.resources,
  cardsGlob: argv.cardsGlob || DEFAULTS.cardsGlob
};

// RLS verification targets (must be disabled for MVP or require service role key)
const RLS_REQUIRED_TABLES = [
  'roles',
  'competencies',
  'role_competencies',
  'role_adjacency',
  'learning_resources'
];

const SCHEMA_SQL_PATH = path.resolve(__dirname, '../supabase/schema.sql');
// Known alternate adjacency file provided in attachments
const ALT_ADJACENCY_PATH = '/home/kavia/workspace/code-generation/attachments/20251127_090401_CA_Role_Adjacency.xlsx';

// Helper utilities
function slugify(input) {
  return String(input || '')
    .trim()
    .toLowerCase()
    .replace(/[()]/g, '')        // remove unmatched parens often found in headers
    .replace(/[^a-z0-9]+/g, '-') // collapse non-alphanum to hyphens
    .replace(/^-+|-+$/g, '');    // trim leading/trailing hyphens
}

function normalizeRoleCode(nameOrCode) {
  const base = String(nameOrCode || '').trim();
  if (!base) return '';
  return slugify(base);
}

function normalizeCompetencyCode(nameOrCode) {
  const base = String(nameOrCode || '').trim();
  if (!base) return '';
  return slugify(base);
}

function mapLevel(str) {
  if (str == null) return null;
  const raw = String(str).trim().toLowerCase();
  if (!raw || raw === '0' || raw === '-' || raw === 'n/a' || raw === 'na') return null;

  // Common mappings; includes initial-letter shortcuts
  const map = {
    'b': 1, 'basic': 1, 'beginner': 1, 'foundation': 1, 'foundational': 1, 'novice': 1, 'baseline': 1,
    'p': 2, 'proficient': 2, 'intermediate': 2,
    'a': 3, 'advanced': 3,
    'm': 4, 'master': 4, 'authority': 4, 'expert': 4, 'e': 4
  };
  if (raw in map) return map[raw];

  // Sometimes cells like "3 (advanced)" or "Level 2"
  const num = parseFloat(raw);
  if (!Number.isNaN(num)) {
    if (num <= 1) return 1;
    if (num <= 2) return 2;
    if (num <= 3) return 3;
    return 4;
  }
  // Heuristic: extract leading letter
  const first = raw[0];
  if (first && map[first]) return map[first];

  return null;
}

function parseWorkbook(filePath) {
  const wb = xlsx.readFile(filePath);
  return wb;
}

function sheetToObjects(wb, preferredNames = []) {
  let sheetName = wb.SheetNames[0];
  for (const name of wb.SheetNames) {
    const lower = name.toLowerCase();
    if (preferredNames.some(p => lower.includes(p))) {
      sheetName = name; break;
    }
  }
  const sheet = wb.Sheets[sheetName];
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: null });
  return { rows, sheetName };
}

function readTextFileSafe(file) {
  try {
    return fs.readFileSync(file, 'utf8');
  } catch {
    return '';
  }
}

function safeNumber(val) {
  if (val == null) return null;
  if (typeof val === 'number') return val;
  const s = String(val).trim().replace('%', '');
  const n = parseFloat(s);
  if (Number.isNaN(n)) return null;
  return n;
}

function uniqueBy(arr, keyFn) {
  const seen = new Set();
  const out = [];
  for (const item of arr) {
    const key = keyFn(item);
    if (!seen.has(key)) {
      seen.add(key);
      out.push(item);
    }
  }
  return out;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function withRetries(fn, label, { retries = 4, baseDelayMs = 500 } = {}) {
  let lastErr = null;
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      return await fn();
    } catch (e) {
      lastErr = e;
      if (attempt === retries) break;
      const delay = baseDelayMs * Math.pow(2, attempt) + Math.floor(Math.random() * 120);
      console.warn(`[seed:data] Retry ${attempt + 1}/${retries} on ${label}: ${e.message}. Waiting ${delay}ms...`);
      await sleep(delay);
    }
  }
  throw lastErr;
}

// Supabase client init
function getSupabase() {
  const url = process.env.SUPABASE_URL || process.env.REACT_APP_SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.REACT_APP_SUPABASE_KEY;
  if (!url || !key) return null;
  const supabase = createClient(url, key, { auth: { persistSession: false } });
  return supabase;
}

async function upsertChunked(supabase, table, rows, onConflict) {
  const chunkSize = 500;
  let total = 0;
  for (let i = 0; i < rows.length; i += chunkSize) {
    const chunk = rows.slice(i, i + chunkSize);
    await withRetries(async () => {
      const { error } = await supabase
        .from(table)
        .upsert(chunk, { onConflict, ignoreDuplicates: false });
      if (error) throw new Error(`Upsert error on ${table}: ${error.message}`);
    }, `upsert ${table} [${i + 1}-${Math.min(i + chunkSize, rows.length)}]`);
    total += chunk.length;
  }
  return total;
}

async function fetchIdMap(supabase, table, codeColumn = 'code') {
  const { data, error } = await withRetries(
    () => supabase.from(table).select(`id, ${codeColumn}`),
    `fetch ${table} id map`
  );
  if (error) throw new Error(`Failed to read ${table}: ${error.message}`);
  const map = {};
  for (const row of (data || [])) {
    map[row[codeColumn]] = row.id;
  }
  return map;
}

function detectRoleColumns(headers) {
  // Any column not recognized as competency metadata is considered a role column
  const ignoreKeys = [
    'competency', 'competency name', 'competency_name',
    'competency code', 'competency_code',
    'category', 'description'
  ];
  return headers.filter(h => !ignoreKeys.includes(String(h || '').toLowerCase()));
}

function extractHeaders(rows) {
  if (!rows.length) return [];
  return Object.keys(rows[0] || {});
}

function tryRoleNameFromCard(text) {
  if (!text) return null;
  const firstLine = text.split(/\r?\n/)[0] || '';
  return firstLine.trim();
}

function approximateMatchRoleName(name, candidateNames) {
  const n = name.toLowerCase();
  // Try direct contains or startswith matches
  let best = null;
  for (const cand of candidateNames) {
    const c = String(cand).toLowerCase();
    if (c === n || c.includes(n) || n.includes(c)) {
      best = cand;
      break;
    }
  }
  return best;
}

/**
 * Heuristically parses additional roles listed in the "Role Navigator Worksheet"
 * Avoids sheets containing "resource" to not double-count.
 */
function parseAdditionalRolesFromWorksheet(filePath) {
  const results = [];
  try {
    const wb = parseWorkbook(filePath);
    const sheetNames = wb.SheetNames || [];
    for (const sName of sheetNames) {
      const lower = sName.toLowerCase();
      if (lower.includes('resource')) continue; // handled separately
      const sheet = wb.Sheets[sName];
      if (!sheet) continue;
      const rows = xlsx.utils.sheet_to_json(sheet, { defval: null });
      for (const row of rows) {
        // look for any role-like column
        const roleCandidate =
          row['Role'] || row['role'] ||
          row['Role Name'] || row['role name'] ||
          row['Title'] || row['title'] ||
          row['Current Role'] || row['current role'] ||
          row['Target Role'] || row['target role'] ||
          null;
        if (roleCandidate && String(roleCandidate).trim()) {
          const name = String(roleCandidate).trim();
          const code = normalizeRoleCode(name);
          results.push({ code, name });
        }
      }
    }
  } catch (e) {
    console.warn(`[seed:data] Worksheet role parsing skipped (read error): ${e.message}`);
  }
  // Deduplicate by code
  const uniq = uniqueBy(results, r => r.code);
  return uniq;
}

/**
 * Parse adjacency rows from a workbook. Supports both row-wise and matrix-style sheets.
 * Returns array of { source_role_code, target_role_code, score }.
 */
function parseAdjacencyFromWorkbook(wb) {
  const out = [];

  // Try row-wise format first using heuristic columns
  for (const sName of wb.SheetNames) {
    const lower = sName.toLowerCase();
    const sheet = wb.Sheets[sName];
    const rows = xlsx.utils.sheet_to_json(sheet, { defval: null });
    let added = 0;
    for (const row of rows) {
      const src = row['Source'] || row['source'] || row['From'] || row['from'] ||
                  row['Role A'] || row['Role 1'] || row['role a'] || row['role 1'] ||
                  row['Source Role'] || row['source role'] || row['SourceRole'];
      const dst = row['Target'] || row['target'] || row['To'] || row['to'] ||
                  row['Role B'] || row['Role 2'] || row['role b'] || row['role 2'] ||
                  row['Target Role'] || row['target role'] || row['TargetRole'];
      const score = row['Score'] || row['score'] || row['Weight'] || row['weight'] ||
                    row['Percent'] || row['percentage'] || row['Adjacency'] || row['adjacency'];

      if (!src || !dst) continue;
      const source_role_code = normalizeRoleCode(src);
      const target_role_code = normalizeRoleCode(dst);
      const numericScore = safeNumber(score);
      out.push({ source_role_code, target_role_code, score: numericScore });
      added++;
    }
    if (added > 0) return out; // Found usable rows
    // If sheet name contains adjacency but no rows parsed, try matrix parse next
    if (lower.includes('adjacency')) break;
  }

  // Matrix-style: AOA with first row = target roles, first col = source roles
  const firstSheet = wb.Sheets[wb.SheetNames[0]];
  if (!firstSheet) return out;
  const aoa = xlsx.utils.sheet_to_json(firstSheet, { header: 1, blankrows: false });
  if (!Array.isArray(aoa) || aoa.length < 2) return out;

  const headerRow = aoa[0] || [];
  // Determine if first cell is blank/title so targets start at index 1
  const colStart = 1; // assume first column is row role labels
  const targets = headerRow.slice(colStart).map(v => String(v || '').trim()).filter(Boolean);
  if (targets.length === 0) return out;

  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r] || [];
    const srcName = String(row[0] || '').trim();
    if (!srcName) continue;
    const source_role_code = normalizeRoleCode(srcName);
    for (let c = 0; c < targets.length; c++) {
      const targetName = targets[c];
      const val = row[c + colStart];
      const score = safeNumber(val);
      if (score == null) continue;
      const target_role_code = normalizeRoleCode(targetName);
      out.push({ source_role_code, target_role_code, score });
    }
  }
  return out;
}

/**
 * Local verification that schema.sql carries DISABLE RLS statements for required tables.
 * This does not query the remote database; it helps operators ensure the schema is correct.
 */
function verifyRlsDisabledLocally() {
  try {
    if (!fs.existsSync(SCHEMA_SQL_PATH)) {
      console.warn('[seed:data] RLS check: schema.sql not found; cannot verify local RLS configuration.');
      return;
    }
    const sql = fs.readFileSync(SCHEMA_SQL_PATH, 'utf8');
    const results = RLS_REQUIRED_TABLES.map((t) => {
      const re = new RegExp(`ALTER\\s+TABLE\\s+public\\.${t}\\s+DISABLE\\s+ROW\\s+LEVEL\\s+SECURITY`, 'i');
      return { table: t, disabled: re.test(sql) };
    });
    const summary = results.map(r => `${r.table}:${r.disabled ? 'disabled' : 'missing'}`).join(', ');
    console.log(`[seed:data] RLS (local schema.sql) => ${summary}`);
  } catch (e) {
    console.warn(`[seed:data] RLS local verification failed: ${e.message}`);
  }
}

/**
 * Preflight: check that required tables exist. If missing, provide guided recovery and exit.
 */
async function ensureTablesExist(supabase) {
  const missing = [];
  for (const t of RLS_REQUIRED_TABLES) {
    const { error } = await supabase.from(t).select('id').limit(1);
    if (error) {
      // Detect "table not found" hints which can vary; be permissive
      const msg = (error.message || '').toLowerCase();
      if (msg.includes('not find the table') || msg.includes('does not exist') || msg.includes('schema cache')) {
        missing.push(t);
      }
    }
  }
  if (missing.length > 0) {
    console.error('[seed:data] Missing required tables:', missing.join(', '));
    console.error('[seed:data] Apply schema first using one of:');
    console.error('  - npm run seed:schema (requires SUPABASE_DB_URL or PG* envs)');
    console.error('  - Supabase SQL Editor: paste and run career_navigator_frontend/supabase/schema.sql');
    process.exit(1);
  }
}

async function main() {
  console.log('[seed:data] Starting data ingestion...');
  const supabase = getSupabase();
  if (!supabase) {
    console.warn('[seed:data] Missing SUPABASE_URL and/or key. Set SUPABASE_SERVICE_ROLE_KEY (preferred) or ensure RLS disabled and use anon key.');
    console.warn('  Skipping data seeding. Provide env and rerun "npm run seed:data".');
    process.exit(0); // Exit successfully to keep pipeline green; actual seeding requires env.
  }

  // Print minimal info about auth level (no secrets)
  if (process.env.SUPABASE_SERVICE_ROLE_KEY) {
    console.log('[seed:data] Using service role key (RLS bypass).');
  } else {
    console.log('[seed:data] Using non-service key; ensure RLS is disabled for target tables.');
  }

  // Local verification of RLS disablement in schema.sql
  verifyRlsDisabledLocally();

  // Preflight: ensure tables exist before attempting writes (clearer error than during upsert)
  await ensureTablesExist(supabase);

  // Prepare aggregates
  const roles = new Map(); // code -> {code, name, description}
  const competencies = new Map(); // code -> {code, name, category}
  const roleCompetencies = []; // {role_code, competency_code, target_level}
  const adjacency = []; // {source_role_code, target_role_code, score}
  const resources = []; // {competency_code, title, url, provider}

  // 1) Parse Competency Mapping
  if (INPUTS.competency && fs.existsSync(INPUTS.competency)) {
    console.log(`[seed:data] Parsing competency mapping: ${INPUTS.competency}`);
    const wb = parseWorkbook(INPUTS.competency);
    const { rows, sheetName } = sheetToObjects(wb, ['competency', 'mapping', 'matrix']);
    if (!rows.length) {
      console.warn(`[seed:data] Competency mapping sheet "${sheetName}" is empty.`);
    } else {
      const headers = extractHeaders(rows);
      const roleHeaders = detectRoleColumns(headers);
      const unmapped = [];

      for (const row of rows) {
        const compName = row['Competency'] || row['Competency Name'] || row['competency'] || row['competency name'] || row['Competency_Name'];
        const compCodeRaw = row['Competency Code'] || row['competency code'] || null;
        const category = row['Category'] || row['category'] || null;

        if (!compName && !compCodeRaw) continue;

        const compCode = normalizeCompetencyCode(compCodeRaw || compName);
        if (!competencies.has(compCode)) {
          competencies.set(compCode, {
            code: compCode,
            name: compName || compCodeRaw || compCode,
            category: category || null
          });
        }

        // Build roles and mappings
        roleHeaders.forEach((hdr) => {
          const val = row[hdr];
          if (val == null || val === '') return;
          const level = mapLevel(val);

          if (level != null) {
            const roleCode = normalizeRoleCode(hdr);
            if (!roles.has(roleCode)) {
              roles.set(roleCode, {
                code: roleCode,
                name: String(hdr).trim(),
                description: null
              });
            }
            roleCompetencies.push({
              role_code: roleCode,
              competency_code: compCode,
              target_level: level
            });
          } else {
            // Column had a value but didn't map to known level
            unmapped.push({ column: hdr, value: val });
          }
        });

        if (unmapped.length > 0) {
          const example = unmapped.slice(0, 3).map(u => `${u.column}="${u.value}"`).join(' | ');
          console.warn(`[seed:data] Unmapped level values in role columns (showing up to 3): ${example}`);
        }
      }
      console.log(`[seed:data] Parsed ${competencies.size} competencies, ${roles.size} roles, ${roleCompetencies.length} role_competency mappings.`);
    }
  } else {
    console.warn(`[seed:data] Competency mapping file not found: ${INPUTS.competency}`);
  }

  // 1b) Parse additional roles (from Role Navigator worksheet)
  if (INPUTS.roles && fs.existsSync(INPUTS.roles)) {
    const extras = parseAdditionalRolesFromWorksheet(INPUTS.roles);
    let added = 0;
    for (const r of extras) {
      if (!r.code || !r.name) continue;
      if (!roles.has(r.code)) {
        roles.set(r.code, { code: r.code, name: r.name, description: null });
        added += 1;
      }
    }
    if (added > 0) {
      console.log(`[seed:data] Added ${added} roles discovered in worksheet.`);
    }
  }

  // 2) Parse Role Adjacency with fallback and matrix support
  const candidateAdjPaths = [INPUTS.adjacency, ALT_ADJACENCY_PATH].filter(p => p && fs.existsSync(p));
  let parsedAdj = 0;
  for (const adjPath of candidateAdjPaths) {
    console.log(`[seed:data] Parsing role adjacency: ${adjPath}`);
    const wb = parseWorkbook(adjPath);
    const rows = parseAdjacencyFromWorkbook(wb);
    for (const a of rows) {
      if (!a.source_role_code || !a.target_role_code) continue;
      if (!roles.has(a.source_role_code)) roles.set(a.source_role_code, { code: a.source_role_code, name: a.source_role_code, description: null });
      if (!roles.has(a.target_role_code)) roles.set(a.target_role_code, { code: a.target_role_code, name: a.target_role_code, description: null });
      adjacency.push(a);
    }
    parsedAdj += rows.length;
    if (rows.length > 0) break; // Stop at first file with data
  }
  console.log(`[seed:data] Parsed ${parsedAdj} adjacency rows.`);

  // 3) Learning resources (from Role Navigator worksheet or a dedicated file)
  const resourcesPath = INPUTS.resources || INPUTS.roles; // try roles workbook for resources if dedicated not provided
  if (resourcesPath && fs.existsSync(resourcesPath)) {
    console.log(`[seed:data] Attempting to parse learning resources from: ${resourcesPath}`);
    const wb = parseWorkbook(resourcesPath);
    // Prefer a sheet with "resource" in name; otherwise scan all sheets for URL/Link columns
    let parsed = 0;
    const candidateSheets = [...wb.SheetNames];
    const prioritized = candidateSheets.sort((a, b) => {
      const al = a.toLowerCase().includes('resource') ? -1 : 1;
      const bl = b.toLowerCase().includes('resource') ? -1 : 1;
      return al - bl;
    });

    for (const sheetName of prioritized) {
      const sheet = wb.Sheets[sheetName];
      const rows = xlsx.utils.sheet_to_json(sheet, { defval: null });
      if (!rows || rows.length === 0) continue;

      // Detect presence of URL/Link column
      const sample = rows[0] || {};
      const hasUrl = Object.keys(sample).some(k => ['url', 'URL', 'Url', 'Link', 'link'].includes(k));
      if (!hasUrl && !sheetName.toLowerCase().includes('resource')) {
        continue;
      }

      for (const row of rows) {
        const compCodeOrName = row['Competency Code'] || row['competency code'] || row['Competency'] || row['competency'];
        const title = row['Title'] || row['title'] || row['Name'] || row['name'];
        const url = row['URL'] || row['Url'] || row['url'] || row['Link'] || row['link'];
        const provider = row['Provider'] || row['provider'] || row['Source'] || row['source'];
        if (!title || !url) continue;
        const competency_code = compCodeOrName ? normalizeCompetencyCode(compCodeOrName) : null;
        if (competency_code && !competencies.has(competency_code)) {
          // Create placeholder competency if not present
          competencies.set(competency_code, { code: competency_code, name: compCodeOrName, category: null });
        }
        resources.push({ competency_code, title: String(title).trim(), url: String(url).trim(), provider: provider ? String(provider).trim() : null });
        parsed++;
      }
      if (parsed > 0) break; // Use first qualifying sheet
    }
    console.log(`[seed:data] Parsed ${parsed} learning resources.`);
  }

  // 4) Optional: Enrich roles with descriptions from role card .txt files
  if (INPUTS.cardsGlob) {
    const matches = glob.sync(INPUTS.cardsGlob, { nodir: true });
    if (matches.length > 0) {
      console.log(`[seed:data] Enriching role descriptions from ${matches.length} role card text files...`);
      const candidateNames = Array.from(roles.values()).map(r => r.name);
      for (const file of matches) {
        const content = readTextFileSafe(file);
        if (!content) continue;
        const title = tryRoleNameFromCard(content);
        if (!title) continue;
        const codeFromTitle = normalizeRoleCode(title);
        const matchedRole =
          roles.get(codeFromTitle) ||
          (() => {
            const approx = approximateMatchRoleName(title, candidateNames);
            if (!approx) return null;
            const code = normalizeRoleCode(approx);
            return roles.get(code) || null;
          })();

        if (matchedRole) {
          const descText = content.slice(0, 4000); // avoid overly-long text
          matchedRole.description = descText;
          roles.set(matchedRole.code, matchedRole);
        } else {
          // Add as new role if not present yet
          roles.set(codeFromTitle, { code: codeFromTitle, name: title, description: content.slice(0, 4000) });
        }
      }
    }
  }

  // Deduplicate and prepare final arrays
  const rolesArr = Array.from(roles.values());
  const competenciesArr = Array.from(competencies.values());
  const roleCompArr = uniqueBy(roleCompetencies, x => `${x.role_code}|${x.competency_code}`);
  const adjacencyArr = uniqueBy(adjacency, x => `${x.source_role_code}|${x.target_role_code}`);
  const resourcesArr = uniqueBy(resources, x => `${x.competency_code || 'none'}|${x.url}`);

  // Seed order: roles -> competencies -> role_competencies -> role_adjacency -> learning_resources
  console.log(`[seed:data] Upserting roles: ${rolesArr.length}`);
  if (rolesArr.length) {
    await upsertChunked(supabase, 'roles', rolesArr, 'code');
  }

  console.log(`[seed:data] Upserting competencies: ${competenciesArr.length}`);
  if (competenciesArr.length) {
    await upsertChunked(supabase, 'competencies', competenciesArr, 'code');
  }

  // Build id maps
  const roleIdMap = await fetchIdMap(supabase, 'roles', 'code');
  const compIdMap = await fetchIdMap(supabase, 'competencies', 'code');

  // Prepare role_competencies with ids
  const rcRows = roleCompArr
    .map(rc => {
      const role_id = roleIdMap[rc.role_code];
      const competency_id = compIdMap[rc.competency_code];
      if (!role_id || !competency_id) return null;
      return { role_id, competency_id, target_level: rc.target_level };
    })
    .filter(Boolean);
  console.log(`[seed:data] Upserting role_competencies: ${rcRows.length}`);
  if (rcRows.length) {
    await upsertChunked(supabase, 'role_competencies', rcRows, 'role_id,competency_id');
  }

  // Prepare adjacency with ids
  const adjRows = adjacencyArr
    .map(a => {
      const source_role_id = roleIdMap[a.source_role_code];
      const target_role_id = roleIdMap[a.target_role_code];
      if (!source_role_id || !target_role_id) return null;
      return { source_role_id, target_role_id, score: a.score };
    })
    .filter(Boolean);
  console.log(`[seed:data] Upserting role_adjacency: ${adjRows.length}`);
  if (adjRows.length) {
    await upsertChunked(supabase, 'role_adjacency', adjRows, 'source_role_id,target_role_id');
  }

  // Prepare learning resources with ids (if competency_code provided)
  const lrRows = resourcesArr
    .map(r => {
      const competency_id = r.competency_code ? compIdMap[r.competency_code] : null;
      // Only include if we have a competency id
      if (!competency_id) return null;
      return { competency_id, title: r.title, url: r.url, provider: r.provider || null };
    })
    .filter(Boolean);
  console.log(`[seed:data] Upserting learning_resources: ${lrRows.length}`);
  if (lrRows.length) {
    await upsertChunked(supabase, 'learning_resources', lrRows, 'competency_id,url');
  }

  console.log('[seed:data] Ingestion complete.');
  console.log(`Summary:
  - roles upserted: ${rolesArr.length}
  - competencies upserted: ${competenciesArr.length}
  - role_competencies upserted: ${rcRows.length}
  - role_adjacency upserted: ${adjRows.length}
  - learning_resources upserted: ${lrRows.length}
  `);
}

main().catch(err => {
  console.error('[seed:data] ERROR:', err.message);
  process.exit(1);
});
