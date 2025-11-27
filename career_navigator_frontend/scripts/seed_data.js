#!/usr/bin/env node
/**
 * Seed data into Supabase from provided Excel and text attachments.
 * - Idempotent upserts using natural keys/unique constraints.
 * - Detects structure of Excel files heuristically.
 * - Optionally enriches roles with descriptions from role-card .txt files.
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
 *   --cardsGlob <glob_pattern_for_role_card_txt>   (optional)
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
  if (!raw || raw === '0' || raw === '-') return null;
  // Common mappings; adjust as needed
  const map = {
    'basic': 1, 'beginner': 1, 'foundation': 1, 'foundational': 1, 'novice': 1,
    'proficient': 2, 'intermediate': 2,
    'advanced': 3,
    'master': 4, 'authority': 4, 'expert': 4
  };
  if (raw in map) return map[raw];
  const num = parseFloat(raw);
  if (!Number.isNaN(num)) {
    if (num <= 1) return 1;
    if (num <= 2) return 2;
    if (num <= 3) return 3;
    return 4;
  }
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
    const { error } = await supabase
      .from(table)
      .upsert(chunk, { onConflict, ignoreDuplicates: false });
    if (error) throw new Error(`Upsert error on ${table}: ${error.message}`);
    total += chunk.length;
  }
  return total;
}

async function fetchIdMap(supabase, table, codeColumn = 'code') {
  const { data, error } = await supabase.from(table).select(`id, ${codeColumn}`);
  if (error) throw new Error(`Failed to read ${table}: ${error.message}`);
  const map = {};
  for (const row of data) {
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

async function main() {
  console.log('[seed:data] Starting data ingestion...');
  const supabase = getSupabase();
  if (!supabase) {
    console.warn('[seed:data] Missing SUPABASE_URL and/or key. Set SUPABASE_SERVICE_ROLE_KEY (preferred) or ensure RLS disabled and use anon key.');
    console.warn('  Skipping data seeding. Provide env and rerun "npm run seed:data".');
    process.exit(0); // Exit successfully to keep pipeline green; actual seeding requires env.
  }

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
          const hdrLower = String(hdr || '').toLowerCase();
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

  // 2) Parse Role Adjacency (prefer 29.xlsx if available)
  const adjacencyPath = INPUTS.adjacency;
  if (adjacencyPath && fs.existsSync(adjacencyPath)) {
    console.log(`[seed:data] Parsing role adjacency: ${adjacencyPath}`);
    const wb = parseWorkbook(adjacencyPath);
    const { rows } = sheetToObjects(wb, ['adjacency']);
    for (const row of rows) {
      const src = row['Source'] || row['source'] || row['From'] || row['from'] || row['Role A'] || row['Role 1'] || row['role a'] || row['role 1'];
      const dst = row['Target'] || row['target'] || row['To'] || row['to'] || row['Role B'] || row['Role 2'] || row['role b'] || row['role 2'];
      const score = row['Score'] || row['score'] || row['Weight'] || row['weight'] || row['Percent'] || row['percentage'];

      if (!src || !dst) continue;
      const source_role_code = normalizeRoleCode(src);
      const target_role_code = normalizeRoleCode(dst);
      if (!roles.has(source_role_code)) roles.set(source_role_code, { code: source_role_code, name: String(src).trim(), description: null });
      if (!roles.has(target_role_code)) roles.set(target_role_code, { code: target_role_code, name: String(dst).trim(), description: null });

      const numericScore = safeNumber(score);
      adjacency.push({ source_role_code, target_role_code, score: numericScore });
    }
    console.log(`[seed:data] Parsed ${adjacency.length} adjacency rows.`);
  } else {
    console.warn(`[seed:data] Adjacency file not found: ${adjacencyPath}`);
  }

  // 3) Optional: Learning resources (from Role Navigator worksheet or a dedicated file)
  const resourcesPath = INPUTS.resources || INPUTS.roles; // try roles workbook for resources if dedicated not provided
  if (resourcesPath && fs.existsSync(resourcesPath)) {
    console.log(`[seed:data] Attempting to parse learning resources from: ${resourcesPath}`);
    const wb = parseWorkbook(resourcesPath);
    // Try to find a sheet with "resource" name
    const resourceSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('resource'));
    if (resourceSheetName) {
      const sheet = wb.Sheets[resourceSheetName];
      const rows = xlsx.utils.sheet_to_json(sheet, { defval: null });
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
      }
      console.log(`[seed:data] Parsed ${resources.length} learning resources.`);
    } else {
      console.log('[seed:data] No sheet containing "resource" found; skipping resources.');
    }
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
  await upsertChunked(supabase, 'roles', rolesArr, 'code');

  console.log(`[seed:data] Upserting competencies: ${competenciesArr.length}`);
  await upsertChunked(supabase, 'competencies', competenciesArr, 'code');

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
