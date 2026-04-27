/**
 * SiteZeus RevOps Dashboard — Excel Online Sync Script
 *
 * HOW TO USE:
 * 1. In Excel Online, click the "Automate" tab
 * 2. Click "New Script"
 * 3. Delete all placeholder code
 * 4. Paste this entire file
 * 5. Update GITHUB_TOKEN below with your new token (regenerate at GitHub → Settings → Developer Settings → PATs)
 * 6. Click "Run" to sync data to the dashboard
 *
 * ⚠️  IMPORTANT: Regenerate your GitHub token and paste it here.
 *     The previous token was shared in chat and should be treated as compromised.
 */

async function main(workbook: ExcelScript.Workbook): Promise<void> {

  // ── CONFIG — update token after regenerating ──────────────
  const GITHUB_TOKEN = "YOUR_NEW_TOKEN_HERE";
  const GITHUB_OWNER = "GLuchetti";
  const GITHUB_REPO  = "revops-dashboard";
  const FILE_PATH    = "data.json";
  // ─────────────────────────────────────────────────────────

  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getUsedRange();

  if (!range) {
    console.log("Sheet appears to be empty.");
    return;
  }

  const values = range.getValues();

  if (values.length < 2) {
    console.log("No data rows found — make sure your sheet has a header row and at least one data row.");
    return;
  }

  // Build header array from row 1
  const headers = (values[0] as (string | number | boolean)[]).map(h => String(h).trim());

  // Convert each row to an object, skip completely blank rows
  const rows: Record<string, string | number | boolean>[] = [];
  for (let i = 1; i < values.length; i++) {
    const cells = values[i] as (string | number | boolean)[];
    if (cells.every(c => c === "" || c === null || c === undefined)) continue;
    const row: Record<string, string | number | boolean> = {};
    headers.forEach((h, j) => { row[h] = cells[j] ?? ""; });
    rows.push(row);
  }

  // Build the JSON payload
  const payload = {
    lastUpdated: new Date().toISOString(),
    rowCount: rows.length,
    rows
  };

  // Base64-encode (handles Unicode safely)
  const jsonStr  = JSON.stringify(payload);
  const encoded  = btoa(
    encodeURIComponent(jsonStr).replace(/%([0-9A-F]{2})/g, (_, p1) =>
      String.fromCharCode(parseInt(p1, 16))
    )
  );

  const apiUrl = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${FILE_PATH}`;
  const headers2 = {
    "Authorization":          `Bearer ${GITHUB_TOKEN}`,
    "Accept":                 "application/vnd.github.v3+json",
    "X-GitHub-Api-Version":   "2022-11-28",
    "Content-Type":           "application/json"
  };

  // Fetch existing file SHA (required by GitHub API for updates)
  let sha: string | undefined;
  try {
    const getResp = await fetch(apiUrl, { headers: headers2 });
    if (getResp.ok) {
      const existing = await getResp.json() as { sha: string };
      sha = existing.sha;
    }
  } catch {
    // File doesn't exist yet — first push, no SHA needed
  }

  // Push to GitHub
  const body: { message: string; content: string; sha?: string } = {
    message: `RevOps sync — ${new Date().toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}`,
    content: encoded
  };
  if (sha) body.sha = sha;

  const putResp = await fetch(apiUrl, {
    method: "PUT",
    headers: headers2,
    body: JSON.stringify(body)
  });

  if (putResp.ok) {
    console.log(`✅ Dashboard synced! ${rows.length} rows pushed to GitHub.`);
    console.log(`   Dashboard: https://gluchetti.github.io/revops-dashboard`);
  } else {
    const err = await putResp.json() as { message: string };
    console.log(`❌ Sync failed: ${err.message}`);
    console.log("   Check that your GITHUB_TOKEN is valid and has 'repo' scope.");
  }
}
