/* ============================================================
   IMAGE LAB

   A diagnostic page: pick one tab and one cell reference that
   you know holds a pasted-in image, and this runs every method
   the site knows about for reading in-cell image content against
   that single cell, side by side, so you can see which ones
   actually work for your sheet without deploying anything.

   All client-side — no server, no Apps Script required for this
   page to run (though the Apps Script method is included and
   will show as "not configured" if APPS_SCRIPT_URL is blank,
   for comparison against everything else).

   Depends on helpers defined in app.js: SPREADSHEET_ID, API_KEY,
   TAB_GIDS, titleByKey, APPS_SCRIPT_URL, escapeHtml,
   extractImageUrl, parseA1Cell, readSheetImageAnchors,
   OOXML_REL_NS. app.js must load before this file.
   ============================================================ */

function renderLabPage(){

  const el =
    document.getElementById("tab-lab");

  if(!el){
    return;
  }

  const sheetOptions =
    Object.keys(TAB_GIDS)
      .map(key => {

        const label =
          titleByKey[key] || key;

        return `<option value="${key}">${escapeHtml(label)}</option>`;
      })
      .join("");

  el.innerHTML =
    `
      <div class="lab-page">

        <div class="lab-intro">

          <h3>Image Method Lab</h3>

          <p>
            Pick a tab and a cell reference you know holds a
            pasted-in picture — e.g. from an Apps Script execution
            log ("In-cell image found at P166" means the Matchup
            tab, cell P166) — then run every method this site
            knows about against that one cell, side by side, to
            see exactly which ones work for your sheet.
          </p>

          <div class="lab-controls">

            <label>
              Tab
              <select id="lab-sheet">${sheetOptions}</select>
            </label>

            <label>
              Cell
              <input
                type="text"
                id="lab-cell"
                value="P166"
                placeholder="e.g. P166"
              >
            </label>

            <button class="wav-btn" id="lab-run">
              Run all methods
            </button>

          </div>

        </div>

        <div class="lab-results" id="lab-results"></div>

      </div>
    `;

  document
    .getElementById("lab-run")
    .addEventListener("click",runLabMethods);
}


const LAB_METHODS = [
  {
    id:"values-formula",
    name:"Sheets API — values.get (FORMULA)",
    run:labMethodValuesFormula
  },
  {
    id:"values-unformatted",
    name:"Sheets API — values.get (UNFORMATTED_VALUE)",
    run:labMethodValuesUnformatted
  },
  {
    id:"full-cell-dump",
    name:"Sheets API — spreadsheets.get, full cell + includeGridData",
    run:labMethodFullCellDump
  },
  {
    id:"gviz",
    name:"Google Visualization (gviz) endpoint",
    run:labMethodGviz
  },
  {
    id:"csv-export",
    name:"CSV export of just this cell",
    run:labMethodCsvExport
  },
  {
    id:"xlsx-direct",
    name:"Direct .xlsx export fetch (no upload, no proxy)",
    run:labMethodXlsxDirect
  },
  {
    id:"xlsx-proxy",
    name:"Direct .xlsx export via a public CORS proxy",
    run:labMethodXlsxProxy
  },
  {
    id:"drive-api",
    name:"Drive API v3 — files.export",
    run:labMethodDriveApi
  },
  {
    id:"apps-script",
    name:"Apps Script Web App (CellImageExport.gs)",
    run:labMethodAppsScript
  }
];


async function runLabMethods(){

  const sheetKey =
    document.getElementById("lab-sheet").value;

  const cellRaw =
    document.getElementById("lab-cell").value
      .trim()
      .toUpperCase();

  const sheetTitle =
    titleByKey[sheetKey];

  const parsed =
    parseA1Cell(cellRaw);

  const resultsEl =
    document.getElementById("lab-results");

  if(!sheetTitle || !parsed){

    resultsEl.innerHTML =
      `<p class="lab-error">
        Enter a valid cell reference (e.g. P166), and make sure
        the spreadsheet has finished loading first.
      </p>`;

    return;
  }

  const ctx = {
    tabKey:sheetKey,
    sheetTitle:sheetTitle,
    cell:cellRaw,
    row:parsed.row,
    col:parsed.col
  };

  resultsEl.innerHTML =
    LAB_METHODS
      .map(m => labMethodCardHtml(m))
      .join("");

  LAB_METHODS.forEach(async method => {

    let outcome;

    try{
      outcome = await method.run(ctx);
    }catch(err){
      outcome = {
        success:false,
        note:`Threw an error: ${err.message}`
      };
    }

    updateLabMethodCard(method.id,outcome);
  });
}


function labMethodCardHtml(method){

  return `
    <div class="lab-card" id="lab-card-${method.id}" data-status="pending">

      <div class="lab-card-head">
        <span class="lab-status-dot"></span>
        <strong>${escapeHtml(method.name)}</strong>
      </div>

      <div class="lab-card-body">
        <p class="lab-note">Running…</p>
      </div>

    </div>
  `;
}


function updateLabMethodCard(id,outcome){

  const card =
    document.getElementById(`lab-card-${id}`);

  if(!card){
    return;
  }

  card.dataset.status =
    outcome.success ? "success" : "fail";

  const body =
    card.querySelector(".lab-card-body");

  const rawTrimmed =
    outcome.raw && outcome.raw.length > 1500
      ? outcome.raw.slice(0,1500) + "\n…(truncated)"
      : outcome.raw;

  body.innerHTML =
    `
      <p class="lab-note">
        ${escapeHtml(
          outcome.note ||
          (outcome.success ? "Worked." : "Didn't work.")
        )}
      </p>

      ${
        outcome.imageUrl
          ? `
            <div class="lab-image-preview">
              <img src="${outcome.imageUrl}" alt="" loading="lazy">
            </div>
            <input
              type="text"
              readonly
              value="${escapeHtml(outcome.imageUrl)}"
              class="lab-url-input"
              onclick="this.select()"
            >
          `
          : ""
      }

      ${
        rawTrimmed
          ? `
            <details class="lab-raw">
              <summary>Raw response</summary>
              <pre>${escapeHtml(rawTrimmed)}</pre>
            </details>
          `
          : ""
      }
    `;
}


/*
 * labFetchJson: fetch + best-effort JSON parse, never throws —
 * every method below treats "couldn't get this" as a normal
 * outcome to report, not a crash.
 */
async function labFetchJson(url){

  try{

    const res =
      await fetch(url,{cache:"no-store"});

    const text =
      await res.text();

    let json = null;

    try{
      json = JSON.parse(text);
    }catch(err){
      // Not JSON — fine, some of these endpoints return CSV or
      // wrapped text, callers handle .text themselves.
    }

    return {ok:res.ok,status:res.status,text:text,json:json};

  }catch(err){

    return {
      ok:false,
      status:0,
      text:String(err),
      json:null,
      threw:true
    };
  }
}


/* ---------- Method 1: values.get, FORMULA render ---------- */

async function labMethodValuesFormula(ctx){

  const range =
    `${ctx.sheetTitle}!${ctx.cell}`;

  const url =
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}` +
    `/values/${encodeURIComponent(range)}` +
    `?valueRenderOption=FORMULA&key=${API_KEY}`;

  const r =
    await labFetchJson(url);

  if(!r.ok){
    return {success:false,note:`HTTP ${r.status}`,raw:r.text};
  }

  const cellVal =
    r.json && r.json.values && r.json.values[0]
      ? r.json.values[0][0]
      : undefined;

  const imgUrl =
    extractImageUrl(cellVal);

  return {
    success:!!imgUrl,
    note:
      imgUrl
        ? "Found an =IMAGE() formula in the cell."
        : cellVal === undefined
          ? "Cell came back empty — this is exactly what happens for a real pasted-in image; the REST API just has nothing to say about it."
          : `Cell value: ${JSON.stringify(cellVal)}`,
    imageUrl:imgUrl,
    raw:JSON.stringify(r.json,null,2)
  };
}


/* ---------- Method 2: values.get, UNFORMATTED_VALUE ---------- */

async function labMethodValuesUnformatted(ctx){

  const range =
    `${ctx.sheetTitle}!${ctx.cell}`;

  const url =
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}` +
    `/values/${encodeURIComponent(range)}` +
    `?valueRenderOption=UNFORMATTED_VALUE&key=${API_KEY}`;

  const r =
    await labFetchJson(url);

  if(!r.ok){
    return {success:false,note:`HTTP ${r.status}`,raw:r.text};
  }

  const cellVal =
    r.json && r.json.values && r.json.values[0]
      ? r.json.values[0][0]
      : undefined;

  return {
    success:false,
    note:
      cellVal === undefined
        ? "Cell came back empty here too — same story, no image data on this render mode either."
        : `Cell value: ${JSON.stringify(cellVal)}`,
    raw:JSON.stringify(r.json,null,2)
  };
}


/* ---------- Method 3: full cell dump via spreadsheets.get ---------- */

async function labMethodFullCellDump(ctx){

  const range =
    `${ctx.sheetTitle}!${ctx.cell}`;

  const url =
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}` +
    `?ranges=${encodeURIComponent(range)}` +
    `&includeGridData=true&key=${API_KEY}`;

  const r =
    await labFetchJson(url);

  if(!r.ok){
    return {success:false,note:`HTTP ${r.status}`,raw:r.text};
  }

  const cellObj =
    r.json &&
    r.json.sheets &&
    r.json.sheets[0] &&
    r.json.sheets[0].data &&
    r.json.sheets[0].data[0] &&
    r.json.sheets[0].data[0].rowData &&
    r.json.sheets[0].data[0].rowData[0] &&
    r.json.sheets[0].data[0].rowData[0].values
      ? r.json.sheets[0].data[0].rowData[0].values[0]
      : null;

  const dump =
    JSON.stringify(cellObj,null,2) || "(nothing)";

  const urlMatch =
    dump.match(/https?:\/\/[^\s"]+/);

  return {
    success:!!urlMatch,
    note:
      urlMatch
        ? "Found an embedded URL somewhere in the full cell metadata."
        : "No URL anywhere in the fully-expanded cell object — this is the most complete view the REST API can give, and it genuinely has nothing for an in-cell image.",
    imageUrl:urlMatch ? urlMatch[0] : null,
    raw:dump
  };
}


/* ---------- Method 4: gviz endpoint ---------- */

async function labMethodGviz(ctx){

  const url =
    `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}` +
    `/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(ctx.sheetTitle)}` +
    `&range=${encodeURIComponent(ctx.cell)}`;

  const r =
    await labFetchJson(url);

  if(r.threw){
    return {
      success:false,
      note:`Request failed — most likely CORS: ${r.text}`
    };
  }

  let extracted = null;

  const m =
    r.text.match(/setResponse\(([\s\S]*)\);?\s*$/);

  if(m){
    try{
      extracted = JSON.parse(m[1]);
    }catch(err){
      // Leave extracted null, fall through to raw text below.
    }
  }

  const cellData =
    extracted &&
    extracted.table &&
    extracted.table.rows &&
    extracted.table.rows[0] &&
    extracted.table.rows[0].c
      ? extracted.table.rows[0].c[0]
      : null;

  return {
    success:false,
    note:
      r.ok
        ? "gviz reached the sheet, but this endpoint only ever carries text/number values — never image data, by design."
        : `HTTP ${r.status}.`,
    raw:cellData ? JSON.stringify(cellData,null,2) : r.text.slice(0,800)
  };
}


/* ---------- Method 5: CSV export of just this cell ---------- */

async function labMethodCsvExport(ctx){

  const gid =
    TAB_GIDS[ctx.tabKey];

  const url =
    `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}` +
    `/export?format=csv&gid=${gid}&range=${encodeURIComponent(ctx.cell)}`;

  const r =
    await labFetchJson(url);

  if(r.threw){
    return {
      success:false,
      note:`Request failed — most likely CORS: ${r.text}`
    };
  }

  return {
    success:false,
    note:
      r.ok
        ? `CSV export returned: "${r.text.trim().slice(0,200)}" — CSV is plain text by definition, it can never carry image data.`
        : `HTTP ${r.status}.`,
    raw:r.text.slice(0,500)
  };
}


/* ---------- Method 6: direct .xlsx export fetch ---------- */

async function labMethodXlsxDirect(ctx){

  const url =
    `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=xlsx`;

  try{

    const res =
      await fetch(url,{cache:"no-store"});

    if(!res.ok){
      return {success:false,note:`HTTP ${res.status} fetching the .xlsx export directly.`};
    }

    const buffer =
      await res.arrayBuffer();

    const imageUrl =
      await labFindXlsxImageAtCell(buffer,ctx.sheetTitle,ctx.row,ctx.col);

    return imageUrl
      ? {
          success:true,
          note:"Fetched the live .xlsx export directly — no upload, no proxy, no Apps Script — and found the image anchored at that exact cell.",
          imageUrl:imageUrl
        }
      : {
          success:false,
          note:"Fetched the .xlsx export fine, but no image was anchored at that exact cell — double-check the row/column."
        };

  }catch(err){

    return {
      success:false,
      note:
        `Blocked — almost certainly CORS: ${err.message}. ` +
        `Google's export endpoint doesn't send an ` +
        `Access-Control-Allow-Origin header, so the browser ` +
        `refuses to let this page's JavaScript read the response, ` +
        `even though the file itself is publicly viewable.`
    };
  }
}


/* ---------- Method 7: .xlsx export via a public CORS proxy ---------- */

async function labMethodXlsxProxy(ctx){

  const target =
    `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=xlsx`;

  const proxyUrl =
    `https://corsproxy.io/?url=${encodeURIComponent(target)}`;

  try{

    const res =
      await fetch(proxyUrl,{cache:"no-store"});

    if(!res.ok){
      return {success:false,note:`Proxy returned HTTP ${res.status}.`};
    }

    const buffer =
      await res.arrayBuffer();

    const imageUrl =
      await labFindXlsxImageAtCell(buffer,ctx.sheetTitle,ctx.row,ctx.col);

    return imageUrl
      ? {
          success:true,
          note:
            "Worked, routed through a public third-party CORS proxy. " +
            "Usable in a pinch, but it depends on a service you " +
            "don't control staying up, staying free, and not rate-limiting " +
            "you — and it routes your sheet's contents through their " +
            "servers. Treat this as a fallback, not a first choice.",
          imageUrl:imageUrl
        }
      : {
          success:false,
          note:"Proxy fetch succeeded but no image was anchored at that exact cell."
        };

  }catch(err){

    return {
      success:false,
      note:`Proxy request failed: ${err.message}`
    };
  }
}


/* ---------- Method 8: Drive API v3 files.export ---------- */

async function labMethodDriveApi(ctx){

  const mimeType =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  const url =
    `https://www.googleapis.com/drive/v3/files/${SPREADSHEET_ID}` +
    `/export?mimeType=${encodeURIComponent(mimeType)}&key=${API_KEY}`;

  const r =
    await labFetchJson(url);

  return {
    success:false,
    note:
      r.ok
        ? "Unexpectedly returned something — check the raw response below."
        : `HTTP ${r.status} — the Drive API's export endpoint requires an OAuth-authorized user for file content, an API key alone isn't enough. This failure is expected.`,
    raw:r.text.slice(0,500)
  };
}


/* ---------- Method 9: the known-working Apps Script comparison ---------- */

async function labMethodAppsScript(ctx){

  if(!APPS_SCRIPT_URL){

    return {
      success:false,
      note:
        "APPS_SCRIPT_URL isn't set. This is the one method known " +
        "to actually work end-to-end — it just needs the one-time " +
        "deployment described in CellImageExport.gs."
    };
  }

  const url =
    `${APPS_SCRIPT_URL}?key=${encodeURIComponent(ctx.tabKey)}` +
    `&startRow=${ctx.row + 1}&endRow=${ctx.row + 1}`;

  const r =
    await labFetchJson(url);

  if(!r.ok || !r.json){
    return {
      success:false,
      note:`HTTP ${r.status} or not valid JSON from the Apps Script endpoint.`,
      raw:r.text.slice(0,500)
    };
  }

  const match =
    (r.json.found || []).find(f => f.col === ctx.col + 1);

  return match
    ? {
        success:true,
        note:"Confirmed working via the deployed Apps Script endpoint.",
        imageUrl:match.url
      }
    : {
        success:false,
        note:"Reached the Apps Script endpoint fine, but it didn't report an image at that exact cell."
      };
}


/*
 * labFindXlsxImageAtCell: given an already-fetched .xlsx file
 * (as an ArrayBuffer) and a sheet name, finds whichever embedded
 * image (if any) is anchored at the given 0-indexed row/col.
 * Reuses readSheetImageAnchors from app.js — same code path the
 * Data Sync tab's manual upload uses — just scoped to one cell
 * instead of a whole tab.
 */
async function labFindXlsxImageAtCell(buffer,sheetTitle,targetRow,targetCol){

  if(typeof JSZip === "undefined"){
    throw new Error("JSZip didn't load.");
  }

  const zip =
    await JSZip.loadAsync(buffer);

  const workbookFile =
    zip.file("xl/workbook.xml");

  const workbookRelsFile =
    zip.file("xl/_rels/workbook.xml.rels");

  if(!workbookFile || !workbookRelsFile){
    return null;
  }

  const wbDoc =
    new DOMParser().parseFromString(
      await workbookFile.async("string"),
      "application/xml"
    );

  const relsDoc =
    new DOMParser().parseFromString(
      await workbookRelsFile.async("string"),
      "application/xml"
    );

  const relTargets = {};

  relsDoc
    .querySelectorAll("Relationship")
    .forEach(r => {
      relTargets[r.getAttribute("Id")] = r.getAttribute("Target");
    });

  let sheetPath = null;

  wbDoc
    .querySelectorAll("sheet")
    .forEach(s => {

      if(s.getAttribute("name") === sheetTitle){

        const rid =
          s.getAttributeNS(OOXML_REL_NS,"id") ||
          s.getAttribute("r:id");

        const target =
          relTargets[rid];

        if(target){
          sheetPath = "xl/" + target.replace(/^\/?/,"");
        }
      }
    });

  if(!sheetPath){
    return null;
  }

  const anchors =
    await readSheetImageAnchors(zip,sheetPath,"__lab__");

  const hit =
    anchors.find(a => a.row === targetRow && a.col === targetCol);

  return hit ? hit.imageUrl : null;
}
