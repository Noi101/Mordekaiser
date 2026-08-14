/* ============================================================
   GOOGLE SHEETS CONFIG
   ============================================================ */

const API_KEY =
  "AIzaSyCgxLgJKYU0U4jJ0m_2h_T3S-xO56Um8qE";

const SPREADSHEET_ID =
  "1BVH44IPxZNBQshDf90XZRJbfKvavVCUIIo_boyBVF8o";

const TAB_GIDS = {
  home:        848078357,
  matchup:     1046162980,
  itemization: 1054535084,
  runes:       779082056,
  altsetups:   2121273139,
  content:     225968364
};


/*
 * Paste the /exec URL from your CellImageExport.gs Web App
 * deployment here (see that file for setup steps). When this is
 * set, in-cell pasted images (item icons on Matchup, rune icons
 * on Runes) are read live from the actual sheet on every load —
 * no .xlsx export/upload needed at all, and nothing can ever go
 * stale since it's fetched fresh each time. Leave it blank to
 * skip straight to the .xlsx-based fallbacks (a file committed
 * to the repo, or a manual upload).
 */
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwuAhdgtbgEK0jfTQbvr3F_oxxv_gtev0Ly0p52D6WovhGMDBcUepX6ny4UEx3u1XoF/exec";


/* ============================================================
   DATA DRAGON CONFIG

   NOTE ON VERSIONING: /api/versions.json always returns every
   patch version with the newest one first, so versions[0] below
   is always the latest patch. Every image URL built from
   ddragonVersion (champion square icons, item icons) is
   therefore already current. Champion "loading" art and rune
   icons live at unversioned CDN paths by design (Riot doesn't
   version those folders) and always reflect the latest assets
   Riot has published, so nothing extra is needed there either.
   ============================================================ */

const DDRAGON_BASE =
  "https://ddragon.leagueoflegends.com";

let ddragonVersion = null;
let ddragonChampions = {};
let ddragonItems = {};
let ddragonRunes = {};


/* ============================================================
   DOM
   ============================================================ */

const statusEl =
  document.getElementById("status");

const syncLabel =
  document.getElementById("sync-label");

const tabnav =
  document.getElementById("tabnav");

const navToggle =
  document.getElementById("nav-toggle");

const tabsInner =
  document.getElementById("tabs-inner");

const searchWrap =
  document.getElementById("search-wrap");

const searchEl =
  document.getElementById("search");

const overlay =
  document.getElementById("overlay");

const sheetContent =
  document.getElementById("sheet-content");

const matchupGrid =
  document.getElementById("matchup-grid");

const nightfallVeil =
  document.getElementById("nightfall-veil");


let champions = [];
let activeTab = "matchup";

/*
 * titleByKey: our tab keys (matchup, runes, etc.) mapped to the
 * real sheet tab titles — set inside fetchAll(), reused by the
 * xlsx importer to find the matching sheet inside an uploaded
 * .xlsx export.
 *
 * liveRowsByKey: the last live rowsByKey fetched from the API —
 * kept around so we can re-resolve "what name is at this row
 * right now" whenever we need to check an xlsx image for
 * staleness, without re-fetching.
 *
 * xlsxSnapshot / xlsxNameIndex: the parsed contents of a
 * manually-uploaded .xlsx export. xlsxSnapshot holds raw
 * {row,col,imageUrl,nameAtUpload} entries per tab; xlsxNameIndex
 * is that same data indexed by lowercased name for fast lookup.
 * Session-only — cleared on page reload, since re-uploading is
 * how you refresh it anyway.
 *
 * lastSyncWarnings: non-fatal problems found on the last xlsx
 * upload (e.g. a tab that had no embedded images, or a sheet
 * name that didn't match) — this is what used to fail silently.
 * A failed upload throws and shows an error; a "succeeded but
 * found nothing useful" upload now shows these warnings instead
 * of just quietly doing nothing.
 */
let titleByKey = {};
let liveRowsByKey = {};
let xlsxSnapshot = {};
let xlsxNameIndex = {};
let lastSyncWarnings = [];
let xlsxSnapshotSource = null;
let xlsxSnapshotChecking = true;


/* ============================================================
   HELPERS
   ============================================================ */

function escapeHtml(s){
  return String(s)
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;")
    .replace(/'/g,"&#039;");
}


function normalizeName(name){
  return String(name || "")
    .toLowerCase()
    .replace(/[’']/g,"")
    .replace(/&/g,"and")
    .replace(/[^a-z0-9]/g,"");
}


/*
 * Plain Levenshtein edit distance, used only as a last-resort
 * fuzzy fallback when looking up champion art (see
 * findChampionData below) so a small typo in the spreadsheet's
 * "C:" name cell — "Ganglank" instead of "Gangplank" — doesn't
 * just silently fail to find a portrait. Generic on purpose:
 * this isn't a hardcoded fix for one champion, it protects
 * against the same kind of typo for any champion, now or later.
 */
function levenshtein(a,b){

  const m = a.length;
  const n = b.length;

  const dp =
    Array.from({length:m + 1},() => new Array(n + 1).fill(0));

  for(let i = 0; i <= m; i++){ dp[i][0] = i; }
  for(let j = 0; j <= n; j++){ dp[0][j] = j; }

  for(let i = 1; i <= m; i++){
    for(let j = 1; j <= n; j++){

      dp[i][j] =
        a[i - 1] === b[j - 1]
          ? dp[i - 1][j - 1]
          : 1 + Math.min(
              dp[i - 1][j],
              dp[i][j - 1],
              dp[i - 1][j - 1]
            );
    }
  }

  return dp[m][n];
}


/*
 * cellText: gives you the *displayable text* of a cell.
 * - Plain values pass through.
 * - =HYPERLINK("url","label") resolves to its label.
 * - =IMAGE("url") has no text label, so this returns "" —
 *   use extractImageUrl() below to get the picture itself.
 */
function cellText(cell){
  if(cell === null || cell === undefined){
    return "";
  }

  const s = String(cell);

  if(s.startsWith("=")){
    const link = extractHyperlink(s);

    if(link){
      return link.label;
    }

    return "";
  }

  return s.trim();
}


function extractHyperlink(cell){
  if(!cell){
    return null;
  }

  const m = String(cell).match(
    /=HYPERLINK\(\s*"([^"]+)"\s*(?:,\s*"([^"]*)")?\s*\)/i
  );

  if(!m){
    return null;
  }

  return {
    url:m[1],
    label:m[2] || m[1]
  };
}


/*
 * Cells where a rune/item picture was inserted with Sheets'
 * "insert image in cell" feature come back (with
 * valueRenderOption=FORMULA) as =IMAGE("https://..."). This
 * pulls that URL out directly so we show the exact picture the
 * spreadsheet author chose, instead of guessing from text.
 */
function extractImageUrl(cell){
  if(!cell){
    return null;
  }

  const s = String(cell);

  if(!s.startsWith("=")){
    return null;
  }

  const m = s.match(
    /=IMAGE\(\s*"([^"]+)"/i
  );

  return m ? m[1] : null;
}


function diffClass(n){
  return n === null ? "" : "diff-" + n;
}


/*
 * cellLinkHtml: like cellText, but when the cell is a
 * =HYPERLINK("url","label") formula, renders an actual
 * clickable <a> instead of flattening it down to plain label
 * text. Used anywhere we want links to stay clickable (Alt
 * Setups, Content/Creators).
 */
function cellLinkHtml(cell){

  const link =
    extractHyperlink(cell);

  if(link){
    return (
      `<a href="${escapeHtml(link.url)}" target="_blank" rel="noopener">` +
      `${escapeHtml(link.label)}</a>`
    );
  }

  return escapeHtml(cellText(cell));
}


/*
 * parseConditionalBlocks: turns a free-text cell like

     Standard: Comet -> ... -> Bone Plating
     If he builds full tank, swap Comet for Grasp.
     If he's ahead, Conqueror is fine instead.

 * into a lead paragraph plus a list of labeled condition
 * blocks. This isn't specific to any one champion — any
 * matchup's rune build or itemization text can use one or more
 * "If ..." lines and it'll render as its own card instead of
 * getting mashed into one paragraph. Right now Aatrox might be
 * the only one written that way in the sheet, but the parser
 * doesn't care who it is — it just looks for the pattern, so it
 * keeps working as more matchups get conditional notes added.
 */
function parseConditionalBlocks(text){

  if(!text){
    return {lead:"",conditions:[]};
  }

  const lines =
    text
      .split(/\n+/)
      .map(l => l.trim())
      .filter(Boolean);

  const lead = [];
  const conditions = [];
  let current = null;

  lines.forEach(line => {

    if(/^if\b/i.test(line)){

      if(current){
        conditions.push(current);
      }

      const m =
        line.match(/^(if\s[^,:]+)[,:]?\s*(.*)$/i);

      current = {
        label: m ? m[1].trim() : line,
        text: m && m[2] ? m[2].trim() : ""
      };

    }else if(current){

      current.text =
        current.text
          ? current.text + " " + line
          : line;

    }else{
      lead.push(line);
    }
  });

  if(current){
    conditions.push(current);
  }

  return {
    lead:lead.join(" "),
    conditions:conditions
  };
}


/*
 * renderConditionalHtml: renders parseConditionalBlocks() output
 * — plain paragraph for the lead text, then a styled card per
 * "If ..." condition. Falls back to the old plain <pre> block
 * when there's nothing conditional in the text at all, so
 * ordinary matchups look exactly like they did before.
 */
function renderConditionalHtml(text,className){

  const parsed =
    parseConditionalBlocks(text);

  if(!parsed.conditions.length){

    return text
      ? `<pre class="${className}">${escapeHtml(text)}</pre>`
      : "";
  }

  let html = "";

  if(parsed.lead){
    html +=
      `<pre class="${className}">${escapeHtml(parsed.lead)}</pre>`;
  }

  html +=
    `
      <div class="conditional-list">
        ${
          parsed.conditions.map(
            c => `
              <div class="conditional-card">
                <div class="conditional-label">
                  ${escapeHtml(c.label)}
                </div>
                ${
                  c.text
                    ? `<div class="conditional-text">${escapeHtml(c.text)}</div>`
                    : ""
                }
              </div>
            `
          ).join("")
        }
      </div>
    `;

  return html;
}


/* ============================================================
   DATA DRAGON
   ============================================================ */

async function loadDataDragon(){

  const versionsResponse =
    await fetch(
      DDRAGON_BASE + "/api/versions.json",
      {cache:"no-store"}
    );

  if(!versionsResponse.ok){
    throw new Error(
      "Could not load Riot Data Dragon versions."
    );
  }

  const versions =
    await versionsResponse.json();

  if(!versions.length){
    throw new Error(
      "Riot Data Dragon returned no versions."
    );
  }

  // versions[0] is always the newest patch — see note above.
  ddragonVersion = versions[0];

  /*
   * Champion data.
   */
  const championResponse =
    await fetch(
      `${DDRAGON_BASE}/cdn/${ddragonVersion}/data/en_US/champion.json`,
      {cache:"no-store"}
    );

  if(championResponse.ok){

    const championJson =
      await championResponse.json();

    ddragonChampions =
      championJson.data || {};
  }

  /*
   * Item data.
   */
  const itemResponse =
    await fetch(
      `${DDRAGON_BASE}/cdn/${ddragonVersion}/data/en_US/item.json`,
      {cache:"no-store"}
    );

  if(itemResponse.ok){

    const itemJson =
      await itemResponse.json();

    ddragonItems =
      itemJson.data || {};
  }

  /*
   * Rune data.
   */
  const runeResponse =
    await fetch(
      `${DDRAGON_BASE}/cdn/${ddragonVersion}/data/en_US/runesReforged.json`,
      {cache:"no-store"}
    );

  if(runeResponse.ok){

    const runeJson =
      await runeResponse.json();

    ddragonRunes = {};

    runeJson.forEach(style => {

      if(style.name){
        ddragonRunes[
          normalizeName(style.name)
        ] = {
          name:style.name,
          icon:`${DDRAGON_BASE}/cdn/img/${style.icon}`
        };
      }

      (style.slots || []).forEach(slot => {

        (slot.runes || []).forEach(rune => {

          if(rune.name){

            ddragonRunes[
              normalizeName(rune.name)
            ] = {
              name:rune.name,
              icon:`${DDRAGON_BASE}/cdn/img/${rune.icon}`
            };
          }

        });

      });

    });
  }
}


/* ============================================================
   CHAMPION IMAGE LOOKUP
   ============================================================ */

function findChampionData(name){

  const wanted =
    normalizeName(name);

  /*
   * First try exact normalized name.
   */
  for(const key in ddragonChampions){

    const champ =
      ddragonChampions[key];

    if(
      normalizeName(champ.name) === wanted
    ){
      return champ;
    }
  }

  /*
   * Handle common names containing punctuation.
   */
  const aliases = {
    "wukong":"MonkeyKing",
    "chogath":"Chogath",
    "drmundo":"DrMundo",
    "jarvaniv":"JarvanIV",
    "leesin":"LeeSin",
    "masteryi":"MasterYi",
    "missfortune":"MissFortune",
    "reksai":"RekSai",
    "renataglasc":"Renata",
    "tahmkench":"TahmKench",
    "twistedfate":"TwistedFate",
    "xinzhao":"XinZhao"
  };

  if(aliases[wanted] && ddragonChampions[aliases[wanted]]){
    return ddragonChampions[aliases[wanted]];
  }

  /*
   * Fuzzy fallback: catches small typos in the spreadsheet's
   * "C:" name cell (e.g. "Ganglank" vs "Gangplank") instead of
   * just showing no portrait. Only accepts close matches — edit
   * distance capped at ~20% of the candidate's length, and the
   * lengths have to be close to begin with — so it can't
   * accidentally snap to an unrelated champion. If your
   * spreadsheet name is spelled correctly and this still
   * doesn't resolve, click the 🔗 debug link on that champion's
   * detail card — it shows exactly what name we tried to match.
   */
  let fuzzyBest = null;
  let fuzzyDist = Infinity;

  for(const key in ddragonChampions){

    const champ = ddragonChampions[key];
    const candidate = normalizeName(champ.name);

    if(Math.abs(candidate.length - wanted.length) > 2){
      continue;
    }

    const dist = levenshtein(wanted,candidate);
    const limit = Math.max(1,Math.floor(candidate.length * 0.2));

    if(dist <= limit && dist < fuzzyDist){
      fuzzyDist = dist;
      fuzzyBest = champ;
    }
  }

  return fuzzyBest;
}


/*
 * Uses the "loading" splash art (308x560, always the latest
 * skin Riot has published for that champion) instead of the
 * 120x120 square icon. The card/overlay CSS boxes are sized to
 * this exact ratio, so the whole portrait shows — nothing gets
 * stretched into a square-in-a-tall-box crop anymore.
 */
function championImage(name){

  const champ =
    findChampionData(name);

  if(!champ){
    return null;
  }

  return (
    `${DDRAGON_BASE}/cdn/img/champion/loading/${champ.id}_0.jpg`
  );
}


/*
 * The classic 120x120 square icon — used for the Whack-a-Vayne
 * game holes, where a tall splash crop wouldn't read at that
 * size.
 */
function championSquareImage(name){

  const champ =
    findChampionData(name);

  if(!champ || !ddragonVersion){
    return null;
  }

  return (
    `${DDRAGON_BASE}/cdn/${ddragonVersion}/img/champion/${champ.id}.png`
  );
}


/* ============================================================
   ITEM IMAGE LOOKUP (text-name fallback, used only when a
   cell has no inserted =IMAGE() picture to read directly)
   ============================================================ */

function findItemData(text){

  const wanted =
    normalizeName(text);

  if(!wanted){
    return null;
  }

  /*
   * Exact item-name matching.
   */
  for(const id in ddragonItems){

    const item =
      ddragonItems[id];

    if(!item.name){
      continue;
    }

    if(
      normalizeName(item.name) === wanted
    ){
      return {
        id:id,
        name:item.name,
        icon:
          `${DDRAGON_BASE}/cdn/${ddragonVersion}` +
          `/img/item/${id}.png`
      };
    }
  }

  /*
   * If the spreadsheet cell contains a longer build
   * description, search for an item name inside it. We keep
   * the LONGEST matching item name (not just the first one
   * object-key order happens to hit) so e.g. "Death's Dance"
   * isn't shadowed by a shorter unrelated match.
   */
  let best = null;

  for(const id in ddragonItems){

    const item =
      ddragonItems[id];

    if(!item.name){
      continue;
    }

    const itemName =
      normalizeName(item.name);

    if(
      itemName.length >= 4 &&
      wanted.includes(itemName) &&
      (!best || itemName.length > normalizeName(best.name).length)
    ){
      best = {
        id:id,
        name:item.name,
        icon:
          `${DDRAGON_BASE}/cdn/${ddragonVersion}` +
          `/img/item/${id}.png`
      };
    }
  }

  return best;
}


/* ============================================================
   RUNE IMAGE LOOKUP (text-name fallback)
   ============================================================ */

function findRuneData(text){

  const wanted =
    normalizeName(text);

  if(!wanted){
    return null;
  }

  if(ddragonRunes[wanted]){
    return ddragonRunes[wanted];
  }

  let best = null;

  for(const key in ddragonRunes){

    if(
      key.length >= 4 &&
      wanted.includes(key) &&
      (!best || key.length > normalizeName(best.name).length)
    ){
      best = ddragonRunes[key];
    }
  }

  return best;
}


/*
 * classifyBuildText: counts how many official Riot rune names
 * and how many official Riot item names appear inside a block
 * of free text. Used by parseChampions to decide whether a cell
 * is describing a rune build or an item build — see the comment
 * at its call site for why this replaces the old "does it
 * contain an arrow" guess.
 */
function classifyBuildText(text){

  const norm =
    normalizeName(text);

  let itemScore = 0;
  let runeScore = 0;

  for(const id in ddragonItems){

    const nm = ddragonItems[id].name;

    if(!nm){
      continue;
    }

    const key = normalizeName(nm);

    if(key.length >= 4 && norm.includes(key)){
      itemScore++;
    }
  }

  for(const key in ddragonRunes){

    if(key.length >= 4 && norm.includes(key)){
      runeScore++;
    }
  }

  return {itemScore,runeScore};
}


/* ============================================================
   GOOGLE SHEETS FETCHING
   ============================================================ */

async function fetchAll(){

  if(
    !API_KEY ||
    API_KEY.startsWith("IDE_")
  ){
    showConfigWarning(
      "A Google Sheets API key is missing."
    );

    return;
  }

  try{

    syncLabel.textContent =
      "Loading Riot assets…";

    await loadDataDragon();

    syncLabel.textContent =
      "Loading spreadsheet…";

    const metaRes =
      await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/` +
        `${SPREADSHEET_ID}?key=${API_KEY}` +
        `&fields=sheets.properties`
      );

    const metaJson =
      await metaRes.json();

    if(metaJson.error){
      throw new Error(
        metaJson.error.message
      );
    }

    titleByKey = {};

    for(const key in TAB_GIDS){

      const props =
        metaJson.sheets.find(
          s =>
            s.properties.sheetId ===
            TAB_GIDS[key]
        );

      if(!props){

        throw new Error(
          `Could not find sheet for ${key} ` +
          `(gid ${TAB_GIDS[key]})`
        );
      }

      titleByKey[key] =
        props.properties.title;
    }


    const rangeParams =
      Object.values(titleByKey)
        .map(
          t =>
            `ranges=${encodeURIComponent(
              "'" + t + "'!A1:AA3000"
            )}`
        )
        .join("&");


    const batchRes =
      await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/` +
        `${SPREADSHEET_ID}/values:batchGet?` +
        `${rangeParams}` +
        `&valueRenderOption=FORMULA` +
        `&key=${API_KEY}`
      );


    const batchJson =
      await batchRes.json();

    if(batchJson.error){
      throw new Error(
        batchJson.error.message
      );
    }


    const rowsByKey = {};
    const keys =
      Object.keys(titleByKey);


    batchJson.valueRanges.forEach(
      (vr,i) => {
        rowsByKey[keys[i]] =
          vr.values || [];
      }
    );


    liveRowsByKey = rowsByKey;


    /*
     * Image snapshot resolution (live Apps Script -> repo file
     * -> browser storage) now runs in the BACKGROUND — not
     * awaited here. It used to block the entire page behind it:
     * if the Apps Script endpoint was slow to respond (cold
     * starts, a big sheet to scan), the whole site sat frozen on
     * "Loading spreadsheet data…" until it finally resolved or
     * timed out. None of the matchup text, ratings, or guides
     * depend on this — only a handful of pasted-in icons do — so
     * there's no reason to make every visitor wait on it.
     * resolveImageSnapshotInBackground() patches the icons in
     * once it finishes, whenever that ends up being.
     */
    resolveImageSnapshotInBackground();


    champions =
      parseChampions(
        rowsByKey.matchup
      );


    renderMatchupGrid(
      champions
    );


    const itemizationSections =
      parseTextSections(
        rowsByKey.itemization,
        "item"
      );

    renderIconSections(
      "tab-itemization",
      itemizationSections
    );

    itemizationFlat =
      itemizationSections
        .flatMap(s => s.items)
        .filter(it => it.icon && it.label && it.text);


    renderIconSections(
      "tab-runes",
      parseTextSections(
        rowsByKey.runes,
        "rune",
        "runes"
      )
    );


    renderAltSetups(
      rowsByKey.altsetups
    );


    renderContent(
      rowsByKey.content
    );


    renderHome(
      rowsByKey.home
    );


    renderGamesPage();

    renderSyncPage();


    syncLabel.textContent =
      `Live sync · ${champions.length} matchups · ` +
      `Data Dragon ${ddragonVersion}`;


    statusEl.style.display =
      "none";

    tabnav.style.display =
      "flex";

    switchTab("matchup");


  }catch(err){

    console.error(err);

    statusEl.style.display =
      "block";

    statusEl.innerHTML =
      `<div class="config-warning">
        <strong>Synchronization error</strong>
        <br><br>
        ${escapeHtml(err.message)}
        <br><br>
        Check your Google Sheets API key,
        spreadsheet permissions, and Data Dragon connection.
      </div>`;
  }
}


function showConfigWarning(msg){

  statusEl.innerHTML =
    `<div class="config-warning">
      <strong>Configuration problem</strong>
      <br><br>
      ${escapeHtml(msg)}
    </div>`;
}


/* ============================================================
   MATCHUP PARSER
   ============================================================ */

function parseChampions(rows){

  const anchors = [];

  rows.forEach(
    (row,idx) => {

      const a =
        cellText(row[0]);

      if(
        a &&
        a.toLowerCase().startsWith("c:")
      ){

        anchors.push({
          idx:idx,
          name:a.slice(2).trim()
        });
      }
    }
  );


  const result = [];


  for(
    let i = 0;
    i < anchors.length;
    i++
  ){

    const start =
      anchors[i].idx;

    const end =
      i + 1 < anchors.length
        ? anchors[i + 1].idx
        : rows.length;


    const block =
      rows.slice(start,end);


    const allCellsFlat = [];


    block.forEach(
      row => {

        (row || []).forEach(
          cell => {

            const text =
              cellText(cell);

            if(text){
              allCellsFlat.push(text);
            }

          }
        );

      }
    );


    /*
     * Get the champion image from Data Dragon,
     * NOT from Google Sheets. If the "C:" name has a typo,
     * findChampionData's fuzzy fallback will usually still
     * resolve it — but the name shown on the card always comes
     * straight from this cell, exactly as typed in the sheet.
     */
    const portrait =
      championImage(
        anchors[i].name
      );


    const joined =
      allCellsFlat.join(" | ");


    const diff = {};


    [
      "Early",
      "Mid",
      "Late",
      "Overall"
    ].forEach(
      k => {

        const m =
          joined.match(
            new RegExp(
              k +
              "\\s*:?\\s*(\\d)\\s*/\\s*5",
              "i"
            )
          );

        diff[
          k.toLowerCase()
        ] =
          m
            ? parseInt(m[1],10)
            : null;
      }
    );


    let gameplay = "";

    allCellsFlat.forEach(
      t => {

        if(
          t.length >
          gameplay.length
        ){
          gameplay = t;
        }

      }
    );


    const remaining =
      allCellsFlat
        .filter(
          t => t !== gameplay
        )
        .sort(
          (a,b) =>
            b.length - a.length
        );


    /*
     * Both the rune-build text and the itemization text can
     * contain "->" and parentheses (e.g. "Bramble Vest Rush ->
     * Boots" reads exactly like a rune sequence would), so
     * picking whichever cell "has an arrow" isn't reliable — it
     * picked the wrong cell for some champions (Aatrox showed
     * item names under "Rune Build"). Instead, score each
     * remaining cell against Riot's actual rune and item name
     * lists (classifyBuildText below) and assign it to whichever
     * side it matches more strongly. This is the "official
     * standard" check: rune names only come from Riot's rune
     * tree, item names only come from Riot's item list, so a
     * cell full of item names will always out-score as
     * itemization, and vice versa, regardless of punctuation.
     */
    const scored =
      remaining.map(t => ({
        text:t,
        ...classifyBuildText(t)
      }));

    let runeBuild = "";
    let itemization = "";

    const runeCandidate =
      scored
        .filter(s => s.runeScore > 0 && s.runeScore >= s.itemScore)
        .sort((a,b) => b.runeScore - a.runeScore)[0];

    if(runeCandidate){
      runeBuild = runeCandidate.text;
    }

    const itemCandidate =
      scored
        .filter(
          s =>
            s.text !== runeBuild &&
            s.itemScore > 0 &&
            s.itemScore >= s.runeScore
        )
        .sort((a,b) => b.itemScore - a.itemScore)[0];

    if(itemCandidate){
      itemization = itemCandidate.text;
    }

    /*
     * Fallback to the old pattern-based guess only if the vocab
     * check found nothing at all — e.g. a very short note with
     * no recognizable rune or item names in it yet.
     */
    if(!runeBuild){
      runeBuild =
        remaining.find(t => t.includes("->")) || "";
    }

    if(!itemization){
      itemization =
        remaining.find(
          t =>
            t !== runeBuild &&
            /\(.*\)/.test(t) &&
            t.length > 20
        ) || "";
    }


    result.push({
      name:anchors[i].name,
      portrait:portrait,
      difficulty:diff,
      gameplay:gameplay,
      runeBuild:runeBuild,
      itemization:itemization
    });
  }


  return result.filter(
    c => c.name
  );
}


/* ============================================================
   MATCHUP RENDER
   ============================================================ */

function renderMatchupGrid(list){

  matchupGrid.innerHTML =
    list.map(
      (c,i) => {

        const image =
          c.portrait
            ? `<img
                class="portrait"
                src="${c.portrait}"
                alt="${escapeHtml(c.name)}"
                loading="lazy"
              >`
            : `<div
                style="
                  width:100%;
                  height:100%;
                  display:flex;
                  align-items:center;
                  justify-content:center;
                  color:var(--ink-dim);
                  font-size:12px;
                  padding:15px;
                  text-align:center;
                "
              >
                ${escapeHtml(c.name)}
              </div>`;


        return `
          <div
            class="card"
            data-idx="${i}"
          >
            <div class="portrait-wrap">

              ${image}

              ${
                c.difficulty.overall !== null
                  ? `
                    <div
                      class="diff-badge ${diffClass(
                        c.difficulty.overall
                      )}"
                    >
                      ${c.difficulty.overall}/5
                    </div>
                  `
                  : ""
              }

              <div class="name">
                ${escapeHtml(c.name)}
              </div>

            </div>
          </div>
        `;
      }
    )
    .join("");


  matchupGrid
    .querySelectorAll(".card")
    .forEach(
      card => {

        card.addEventListener(
          "click",
          () => {

            openDetail(
              list[
                parseInt(
                  card.dataset.idx
                )
              ]
            );

          }
        );

      }
    );
}


/* ============================================================
   DIFFICULTY BAR
   ============================================================ */

function forgeBar(label,value){

  let cls =
    "filled";

  if(value >= 4){
    cls =
      "filled high";
  }else if(value === 3){
    cls =
      "filled mid";
  }


  const notches =
    Array.from(
      {length:5},
      (_,i) =>
        `<div class="notch ${
          i < value ? cls : ""
        }"></div>`
    ).join("");


  return `
    <div>
      <div class="forge-label">
        ${label}
      </div>

      <div class="forge-bar">
        ${notches}
      </div>
    </div>
  `;
}


/* ============================================================
   MATCHUP DETAIL
   ============================================================ */

function openDetail(c){

  const d =
    c.difficulty;

  const xlsxItems =
    lookupXlsxImages("matchup",c.name);


  sheetContent.innerHTML =
    `
      <div class="sheet-head">

        <button
          class="close-btn"
          id="close-btn"
        >
          ✕
        </button>

        ${
          c.portrait
            ? `
              <img
                class="portrait"
                src="${c.portrait}"
                alt="${escapeHtml(c.name)}"
              >
            `
            : ""
        }

        <div class="sheet-head-info">

          <h2>
            ${escapeHtml(c.name)}
            <button
              class="debug-link-toggle"
              id="debug-link-toggle"
              title="Show the image URL / lookup status for this champion"
            >
              🔗
            </button>
          </h2>

          <div
            class="debug-link-row"
            id="debug-link-row"
            style="display:none;"
          >
            ${
              c.portrait
                ? `
                  <input
                    type="text"
                    readonly
                    value="${escapeHtml(c.portrait)}"
                    id="debug-link-input"
                  >
                  <button id="debug-link-copy">Copy</button>
                `
                : `
                  <span class="debug-link-missing">
                    No Data Dragon match for "${escapeHtml(c.name)}" —
                    double check the spelling in the spreadsheet's "C:" name
                    cell against Data Dragon's champion list. Small typos
                    are usually forgiven automatically; this means the
                    name is too far off for the fuzzy match to trust.
                  </span>
                `
            }
          </div>

          <div class="vs">
            Mordekaiser vs. ${escapeHtml(c.name)}
          </div>

          <div class="forge">

            ${
              d.early !== null
                ? forgeBar(
                    "Early",
                    d.early
                  )
                : ""
            }

            ${
              d.mid !== null
                ? forgeBar(
                    "Mid",
                    d.mid
                  )
                : ""
            }

            ${
              d.late !== null
                ? forgeBar(
                    "Late",
                    d.late
                  )
                : ""
            }

            ${
              d.overall !== null
                ? forgeBar(
                    "Overall",
                    d.overall
                  )
                : ""
            }

          </div>

        </div>
      </div>


      <div class="sheet-body">

        ${
          c.runeBuild
            ? `
              <div class="section-title">
                Rune Build
              </div>

              ${renderConditionalHtml(c.runeBuild,"rune-text")}
            `
            : ""
        }


        ${
          c.itemization || xlsxItems.length
            ? `
              <div class="section-title">
                Itemization
              </div>

              ${
                xlsxItems.length
                  ? `
                    <div class="xlsx-item-row">
                      ${
                        xlsxItems.map(
                          entry => `
                            <div class="xlsx-item-chip">
                              <img
                                src="${entry.imageUrl}"
                                alt=""
                                loading="lazy"
                              >
                              ${
                                isXlsxEntryStale("matchup",entry)
                                  ? `<span class="stale-badge" title="This row's content has changed since your last .xlsx upload">!</span>`
                                  : ""
                              }
                            </div>
                          `
                        ).join("")
                      }
                    </div>
                  `
                  : ""
              }

              ${
                c.itemization
                  ? renderConditionalHtml(c.itemization,"item-text")
                  : ""
              }
            `
            : ""
        }


        ${
          c.gameplay
            ? `
              <div class="section-title">
                Gameplay Guide
              </div>

              <pre class="guide-text">${
                escapeHtml(
                  c.gameplay
                )
              }</pre>
            `
            : ""
        }

      </div>
    `;


  overlay.classList.add(
    "open"
  );


  document
    .getElementById("close-btn")
    .addEventListener(
      "click",
      closeDetail
    );


  const debugToggle =
    document.getElementById("debug-link-toggle");

  const debugRow =
    document.getElementById("debug-link-row");

  if(debugToggle && debugRow){

    debugToggle.addEventListener(
      "click",
      () => {
        debugRow.style.display =
          debugRow.style.display === "none"
            ? "flex"
            : "none";
      }
    );
  }


  const debugCopy =
    document.getElementById("debug-link-copy");

  if(debugCopy){

    debugCopy.addEventListener(
      "click",
      () => {

        const input =
          document.getElementById("debug-link-input");

        input.select();

        if(navigator.clipboard){
          navigator.clipboard.writeText(input.value);
        }

        debugCopy.textContent = "Copied!";

        setTimeout(() => {
          debugCopy.textContent = "Copy";
        },1200);

      }
    );
  }
}


function closeDetail(){
  overlay.classList.remove(
    "open"
  );
}


overlay.addEventListener(
  "click",
  e => {

    if(
      e.target === overlay
    ){
      closeDetail();
    }

  }
);


/* ============================================================
   TEXT / ICON SECTIONS (Rune Guide + Itemization tabs)

   The sheet repeats column pairs across a row, exactly like:
   [Rune][Explanation][Rune][Explanation][Rune][Explanation]
   (or [Item][Explanation] on the itemization tab). This parser
   finds the header row that declares those pairs, then reads
   every data row strictly by that pairing — column i is the
   icon, column i+1 is ALWAYS its explanation. That's what fixes
   "items on the same line getting treated as different items":
   they were being scanned independently before instead of
   paired to the column layout the sheet actually uses.
   ============================================================ */

/*
 * The Runes tab labels its pairs plainly: "Rune" / "Explanation".
 * The Itemization tab instead labels each pair by build order —
 * "First Items", "Second Items", "Third Items", "Fourth+ Items"
 * (see the real sheet header row) — with no separate
 * "Explanation" header on the second column at all. Both are
 * valid ways of declaring "this column is a picture, the next
 * one is its write-up", so NAME_HEADER_RE matches either style
 * and we no longer require the second column to say
 * "Explanation" — we just take whatever's immediately to its
 * right.
 */
const NAME_HEADER_RE =
  /^(rune|item|items|first\s*items?|second\s*items?|third\s*items?|fourth\+?\s*items?|\d+(st|nd|rd|th)\s*items?)$/i;


function detectPairColumns(row){

  const cells =
    (row || []).map(cellText);

  const pairs = [];

  for(let i = 0; i < cells.length - 1; i++){

    if(NAME_HEADER_RE.test(cells[i])){
      pairs.push([i, i + 1]);
    }
  }

  return pairs;
}


function parseTextSections(rows,iconType,xlsxTabKey){

  const sections = [];

  let current = null;
  let pairCols = [];


  function flush(){

    if(
      current &&
      (current.items.length || current.notes.length)
    ){
      sections.push(current);
    }
  }


  function startSection(title){

    flush();

    current = {
      title:title,
      items:[],
      notes:[]
    };
  }


  (rows || []).forEach(row => {

    const cells =
      (row || []).map(cellText);

    const nonEmpty =
      cells.filter(Boolean);

    if(!nonEmpty.length){
      return;
    }

    /*
     * Header row declaring the Icon/Explanation column pairs
     * for the rows that follow (can repeat per section, as in
     * the screenshot where each new block re-declares it).
     */
    const detected =
      detectPairColumns(row);

    if(detected.length){

      pairCols = detected;

      if(!current){
        startSection(null);
      }

      return;
    }

    /*
     * A lone short cell is a section heading
     * (e.g. a category name above a block of runes/items).
     */
    if(
      nonEmpty.length === 1 &&
      nonEmpty[0].length < 70
    ){
      startSection(nonEmpty[0]);
      return;
    }

    if(!current){
      startSection(null);
    }

    if(!pairCols.length){
      // No header seen yet for this block — fall back to
      // treating the row as freeform notes rather than
      // guessing at column meaning.
      current.notes.push(cells.filter(Boolean).join(" — "));
      return;
    }

    pairCols.forEach(([nameCol,explCol]) => {

      const rawName =
        (row || [])[nameCol];

      const nameText =
        cellText(rawName);

      const explText =
        cellText((row || [])[explCol]);

      const directImage =
        extractImageUrl(rawName);

      let icon = directImage;
      let label = nameText;
      let stale = false;

      if(!icon && nameText){

        const found =
          iconType === "item"
            ? findItemData(nameText)
            : findRuneData(nameText);

        if(found){
          icon = found.icon;
          label = found.name;
        }
      }

      if(!icon && !nameText && explText){

        // Last resort: the name cell was blank (image the
        // Sheets API couldn't expose) — try to infer from the
        // explanation text itself.
        const found =
          iconType === "item"
            ? findItemData(explText)
            : findRuneData(explText);

        if(found){
          icon = found.icon;
          label = found.name;
        }
      }

      if(!icon && xlsxTabKey){

        // Last-last resort: an image was pasted directly into
        // the cell (not an =IMAGE() formula, not a name Data
        // Dragon recognizes) — only readable from an uploaded
        // .xlsx snapshot, matched by name.
        const entry =
          lookupXlsxImage(
            xlsxTabKey,
            label || explText
          );

        if(entry){
          icon = entry.imageUrl;
          stale = isXlsxEntryStale(xlsxTabKey,entry);
        }
      }

      if(icon || explText || label){

        current.items.push({
          icon:icon,
          label:label,
          text:explText,
          stale:stale
        });
      }

    });

  });


  flush();

  return sections;
}


function renderIconSections(containerId,sections){

  const el =
    document.getElementById(containerId);


  if(!sections.length){

    el.innerHTML =
      `<p style="
        color:var(--ink-dim);
        text-align:center;
      ">
        No displayable data on this tab.
      </p>`;

    return;
  }


  el.innerHTML =
    sections.map(
      sec => `

        <div class="icon-section">

          ${
            sec.title
              ? `
                <h3>
                  ${escapeHtml(
                    sec.title
                  )}
                </h3>
              `
              : ""
          }


          ${
            sec.items.length
              ? `
                <div class="icon-grid">

                  ${
                    sec.items.map(
                      it => `
                        <div class="icon-card">

                          ${
                            it.icon
                              ? `
                                <div class="icon-img-wrap">
                                  <img
                                    src="${it.icon}"
                                    alt=""
                                    loading="lazy"
                                  >
                                  ${
                                    it.stale
                                      ? `<span class="stale-badge" title="This row's content has changed since your last .xlsx upload">!</span>`
                                      : ""
                                  }
                                </div>
                              `
                              : `
                                <div class="icon-placeholder">
                                  ${
                                    it.label
                                      ? escapeHtml(it.label)
                                      : "?"
                                  }
                                </div>
                              `
                          }

                          <p>
                            ${
                              it.label
                                ? `<strong>${escapeHtml(it.label)}</strong>`
                                : ""
                            }
                            ${escapeHtml(
                              it.text ||
                              "(no description)"
                            )}
                          </p>

                        </div>
                      `
                    ).join("")
                  }

                </div>
              `
              : ""
          }


          ${
            sec.notes.length
              ? `
                <div class="icon-notes">
                  ${
                    sec.notes
                      .map(
                        n =>
                          escapeHtml(n)
                      )
                      .join(
                        "<br><br>"
                      )
                  }
                </div>
              `
              : ""
          }

        </div>
      `
    ).join("");
}


/* ============================================================
   ALTERNATE SETUPS
   ============================================================ */

function renderAltSetups(rows){

  const el =
    document.getElementById(
      "tab-altsetups"
    );


  const headerIdx =
    (rows || []).findIndex(
      r =>
        cellText(r[0]) ===
        "Build Concept"
    );


  if(headerIdx === -1){

    el.innerHTML =
      `<p style="
        color:var(--ink-dim);
        text-align:center;
      ">
        No data found.
      </p>`;

    return;
  }


  let html =
    `
      <div class="data-block">

        <table class="data-table">

          <thead>
            <tr>
              <th>Build</th>
              <th>Runes</th>
              <th>Items</th>
              <th>Example</th>
              <th>Notes</th>
            </tr>
          </thead>

          <tbody>
    `;


  for(
    let i = headerIdx + 1;
    i < rows.length;
    i++
  ){

    const row =
      rows[i] || [];


    const name =
      cellText(row[0]);


    if(!name){
      continue;
    }


    const runes =
      cellText(row[2]);

    const items =
      cellText(row[3]);

    const example =
      cellText(row[4]);

    const comments =
      cellText(row[6]);


    if(
      !runes &&
      !items &&
      !example &&
      !comments
    ){

      html +=
        `
          </tbody>
        </table>
      </div>

      <div class="data-block">

        <h3>
          ${escapeHtml(name)}
        </h3>

        <table class="data-table">

          <thead>
            <tr>
              <th>Build</th>
              <th>Runes</th>
              <th>Items</th>
              <th>Example</th>
              <th>Notes</th>
            </tr>
          </thead>

          <tbody>
        `;

      continue;
    }


    html +=
      `
        <tr>

          <td>
            <strong style="
              color:var(--rune-gold)
            ">
              ${cellLinkHtml(row[0])}
            </strong>
          </td>

          <td>
            ${cellLinkHtml(row[2])}
          </td>

          <td>
            ${cellLinkHtml(row[3])}
          </td>

          <td>
            ${cellLinkHtml(row[4])}
          </td>

          <td>
            ${cellLinkHtml(row[6])}
          </td>

        </tr>
      `;
  }


  html +=
    `
          </tbody>
        </table>

      </div>
    `;


  el.innerHTML =
    html;
}


/* ============================================================
   CONTENT / CREATORS
   ============================================================ */

function renderContent(rows){

  const el =
    document.getElementById(
      "tab-content"
    );


  let html = "";
  let i = 0;

  const R =
    rows || [];


  while(i < R.length){

    const row =
      R[i] || [];


    const cells =
      row.map(cellText);


    const nonEmptyCount =
      cells.filter(
        c => c
      ).length;


    if(
      nonEmptyCount === 1 &&
      cells[0] &&
      cells[0].length > 3 &&
      !/^(Region)$/.test(
        cells[0]
      )
    ){

      html +=
        `
          <div class="data-block">
            <h3>
              ${escapeHtml(
                cells[0]
              )}
            </h3>
        `;


      i++;


      const headerRow =
        (R[i] || [])
          .map(cellText);


      if(
        headerRow[0] ===
        "Region"
      ){

        html +=
          `
            <table class="data-table">
              <thead>
                <tr>
                  ${
                    headerRow
                      .filter(
                        (h,idx) =>
                          idx < 8
                      )
                      .map(
                        h =>
                          `<th>${escapeHtml(h)}</th>`
                      )
                      .join("")
                  }
                </tr>
              </thead>

              <tbody>
          `;


        i++;


        while(
          i < R.length
        ){

          const rawDr =
            R[i] || [];

          const dr =
            rawDr.map(cellText);


          if(
            dr.filter(
              c => c
            ).length <= 1
          ){
            break;
          }


          html +=
            `
              <tr>
                ${
                  rawDr
                    .slice(0,8)
                    .map(
                      c =>
                        `<td>${cellLinkHtml(c)}</td>`
                    )
                    .join("")
                }
              </tr>
            `;


          i++;
        }


        html +=
          `
              </tbody>
            </table>
          `;
      }


      html +=
        `
          </div>
        `;


      continue;
    }


    i++;
  }


  el.innerHTML =
    html ||
    `<p style="
      color:var(--ink-dim);
      text-align:center;
    ">
      No data found.
    </p>`;
}


/* ============================================================
   PATCH NOTES / HOME
   ============================================================ */

function renderHome(rows){

  const el =
    document.getElementById(
      "tab-home"
    );


  let html = "";


  (rows || []).forEach(
    row => {

      const texts =
        (row || [])
          .map(cellText)
          .filter(t => t);


      if(!texts.length){
        return;
      }


      html +=
        `
          <div class="paragraph-block">
            ${
              texts
                .map(
                  t =>
                    escapeHtml(t)
                )
                .join(" — ")
            }
          </div>
        `;
    }
  );


  el.innerHTML =
    html ||
    `<p style="
      color:var(--ink-dim);
      text-align:center;
    ">
      No data found.
    </p>`;
}


/* ============================================================
   XLSX SNAPSHOT ENGINE

   Some images (pasted directly into cells with Sheets' "insert
   image in cell" feature, not an =IMAGE() formula) can't be read
   through the Sheets API at all — the API just returns nothing
   for that cell. The only way to get the actual picture is to
   export the sheet as .xlsx (which bundles real image files)
   and read it locally.

   The flow: you export → .xlsx from Google Sheets, upload it on
   the Data Sync tab, and we dig through the file's internal XML
   to find every embedded image, which row/column it's anchored
   to, and — by cross-referencing that row against the live data
   we already fetched — what champion/rune/item it belongs to.

   Because that mapping is captured once at upload time, if the
   live sheet later changes what's sitting in that row, we can
   tell the image is now stale (isXlsxEntryStale) just by
   re-checking what name is at that row *now* vs. what it was
   *at upload time*. No stored image data is compared — only the
   row's identity — which is intentionally simple.

   NOTE: this is session-only. Nothing is persisted across page
   reloads (images are held as in-memory blob URLs), which is
   also why the tab always starts back at "no snapshot uploaded".

   WHERE THE UPLOADED FILE IS USED: it only backs the Matchup
   tab's itemization chips and the Runes tab icon grid — anywhere
   a cell has an image pasted in directly rather than typed as a
   name or an =IMAGE() formula (see the "last-last resort" branch
   in parseTextSections and the xlsxItems lookup in openDetail).
   All the *text* on every tab always comes straight from the
   live Sheets API call in fetchAll(), never from the upload —
   the upload is strictly an image patch, not a data source.
   ============================================================ */

function resolveMatchupRowName(rowIndex){

  const rows =
    liveRowsByKey.matchup || [];

  let owner = null;

  for(
    let i = 0;
    i < rows.length && i <= rowIndex;
    i++
  ){

    const a =
      cellText((rows[i] || [])[0]);

    if(
      a &&
      a.toLowerCase().startsWith("c:")
    ){
      owner = a.slice(2).trim();
    }
  }

  return owner;
}


/*
 * Approximate: the runes tab repeats [Name][Explanation] column
 * pairs at different offsets per section, so rather than fully
 * re-deriving the header layout here too, we just scan a small
 * window of columns around where the image was anchored for any
 * text — in practice that's either the name cell itself or the
 * explanation cell right beside it.
 */
function resolveIconRowName(tabKey,rowIndex,col){

  const rows =
    liveRowsByKey[tabKey] || [];

  const row =
    rows[rowIndex] || [];

  for(const delta of [0,-1,1,-2,2]){

    const text =
      cellText(row[col + delta]);

    if(text){
      return text;
    }
  }

  return null;
}


function rebuildXlsxNameIndex(){

  xlsxNameIndex = {};

  for(const tabKey in xlsxSnapshot){

    xlsxNameIndex[tabKey] = {};

    xlsxSnapshot[tabKey].forEach(entry => {

      const key =
        (entry.nameAtUpload || "")
          .trim()
          .toLowerCase();

      if(!key){
        return;
      }

      if(!xlsxNameIndex[tabKey][key]){
        xlsxNameIndex[tabKey][key] = [];
      }

      xlsxNameIndex[tabKey][key].push(entry);
    });

    Object.values(
      xlsxNameIndex[tabKey]
    ).forEach(
      list =>
        list.sort(
          (a,b) =>
            a.row - b.row ||
            a.col - b.col
        )
    );
  }
}


/*
 * XLSX SNAPSHOT PERSISTENCE

 * Previously the uploaded snapshot only lived in memory for the
 * current tab — reload the page and you were back to "no
 * snapshot uploaded, please upload one" every single time, even
 * if nothing had changed. Two ways it can now be found
 * automatically, checked in this order, before the Data Sync
 * tab ever asks for a manual upload:

 * 1. A snapshot file checked into the repo itself, right next
 *    to index.html — see XLSX_REPO_PATH below. This is the
 *    "from the git side" option: commit an .xlsx export under
 *    that exact filename and every visitor gets it automatically,
 *    no per-browser upload needed at all.
 * 2. A snapshot saved locally in this browser from a previous
 *    manual upload (IndexedDB, not localStorage — localStorage
 *    caps out around 5-10MB and these image-heavy snapshots blow
 *    past that easily, which is what was throwing
 *    QuotaExceededError before).

 * Only if neither is found does it fall back to asking for a
 * file to be uploaded by hand.
 */

const XLSX_REPO_PATH = "snapshot.xlsx";

const IDB_NAME = "mordeXlsxDB";
const IDB_STORE = "snapshots";
const IDB_KEY = "xlsxSnapshotV1";


function openXlsxDb(){

  return new Promise((resolve,reject) => {

    if(!window.indexedDB){
      reject(new Error("IndexedDB isn't available in this browser."));
      return;
    }

    const req =
      indexedDB.open(IDB_NAME,1);

    req.onupgradeneeded = () => {
      req.result.createObjectStore(IDB_STORE);
    };

    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}


async function saveXlsxSnapshotToStorage(){

  try{

    const db = await openXlsxDb();

    await new Promise((resolve,reject) => {

      const tx =
        db.transaction(IDB_STORE,"readwrite");

      tx.objectStore(IDB_STORE).put(
        {
          savedAt:Date.now(),
          titleByKey:{
            matchup:titleByKey.matchup,
            runes:titleByKey.runes
          },
          snapshot:xlsxSnapshot
        },
        IDB_KEY
      );

      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });

    return true;

  }catch(err){

    // IndexedDB has a much bigger quota than localStorage did,
    // but browsers can still refuse to grant storage (private
    // browsing, storage already full from something else, etc).
    // Not fatal either way: the snapshot still works for this
    // session, it just won't be there automatically next time.
    console.warn("Could not save xlsx snapshot for next time:",err);

    return false;
  }
}


/*
 * Returns true if a saved snapshot was found and restored from
 * this browser's local IndexedDB, false if there was nothing to
 * restore.
 */
async function tryRestoreXlsxSnapshotFromStorage(){

  let parsed;

  try{

    const db = await openXlsxDb();

    parsed = await new Promise((resolve,reject) => {

      const tx =
        db.transaction(IDB_STORE,"readonly");

      const req =
        tx.objectStore(IDB_STORE).get(IDB_KEY);

      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });

  }catch(err){
    return false;
  }

  if(!parsed || !parsed.snapshot){
    return false;
  }

  /*
   * If the live spreadsheet's tab names have changed since this
   * was saved, the row/column anchors it recorded may no longer
   * line up with anything meaningful — safer to treat it as not
   * found and let the person re-upload than to show images
   * against the wrong rows.
   */
  const savedTitles = parsed.titleByKey || {};

  if(
    savedTitles.matchup !== titleByKey.matchup ||
    savedTitles.runes !== titleByKey.runes
  ){
    return false;
  }

  xlsxSnapshot = parsed.snapshot;
  rebuildXlsxNameIndex();

  return true;
}


async function clearXlsxSnapshotStorage(){

  try{

    const db = await openXlsxDb();

    await new Promise(resolve => {

      const tx =
        db.transaction(IDB_STORE,"readwrite");

      tx.objectStore(IDB_STORE).delete(IDB_KEY);

      tx.oncomplete = resolve;
      tx.onerror = resolve;
    });

  }catch(err){
    // Nothing to do — worst case it just gets overwritten
    // next successful save.
  }
}


/*
 * parseA1Cell: turns "P166" into a 0-indexed {row,col} pair
 * matching how rowsByKey/liveRowsByKey are indexed elsewhere in
 * this file. Handles multi-letter columns (AA, AB, ...) the
 * same way Sheets does.
 */
function parseA1Cell(ref){

  const m =
    String(ref || "").match(/^([A-Z]+)(\d+)$/i);

  if(!m){
    return null;
  }

  const letters =
    m[1].toUpperCase();

  let col = 0;

  for(let i = 0; i < letters.length; i++){
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }

  return {
    row:parseInt(m[2],10) - 1,
    col:col - 1
  };
}


/*
 * Apps Script Web Apps can be genuinely slow — cold starts and
 * a full-sheet scan (see CellImageExport.gs) can take a long
 * time, sometimes tens of seconds. Without a timeout, fetchAll()
 * would sit there waiting indefinitely and the whole site would
 * be stuck on "Loading spreadsheet data…" until it finally
 * resolved. This caps how long any single attempt is allowed to
 * take before giving up and falling through to the next source.
 */
/*
 * Apps Script Web Apps can be genuinely slow — cold starts and
 * a full-sheet scan (see CellImageExport.gs) can take a long
 * time. This used to need to be short because it blocked the
 * whole page; now that resolveImageSnapshotInBackground() runs
 * unawaited, a longer timeout just means icons might show up a
 * little later, not a frozen page — so it's fine to be generous
 * here rather than give up on a legitimately-slow-but-working
 * deployment.
 */
const APPS_SCRIPT_TIMEOUT_MS = 20000;


/*
 * Live in-cell image sync — this is the "get photos that look
 * like this" answer: the Sheets REST API can't read in-cell
 * image content at all, only Apps Script's SpreadsheetApp can
 * (see CellImageExport.gs). If APPS_SCRIPT_URL is configured,
 * this fetches that Web App's JSON output and feeds it into the
 * same {row,col,imageUrl,nameAtUpload} shape the .xlsx snapshot
 * system already uses — so lookupXlsxImages, the stale-badge
 * logic, everything downstream just works unchanged. The one
 * difference: since this is read fresh from the live sheet on
 * every page load rather than a one-time export, "nameAtUpload"
 * is always current, so it can never actually go stale.
 */
async function tryFetchLiveCellImages(){

  if(!APPS_SCRIPT_URL){
    return null;
  }

  const controller =
    new AbortController();

  const timeoutId =
    setTimeout(
      () => controller.abort(),
      APPS_SCRIPT_TIMEOUT_MS
    );

  let res;

  try{

    res = await fetch(
      APPS_SCRIPT_URL,
      {cache:"no-store",signal:controller.signal}
    );

  }catch(err){

    if(err.name === "AbortError"){
      console.warn(
        `Apps Script image endpoint didn't respond within ` +
        `${APPS_SCRIPT_TIMEOUT_MS}ms — skipping it for this load.`
      );
    }else{
      console.warn("Could not reach the Apps Script image endpoint:",err);
    }

    return null;

  }finally{
    clearTimeout(timeoutId);
  }

  if(!res.ok){
    return null;
  }

  let json;

  try{
    json = await res.json();
  }catch(err){
    console.warn("Apps Script image endpoint didn't return valid JSON:",err);
    return null;
  }

  const snapshot = {};
  const warnings = [];

  ["matchup","runes"].forEach(key => {

    const title =
      titleByKey[key];

    const entries =
      title ? json[title] : null;

    if(!entries || !entries.length){

      snapshot[key] = [];

      if(title){
        warnings.push(
          `The live image endpoint didn't find any in-cell pictures ` +
          `on the "${title}" tab.`
        );
      }

      return;
    }

    snapshot[key] =
      entries.map(entry => {

        const parsed =
          parseA1Cell(entry.cell) || {
            row:(entry.row || 1) - 1,
            col:(entry.col || 1) - 1
          };

        return {
          row:parsed.row,
          col:parsed.col,
          imageUrl:entry.url,
          nameAtUpload:
            key === "matchup"
              ? resolveMatchupRowName(parsed.row)
              : resolveIconRowName(key,parsed.row,parsed.col)
        };
      });
  });

  return {snapshot,warnings};
}


/*
 * Checks for an .xlsx file sitting in the site's own files,
 * right next to index.html — see XLSX_REPO_PATH above. Returns
 * the parsed {snapshot,warnings} on success, or null if there's
 * no such file (a normal 404, not an error worth surfacing) or
 * it couldn't be read as a valid workbook.
 */
async function tryFetchXlsxSnapshotFromRepo(){

  let res;

  try{
    res = await fetch(XLSX_REPO_PATH,{cache:"no-store"});
  }catch(err){
    // Network/CORS failure (e.g. running from a local file://
    // path) — treat the same as "not found".
    return null;
  }

  if(!res.ok){
    return null;
  }

  try{

    const buffer =
      await res.arrayBuffer();

    return await loadXlsxSnapshot(buffer);

  }catch(err){

    console.warn(
      `Found "${XLSX_REPO_PATH}" but couldn't read it as an .xlsx file:`,
      err
    );

    return null;
  }
}


/*
 * Runs the three image-snapshot sources in priority order (live
 * Apps Script -> repo file -> browser storage) and applies
 * whichever one succeeds. Called from fetchAll() WITHOUT an
 * await — see the comment at that call site for why. Once this
 * settles (could be instantly, could be several seconds later),
 * it re-renders the Runes tab and the Data Sync tab so whatever
 * it found actually shows up, without the rest of the page ever
 * having been blocked on it.
 */
async function resolveImageSnapshotInBackground(){

  try{

    const liveResult =
      await tryFetchLiveCellImages();

    if(liveResult){

      xlsxSnapshot = liveResult.snapshot;
      lastSyncWarnings = liveResult.warnings;
      xlsxSnapshotSource = "live";

      rebuildXlsxNameIndex();

      // Cache it too, so something still shows up instantly on
      // a repeat visit even before this fetch resolves.
      saveXlsxSnapshotToStorage();

    }else{

      const repoResult =
        await tryFetchXlsxSnapshotFromRepo();

      if(repoResult){

        xlsxSnapshot = repoResult.snapshot;
        lastSyncWarnings = repoResult.warnings;
        xlsxSnapshotSource = "repo";

        rebuildXlsxNameIndex();

        saveXlsxSnapshotToStorage();

      }else{

        const restored =
          await tryRestoreXlsxSnapshotFromStorage();

        xlsxSnapshotSource =
          restored ? "storage" : null;
      }
    }

  }catch(err){

    // Whatever went wrong, the rest of the site already
    // rendered fine without this — just log it and leave the
    // Data Sync tab to report "nothing found" as usual.
    console.warn("Image snapshot resolution failed:",err);

  }finally{

    xlsxSnapshotChecking = false;
  }

  if(liveRowsByKey.runes){

    renderIconSections(
      "tab-runes",
      parseTextSections(
        liveRowsByKey.runes,
        "rune",
        "runes"
      )
    );
  }

  renderSyncPage();
}


function lookupXlsxImage(tabKey,name){

  const matches =
    lookupXlsxImages(tabKey,name);

  return matches.length ? matches[0] : null;
}


function lookupXlsxImages(tabKey,name){

  const key =
    (name || "").trim().toLowerCase();

  if(!key){
    return [];
  }

  const bucket =
    xlsxNameIndex[tabKey] &&
    xlsxNameIndex[tabKey][key];

  return bucket || [];
}


function isXlsxEntryStale(tabKey,entry){

  const currentName =
    tabKey === "matchup"
      ? resolveMatchupRowName(entry.row)
      : resolveIconRowName(tabKey,entry.row,entry.col);

  if(!currentName){
    return true;
  }

  return (
    currentName.trim().toLowerCase() !==
    (entry.nameAtUpload || "").trim().toLowerCase()
  );
}


function computeStaleSummary(){

  const list = [];

  ["matchup","runes"].forEach(tabKey => {

    (xlsxSnapshot[tabKey] || []).forEach(entry => {

      if(isXlsxEntryStale(tabKey,entry)){

        list.push({
          tab:tabKey,
          uploadedName:entry.nameAtUpload || "(unrecognized row)"
        });
      }

    });

  });

  return list;
}


/*
 * resolvePath: resolves a relative OOXML "Target" path (e.g.
 * "../media/image3.png") against the folder the relationship
 * file lives in.
 */
function resolvePath(baseFolder,relativeTarget){

  if(relativeTarget.startsWith("/")){
    return relativeTarget.replace(/^\//,"");
  }

  const parts =
    baseFolder.split("/");

  relativeTarget
    .split("/")
    .forEach(part => {

      if(part === ".."){
        parts.pop();
      }else if(part !== "."){
        parts.push(part);
      }

    });

  return parts.join("/");
}


const OOXML_REL_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";


async function readSheetImageAnchors(zip,sheetPath,tabKey){

  const folder =
    sheetPath.substring(0,sheetPath.lastIndexOf("/"));

  const fileName =
    sheetPath.substring(sheetPath.lastIndexOf("/") + 1);

  const sheetRelsFile =
    zip.file(`${folder}/_rels/${fileName}.rels`);

  if(!sheetRelsFile){
    return [];
  }

  const sheetRelsDoc =
    new DOMParser().parseFromString(
      await sheetRelsFile.async("string"),
      "application/xml"
    );

  let drawingTarget = null;

  sheetRelsDoc
    .querySelectorAll("Relationship")
    .forEach(r => {

      if((r.getAttribute("Type") || "").endsWith("/drawing")){
        drawingTarget = r.getAttribute("Target");
      }

    });

  if(!drawingTarget){
    return [];
  }

  const drawingPath =
    resolvePath(folder,drawingTarget);

  const drawingFile =
    zip.file(drawingPath);

  if(!drawingFile){
    return [];
  }

  const drawingDoc =
    new DOMParser().parseFromString(
      await drawingFile.async("string"),
      "application/xml"
    );

  const drawingFolder =
    drawingPath.substring(0,drawingPath.lastIndexOf("/"));

  const drawingFileName =
    drawingPath.substring(drawingPath.lastIndexOf("/") + 1);

  const drawingRelsFile =
    zip.file(`${drawingFolder}/_rels/${drawingFileName}.rels`);

  const embedTargets = {};

  if(drawingRelsFile){

    const drawingRelsDoc =
      new DOMParser().parseFromString(
        await drawingRelsFile.async("string"),
        "application/xml"
      );

    drawingRelsDoc
      .querySelectorAll("Relationship")
      .forEach(r => {
        embedTargets[r.getAttribute("Id")] =
          r.getAttribute("Target");
      });
  }

  const anchors =
    drawingDoc.getElementsByTagNameNS("*","twoCellAnchor").length
      ? Array.from(drawingDoc.getElementsByTagNameNS("*","twoCellAnchor"))
      : Array.from(drawingDoc.getElementsByTagNameNS("*","oneCellAnchor"));

  const entries = [];

  for(const anchor of anchors){

    const fromEl =
      anchor.getElementsByTagNameNS("*","from")[0];

    if(!fromEl){
      continue;
    }

    const colEl =
      fromEl.getElementsByTagNameNS("*","col")[0];

    const rowEl =
      fromEl.getElementsByTagNameNS("*","row")[0];

    const col =
      parseInt(colEl ? colEl.textContent : "0",10);

    const row =
      parseInt(rowEl ? rowEl.textContent : "0",10);

    const blip =
      anchor.getElementsByTagNameNS("*","blip")[0];

    if(!blip){
      continue;
    }

    const rEmbed =
      blip.getAttributeNS(OOXML_REL_NS,"embed") ||
      blip.getAttribute("r:embed");

    const target =
      embedTargets[rEmbed];

    if(!target){
      continue;
    }

    const mediaPath =
      resolvePath(drawingFolder,target);

    const mediaFile =
      zip.file(mediaPath);

    if(!mediaFile){
      continue;
    }

    /*
     * Data URL instead of a blob URL: blob URLs only live as
     * long as the tab does, so nothing could ever be saved for
     * next time. A data URL is just a string, so it can be
     * stored directly (see saveXlsxSnapshotToStorage below) and
     * read back exactly as it was without re-uploading the file.
     */
    const base64 =
      await mediaFile.async("base64");

    const ext =
      (mediaPath.split(".").pop() || "png").toLowerCase();

    const mime =
      ext === "jpg" ? "jpeg" : ext;

    const imageUrl =
      `data:image/${mime};base64,${base64}`;

    const nameAtUpload =
      tabKey === "matchup"
        ? resolveMatchupRowName(row)
        : resolveIconRowName(tabKey,row,col);

    entries.push({
      row:row,
      col:col,
      imageUrl:imageUrl,
      nameAtUpload:nameAtUpload
    });
  }

  return entries;
}


async function loadXlsxSnapshot(file){

  if(typeof JSZip === "undefined"){
    throw new Error(
      "JSZip didn't load — check your connection and try again."
    );
  }

  const zip =
    await JSZip.loadAsync(file);

  const workbookFile =
    zip.file("xl/workbook.xml");

  const workbookRelsFile =
    zip.file("xl/_rels/workbook.xml.rels");

  if(!workbookFile || !workbookRelsFile){
    throw new Error(
      "That doesn't look like a valid .xlsx file."
    );
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
      relTargets[r.getAttribute("Id")] =
        r.getAttribute("Target");
    });

  const sheetPaths = {};

  wbDoc
    .querySelectorAll("sheet")
    .forEach(s => {

      const name =
        s.getAttribute("name");

      const rid =
        s.getAttributeNS(OOXML_REL_NS,"id") ||
        s.getAttribute("r:id");

      const target =
        relTargets[rid];

      if(name && target){
        sheetPaths[name] =
          "xl/" + target.replace(/^\/?/,"");
      }

    });

  const wanted = {
    matchup:titleByKey.matchup,
    runes:titleByKey.runes
  };

  const snapshot = {};
  const warnings = [];

  /*
   * This is the part that used to be a silent no-op: if the
   * uploaded file didn't have a sheet matching our live tab
   * title (wrong export, renamed tab, etc.) we'd just set an
   * empty array and move on — the upload looked "successful"
   * with nothing to show for it and no way to know why. Now
   * every gap gets a warning that renderSyncPage actually
   * displays.
   */
  for(const key in wanted){

    const title =
      wanted[key];

    const path =
      title ? sheetPaths[title] : null;

    if(!path){

      snapshot[key] = [];

      warnings.push(
        `Couldn't find a sheet named "${title || key}" inside this ` +
        `file — make sure you exported the same spreadsheet that's ` +
        `live-synced, with tab names unchanged.`
      );

      continue;
    }

    const entries =
      await readSheetImageAnchors(zip,path,key);

    snapshot[key] = entries;

    if(!entries.length){

      warnings.push(
        `The "${title}" tab was found, but this file has no ` +
        `embedded images on it — nothing pasted directly into a cell ` +
        `there to read.`
      );
    }
  }

  return {snapshot,warnings};
}


function renderSyncPage(){

  const el =
    document.getElementById("tab-sync");

  if(!el){
    return;
  }

  const hasSnapshot =
    Object.keys(xlsxSnapshot).some(
      k => (xlsxSnapshot[k] || []).length
    );

  const staleList =
    computeStaleSummary();

  el.innerHTML =
    `
      <div class="sync-page">

        <div class="sync-card">

          <h3>
            Image Snapshot Sync
          </h3>

          <p>
            Live text always comes straight from the spreadsheet.
            Some images — item icons pasted directly into the
            Matchup tab, and rune icons pasted into the Runes tab —
            can't be read through the Sheets REST API at all, only
            through Apps Script. This is found automatically three
            ways, checked in this order:
            ${
              APPS_SCRIPT_URL
                ? "a live Apps Script endpoint (configured, reading straight off the sheet — nothing ever goes stale),"
                : `a live Apps Script endpoint (<strong>not configured yet</strong> — see <code>CellImageExport.gs</code> to set one up, it's the best option and skips needing an .xlsx at all),`
            }
            a file named <code>${escapeHtml(XLSX_REPO_PATH)}</code>
            committed next to <code>index.html</code> in the site's
            own files, or a copy saved locally from a manual
            upload in this browser. Only if none of those exist
            does it ask you to upload one below.
          </p>

          <input
            type="file"
            id="xlsx-input"
            accept=".xlsx"
          >

          <div class="sync-status" id="sync-status">
            ${
              hasSnapshot
                ? (
                    xlsxSnapshotSource === "live"
                      ? "Reading live from the sheet via Apps Script — always current, nothing to upload."
                      : xlsxSnapshotSource === "repo"
                        ? `Loaded automatically from "${escapeHtml(XLSX_REPO_PATH)}" in the site files.`
                        : xlsxSnapshotSource === "storage"
                          ? "Restored automatically from your last upload in this browser."
                          : "Snapshot loaded and saved for next time."
                  )
                : xlsxSnapshotChecking
                  ? "Checking for a live Apps Script feed, a repo file, and a saved copy in this browser… the rest of the page doesn't wait on this, so feel free to look around."
                  : `No snapshot found — checked the live Apps Script endpoint${APPS_SCRIPT_URL ? "" : " (not configured)"}, "${escapeHtml(XLSX_REPO_PATH)}" in the site files, and this browser's saved copy, found none. Upload an .xlsx export below, or set up the live endpoint for something that never needs re-uploading.`
            }
          </div>

          ${
            lastSyncWarnings.length
              ? `
                <div class="sync-warning">
                  ${
                    lastSyncWarnings.map(
                      w => `<p>⚠ ${escapeHtml(w)}</p>`
                    ).join("")
                  }
                </div>
              `
              : ""
          }

          ${
            xlsxSnapshotSource === "live"
              ? `
                <p class="sync-disclaimer">
                  This is being read live off the sheet, not stored
                  anywhere — there's nothing to clear. To stop
                  using it, remove the URL from
                  <code>APPS_SCRIPT_URL</code> in App.js.
                </p>
              `
              : `
                <button
                  class="wav-btn"
                  id="xlsx-clear"
                  ${hasSnapshot ? "" : "disabled"}
                >
                  Clear snapshot
                </button>
              `
          }

          <p class="sync-disclaimer">
            Uploaded images are a snapshot from the moment you
            exported the file — they can drift out of sync with
            the live sheet over time. Anything listed below has
            changed on the live sheet since your last upload, so
            treat that image as possibly outdated; it'll disappear
            from this list on its own once the data lines up again.
          </p>

          <div id="sync-stale-list">
            ${
              staleList.length
                ? `
                  <ul>
                    ${
                      staleList.map(
                        s => `
                          <li>
                            <strong>${escapeHtml(s.uploadedName)}</strong>
                            (${escapeHtml(s.tab)} tab) —
                            row content has changed since upload
                          </li>
                        `
                      ).join("")
                    }
                  </ul>
                `
                : hasSnapshot
                  ? `<p class="sync-ok">Everything lines up — no stale images right now.</p>`
                  : ""
            }
          </div>

        </div>

      </div>
    `;

  updateSyncTabDot(staleList.length);

  document
    .getElementById("xlsx-input")
    .addEventListener("change",handleXlsxUpload);

  const clearBtn =
    document.getElementById("xlsx-clear");

  if(clearBtn){

    clearBtn.addEventListener("click",async () => {

      xlsxSnapshot = {};
      xlsxNameIndex = {};
      lastSyncWarnings = [];
      xlsxSnapshotSource = null;

      await clearXlsxSnapshotStorage();

      renderSyncPage();

      renderIconSections(
        "tab-runes",
        parseTextSections(
          liveRowsByKey.runes,
          "rune",
          "runes"
        )
      );

    });
  }
}


/*
 * Puts a small red dot on the "Data Sync" tab button itself so
 * a live-vs-snapshot mismatch is visible without having to open
 * that tab to notice it.
 */
function updateSyncTabDot(staleCount){

  const btn =
    document.querySelector('button[data-tab="sync"]');

  if(!btn){
    return;
  }

  let dot =
    btn.querySelector(".stale-dot");

  if(staleCount > 0){

    if(!dot){
      dot = document.createElement("span");
      dot.className = "stale-dot";
      btn.appendChild(dot);
    }

    dot.title =
      `${staleCount} image${staleCount === 1 ? "" : "s"} out of sync with the live sheet`;

  }else if(dot){
    dot.remove();
  }
}


async function handleXlsxUpload(e){

  const file =
    e.target.files[0];

  if(!file){
    return;
  }

  const statusEl =
    document.getElementById("sync-status");

  if(statusEl){
    statusEl.textContent =
      "Reading file…";
  }

  try{

    const result =
      await loadXlsxSnapshot(file);

    xlsxSnapshot = result.snapshot;
    lastSyncWarnings = result.warnings;
    xlsxSnapshotSource = "upload";

    rebuildXlsxNameIndex();

    const saved =
      await saveXlsxSnapshotToStorage();

    if(!saved){
      lastSyncWarnings = [
        ...lastSyncWarnings,
        "This snapshot is a bit large to save for next time — " +
        "it'll work for this session, but you'll need to " +
        "re-upload after refreshing the page."
      ];
    }

    renderIconSections(
      "tab-runes",
      parseTextSections(
        liveRowsByKey.runes,
        "rune",
        "runes"
      )
    );

    renderSyncPage();

  }catch(err){

    console.error(err);

    lastSyncWarnings = [];

    if(statusEl){
      statusEl.textContent =
        "Couldn't read that file: " + err.message;
    }
  }
}


/* ============================================================
   TABS
   ============================================================ */

function switchTab(tab){

  activeTab =
    tab;


  document
    .querySelectorAll(
      "nav.tabs button[data-tab]"
    )
    .forEach(
      b =>
        b.classList.toggle(
          "active",
          b.dataset.tab === tab
        )
    );


  matchupGrid.style.display =
    tab === "matchup"
      ? "grid"
      : "none";


  searchWrap.style.display =
    tab === "matchup"
      ? "block"
      : "none";


  document
    .querySelectorAll(
      ".tab-page"
    )
    .forEach(
      p =>
        p.style.display =
          "none"
    );


  if(tab !== "matchup"){

    const page =
      document.getElementById(
        "tab-" + tab
      );

    if(page){
      page.style.display =
        "block";
    }
  }
}


tabnav.addEventListener(
  "click",
  e => {

    const btn =
      e.target.closest(
        "button[data-tab]"
      );

    if(btn){

      switchTab(
        btn.dataset.tab
      );

      if(navToggle){
        navToggle.classList.remove("open");
      }

      if(tabsInner){
        tabsInner.classList.remove("open");
      }
    }

  }
);


/* ============================================================
   MOBILE NAV TOGGLE
   ============================================================ */

if(navToggle && tabsInner){

  navToggle.addEventListener(
    "click",
    () => {
      navToggle.classList.toggle("open");
      tabsInner.classList.toggle("open");
    }
  );
}


/* ============================================================
   SEARCH
   ============================================================ */

searchEl.addEventListener(
  "input",
  () => {

    const q =
      searchEl.value
        .trim()
        .toLowerCase();


    const filtered =
      q
        ? champions.filter(
            c =>
              c.name
                .toLowerCase()
                .includes(q)
          )
        : champions;


    renderMatchupGrid(
      filtered
    );
  }
);


searchEl.addEventListener(
  "keydown",
  e => {

    if(e.key === "Enter"){

      const q =
        searchEl.value
          .trim()
          .toLowerCase();


      const match =
        champions.find(
          c =>
            c.name
              .toLowerCase()
              .includes(q)
        );


      if(match){
        openDetail(match);
      }
    }

  }
);


/* ============================================================
   MORDEKAISER MACE CURSOR

   Body classes swap the CSS cursor (drawn in the stylesheet):
   default resting mace, a raised mace on hover over anything
   clickable, a swinging mace on mousedown, and — after the
   page has sat idle for a while — the dim "Realm of Death"
   mace plus a slow shadow veil over the whole page (Nightfall).
   Any movement, click, key press, touch, or scroll cancels
   Nightfall and restarts the idle clock.
   ============================================================ */

const NIGHTFALL_DELAY_MS = 6000;

const HOVER_SELECTOR =
  "a, button, input, textarea, .card, .icon-card, " +
  ".wav-hole, [data-tab], .close-btn";

let nightfallTimer = null;


function clearNightfall(){

  document.body.classList.remove("nightfall-active");

  if(nightfallVeil){
    nightfallVeil.classList.remove("show");
  }
}


function armNightfall(){

  clearTimeout(nightfallTimer);

  nightfallTimer =
    setTimeout(() => {

      document.body.classList.add("nightfall-active");

      if(nightfallVeil){
        nightfallVeil.classList.add("show");
      }

    },NIGHTFALL_DELAY_MS);
}


function registerActivity(){
  clearNightfall();
  armNightfall();
}


[
  "mousemove",
  "mousedown",
  "keydown",
  "touchstart",
  "scroll"
].forEach(evt => {

  window.addEventListener(
    evt,
    registerActivity,
    {passive:true}
  );

});


document.addEventListener(
  "mouseover",
  e => {

    if(e.target.closest(HOVER_SELECTOR)){
      document.body.classList.add("cursor-hover");
    }

  }
);


document.addEventListener(
  "mouseout",
  e => {

    const stillOverHoverable =
      e.relatedTarget &&
      e.relatedTarget.closest &&
      e.relatedTarget.closest(HOVER_SELECTOR);

    if(
      e.target.closest(HOVER_SELECTOR) &&
      !stillOverHoverable
    ){
      document.body.classList.remove("cursor-hover");
    }

  }
);


document.addEventListener(
  "mousedown",
  () => {
    document.body.classList.add("cursor-click");
  }
);


document.addEventListener(
  "mouseup",
  () => {
    document.body.classList.remove("cursor-click");
  }
);


armNightfall();


/* ============================================================
   START
   ============================================================ */

fetchAll();