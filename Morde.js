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
 */
let titleByKey = {};
let liveRowsByKey = {};
let xlsxSnapshot = {};
let xlsxNameIndex = {};


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

  return null;
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


    champions =
      parseChampions(
        rowsByKey.matchup
      );


    renderMatchupGrid(
      champions
    );


    renderIconSections(
      "tab-itemization",
      parseTextSections(
        rowsByKey.itemization,
        "item"
      )
    );


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
     * NOT from Google Sheets.
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


    const runeBuild =
      remaining.find(
        t => t.includes("->")
      ) || "";


    const itemization =
      remaining.find(
        t =>
          t !== runeBuild &&
          /\(.*\)/.test(t) &&
          t.length > 20
      ) || "";


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
                    check the spelling in the spreadsheet's "C:" name
                    cell against Data Dragon's champion list.
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

              <pre class="rune-text">${
                escapeHtml(
                  c.runeBuild
                )
              }</pre>
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
                  ? `
                    <pre class="item-text">${
                      escapeHtml(
                        c.itemization
                      )
                    }</pre>
                  `
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

const NAME_HEADER_RE = /^(rune|item|items)$/i;
const EXPLANATION_HEADER_RE = /^explanation$/i;


function detectPairColumns(row){

  const cells =
    (row || []).map(cellText);

  const pairs = [];

  for(let i = 0; i < cells.length - 1; i++){

    if(
      NAME_HEADER_RE.test(cells[i]) &&
      EXPLANATION_HEADER_RE.test(cells[i + 1])
    ){
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

    const blob =
      await mediaFile.async("blob");

    const imageUrl =
      URL.createObjectURL(blob);

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

  for(const key in wanted){

    const title =
      wanted[key];

    const path =
      title ? sheetPaths[title] : null;

    if(!path){
      snapshot[key] = [];
      continue;
    }

    snapshot[key] =
      await readSheetImageAnchors(zip,path,key);
  }

  return snapshot;
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
            can't be read through the Sheets API at all, so upload
            a fresh <code>.xlsx</code> export of the sheet here
            whenever you update those pictures.
          </p>

          <input
            type="file"
            id="xlsx-input"
            accept=".xlsx"
          >

          <div class="sync-status" id="sync-status">
            ${
              hasSnapshot
                ? "Snapshot loaded for this session."
                : "No snapshot uploaded yet — pasted-in images will show as blank until you upload one."
            }
          </div>

          <button
            class="wav-btn"
            id="xlsx-clear"
            ${hasSnapshot ? "" : "disabled"}
          >
            Clear snapshot
          </button>

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

  document
    .getElementById("xlsx-input")
    .addEventListener("change",handleXlsxUpload);

  const clearBtn =
    document.getElementById("xlsx-clear");

  if(clearBtn){

    clearBtn.addEventListener("click",() => {

      xlsxSnapshot = {};
      xlsxNameIndex = {};

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

    xlsxSnapshot =
      await loadXlsxSnapshot(file);

    rebuildXlsxNameIndex();

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

    if(statusEl){
      statusEl.textContent =
        "Couldn't read that file: " + err.message;
    }
  }
}


/* ============================================================
   GAMES PAGE
   ============================================================ */

let wavState = null;


function renderGamesPage(){

  const el =
    document.getElementById("tab-games");

  if(!el){
    return;
  }

  const storedBest =
    parseInt(
      localStorage.getItem("wavBest") || "0",
      10
    );

  el.innerHTML =
    `
      <div class="games-page">

        <div class="game-card" id="wav-card">

          <h3>
            Whack-a-Vayne
          </h3>

          <p>
            She keeps peeking out of the shadows. Click her
            before she slips back into stealth. 30 seconds,
            as many hits as you can land.
          </p>

          <div class="wav-hud">
            <div>Score <span id="wav-score">0</span></div>
            <div>Time <span id="wav-time">30</span>s</div>
            <div>Best <span id="wav-best">${storedBest}</span></div>
          </div>

          <div class="wav-grid" id="wav-grid"></div>

          <button class="wav-btn" id="wav-start">
            Start
          </button>

          <div class="wav-end" id="wav-end"></div>

        </div>

      </div>
    `;


  const grid =
    document.getElementById("wav-grid");

  grid.innerHTML =
    Array.from(
      {length:9},
      (_,i) =>
        `<div class="wav-hole" data-hole="${i}">
          <img class="wav-vayne" alt="Vayne" draggable="false">
        </div>`
    ).join("");


  const vayneIcon =
    championSquareImage("Vayne");

  if(vayneIcon){

    grid
      .querySelectorAll(".wav-vayne")
      .forEach(img => {
        img.src = vayneIcon;
      });
  }


  grid
    .querySelectorAll(".wav-hole")
    .forEach(hole => {

      hole.addEventListener(
        "click",
        () => whackHole(hole)
      );

    });


  document
    .getElementById("wav-start")
    .addEventListener(
      "click",
      startWhackAVayne
    );
}


function whackHole(hole){

  if(!wavState || !wavState.running){
    return;
  }

  if(!hole.classList.contains("active")){
    return;
  }

  clearTimeout(hole._wavTimeout);

  hole.classList.remove("active");
  hole.classList.add("hit");

  setTimeout(
    () => hole.classList.remove("hit"),
    150
  );

  wavState.score++;

  document.getElementById("wav-score").textContent =
    wavState.score;
}


function startWhackAVayne(){

  const grid =
    document.getElementById("wav-grid");

  const holes =
    Array.from(
      grid.querySelectorAll(".wav-hole")
    );

  const startBtn =
    document.getElementById("wav-start");

  const endEl =
    document.getElementById("wav-end");


  if(wavState){
    clearInterval(wavState.tickTimer);
    clearTimeout(wavState.spawnTimer);
  }


  holes.forEach(h => {
    h.classList.remove("active","hit");
    clearTimeout(h._wavTimeout);
  });


  endEl.textContent = "";
  startBtn.disabled = true;
  startBtn.textContent = "Whacking…";


  wavState = {
    running:true,
    score:0,
    timeLeft:30,
    spawnTimer:null,
    tickTimer:null
  };


  document.getElementById("wav-score").textContent = "0";
  document.getElementById("wav-time").textContent = "30";


  const spawnLoop = () => {

    if(!wavState || !wavState.running){
      return;
    }

    const idle =
      holes.filter(
        h => !h.classList.contains("active")
      );

    if(idle.length){

      const hole =
        idle[
          Math.floor(Math.random() * idle.length)
        ];

      hole.classList.add("active");

      const upTime =
        500 + Math.random() * 500;

      hole._wavTimeout =
        setTimeout(() => {
          hole.classList.remove("active");
        },upTime);
    }

    const nextSpawn =
      450 + Math.random() * 500;

    wavState.spawnTimer =
      setTimeout(spawnLoop,nextSpawn);
  };


  spawnLoop();


  wavState.tickTimer =
    setInterval(() => {

      wavState.timeLeft--;

      document.getElementById("wav-time").textContent =
        Math.max(0,wavState.timeLeft);

      if(wavState.timeLeft <= 0){
        endWhackAVayne();
      }

    },1000);
}


function endWhackAVayne(){

  if(!wavState){
    return;
  }

  wavState.running = false;

  clearInterval(wavState.tickTimer);
  clearTimeout(wavState.spawnTimer);


  const grid =
    document.getElementById("wav-grid");

  grid
    .querySelectorAll(".wav-hole")
    .forEach(h => {
      h.classList.remove("active");
      clearTimeout(h._wavTimeout);
    });


  const best =
    parseInt(
      localStorage.getItem("wavBest") || "0",
      10
    );

  if(wavState.score > best){
    localStorage.setItem("wavBest",wavState.score);
  }

  document.getElementById("wav-best").textContent =
    Math.max(best,wavState.score);

  document.getElementById("wav-end").textContent =
    `Time's up — you landed ${wavState.score} hit${
      wavState.score === 1 ? "" : "s"
    } on Vayne.`;


  const startBtn =
    document.getElementById("wav-start");

  startBtn.disabled = false;
  startBtn.textContent = "Play again";
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