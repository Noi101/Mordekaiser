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

let wavState = null;
let mordekaiserSpells = null;
let itemizationFlat = [];
let gameActiveCleanup = null;


const GAME_LIST = [
  {id:"whack",  name:"Whack-a-Vayne",           blurb:"Click her before she slips back into stealth."},
  {id:"dodge",  name:"Realm of Death Dodge",    blurb:"Steer clear of the spectral bolts as long as you can."},
  {id:"combo",  name:"Combo Trainer",           blurb:"Watch the ability sequence, then repeat it back."},
  {id:"cs",     name:"CS Practice",             blurb:"Time your last hit on a shrinking-HP minion."},
  {id:"guess",  name:"Guess the Champion",      blurb:"Splash art sharpens — name them before it's clear."},
  {id:"trivia", name:"Matchup Trivia",          blurb:"Quiz yourself on your own written matchup ratings."},
  {id:"runes",  name:"Rune Path Builder",       blurb:"Click his rune page together from memory."},
  {id:"items",  name:"Itemization Speed Round", blurb:"Match the description to the right item."},
  {id:"memory", name:"Ghost Memory",            blurb:"Classic pairs, played with champion icons."},
  {id:"runner", name:"Realm Runner",            blurb:"Endless runner through the Realm of Death."},
  {id:"souls",  name:"Soul Tracker",            blurb:"Collect drifting souls before they fade — bigger ones are worth more."}
];


function renderGamesPage(){

  const el =
    document.getElementById("tab-games");

  if(!el){
    return;
  }

  el.innerHTML =
    `
      <div class="games-menu" id="games-menu"></div>
      <div class="game-stage" id="game-stage" style="display:none;"></div>
    `;

  showGamesMenu();
}


function showGamesMenu(){

  if(gameActiveCleanup){
    gameActiveCleanup();
    gameActiveCleanup = null;
  }

  const menu =
    document.getElementById("games-menu");

  const stage =
    document.getElementById("game-stage");

  if(!menu || !stage){
    return;
  }

  stage.style.display = "none";
  stage.innerHTML = "";
  menu.style.display = "grid";

  menu.innerHTML =
    GAME_LIST.map(
      g => `
        <div class="arcade-card" data-game="${g.id}">
          <h3>${escapeHtml(g.name)}</h3>
          <p>${escapeHtml(g.blurb)}</p>
          <button class="wav-btn">Play</button>
        </div>
      `
    ).join("");

  menu
    .querySelectorAll(".arcade-card")
    .forEach(card => {
      card.addEventListener(
        "click",
        () => launchGame(card.dataset.game)
      );
    });
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

    document.getElementById("dodge-time").textContent =
      ((Date.now()-startTime)/1000).toFixed(1);

    rafId = requestAnimationFrame(loop);
  }

  function startDodge(){

    bolts = [];
    player = {x:170,y:170,r:9};
    running = true;
    startTime = Date.now();

    document.getElementById("dodge-end").textContent = "";
    document.getElementById("dodge-start").disabled = true;
    document.getElementById("dodge-start").textContent = "Running…";

    clearInterval(spawnTimer);
    spawnTimer = setInterval(spawnBolt,550);

    loop();
  }

  function endDodge(){

    running = false;
    cancelAnimationFrame(rafId);
    clearInterval(spawnTimer);

    const elapsed = (Date.now()-startTime) / 1000;
    const best = parseFloat(localStorage.getItem("dodgeBest") || "0");

    if(elapsed > best){
      localStorage.setItem("dodgeBest",elapsed.toFixed(1));
    }

    document.getElementById("dodge-best").textContent =
      Math.max(best,elapsed).toFixed(1);

    document.getElementById("dodge-end").textContent =
      `Caught by the shadow realm at ${elapsed.toFixed(1)}s.`;

    const btn = document.getElementById("dodge-start");
    btn.disabled = false;
    btn.textContent = "Try again";
  }

  document.getElementById("dodge-start").addEventListener("click",startDodge);

  gameActiveCleanup = () => {
    running = false;
    cancelAnimationFrame(rafId);
    clearInterval(spawnTimer);
  };
}


/* ============================================================
   COMBO TRAINER
   ============================================================ */

async function loadMordekaiserSpells(){

  if(mordekaiserSpells){
    return mordekaiserSpells;
  }

  const res =
    await fetch(
      `${DDRAGON_BASE}/cdn/${ddragonVersion}/data/en_US/champion/Mordekaiser.json`,
      {cache:"no-store"}
    );

  const json = await res.json();
  const spells = json.data.Mordekaiser.spells;

  mordekaiserSpells =
    ["Q","W","E","R"].map((key,i) => ({
      key:key,
      name:spells[i].name,
      icon:`${DDRAGON_BASE}/cdn/${ddragonVersion}/img/spell/${spells[i].image.full}`
    }));

  return mordekaiserSpells;
}


async function setupComboGame(container){

  container.innerHTML =
    `<div class="game-card"><p class="loading-small">Loading ability data…</p></div>`;

  let spells;

  try{
    spells = await loadMordekaiserSpells();
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

        answered = true;

        if(btn.dataset.label === correct.label){
          score++;
          btn.classList.add("rune-used");
          document.getElementById("item-score").textContent = score;
        }else{
          btn.classList.add("trivia-wrong");
        }

        setTimeout(nextRound,700);
      });

    });
  }

  function finishItemGame(){

    const best = parseInt(localStorage.getItem("itemSpeedBest") || "0",10);

    if(score > best){
      localStorage.setItem("itemSpeedBest",score);
    }

    document.getElementById("item-best").textContent = Math.max(best,score);
    document.getElementById("item-end").textContent = `Final score: ${score}/${ROUNDS}.`;
    document.getElementById("item-choices").innerHTML = "";

    document.getElementById("item-start").disabled = false;
    document.getElementById("item-start").textContent = "Play again";
  }

  document.getElementById("item-start").addEventListener("click",() => {

    round = 0;
    score = 0;

    document.getElementById("item-score").textContent = "0";
    document.getElementById("item-end").textContent = "";
    document.getElementById("item-start").disabled = true;
    document.getElementById("item-start").textContent = "Quizzing…";

    nextRound();
  });

  gameActiveCleanup = () => {};
}


/* ============================================================
   GHOST MEMORY
   ============================================================ */

function setupMemoryGame(container){

  const storedBest =
    parseInt(localStorage.getItem("memoryBest") || "0",10);

  const allChamps = Object.values(ddragonChampions);

  if(allChamps.length < 8){
    container.innerHTML =
      `<div class="game-card"><p>Champion data hasn't loaded yet — try again in a moment.</p></div>`;
    return;
  }

  container.innerHTML =
    `
      <div class="game-card">
        <h3>Ghost Memory</h3>
        <p>Classic pairs — flip two cards, find the matching champion.</p>
        <div class="wav-hud">
          <div>Moves <span id="mem-moves">0</span></div>
          <div>Best <span id="mem-best">${storedBest || "—"}</span></div>
        </div>
        <div class="memory-grid" id="memory-grid"></div>
        <button class="wav-btn" id="mem-start">Shuffle &amp; Start</button>
        <div class="wav-end" id="mem-end"></div>
      </div>
    `;

  let flipped = [];
  let matched = 0;
  let moves = 0;
  let lock = false;

  function flipCard(card){

    if(lock || card.classList.contains("flipped") || card.classList.contains("matched")){
      return;
    }

    card.classList.add("flipped");
    flipped.push(card);

    if(flipped.length === 2){

      moves++;
      document.getElementById("mem-moves").textContent = moves;
      lock = true;

      const [a,b] = flipped;

      if(a.dataset.id === b.dataset.id){

        a.classList.add("matched");
        b.classList.add("matched");
        flipped = [];
        lock = false;
        matched++;

        if(matched === 8){
          finishMemory();
        }

      }else{

        setTimeout(() => {
          a.classList.remove("flipped");
          b.classList.remove("flipped");
          flipped = [];
          lock = false;
        },700);
      }
    }
  }

  function finishMemory(){

    const best = parseInt(localStorage.getItem("memoryBest") || "0",10);

    if(!best || moves < best){
      localStorage.setItem("memoryBest",moves);
    }

    document.getElementById("mem-best").textContent =
      best ? Math.min(best,moves) : moves;

    document.getElementById("mem-end").textContent = `Solved in ${moves} moves.`;
  }

  function buildBoard(){

    const chosen =
      [...allChamps].sort(() => Math.random()-0.5).slice(0,8);

    const cards =
      [...chosen,...chosen]
        .map(c => ({id:c.id,name:c.name}))
        .sort(() => Math.random()-0.5);

    matched = 0;
    moves = 0;
    flipped = [];
    lock = false;

    document.getElementById("mem-moves").textContent = "0";
    document.getElementById("mem-end").textContent = "";

    const grid = document.getElementById("memory-grid");

    grid.innerHTML =
      cards.map(
        (c,i) => `
          <button class="memory-card" data-idx="${i}" data-id="${c.id}">
            <span class="memory-back">?</span>
            <img class="memory-face" src="${championSquareImage(c.name) || ""}" alt="${escapeHtml(c.name)}">
          </button>
        `
      ).join("");

    grid.querySelectorAll(".memory-card").forEach(card => {
      card.addEventListener("click",() => flipCard(card));
    });
  }

  document.getElementById("mem-start").addEventListener("click",buildBoard);

  gameActiveCleanup = () => {};
}


/* ============================================================
   REALM RUNNER
   ============================================================ */

function setupRunnerGame(container){

  const storedBest =
    parseInt(localStorage.getItem("runnerBest") || "0",10);

  container.innerHTML =
    `
      <div class="game-card">
        <h3>Realm Runner</h3>
        <p>Space, tap, or click to jump. Obstacles get faster the longer you last.</p>
        <div class="wav-hud">
          <div>Score <span id="runner-score">0</span></div>
          <div>Best <span id="runner-best">${storedBest}</span></div>
        </div>
        <canvas id="runner-canvas" width="340" height="170"></canvas>
        <button class="wav-btn" id="runner-start">Start</button>
        <div class="wav-end" id="runner-end"></div>
      </div>
    `;

  const canvas = document.getElementById("runner-canvas");
  const ctx = canvas.getContext("2d");

  const groundY = 140;

  let player = {x:40,y:groundY,vy:0,r:11,onGround:true};
  let obstacles = [];
  let speed = 3;
  let running = false;
  let rafId = null;
  let spawnTimer = null;
  let score = 0;

  function jump(){
    if(running && player.onGround){
      player.vy = -8.5;
      player.onGround = false;
    }
  }

  function keyHandler(e){
    if(e.code === "Space"){
      e.preventDefault();
      jump();
    }
  }

  window.addEventListener("keydown",keyHandler);
  canvas.addEventListener("mousedown",jump);
  canvas.addEventListener("touchstart",e => { jump(); e.preventDefault(); },{passive:false});

  function spawnObstacle(){
    obstacles.push({x:canvas.width+10,w:12+Math.random()*10,h:18+Math.random()*18});
  }

  function loop(){

    if(!running){
      return;
    }

    ctx.clearRect(0,0,canvas.width,canvas.height);

    ctx.strokeStyle = "#2a2732";
    ctx.beginPath();
    ctx.moveTo(0,groundY+11);
    ctx.lineTo(canvas.width,groundY+11);
    ctx.stroke();

    player.vy += 0.5;
    player.y += player.vy;

    if(player.y >= groundY){
      player.y = groundY;
      player.vy = 0;
      player.onGround = true;
    }

    ctx.fillStyle = "#4fe3b0";
    ctx.beginPath();
    ctx.arc(player.x,player.y,player.r,0,Math.PI*2);
    ctx.fill();

    speed += 0.0015;

    obstacles.forEach(o => { o.x -= speed; });
    obstacles = obstacles.filter(o => o.x + o.w > -10);

    let hit = false;

    obstacles.forEach(o => {

      ctx.fillStyle = "#b54b3d";
      ctx.fillRect(o.x,groundY+11-o.h,o.w,o.h);

      const closestX = Math.max(o.x,Math.min(player.x,o.x+o.w));
      const closestY = Math.max(groundY+11-o.h,Math.min(player.y,groundY+11));
      const dist = Math.hypot(player.x-closestX,player.y-closestY);

      if(dist < player.r){
        hit = true;
      }
    });

    if(hit){
      endRunner();
      return;
    }

    score += 1;
    document.getElementById("runner-score").textContent = Math.floor(score/6);

    rafId = requestAnimationFrame(loop);
  }

  function startRunner(){

    obstacles = [];
    player = {x:40,y:groundY,vy:0,r:11,onGround:true};
    speed = 3;
    score = 0;
    running = true;

    document.getElementById("runner-score").textContent = "0";
    document.getElementById("runner-end").textContent = "";
    document.getElementById("runner-start").disabled = true;
    document.getElementById("runner-start").textContent = "Running…";

    clearInterval(spawnTimer);
    spawnTimer = setInterval(spawnObstacle,1400);

    rafId = requestAnimationFrame(loop);
  }

  function endRunner(){

    running = false;
    cancelAnimationFrame(rafId);
    clearInterval(spawnTimer);

    const finalScore = Math.floor(score/6);
    const best = parseInt(localStorage.getItem("runnerBest") || "0",10);

    if(finalScore > best){
      localStorage.setItem("runnerBest",finalScore);
    }

    document.getElementById("runner-best").textContent = Math.max(best,finalScore);
    document.getElementById("runner-end").textContent = `Caught at ${finalScore}m.`;

    document.getElementById("runner-start").disabled = false;
    document.getElementById("runner-start").textContent = "Try again";
  }

  document.getElementById("runner-start").addEventListener("click",startRunner);

  gameActiveCleanup = () => {
    running = false;
    cancelAnimationFrame(rafId);
    clearInterval(spawnTimer);
    window.removeEventListener("keydown",keyHandler);
  };
}


/* ============================================================
   SOUL TRACKER

   Mordekaiser drags enemy souls into his realm — so here, souls
   of three tiers (small/common, mid, and rare/high-value) drift
   and bounce around the canvas at different speeds and fade at
   different rates. Click a soul to bank its value before it
   fades out; bigger souls are worth more but disappear faster,
   so it's a constant judgment call about which one to chase.
   ============================================================ */

function setupSoulTrackerGame(container){

  const storedBest =
    parseInt(localStorage.getItem("soulBest") || "0",10);

  const ROUND_SECONDS = 30;

  container.innerHTML =
    `
      <div class="game-card">
        <h3>Soul Tracker</h3>
        <p>Souls of different value drift around the Realm of Death. Click them before they fade — bigger, brighter souls are worth more but don't stick around long.</p>
        <div class="wav-hud">
          <div>Score <span id="soul-score">0</span></div>
          <div>Time <span id="soul-time">${ROUND_SECONDS}</span>s</div>
          <div>Best <span id="soul-best">${storedBest}</span></div>
        </div>
        <canvas id="soul-canvas" width="340" height="300"></canvas>
        <button class="wav-btn" id="soul-start">Start</button>
        <div class="wav-end" id="soul-end"></div>
      </div>
    `;

  const canvas = document.getElementById("soul-canvas");
  const ctx = canvas.getContext("2d");

  const SOUL_TIERS = [
    {value:1, r:9,  color:"#4fe3b0", life:3400, weight:5},
    {value:3, r:13, color:"#c9a15a", life:2500, weight:3},
    {value:7, r:17, color:"#b54b3d", life:1700, weight:1}
  ];

  function pickTier(){

    const total =
      SOUL_TIERS.reduce((s,t) => s + t.weight,0);

    let roll = Math.random() * total;

    for(const t of SOUL_TIERS){

      if(roll < t.weight){
        return t;
      }

      roll -= t.weight;
    }

    return SOUL_TIERS[0];
  }

  let souls = [];
  let score = 0;
  let timeLeft = ROUND_SECONDS;
  let running = false;
  let rafId = null;
  let spawnTimer = null;
  let tickTimer = null;

  function spawnSoul(){

    if(souls.length >= 9){
      return;
    }

    const tier = pickTier();
    const angle = Math.random() * Math.PI * 2;
    const speed = 0.4 + Math.random() * 0.9;

    souls.push({
      x: tier.r + Math.random() * (canvas.width - tier.r*2),
      y: tier.r + Math.random() * (canvas.height - tier.r*2),
      vx: Math.cos(angle) * speed,
      vy: Math.sin(angle) * speed,
      r: tier.r,
      color: tier.color,
      value: tier.value,
      born: Date.now(),
      life: tier.life + Math.random() * 700
    });
  }

  function loop(){

    if(!running){
      return;
    }

    ctx.clearRect(0,0,canvas.width,canvas.height);

    const now = Date.now();

    souls.forEach(s => {

      s.x += s.vx;
      s.y += s.vy;

      if(s.x < s.r || s.x > canvas.width - s.r){ s.vx *= -1; }
      if(s.y < s.r || s.y > canvas.height - s.r){ s.vy *= -1; }

      const age = now - s.born;
      const fade = Math.max(0,1 - age / s.life);

      ctx.globalAlpha = 0.25 + fade * 0.75;
      ctx.fillStyle = s.color;
      ctx.beginPath();
      ctx.arc(s.x,s.y,s.r,0,Math.PI*2);
      ctx.fill();
      ctx.globalAlpha = 1;

      ctx.fillStyle = "#0b0a0d";
      ctx.font = "10px 'IBM Plex Mono', monospace";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(s.value,s.x,s.y);
    });

    souls = souls.filter(s => now - s.born < s.life);

    rafId = requestAnimationFrame(loop);
  }

  function handleClick(e){

    if(!running){
      return;
    }

    const rect = canvas.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;
    const x = clientX - rect.left;
    const y = clientY - rect.top;

    for(let i = souls.length - 1; i >= 0; i--){

      const s = souls[i];

      if(Math.hypot(x - s.x,y - s.y) <= s.r){

        score += s.value;
        souls.splice(i,1);

        document.getElementById("soul-score").textContent = score;
        break;
      }
    }
  }

  canvas.addEventListener("click",handleClick);
  canvas.addEventListener("touchstart",e => { handleClick(e); e.preventDefault(); },{passive:false});

  function startSoulTracker(){

    souls = [];
    score = 0;
    timeLeft = ROUND_SECONDS;
    running = true;

    document.getElementById("soul-score").textContent = "0";
    document.getElementById("soul-time").textContent = ROUND_SECONDS;
    document.getElementById("soul-end").textContent = "";
    document.getElementById("soul-start").disabled = true;
    document.getElementById("soul-start").textContent = "Tracking…";

    clearInterval(spawnTimer);
    clearInterval(tickTimer);

    spawnTimer = setInterval(spawnSoul,550);

    tickTimer = setInterval(() => {

      timeLeft--;

      document.getElementById("soul-time").textContent =
        Math.max(0,timeLeft);

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