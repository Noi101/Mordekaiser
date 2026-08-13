/* ============================================================
   GAMES

   Split out of the main sync/render logic (app.js) into its own
   file since this section isn't about spreadsheet syncing at
   all — it's a self-contained arcade. app.js must load before
   this file: these games call back into helpers it defines
   (championImage, championSquareImage, findRuneData, ddragon*,
   normalizeName, escapeHtml, champions, itemizationFlat).
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


function launchGame(id){

  const menu =
    document.getElementById("games-menu");

  const stage =
    document.getElementById("game-stage");

  menu.style.display = "none";
  stage.style.display = "block";

  stage.innerHTML =
    `
      <button class="wav-btn back-btn" id="game-back">
        ← Back to Games
      </button>

      <div id="game-inner"></div>
    `;

  document
    .getElementById("game-back")
    .addEventListener("click",showGamesMenu);

  const inner =
    document.getElementById("game-inner");

  switch(id){
    case "whack":  setupWhackAVayneGame(inner); break;
    case "dodge":  setupDodgeGame(inner); break;
    case "combo":  setupComboGame(inner); break;
    case "cs":     setupCsGame(inner); break;
    case "guess":  setupGuessGame(inner); break;
    case "trivia": setupTriviaGame(inner); break;
    case "runes":  setupRuneBuilderGame(inner); break;
    case "items":  setupItemSpeedGame(inner); break;
    case "memory": setupMemoryGame(inner); break;
    case "runner": setupRunnerGame(inner); break;
    case "souls":  setupSoulTrackerGame(inner); break;
  }
}


function setupWhackAVayneGame(container){

  const storedBest =
    parseInt(
      localStorage.getItem("wavBest") || "0",
      10
    );

  container.innerHTML =
    `
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

  gameActiveCleanup = () => {

    if(wavState){
      clearInterval(wavState.tickTimer);
      clearTimeout(wavState.spawnTimer);
      wavState = null;
    }
  };
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
   REALM OF DEATH DODGE
   ============================================================ */

function setupDodgeGame(container){

  const storedBest =
    parseFloat(localStorage.getItem("dodgeBest") || "0");

  container.innerHTML =
    `
      <div class="game-card">
        <h3>Realm of Death Dodge</h3>
        <p>Move your mouse (or finger) to steer. Survive the spectral bolts as long as you can.</p>
        <div class="wav-hud">
          <div>Time <span id="dodge-time">0.0</span>s</div>
          <div>Best <span id="dodge-best">${storedBest.toFixed(1)}</span>s</div>
        </div>
        <canvas id="dodge-canvas" width="340" height="340"></canvas>
        <button class="wav-btn" id="dodge-start">Start</button>
        <div class="wav-end" id="dodge-end"></div>
      </div>
    `;

  const canvas = document.getElementById("dodge-canvas");
  const ctx = canvas.getContext("2d");

  let player = {x:170,y:170,r:9};
  let bolts = [];
  let running = false;
  let startTime = 0;
  let rafId = null;
  let spawnTimer = null;

  function pointerMove(e){
    const rect = canvas.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;
    player.x = Math.min(canvas.width,Math.max(0,clientX - rect.left));
    player.y = Math.min(canvas.height,Math.max(0,clientY - rect.top));
  }

  canvas.addEventListener("mousemove",pointerMove);
  canvas.addEventListener("touchmove",e => { pointerMove(e); e.preventDefault(); },{passive:false});

  function spawnBolt(){

    const edge = Math.floor(Math.random() * 4);
    let x,y;

    if(edge === 0){ x = Math.random()*canvas.width; y = -10; }
    else if(edge === 1){ x = canvas.width+10; y = Math.random()*canvas.height; }
    else if(edge === 2){ x = Math.random()*canvas.width; y = canvas.height+10; }
    else{ x = -10; y = Math.random()*canvas.height; }

    const angle = Math.atan2(170-y,170-x) + (Math.random()-0.5) * 1.1;
    const elapsedSec = (Date.now()-startTime) / 1000;
    const speed = 1.5 + Math.random()*1.4 + elapsedSec/10;

    bolts.push({x,y,vx:Math.cos(angle)*speed,vy:Math.sin(angle)*speed,r:6});
  }

  function loop(){

    if(!running){
      return;
    }

    ctx.clearRect(0,0,canvas.width,canvas.height);

    ctx.fillStyle = "rgba(36,27,46,0.4)";
    ctx.beginPath();
    ctx.arc(170,170,168,0,Math.PI*2);
    ctx.fill();

    bolts.forEach(b => {
      b.x += b.vx;
      b.y += b.vy;

      ctx.fillStyle = "#4fe3b0";
      ctx.beginPath();
      ctx.arc(b.x,b.y,b.r,0,Math.PI*2);
      ctx.fill();
    });

    bolts = bolts.filter(
      b => b.x > -20 && b.x < canvas.width+20 && b.y > -20 && b.y < canvas.height+20
    );

    ctx.fillStyle = "#c9a15a";
    ctx.beginPath();
    ctx.arc(player.x,player.y,player.r,0,Math.PI*2);
    ctx.fill();

    const hit =
      bolts.some(b => Math.hypot(b.x-player.x,b.y-player.y) < b.r+player.r);

    if(hit){
      endDodge();
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
    container.innerHTML =
      `<div class="game-card"><p>Couldn't load ability data. Try again in a moment.</p></div>`;
    return;
  }

  const storedBest =
    parseInt(localStorage.getItem("comboBest") || "0",10);

  container.innerHTML =
    `
      <div class="game-card">
        <h3>Combo Trainer</h3>
        <p>Watch the sequence flash, then repeat it back. Gets one ability longer each round.</p>
        <div class="wav-hud">
          <div>Round <span id="combo-round">0</span></div>
          <div>Best <span id="combo-best">${storedBest}</span></div>
        </div>
        <div class="combo-icons" id="combo-icons">
          ${
            spells.map(
              s => `
                <button class="combo-icon" data-key="${s.key}" title="${escapeHtml(s.name)}">
                  <img src="${s.icon}" alt="${escapeHtml(s.key)}">
                  <span>${s.key}</span>
                </button>
              `
            ).join("")
          }
        </div>
        <button class="wav-btn" id="combo-start">Start</button>
        <div class="wav-end" id="combo-end"></div>
      </div>
    `;

  let sequence = [];
  let playerIndex = 0;
  let accepting = false;

  const buttons =
    Array.from(container.querySelectorAll(".combo-icon"));

  function flash(key,cls){
    const btn = buttons.find(b => b.dataset.key === key);
    if(!btn) return;
    btn.classList.add(cls);
    setTimeout(() => btn.classList.remove(cls),280);
  }

  async function playSequence(){

    accepting = false;

    for(const key of sequence){
      await new Promise(r => setTimeout(r,320));
      flash(key,"flash-show");
      await new Promise(r => setTimeout(r,320));
    }

    playerIndex = 0;
    accepting = true;
  }

  function nextRound(){

    const keys = ["Q","W","E","R"];
    sequence.push(keys[Math.floor(Math.random()*4)]);

    document.getElementById("combo-round").textContent = sequence.length;

    playSequence();
  }

  function endCombo(){

    accepting = false;

    const score = sequence.length - 1;
    const best = parseInt(localStorage.getItem("comboBest") || "0",10);

    if(score > best){
      localStorage.setItem("comboBest",score);
    }

    document.getElementById("combo-best").textContent = Math.max(best,score);

    document.getElementById("combo-end").textContent =
      `Broke the combo at round ${sequence.length}. Final: ${score} correct.`;

    const btn = document.getElementById("combo-start");
    btn.disabled = false;
    btn.textContent = "Try again";
  }

  buttons.forEach(btn => {

    btn.addEventListener("click",() => {

      if(!accepting){
        return;
      }

      const key = btn.dataset.key;

      if(key === sequence[playerIndex]){

        flash(key,"flash-correct");
        playerIndex++;

        if(playerIndex === sequence.length){
          setTimeout(nextRound,500);
        }

      }else{

        flash(key,"flash-wrong");
        endCombo();
      }

    });

  });

  document.getElementById("combo-start").addEventListener("click",() => {

    sequence = [];

    document.getElementById("combo-end").textContent = "";
    document.getElementById("combo-start").disabled = true;
    document.getElementById("combo-start").textContent = "Watching…";

    nextRound();
  });

  gameActiveCleanup = () => { accepting = false; };
}


/* ============================================================
   CS PRACTICE
   ============================================================ */

function setupCsGame(container){

  const storedBest =
    parseInt(localStorage.getItem("csBest") || "0",10);

  const ROUNDS = 10;

  container.innerHTML =
    `
      <div class="game-card">
        <h3>CS Practice</h3>
        <p>Click "Last Hit" the instant you think the minion's about to die from your allies' damage — not before, not after.</p>
        <div class="wav-hud">
          <div>Hits <span id="cs-hits">0</span>/<span id="cs-total">${ROUNDS}</span></div>
          <div>Best <span id="cs-best">${storedBest}</span></div>
        </div>
        <div class="cs-bar-wrap"><div class="cs-bar" id="cs-bar"></div></div>
        <div class="cs-hp" id="cs-hp">— HP —</div>
        <button class="wav-btn" id="cs-hit" disabled>Last Hit!</button>
        <button class="wav-btn" id="cs-start">Start</button>
        <div class="wav-end" id="cs-end"></div>
      </div>
    `;

  let round = 0;
  let hits = 0;
  let hp = 0;
  let maxHp = 0;
  let tickTimer = null;
  let dead = false;
  let running = false;

  function updateBar(){
    const pct = Math.max(0,(hp/maxHp)*100);
    document.getElementById("cs-bar").style.width = pct + "%";
    document.getElementById("cs-hp").textContent = Math.max(0,Math.round(hp)) + " HP";
  }

  function afterRound(success){

    round++;

    if(success){
      hits++;
    }

    document.getElementById("cs-hits").textContent = hits;

    if(round >= ROUNDS){
      finishCs();
    }else{
      nextMinion();
    }
  }

  function nextMinion(){

    maxHp = 90 + Math.random()*140;
    hp = maxHp;
    dead = false;

    updateBar();

    clearInterval(tickTimer);

    tickTimer = setInterval(() => {

      hp -= 4 + Math.random()*10;

      if(hp <= 0){
        hp = 0;
        dead = true;
        clearInterval(tickTimer);
        setTimeout(() => afterRound(false),400);
      }

      updateBar();

    },260);
  }

  function finishCs(){

    running = false;
    clearInterval(tickTimer);

    const best = parseInt(localStorage.getItem("csBest") || "0",10);

    if(hits > best){
      localStorage.setItem("csBest",hits);
    }

    document.getElementById("cs-best").textContent = Math.max(best,hits);
    document.getElementById("cs-end").textContent = `Final: ${hits}/${ROUNDS} minions last-hit.`;

    document.getElementById("cs-start").disabled = false;
    document.getElementById("cs-start").textContent = "Try again";
    document.getElementById("cs-hit").disabled = true;
  }

  document.getElementById("cs-hit").addEventListener("click",() => {

    if(!running || dead){
      return;
    }

    const success = hp > 0 && hp <= maxHp * 0.18;

    clearInterval(tickTimer);
    afterRound(success);
  });

  document.getElementById("cs-start").addEventListener("click",() => {

    round = 0;
    hits = 0;
    running = true;

    document.getElementById("cs-hits").textContent = "0";
    document.getElementById("cs-end").textContent = "";
    document.getElementById("cs-start").disabled = true;
    document.getElementById("cs-start").textContent = "Practicing…";
    document.getElementById("cs-hit").disabled = false;

    nextMinion();
  });

  gameActiveCleanup = () => { running = false; clearInterval(tickTimer); };
}


/* ============================================================
   GUESS THE CHAMPION
   ============================================================ */

function setupGuessGame(container){

  const storedBest =
    parseInt(localStorage.getItem("guessBest") || "0",10);

  const ROUNDS = 10;
  const pool = Object.values(ddragonChampions);

  if(!pool.length){
    container.innerHTML =
      `<div class="game-card"><p>Champion data hasn't loaded yet — try again in a moment.</p></div>`;
    return;
  }

  container.innerHTML =
    `
      <div class="game-card">
        <h3>Guess the Champion</h3>
        <p>The splash art sharpens over a few seconds. Type the champion's name before it's fully clear.</p>
        <div class="wav-hud">
          <div>Score <span id="guess-score">0</span></div>
          <div>Round <span id="guess-round">0</span>/<span id="guess-total">${ROUNDS}</span></div>
          <div>Best <span id="guess-best">${storedBest}</span></div>
        </div>
        <div class="guess-portrait-wrap">
          <img id="guess-portrait" class="guess-portrait" alt="">
        </div>
        <input type="text" id="guess-input" placeholder="Champion name…" autocomplete="off" disabled>
        <button class="wav-btn" id="guess-start">Start</button>
        <div class="wav-end" id="guess-end"></div>
      </div>
    `;

  let round = 0;
  let score = 0;
  let current = null;
  let blurTimer = null;
  let running = false;

  function nextRound(){

    round++;
    document.getElementById("guess-round").textContent = round;

    if(round > ROUNDS){
      finishGuess();
      return;
    }

    current = pool[Math.floor(Math.random()*pool.length)];

    const img = document.getElementById("guess-portrait");
    img.src = `${DDRAGON_BASE}/cdn/img/champion/loading/${current.id}_0.jpg`;
    img.style.filter = "blur(22px)";

    const input = document.getElementById("guess-input");
    input.value = "";
    input.disabled = false;
    input.focus();

    let blur = 22;
    clearInterval(blurTimer);

    blurTimer = setInterval(() => {

      blur -= 1;
      img.style.filter = `blur(${Math.max(0,blur)}px)`;

      if(blur <= 0){
        clearInterval(blurTimer);
        checkGuess(true);
      }

    },350);
  }

  function checkGuess(timedOut){

    if(!running){
      return;
    }

    const input = document.getElementById("guess-input");
    const guess = normalizeName(input.value);
    const answer = normalizeName(current.name);

    clearInterval(blurTimer);
    document.getElementById("guess-portrait").style.filter = "blur(0px)";
    input.disabled = true;

    if(!timedOut && guess.length >= 3 && answer.includes(guess)){
      score++;
      document.getElementById("guess-score").textContent = score;
    }

    setTimeout(nextRound,900);
  }

  function finishGuess(){

    running = false;

    const best = parseInt(localStorage.getItem("guessBest") || "0",10);

    if(score > best){
      localStorage.setItem("guessBest",score);
    }

    document.getElementById("guess-best").textContent = Math.max(best,score);
    document.getElementById("guess-end").textContent = `Final score: ${score}/${ROUNDS}.`;

    document.getElementById("guess-start").disabled = false;
    document.getElementById("guess-start").textContent = "Play again";
  }

  document.getElementById("guess-input").addEventListener("keydown",e => {
    if(e.key === "Enter" && running){
      checkGuess(false);
    }
  });

  document.getElementById("guess-start").addEventListener("click",() => {

    round = 0;
    score = 0;
    running = true;

    document.getElementById("guess-score").textContent = "0";
    document.getElementById("guess-end").textContent = "";
    document.getElementById("guess-start").disabled = true;
    document.getElementById("guess-start").textContent = "Guessing…";

    nextRound();
  });

  gameActiveCleanup = () => { running = false; clearInterval(blurTimer); };
}


/* ============================================================
   MATCHUP TRIVIA
   ============================================================ */

function setupTriviaGame(container){

  const storedBest =
    parseInt(localStorage.getItem("triviaBest") || "0",10);

  const pool =
    champions.filter(
      c =>
        c.difficulty.early !== null ||
        c.difficulty.mid !== null ||
        c.difficulty.late !== null ||
        c.difficulty.overall !== null
    );

  if(!pool.length){
    container.innerHTML =
      `<div class="game-card"><p>No difficulty ratings found in the matchup tab yet.</p></div>`;
    return;
  }

  const ROUNDS = Math.min(10,pool.length);

  container.innerHTML =
    `
      <div class="game-card">
        <h3>Matchup Trivia</h3>
        <p>Quick quiz on your own written matchup ratings.</p>
        <div class="wav-hud">
          <div>Score <span id="trivia-score">0</span></div>
          <div>Round <span id="trivia-round">0</span>/<span id="trivia-total">${ROUNDS}</span></div>
          <div>Best <span id="trivia-best">${storedBest}</span></div>
        </div>
        <div class="trivia-question" id="trivia-question"></div>
        <div class="rune-choices" id="trivia-options"></div>
        <button class="wav-btn" id="trivia-start">Start</button>
        <div class="wav-end" id="trivia-end"></div>
      </div>
    `;

  let round = 0;
  let score = 0;
  let answered = false;

  function nextQuestion(){

    round++;
    answered = false;

    document.getElementById("trivia-round").textContent = round;

    if(round > ROUNDS){
      finishTrivia();
      return;
    }

    const champ = pool[Math.floor(Math.random()*pool.length)];

    const stages =
      ["early","mid","late","overall"].filter(
        s => champ.difficulty[s] !== null
      );

    const stage = stages[Math.floor(Math.random()*stages.length)];
    const correct = champ.difficulty[stage];

    document.getElementById("trivia-question").textContent =
      `What's Mordekaiser's ${stage} difficulty vs. ${champ.name}?`;

    const options = new Set([correct]);

    while(options.size < 4){
      options.add(1 + Math.floor(Math.random()*5));
    }

    const shuffled =
      Array.from(options).sort(() => Math.random()-0.5);

    document.getElementById("trivia-options").innerHTML =
      shuffled.map(
        o => `<button class="rune-choice trivia-opt" data-value="${o}"><span>${o}/5</span></button>`
      ).join("");

    document.getElementById("trivia-options").querySelectorAll(".trivia-opt").forEach(btn => {

      btn.addEventListener("click",() => {

        if(answered){
          return;
        }

        answered = true;

        const chosen = parseInt(btn.dataset.value,10);

        if(chosen === correct){
          score++;
          btn.classList.add("rune-used");
          document.getElementById("trivia-score").textContent = score;
        }else{
          btn.classList.add("trivia-wrong");
        }

        setTimeout(nextQuestion,700);
      });

    });
  }

  function finishTrivia(){

    const best = parseInt(localStorage.getItem("triviaBest") || "0",10);

    if(score > best){
      localStorage.setItem("triviaBest",score);
    }

    document.getElementById("trivia-best").textContent = Math.max(best,score);
    document.getElementById("trivia-end").textContent = `Final score: ${score}/${ROUNDS}.`;
    document.getElementById("trivia-options").innerHTML = "";

    document.getElementById("trivia-start").disabled = false;
    document.getElementById("trivia-start").textContent = "Play again";
  }

  document.getElementById("trivia-start").addEventListener("click",() => {

    round = 0;
    score = 0;

    document.getElementById("trivia-score").textContent = "0";
    document.getElementById("trivia-end").textContent = "";
    document.getElementById("trivia-start").disabled = true;
    document.getElementById("trivia-start").textContent = "Quizzing…";

    nextQuestion();
  });

  gameActiveCleanup = () => {};
}


/* ============================================================
   RUNE PATH BUILDER
   ============================================================ */

const MORDE_RUNE_PATH = [
  "Conqueror","Triumph","Legend: Alacrity","Last Stand","Shield Bash","Bone Plating"
];


function setupRuneBuilderGame(container){

  const storedBest =
    parseInt(localStorage.getItem("runeBuilderBest") || "0",10);

  const runeData =
    MORDE_RUNE_PATH.map(name => {
      const found = findRuneData(name);
      return {name:name,icon:found ? found.icon : null};
    }).filter(r => r.icon);

  if(runeData.length < 4){
    container.innerHTML =
      `<div class="game-card"><p>Rune data hasn't loaded yet — try again in a moment.</p></div>`;
    return;
  }

  container.innerHTML =
    `
      <div class="game-card">
        <h3>Rune Path Builder</h3>
        <p>Click Mordekaiser's core runes in the correct order, from keystone to the last secondary rune.</p>
        <div class="wav-hud">
          <div>Correct <span id="rune-progress">0</span>/<span id="rune-total">${runeData.length}</span></div>
          <div>Best streak <span id="rune-best">${storedBest}</span></div>
        </div>
        <div class="rune-slots" id="rune-slots">
          ${runeData.map((_,i) => `<div class="rune-slot" id="rune-slot-${i}"></div>`).join("")}
        </div>
        <div class="rune-choices" id="rune-choices"></div>
        <button class="wav-btn" id="rune-start">Start</button>
        <div class="wav-end" id="rune-end"></div>
      </div>
    `;

  let progress = 0;
  let running = false;

  function shuffleChoices(){

    const shuffled = [...runeData].sort(() => Math.random()-0.5);

    document.getElementById("rune-choices").innerHTML =
      shuffled.map(
        r => `
          <button class="rune-choice" data-name="${escapeHtml(r.name)}">
            <img src="${r.icon}" alt="">
            <span>${escapeHtml(r.name)}</span>
          </button>
        `
      ).join("");

    document.getElementById("rune-choices").querySelectorAll(".rune-choice").forEach(btn => {
      btn.addEventListener("click",() => pickRune(btn));
    });
  }

  function pickRune(btn){

    if(!running){
      return;
    }

    const name = btn.dataset.name;
    const expected = runeData[progress].name;

    if(name === expected){

      const slot = document.getElementById(`rune-slot-${progress}`);
      slot.innerHTML = `<img src="${btn.querySelector("img").src}" alt="">`;
      slot.classList.add("filled");

      btn.disabled = true;
      btn.classList.add("rune-used");

      progress++;
      document.getElementById("rune-progress").textContent = progress;

      if(progress === runeData.length){
        finishRuneBuild(true);
      }

    }else{
      finishRuneBuild(false);
    }
  }

  function finishRuneBuild(success){

    running = false;

    const best = parseInt(localStorage.getItem("runeBuilderBest") || "0",10);

    if(progress > best){
      localStorage.setItem("runeBuilderBest",progress);
    }

    document.getElementById("rune-best").textContent = Math.max(best,progress);

    document.getElementById("rune-end").textContent =
      success
        ? "Full build assembled correctly."
        : `Wrong pick at step ${progress+1}.`;

    document.getElementById("rune-choices").querySelectorAll(".rune-choice").forEach(b => {
      b.disabled = true;
    });

    document.getElementById("rune-start").disabled = false;
    document.getElementById("rune-start").textContent = "Try again";
  }

  document.getElementById("rune-start").addEventListener("click",() => {

    progress = 0;
    running = true;

    document.getElementById("rune-progress").textContent = "0";
    document.getElementById("rune-end").textContent = "";
    document.getElementById("rune-start").disabled = true;
    document.getElementById("rune-start").textContent = "Building…";

    runeData.forEach((_,i) => {
      const slot = document.getElementById(`rune-slot-${i}`);
      slot.innerHTML = "";
      slot.classList.remove("filled");
    });

    shuffleChoices();
  });

  gameActiveCleanup = () => { running = false; };
}


/* ============================================================
   ITEMIZATION SPEED ROUND
   ============================================================ */

function setupItemSpeedGame(container){

  const storedBest =
    parseInt(localStorage.getItem("itemSpeedBest") || "0",10);

  if(itemizationFlat.length < 4){
    container.innerHTML =
      `<div class="game-card"><p>Not enough labeled items on the Itemization tab yet to quiz on.</p></div>`;
    return;
  }

  const ROUNDS = Math.min(10,itemizationFlat.length);

  container.innerHTML =
    `
      <div class="game-card">
        <h3>Itemization Speed Round</h3>
        <p>Read the note, pick the item it's describing.</p>
        <div class="wav-hud">
          <div>Score <span id="item-score">0</span></div>
          <div>Round <span id="item-round">0</span>/<span id="item-total">${ROUNDS}</span></div>
          <div>Best <span id="item-best">${storedBest}</span></div>
        </div>
        <div class="trivia-question" id="item-question"></div>
        <div class="rune-choices" id="item-choices"></div>
        <button class="wav-btn" id="item-start">Start</button>
        <div class="wav-end" id="item-end"></div>
      </div>
    `;

  let round = 0;
  let score = 0;
  let answered = false;

  function nextRound(){

    round++;
    answered = false;

    document.getElementById("item-round").textContent = round;

    if(round > ROUNDS){
      finishItemGame();
      return;
    }

    const correct =
      itemizationFlat[Math.floor(Math.random()*itemizationFlat.length)];

    document.getElementById("item-question").textContent = correct.text;

    const choices = new Set([correct]);
    let guard = 0;

    while(choices.size < Math.min(4,itemizationFlat.length) && guard < 50){
      choices.add(itemizationFlat[Math.floor(Math.random()*itemizationFlat.length)]);
      guard++;
    }

    const shuffled =
      Array.from(choices).sort(() => Math.random()-0.5);

    document.getElementById("item-choices").innerHTML =
      shuffled.map(
        it => `
          <button class="rune-choice" data-label="${escapeHtml(it.label)}">
            <img src="${it.icon}" alt="">
            <span>${escapeHtml(it.label)}</span>
          </button>
        `
      ).join("");

    document.getElementById("item-choices").querySelectorAll(".rune-choice").forEach(btn => {

      btn.addEventListener("click",() => {

        if(answered){
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

      if(timeLeft <= 0){
        endSoulTracker();
      }

    },1000);

    rafId = requestAnimationFrame(loop);
  }

  function endSoulTracker(){

    running = false;

    cancelAnimationFrame(rafId);
    clearInterval(spawnTimer);
    clearInterval(tickTimer);

    const best = parseInt(localStorage.getItem("soulBest") || "0",10);

    if(score > best){
      localStorage.setItem("soulBest",score);
    }

    document.getElementById("soul-best").textContent = Math.max(best,score);
    document.getElementById("soul-end").textContent = `Collected ${score} soul value.`;

    const btn = document.getElementById("soul-start");
    btn.disabled = false;
    btn.textContent = "Play again";
  }

  document.getElementById("soul-start").addEventListener("click",startSoulTracker);

  gameActiveCleanup = () => {
    running = false;
    cancelAnimationFrame(rafId);
    clearInterval(spawnTimer);
    clearInterval(tickTimer);
  };
}