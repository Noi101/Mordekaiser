/* ============================================================
   GAMES — WHACK A VAYNE
   ============================================================ */

let whackTimer = null;
let whackSpawnTimer = null;
let whackScore = 0;
let whackMisses = 0;
let whackTime = 30;
let whackRunning = false;
let whackActiveHole = -1;

function renderGames(){
  const el = document.getElementById("tab-games");
  if(!el) return;

  const vayneImage =
    `${DDRAGON_BASE}/cdn/img/champion/loading/Vayne_0.jpg`;

  el.innerHTML = `
    <div class="games-list">
      <h2 class="games-title">Games</h2>

      <div class="game-list-grid">

        <button
          class="game-card active"
          type="button"
          data-game="whack-vayne"
        >
          <span class="game-card-icon">⚒</span>

          <span>
            <strong>Whack A Vayne</strong>
            <small>
              30 seconds · click Vayne before she disappears.
            </small>
          </span>
        </button>

        <div
          class="game-card disabled"
          aria-disabled="true"
        >
          <span class="game-card-icon">☠</span>

          <span>
            <strong>Mordekaiser Arena</strong>
            <small>Coming soon.</small>
          </span>
        </div>

        <div
          class="game-card disabled"
          aria-disabled="true"
        >
          <span class="game-card-icon">♜</span>

          <span>
            <strong>Realm of Death</strong>
            <small>Coming soon.</small>
          </span>
        </div>

      </div>
    </div>

    <section
      class="wav-game"
      aria-label="Whack A Vayne"
    >

      <h2>Whack A Vayne</h2>

      <p class="wav-intro">
        Vayne has entered the Realm of Death.
        Bonk her with your Morde mace.
        Hit as many as you can before the timer reaches zero.
      </p>

      <div class="wav-hud">

        <div class="wav-stat">
          <span class="wav-stat-label">Score</span>
          <span
            class="wav-stat-value"
            id="wav-score"
          >
            0
          </span>
        </div>

        <div class="wav-stat">
          <span class="wav-stat-label">Misses</span>
          <span
            class="wav-stat-value"
            id="wav-misses"
          >
            0
          </span>
        </div>

        <div class="wav-stat">
          <span class="wav-stat-label">Time</span>
          <span
            class="wav-stat-value"
            id="wav-time"
          >
            30
          </span>
        </div>

        <button
          class="wav-start-btn"
          id="wav-start"
          type="button"
        >
          Start Game
        </button>

      </div>

      <div
        class="wav-board"
        id="wav-board"
      >

        ${Array.from({length:9}, (_,i) => `
          <div
            class="wav-hole"
            data-hole="${i}"
          >

            <div class="wav-mound"></div>

            <button
              class="wav-mole"
              type="button"
              aria-label="Vayne"
              tabindex="-1"
            >
              <img
                src="${vayneImage}"
                alt="Vayne"
              >
            </button>

          </div>
        `).join("")}

      </div>

      <div
        class="wav-result"
        id="wav-result"
        aria-live="polite"
      >
        Press Start Game to begin.
      </div>

    </section>
  `;

  const start =
    document.getElementById("wav-start");

  start.addEventListener(
    "click",
    startWhackGame
  );

  document
    .querySelectorAll(".wav-mole")
    .forEach(mole => {

      mole.addEventListener(
        "click",
        e => {

          e.preventDefault();

          if(!whackRunning)
            return;

          const hole =
            mole.closest(".wav-hole");

          if(!hole.classList.contains("up"))
            return;

          whackScore++;

          hole.classList.remove("up");
          hole.classList.add("hit");

          whackActiveHole = -1;

          updateWhackHud();

          setTimeout(
            () => hole.classList.remove("hit"),
            180
          );

          scheduleWhackSpawn(160);
        }
      );

    });
}


function updateWhackHud(){

  const score =
    document.getElementById("wav-score");

  const misses =
    document.getElementById("wav-misses");

  const time =
    document.getElementById("wav-time");

  if(score)
    score.textContent = whackScore;

  if(misses)
    misses.textContent = whackMisses;

  if(time)
    time.textContent = whackTime;
}


function scheduleWhackSpawn(
  delay = 420
){

  clearTimeout(
    whackSpawnTimer
  );

  whackSpawnTimer =
    setTimeout(
      spawnVayne,
      delay
    );
}


function spawnVayne(){

  if(!whackRunning)
    return;

  const holes =
    [...document.querySelectorAll(".wav-hole")];

  if(!holes.length)
    return;

  // If the previous Vayne was still visible,
  // count it as a miss.
  if(whackActiveHole >= 0){

    const previous =
      holes[whackActiveHole];

    if(
      previous &&
      previous.classList.contains("up")
    ){

      whackMisses++;

      previous.classList.remove("up");

      updateWhackHud();
    }
  }

  let next =
    Math.floor(
      Math.random() * holes.length
    );

  // Don't immediately use the same hole.
  if(
    next === whackActiveHole &&
    holes.length > 1
  ){
    next =
      (next + 1) % holes.length;
  }

  whackActiveHole = next;

  holes[next].classList.add("up");

  scheduleWhackSpawn(
    650 + Math.random() * 550
  );
}


function startWhackGame(){

  clearInterval(whackTimer);
  clearTimeout(whackSpawnTimer);

  whackScore = 0;
  whackMisses = 0;
  whackTime = 30;

  whackRunning = true;
  whackActiveHole = -1;

  const start =
    document.getElementById("wav-start");

  const result =
    document.getElementById("wav-result");

  document
    .querySelectorAll(".wav-hole")
    .forEach(h => {
      h.classList.remove(
        "up",
        "hit"
      );
    });

  start.disabled = true;
  start.textContent = "Game Running…";

  result.textContent =
    "Get her!";

  updateWhackHud();

  spawnVayne();

  whackTimer =
    setInterval(() => {

      whackTime--;

      updateWhackHud();

      if(whackTime <= 0){
        endWhackGame();
      }

    }, 1000);
}


function endWhackGame(){

  clearInterval(whackTimer);
  clearTimeout(whackSpawnTimer);

  whackRunning = false;

  document
    .querySelectorAll(".wav-hole")
    .forEach(h => {
      h.classList.remove("up");
    });

  const start =
    document.getElementById("wav-start");

  const result =
    document.getElementById("wav-result");

  if(start){

    start.disabled = false;

    start.textContent =
      "Play Again";
  }

  if(result){

    result.textContent =
      `Time! You whacked Vayne ${
        whackScore
      } time${
        whackScore === 1 ? "" : "s"
      }.`;
  }
}