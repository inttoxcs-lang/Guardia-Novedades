(() => {
  // =========================
  // CONFIG
  // =========================
  const SHEET_URL =
    "https://docs.google.com/spreadsheets/d/120WSaF1Zu6h4-Edid-Yc7GIWKwtQB61GQ3rNexH-MXc/edit?gid=0#gid=0";

  const DEFAULT_GID = "0";
  const DEFAULT_HEADER_ROW = 1;
  const DEFAULT_METRIC_COL = 1;
  const DAYS_BACK = 6;

  // 🔄 Auto refresh (minutos)
  const REFRESH_MINUTES = 5;
  const REFRESH_MS = REFRESH_MINUTES * 60 * 1000;

  const cardsGrid = document.getElementById("cardsGrid");
  let dayCards = [];
  let isLoading = false;
  let refreshTimer = null;

  // =========================
  // HELPERS
  // =========================
  const escapeHtml = s =>
    String(s ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");

  const normalize = s =>
    String(s ?? "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim();

  const startOfDay = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());

  const addDays = (d, n) => {
    const x = new Date(d);
    x.setDate(x.getDate() + n);
    return x;
  };

  const extractLegajos = v =>
    String(v ?? "")
      .split(/[\s,;|]+/)
      .filter(x => /^\d{3,}$/.test(x));

  const uniqueSortLegajos = arr =>
    Array.from(new Set(arr)).sort((a, b) => Number(a) - Number(b));

  const parseNumber = v => {
    const m = String(v ?? "").match(/-?\d+/);
    return m ? Number(m[0]) : 0;
  };

  const formatDate = d =>
    `${String(d.getDate()).padStart(2, "0")}/${String(
      d.getMonth() + 1
    ).padStart(2, "0")}/${d.getFullYear()}`;

  // =========================
  // FETCH + PARSE
  // =========================
  async function fetchCsv() {
    const id = SHEET_URL.match(/\/d\/([^/]+)/)[1];
    const url = `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${DEFAULT_GID}`;
    const r = await fetch(url, { cache: "no-store" });
    const t = await r.text();
    if (t.startsWith("<")) throw "Sheet no público";
    return t;
  }

  function parseCsv(t) {
    return t
      .trim()
      .split("\n")
      .map(r => r.split(",").map(c => c.trim()));
  }

  function parseDate(label) {
    const m = label.match(/(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?/);
    if (!m) return null;
    return startOfDay(
      new Date(m[3] || new Date().getFullYear(), m[2] - 1, m[1])
    );
  }

  // =========================
  // BUILD
  // =========================
  function buildCards(matrix) {
    const header = matrix[0];
    const rows = matrix.slice(1);
    const cols = [];

    header.forEach((h, i) => {
      if (i === 0) return;
      const d = parseDate(h);
      if (d) cols.push({ i, d });
    });

    return cols.map(({ i, d }) => {
      let lineaTM = "—",
        lineaTT = "—",
        legajos = [];
      const table = [];

      for (let r = 0; r < rows.length; r++) {
        const name = rows[r][0];
        const val = rows[r][i];
        if (!name) continue;
        const n = normalize(name);

        if (n === "linea tm") lineaTM = val;
        else if (n === "linea tt") lineaTT = val;
        else if (n === "inasistencias tm") {
          const cant = parseNumber(val);
          let j = r + 1;
          const collected = [];
          while (j < rows.length && !rows[j][0]) {
            collected.push(...extractLegajos(rows[j][i]));
            j++;
          }
          if (cant) legajos = collected;
          r = j - 1;
        } else {
          table.push({ name, val });
        }
      }

      return {
        date: d,
        lineaTM,
        lineaTT,
        legajos: uniqueSortLegajos(legajos),
        table
      };
    });
  }

  function windowCards(cards) {
    const today = startOfDay(new Date());
    const valid = cards.filter(c => c.date <= today);
    const anchor = valid.reduce((a, b) => (b.date > a ? b.date : a), valid[0].date);
    const min = addDays(anchor, -DAYS_BACK);
    return valid
      .filter(c => c.date >= min)
      .sort((a, b) => b.date - a.date);
  }

  // =========================
  // RENDER
  // =========================
  function render() {
    cardsGrid.innerHTML = "";
    const today = startOfDay(new Date());

    windowCards(dayCards).forEach(c => {
      const isToday = c.date.getTime() === today.getTime();
      const hasInasist = c.legajos.length > 0;

      const el = document.createElement("article");
      el.className = `card ${hasInasist ? "card--alert" : "card--ok"} ${
        isToday ? "card--today" : ""
      }`;

      el.innerHTML = `
        <div class="card-body">
          <div class="card-date">
            <div class="date-pill">
              <span class="dot ${hasInasist ? "dot--red" : "dot--green"}"></span>
              ${formatDate(c.date)}
            </div>
          </div>

          <div class="kpi-row">
            <div class="kpi"><div class="k">Línea TM</div><div class="v">${escapeHtml(
              c.lineaTM
            )}</div></div>
            <div class="kpi"><div class="k">Línea TT</div><div class="v">${escapeHtml(
              c.lineaTT
            )}</div></div>
            <div class="kpi"><div class="k">Legajo inasistencia</div><div class="v">${
              c.legajos.length ? c.legajos.join(", ") : "—"
            }</div></div>
          </div>

          <div class="table table-scroll">
            ${c.table
              .map(
                r => `<div class="row"><div class="key">${escapeHtml(
                  r.name
                )}</div><div class="val">${escapeHtml(r.val || "—")}</div></div>`
              )
              .join("")}
          </div>
        </div>
      `;
      cardsGrid.appendChild(el);
    });
  }

  // =========================
  // LOAD + AUTO REFRESH
  // =========================
  async function load() {
    if (isLoading) return;
    isLoading = true;
    try {
      const csv = await fetchCsv();
      dayCards = buildCards(parseCsv(csv));
      render();
    } finally {
      isLoading = false;
    }
  }

  load();
  refreshTimer = setInterval(() => {
    if (document.visibilityState === "visible") load();
  }, REFRESH_MS);
})();
