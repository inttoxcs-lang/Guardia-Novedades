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

  // =========================
  // DOM
  // =========================
  const gidInput = document.getElementById("gidInput");
  const headerRowInput = document.getElementById("headerRowInput");
  const metricColInput = document.getElementById("metricColInput");
  const reloadBtn = document.getElementById("reloadBtn");
  const applyBtn = document.getElementById("applyBtn");
  const searchInput = document.getElementById("searchInput");
  const cardsGrid = document.getElementById("cardsGrid");

  // =========================
  // STATE
  // =========================
  let dayCards = [];

  // =========================
  // HELPERS
  // =========================
  function escapeHtml(str) {
    return String(str ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function extractSpreadsheetId(url) {
    const m = String(url).match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return m ? m[1] : "";
  }

  function startOfDay(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function addDays(d, days) {
    const x = new Date(d);
    x.setDate(x.getDate() + days);
    return x;
  }

  // =========================
  // FETCH CSV
  // =========================
  async function fetchCsv(spreadsheetId, gid) {
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=${gid}`;
    const res = await fetch(url, { cache: "no-store" });
    const text = await res.text();

    if (!res.ok || text.startsWith("<")) {
      throw new Error("El Google Sheet no es público o no está publicado.");
    }
    return text;
  }

  // =========================
  // CSV → MATRIX
  // =========================
  function parseCsv(text) {
    const rows = [];
    let row = [];
    let field = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      const n = text[i + 1];

      if (c === '"' && inQuotes && n === '"') {
        field += '"';
        i++;
        continue;
      }
      if (c === '"') {
        inQuotes = !inQuotes;
        continue;
      }
      if (!inQuotes && c === ",") {
        row.push(field);
        field = "";
        continue;
      }
      if (!inQuotes && c === "\n") {
        row.push(field);
        rows.push(row);
        row = [];
        field = "";
        continue;
      }
      if (c !== "\r") field += c;
    }

    row.push(field);
    rows.push(row);
    return rows.map(r => r.map(c => String(c ?? "").trim()));
  }

  // =========================
  // FECHAS
  // =========================
  function parseDate(label) {
    const m = String(label).match(/(\d{1,2})\/(\d{1,2})\/?(\d{2,4})?/);
    if (!m) return null;

    const d = Number(m[1]);
    const mo = Number(m[2]) - 1;
    const y = m[3] ? Number(m[3]) : new Date().getFullYear();
    const date = new Date(y < 100 ? 2000 + y : y, mo, d);
    return isNaN(date) ? null : startOfDay(date);
  }

  // =========================
  // BUILD CARDS
  // =========================
  function buildCards(matrix, headerRow, metricCol) {
    const h = headerRow - 1;
    const m = metricCol - 1;
    const header = matrix[h];

    const cols = [];
    header.forEach((label, c) => {
      if (c === m) return;
      const d = parseDate(label);
      if (d) cols.push({ c, d });
    });

    const rows = matrix.slice(h + 1);

    return cols.map(({ c, d }) => {
      const metrics = [];
      let inasist = 0;

      rows.forEach(r => {
        const name = r[m];
        if (!name) return;
        const val = r[c];
        metrics.push({ name, value: val });

        if (name.toLowerCase().includes("inasist")) {
          const n = Number(String(val).replace(",", "."));
          if (!isNaN(n)) inasist += n;
        }
      });

      return { date: d, metrics, inasist };
    });
  }

  // =========================
  // WINDOW & SORT
  // =========================
  function windowCards(cards) {
    const today = startOfDay(new Date());
    const valid = cards.filter(c => c.date <= today);
    const anchor = valid.reduce((a, b) => (b.date > a ? b.date : a), valid[0]?.date);
    const min = addDays(anchor, -DAYS_BACK);

    return valid
      .filter(c => c.date >= min && c.date <= anchor)
      .sort((a, b) => b.date - a.date);
  }

  // =========================
  // RENDER (SIN HEADER)
  // =========================
  function render() {
    cardsGrid.innerHTML = "";
    const q = searchInput.value.toLowerCase();

    windowCards(dayCards)
      .filter(c =>
        !q ||
        c.metrics.some(m =>
          `${m.name} ${m.value}`.toLowerCase().includes(q)
        )
      )
      .forEach(card => {
        const el = document.createElement("article");
        el.className = "card";

        el.innerHTML = `
          <div class="card-body">
            <div class="kpi-row">
              <div class="kpi"><div class="k">Con valor</div><div class="v">${card.metrics.filter(m => m.value).length}</div></div>
              <div class="kpi"><div class="k">Total</div><div class="v">${card.metrics.length}</div></div>
              <div class="kpi"><div class="k">Inasistencias</div><div class="v">${card.inasist}</div></div>
            </div>

            <div class="table table-scroll">
              ${card.metrics.map(m => `
                <div class="row">
                  <div class="key">${escapeHtml(m.name)}</div>
                  <div class="val">${escapeHtml(m.value || "—")}</div>
                </div>
              `).join("")}
            </div>
          </div>
        `;

        cardsGrid.appendChild(el);
      });
  }

  // =========================
  // LOAD
  // =========================
  async function load() {
    const id = extractSpreadsheetId(SHEET_URL);
    const csv = await fetchCsv(id, gidInput.value || DEFAULT_GID);
    const matrix = parseCsv(csv);

    dayCards = buildCards(
      matrix,
      Number(headerRowInput.value || DEFAULT_HEADER_ROW),
      Number(metricColInput.value || DEFAULT_METRIC_COL)
    );

    render();
  }

  reloadBtn.onclick = load;
  applyBtn.onclick = load;
  searchInput.oninput = render;

  load();
})();
