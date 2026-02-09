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

  console.log("APP VERSION FINAL KPI TM/TT");

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

  function normalize(s) {
    return String(s ?? "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function parseNumber(value) {
    const s = String(value ?? "").trim();
    if (!s) return 0;
    if (/^\d{1,2}:\d{2}$/.test(s)) return 0; // evita horas
    const m = s.match(/-?\d+(?:[.,]\d+)?/);
    if (!m) return 0;
    const n = Number(m[0].replace(",", "."));
    return Number.isFinite(n) ? n : 0;
  }

  // =========================
  // FETCH CSV
  // =========================
  async function fetchCsv(spreadsheetId, gid) {
    const url = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(
      spreadsheetId
    )}/export?format=csv&gid=${encodeURIComponent(gid)}`;

    const res = await fetch(url, { cache: "no-store" });
    const text = await res.text();

    if (!res.ok || text.trim().startsWith("<")) {
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
  // FECHAS (dd/mm o dd/mm/yyyy)
  // =========================
  function parseDate(label) {
    const s = String(label ?? "");
    const m = s.match(/(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/);
    if (!m) return null;

    const dd = Number(m[1]);
    const mm = Number(m[2]) - 1;
    const yRaw = m[3] ? Number(m[3]) : new Date().getFullYear();
    const yyyy = yRaw < 100 ? 2000 + yRaw : yRaw;

    const dt = new Date(yyyy, mm, dd);
    return Number.isNaN(dt.getTime()) ? null : startOfDay(dt);
  }

  // =========================
  // BUILD CARDS
  // =========================
  function buildCards(matrix, headerRow, metricCol) {
    const h = headerRow - 1;
    const mCol = metricCol - 1;
    const header = matrix[h];
    if (!header) throw new Error("Fila de fechas inválida.");

    const cols = [];
    header.forEach((label, c) => {
      if (c === mCol) return;
      const d = parseDate(label);
      if (d) cols.push({ c, d });
    });

    const rows = matrix.slice(h + 1);

    return cols.map(({ c, d }) => {
      const metrics = [];
      let lineaTM = 0;
      let lineaTT = 0;
      let inasist = 0;

      rows.forEach(r => {
        const name = r[mCol];
        if (!name) return;

        const value = r[c];
        metrics.push({ name, value });

        const n = normalize(name);
        if (n.includes("linea tm")) lineaTM = parseNumber(value);
        if (n.includes("linea tt")) lineaTT = parseNumber(value);
        if (n.includes("inasist")) inasist += parseNumber(value);
      });

      return { date: d, metrics, lineaTM, lineaTT, inasist };
    });
  }

  // =========================
  // WINDOW + SORT
  // =========================
  function windowCards(cards) {
    const today = startOfDay(new Date());
    const valid = cards.filter(c => c.date && c.date <= today);
    if (!valid.length) return [];

    let anchor = valid[0].date;
    for (const c of valid) if (c.date > anchor) anchor = c.date;

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
    const q = (searchInput.value || "").toLowerCase().trim();

    windowCards(dayCards)
      .filter(card => {
        if (!q) return true;
        return card.metrics.some(m =>
          `${m.name} ${m.value}`.toLowerCase().includes(q)
        );
      })
      .forEach(card => {
        const el = document.createElement("article");
        el.className = "card";

        el.innerHTML = `
          <div class="card-body">
            <div class="kpi-row">
              <div class="kpi">
                <div class="k">Línea TM</div>
                <div class="v">${escapeHtml(String(card.lineaTM))}</div>
              </div>
              <div class="kpi">
                <div class="k">Línea TT</div>
                <div class="v">${escapeHtml(String(card.lineaTT))}</div>
              </div>
              <div class="kpi">
                <div class="k">Inasistencias</div>
                <div class="v">${escapeHtml(String(card.inasist))}</div>
              </div>
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
    const gid = gidInput?.value || DEFAULT_GID;

    const csv = await fetchCsv(id, gid);
    const matrix = parseCsv(csv);

    dayCards = buildCards(
      matrix,
      Number(headerRowInput?.value || DEFAULT_HEADER_ROW),
      Number(metricColInput?.value || DEFAULT_METRIC_COL)
    );

    render();
  }

  reloadBtn && (reloadBtn.onclick = load);
  applyBtn && (applyBtn.onclick = load);
  searchInput && (searchInput.oninput = render);

  load().catch(err => console.error(err));
})();
