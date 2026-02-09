(() => {
  // =========================
  // CONFIG
  // =========================
  const SHEET_URL =
    "https://docs.google.com/spreadsheets/d/120WSaF1Zu6h4-Edid-Yc7GIWKwtQB61GQ3rNexH-MXc/edit?gid=0#gid=0";

  const DEFAULT_GID = "0";
  const DEFAULT_HEADER_ROW = 1; // 1-based
  const DEFAULT_METRIC_COL = 1; // 1-based
  const DAYS_BACK = 6; // ancla + 6 atrás = 7 tarjetas

  // 🔄 Auto refresh (minutos)
  const REFRESH_MINUTES = 5;
  const REFRESH_MS = REFRESH_MINUTES * 60 * 1000;

  // =========================
  // DOM
  // =========================
  const cardsGrid = document.getElementById("cardsGrid");

  // =========================
  // STATE
  // =========================
  let dayCards = [];
  let isLoading = false;
  let refreshTimer = null;

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

  function normalize(s) {
    return String(s ?? "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function startOfDay(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function addDays(d, n) {
    const x = new Date(d);
    x.setDate(x.getDate() + n);
    return x;
  }

  function formatDate(d) {
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    return `${dd}/${mm}/${d.getFullYear()}`;
  }

  function extractSpreadsheetId(url) {
    const s = String(url || "");
    const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return m ? m[1] : "";
  }

  function parseDate(label) {
    const s = String(label ?? "").trim();
    const m = s.match(/(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/);
    if (!m) return null;

    const dd = Number(m[1]);
    const mm = Number(m[2]) - 1;
    const yRaw = m[3] ? Number(m[3]) : new Date().getFullYear();
    const yyyy = yRaw < 100 ? 2000 + yRaw : yRaw;

    const dt = new Date(yyyy, mm, dd);
    return Number.isNaN(dt.getTime()) ? null : startOfDay(dt);
  }

  function parseNumber(value) {
    const s = String(value ?? "").trim();
    if (!s) return 0;
    if (/^\d{1,2}:\d{2}$/.test(s)) return 0; // horas no cuentan
    const m = s.match(/-?\d+(?:[.,]\d+)?/);
    if (!m) return 0;
    const n = Number(m[0].replace(",", "."));
    return Number.isFinite(n) ? n : 0;
  }

  function extractLegajos(value) {
    // tokens numéricos (>=3 dígitos)
    return String(value ?? "")
      .split(/[\s,;|]+/)
      .map(v => v.trim())
      .filter(v => /^\d{3,}$/.test(v));
  }

  function uniqueSortLegajos(legajos) {
    const uniq = Array.from(new Set((legajos || []).map(String)));
    return uniq.sort((a, b) => Number(a) - Number(b));
  }

  // Ocultar en tabla:
  function shouldHideInTable(metricName) {
    const n = normalize(metricName);
    if (n === "linea tm") return true;
    if (n === "linea tt") return true;
    if (n.includes("legajo") && n.includes("inasist")) return true;
    if (n === "inasistencias tm") return true; // queda solo KPI
    return false;
  }

  // =========================
  // CSV PARSER (robusto con comillas/comas)
  // =========================
  function parseCsvToMatrix(text) {
    const rows = [];
    let curRow = [];
    let cur = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      const next = text[i + 1];

      if (ch === '"' && inQuotes && next === '"') {
        cur += '"';
        i++;
        continue;
      }
      if (ch === '"') {
        inQuotes = !inQuotes;
        continue;
      }
      if (!inQuotes && ch === ",") {
        curRow.push(cur);
        cur = "";
        continue;
      }
      if (!inQuotes && ch === "\n") {
        curRow.push(cur);
        rows.push(curRow);
        curRow = [];
        cur = "";
        continue;
      }
      if (ch !== "\r") cur += ch;
    }

    curRow.push(cur);
    rows.push(curRow);

    return rows.map(r => r.map(c => String(c ?? "").trim()));
  }

  // =========================
  // FETCH CSV (NO CACHE)
  // =========================
  async function fetchCsv(spreadsheetId, gid) {
    const url =
      `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}` +
      `/export?format=csv&gid=${encodeURIComponent(gid)}`;

    const res = await fetch(url, { cache: "no-store" });
    const text = await res.text();

    if (!res.ok || text.trim().startsWith("<")) {
      throw new Error("No se pudo leer el Sheet. Asegurate que esté publicado o público.");
    }
    return text;
  }

  // =========================
  // BUILD CARDS (KPIs + tabla + legajos completos)
  // =========================
  function buildCards(matrix, headerRow1, metricCol1) {
    const h = headerRow1 - 1;
    const mCol = metricCol1 - 1;

    const header = matrix[h];
    if (!header) throw new Error("Fila de fechas inválida (headerRow).");

    const cols = [];
    for (let c = 0; c < header.length; c++) {
      if (c === mCol) continue;
      const d = parseDate(header[c]);
      if (d) cols.push({ c, d });
    }
    if (!cols.length) throw new Error("No se detectaron columnas con fechas en el header.");

    const rows = matrix.slice(h + 1);

    return cols.map(({ c, d }) => {
      let lineaTM = "—";
      let lineaTT = "—";
      let legajos = [];
      const table = [];

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i] || [];
        const name = String(row[mCol] ?? "").trim();
        const val = String(row[c] ?? "").trim();

        // si no hay métrica, no es fila principal
        if (!name) continue;

        const n = normalize(name);

        // KPIs TM / TT
        if (n === "linea tm" || n.includes("linea tm")) lineaTM = val || "—";
        if (n === "linea tt" || n.includes("linea tt")) lineaTT = val || "—";

        // fila explícita de legajos (si existiera)
        if (n.includes("legajo") && n.includes("inasist")) {
          legajos = extractLegajos(val);
        }

        // ✅ FIX ROBUSTO: "Inasistencias TM" y legajos debajo (aunque estén más abajo)
        if (n === "inasistencias tm") {
          const expected = parseNumber(val); // cantidad, puede ayudar pero no limita

          const collected = [];
          let j = i + 1;

          // Recorre TODO lo que venga debajo hasta la próxima métrica "real"
          while (j < rows.length) {
            const nextName = String(rows[j]?.[mCol] ?? "").trim();
            const nextVal = String(rows[j]?.[c] ?? "").trim();

            // Si aparece otra métrica real (texto), corta.
            // OJO: hay sheets que ponen números en la columna de métrica por error,
            // por eso evitamos cortar si nextName es solo dígitos.
            if (nextName && !/^\d+$/.test(nextName)) break;

            // Acumula cualquier legajo numérico encontrado en el valor del día
            if (nextVal) collected.push(...extractLegajos(nextVal));

            j++;
          }

          // Si el sheet dice que hay inasistencias, pero no encontramos nada, queda vacío.
          // (si querés mostrar "—" explícito, se hace en render)
          if (expected > 0 && collected.length) {
            legajos = collected;
          } else if (collected.length) {
            // aunque expected sea 0, si hay legajos, los tomamos
            legajos = collected;
          }

          // saltar las filas consumidas
          i = j - 1;
        }

        // Tabla: todo excepto lo que pediste ocultar
        if (!shouldHideInTable(name)) {
          table.push({ name, val });
        }
      }

      legajos = uniqueSortLegajos(legajos);

      return { date: d, lineaTM, lineaTT, legajos, table };
    });
  }

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
  // RENDER
  // =========================
  function render() {
    if (!cardsGrid) return;
    cardsGrid.innerHTML = "";

    const today = startOfDay(new Date());

    windowCards(dayCards).forEach(card => {
      const isToday = card.date.getTime() === today.getTime();
      const hasInasist = card.legajos.length > 0;

      const el = document.createElement("article");
      el.className = `card ${hasInasist ? "card--alert" : "card--ok"} ${
        isToday ? "card--today" : ""
      }`;

      el.innerHTML = `
        <div class="card-body">
          <div class="card-date">
            <div class="date-pill">
              <span class="dot ${hasInasist ? "dot--red" : "dot--green"}"></span>
              ${escapeHtml(formatDate(card.date))}
            </div>
          </div>

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
              <div class="k">Legajo inasistencia</div>
              <div class="v">${escapeHtml(card.legajos.length ? card.legajos.join(", ") : "—")}</div>
            </div>
          </div>

          <div class="table table-scroll">
            ${card.table
              .map(
                r => `
                <div class="row">
                  <div class="key">${escapeHtml(r.name)}</div>
                  <div class="val">${escapeHtml(r.val || "—")}</div>
                </div>`
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
      const id = extractSpreadsheetId(SHEET_URL);
      if (!id) throw new Error("No pude extraer el ID del Sheet desde la URL.");

      const csv = await fetchCsv(id, DEFAULT_GID);
      const matrix = parseCsvToMatrix(csv);

      dayCards = buildCards(matrix, DEFAULT_HEADER_ROW, DEFAULT_METRIC_COL);
      render();
    } catch (err) {
      console.error(err);
      if (cardsGrid) cardsGrid.innerHTML = "";
    } finally {
      isLoading = false;
    }
  }

  function startAutoRefresh() {
    if (refreshTimer) clearInterval(refreshTimer);
    refreshTimer = setInterval(() => {
      if (document.visibilityState === "visible") load();
    }, REFRESH_MS);
  }

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") load();
  });

  load();
  startAutoRefresh();
})();
