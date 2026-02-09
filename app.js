(() => {
  // =========================
  // CONFIG
  // =========================
  const SHEET_URL =
    "https://docs.google.com/spreadsheets/d/120WSaF1Zu6h4-Edid-Yc7GIWKwtQB61GQ3rNexH-MXc/edit?gid=0#gid=0";

  const DEFAULT_GID = "0";
  const DEFAULT_HEADER_ROW = 1;
  const DEFAULT_METRIC_COL = 1;

  // 7 tarjetas: ancla (último día con datos <= hoy) + 6 días atrás
  const DAYS_BACK = 6;

  // =========================
  // DOM
  // =========================
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

  function extractLegajos(value) {
    // captura números (>=3 dígitos)
    return String(value ?? "")
      .split(/[\s,;|]+/)
      .map(v => v.trim())
      .filter(v => /^\d{3,}$/.test(v));
  }

  function uniqueSortLegajos(legajos) {
    // 1) normaliza a string, 2) quita duplicados, 3) orden numérico
    const uniq = Array.from(new Set((legajos || []).map(String)));

    // sort numérico: si alguno no es numérico, lo manda al final por string
    return uniq.sort((a, b) => {
      const na = Number(a);
      const nb = Number(b);
      const aIsNum = Number.isFinite(na);
      const bIsNum = Number.isFinite(nb);

      if (aIsNum && bIsNum) return na - nb;
      if (aIsNum && !bIsNum) return -1;
      if (!aIsNum && bIsNum) return 1;
      return a.localeCompare(b, "es");
    });
  }

  function formatDateDDMMYYYY(date) {
    const d = String(date.getDate()).padStart(2, "0");
    const m = String(date.getMonth() + 1).padStart(2, "0");
    return `${d}/${m}/${date.getFullYear()}`;
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
  // CSV -> MATRIX
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
  // FILTRO DE FILAS EN TABLA
  // =========================
  function shouldHideInTable(metricName) {
    const n = normalize(metricName);
    if (n === "linea tm") return true;
    if (n === "linea tt") return true;
    if (n.includes("legajo") && n.includes("inasist")) return true;

    // si también querés ocultar "Inasistencias TM" de la tabla:
    if (n === "inasistencias tm") return true;

    return false;
  }

  // =========================
  // BUILD CARDS ✅ (incluye legajos bajo "Inasistencias TM")
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
      let lineaTM = "—";
      let lineaTT = "—";
      let legajosInasist = [];

      const metricsTable = [];

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const name = (r[mCol] ?? "").trim();
        const value = (r[c] ?? "").trim();

        // filas sin nombre: por defecto no van a tabla
        if (!name) continue;

        const n = normalize(name);

        // KPI TM/TT
        if (n.includes("linea tm")) lineaTM = value || "—";
        if (n.includes("linea tt")) lineaTT = value || "—";

        // KPI legajos si existe fila explícita "legajos ... inasist"
        if (n.includes("legajo") && n.includes("inasist")) {
          legajosInasist = extractLegajos(value);
        }

        // ✅ Si es "Inasistencias TM", capturar legajos en filas siguientes sin nombre
        if (n === "inasistencias tm") {
          const expected = parseNumber(value); // ej: 3
          const collected = [];

          let j = i + 1;
          while (j < rows.length) {
            const nextName = (rows[j][mCol] ?? "").trim();
            const nextVal = (rows[j][c] ?? "").trim();

            if (nextName) break; // llegó otra métrica, cortar
            const legs = extractLegajos(nextVal);
            if (legs.length) collected.push(...legs);
            j++;
          }

          if (collected.length) {
            legajosInasist = collected;
          } else if (expected > 0 && !legajosInasist.length) {
            // si hay inasist pero no legajos listados, dejamos vacío "—"
            legajosInasist = [];
          }

          i = j - 1; // saltar las filas vacías consumidas
        }

        // Tabla: incluir todo excepto lo oculto
        if (!shouldHideInTable(name)) {
          metricsTable.push({ name, value });
        }
      }

      // ✅ ordenar numéricamente y quitar duplicados
      legajosInasist = uniqueSortLegajos(legajosInasist);

      return {
        date: d,
        lineaTM,
        lineaTT,
        legajosInasist,
        metricsTable,
      };
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
  // RENDER
  // =========================
  function render() {
    cardsGrid.innerHTML = "";

    windowCards(dayCards).forEach(card => {
      const el = document.createElement("article");
      el.className = "card";

      const fecha = formatDateDDMMYYYY(card.date);
      const legajosTxt = card.legajosInasist.length ? card.legajosInasist.join(", ") : "—";

      el.innerHTML = `
        <div class="card-body">
          <div class="card-date">
            <div class="date-pill">
              <span class="dot"></span>
              ${escapeHtml(fecha)}
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
              <div class="v">${escapeHtml(legajosTxt)}</div>
            </div>
          </div>

          <div class="table table-scroll">
            ${card.metricsTable.map(m => `
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
    const csv = await fetchCsv(id, DEFAULT_GID);
    const matrix = parseCsv(csv);

    dayCards = buildCards(matrix, DEFAULT_HEADER_ROW, DEFAULT_METRIC_COL);
    render();
  }

  load().catch(err => console.error(err));
})();
