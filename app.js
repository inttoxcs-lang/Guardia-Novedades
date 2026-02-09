(() => {
  // =========================
  // CONFIG – TU GOOGLE SHEET
  // =========================
  const SHEET_URL =
    "https://docs.google.com/spreadsheets/d/120WSaF1Zu6h4-Edid-Yc7GIWKwtQB61GQ3rNexH-MXc/edit?gid=0#gid=0";

  const DEFAULT_GID = "0";
  const DEFAULT_HEADER_ROW = 1; // fila donde están las fechas (1-based)
  const DEFAULT_METRIC_COL = 1; // columna métricas (A=1)

  // ✅ Ventana: hoy y 7 días hacia atrás
  const DAYS_BACK = 7;

  // =========================
  // DOM
  // =========================
  const sheetUrlInput = document.getElementById("sheetUrl");
  const gidInput = document.getElementById("gidInput");
  const headerRowInput = document.getElementById("headerRowInput");
  const metricColInput = document.getElementById("metricColInput");
  const reloadBtn = document.getElementById("reloadBtn");
  const applyBtn = document.getElementById("applyBtn");
  const openSheetBtn = document.getElementById("openSheetBtn");
  const searchInput = document.getElementById("searchInput");
  const cardsGrid = document.getElementById("cardsGrid");
  const hint = document.getElementById("hint");

  const detailModal = document.getElementById("detailModal");
  const closeModalBtn = document.getElementById("closeModalBtn");
  const modalTitle = document.getElementById("modalTitle");
  const modalSubtitle = document.getElementById("modalSubtitle");
  const modalBody = document.getElementById("modalBody");

  // =========================
  // STATE
  // =========================
  let dayCards = []; // [{ dayLabel, dayDate, metrics, searchBlob }]
  let filteredCards = [];

  // =========================
  // INIT UI
  // =========================
  sheetUrlInput.value = SHEET_URL;
  gidInput.value = DEFAULT_GID;
  headerRowInput.value = DEFAULT_HEADER_ROW;
  metricColInput.value = DEFAULT_METRIC_COL;

  openSheetBtn.addEventListener("click", () => {
    window.open(SHEET_URL, "_blank", "noopener,noreferrer");
  });

  // =========================
  // HELPERS
  // =========================
  function showHint(text, tone = "info") {
    const colors = {
      info: "rgba(255,255,255,0.68)",
      warn: "rgba(245,158,11,0.95)",
      danger: "rgba(239,68,68,0.95)",
      ok: "rgba(34,197,94,0.95)",
    };
    hint.textContent = text;
    hint.style.color = colors[tone] || colors.info;
  }

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

  // =========================
  // FETCH CSV (blindado)
  // =========================
  async function fetchCsv(spreadsheetId, gid) {
    const csvUrl =
      `https://docs.google.com/spreadsheets/d/${encodeURIComponent(
        spreadsheetId
      )}/export?format=csv&gid=${encodeURIComponent(gid)}`;

    const res = await fetch(csvUrl, { cache: "no-store" });
    const contentType = res.headers.get("content-type") || "";
    const text = await res.text();

    console.log("CSV URL:", csvUrl);
    console.log("HTTP:", res.status, res.statusText);
    console.log("Content-Type:", contentType);
    console.log("Preview:", text.slice(0, 300));

    if (!res.ok) {
      throw new Error(`HTTP ${res.status}. El Sheet no es accesible públicamente.`);
    }

    if (
      contentType.includes("text/html") ||
      text.trim().startsWith("<") ||
      text.includes("<html")
    ) {
      throw new Error(
        "Google devolvió HTML (login/permisos). Publicá el Sheet: Archivo → Publicar en la web."
      );
    }

    return text;
  }

  // =========================
  // CSV → MATRIZ
  // =========================
  function parseCsvToMatrix(text) {
    const rows = [];
    let cur = [];
    let field = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      const next = text[i + 1];

      if (ch === '"' && inQuotes && next === '"') {
        field += '"';
        i++;
        continue;
      }
      if (ch === '"') {
        inQuotes = !inQuotes;
        continue;
      }

      if (!inQuotes && ch === ",") {
        cur.push(field);
        field = "";
        continue;
      }
      if (!inQuotes && ch === "\n") {
        cur.push(field);
        rows.push(cur);
        cur = [];
        field = "";
        continue;
      }
      if (ch !== "\r") field += ch;
    }

    cur.push(field);
    rows.push(cur);

    return rows.map((r) => r.map((c) => String(c ?? "").trim()));
  }

  // =========================
  // Fecha: parse de headers tipo 01/02, 01/02/2026, 2026-02-01
  // =========================
  function parseDayLabelToDate(label) {
    const s = String(label || "").trim();
    if (!s) return null;

    const today = new Date();
    const y = today.getFullYear();

    // yyyy-mm-dd
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) {
      const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
      return isNaN(dt.getTime()) ? null : dt;
    }

    // dd/mm/yyyy (o dd/mm/yy)
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) {
      const dd = Number(m[1]);
      const mm = Number(m[2]);
      let yy = Number(m[3]);
      const yyyy = yy < 100 ? 2000 + yy : yy;
      const dt = new Date(yyyy, mm - 1, dd);
      return isNaN(dt.getTime()) ? null : dt;
    }

    // dd/mm (sin año) -> asumimos año actual
    m = s.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (m) {
      const dd = Number(m[1]);
      const mm = Number(m[2]);
      const dt = new Date(y, mm - 1, dd);
      return isNaN(dt.getTime()) ? null : dt;
    }

    // dd-mm (sin año)
    m = s.match(/^(\d{1,2})-(\d{1,2})$/);
    if (m) {
      const dd = Number(m[1]);
      const mm = Number(m[2]);
      const dt = new Date(y, mm - 1, dd);
      return isNaN(dt.getTime()) ? null : dt;
    }

    return null;
  }

  function startOfDay(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function addDays(d, days) {
    const x = new Date(d);
    x.setDate(x.getDate() + days);
    return x;
  }

  function withinLastNDaysInclusive(dayDate, nDaysBack) {
    const today = startOfDay(new Date());
    const min = addDays(today, -nDaysBack); // hoy - N
    const max = today; // hoy
    const d = startOfDay(dayDate);

    // Solo hasta hoy, sin futuros
    return d >= min && d <= max;
  }

  // =========================
  // Construcción de cards
  // =========================
  function buildCardsFromMatrix(matrix, headerRow1, metricCol1) {
    const h = headerRow1 - 1;
    const m = metricCol1 - 1;

    const header = matrix[h];
    if (!header) throw new Error("Fila de fechas inexistente.");

    // columnas con fecha válida
    const dayCols = [];
    header.forEach((label, c) => {
      if (c === m) return;
      const dayDate = parseDayLabelToDate(label);
      if (dayDate) dayCols.push({ c, label, dayDate });
    });

    if (!dayCols.length) throw new Error("No se detectaron columnas de fecha/día.");

    const rows = matrix.slice(h + 1);

    const cards = dayCols.map(({ c, label, dayDate }) => {
      const metrics = [];
      rows.forEach((r) => {
        const name = (r[m] ?? "").trim();
        if (!name) return;
        metrics.push({ name, value: (r[c] ?? "").trim() });
      });

      return {
        dayLabel: label,
        dayDate,
        metrics,
        searchBlob: (
          label +
          " " +
          metrics.map((x) => `${x.name} ${x.value}`).join(" ")
        ).toLowerCase(),
      };
    });

    return cards;
  }

  // =========================
  // RENDER
  // =========================
  function buildCard(card) {
    const el = document.createElement("article");
    el.className = "card";

    el.innerHTML = `
      <div class="card-header">
        <div class="badge">
          <span class="dot"></span>${escapeHtml(card.dayLabel)} · ${card.metrics.length} métricas
        </div>
        <div class="card-actions">
          <button class="icon-btn" title="Ver detalle">🔎</button>
        </div>
      </div>
      <div class="card-body">
        <div class="table">
          ${card.metrics
            .slice(0, 10)
            .map(
              (m) => `
              <div class="row">
                <div class="key">${escapeHtml(m.name)}</div>
                <div class="val">${escapeHtml(m.value || "—")}</div>
              </div>`
            )
            .join("")}
        </div>
      </div>
    `;

    el.querySelector(".icon-btn").onclick = () => openDetail(card);
    return el;
  }

  function renderCards() {
    cardsGrid.innerHTML = "";

    // ✅ filtro por fecha: hoy → 7 días atrás (y nada futuro)
    const windowed = dayCards
      .filter((c) => c.dayDate && withinLastNDaysInclusive(c.dayDate, DAYS_BACK))
      // ✅ orden: hoy primero (desc)
      .sort((a, b) => startOfDay(b.dayDate).getTime() - startOfDay(a.dayDate).getTime());

    const q = searchInput.value.trim().toLowerCase();
    filteredCards = !q ? windowed : windowed.filter((c) => c.searchBlob.includes(q));

    filteredCards.forEach((c) => cardsGrid.appendChild(buildCard(c)));

    if (!filteredCards.length) {
      showHint(`No hay datos en el rango (hoy y ${DAYS_BACK} días atrás) o no coincide la búsqueda.`, "warn");
    } else {
      showHint(`Mostrando ${filteredCards.length} día(s): hoy → ${DAYS_BACK} días atrás.`, "ok");
    }
  }

  function openDetail(card) {
    modalTitle.textContent = `Detalle ${card.dayLabel}`;
    modalSubtitle.textContent = `${card.metrics.length} métricas (columna completa)`;
    modalBody.innerHTML = `
      <div class="table">
        ${card.metrics
          .map(
            (m) => `
          <div class="row">
            <div class="key">${escapeHtml(m.name)}</div>
            <div class="val">${escapeHtml(m.value || "—")}</div>
          </div>`
          )
          .join("")}
      </div>
    `;
    detailModal.showModal();
  }

  closeModalBtn.onclick = () => detailModal.close();

  // =========================
  // LOAD
  // =========================
  async function load() {
    try {
      showHint("Cargando datos del Google Sheet…", "info");
      cardsGrid.innerHTML = "";

      const id = extractSpreadsheetId(SHEET_URL);
      const csv = await fetchCsv(id, gidInput.value);
      const matrix = parseCsvToMatrix(csv);

      dayCards = buildCardsFromMatrix(
        matrix,
        Number(headerRowInput.value),
        Number(metricColInput.value)
      );

      renderCards();
    } catch (err) {
      console.error(err);
      showHint(err.message, "danger");
    }
  }

  reloadBtn.onclick = load;
  applyBtn.onclick = load;
  searchInput.oninput = renderCards;

  // AUTOCARGA
  load();
})();
