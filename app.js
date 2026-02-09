(() => {
  // =========================
  // CONFIG – TU GOOGLE SHEET
  // =========================
  const SHEET_URL =
    "https://docs.google.com/spreadsheets/d/120WSaF1Zu6h4-Edid-Yc7GIWKwtQB61GQ3rNexH-MXc/edit?gid=0#gid=0";

  const DEFAULT_GID = "0";
  const DEFAULT_HEADER_ROW = 1; // fila donde están las fechas (1-based)
  const DEFAULT_METRIC_COL = 1; // columna métricas (A=1)

  // ✅ 7 tarjetas: (último día con datos <= hoy) + 6 días atrás
  const DAYS_BACK = 6;

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
  // [{ dayLabel, dayDate, metrics:[{name,value}], inasistCount, legajosInasist[], searchBlob }]
  let dayCards = [];
  let filteredCards = [];

  // =========================
  // INIT UI
  // =========================
  sheetUrlInput.value = SHEET_URL;
  gidInput.value = DEFAULT_GID;
  headerRowInput.value = String(DEFAULT_HEADER_ROW);
  metricColInput.value = String(DEFAULT_METRIC_COL);

  openSheetBtn?.addEventListener("click", () => {
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

  function startOfDay(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function addDays(d, days) {
    const x = new Date(d);
    x.setDate(x.getDate() + days);
    return x;
  }

  function iso(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${dd}`;
  }

  // =========================
  // FETCH CSV (blindado)
  // =========================
  async function fetchCsv(spreadsheetId, gid) {
    const csvUrl =
      `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}` +
      `/export?format=csv&gid=${encodeURIComponent(gid)}`;

    const res = await fetch(csvUrl, { cache: "no-store" });
    const contentType = res.headers.get("content-type") || "";
    const text = await res.text();

    console.log("CSV URL:", csvUrl);
    console.log("HTTP:", res.status, res.statusText);
    console.log("Content-Type:", contentType);
    console.log("Preview:", text.slice(0, 250));

    if (!res.ok) throw new Error(`HTTP ${res.status}. El Sheet no es accesible públicamente.`);

    if (
      contentType.includes("text/html") ||
      text.trim().startsWith("<") ||
      text.toLowerCase().includes("<html")
    ) {
      throw new Error(
        "Google devolvió HTML (login/permisos). Publicá el Sheet: Archivo → Publicar en la web."
      );
    }

    return text;
  }

  // =========================
  // CSV -> Matrix
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
  // Parse fecha desde headers
  // =========================
  function parseDayLabelToDate(label) {
    const raw = String(label || "").trim();
    if (!raw) return null;

    // ISO escondido
    let m = raw.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (m) {
      const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
      return isNaN(dt.getTime()) ? null : startOfDay(dt);
    }

    // dd/mm/yyyy
    m = raw.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) {
      const dd = Number(m[1]);
      const mm = Number(m[2]);
      const yy = Number(m[3]);
      const yyyy = yy < 100 ? 2000 + yy : yy;
      const dt = new Date(yyyy, mm - 1, dd);
      return isNaN(dt.getTime()) ? null : startOfDay(dt);
    }

    // dd/mm (sin año)
    m = raw.match(/(\d{1,2})\/(\d{1,2})/);
    if (m) {
      const dd = Number(m[1]);
      const mm = Number(m[2]);
      const yyyy = new Date().getFullYear();
      const dt = new Date(yyyy, mm - 1, dd);
      return isNaN(dt.getTime()) ? null : startOfDay(dt);
    }

    return null;
  }

  // =========================
  // Inasistencias + legajos
  // =========================
  function normalizeStr(s) {
    return String(s ?? "").trim().toLowerCase();
  }

  function isInasistMetricName(name) {
    const n = normalizeStr(name);
    // incluye "inasist..." pero excluye cosas de hora tipo "salida último servicio"
    return n.includes("inasist");
  }

  function parseNumericSafe(value) {
    const s = String(value ?? "").trim();
    if (!s) return 0;
    // si es hora (6:26) no lo tomamos como número
    if (/^\d{1,2}:\d{2}$/.test(s)) return 0;
    // tomar primer número (permite "2", "2.0", "2,0")
    const m = s.match(/-?\d+(?:[.,]\d+)?/);
    if (!m) return 0;
    return Number(m[0].replace(",", ".")) || 0;
  }

  function extractLegajosFromText(value) {
    // tolerante: separa por coma/espacio/punto y coma / salto
    // y se queda con "tokens" que tengan al menos 3 dígitos (ajustable)
    const s = String(value ?? "").trim();
    if (!s) return [];

    return s
      .split(/[\s,;|\n]+/g)
      .map((t) => t.trim())
      .filter(Boolean)
      .filter((t) => /\d{3,}/.test(t)); // "legajo" suele ser numérico
  }

  function computeInasistAndLegajos(metrics) {
    let inasistCount = 0;

    // Si el sheet trae una fila específica con legajos, la usamos
    let legajosInasist = [];

    for (const m of metrics) {
      const name = normalizeStr(m.name);
      const val = m.value;

      // suma de inasistencias (si el nombre contiene inasist...)
      if (isInasistMetricName(name)) {
        inasistCount += parseNumericSafe(val);
      }

      // detectar fila de legajos (varias posibilidades)
      const looksLikeLegajosRow =
        (name.includes("legajo") || name.includes("legajos")) &&
        name.includes("inasist");

      if (looksLikeLegajosRow) {
        legajosInasist = extractLegajosFromText(val);
      }
    }

    return { inasistCount, legajosInasist };
  }

  // =========================
  // Build cards (matriz -> columnas por día)
  // =========================
  function buildCardsFromMatrix(matrix, headerRow1, metricCol1) {
    const h = Math.max(1, Number(headerRow1 || 1)) - 1;
    const m = Math.max(1, Number(metricCol1 || 1)) - 1;

    const header = matrix[h];
    if (!header) throw new Error("Fila de fechas inexistente (revisá 'Fila de fechas').");

    const dayCols = [];
    for (let c = 0; c < header.length; c++) {
      if (c === m) continue;
      const dayDate = parseDayLabelToDate(header[c]);
      if (dayDate) dayCols.push({ c, label: header[c], dayDate });
    }

    if (!dayCols.length) throw new Error("No se detectaron columnas de fecha/día en esa fila.");

    const rows = matrix.slice(h + 1);

    return dayCols.map(({ c, label, dayDate }) => {
      const metrics = [];
      for (const r of rows) {
        const name = (r[m] ?? "").trim();
        if (!name) continue;
        metrics.push({ name, value: (r[c] ?? "").trim() });
      }

      const { inasistCount, legajosInasist } = computeInasistAndLegajos(metrics);

      return {
        dayLabel: label,
        dayDate,
        metrics,
        inasistCount,
        legajosInasist,
        searchBlob: (
          String(label) +
          " " +
          metrics.map((x) => `${x.name} ${x.value}`).join(" ")
        ).toLowerCase(),
      };
    });
  }

  // =========================
  // Ventana: ancla = max fecha <= hoy
  // =========================
  function windowAndSort(cards) {
    const today = startOfDay(new Date());

    const nonFuture = cards.filter((c) => c.dayDate && c.dayDate.getTime() <= today.getTime());
    if (!nonFuture.length) return { windowed: [], anchor: null, min: null, max: null };

    let anchor = nonFuture[0].dayDate;
    for (const c of nonFuture) if (c.dayDate.getTime() > anchor.getTime()) anchor = c.dayDate;

    const min = addDays(anchor, -DAYS_BACK);
    const max = anchor;

    const windowed = nonFuture
      .filter((c) => c.dayDate.getTime() >= min.getTime() && c.dayDate.getTime() <= max.getTime())
      .sort((a, b) => b.dayDate.getTime() - a.dayDate.getTime());

    return { windowed, anchor, min, max };
  }

  // =========================
  // Semáforo
  // =========================
  function trafficStatus(card) {
    // rojo si hay inasistencias > 0
    if (typeof card.inasistCount === "number") {
      if (card.inasistCount > 0) return "red";
      return "green";
    }
    return "yellow";
  }

  function dotHtml(status) {
    const color =
      status === "red" ? "#ef4444" :
      status === "green" ? "#22c55e" :
      "#f59e0b";

    return `<span class="dot" style="background:${color}; box-shadow:0 0 0 4px color-mix(in srgb, ${color} 25%, transparent);"></span>`;
  }

  // =========================
  // RENDER
  // =========================
  function buildCard(card) {
    const status = trafficStatus(card);

    // KPIs simples
    const nonEmpty = card.metrics.filter((m) => String(m.value ?? "").trim() !== "").length;

    // ✅ Mostrar TODAS las métricas en la tarjeta
    const rowsHtml = card.metrics
      .map(
        (m) => `
        <div class="row">
          <div class="key">${escapeHtml(m.name)}</div>
          <div class="val">${escapeHtml(m.value || "—")}</div>
        </div>`
      )
      .join("");

    const el = document.createElement("article");
    el.className = "card";

    el.innerHTML = `
      <div class="card-header">
        <div class="badge">
          ${dotHtml(status)}
          ${escapeHtml(card.dayLabel)} · ${card.metrics.length} métricas
          ${status === "red" ? `· ⚠️ Inasistencias: ${escapeHtml(String(card.inasistCount))}` : ""}
        </div>
        <div class="card-actions">
          <button class="icon-btn" title="Detalles (legajos con inasistencia)">🔎</button>
        </div>
      </div>

      <div class="card-body">
        <div class="kpi-row">
          <div class="kpi"><div class="k">Con valor</div><div class="v">${escapeHtml(String(nonEmpty))}</div></div>
          <div class="kpi"><div class="k">Total métricas</div><div class="v">${escapeHtml(String(card.metrics.length))}</div></div>
          <div class="kpi"><div class="k">Inasistencias</div><div class="v">${escapeHtml(String(card.inasistCount ?? 0))}</div></div>
        </div>

        <div class="table table-scroll">
          ${rowsHtml}
        </div>
      </div>
    `;

    el.querySelector(".icon-btn").onclick = () => openDetail(card);
    return el;
  }

  function renderCards() {
    cardsGrid.innerHTML = "";

    const { windowed, anchor, min, max } = windowAndSort(dayCards);

    const q = searchInput.value.trim().toLowerCase();
    filteredCards = !q ? windowed : windowed.filter((c) => c.searchBlob.includes(q));

    filteredCards.forEach((c) => cardsGrid.appendChild(buildCard(c)));

    if (!anchor) {
      showHint("No encontré fechas válidas (o todas son futuras). Revisá la fila de fechas.", "warn");
      return;
    }

    if (!filteredCards.length) {
      showHint(`No hay tarjetas en el rango ${iso(min)} → ${iso(max)} o no coincide la búsqueda.`, "warn");
      return;
    }

    showHint(`✅ Ordenado DESC. Mostrando ${filteredCards.length} día(s): ${iso(max)} → ${iso(min)}.`, "ok");
  }

  // ✅ Modal: SOLO legajos con inasistencia
  function openDetail(card) {
    modalTitle.textContent = `Detalles ${card.dayLabel}`;
    modalSubtitle.textContent =
      `Legajos con inasistencia · Inasistencias: ${String(card.inasistCount ?? 0)} · Fecha: ${iso(card.dayDate)}`;

    const legajos = Array.isArray(card.legajosInasist) ? card.legajosInasist : [];

    if (!legajos.length) {
      modalBody.innerHTML = `
        <div style="padding:12px; border:1px dashed rgba(255,255,255,0.14); border-radius:14px; color:rgba(255,255,255,0.72);">
          No se encontraron <b>legajos</b> en el Sheet para este día.
          <br/><br/>
          Para que aparezcan, el Google Sheet debe tener alguna fila tipo:
          <div style="margin-top:8px; font-family:var(--mono); font-size:12px; opacity:.9;">
            "Legajos con inasistencia" → "1234, 5678, 9012"
          </div>
        </div>
      `;
      detailModal.showModal();
      return;
    }

    modalBody.innerHTML = `
      <div style="display:flex; flex-wrap:wrap; gap:10px;">
        ${legajos
          .map(
            (l) => `
          <div style="
            padding:10px 12px;
            border-radius:999px;
            border:1px solid rgba(255,255,255,0.14);
            background:rgba(255,255,255,0.06);
            font-family:var(--mono);
            font-weight:800;
          ">${escapeHtml(l)}</div>`
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
      const gid = String(gidInput.value || DEFAULT_GID).trim();

      const csv = await fetchCsv(id, gid);
      const matrix = parseCsvToMatrix(csv);

      dayCards = buildCardsFromMatrix(
        matrix,
        Number(headerRowInput.value || DEFAULT_HEADER_ROW),
        Number(metricColInput.value || DEFAULT_METRIC_COL)
      );

      console.log(
        "Parsed cards:",
        dayCards.map((c) => ({
          label: c.dayLabel,
          date: c.dayDate ? iso(c.dayDate) : null,
          inasist: c.inasistCount,
          legajos: c.legajosInasist?.length || 0,
        }))
      );

      renderCards();
    } catch (err) {
      console.error(err);
      showHint(`Error: ${err.message}`, "danger");
    }
  }

  reloadBtn.onclick = load;
  applyBtn.onclick = load;
  searchInput.oninput = renderCards;

  // Autocarga
  load();
})();
