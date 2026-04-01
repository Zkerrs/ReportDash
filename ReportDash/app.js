(function () {
  "use strict";

  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const fileNameEl = document.getElementById("fileName");
  const importConfig = document.getElementById("importConfig");
  const importStats = document.getElementById("importStats");
  const importCols = document.getElementById("importCols");
  const importPrefilters = document.getElementById("importPrefilters");
  const importRowLimit = document.getElementById("importRowLimit");
  const btnColsAll = document.getElementById("btnColsAll");
  const btnColsNone = document.getElementById("btnColsNone");
  const btnAddPrefilter = document.getElementById("btnAddPrefilter");
  const btnReport = document.getElementById("btnReport");
  const reportBarBusy = document.getElementById("reportBarBusy");
  const reportFullscreen = document.getElementById("reportFullscreen");
  const reportScroll = document.getElementById("reportScroll");
  const reportTable = document.getElementById("reportTable");
  const reportLoading = document.getElementById("reportLoading");
  const reportMeta = document.getElementById("reportMeta");
  const btnCloseReport = document.getElementById("btnCloseReport");
  const exportWorkspace = document.getElementById("exportWorkspace");

  /** Dados brutos da planilha (após carregar o arquivo). */
  let parsedRows = [];
  let parsedHeaders = [];
  /** Subconjunto efetivo exibido no relatório. */
  let rows = [];
  let headers = [];
  /** @type {string} */
  let loadedFileName = "";
  let reportRenderToken = 0;
  /** @type {((this: Window, ev: UIEvent) => void) | null} */
  let reportResizeHandler = null;

  /** Só após «Abrir relatório»: spinner na barra + botão desabilitado até terminar. */
  function setReportBusy(busy) {
    if (reportBarBusy) reportBarBusy.hidden = !busy;
    if (btnReport) btnReport.disabled = !!busy;
  }

  function layoutReportScroll() {
    if (!reportFullscreen || !reportScroll) return;
    const bar = reportFullscreen.querySelector(".report-fullscreen__bar");
    let h = 52;
    if (bar) {
      const rect = bar.getBoundingClientRect();
      const measured = Math.ceil(rect.height);
      /* Altura 0 antes do paint deixava top:0 e a tabela ia para trás da barra (tela “em branco”) */
      if (measured >= 32) {
        h = measured;
      }
    }
    reportScroll.style.top = h + "px";
  }

  function parseWorkbook(arrayBuffer) {
    if (typeof XLSX === "undefined") {
      throw new Error("Biblioteca XLSX não carregou (sem internet ou CDN bloqueado).");
    }
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    if (!wb.SheetNames || !wb.SheetNames.length) {
      parsedRows = [];
      parsedHeaders = [];
      rows = [];
      headers = [];
      return;
    }
    const firstName = wb.SheetNames[0];
    const sheet = wb.Sheets[firstName];
    const json = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: false,
      cellDates: true,
    });
    if (!json.length) {
      parsedRows = [];
      parsedHeaders = [];
      rows = [];
      headers = [];
      return;
    }
    parsedHeaders = Object.keys(json[0]);
    parsedRows = json;
    rows = [];
    headers = [];
  }

  /** Colunas cujo nome sugere data (evita «Lançamento contábil» = ID numérico). */
  function columnLooksLikeDate(name) {
    const s = String(name || "");
    if (/\bdata\b/i.test(s) || /\bdate\b/i.test(s)) return true;
    if (/vencimento|venc\.|emiss[aã]o|compensa(ção|cao)?\b|período|periodo/i.test(s)) {
      return true;
    }
    return false;
  }

  function formatDateDMY(d) {
    const day = String(d.getDate()).padStart(2, "0");
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const year = d.getFullYear();
    return `${day}/${month}/${year}`;
  }

  /** Converte serial Excel (com ou sem hora fracionária) em Date local. */
  function excelSerialToDate(serial) {
    const n = Number(serial);
    if (!Number.isFinite(n)) return null;
    const days = Math.floor(n);
    if (days < 1 || days > 2958465) return null;
    const utcMs = (days - 25569) * 86400 * 1000;
    const d = new Date(utcMs);
    if (isNaN(d.getTime())) return null;
    const frac = n - days;
    if (frac > 1e-12) {
      let secs = Math.round(frac * 86400);
      if (secs >= 86400) secs = 86399;
      const hh = Math.floor(secs / 3600);
      const mm = Math.floor((secs % 3600) / 60);
      const ss = secs % 60;
      d.setHours(hh, mm, ss, 0);
    }
    return d;
  }

  /**
   * D/M/Y com separador / - ou . e ano de 2 ou 4 dígitos (ex.: 1/1/25 → 01/01/25).
   * Ordem: dia / mês / ano (pt-BR). Opcional: hora depois do texto.
   */
  function padBrazilianDateString(s) {
    const t = String(s).trim();
    const sepClass = "[.\\/\\-]";
    const yearPart = "(\\d{2}|\\d{4})";
    const withTime = new RegExp(
      "^(\\d{1,2})(" + sepClass + ")(\\d{1,2})\\2" + yearPart + "(\\s+.+)$"
    );
    const m2 = t.match(withTime);
    if (m2) {
      const day = m2[1].padStart(2, "0");
      const sep = m2[2];
      const month = m2[3].padStart(2, "0");
      return `${day}${sep}${month}${sep}${m2[4]}${m2[5]}`;
    }
    const dateOnly = new RegExp(
      "^(\\d{1,2})(" + sepClass + ")(\\d{1,2})\\2" + yearPart + "$"
    );
    const m = t.match(dateOnly);
    if (m) {
      const day = m[1].padStart(2, "0");
      const sep = m[2];
      const month = m[3].padStart(2, "0");
      return `${day}${sep}${month}${sep}${m[4]}`;
    }
    return t;
  }

  function cellText(v, columnKey) {
    if (v == null) return "";
    if (v instanceof Date && !isNaN(v.getTime())) {
      return formatDateDMY(v);
    }
    if (typeof v === "number" && Number.isFinite(v)) {
      if (columnLooksLikeDate(columnKey)) {
        const d = excelSerialToDate(v);
        if (d) {
          const n = v;
          const frac = n - Math.floor(n);
          if (frac > 1e-12) {
            const hh = String(d.getHours()).padStart(2, "0");
            const mm = String(d.getMinutes()).padStart(2, "0");
            const ss = String(d.getSeconds()).padStart(2, "0");
            return `${formatDateDMY(d)} ${hh}:${mm}:${ss}`;
          }
          return formatDateDMY(d);
        }
      }
      return String(v);
    }
    if (typeof v === "object") {
      try {
        return JSON.stringify(v);
      } catch (e) {
        return "";
      }
    }
    return padBrazilianDateString(String(v));
  }

  const FILTER_MAX_LIST = 500;
  /** Máximo de valores distintos na lista/datalist dos prefiltros (evita DOM enorme). */
  const PREFILTER_UNIQUE_UI_CAP = 1200;
  /** Valor interno do select para células vazias (evita conflito com placeholder value=""). */
  const PREFILTER_EMPTY_TOKEN = "\uE000";
  let prefilterRowId = 0;
  /** @type {HTMLTableSectionElement|null} */
  let reportTbodyEl = null;
  let columnFilters = Object.create(null);
  /** @type {HTMLDivElement|null} */
  let columnFilterPanel = null;
  let columnFilterOutsideHandler = null;
  let columnFilterEscHandler = null;
  let columnFilterResizeHandler = null;
  /** @type {HTMLButtonElement|null} */
  let columnFilterAnchorBtn = null;
  let filterPanelUiInited = false;

  function getRowCellValue(row, colKey) {
    return row && typeof row === "object" ? cellText(row[colKey], colKey) : "";
  }

  function escapeHtml(s) {
    const div = document.createElement("div");
    div.textContent = s;
    return div.innerHTML;
  }

  function populateImportColumns() {
    if (!importCols) return;
    importCols.innerHTML = "";
    parsedHeaders.forEach(function (h) {
      const lab = document.createElement("label");
      lab.className = "import-config__col";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = true;
      cb.dataset.header = h;
      const span = document.createElement("span");
      span.textContent = String(h);
      lab.appendChild(cb);
      lab.appendChild(span);
      importCols.appendChild(lab);
    });
  }

  function resetImportPrefilters() {
    if (importPrefilters) importPrefilters.innerHTML = "";
  }

  function getUniqueValuesForParsedColumn(colKey) {
    if (!colKey) return [];
    const set = new Set();
    for (let i = 0; i < parsedRows.length; i++) {
      set.add(getRowCellValue(parsedRows[i], colKey));
    }
    return Array.from(set).sort(function (a, b) {
      return a.localeCompare(b, "pt-BR", { numeric: true, sensitivity: "base" });
    });
  }

  function refreshPrefilterValueOptions(rowEl) {
    const colSel = rowEl.querySelector(".import-prefilter-col");
    const dl = rowEl.querySelector(".import-prefilter-datalist");
    const quick = rowEl.querySelector(".import-prefilter-quick");
    if (!colSel || !dl || !quick) return;
    const col = colSel.value;
    dl.innerHTML = "";
    quick.innerHTML = "";
    if (!col) {
      const ph = document.createElement("option");
      ph.value = "";
      ph.textContent = "Valores da coluna…";
      quick.appendChild(ph);
      return;
    }
    const all = getUniqueValuesForParsedColumn(col);
    const list = all.slice(0, PREFILTER_UNIQUE_UI_CAP);
    const truncated = all.length > list.length;
    list.forEach(function (u) {
      const dOpt = document.createElement("option");
      dOpt.value = u;
      dl.appendChild(dOpt);
    });
    const ph2 = document.createElement("option");
    ph2.value = "";
    ph2.textContent = truncated
      ? "Lista (" + list.length.toLocaleString("pt-BR") + "+ valores distintos)…"
      : "Ver valores (" + list.length.toLocaleString("pt-BR") + ")…";
    quick.appendChild(ph2);
    list.forEach(function (u) {
      const o = document.createElement("option");
      o.value = u === "" ? PREFILTER_EMPTY_TOKEN : u;
      let t = String(u);
      if (t.length === 0) t = "(vazio)";
      else if (t.length > 96) t = t.slice(0, 93) + "…";
      o.textContent = t;
      quick.appendChild(o);
    });
  }

  function addPrefilterRow() {
    if (!importPrefilters) return;
    const row = document.createElement("div");
    row.className = "import-prefilter-row";
    const sel = document.createElement("select");
    sel.className = "import-prefilter-col import-config__select";
    const opt0 = document.createElement("option");
    opt0.value = "";
    opt0.textContent = "Coluna…";
    sel.appendChild(opt0);
    parsedHeaders.forEach(function (h) {
      const o = document.createElement("option");
      o.value = h;
      o.textContent = String(h);
      sel.appendChild(o);
    });
    const rowN = ++prefilterRowId;
    const datalist = document.createElement("datalist");
    datalist.className = "import-prefilter-datalist";
    datalist.id = "import-prefilter-dl-" + rowN;
    const inp = document.createElement("input");
    inp.type = "text";
    inp.className = "import-prefilter-val";
    inp.setAttribute("list", datalist.id);
    inp.placeholder = "Contém…";
    inp.setAttribute("aria-label", "Filtrar por texto ou valor da lista");
    inp.title =
      "Digite parte do texto ou use a lista «Ver valores» para escolher um valor existente na coluna.";
    const quick = document.createElement("select");
    quick.className = "import-prefilter-quick import-config__select";
    quick.setAttribute("aria-label", "Valores distintos da coluna");
    const rm = document.createElement("button");
    rm.type = "button";
    rm.className = "import-prefilter-remove import-config__mini";
    rm.textContent = "Remover";
    rm.addEventListener("click", function () {
      row.remove();
    });
    sel.addEventListener("change", function () {
      refreshPrefilterValueOptions(row);
    });
    quick.addEventListener("change", function () {
      const v = quick.value;
      if (v === "") return;
      inp.value = v === PREFILTER_EMPTY_TOKEN ? "" : v;
      quick.selectedIndex = 0;
    });
    const valWrap = document.createElement("div");
    valWrap.className = "import-prefilter-val-wrap";
    valWrap.appendChild(inp);
    valWrap.appendChild(datalist);
    row.appendChild(sel);
    row.appendChild(valWrap);
    row.appendChild(quick);
    row.appendChild(rm);
    importPrefilters.appendChild(row);
    refreshPrefilterValueOptions(row);
  }

  function showImportConfigAfterParse() {
    if (!importConfig || !importStats) return;
    populateImportColumns();
    resetImportPrefilters();
    importStats.innerHTML =
      '<strong class="import-config__fname">' +
      escapeHtml(loadedFileName) +
      '</strong><p class="import-config__stat-line">Planilha: <span class="import-config__stat-num">' +
      parsedRows.length.toLocaleString("pt-BR") +
      '</span> linhas · <span class="import-config__stat-num">' +
      parsedHeaders.length +
      "</span> colunas (primeira aba)</p>";
    importConfig.hidden = false;
  }

  function buildRowsFromImportSettings() {
    if (!importCols || !importRowLimit) return false;
    const colChecks = importCols.querySelectorAll('input[type="checkbox"][data-header]');
    const selected = [];
    colChecks.forEach(function (cb) {
      if (cb.checked) selected.push(cb.dataset.header);
    });
    if (selected.length === 0) {
      alert("Marque pelo menos uma coluna.");
      return false;
    }
    let r = parsedRows.slice();
    if (importPrefilters) {
      importPrefilters.querySelectorAll(".import-prefilter-row").forEach(function (rowEl) {
        const sel = rowEl.querySelector(".import-prefilter-col");
        const inp = rowEl.querySelector(".import-prefilter-val");
        const col = sel && sel.value;
        const q = inp && String(inp.value || "").trim();
        if (!col || !q) return;
        r = r.filter(function (row) {
          return getRowCellValue(row, col).toLowerCase().includes(q.toLowerCase());
        });
      });
    }
    const lim = importRowLimit.value;
    if (lim !== "all") {
      const n = parseInt(lim, 10);
      if (!Number.isFinite(n) || n < 1) {
        alert("Limite de linhas inválido.");
        return false;
      }
      r = r.slice(0, n);
    } else if (r.length > 15000) {
      if (
        !confirm(
          "Carregar " +
            r.length.toLocaleString("pt-BR") +
            " linhas pode deixar o navegador bem lento. Continuar?"
        )
      ) {
        return false;
      }
    }
    headers = selected;
    rows = r.map(function (row) {
      const o = {};
      selected.forEach(function (h) {
        o[h] = row && typeof row === "object" ? row[h] : "";
      });
      return o;
    });
    return true;
  }

  function getUniqueEntriesForColumn(colKey) {
    const map = new Map();
    for (let r = 0; r < rows.length; r++) {
      const v = getRowCellValue(rows[r], colKey);
      map.set(v, (map.get(v) || 0) + 1);
    }
    return Array.from(map.entries()).sort(function (a, b) {
      return a[0].localeCompare(b[0], "pt-BR", { numeric: true, sensitivity: "base" });
    });
  }

  function applyColumnFiltersToDom() {
    if (!reportTbodyEl) return;
    const trs = reportTbodyEl.rows;
    for (let i = 0; i < rows.length; i++) {
      const tr = trs[i];
      if (!tr) continue;
      const row = rows[i];
      let visible = true;
      for (let c = 0; c < headers.length; c++) {
        const col = headers[c];
        const f = columnFilters[col];
        if (f == null) continue;
        const val = getRowCellValue(row, col);
        if (f instanceof Set) {
          if (!f.has(val)) visible = false;
        } else if (f && f.type === "contains") {
          const q = String(f.q || "").toLowerCase();
          if (!String(val).toLowerCase().includes(q)) visible = false;
        }
      }
      tr.style.display = visible ? "" : "none";
    }
  }

  function updateThFilterActiveState() {
    if (!reportTable) return;
    reportTable.querySelectorAll("th[data-col-index]").forEach(function (th) {
      const idx = parseInt(th.getAttribute("data-col-index"), 10);
      const col = headers[idx];
      const btn = th.querySelector(".th-filter__btn");
      if (!btn || col == null) return;
      if (columnFilters[col] != null) btn.setAttribute("data-active", "1");
      else btn.removeAttribute("data-active");
    });
  }

  function closeColumnFilterPanel() {
    if (columnFilterPanel) columnFilterPanel.setAttribute("hidden", "");
    if (columnFilterAnchorBtn) {
      columnFilterAnchorBtn.setAttribute("aria-expanded", "false");
      columnFilterAnchorBtn = null;
    }
    if (columnFilterOutsideHandler) {
      document.removeEventListener("mousedown", columnFilterOutsideHandler, true);
      columnFilterOutsideHandler = null;
    }
    if (columnFilterEscHandler) {
      document.removeEventListener("keydown", columnFilterEscHandler, true);
      columnFilterEscHandler = null;
    }
    if (columnFilterResizeHandler) {
      window.removeEventListener("resize", columnFilterResizeHandler);
      columnFilterResizeHandler = null;
    }
  }

  function positionColumnFilterPanel(anchor) {
    if (!columnFilterPanel || !anchor) return;
    columnFilterPanel.style.position = "fixed";
    columnFilterPanel.style.zIndex = "400";
    const rect = anchor.getBoundingClientRect();
    const margin = 8;
    columnFilterPanel.style.width = "min(320px, calc(100vw - 16px))";
    void columnFilterPanel.offsetWidth;
    let left = rect.left;
    let top = rect.bottom + 6;
    const pw = columnFilterPanel.offsetWidth;
    const ph = columnFilterPanel.offsetHeight;
    if (left + pw > window.innerWidth - margin) {
      left = window.innerWidth - pw - margin;
    }
    if (left < margin) left = margin;
    if (top + ph > window.innerHeight - margin) {
      top = rect.top - ph - 6;
    }
    if (top < margin) top = margin;
    columnFilterPanel.style.left = left + "px";
    columnFilterPanel.style.top = top + "px";
  }

  function ensureColumnFilterPanel() {
    if (columnFilterPanel) return columnFilterPanel;
    const el = document.createElement("div");
    el.className = "col-filter-panel";
    el.setAttribute("hidden", "");
    el.innerHTML =
      '<div class="col-filter-panel__head">' +
      '<span class="col-filter-panel__title"></span>' +
      '<button type="button" class="col-filter-panel__x" aria-label="Fechar">×</button>' +
      "</div>" +
      '<p class="col-filter-panel__notice" hidden></p>' +
      '<div class="col-filter-panel__contains" hidden>' +
      '<label class="col-filter-panel__contains-label">Contém o texto</label>' +
      '<input type="text" class="col-filter-panel__contains-input" placeholder="Digite para filtrar…" />' +
      "</div>" +
      '<div class="col-filter-panel__search">' +
      '<input type="search" class="col-filter-panel__search-input" placeholder="Buscar na lista…" />' +
      "</div>" +
      '<div class="col-filter-panel__toggles">' +
      '<button type="button" class="col-filter-panel__link" data-all>Marcar todos</button>' +
      '<span class="col-filter-panel__dot">·</span>' +
      '<button type="button" class="col-filter-panel__link" data-none>Desmarcar todos</button>' +
      "</div>" +
      '<div class="col-filter-panel__list"></div>' +
      '<div class="col-filter-panel__foot">' +
      '<button type="button" class="col-filter-panel__btn col-filter-panel__btn--primary" data-apply>Aplicar</button>' +
      '<button type="button" class="col-filter-panel__btn" data-clear>Limpar filtro</button>' +
      "</div>";
    document.body.appendChild(el);
    columnFilterPanel = el;
    return el;
  }

  function initColumnFilterPanelUiOnce() {
    if (filterPanelUiInited) return;
    filterPanelUiInited = true;
    const panel = ensureColumnFilterPanel();
    panel.querySelector(".col-filter-panel__x").addEventListener("click", function (e) {
      e.stopPropagation();
      closeColumnFilterPanel();
    });
    panel.querySelector("[data-all]").addEventListener("click", function () {
      panel.querySelectorAll(".col-filter-panel__list input[type=checkbox]").forEach(function (cb) {
        cb.checked = true;
      });
    });
    panel.querySelector("[data-none]").addEventListener("click", function () {
      panel.querySelectorAll(".col-filter-panel__list input[type=checkbox]").forEach(function (cb) {
        cb.checked = false;
      });
    });
    panel.querySelector("[data-apply]").addEventListener("click", function () {
      const idx = parseInt(panel.getAttribute("data-open-col-index"), 10);
      if (Number.isNaN(idx) || !headers[idx]) return;
      const colKey = headers[idx];
      const containsWrap = panel.querySelector(".col-filter-panel__contains");
      if (!containsWrap.hidden) {
        const q = panel.querySelector(".col-filter-panel__contains-input").value.trim();
        if (q === "") delete columnFilters[colKey];
        else columnFilters[colKey] = { type: "contains", q: q };
      } else {
        const allCbs = panel.querySelectorAll(".col-filter-panel__list input[type=checkbox]");
        let checked = 0;
        const allowed = new Set();
        allCbs.forEach(function (cb) {
          if (cb.checked) {
            checked += 1;
            let v = cb.value;
            try {
              v = decodeURIComponent(v);
            } catch (e) {
              /* valor com % inválido */
            }
            allowed.add(v);
          }
        });
        if (checked === 0) {
          alert("Marque pelo menos um valor, ou use «Limpar filtro».");
          return;
        }
        if (checked === allCbs.length) delete columnFilters[colKey];
        else columnFilters[colKey] = allowed;
      }
      updateThFilterActiveState();
      applyColumnFiltersToDom();
      closeColumnFilterPanel();
    });
    panel.querySelector("[data-clear]").addEventListener("click", function () {
      const idx = parseInt(panel.getAttribute("data-open-col-index"), 10);
      if (!Number.isNaN(idx) && headers[idx]) delete columnFilters[headers[idx]];
      updateThFilterActiveState();
      applyColumnFiltersToDom();
      closeColumnFilterPanel();
    });
  }

  function openColumnFilterPanel(colIndex, btn) {
    const colKey = headers[colIndex];
    if (!colKey) return;
    initColumnFilterPanelUiOnce();
    closeColumnFilterPanel();
    columnFilterAnchorBtn = btn;
    btn.setAttribute("aria-expanded", "true");
    const panel = ensureColumnFilterPanel();
    panel.setAttribute("data-open-col-index", String(colIndex));
    const entries = getUniqueEntriesForColumn(colKey);
    const useList = entries.length <= FILTER_MAX_LIST;
    panel.querySelector(".col-filter-panel__title").textContent = colKey;
    const notice = panel.querySelector(".col-filter-panel__notice");
    const containsWrap = panel.querySelector(".col-filter-panel__contains");
    const searchWrap = panel.querySelector(".col-filter-panel__search");
    const toggles = panel.querySelector(".col-filter-panel__toggles");
    const listEl = panel.querySelector(".col-filter-panel__list");
    const searchInput = panel.querySelector(".col-filter-panel__search-input");
    listEl.innerHTML = "";
    searchInput.value = "";
    searchInput.oninput = null;

    const cur = columnFilters[colKey];
    let selectedSet = null;
    if (cur instanceof Set) selectedSet = new Set(cur);
    else if (cur == null) selectedSet = null;
    else selectedSet = new Set();

    const containsInput = panel.querySelector(".col-filter-panel__contains-input");
    const containsLabel = panel.querySelector(".col-filter-panel__contains-label");
    if (useList) {
      notice.hidden = true;
      containsWrap.hidden = true;
      containsInput.removeAttribute("id");
      containsLabel.removeAttribute("for");
      searchWrap.hidden = false;
      toggles.hidden = false;
      if (selectedSet == null || (cur && cur.type === "contains")) {
        selectedSet = new Set(entries.map(function (e) {
          return e[0];
        }));
      }
      const frag = document.createDocumentFragment();
      entries.forEach(function (pair) {
        const val = pair[0];
        const count = pair[1];
        const label = document.createElement("label");
        label.className = "col-filter-panel__row";
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = encodeURIComponent(val);
        cb.checked = selectedSet.has(val);
        const span = document.createElement("span");
        span.className = "col-filter-panel__row-text";
        span.textContent = val === "" ? "(vazio)" : val;
        const cnt = document.createElement("span");
        cnt.className = "col-filter-panel__row-count";
        cnt.textContent = count.toLocaleString("pt-BR");
        label.appendChild(cb);
        label.appendChild(span);
        label.appendChild(cnt);
        frag.appendChild(label);
      });
      listEl.appendChild(frag);
      searchInput.oninput = function () {
        const q = searchInput.value.toLowerCase();
        listEl.querySelectorAll(".col-filter-panel__row").forEach(function (lbl) {
          const t = lbl.querySelector(".col-filter-panel__row-text").textContent.toLowerCase();
          lbl.style.display = t.includes(q) ? "" : "none";
        });
      };
    } else {
      notice.hidden = false;
      notice.textContent =
        "Esta coluna tem " +
        entries.length.toLocaleString("pt-BR") +
        " valores distintos. Use o filtro por texto abaixo (estilo «Contém»).";
      containsWrap.hidden = false;
      searchWrap.hidden = true;
      toggles.hidden = true;
      listEl.innerHTML =
        '<p class="col-filter-panel__hint">Digite parte do texto e clique em Aplicar.</p>';
      const lid = "col-filter-contains-" + colIndex;
      containsInput.id = lid;
      containsLabel.setAttribute("for", lid);
      containsInput.value = cur && cur.type === "contains" ? cur.q : "";
    }

    panel.removeAttribute("hidden");
    positionColumnFilterPanel(btn);
    columnFilterOutsideHandler = function (ev) {
      if (
        columnFilterPanel.contains(ev.target) ||
        ev.target === btn ||
        (btn && btn.contains(ev.target))
      ) {
        return;
      }
      closeColumnFilterPanel();
    };
    document.addEventListener("mousedown", columnFilterOutsideHandler, true);
    columnFilterEscHandler = function (ev) {
      if (ev.key === "Escape") closeColumnFilterPanel();
    };
    document.addEventListener("keydown", columnFilterEscHandler, true);
    columnFilterResizeHandler = function () {
      if (columnFilterAnchorBtn) positionColumnFilterPanel(columnFilterAnchorBtn);
    };
    window.addEventListener("resize", columnFilterResizeHandler);
  }

  function bindReportTableFilters(thead, tbody) {
    reportTbodyEl = tbody;
    columnFilters = Object.create(null);
    initColumnFilterPanelUiOnce();
    if (!thead._reportFilterBound) {
      thead._reportFilterBound = true;
      thead.addEventListener("click", function (ev) {
        const btn = ev.target.closest(".th-filter__btn");
        if (!btn) return;
        ev.preventDefault();
        ev.stopPropagation();
        const th = btn.closest("th");
        const idx = parseInt(th.getAttribute("data-col-index"), 10);
        if (Number.isNaN(idx)) return;
        if (
          columnFilterPanel &&
          !columnFilterPanel.hasAttribute("hidden") &&
          columnFilterAnchorBtn === btn
        ) {
          closeColumnFilterPanel();
          return;
        }
        openColumnFilterPanel(idx, btn);
      });
    }
  }

  function renderReportTableBatched() {
    reportRenderToken += 1;
    const token = reportRenderToken;
    setReportBusy(true);

    reportScroll.querySelectorAll(".report-error").forEach(function (el) {
      el.remove();
    });
    reportLoading.hidden = false;
    reportTable.innerHTML = "";

    try {
      if (!headers.length) {
        reportLoading.hidden = true;
        setReportBusy(false);
        const cap = document.createElement("caption");
        cap.className = "report-empty";
        cap.textContent = "Nenhuma coluna para exibir. Recarregue um Excel válido.";
        reportTable.appendChild(cap);
        return;
      }

      const thead = document.createElement("thead");
      const trh = document.createElement("tr");
      headers.forEach((h, colIndex) => {
        const th = document.createElement("th");
        th.classList.add("data-table__th--filter");
        th.setAttribute("data-col-index", String(colIndex));
        const wrap = document.createElement("div");
        wrap.className = "th-filter";
        const label = document.createElement("span");
        label.className = "th-filter__label";
        label.textContent = String(h);
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "th-filter__btn";
        btn.setAttribute("aria-label", "Filtrar coluna " + String(h));
        btn.setAttribute("aria-expanded", "false");
        const ic = document.createElement("span");
        ic.className = "th-filter__icon";
        ic.setAttribute("aria-hidden", "true");
        btn.appendChild(ic);
        wrap.appendChild(label);
        wrap.appendChild(btn);
        th.appendChild(wrap);
        trh.appendChild(th);
      });
      thead.appendChild(trh);
      reportTable.appendChild(thead);

      const tbody = document.createElement("tbody");
      reportTable.appendChild(tbody);

      layoutReportScroll();

      const batchSize = 400;
      let i = 0;
      const total = rows.length;

      function step() {
        if (token !== reportRenderToken) return;
        const end = Math.min(i + batchSize, total);
        const frag = document.createDocumentFragment();
        for (; i < end; i++) {
          const tr = document.createElement("tr");
          const row = rows[i];
          headers.forEach((h) => {
            const td = document.createElement("td");
            td.textContent =
              row && typeof row === "object" ? cellText(row[h], h) : "";
            tr.appendChild(td);
          });
          frag.appendChild(tr);
        }
        tbody.appendChild(frag);
        if (frag.childNodes.length > 0) {
          reportLoading.hidden = true;
        }
        if (i < total) {
          requestAnimationFrame(step);
        } else {
          reportLoading.hidden = true;
          setReportBusy(false);
          if (token === reportRenderToken) {
            bindReportTableFilters(thead, tbody);
          }
        }
      }

      step();
    } catch (err) {
      console.error(err);
      reportLoading.hidden = true;
      setReportBusy(false);
      reportTable.innerHTML = "";
      const msg = document.createElement("p");
      msg.className = "report-error";
      msg.textContent =
        "Erro ao montar a tabela: " +
        (err && err.message ? err.message : String(err));
      reportScroll.insertBefore(msg, reportTable);
    }
  }

  function updateReportMeta() {
    const name = loadedFileName ? ` · ${loadedFileName}` : "";
    reportMeta.textContent = `${rows.length.toLocaleString("pt-BR")} linhas · ${headers.length} colunas${name}`;
  }

  function openReportFullscreen() {
    updateReportMeta();
    if (exportWorkspace) exportWorkspace.hidden = true;
    reportFullscreen.hidden = false;
    document.body.classList.add("report-open");
    void reportFullscreen.offsetHeight;
    layoutReportScroll();
    reportResizeHandler = function () {
      layoutReportScroll();
    };
    window.addEventListener("resize", reportResizeHandler);
  }

  function closeReportFullscreen() {
    closeColumnFilterPanel();
    reportTbodyEl = null;
    columnFilters = Object.create(null);
    if (reportResizeHandler) {
      window.removeEventListener("resize", reportResizeHandler);
      reportResizeHandler = null;
    }
    reportFullscreen.hidden = true;
    if (exportWorkspace) exportWorkspace.hidden = false;
    document.body.classList.remove("report-open");
    reportRenderToken += 1;
    reportLoading.hidden = true;
    setReportBusy(false);
  }

  function onFile(file) {
    if (!file) return;
    const name = file.name.toLowerCase();
    if (!name.endsWith(".xlsx") && !name.endsWith(".xls")) {
      alert("Use um arquivo .xlsx ou .xls");
      return;
    }
    loadedFileName = file.name;
    fileNameEl.textContent = file.name;
    fileNameEl.hidden = false;
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        parseWorkbook(e.target.result);
        if (!parsedRows.length || !parsedHeaders.length) {
          if (importConfig) importConfig.hidden = true;
          rows = [];
          headers = [];
          closeReportFullscreen();
          alert(
            "Não encontrei dados na primeira aba da planilha (vazia ou sem linhas de conteúdo)."
          );
          return;
        }
        showImportConfigAfterParse();
        closeReportFullscreen();
      } catch (err) {
        console.error(err);
        parsedRows = [];
        parsedHeaders = [];
        rows = [];
        headers = [];
        if (importConfig) importConfig.hidden = true;
        alert(
          err && err.message
            ? err.message
            : "Erro ao ler o Excel. Verifique o arquivo."
        );
      }
    };
    reader.readAsArrayBuffer(file);
  }

  dropzone.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => {
    const f = fileInput.files && fileInput.files[0];
    onFile(f);
    fileInput.value = "";
  });
  dropzone.addEventListener("dragover", (ev) => {
    ev.preventDefault();
    dropzone.classList.add("dragover");
  });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
  dropzone.addEventListener("drop", (ev) => {
    ev.preventDefault();
    dropzone.classList.remove("dragover");
    const f = ev.dataTransfer.files && ev.dataTransfer.files[0];
    onFile(f);
  });

  btnReport.addEventListener("click", () => {
    if (!parsedRows.length || !parsedHeaders.length) {
      alert("Carregue um Excel com dados primeiro.");
      return;
    }
    if (!buildRowsFromImportSettings()) return;
    openReportFullscreen();
    /* Dois frames: medida da barra confiável antes de posicionar o scroll */
    requestAnimationFrame(() => {
      layoutReportScroll();
      requestAnimationFrame(() => {
        layoutReportScroll();
        renderReportTableBatched();
      });
    });
  });

  if (btnColsAll && importCols) {
    btnColsAll.addEventListener("click", function () {
      importCols.querySelectorAll('input[type="checkbox"][data-header]').forEach(function (cb) {
        cb.checked = true;
      });
    });
  }
  if (btnColsNone && importCols) {
    btnColsNone.addEventListener("click", function () {
      importCols.querySelectorAll('input[type="checkbox"][data-header]').forEach(function (cb) {
        cb.checked = false;
      });
    });
  }
  if (btnAddPrefilter) {
    btnAddPrefilter.addEventListener("click", function () {
      addPrefilterRow();
    });
  }

  btnCloseReport.addEventListener("click", closeReportFullscreen);
})();
