(function () {
  "use strict";

  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const fileNameEl = document.getElementById("fileName");
  const actions = document.getElementById("actions");
  const btnReport = document.getElementById("btnReport");
  const reportBarBusy = document.getElementById("reportBarBusy");
  const reportFullscreen = document.getElementById("reportFullscreen");
  const reportScroll = document.getElementById("reportScroll");
  const reportTable = document.getElementById("reportTable");
  const reportLoading = document.getElementById("reportLoading");
  const reportMeta = document.getElementById("reportMeta");
  const btnCloseReport = document.getElementById("btnCloseReport");
  const mainHeader = document.getElementById("mainHeader");
  const mainHome = document.getElementById("mainHome");

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
      rows = [];
      headers = [];
      return;
    }
    headers = Object.keys(json[0]);
    rows = json;
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
   * Garante dia/mês com dois dígitos em datas tipo 3/10/2025 → 03/10/2025 (pt-BR).
   * Também aceita data com hora depois do texto.
   */
  function padBrazilianDateString(s) {
    const t = String(s).trim();
    const withTime = /^(\d{1,2})([\/\-])(\d{1,2})\2(\d{4})(\s+.+)$/;
    const m2 = t.match(withTime);
    if (m2) {
      const day = m2[1].padStart(2, "0");
      const sep = m2[2];
      const month = m2[3].padStart(2, "0");
      return `${day}${sep}${month}${sep}${m2[4]}${m2[5]}`;
    }
    const dateOnly = /^(\d{1,2})([\/\-])(\d{1,2})\2(\d{4})$/;
    const m = t.match(dateOnly);
    if (m) {
      return `${m[1].padStart(2, "0")}${m[2]}${m[3].padStart(2, "0")}${m[2]}${m[4]}`;
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
      headers.forEach((h) => {
        const th = document.createElement("th");
        th.textContent = String(h);
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
    mainHeader.hidden = true;
    mainHome.hidden = true;
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
    if (reportResizeHandler) {
      window.removeEventListener("resize", reportResizeHandler);
      reportResizeHandler = null;
    }
    reportFullscreen.hidden = true;
    mainHeader.hidden = false;
    mainHome.hidden = false;
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
        if (!rows.length || !headers.length) {
          actions.hidden = true;
          closeReportFullscreen();
          alert(
            "Não encontrei dados na primeira aba da planilha (vazia ou sem linhas de conteúdo)."
          );
          return;
        }
        actions.hidden = false;
        closeReportFullscreen();
      } catch (err) {
        console.error(err);
        rows = [];
        headers = [];
        actions.hidden = true;
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
    if (!rows.length || !headers.length) {
      alert("Carregue um Excel com dados primeiro.");
      return;
    }
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

  btnCloseReport.addEventListener("click", closeReportFullscreen);
})();
