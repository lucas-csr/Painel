/* =====================================================================
   Dashboard "Andamento das Aulas" — busca dados ao vivo do Google Sheets
   via API pública gviz (JSONP), sem precisar de backend nem de chave de API.
   Funciona porque as planilhas foram publicadas na web (Arquivo > Compartilhar
   > Publicar na Web).
   ===================================================================== */

const HOJE = new Date();
HOJE.setHours(0, 0, 0, 0);

// Fim do ano letivo — usado para estimar- quantas aulas ainda faltam
const FIM_ANO_LETIVO = new Date(2026, 11, 10); // 10/12/2026
FIM_ANO_LETIVO.setHours(0, 0, 0, 0);

const MS_POR_SEMANA = 7 * 24 * 60 * 60 * 1000;

// Colunas fixas observadas na estrutura das planilhas:
// col 1 = Trimestre (célula mesclada) | col 2 = Data | col 3 = Atividade
const COL_TRIMESTRE = 1;
const COL_DATA = 2;
const COL_ATIVIDADE = 3;

// Estima quantas aulas semanais ainda vão acontecer entre a última aula
// já cadastrada na planilha (ou hoje, o que for mais tarde) e o fim do ano letivo.
function estimarAulasRestantes(ultimaDataCadastrada) {
  const partida = ultimaDataCadastrada && ultimaDataCadastrada > HOJE ? ultimaDataCadastrada : HOJE;
  if (partida >= FIM_ANO_LETIVO) return 0;
  return Math.max(Math.round((FIM_ANO_LETIVO - partida) / MS_POR_SEMANA), 0);
}

const DATE_REGEX = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/;

let turmasResultado = []; // resultado final, um item por turma

/* ---------------------- Busca via JSONP (gviz) ---------------------- */

function loadSheetJSONP(spreadsheetId, sheetName) {
  return new Promise((resolve) => {
    const callbackName = "gvizCB_" + Math.random().toString(36).slice(2);
    let finished = false;

    const cleanup = () => {
      delete window[callbackName];
      if (script.parentNode) script.parentNode.removeChild(script);
    };

    const timeout = setTimeout(() => {
      if (!finished) {
        finished = true;
        cleanup();
        resolve({ error: "timeout" });
      }
    }, 15000);

    window[callbackName] = function (response) {
      if (finished) return;
      finished = true;
      clearTimeout(timeout);
      cleanup();
      if (!response || response.status === "error") {
        resolve({ error: "sheet_not_found" });
        return;
      }
      resolve({ data: response.table });
    };

    const url =
      `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq` +
      `?sheet=${encodeURIComponent(sheetName)}` +
      `&tqx=responseHandler:${callbackName}`;

    const script = document.createElement("script");
    script.src = url;
    script.onerror = () => {
      if (finished) return;
      finished = true;
      clearTimeout(timeout);
      cleanup();
      resolve({ error: "network" });
    };
    document.body.appendChild(script);
  });
}

/* ---------------------- Parsing das linhas da aba ---------------------- */

function cellText(cell) {
  if (!cell) return "";
  if (cell.f !== undefined && cell.f !== null) return String(cell.f).trim();
  if (cell.v === undefined || cell.v === null) return "";
  if (cell.v instanceof Date) {
    const d = cell.v;
    return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
  }
  return String(cell.v).trim();
}

function parseDateBR(str) {
  const m = DATE_REGEX.exec(str);
  if (!m) return null;
  let [, d, mo, y] = m;
  d = parseInt(d, 10);
  mo = parseInt(mo, 10);
  y = parseInt(y, 10);
  if (y < 100) y += 2000;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  const dt = new Date(y, mo - 1, d);
  dt.setHours(0, 0, 0, 0);
  return dt;
}

function extractLessons(table) {
  const trimestres = {}; // nome -> [ {data, atividade, status} ]
  let currentTrimestre = null;

  (table.rows || []).forEach((row) => {
    const cells = row.c || [];

    // Atualiza o trimestre atual (célula mesclada só aparece na 1ª linha do bloco)
    const triText = cellText(cells[COL_TRIMESTRE]);
    if (triText && /TRIMESTRE/i.test(triText)) {
      currentTrimestre = triText.replace(/^\[merged\]\s*/i, "").trim();
      if (!trimestres[currentTrimestre]) trimestres[currentTrimestre] = [];
    }
    if (triText && /^TOTAL$/i.test(triText.replace(/^\[merged\]\s*/i, "").trim())) {
      return; // linha de total, ignora
    }

    // Tenta achar a data na coluna fixa; se não achar, varre a linha toda
    let dataStr = cellText(cells[COL_DATA]);
    let dt = parseDateBR(dataStr);
    let atividade = cellText(cells[COL_ATIVIDADE]);

    if (!dt) {
      // fallback: procura em qualquer célula da linha
      for (let i = 0; i < cells.length; i++) {
        const t = cellText(cells[i]);
        const tryDt = parseDateBR(t);
        if (tryDt) {
          dt = tryDt;
          // atividade = maior texto entre as outras células, exceto a de trimestre
          let best = "";
          cells.forEach((c, idx) => {
            if (idx === i || idx === COL_TRIMESTRE) return;
            const txt = cellText(c);
            if (txt.length > best.length && !/TRIMESTRE|^TOTAL$/i.test(txt)) best = txt;
          });
          atividade = best;
          break;
        }
      }
    }

    if (!dt || !currentTrimestre) return; // linha sem data válida, ignora
    if (!atividade) atividade = "(sem descrição)";

    const status = dt <= HOJE ? "realizada" : "pendente";
    trimestres[currentTrimestre].push({
      data: `${String(dt.getDate()).padStart(2, "0")}/${String(dt.getMonth() + 1).padStart(2, "0")}/${dt.getFullYear()}`,
      atividade,
      status,
      _sortKey: dt.getTime(),
    });
  });

  // ordena aulas de cada trimestre por data
  Object.values(trimestres).forEach((arr) => arr.sort((a, b) => a._sortKey - b._sortKey));

  return trimestres;
}

/* ---------------------- Orquestração ---------------------- */

async function carregarTudo() {
  setStatus("loading", "Carregando dados das planilhas…");
  const resultados = [];

  for (const cfg of TURMAS_CONFIG) {
    const resp = await loadSheetJSONP(cfg.spreadsheetId, cfg.sheetName);
    if (resp.error) {
      resultados.push({
        ...cfg,
        erro:
          resp.error === "sheet_not_found"
            ? `Aba "${cfg.sheetName}" não encontrada — confira o nome exato no config.`
            : "Não foi possível carregar (rede ou planilha não publicada).",
      });
      continue;
    }
    const trimestres = extractLessons(resp.data);
    let cadastradas = 0,
      realizadas = 0,
      ultimaDataTs = 0;
    Object.values(trimestres).forEach((arr) => {
      cadastradas += arr.length;
      realizadas += arr.filter((a) => a.status === "realizada").length;
      arr.forEach((a) => {
        if (a._sortKey > ultimaDataTs) ultimaDataTs = a._sortKey;
      });
    });
    const ultimaData = ultimaDataTs ? new Date(ultimaDataTs) : null;
    const aulasRestantesEstimadas = estimarAulasRestantes(ultimaData);
    const totalEstimadoAno = cadastradas + aulasRestantesEstimadas;

    resultados.push({
      ...cfg,
      trimestres,
      aulas_cadastradas: cadastradas,
      aulas_realizadas: realizadas,
      aulas_restantes_estimadas: aulasRestantesEstimadas,
      total_aulas: totalEstimadoAno, // total estimado até 10/12/2026
      progresso_pct: totalEstimadoAno ? Math.round((realizadas / totalEstimadoAno) * 1000) / 10 : 0,
    });
  }

  turmasResultado = resultados;
  const comErro = resultados.filter((r) => r.erro).length;
  if (comErro === 0) {
    setStatus("ok", `Dados atualizados agora · posição em ${formatHoje()}`);
  } else {
    setStatus("error", `${comErro} turma(s) com erro ao carregar — veja os cards em vermelho`);
  }
  renderAll();
}

function formatHoje() {
  return `${String(HOJE.getDate()).padStart(2, "0")}/${String(HOJE.getMonth() + 1).padStart(2, "0")}/${HOJE.getFullYear()}`;
}

function setStatus(type, text) {
  const dot = document.getElementById("statusDot");
  dot.className = "status-dot" + (type === "loading" ? " loading" : type === "error" ? " error" : "");
  document.getElementById("statusText").textContent = text;
}

/* ---------------------- Render ---------------------- */

function statusColor(pct) {
  if (pct >= 95) return { cls: "ok", color: "#2e9e6f" };
  if (pct >= 80) return { cls: "warn", color: "#e0a63b" };
  return { cls: "late", color: "#d3564a" };
}

function getFiltered() {
  const turno = document.getElementById("filterTurno").value;
  return turmasResultado.filter((t) => turno === "todos" || t.turno === turno);
}

function renderKPIs() {
  const ok = getFiltered().filter((t) => !t.erro);
  const totalTurmas = ok.length;
  const totalAulas = ok.reduce((s, t) => s + t.total_aulas, 0);
  const totalRealizadas = ok.reduce((s, t) => s + t.aulas_realizadas, 0);
  const media = totalAulas ? (totalRealizadas / totalAulas) * 100 : 0;
  const atencao = ok.filter((t) => t.progresso_pct < 90).length;

  const kpis = [
    { label: "Turmas", value: totalTurmas, sub: "monitoradas" },
    { label: "Aulas já dadas", value: totalRealizadas, sub: `de ~${totalAulas} até 10/12` },
    { label: "Progresso médio do ano", value: media.toFixed(1) + "%", sub: formatHoje() },
    { label: "Turmas em atenção", value: atencao, sub: "progresso < 90% do esperado" },
  ];

  document.getElementById("kpiRow").innerHTML = kpis
    .map(
      (k) => `
    <div class="kpi-card">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value">${k.value}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>`
    )
    .join("");
}

function renderTurmas() {
  const grid = document.getElementById("turmasGrid");
  const sorted = getFiltered().sort((a, b) => {
    if (a.turno !== b.turno) return a.turno === "Manhã" ? -1 : 1;
    return parseInt(a.serie) - parseInt(b.serie);
  });

  grid.innerHTML = sorted
    .map((t) => {
      const badgeClass = t.turno === "Manhã" ? "turno-manha" : "turno-tarde";
      if (t.erro) {
        return `
        <div class="turma-card err">
          <div class="row-top">
            <div class="nome">${t.turma}</div>
            <div class="badge ${badgeClass}">${t.turno}</div>
          </div>
          <div class="err-msg">⚠ ${t.erro}</div>
        </div>`;
      }
      const sc = statusColor(t.progresso_pct);
      return `
      <div class="turma-card" data-turma="${t.turma}" data-turno="${t.turno}">
        <div class="row-top">
          <div class="nome">${t.turma}</div>
          <div class="badge ${badgeClass}">${t.turno}</div>
        </div>
        <div class="progress-bar-bg">
          <div class="progress-bar-fill" style="width:${t.progresso_pct}%; background:${sc.color};"></div>
        </div>
        <div class="stats">
          <span>${t.aulas_realizadas} de ~${t.total_aulas} aulas (estim.)</span>
          <span class="pct ${sc.cls}">${t.progresso_pct}%</span>
        </div>
      </div>`;
    })
    .join("");

  grid.querySelectorAll(".turma-card:not(.err)").forEach((card) => {
    card.addEventListener("click", () => {
      const t = turmasResultado.find(
        (x) => x.turma === card.dataset.turma && x.turno === card.dataset.turno
      );
      if (t) openModal(t);
    });
  });
}

function openModal(turma) {
  document.getElementById("modalTitle").textContent = `${turma.turma} — ${turma.turno}`;
  document.getElementById("modalSub").textContent =
    `${turma.aulas_realizadas} dadas · ~${turma.aulas_restantes_estimadas} faltam até 10/12/2026 · ${turma.progresso_pct}% do ano · posição em ${formatHoje()}`;

  let html = "";
  Object.entries(turma.trimestres).forEach(([nome, aulas]) => {
    const realizadas = aulas.filter((a) => a.status === "realizada").length;
    html += `<div style="margin-bottom:16px;">
      <h4 style="font-size:12.5px;text-transform:uppercase;letter-spacing:.4px;color:#647184;margin:0 0 8px;border-bottom:1px solid #e7eaf0;padding-bottom:6px;">
        ${nome} ${aulas.length ? `(${realizadas}/${aulas.length})` : "(sem registros)"}
      </h4>`;
    if (aulas.length === 0) {
      html += `<div class="lesson-row"><span class="lesson-ativ">Nenhuma aula cadastrada ainda.</span></div>`;
    } else {
      aulas.forEach((a) => {
        html += `
        <div class="lesson-row ${a.status}">
          <span class="lesson-dot ${a.status}"></span>
          <span class="lesson-data">${a.data}</span>
          <span class="lesson-ativ">${a.atividade}</span>
        </div>`;
      });
    }
    html += `</div>`;
  });

  document.getElementById("modalBody").innerHTML = html;
  document.getElementById("modalOverlay").classList.add("active");
}

function renderAll() {
  renderKPIs();
  renderTurmas();
}

/* ---------------------- Eventos ---------------------- */

document.getElementById("closeModal").addEventListener("click", () => {
  document.getElementById("modalOverlay").classList.remove("active");
});
document.getElementById("modalOverlay").addEventListener("click", (e) => {
  if (e.target.id === "modalOverlay") document.getElementById("modalOverlay").classList.remove("active");
});
document.getElementById("filterTurno").addEventListener("change", renderAll);
document.getElementById("refreshBtn").addEventListener("click", carregarTudo);

carregarTudo();
