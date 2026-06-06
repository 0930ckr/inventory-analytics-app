const SOURCES = [
  {
    id: "vending",
    label: "자판기 상품",
    spreadsheetId: "1Za7VZOICZZeyPEFpRUxFPsYubT6WtR9jVQ0Yyh9vfRg",
    refSheet: "REF",
    priceColumnIndex: 4,
    salePriceColumnIndex: 5,
    branches: ["평촌점", "강남점", "선릉점", "여의도점", "강남구청역점", "목동역점"],
  },
  {
    id: "consumable",
    label: "경상소모품",
    spreadsheetId: "1E0E8525LczrWl2Mf4lBoOpQ3WGPhMgT9tYQBunJkppw",
    refSheet: "REF",
    priceColumnIndex: 4,
    salePriceColumnIndex: null,
    branches: ["평촌점", "강남점", "선릉점", "여의도점", "강남구청역점", "목동역점"],
  },
];

// Paste a deployed Google Apps Script Web App URL here to share 기타 메모 online.
// Empty value keeps the current browser-only storage fallback.
const OTHER_MEMO_API_URL = "https://script.google.com/macros/s/AKfycbwLlb6nQBD36fbLQSnsfW0AgIBhZpQLUvdG5zN_knxAp-BrLcRxqd4--1hun9jmEKuc/exec";
const LIVE_DATA_API_URL = "";

const state = {
  rows: [],
  filteredRows: [],
  branches: [],
  history: {},
  warehouse: { bySourceCode: {}, basisLabel: "" },
  currentStockBasis: {},
  lastDataStatus: "",
  activeTab: "overview",
  discardSelection: { period: null, branch: null },
  salesSelection: { period: null, branch: null },
  otherSelection: { period: null, branch: null },
  otherNotes: loadOtherNotes(),
  appSettings: loadAppSettings(),
  memoSyncStatus: OTHER_MEMO_API_URL ? "online" : "local",
  activeOtherNote: null,
  exclusionPanelOpen: false,
};

const $ = (id) => document.getElementById(id);
const nf = new Intl.NumberFormat("ko-KR");
const df = new Intl.NumberFormat("ko-KR", { maximumFractionDigits: 1 });
const cf = new Intl.NumberFormat("ko-KR", { style: "currency", currency: "KRW", maximumFractionDigits: 0 });

function qtyText(row, value, formatter = nf) {
  return `${formatter.format(Number(value) || 0)} ${row.unit || "EA"}`;
}

function eaText(value, formatter = nf) {
  return `${formatter.format(Number(value) || 0)} EA`;
}

function stockCell(row) {
  return escapeHtml(qtyText(row, row.stock));
}

function loadOtherNotes() {
  try {
    return JSON.parse(localStorage.getItem("inventory.vendingOtherNotes") || "{}");
  } catch {
    return {};
  }
}

function saveOtherNotes() {
  try {
    localStorage.setItem("inventory.vendingOtherNotes", JSON.stringify(state.otherNotes));
  } catch (error) {
    console.warn("memo save failed", error);
  }
}

function defaultAppSettings() {
  return {
    branchAdds: [],
    branchRenames: {},
    branchHidden: [],
    vendingExclusions: [],
  };
}

function loadAppSettings() {
  try {
    const parsed = JSON.parse(localStorage.getItem("inventory.appSettings") || "{}");
    return {
      ...defaultAppSettings(),
      ...parsed,
      branchAdds: Array.isArray(parsed.branchAdds) ? parsed.branchAdds : [],
      branchRenames: parsed.branchRenames && typeof parsed.branchRenames === "object" ? parsed.branchRenames : {},
      branchHidden: Array.isArray(parsed.branchHidden) ? parsed.branchHidden : [],
      vendingExclusions: Array.isArray(parsed.vendingExclusions) ? parsed.vendingExclusions : [],
    };
  } catch {
    return defaultAppSettings();
  }
}

function saveAppSettings() {
  try {
    localStorage.setItem("inventory.appSettings", JSON.stringify(state.appSettings));
  } catch (error) {
    console.warn("app settings save failed", error);
  }
}

function otherMemoApiUrl(params = {}) {
  if (!OTHER_MEMO_API_URL) return "";
  const url = new URL(OTHER_MEMO_API_URL);
  Object.entries(params).forEach(([key, value]) => url.searchParams.set(key, value));
  return url.href;
}

function requestOtherMemoApi(params) {
  return new Promise((resolve, reject) => {
    if (!OTHER_MEMO_API_URL) {
      resolve({ ok: true, notes: state.otherNotes });
      return;
    }
    const callbackName = `__otherMemoCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement("script");
    const timeout = window.setTimeout(() => {
      cleanup();
      reject(new Error("memo api timeout"));
    }, 15000);
    function cleanup() {
      window.clearTimeout(timeout);
      delete window[callbackName];
      script.remove();
    }
    window[callbackName] = (payload) => {
      cleanup();
      resolve(payload || {});
    };
    script.onerror = () => {
      cleanup();
      reject(new Error("memo api script load failed"));
    };
    script.src = otherMemoApiUrl({ ...params, callback: callbackName, v: Date.now() });
    document.head.appendChild(script);
  });
}

function requestJsonp(url, params = {}) {
  return new Promise((resolve, reject) => {
    if (!url) {
      reject(new Error("missing JSONP URL"));
      return;
    }
    const callbackName = `__jsonpCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement("script");
    const timeout = window.setTimeout(() => {
      cleanup();
      reject(new Error("JSONP request timeout"));
    }, 30000);
    function cleanup() {
      window.clearTimeout(timeout);
      delete window[callbackName];
      script.remove();
    }
    window[callbackName] = (payload) => {
      cleanup();
      resolve(payload || {});
    };
    script.onerror = () => {
      cleanup();
      reject(new Error("JSONP script load failed"));
    };
    const target = new URL(url);
    Object.entries({ ...params, callback: callbackName, v: Date.now() }).forEach(([key, value]) => {
      target.searchParams.set(key, value);
    });
    script.src = target.href;
    document.head.appendChild(script);
  });
}

function setOtherMemoSyncStatus(message, isError = false) {
  const target = $("otherMemoSyncStatus");
  if (!target) return;
  target.textContent = message;
  target.classList.toggle("error", isError);
}

function startOtherMemoSync() {
  if (!OTHER_MEMO_API_URL) return;
  const refresh = async () => {
    if (state.activeOtherNote) return;
    await loadSharedOtherNotes({ silent: true });
    render();
  };
  window.addEventListener("focus", refresh);
  setInterval(refresh, 15000);
}

async function loadSharedOtherNotes({ silent = false } = {}) {
  if (!OTHER_MEMO_API_URL) {
    state.memoSyncStatus = "local";
    setOtherMemoSyncStatus("현재 이 브라우저에만 메모가 저장됩니다. 온라인 공유는 Apps Script URL 설정 후 사용합니다.");
    return;
  }
  try {
    if (!silent) setOtherMemoSyncStatus("온라인 메모를 불러오는 중...");
    const payload = await requestOtherMemoApi({ action: "list" });
    if (payload.ok === false) throw new Error(payload.error || "memo list failed");
    state.otherNotes = payload.notes && typeof payload.notes === "object" ? payload.notes : {};
    saveOtherNotes();
    state.memoSyncStatus = "online";
    if (!silent) setOtherMemoSyncStatus("온라인 공유 메모와 연결됨");
  } catch (error) {
    console.warn("online memo load failed", error);
    state.memoSyncStatus = "fallback";
    setOtherMemoSyncStatus("온라인 메모 연결 실패: Apps Script 배포 권한을 '모든 사용자'로 설정해야 링크 접속자도 같은 메모를 볼 수 있습니다.", true);
  }
}

async function persistOtherMemo(key, label, note) {
  if (note) state.otherNotes[key] = note;
  else delete state.otherNotes[key];
  saveOtherNotes();
  if (!OTHER_MEMO_API_URL) return;
  const result = await requestOtherMemoApi({
    action: note ? "save" : "delete",
    key,
    label,
    note,
    updatedAt: new Date().toISOString(),
  });
  if (result.ok === false) throw new Error(result.error || "memo save failed");
  if (result.notes && typeof result.notes === "object") {
    state.otherNotes = result.notes;
    saveOtherNotes();
  }
}

function otherNoteKey(row) {
  return `${row.branch}__${row.code}`;
}

function openOtherMemo(key, label) {
  state.activeOtherNote = { key, label };
  $("otherMemoTarget").textContent = label;
  $("otherMemoInput").value = state.otherNotes[key] || "";
  $("otherMemoPanel").classList.remove("hidden");
  $("otherMemoInput").focus();
}

document.addEventListener("DOMContentLoaded", () => {
  bindEvents();
  loadSnapshot();
  startOtherMemoSync();
});

function bindEvents() {
  $("reloadButton").addEventListener("click", loadSnapshot);
  $("datasetFilter").addEventListener("change", () => {
    fillBranchFilter();
    render();
  });
  $("branchFilter").addEventListener("change", render);
  $("historyStart").addEventListener("change", () => {
    if ($("historyStart").value > $("historyEnd").value) $("historyEnd").value = $("historyStart").value;
    render();
  });
  $("historyEnd").addEventListener("change", () => {
    if ($("historyEnd").value < $("historyStart").value) $("historyStart").value = $("historyEnd").value;
    render();
  });
  $("safetyMonths").addEventListener("input", render);
  $("deliveryWeeks").addEventListener("input", render);
  $("searchInput").addEventListener("input", render);
  $("exportPurchase").addEventListener("click", exportPurchaseXlsx);
  $("exportDelivery").addEventListener("click", exportDeliveryXlsx);
  $("clearDiscardSelection").addEventListener("click", () => {
    state.discardSelection = { period: null, branch: null };
    render();
  });
  $("clearSalesSelection").addEventListener("click", () => {
    state.salesSelection = { period: null, branch: null };
    render();
  });
  $("clearOtherSelection").addEventListener("click", () => {
    state.otherSelection = { period: null, branch: null };
    render();
  });
  $("vendingSalesMonthChart").addEventListener("click", (event) => {
    const button = event.target.closest("[data-sales-period]");
    if (!button) return;
    state.salesSelection = {
      period: button.getAttribute("data-sales-period") || null,
      branch: null,
    };
    render();
  });
  $("vendingSalesBranchChart").addEventListener("click", (event) => {
    const button = event.target.closest("[data-sales-period]");
    if (!button) return;
    state.salesSelection = {
      period: button.getAttribute("data-sales-period") || null,
      branch: button.getAttribute("data-sales-branch") || null,
    };
    render();
  });
  $("vendingOtherMonthChart").addEventListener("click", (event) => {
    const button = event.target.closest("[data-other-period]");
    if (!button) return;
    state.otherSelection = {
      period: button.getAttribute("data-other-period") || null,
      branch: null,
    };
    render();
  });
  $("vendingOtherBranchChart").addEventListener("click", (event) => {
    const button = event.target.closest("[data-other-period]");
    if (!button) return;
    state.otherSelection = {
      period: button.getAttribute("data-other-period") || null,
      branch: button.getAttribute("data-other-branch") || null,
    };
    render();
  });
  $("vendingOtherBody").addEventListener("click", (event) => {
    const memoCell = event.target.closest("[data-other-note-key]");
    if (!memoCell) return;
    openOtherMemo(memoCell.dataset.otherNoteKey, memoCell.dataset.otherNoteLabel);
  });
  $("vendingOtherBody").addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    const memoCell = event.target.closest("[data-other-note-key]");
    if (!memoCell) return;
    event.preventDefault();
    openOtherMemo(memoCell.dataset.otherNoteKey, memoCell.dataset.otherNoteLabel);
  });
  $("saveOtherMemo").addEventListener("click", async () => {
    if (!state.activeOtherNote) return;
    const value = $("otherMemoInput").value.trim();
    const active = { ...state.activeOtherNote };
    try {
      await persistOtherMemo(active.key, active.label, value);
      setOtherMemoSyncStatus(OTHER_MEMO_API_URL ? "온라인 공유 메모 저장 완료" : "이 브라우저에 메모 저장 완료");
      render();
      openOtherMemo(active.key, active.label);
    } catch (error) {
      console.error(error);
      setOtherMemoSyncStatus("메모 저장 실패: Apps Script 배포 권한을 '모든 사용자'로 설정해야 온라인 공유 저장이 됩니다.", true);
    }
  });
  $("deleteOtherMemo").addEventListener("click", async () => {
    if (!state.activeOtherNote) return;
    const active = { ...state.activeOtherNote };
    try {
      await persistOtherMemo(active.key, active.label, "");
      setOtherMemoSyncStatus(OTHER_MEMO_API_URL ? "온라인 공유 메모 삭제 완료" : "이 브라우저의 메모 삭제 완료");
      render();
      openOtherMemo(active.key, active.label);
    } catch (error) {
      console.error(error);
      setOtherMemoSyncStatus("메모 삭제 실패: Apps Script 배포 권한을 '모든 사용자'로 설정해야 온라인 공유 삭제가 됩니다.", true);
    }
  });
  $("closeOtherMemo").addEventListener("click", () => {
    state.activeOtherNote = null;
    $("otherMemoPanel").classList.add("hidden");
  });
  $("vendingDiscardChart").addEventListener("click", (event) => {
    const button = event.target.closest("[data-discard-period]");
    if (!button) return;
    state.discardSelection = {
      period: button.dataset.discardPeriod || null,
      branch: button.dataset.discardBranch || null,
    };
    render();
  });
  $("vendingDiscardMonthChart").addEventListener("click", (event) => {
    const button = event.target.closest("[data-discard-period]");
    if (!button) return;
    state.discardSelection = {
      period: button.getAttribute("data-discard-period") || null,
      branch: null,
    };
    render();
  });
  $("vendingBranchButtons").addEventListener("change", (event) => {
    if (event.target.id === "excludeCodeInput") {
      fillVendingExclusionProduct();
      return;
    }
  });
  $("vendingBranchButtons").addEventListener("input", (event) => {
    if (event.target.id === "excludeCodeInput") fillVendingExclusionProduct();
  });
  $("vendingBranchButtons").addEventListener("click", handleVendingSettingsClick);
  $("openExclusionPanel").addEventListener("click", () => {
    state.exclusionPanelOpen = true;
    $("exclusionPanel")?.classList.remove("hidden");
  });

  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeTab = button.dataset.tab;
      document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("active", tab === button));
      document.querySelectorAll(".tab-page").forEach((page) => page.classList.toggle("active", page.id === state.activeTab));
      render();
    });
  });
}

function handleVendingSettingsClick(event) {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  if (target.id === "closeExclusionPanel") {
    state.exclusionPanelOpen = false;
    $("exclusionPanel")?.classList.add("hidden");
    return;
  }
  if (target.id === "excludeAddButton") {
    const code = normalizeCode($("excludeCodeInput")?.value);
    const name = $("excludeNameInput")?.value.trim() || "";
    const spec = $("excludeSpecInput")?.value.trim() || "";
    const startDate = $("excludeStartDateInput")?.value || "";
    const editingId = $("excludeEditingId")?.value || "";
    if (!code || !startDate) return;
    if (editingId) {
      state.appSettings.vendingExclusions = state.appSettings.vendingExclusions.map((item) =>
        item.id === editingId ? { ...item, code, name, spec, startDate } : item
      );
    } else {
      const id = `${code}-${startDate}-${Date.now()}`;
      state.appSettings.vendingExclusions.push({ id, code, name, spec, startDate });
    }
    state.exclusionPanelOpen = true;
    saveAppSettings();
    render();
    return;
  }
  const editExclusionId = target.getAttribute("data-exclusion-edit");
  if (editExclusionId) {
    const item = state.appSettings.vendingExclusions.find((candidate) => candidate.id === editExclusionId);
    if (!item) return;
    $("excludeCodeInput").value = item.code || "";
    $("excludeNameInput").value = item.name || "";
    $("excludeSpecInput").value = item.spec || "";
    $("excludeStartDateInput").value = item.startDate || "";
    $("excludeEditingId").value = item.id;
    return;
  }
  const exclusionId = target.getAttribute("data-exclusion-delete");
  if (exclusionId) {
    state.appSettings.vendingExclusions = state.appSettings.vendingExclusions.filter((item) => item.id !== exclusionId);
    state.exclusionPanelOpen = true;
    saveAppSettings();
    render();
  }
}

function clearExclusionInputs() {
  $("excludeCodeInput").value = "";
  $("excludeNameInput").value = "";
  $("excludeSpecInput").value = "";
  $("excludeStartDateInput").value = "";
  $("excludeEditingId").value = "";
}

function fillVendingExclusionProduct() {
  const code = normalizeCode($("excludeCodeInput")?.value);
  if (!code) return;
  const row = state.rows.find((item) => item.sourceId === "vending" && normalizeCode(item.code) === code);
  if (!row) return;
  if (!$("excludeNameInput").value.trim()) $("excludeNameInput").value = row.category || "";
  if (!$("excludeSpecInput").value.trim()) $("excludeSpecInput").value = row.spec || "";
}

async function loadSnapshot() {
  const dataUrl = new URL("./data.json", window.location.href).href;
  const isLocalPreview = ["127.0.0.1", "localhost", ""].includes(window.location.hostname);
  const useLiveDataApi = Boolean(LIVE_DATA_API_URL && !isLocalPreview);
  setStatus(useLiveDataApi ? "구글시트 실시간 데이터 읽는 중..." : "로컬 스냅샷 읽는 중...");
  try {
    const payload = useLiveDataApi ? await requestJsonp(LIVE_DATA_API_URL) : await loadLocalSnapshot(dataUrl);
    applyPayload(payload, useLiveDataApi ? "구글시트 실시간" : "스냅샷");
    await loadSharedOtherNotes();
    render();
  } catch (error) {
    console.error(error);
    if (useLiveDataApi) {
      try {
        setStatus("실시간 연결 실패: 로컬 스냅샷으로 임시 표시합니다.", true);
        const payload = await loadLocalSnapshot(dataUrl);
        applyPayload(payload, "스냅샷");
        await loadSharedOtherNotes();
        render();
        return;
      } catch (fallbackError) {
        console.error(fallbackError);
      }
    }
    setStatus(`data.json을 읽지 못했습니다. GitHub Pages에 index.html과 같은 위치로 data.json을 올렸는지 확인하세요. 확인 주소: ${dataUrl}`, true);
  }
}

async function loadLocalSnapshot(dataUrl) {
  const response = await fetch(`${dataUrl}?v=${Date.now()}`);
  if (!response.ok) {
    throw new Error(`data.json fetch failed: ${response.status}`);
  }
  return response.json();
}

function applyPayload(payload, sourceLabel) {
  if (!payload || !Array.isArray(payload.rows)) {
    throw new Error("payload rows missing");
  }
    state.rows = Array.isArray(payload.rows) ? payload.rows : [];
    state.history = payload.history || {};
    state.warehouse = normalizeWarehousePayload(payload.warehouse);
    state.currentStockBasis = payload.currentStockBasis || {};
    if (payload.sourceMode === "google-sheets-live") normalizeLiveRows();
    state.branches = ["전체", ...Array.from(new Set(state.rows.map((row) => row.branch))).sort()];
    fillBranchFilter();
    const generatedAt = payload.generatedAt ? new Date(payload.generatedAt).toLocaleString("ko-KR") : "시간 정보 없음";
    state.lastDataStatus = `${sourceLabel} ${generatedAt} / ${nf.format(state.rows.length)}개`;
    setStatus(state.lastDataStatus);
}

function normalizeLiveRows() {
  const latestBySource = {
    vending: (state.history.vendingMonths || []).slice().sort().at(-1),
    consumable: (state.history.consumableMonths || []).slice().sort().at(-1),
  };
  state.rows.forEach((row) => {
    const latestPeriod = latestBySource[row.sourceId];
    const latest = (row.monthlyHistory || []).find((month) => month.period === latestPeriod);
    if (!latest) return;
    row.monthlyConsume = latest.consume || 0;
    row.monthlyDiscard = latest.discard || 0;
    row.monthlyOther = latest.other || 0;
    row.consumeAmount = latest.amount || 0;
    row.salesAmount = latest.salesAmount || 0;
    row.discardAmount = (latest.discard || 0) * (row.price || 0);
  });

  const vendingPriceByCode = new Map(
    state.rows
      .filter((row) => row.sourceId === "vending" && row.code)
      .map((row) => [row.code, row.price || 0])
  );
  (state.history.vendingOtherRows || []).forEach((row) => {
    if (!row.price) row.price = vendingPriceByCode.get(row.code) || 0;
  });
}

function normalizeWarehousePayload(warehouse) {
  const bySourceCode = {};
  (warehouse?.rows || []).forEach((row) => {
    const code = normalizeCode(row.code);
    if (!code) return;
    const sourceIds = !row.sourceId || row.sourceId === "all" ? ["vending", "consumable"] : [row.sourceId];
    sourceIds.forEach((sourceId) => {
      const key = `${sourceId}||${code}`;
      const current = bySourceCode[key] || {
        sourceId,
        code,
        category: row.category || "",
        spec: row.spec || "",
        unit: row.unit || "EA",
        stock: 0,
      };
      current.stock += toNumber(row.stock);
      bySourceCode[key] = current;
    });
  });
  return {
    bySourceCode,
    basisLabel: warehouse?.basisLabel || warehouse?.generatedAt || "",
  };
}

function warehouseStockFor(row) {
  return toNumber(state.warehouse.bySourceCode?.[`${row.sourceId}||${row.code}`]?.stock);
}

function originalBranchesForDataset(dataset) {
  return Array.from(new Set(state.rows.filter((row) => row.sourceId === dataset).map((row) => row.branch))).sort((a, b) => a.localeCompare(b, "ko"));
}

function branchDisplayName(branch) {
  return branch;
}

function branchSourceName(displayName) {
  return displayName;
}

function isBranchHidden(branch) {
  return false;
}

function effectiveBranchesForDataset(dataset) {
  return ["\uC804\uCCB4", ...originalBranchesForDataset(dataset)];
}

function applyBranchSettings(row) {
  return row;
}

function latestPeriod(periods) {
  return Array.isArray(periods) && periods.length ? periods.slice().sort().at(-1) : "";
}

function exclusionPeriod(exclusion) {
  return String(exclusion.startDate || "").slice(0, 7);
}

function activeVendingExclusionCodes(analysisPeriods) {
  const latest = latestPeriod(analysisPeriods);
  if (!latest) return new Set();
  return new Set(
    state.appSettings.vendingExclusions
      .filter((exclusion) => normalizeCode(exclusion.code) && exclusionPeriod(exclusion) && exclusionPeriod(exclusion) <= latest)
      .map((exclusion) => normalizeCode(exclusion.code))
  );
}

function isExcludedFromVendingPlan(row, analysisPeriods) {
  return row.sourceId === "vending" && activeVendingExclusionCodes(analysisPeriods).has(normalizeCode(row.code));
}

async function refreshCurrentStocksFromSourceSheets() {
  const overlays = await Promise.all(SOURCES.map((source) => loadSourceCurrentStock(source)));
  const stockMap = new Map();
  state.currentStockBasis = {};

  overlays.forEach((overlay) => {
    if (!overlay) return;
    state.currentStockBasis[overlay.sourceId] = overlay.basisLabel;
    overlay.rows.forEach((row) => {
      stockMap.set([overlay.sourceId, row.branch, row.code].join("||"), row);
    });
  });

  state.rows.forEach((row) => {
    const current = stockMap.get([row.sourceId, row.branch, row.code].join("||"));
    if (!current) return;
    row.stock = current.stock;
    row.actualStock = current.actualStock;
    row.stockGap = current.actualStock === null ? 0 : current.actualStock - current.stock;
    row.currentStockBasisLabel = current.basisLabel;
  });
}

async function loadSourceCurrentStock(source) {
  try {
    const refRows = await fetchSheet(source.spreadsheetId, source.refSheet);
    const itemMap = parseRef(refRows, source);
    const branchSheets = await Promise.all(source.branches.map(async (branch) => ({
      branch,
      rows: await fetchSheet(source.spreadsheetId, branch),
    })));
    const branchBlocks = branchSheets.map(({ branch, rows }) => ({
      branch,
      rows,
      blocks: detectCurrentBlocks(rows, source),
    }));
    const basis = chooseCommonCurrentBasis(branchBlocks);
    const stockRows = branchBlocks.flatMap(({ branch, rows, blocks }) => {
      const block = blocks.find((candidate) => candidate.key === basis.key) || blocks.at(-1);
      return block ? currentRowsForBlock(rows, source, branch, itemMap, block, basis.label) : [];
    });
    return { sourceId: source.id, basisLabel: basis.label, rows: aggregateCurrentRows(stockRows) };
  } catch (error) {
    console.warn(`${source.label} current stock refresh failed`, error);
    return null;
  }
}

function detectCurrentBlocks(rows, source) {
  const dateRow = rows[0] || [];
  const headerRow = rows[1] || [];
  const firstBlockColumn = source.id === "vending" ? 4 : 3;
  const blocks = [];
  for (let col = firstBlockColumn; col < headerRow.length; col += 7) {
    const stockHeader = headerRow[col + 5] || "";
    const actualHeader = headerRow[col + 6] || "";
    const dateInfo = normalizeSheetDate(dateRow[col]);
    if (!dateInfo || !String(actualHeader).includes("실재고") || !stockHeader) continue;
    const actualCount = rows.slice(2).filter((row) => normalizeCode(row[0]) && row[col + 6] !== "").length;
    if (actualCount === 0) continue;
    blocks.push({
      ...dateInfo,
      stock: col + 5,
      actual: col + 6,
      actualCount,
    });
  }
  return blocks;
}

function normalizeSheetDate(value) {
  const text = String(value || "").trim();
  if (!text) return null;
  const dateFunction = text.match(/Date\((\d{4}),\s*(\d{1,2}),\s*(\d{1,2})\)/);
  if (dateFunction) {
    const year = Number(dateFunction[1]);
    const month = Number(dateFunction[2]) + 1;
    const day = Number(dateFunction[3]);
    return sheetDateInfo(year, month, day);
  }
  const normalized = text.replace(/[.]/g, "/").replace(/\s/g, "");
  const parts = normalized.split("/").map((part) => Number(part));
  if (parts.length === 3 && parts.every(Number.isFinite)) {
    const [first, second, third] = parts;
    return first > 31 ? sheetDateInfo(first, second, third) : sheetDateInfo(2000 + third, first, second);
  }
  if (parts.length === 2 && parts.every(Number.isFinite)) {
    return sheetDateInfo(new Date().getFullYear(), parts[0], parts[1]);
  }
  return null;
}

function sheetDateInfo(year, month, day) {
  if (!year || !month || !day) return null;
  const key = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  return {
    key,
    sortValue: new Date(year, month - 1, day).getTime(),
    label: `${month}/${day} 실재고`,
  };
}

function chooseCommonCurrentBasis(branchBlocks) {
  const [first, ...rest] = branchBlocks;
  let common = new Set((first?.blocks || []).map((block) => block.key));
  rest.forEach(({ blocks }) => {
    const keys = new Set(blocks.map((block) => block.key));
    common = new Set([...common].filter((key) => keys.has(key)));
  });
  const candidates = branchBlocks.flatMap(({ blocks }) => blocks).filter((block) => common.has(block.key));
  const chosen = candidates.sort((a, b) => b.sortValue - a.sortValue)[0];
  if (chosen) return { key: chosen.key, label: chosen.label };
  const fallback = branchBlocks
    .flatMap(({ blocks }) => blocks)
    .sort((a, b) => b.sortValue - a.sortValue)[0];
  return fallback ? { key: fallback.key, label: `${fallback.label} 기준` } : { key: "", label: "원본 재고파일" };
}

function currentRowsForBlock(rows, source, branch, itemMap, block, basisLabel) {
  return rows.slice(2).flatMap((row) => {
    const code = normalizeCode(row[0]);
    if (!code) return [];
    const ref = itemMap.get(code) || {};
    const stock = toNumber(row[block.stock]);
    const actualRaw = row[block.actual];
    const actualStock = actualRaw === "" ? null : toNumber(actualRaw);
    return [{
      sourceId: source.id,
      branch,
      code,
      stock: actualStock === null ? stock : actualStock,
      actualStock,
      currentStockBasisLabel: basisLabel,
    }];
  });
}

function aggregateCurrentRows(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const key = [row.branch, row.code].join("||");
    const target = map.get(key) || { ...row, stock: 0, actualStock: null };
    target.stock += row.stock;
    target.actualStock = target.actualStock === null && row.actualStock === null ? null : (target.actualStock || 0) + (row.actualStock || 0);
    map.set(key, target);
  });
  return Array.from(map.values());
}

async function fetchSheet(spreadsheetId, sheetName) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), 12000);
  let response;
  try {
    response = await fetch(url, { credentials: "omit", signal: controller.signal });
  } finally {
    window.clearTimeout(timeout);
  }
  if (!response.ok) {
    throw new Error(`${sheetName} fetch failed: ${response.status}`);
  }
  const text = await response.text();
  if (text.trim().startsWith("<")) {
    throw new Error(`${sheetName} returned HTML instead of CSV`);
  }
  return parseCsv(text);
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (quoted) {
      if (char === '"' && next === '"') {
        cell += '"';
        i += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        cell += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === ",") {
      row.push(cell);
      cell = "";
    } else if (char === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (char !== "\r") {
      cell += char;
    }
  }
  row.push(cell);
  rows.push(row);
  return rows.filter((cells) => cells.some((value) => value !== ""));
}

function parseRef(rows, source) {
  const map = new Map();
  const normalizedHeaders = (rows[0] || []).map((value) => String(value || "").replace(/\s/g, ""));
  const purposeIndex = source.id === "consumable"
    ? ["용도묶음", "용도분류"].map((name) => normalizedHeaders.indexOf(name)).find((index) => index >= 0)
    : -1;
  rows.slice(1).forEach((row) => {
    const code = normalizeCode(row[0]);
    if (!code) return;
    map.set(code, {
      code,
      category: row[1] || "",
      spec: row[2] || "",
      unit: row[3] || "EA",
      price: toNumber(row[source.priceColumnIndex]),
      salePrice: source.salePriceColumnIndex === null ? 0 : toNumber(row[source.salePriceColumnIndex]),
      purposeGroup: purposeIndex >= 0 ? row[purposeIndex] || "" : "",
    });
  });
  return map;
}

function parseBranchSheet(rows, source, branch, itemMap) {
  if (rows.length < 3) return [];
  const headerRow = rows[1] || [];
  const dayBlocks = [];
  const firstBlockColumn = source.id === "vending" ? 4 : 3;
  for (let col = firstBlockColumn; col < headerRow.length; col += 7) {
    const baseHeader = headerRow[col];
    const stockHeader = headerRow[col + 5];
    if (!baseHeader && !stockHeader) continue;
    dayBlocks.push({
      base: col,
      inbound: col + 1,
      consume: col + 2,
      discard: col + 3,
      other: col + 4,
      stock: col + 5,
      actual: col + 6,
    });
  }

  return rows.slice(2).flatMap((row) => {
    const code = normalizeCode(row[0]);
    if (!code) return [];
    const ref = itemMap.get(code) || {};
    const category = ref.category || row[1] || "";
    const spec = ref.spec || row[2] || "";
    const expiry = source.id === "vending" ? row[3] || "" : "";
    const price = toNumber(ref.price);
    const salePrice = toNumber(ref.salePrice);
    const latest = findLatestBlock(row, dayBlocks);
    const monthly = dayBlocks.reduce(
      (sum, block) => {
        sum.inbound += toNumber(row[block.inbound]);
        sum.consume += toNumber(row[block.consume]);
        sum.discard += toNumber(row[block.discard]);
        sum.other += toNumber(row[block.other]);
        return sum;
      },
      { inbound: 0, consume: 0, discard: 0, other: 0 }
    );

    return [{
      sourceId: source.id,
      sourceLabel: source.label,
      branch,
      code,
      category,
      spec,
      purposeGroup: ref.purposeGroup || "",
      expiry,
      unit: ref.unit || "EA",
      price,
      salePrice,
      stock: latest.stock,
      actualStock: latest.actualStock,
      stockGap: latest.actualStock === null ? 0 : latest.actualStock - latest.stock,
      monthlyInbound: monthly.inbound,
      monthlyConsume: monthly.consume,
      monthlyDiscard: monthly.discard,
      monthlyOther: monthly.other,
    }];
  });
}

function findLatestBlock(row, blocks) {
  for (let i = blocks.length - 1; i >= 0; i -= 1) {
    const block = blocks[i];
    const stockRaw = row[block.stock];
    const actualRaw = row[block.actual];
    if (stockRaw !== "" || actualRaw !== "") {
      return {
        stock: toNumber(stockRaw || actualRaw),
        actualStock: actualRaw === "" ? null : toNumber(actualRaw),
      };
    }
  }
  return { stock: 0, actualStock: null };
}

function aggregateRows(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const key = [row.sourceId, row.branch, row.code].join("||");
    const target = map.get(key) || {
      ...row,
      expiryCount: 0,
      stock: 0,
      actualStock: null,
      stockGap: 0,
      monthlyInbound: 0,
      monthlyConsume: 0,
      monthlyDiscard: 0,
      monthlyOther: 0,
    };
    target.expiryCount += row.expiry ? 1 : 0;
    target.stock += row.stock;
    target.actualStock = target.actualStock === null && row.actualStock === null ? null : (target.actualStock || 0) + (row.actualStock || 0);
    target.stockGap += row.stockGap;
    target.monthlyInbound += row.monthlyInbound;
    target.monthlyConsume += row.monthlyConsume;
    target.monthlyDiscard += row.monthlyDiscard;
    target.monthlyOther += row.monthlyOther;
    map.set(key, target);
  });
  return Array.from(map.values()).map((row) => ({
    ...row,
    consumeAmount: row.monthlyConsume * row.price,
    salesAmount: row.monthlyConsume * (row.salePrice || 0),
    discardAmount: row.monthlyDiscard * row.price,
  }));
}

function fillBranchFilter() {
  const dataset = $("datasetFilter")?.value || "consumable";
  const branches = dataset === "all" ? state.branches : effectiveBranchesForDataset(dataset);
  const current = $("branchFilter").value || "전체";
  $("branchFilter").innerHTML = branches.map((branch) => `<option value="${escapeHtml(branch)}">${escapeHtml(branch)}</option>`).join("");
  $("branchFilter").value = branches.includes(current) ? current : branches[0];
}

function render() {
  const dataset = $("datasetFilter").value;
  const vendingMode = dataset === "vending";
  const consumableMode = dataset === "consumable";
  if (!vendingMode && ["delivery", "discard", "other"].includes(state.activeTab)) {
    state.activeTab = "overview";
    document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.tab === state.activeTab));
    document.querySelectorAll(".tab-page").forEach((page) => page.classList.toggle("active", page.id === state.activeTab));
  }
  if (!consumableMode && state.activeTab === "purpose") {
    state.activeTab = "overview";
    document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.tab === state.activeTab));
    document.querySelectorAll(".tab-page").forEach((page) => page.classList.toggle("active", page.id === state.activeTab));
  }
  const branch = $("branchFilter").value || "전체";
  const query = $("searchInput").value.trim().toLowerCase();
  const analysisPeriods = selectedAnalysisPeriods(dataset);
  const analysisMonths = analysisPeriods.length;
  const analysisLabel = describePeriodRange(analysisPeriods);
  const safetyMonths = Math.max(1, toNumber($("safetyMonths").value) || 2);
  const deliveryWeeks = Math.max(1, toNumber($("deliveryWeeks").value) || 2);
  $("safetyLabel").textContent = "구매목표(개월)";
  $("deliveryWeeksField").classList.toggle("hidden", !vendingMode);
  $("deliveryTab").classList.toggle("hidden", !vendingMode);
  $("discardTab").classList.toggle("hidden", !vendingMode);
  $("otherTab").classList.toggle("hidden", !vendingMode);
  $("purposeTab").classList.toggle("hidden", !consumableMode);
  $("purchaseTab").textContent = vendingMode || consumableMode ? "구매필요" : "구매 추천";
  $("overviewTab").textContent = vendingMode ? "판매" : "현황";
  $("consumeMetricLabel").textContent = vendingMode ? "당월 판매 수량" : consumableMode ? "당월 소모 수량" : "월 소모/판매 수량";
  $("consumeAmountMetricLabel").textContent = vendingMode ? "당월 매출" : "월 소모 금액";
  $("purchaseMetricLabel").textContent = vendingMode && state.activeTab === "delivery" ? "배송 필요 수량" : vendingMode || consumableMode ? "구매필요 금액" : "구매 추천 금액";
  $("overviewTitle").textContent = vendingMode ? "자판기 판매 현황" : consumableMode ? "지점별 소모품 현황" : "현재고 현황";
  $("branchAmountTitle").textContent = vendingMode ? "지점별 당월 매출" : "지점별 당월 소모 금액";
  $("topConsumeTitle").textContent = vendingMode ? "판매 상위 품목" : "소모/판매 상위 품목";
  $("categoryAmountTitle").textContent = vendingMode ? "카테고리별 당월 매출" : "카테고리별 소모 금액";
  $("discardAmountTitle").textContent = vendingMode ? "폐기 금액 상위" : "이동 금액 상위";
  $("deliveryTitle").textContent = "배송";
  setVendingHeaders(vendingMode, consumableMode, analysisMonths, safetyMonths);
  renderBranchFilters(dataset, branch, safetyMonths, deliveryWeeks, analysisPeriods);

  state.filteredRows = state.rows
    .filter((row) => dataset === "all" || row.sourceId === dataset)
    .map(applyBranchSettings)
    .filter(Boolean)
    .filter((row) => branch === "전체" || row.branch === branch)
    .filter((row) => {
      if (!query) return true;
      return [row.code, row.category, row.spec, row.purposeGroup, row.branch, row.sourceLabel].join(" ").toLowerCase().includes(query);
    })
    .map((row) => withPurchase(row, safetyMonths, deliveryWeeks, analysisPeriods));

  renderMetrics();
  renderOverview(analysisPeriods);
  renderVendingDiscard(vendingMode, analysisPeriods);
  renderVendingOtherHistory(vendingMode, branch, query, analysisPeriods);
  renderDelivery();
  renderPurchase();
  renderAmount();
  renderPurpose(safetyMonths, analysisPeriods);
  renderExceptions();
}

function renderBranchFilters(dataset, selectedBranch, safetyMonths, deliveryWeeks, analysisPeriods) {
  const panel = $("vendingBranchFilters");
  const vendingMode = dataset === "vending";
  panel.classList.toggle("hidden", !vendingMode);
  $("openExclusionPanel").classList.toggle("hidden", !vendingMode);
  if (!vendingMode) {
    state.exclusionPanelOpen = false;
    $("vendingBranchButtons").innerHTML = "";
    return;
  }
  $("vendingBranchButtons").innerHTML = renderVendingSettingsPanel();
}

function renderVendingSettingsPanel() {
  const exclusions = state.appSettings.vendingExclusions;
  const itemOptions = Array.from(new Map(state.rows
    .filter((row) => row.sourceId === "vending" && row.code)
    .map((row) => [row.code, row]))
    .values())
    .sort((a, b) => a.code.localeCompare(b.code, "ko"));
  return `
    <div id="exclusionPanel" class="exclusion-modal ${state.exclusionPanelOpen ? "" : "hidden"}" role="dialog" aria-modal="true" aria-label="&#54032;&#47588; &#51228;&#50808; &#47785;&#47197;">
      <article class="vending-settings-card">
        <div class="settings-card-title exclusion-modal-title">
          <h3>&#54032;&#47588; &#51228;&#50808; &#47785;&#47197;</h3>
          <button id="closeExclusionPanel" type="button" class="secondary">&#45803;&#44592;</button>
        </div>
        <div class="settings-input-grid exclusion-settings-inputs">
          <label>
            <span>&#54408;&#48264;</span>
            <input id="excludeCodeInput" type="text" list="excludeProductOptions" placeholder="&#54408;&#48264;" />
          </label>
          <label>
            <span>&#54408;&#47749;</span>
            <input id="excludeNameInput" type="text" placeholder="&#49345;&#54408;&#47749;" />
          </label>
          <label>
            <span>&#44508;&#44201;</span>
            <input id="excludeSpecInput" type="text" placeholder="&#44508;&#44201;" />
          </label>
          <label>
            <span>&#54032;&#47588; &#51228;&#50808; &#49884;&#51089;&#51068;</span>
            <input id="excludeStartDateInput" type="date" />
          </label>
          <input id="excludeEditingId" type="hidden" />
          <div class="settings-action-row exclusion-action-row">
            <button id="excludeAddButton" type="button">&#54032;&#47588; &#51228;&#50808;</button>
          </div>
        </div>
        <datalist id="excludeProductOptions">
          ${itemOptions.map((row) => `<option value="${escapeHtml(row.code)}">${escapeHtml(row.category)} ${escapeHtml(row.spec)}</option>`).join("")}
        </datalist>
        <div class="settings-subtitle">&#54032;&#47588; &#51228;&#50808; &#45236;&#50669;</div>
        <div class="settings-list exclusion-list">
          ${exclusions.map((item) => `
            <div class="exclusion-row-card">
              <div>
                <strong>${escapeHtml(item.code)}</strong>
                <span>${escapeHtml(item.name || "-")}</span>
                <small>${escapeHtml(item.spec || "-")}</small>
              </div>
              <em>${escapeHtml(item.startDate || "")}&#48512;&#53552; &#54032;&#47588; &#51228;&#50808;</em>
              <div class="row-actions">
                <button type="button" data-exclusion-edit="${escapeHtml(item.id)}">&#49688;&#51221;</button>
                <button type="button" data-exclusion-delete="${escapeHtml(item.id)}">&#49325;&#51228;</button>
              </div>
            </div>
          `).join("") || `<span class="settings-empty">&#54032;&#47588; &#51228;&#50808; &#49345;&#54408; &#50630;&#51020;</span>`}
        </div>
      </article>
    </div>
  `;
}

function describeHistory(label, periods) {
  if (!Array.isArray(periods) || periods.length === 0) return "";
  const formatPeriod = (period) => {
    const [year, month] = period.split("-");
    return `${year}년 ${Number(month)}월`;
  };
  return `${label}은 ${formatPeriod(periods[0])}부터 ${formatPeriod(periods[periods.length - 1])}까지 ${nf.format(periods.length)}개월 평균`;
}

function describeSelectedHistory(label, periods) {
  if (!Array.isArray(periods) || periods.length === 0) return "";
  return describeHistory(label, periods);
}

function availableHistoryPeriods(dataset) {
  if (dataset === "vending") return state.history.vendingMonths || [];
  if (dataset === "consumable") return state.history.consumableMonths || [];
  return state.history.consumableMonths || state.history.vendingMonths || [];
}

function selectedAnalysisPeriods(dataset) {
  const periods = availableHistoryPeriods(dataset);
  if (!periods.length) return [];
  const startInput = $("historyStart");
  const endInput = $("historyEnd");
  startInput.min = periods[0];
  startInput.max = periods[periods.length - 1];
  endInput.min = periods[0];
  endInput.max = periods[periods.length - 1];
  if (!periods.includes(startInput.value)) startInput.value = periods[0];
  if (!periods.includes(endInput.value)) endInput.value = periods[periods.length - 1];
  const startIndex = periods.indexOf(startInput.value);
  const endIndex = periods.indexOf(endInput.value);
  if (startIndex > endIndex) {
    endInput.value = startInput.value;
    return [startInput.value];
  }
  return periods.slice(startIndex, endIndex + 1);
}

function describePeriodRange(periods) {
  if (!periods.length) return "선택 기간";
  return `${formatPeriodLabel(periods[0])}~${formatPeriodLabel(periods[periods.length - 1])} (${nf.format(periods.length)}개월)`;
}

function formatPeriodLabel(period) {
  const [year, month] = period.split("-");
  return `${year}년 ${Number(month)}월`;
}

function formatPeriodShort(period) {
  const [year, month] = period.split("-");
  return `${year.slice(2)}년 ${Number(month)}월`;
}

function formatDiscardDetailDateLabel(dateText) {
  if (!dateText) return "-";
  const [year, month, day] = String(dateText).split("-");
  if (!year || !month || !day) return String(dateText);
  return `${Number(month)}/${Number(day)}`;
}

function selectedHistory(row, analysisPeriods) {
  const selectedSet = new Set(analysisPeriods);
  const history = (row.monthlyHistory || []).filter((month) => selectedSet.has(month.period));
  const count = history.length;
  const total = (key) => history.reduce((sumValue, month) => sumValue + (Number(month[key]) || 0), 0);
  return {
    count,
    avgConsume: count ? total("consume") / count : null,
    avgDiscard: count ? total("discard") / count : null,
    avgOther: count ? total("other") / count : null,
  };
}

function withPurchase(row, safetyMonths, deliveryWeeks, analysisPeriods) {
  const selected = selectedHistory(row, analysisPeriods);
  const usesHistory = selected.count > 0 && Number.isFinite(selected.avgConsume);
  const planningMonthlyConsume = usesHistory ? selected.avgConsume : row.monthlyConsume;
  const recommendedStock = Math.ceil(planningMonthlyConsume * safetyMonths);
  const purchaseQty = Math.max(0, recommendedStock - row.stock);
  const deliveryMonthlySales = planningMonthlyConsume;
  const deliveryWeeklySales = deliveryMonthlySales / 4.345;
  const deliveryTargetStock = Math.ceil(deliveryWeeklySales * deliveryWeeks);
  const deliveryQty = Math.max(0, deliveryTargetStock - row.stock);
  return {
    ...row,
    avgMonthlyConsume: selected.avgConsume,
    avgMonthlyDiscard: selected.avgDiscard,
    avgMonthlyOther: selected.avgOther,
    analysisMonths: selected.count,
    planningMonthlyConsume,
    planningBasis: usesHistory ? `${selected.count}개월 평균` : "당월",
    recommendedStock,
    purchaseQty,
    purchaseAmount: purchaseQty * row.price,
    deliveryWeeks,
    deliveryMonthlySales,
    deliveryWeeklySales,
    deliveryTargetStock,
    deliveryQty,
    deliveryAmount: deliveryQty * row.price,
  };
}

function renderMetrics() {
  $("totalStockQty").textContent = nf.format(sum(state.filteredRows, "stock"));
  $("totalConsumeQty").textContent = nf.format(sum(state.filteredRows, "monthlyConsume"));
  const revenueKey = $("datasetFilter").value === "vending" ? "salesAmount" : "consumeAmount";
  $("totalConsumeAmount").textContent = cf.format(sum(state.filteredRows, revenueKey));
  if (state.activeTab === "purchase") {
    $("totalPurchaseAmount").textContent = cf.format(sum(purchaseRows(), "purchaseAmount"));
  } else if ($("datasetFilter").value === "vending" && state.activeTab === "delivery") {
    $("totalPurchaseAmount").textContent = eaText(sum(state.filteredRows, "deliveryQty"));
  } else {
    $("totalPurchaseAmount").textContent = cf.format(sum(state.filteredRows, "purchaseAmount"));
  }
}

function renderOverview(analysisPeriods) {
  const vendingMode = $("datasetFilter").value === "vending";
  $("vendingSalesMonthPanel").classList.toggle("hidden", !vendingMode);
  $("vendingSalesBranchPanel").classList.toggle("hidden", !vendingMode);
  $("overviewSummaryPanels").classList.toggle("hidden", vendingMode);
  if (vendingMode) {
    renderVendingSales(analysisPeriods);
    return;
  }
  const rows = [...state.filteredRows].sort(vendingMode
    ? (a, b) => a.code.localeCompare(b.code, "ko") || a.branch.localeCompare(b.branch, "ko")
    : (a, b) => b.consumeAmount - a.consumeAmount || b.monthlyConsume - a.monthlyConsume);
  $("overviewCount").textContent = `${nf.format(rows.length)}건`;
  $("overviewBody").innerHTML = rows.slice(0, 350).map((row) => `
    <tr>
      <td>${typePill(row)}</td>
      <td>${escapeHtml(row.branch)}</td>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${stockCell(row)}</td>
      <td class="number">${escapeHtml(vendingMode ? qtyText(row, row.avgMonthlyConsume, df) : qtyText(row, row.monthlyConsume))}</td>
      ${vendingMode ? "" : `<td class="number">${escapeHtml(qtyText(row, row.monthlyDiscard))}</td>`}
      <td class="number">${cf.format(row.price)}</td>
      ${vendingMode ? `<td class="number">${cf.format(row.salePrice || 0)}</td>` : ""}
      <td class="number">${cf.format(vendingMode ? row.salesAmount : row.consumeAmount)}</td>
    </tr>
  `).join("") || emptyRow(vendingMode ? 11 : 11);

  renderBranchAmounts();
  renderTopConsume();
}

function renderVendingSales(analysisPeriods) {
  const selectedSet = new Set(analysisPeriods);
  if (!selectedSet.has(state.salesSelection.period)) state.salesSelection.period = null;
  const availableBranches = new Set(state.filteredRows.map((row) => row.branch));
  if (state.salesSelection.branch && !availableBranches.has(state.salesSelection.branch)) state.salesSelection.branch = null;

  const rowsWithSales = [...state.filteredRows]
    .map((row) => {
      const salesByPeriod = Object.fromEntries((row.monthlyHistory || [])
        .filter((month) => selectedSet.has(month.period))
        .map((month) => [month.period, {
          qty: Number(month.consume) || 0,
          amount: Number(month.salesAmount) || 0,
        }]));
      const selectedSalesQty = analysisPeriods.reduce((total, period) => total + (salesByPeriod[period]?.qty || 0), 0);
      const selectedSalesAmount = analysisPeriods.reduce((total, period) => total + (salesByPeriod[period]?.amount || 0), 0);
      return { ...row, salesByPeriod, selectedSalesQty, selectedSalesAmount };
    })
    .filter((row) => row.selectedSalesQty > 0 || row.selectedSalesAmount > 0);

  const chartData = buildPeriodBranchChartData(rowsWithSales, analysisPeriods, (row, period) => row.salesByPeriod[period]?.amount || 0);
  renderPeriodLineChart("vendingSalesMonthChart", chartData, {
    selection: state.salesSelection,
    dataPrefix: "sales",
    ariaLabel: "월별 총 매출 곡선 그래프",
    emptyText: "매출 데이터가 없습니다.",
  });
  if (state.salesSelection.period) {
    renderPeriodPieChart("vendingSalesBranchChart", chartData, {
      selection: state.salesSelection,
      dataPrefix: "sales",
      period: state.salesSelection.period,
      label: "매출",
      emptyText: "매출 데이터가 없습니다.",
    });
  } else {
    renderPeriodHeatmap("vendingSalesBranchChart", chartData, {
      selection: state.salesSelection,
      dataPrefix: "sales",
      label: "매출",
      emptyText: "매출 데이터가 없습니다.",
    });
  }

  const detailRows = rowsWithSales
    .map((row) => {
      const detailQty = state.salesSelection.period ? (row.salesByPeriod[state.salesSelection.period]?.qty || 0) : row.selectedSalesQty;
      const detailAmount = state.salesSelection.period ? (row.salesByPeriod[state.salesSelection.period]?.amount || 0) : row.selectedSalesAmount;
      return { ...row, detailSalesQty: detailQty, detailSalesAmount: detailAmount };
    })
    .filter((row) => row.detailSalesQty > 0 || row.detailSalesAmount > 0)
    .filter((row) => !state.salesSelection.branch || row.branch === state.salesSelection.branch)
    .sort((a, b) => b.detailSalesAmount - a.detailSalesAmount || b.detailSalesQty - a.detailSalesQty || a.code.localeCompare(b.code, "ko"));

  const selectedTitle = [
    state.salesSelection.branch,
    state.salesSelection.period ? formatPeriodLabel(state.salesSelection.period) : null,
  ].filter(Boolean).join(" / ");
  $("overviewTitle").textContent = selectedTitle ? `자판기 판매 세부내역 (${selectedTitle})` : "자판기 판매 세부내역";
  $("overviewCount").textContent = `${nf.format(detailRows.length)}건 / 총 ${nf.format(sum(detailRows, "detailSalesQty"))} EA / ${cf.format(sum(detailRows, "detailSalesAmount"))}`;
  $("overviewHeaderRow").innerHTML = `
    <th>지점</th>
    <th>품번</th>
    <th>품명</th>
    <th>규격</th>
    <th>단위</th>
    <th class="number">판매 수량</th>
    <th class="number">구매 단가</th>
    <th class="number">판매 단가</th>
    <th class="number">매출</th>
  `;
  $("overviewBody").innerHTML = detailRows.slice(0, 350).map((row) => `
    <tr>
      <td>${escapeHtml(row.branch)}</td>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${escapeHtml(qtyText(row, row.detailSalesQty))}</td>
      <td class="number">${cf.format(row.price)}</td>
      <td class="number">${cf.format(row.salePrice || 0)}</td>
      <td class="number">${cf.format(row.detailSalesAmount)}</td>
    </tr>
  `).join("") || emptyRow(9);
}

function buildPeriodBranchChartData(rows, analysisPeriods, valueFor) {
  const branches = Array.from(new Set(rows.map((row) => row.branch))).sort((a, b) => a.localeCompare(b, "ko"));
  const cellValues = new Map();
  branches.forEach((branch) => {
    analysisPeriods.forEach((period) => {
      const value = rows
        .filter((row) => row.branch === branch)
        .reduce((total, row) => total + valueFor(row, period), 0);
      cellValues.set(`${branch}__${period}`, value);
    });
  });
  const monthTotals = analysisPeriods.map((period) => ({
    period,
    value: branches.reduce((total, branch) => total + (cellValues.get(`${branch}__${period}`) || 0), 0),
  }));
  return { branches, analysisPeriods, cellValues, monthTotals };
}

function renderPeriodLineChart(targetId, { analysisPeriods, monthTotals }, { selection, dataPrefix, ariaLabel, emptyText }) {
  const formatValue = arguments[2].formatValue || ((value) => cf.format(value));
  const max = Math.max(1, ...monthTotals.map((item) => item.value));
  const width = 920;
  const height = 220;
  const pad = { top: 24, right: 28, bottom: 54, left: 76 };
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const points = monthTotals.map(({ period, value }, index) => {
    const x = pad.left + (analysisPeriods.length === 1 ? plotWidth / 2 : (plotWidth * index) / (analysisPeriods.length - 1));
    const y = pad.top + plotHeight - (value / max) * plotHeight;
    return { period, value, x, y };
  });
  const linePath = points.map((point, index) => `${index ? "L" : "M"} ${point.x.toFixed(1)} ${point.y.toFixed(1)}`).join(" ");
  const areaPath = points.length
    ? `${linePath} L ${points[points.length - 1].x.toFixed(1)} ${pad.top + plotHeight} L ${points[0].x.toFixed(1)} ${pad.top + plotHeight} Z`
    : "";
  const periodAttr = `data-${dataPrefix}-period`;
  $(targetId).innerHTML = points.length ? `
    <svg class="line-chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeHtml(ariaLabel)}">
      <line class="chart-axis" x1="${pad.left}" y1="${pad.top + plotHeight}" x2="${width - pad.right}" y2="${pad.top + plotHeight}"></line>
      <line class="chart-axis" x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + plotHeight}"></line>
      <text class="chart-y-label" x="${pad.left - 10}" y="${pad.top + 4}">${escapeHtml(formatValue(max))}</text>
      <text class="chart-y-label" x="${pad.left - 10}" y="${pad.top + plotHeight + 4}">₩0</text>
      <path class="chart-area" d="${areaPath}"></path>
      <path class="chart-line" d="${linePath}"></path>
      ${points.map((point) => {
        const selected = selection.period === point.period && !selection.branch;
        return `
          <g class="chart-point-group ${selected ? "active" : ""}" tabindex="0" role="button" ${periodAttr}="${escapeHtml(point.period)}" aria-label="${escapeHtml(`${formatPeriodLabel(point.period)} ${cf.format(point.value)}`)}">
            <circle class="chart-point-hit" ${periodAttr}="${escapeHtml(point.period)}" cx="${point.x}" cy="${point.y}" r="16"></circle>
            <circle class="chart-point" ${periodAttr}="${escapeHtml(point.period)}" cx="${point.x}" cy="${point.y}" r="5"></circle>
            <text class="chart-value" ${periodAttr}="${escapeHtml(point.period)}" x="${point.x}" y="${Math.max(14, point.y - 12)}">${escapeHtml(formatValue(point.value))}</text>
            <text class="chart-x-label" ${periodAttr}="${escapeHtml(point.period)}" x="${point.x}" y="${height - 18}">${escapeHtml(formatPeriodShort(point.period))}</text>
          </g>
        `;
      }).join("")}
    </svg>
  ` : `<div class="empty">${escapeHtml(emptyText)}</div>`;
}

function renderPeriodHeatmap(targetId, { branches, analysisPeriods, cellValues, monthTotals }, { selection, dataPrefix, label, emptyText }) {
  const formatValue = arguments[2].formatValue || ((value) => cf.format(value));
  const max = Math.max(1, ...monthTotals.map((item) => item.value), ...Array.from(cellValues.values()));
  const columns = `132px repeat(${analysisPeriods.length}, minmax(86px, 1fr)) 110px`;
  const periodAttr = `data-${dataPrefix}-period`;
  const branchAttr = `data-${dataPrefix}-branch`;
  const isSelected = (period, branch = "") => selection.period === period && (selection.branch || "") === branch;
  $(targetId).style.gridTemplateColumns = columns;
  $(targetId).innerHTML = `
    <div class="heatmap-corner">지점 / 월</div>
    ${analysisPeriods.map((period) => `
      <button class="heatmap-head ${isSelected(period) ? "active" : ""}" ${periodAttr}="${escapeHtml(period)}" ${branchAttr}="" type="button">
        ${escapeHtml(formatPeriodShort(period))}
        <span>${escapeHtml(formatValue(monthTotals.find((item) => item.period === period)?.value || 0))}</span>
      </button>
    `).join("")}
    <div class="heatmap-corner number">합계</div>
    ${branches.map((branch) => {
      const branchTotal = analysisPeriods.reduce((total, period) => total + (cellValues.get(`${branch}__${period}`) || 0), 0);
      return `
        <div class="heatmap-branch">${escapeHtml(branch)}</div>
        ${analysisPeriods.map((period) => {
          const value = cellValues.get(`${branch}__${period}`) || 0;
          const intensity = value ? Math.max(0.12, value / max) : 0;
          return `
            <button
              class="heatmap-cell ${isSelected(period, branch) ? "active" : ""}"
              ${periodAttr}="${escapeHtml(period)}"
              ${branchAttr}="${escapeHtml(branch)}"
              type="button"
              style="--heat:${intensity}"
              title="${escapeHtml(`${branch} ${formatPeriodLabel(period)} ${label} ${formatValue(value)}`)}"
            >
              ${value ? escapeHtml(formatValue(value)) : "-"}
            </button>
          `;
        }).join("")}
        <div class="heatmap-total number">${escapeHtml(formatValue(branchTotal))}</div>
      `;
    }).join("") || `<div class="empty">${escapeHtml(emptyText)}</div>`}
  `;
}

function renderPeriodPieChart(targetId, { branches, cellValues }, { selection, dataPrefix, period, label, emptyText }) {
  const formatValue = arguments[2].formatValue || ((value) => cf.format(value));
  const items = branches
    .map((branch) => ({ branch, value: cellValues.get(`${branch}__${period}`) || 0 }))
    .filter((item) => item.value > 0)
    .sort((a, b) => b.value - a.value);
  const total = items.reduce((sumValue, item) => sumValue + item.value, 0);
  const colors = ["#1b7f5a", "#2f80ed", "#f2994a", "#9b51e0", "#eb5757", "#27ae60", "#56ccf2"];
  const size = 230;
  const center = size / 2;
  const radius = 92;
  let angle = -Math.PI / 2;
  const singleSlice = items.length === 1;
  const slices = items.map((item, index) => {
    const sliceAngle = total ? (item.value / total) * Math.PI * 2 : 0;
    const start = angle;
    const end = angle + sliceAngle;
    angle = end;
    if (singleSlice) return { ...item, color: colors[index % colors.length], ratio: 1, singleSlice: true };
    const x1 = center + Math.cos(start) * radius;
    const y1 = center + Math.sin(start) * radius;
    const x2 = center + Math.cos(end) * radius;
    const y2 = center + Math.sin(end) * radius;
    const largeArc = sliceAngle > Math.PI ? 1 : 0;
    const path = `M ${center} ${center} L ${x1.toFixed(2)} ${y1.toFixed(2)} A ${radius} ${radius} 0 ${largeArc} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} Z`;
    return { ...item, color: colors[index % colors.length], path, ratio: total ? item.value / total : 0 };
  });
  const periodAttr = `data-${dataPrefix}-period`;
  const branchAttr = `data-${dataPrefix}-branch`;
  $(targetId).removeAttribute("style");
  $(targetId).innerHTML = total ? `
    <div class="discard-pie-panel">
      <svg class="discard-pie" viewBox="0 0 ${size} ${size}" role="img" aria-label="${escapeHtml(formatPeriodLabel(period))} 지점별 ${label} 비율">
        ${slices.map((slice) => slice.singleSlice
          ? `<circle class="pie-slice ${selection.branch === slice.branch ? "active" : ""}" cx="${center}" cy="${center}" r="${radius}" fill="${slice.color}" ${periodAttr}="${escapeHtml(period)}" ${branchAttr}="${escapeHtml(slice.branch)}"></circle>`
          : `<path class="pie-slice ${selection.branch === slice.branch ? "active" : ""}" d="${slice.path}" fill="${slice.color}" ${periodAttr}="${escapeHtml(period)}" ${branchAttr}="${escapeHtml(slice.branch)}"></path>`).join("")}
        <circle cx="${center}" cy="${center}" r="48" fill="#fff"></circle>
        <text class="pie-center-title" x="${center}" y="${center - 5}">${escapeHtml(formatPeriodShort(period))}</text>
        <text class="pie-center-value" x="${center}" y="${center + 16}">${escapeHtml(formatValue(total))}</text>
      </svg>
      <div class="pie-legend">
        ${slices.map((slice) => `
          <button class="pie-legend-row ${selection.branch === slice.branch ? "active" : ""}" ${periodAttr}="${escapeHtml(period)}" ${branchAttr}="${escapeHtml(slice.branch)}" type="button">
            <span class="pie-color" style="background:${slice.color}"></span>
            <strong>${escapeHtml(slice.branch)}</strong>
            <span>${df.format(slice.ratio * 100)}%</span>
            <span class="number">${escapeHtml(formatValue(slice.value))}</span>
          </button>
        `).join("")}
      </div>
    </div>
  ` : `<div class="empty">${escapeHtml(formatPeriodLabel(period))} ${escapeHtml(emptyText)}</div>`;
}

function renderVendingDiscard(vendingMode, analysisPeriods) {
  if (!vendingMode) return;
  const selectedSet = new Set(analysisPeriods);
  if (!selectedSet.has(state.discardSelection.period)) state.discardSelection.period = null;
  const availableBranches = new Set(state.filteredRows.map((row) => row.branch));
  if (state.discardSelection.branch && !availableBranches.has(state.discardSelection.branch)) state.discardSelection.branch = null;

  const rowsWithDiscard = [...state.filteredRows]
    .map((row) => {
      const selectedHistory = (row.monthlyHistory || []).filter((month) => selectedSet.has(month.period));
      const discardByPeriod = Object.fromEntries(selectedHistory
        .map((month) => [month.period, Number(month.discard) || 0]));
      const selectedDiscardTotal = analysisPeriods.reduce((total, period) => total + (discardByPeriod[period] || 0), 0);
      const discardDetails = selectedHistory.flatMap((month) => {
        const details = Array.isArray(month.discardDetails) ? month.discardDetails : [];
        if (details.length) {
          return details.map((detail) => ({
            period: month.period,
            date: detail.date || "",
            qty: Number(detail.qty) || 0,
            amount: Number(detail.amount) || ((Number(detail.qty) || 0) * row.price),
          }));
        }
        const qty = Number(month.discard) || 0;
        return qty ? [{ period: month.period, date: "", qty, amount: qty * row.price }] : [];
      });
      return {
        ...row,
        discardByPeriod,
        discardDetails,
        selectedDiscardTotal,
        selectedDiscardAmount: selectedDiscardTotal * row.price,
      };
    })
    .filter((row) => row.selectedDiscardTotal > 0);

  const chartData = buildDiscardChartData(rowsWithDiscard, analysisPeriods);
  renderVendingDiscardMonthChart(chartData);
  renderVendingDiscardChart(chartData);

  const detailRows = rowsWithDiscard
    .flatMap((row) => row.discardDetails
      .filter((detail) => !state.discardSelection.period || detail.period === state.discardSelection.period)
      .map((detail) => ({
        ...row,
        detailDate: detail.date || "",
        detailPeriod: detail.period,
        detailDiscardQty: detail.qty,
        detailDiscardAmount: detail.amount,
      })))
    .filter((row) => row.detailDiscardQty > 0)
    .filter((row) => !state.discardSelection.branch || row.branch === state.discardSelection.branch)
    .sort((a, b) => {
      const dateCompare = (a.detailDate || `${a.detailPeriod}-99`).localeCompare(b.detailDate || `${b.detailPeriod}-99`);
      return dateCompare || a.branch.localeCompare(b.branch, "ko") || a.code.localeCompare(b.code, "ko");
    });

  $("vendingDiscardHeaderRow").innerHTML = `
    <th>날짜</th>
    <th>지점</th>
    <th>품번</th>
    <th>품명</th>
    <th>규격</th>
    <th>단위</th>
    <th class="number">폐기 수량</th>
    <th class="number">구매 단가</th>
    <th class="number">폐기금액</th>
  `;
  const discardTitlePeriod = state.discardSelection.period ? formatPeriodLabel(state.discardSelection.period) : describePeriodRange(analysisPeriods);
  $("vendingDiscardDetailTitle").textContent = `자판기 폐기 세부내역 (${discardTitlePeriod})`;
  $("vendingDiscardBasis").textContent = "";
  $("vendingDiscardBasis").classList.add("hidden");
  $("vendingDiscardCount").textContent = `${nf.format(detailRows.length)}건 / 총 ${nf.format(sum(detailRows, "detailDiscardQty"))} EA / ${cf.format(sum(detailRows, "detailDiscardAmount"))}`;
  $("vendingDiscardBody").innerHTML = detailRows.slice(0, 350).map((row) => `
    <tr>
      <td>${escapeHtml(formatDiscardDetailDateLabel(row.detailDate))}</td>
      <td>${escapeHtml(row.branch)}</td>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${escapeHtml(qtyText(row, row.detailDiscardQty))}</td>
      <td class="number">${cf.format(row.price)}</td>
      <td class="number">${cf.format(row.detailDiscardAmount)}</td>
    </tr>
  `).join("") || emptyRow(9);
}

function buildDiscardChartData(rows, analysisPeriods) {
  const branches = Array.from(new Set(rows.map((row) => row.branch))).sort((a, b) => a.localeCompare(b, "ko"));
  const cellValues = new Map();
  branches.forEach((branch) => {
    analysisPeriods.forEach((period) => {
      const value = rows
        .filter((row) => row.branch === branch)
        .reduce((total, row) => total + ((row.discardByPeriod[period] || 0) * row.price), 0);
      cellValues.set(`${branch}__${period}`, value);
    });
  });
  const monthTotals = analysisPeriods.map((period) => ({
    period,
    value: branches.reduce((total, branch) => total + (cellValues.get(`${branch}__${period}`) || 0), 0),
  }));
  return { branches, analysisPeriods, cellValues, monthTotals };
}

function renderVendingDiscardMonthChart({ analysisPeriods, monthTotals }) {
  const max = Math.max(1, ...monthTotals.map((item) => item.value));
  const width = 920;
  const height = 220;
  const pad = { top: 24, right: 28, bottom: 54, left: 76 };
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const points = monthTotals.map(({ period, value }, index) => {
    const x = pad.left + (analysisPeriods.length === 1 ? plotWidth / 2 : (plotWidth * index) / (analysisPeriods.length - 1));
    const y = pad.top + plotHeight - (value / max) * plotHeight;
    return { period, value, x, y };
  });
  const linePath = points.map((point, index) => `${index ? "L" : "M"} ${point.x.toFixed(1)} ${point.y.toFixed(1)}`).join(" ");
  const areaPath = points.length
    ? `${linePath} L ${points[points.length - 1].x.toFixed(1)} ${pad.top + plotHeight} L ${points[0].x.toFixed(1)} ${pad.top + plotHeight} Z`
    : "";
  $("vendingDiscardMonthChart").innerHTML = points.length ? `
    <svg class="line-chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="월별 총 폐기금액 곡선 그래프">
      <line class="chart-axis" x1="${pad.left}" y1="${pad.top + plotHeight}" x2="${width - pad.right}" y2="${pad.top + plotHeight}"></line>
      <line class="chart-axis" x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + plotHeight}"></line>
      <text class="chart-y-label" x="${pad.left - 10}" y="${pad.top + 4}">${escapeHtml(cf.format(max))}</text>
      <text class="chart-y-label" x="${pad.left - 10}" y="${pad.top + plotHeight + 4}">₩0</text>
      <path class="chart-area" d="${areaPath}"></path>
      <path class="chart-line" d="${linePath}"></path>
      ${points.map((point) => {
        const selected = state.discardSelection.period === point.period && !state.discardSelection.branch;
        return `
          <g class="chart-point-group ${selected ? "active" : ""}" tabindex="0" role="button" data-discard-period="${escapeHtml(point.period)}" aria-label="${escapeHtml(`${formatPeriodLabel(point.period)} 폐기금액 ${cf.format(point.value)}`)}">
            <circle class="chart-point-hit" data-discard-period="${escapeHtml(point.period)}" cx="${point.x}" cy="${point.y}" r="16"></circle>
            <circle class="chart-point" data-discard-period="${escapeHtml(point.period)}" cx="${point.x}" cy="${point.y}" r="5"></circle>
            <text class="chart-value" data-discard-period="${escapeHtml(point.period)}" x="${point.x}" y="${Math.max(14, point.y - 12)}">${escapeHtml(cf.format(point.value))}</text>
            <text class="chart-x-label" data-discard-period="${escapeHtml(point.period)}" x="${point.x}" y="${height - 18}">${escapeHtml(formatPeriodShort(point.period))}</text>
          </g>
        `;
      }).join("")}
    </svg>
  ` : `<div class="empty">폐기 금액 데이터가 없습니다.</div>`;
}

function renderVendingDiscardMonthBars({ monthTotals }) {
  const max = Math.max(1, ...monthTotals.map((item) => item.value));
  return monthTotals.map(({ period, value }) => {
    const selected = state.discardSelection.period === period && !state.discardSelection.branch;
    return `
      <button class="month-bar ${selected ? "active" : ""}" data-discard-period="${escapeHtml(period)}" type="button">
        <span class="month-bar-label">${escapeHtml(formatPeriodShort(period))}</span>
        <span class="month-bar-track"><span class="month-bar-fill" style="width:${Math.max(2, (value / max) * 100)}%"></span></span>
        <strong class="number">${cf.format(value)}</strong>
      </button>
    `;
  }).join("") || `<div class="empty">폐기 금액 데이터가 없습니다.</div>`;
}

function renderVendingDiscardChart({ branches, analysisPeriods, cellValues, monthTotals }) {
  if (state.discardSelection.period) {
    renderVendingDiscardPie({ branches, cellValues, period: state.discardSelection.period });
    return;
  }
  const max = Math.max(1, ...monthTotals.map((item) => item.value), ...Array.from(cellValues.values()));
  const columns = `132px repeat(${analysisPeriods.length}, minmax(86px, 1fr)) 110px`;
  const isSelected = (period, branch = "") => state.discardSelection.period === period && (state.discardSelection.branch || "") === branch;
  $("vendingDiscardChart").style.gridTemplateColumns = columns;
  $("vendingDiscardChart").innerHTML = `
    <div class="heatmap-corner">지점 / 월</div>
    ${analysisPeriods.map((period) => `
      <button class="heatmap-head ${isSelected(period) ? "active" : ""}" data-discard-period="${escapeHtml(period)}" data-discard-branch="" type="button">
        ${escapeHtml(formatPeriodShort(period))}
        <span>${cf.format(monthTotals.find((item) => item.period === period)?.value || 0)}</span>
      </button>
    `).join("")}
    <div class="heatmap-corner number">합계</div>
    ${branches.map((branch) => {
      const branchTotal = analysisPeriods.reduce((total, period) => total + (cellValues.get(`${branch}__${period}`) || 0), 0);
      return `
        <div class="heatmap-branch">${escapeHtml(branch)}</div>
        ${analysisPeriods.map((period) => {
          const value = cellValues.get(`${branch}__${period}`) || 0;
          const intensity = value ? Math.max(0.12, value / max) : 0;
          return `
            <button
              class="heatmap-cell ${isSelected(period, branch) ? "active" : ""}"
              data-discard-period="${escapeHtml(period)}"
              data-discard-branch="${escapeHtml(branch)}"
              type="button"
              style="--heat:${intensity}"
              title="${escapeHtml(`${branch} ${formatPeriodLabel(period)} 폐기금액 ${cf.format(value)}`)}"
            >
              ${value ? cf.format(value) : "-"}
            </button>
          `;
        }).join("")}
        <div class="heatmap-total number">${cf.format(branchTotal)}</div>
      `;
    }).join("") || `<div class="empty">폐기 금액 데이터가 없습니다.</div>`}
  `;
}

function renderVendingDiscardPie({ branches, cellValues, period }) {
  const items = branches
    .map((branch) => ({
      branch,
      value: cellValues.get(`${branch}__${period}`) || 0,
    }))
    .filter((item) => item.value > 0)
    .sort((a, b) => b.value - a.value);
  const total = items.reduce((sumValue, item) => sumValue + item.value, 0);
  const colors = ["#1b7f5a", "#2f80ed", "#f2994a", "#9b51e0", "#eb5757", "#27ae60", "#56ccf2"];
  const size = 230;
  const center = size / 2;
  const radius = 92;
  let angle = -Math.PI / 2;
  const singleSlice = items.length === 1;
  const slices = items.map((item, index) => {
    const sliceAngle = total ? (item.value / total) * Math.PI * 2 : 0;
    const start = angle;
    const end = angle + sliceAngle;
    angle = end;
    if (singleSlice) {
      return { ...item, color: colors[index % colors.length], path: null, ratio: 1, singleSlice: true };
    }
    const x1 = center + Math.cos(start) * radius;
    const y1 = center + Math.sin(start) * radius;
    const x2 = center + Math.cos(end) * radius;
    const y2 = center + Math.sin(end) * radius;
    const largeArc = sliceAngle > Math.PI ? 1 : 0;
    const path = `M ${center} ${center} L ${x1.toFixed(2)} ${y1.toFixed(2)} A ${radius} ${radius} 0 ${largeArc} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} Z`;
    return { ...item, color: colors[index % colors.length], path, ratio: total ? item.value / total : 0 };
  });

  $("vendingDiscardChart").removeAttribute("style");
  $("vendingDiscardChart").innerHTML = total ? `
    <div class="discard-pie-panel">
      <svg class="discard-pie" viewBox="0 0 ${size} ${size}" role="img" aria-label="${escapeHtml(formatPeriodLabel(period))} 지점별 폐기금액 비율">
        ${slices.map((slice) => slice.singleSlice
          ? `
            <circle
              class="pie-slice ${state.discardSelection.branch === slice.branch ? "active" : ""}"
              cx="${center}"
              cy="${center}"
              r="${radius}"
              fill="${slice.color}"
              data-discard-period="${escapeHtml(period)}"
              data-discard-branch="${escapeHtml(slice.branch)}"
            ></circle>
          `
          : `
            <path
              class="pie-slice ${state.discardSelection.branch === slice.branch ? "active" : ""}"
              d="${slice.path}"
              fill="${slice.color}"
              data-discard-period="${escapeHtml(period)}"
              data-discard-branch="${escapeHtml(slice.branch)}"
            ></path>
          `).join("")}
        <circle cx="${center}" cy="${center}" r="48" fill="#fff"></circle>
        <text class="pie-center-title" x="${center}" y="${center - 5}">${escapeHtml(formatPeriodShort(period))}</text>
        <text class="pie-center-value" x="${center}" y="${center + 16}">${escapeHtml(cf.format(total))}</text>
      </svg>
      <div class="pie-legend">
        ${slices.map((slice) => `
          <button
            class="pie-legend-row ${state.discardSelection.branch === slice.branch ? "active" : ""}"
            data-discard-period="${escapeHtml(period)}"
            data-discard-branch="${escapeHtml(slice.branch)}"
            type="button"
          >
            <span class="pie-color" style="background:${slice.color}"></span>
            <strong>${escapeHtml(slice.branch)}</strong>
            <span>${df.format(slice.ratio * 100)}%</span>
            <span class="number">${cf.format(slice.value)}</span>
          </button>
        `).join("")}
      </div>
    </div>
  ` : `<div class="empty">${escapeHtml(formatPeriodLabel(period))} 폐기 금액 데이터가 없습니다.</div>`;
}

function renderPurchase() {
  const rows = purchaseRows().sort((a, b) =>
    b.purchaseAmount - a.purchaseAmount || b.purchaseQty - a.purchaseQty || b.monthlyAverage - a.monthlyAverage || a.code.localeCompare(b.code, "ko"));
  const totalQty = sum(rows, "purchaseQty");
  const totalAmount = sum(rows, "purchaseAmount");
  $("purchaseSummary").innerHTML = `
    <div class="purchase-summary-card">
      <span>구매 필요 품목</span>
      <strong>${nf.format(rows.length)}품목</strong>
    </div>
    <div class="purchase-summary-card">
      <span>구매필요 수량</span>
      <strong>${nf.format(totalQty)} EA</strong>
    </div>
    <div class="purchase-summary-card">
      <span>예상비용</span>
      <strong>${cf.format(totalAmount)}</strong>
    </div>
  `;
  $("purchaseExportHint").textContent = `모든 지점 기말재고와 창고 기말재고를 합산해 계산한 구매필요 품목을 엑셀로 저장합니다.`;
  $("purchaseBody").innerHTML = rows.slice(0, 350).map((row) => `
    <tr>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${escapeHtml(qtyText(row, row.monthlyAverage, df))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.targetStock))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.branchStock))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.warehouseStock || 0))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.totalStock))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.purchaseQty))}</td>
      <td class="number">${cf.format(row.price)}</td>
      <td class="number">${cf.format(row.purchaseAmount)}</td>
    </tr>
  `).join("") || emptyRow(12);
}

function purchaseRows() {
  const dataset = $("datasetFilter").value;
  const query = $("searchInput").value.trim().toLowerCase();
  const safetyMonths = Math.max(1, toNumber($("safetyMonths").value) || 2);
  const analysisPeriods = selectedAnalysisPeriods(dataset);
  const map = new Map();
  state.rows
    .filter((row) => dataset === "all" || row.sourceId === dataset)
    .map(applyBranchSettings)
    .filter(Boolean)
    .filter((row) => !isExcludedFromVendingPlan(row, analysisPeriods))
    .filter((row) => {
      if (!query) return true;
      return [row.code, row.category, row.spec, row.purposeGroup, row.branch, row.sourceLabel].join(" ").toLowerCase().includes(query);
    })
    .map((row) => withPurchase(row, safetyMonths, 2, analysisPeriods))
    .forEach((row) => {
      const key = `${row.sourceId}||${row.code}`;
      const target = map.get(key) || {
      ...row,
      branchStock: 0,
      monthlyAverage: 0,
      planningMonthlyConsume: 0,
      branchCount: 0,
      targetStock: 0,
      warehouseStock: warehouseStockFor(row),
      totalStock: 0,
      purchaseAmount: 0,
    };
      target.branchCount += 1;
      target.branchStock += row.stock;
      target.planningMonthlyConsume += row.planningMonthlyConsume;
      map.set(key, target);
    });
  return Array.from(map.values()).map((row) => {
    const monthlyAverage = row.planningMonthlyConsume;
    const targetStock = Math.ceil(monthlyAverage * safetyMonths);
    const totalStock = row.branchStock + row.warehouseStock;
    const purchaseQty = Math.max(0, targetStock - totalStock);
    return {
      ...row,
      monthlyAverage,
      targetStock,
      totalStock,
      purchaseQty,
      purchaseAmount: purchaseQty * row.price,
    };
  }).filter((row) => row.purchaseQty > 0);
}

function renderDelivery() {
  const deliveryWeeks = Math.max(1, toNumber($("deliveryWeeks").value) || 2);
  const analysisPeriods = selectedAnalysisPeriods("vending");
  const rows = state.filteredRows
    .filter((row) => !isExcludedFromVendingPlan(row, analysisPeriods))
    .filter((row) => row.deliveryQty > 0)
    .sort((a, b) => a.branch.localeCompare(b.branch, "ko") || b.deliveryQty - a.deliveryQty || a.code.localeCompare(b.code, "ko"));
  const branchGroups = groupByBranch(rows);
  const selectedBranch = $("branchFilter").value || "전체";
  const totalQty = rows.reduce((total, row) => total + row.deliveryQty, 0);

  $("deliverySummary").innerHTML = `
    <div class="delivery-summary-card">
      <span>배송 대상 지점</span>
      <strong>${nf.format(branchGroups.length)}곳</strong>
    </div>
    <div class="delivery-summary-card">
      <span>배송 필요 품목</span>
      <strong>${nf.format(rows.length)}품목</strong>
    </div>
    <div class="delivery-summary-card">
      <span>총 배송 수량</span>
      <strong>${nf.format(totalQty)} EA</strong>
    </div>
  `;

  $("deliveryTargetHeader").textContent = `${nf.format(deliveryWeeks)}주 목표`;
  $("deliveryExportHint").textContent = selectedBranch !== "전체"
    ? `${selectedBranch} 배송 필요 품목을 엑셀로 저장합니다.`
    : "현재 표시 중인 배송 필요 품목을 엑셀로 저장합니다.";

  $("deliveryBody").innerHTML = rows.slice(0, 350).map((row) => `
    <tr>
      <td>${escapeHtml(row.branch)}</td>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${escapeHtml(qtyText(row, row.deliveryMonthlySales, df))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.deliveryTargetStock))}</td>
      <td class="number">${stockCell(row)}</td>
      <td class="number delivery-qty">${escapeHtml(qtyText(row, row.deliveryQty))}</td>
    </tr>
  `).join("") || emptyRow(9);
}

function groupByBranch(rows) {
  const map = new Map();
  rows.forEach((row) => {
    if (!map.has(row.branch)) map.set(row.branch, []);
    map.get(row.branch).push(row);
  });
  return Array.from(map.entries()).map(([branch, groupRows]) => ({ branch, rows: groupRows }));
}

function renderVendingOtherHistory(vendingMode, branch, query, analysisPeriods) {
  const panel = $("vendingOtherPanel");
  panel.classList.toggle("hidden", !vendingMode);
  if (!vendingMode) return;
  const selectedPeriods = analysisPeriods;
  const selectedSet = new Set(selectedPeriods);
  if (!selectedSet.has(state.otherSelection.period)) state.otherSelection.period = null;
  const rows = (state.history.vendingOtherRows || [])
    .map(applyBranchSettings)
    .filter(Boolean)
    .filter((row) => branch === "전체" || row.branch === branch)
    .filter((row) => !query || [row.branch, row.code, row.category, row.spec].join(" ").toLowerCase().includes(query))
    .map((row) => {
      const selectedTotal = selectedPeriods.reduce((total, period) => total + (row.otherByPeriod?.[period] || 0), 0);
      const selectedMonthCount = selectedPeriods.filter((period) => (row.otherByPeriod?.[period] || 0) > 0).length;
      return { ...row, selectedTotal, selectedMonthCount, selectedAverage: selectedPeriods.length ? selectedTotal / selectedPeriods.length : 0 };
    })
    .filter((row) => row.selectedTotal > 0);

  const availableBranches = new Set(rows.map((row) => row.branch));
  if (state.otherSelection.branch && !availableBranches.has(state.otherSelection.branch)) state.otherSelection.branch = null;
  const chartData = buildPeriodBranchChartData(rows, selectedPeriods, (row, period) => (row.otherByPeriod?.[period] || 0) * (row.price || 0));
  renderPeriodLineChart("vendingOtherMonthChart", chartData, {
    selection: state.otherSelection,
    dataPrefix: "other",
    ariaLabel: "월별 기타 금액 곡선 그래프",
    emptyText: "기타 데이터가 없습니다.",
  });
  if (state.otherSelection.period) {
    renderPeriodPieChart("vendingOtherBranchChart", chartData, {
      selection: state.otherSelection,
      dataPrefix: "other",
      period: state.otherSelection.period,
      label: "기타 금액",
      emptyText: "기타 데이터가 없습니다.",
    });
  } else {
    renderPeriodHeatmap("vendingOtherBranchChart", chartData, {
      selection: state.otherSelection,
      dataPrefix: "other",
      label: "기타 금액",
      emptyText: "기타 데이터가 없습니다.",
    });
  }

  const detailRows = rows
    .flatMap((row) => {
      const details = Array.isArray(row.otherDetails) ? row.otherDetails : [];
      if (details.length) {
        return details
          .filter((detail) => selectedSet.has(detail.period))
          .filter((detail) => !state.otherSelection.period || detail.period === state.otherSelection.period)
          .map((detail) => ({
            ...row,
            detailDate: detail.date || "",
            detailPeriod: detail.period,
            detailOtherQty: Number(detail.qty) || 0,
            detailOtherAmount: Number(detail.amount) || ((Number(detail.qty) || 0) * (row.price || 0)),
          }));
      }
      const detailQty = state.otherSelection.period ? (row.otherByPeriod?.[state.otherSelection.period] || 0) : row.selectedTotal;
      return [{
        ...row,
        detailDate: "",
        detailPeriod: state.otherSelection.period || "",
        detailOtherQty: detailQty,
        detailOtherAmount: detailQty * (row.price || 0),
      }];
    })
    .filter((row) => row.detailOtherQty > 0)
    .filter((row) => !state.otherSelection.branch || row.branch === state.otherSelection.branch)
    .sort((a, b) => {
      const dateCompare = (a.detailDate || `${a.detailPeriod}-99`).localeCompare(b.detailDate || `${b.detailPeriod}-99`);
      return dateCompare || a.branch.localeCompare(b.branch, "ko") || a.code.localeCompare(b.code, "ko");
    });
  const selectedTitle = [
    state.otherSelection.branch,
    state.otherSelection.period ? formatPeriodLabel(state.otherSelection.period) : null,
  ].filter(Boolean).join(" / ");
  $("vendingOtherTitle").textContent = selectedTitle ? `기타 세부내역 (${selectedTitle})` : "기타 세부내역";
  $("vendingOtherBasis").textContent = selectedTitle
    ? `${selectedTitle} 조건에 해당하는 기타 입력을 구매단가 기준 금액으로 표시합니다. 품목 행을 클릭하면 메모를 남길 수 있습니다.`
    : `${describeSelectedHistory("자판기 상품", selectedPeriods)} 원본 시트의 기타 입력을 구매단가 기준 금액으로 표시합니다. 월 점을 클릭하면 지점별 비율이 원그래프로 표시됩니다.`;
  $("vendingOtherTotalHeader").textContent = "수량";
  $("vendingOtherCount").textContent = `총 ${nf.format(sum(detailRows, "detailOtherQty"))} EA / ${cf.format(sum(detailRows, "detailOtherAmount"))} / ${nf.format(detailRows.length)}품목`;
  $("vendingOtherBody").innerHTML = detailRows.map((row) => {
    const noteKey = otherNoteKey(row);
    const note = state.otherNotes[noteKey] || "";
    const label = `${row.branch} / ${row.code} / ${row.category}`;
    return `
    <tr class="${state.activeOtherNote?.key === noteKey ? "selected-row" : ""}">
      <td>${escapeHtml(formatDiscardDetailDateLabel(row.detailDate))}</td>
      <td>${escapeHtml(row.branch)}</td>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${escapeHtml(qtyText(row, row.detailOtherQty))}</td>
      <td class="number">${cf.format(row.price || 0)}</td>
      <td class="memo-text memo-clickable" data-other-note-key="${escapeHtml(noteKey)}" data-other-note-label="${escapeHtml(label)}" tabindex="0" role="button">${escapeHtml(note || "메모 없음")}</td>
    </tr>
  `;
  }).join("") || emptyRow(9);
}

function setVendingHeaders(vendingMode, consumableMode, analysisMonths, safetyMonths) {
  resetOverviewHeader();
  $("overviewConsumeHeader").textContent = vendingMode ? `${analysisMonths}개월 평균 판매` : "월 소모";
  $("overviewDiscardHeader").textContent = vendingMode ? `${analysisMonths}개월 평균 폐기` : "이동";
  $("overviewDiscardHeader").classList.toggle("hidden", vendingMode);
  $("overviewOtherHeader").textContent = `${analysisMonths}개월 평균 기타`;
  $("overviewOtherHeader").classList.add("hidden");
  $("overviewPriceHeader").textContent = vendingMode ? "구매단가" : "단가";
  $("overviewSalePriceHeader").classList.toggle("hidden", !vendingMode);
  $("purchasePlanningHeader").textContent = vendingMode ? "월 평균 판매" : consumableMode ? "월 평균 소모" : "월 평균 판매/소모";
  $("purchaseTargetHeader").textContent = `${safetyMonths}개월 안전재고 목표`;
  $("purchaseQtyHeader").textContent = "구매필요";
  $("overviewAmountHeader").textContent = vendingMode ? "당월 매출" : "소모 금액";
  $("exceptionConsumeHeader").textContent = vendingMode ? "당월 판매" : "월 소모";
}

function resetOverviewHeader() {
  if ($("overviewConsumeHeader")) return;
  $("overviewHeaderRow").innerHTML = `
    <th>구분</th>
    <th>지점</th>
    <th>품번</th>
    <th>품명</th>
    <th>규격</th>
    <th>단위</th>
    <th class="number">현재고</th>
    <th id="overviewConsumeHeader" class="number">월 소모</th>
    <th id="overviewDiscardHeader" class="number">이동</th>
    <th id="overviewOtherHeader" class="number hidden">평균 기타</th>
    <th id="overviewPriceHeader" class="number">단가</th>
    <th id="overviewSalePriceHeader" class="number hidden">판매단가</th>
    <th id="overviewAmountHeader" class="number">소모 금액</th>
  `;
}

function renderAmount() {
  const vendingMode = $("datasetFilter").value === "vending";
  $("vendingDashboard").classList.toggle("hidden", !vendingMode);
  $("legacyAmountAnalysis").classList.toggle("hidden", vendingMode);
  if (vendingMode) {
    renderVendingDashboard();
    return;
  }
  renderCategoryAmounts();
  const discardRows = [...state.filteredRows].filter((row) => row.discardAmount > 0).sort((a, b) => b.discardAmount - a.discardAmount).slice(0, 10);
  $("discardAmountList").innerHTML = discardRows.map((row, index) => rankRow(index, row.category + " / " + row.branch, row.spec, cf.format(row.discardAmount))).join("") || '<div class="empty">?? ?? ???? ????.</div>';
}

function renderVendingDashboard() {
  const rows = state.filteredRows.filter((row) => row.sourceId === "vending");
  const salesAmount = sum(rows, "salesAmount");
  const discardAmount = sum(rows, "discardAmount");
  const otherAmount = rows.reduce((total, row) => total + vendingOtherAmount(row), 0);
  $("dashboardSummary").innerHTML = [
    dashboardCard("\uB2F9\uC6D4 \uB9E4\uCD9C", cf.format(salesAmount)),
    dashboardCard("\uB2F9\uC6D4 \uD3D0\uAE30\uAE08\uC561", cf.format(discardAmount), "warning"),
    dashboardCard("\uB2F9\uC6D4 \uAE30\uD0C0\uAE08\uC561", cf.format(otherAmount), otherAmount > 0 ? "warning" : ""),
  ].join("");

  const branchRows = Array.from(groupVendingDashboardByBranch(rows).values())
    .sort((a, b) => b.salesAmount - a.salesAmount || b.discardAmount - a.discardAmount || a.branch.localeCompare(b.branch, "ko"));
  $("dashboardBranchBody").innerHTML = branchRows.map((row) => {
    const flags = [];
    if (row.discardAmount > 0) flags.push("\uD3D0\uAE30");
    if (row.otherAmount > 0) flags.push("\uAE30\uD0C0");
    return '\n      <tr>\n        <td><strong>' + escapeHtml(row.branch) + '</strong></td>\n        <td class="number">' + cf.format(row.salesAmount) + '</td>\n        <td class="number">' + cf.format(row.discardAmount) + '</td>\n        <td class="number">' + cf.format(row.otherAmount) + '</td>\n        <td>' + (flags.length ? flags.map((flag) => '<span class="dashboard-flag">' + flag + '</span>').join("") : "-") + '</td>\n      </tr>\n    ';
  }).join("") || emptyRow(5);

  renderDashboardTop("dashboardSalesTop", rows, "salesAmount", (row) => row.monthlyConsume, "\uB9E4\uCD9C \uB370\uC774\uD130\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.");
  renderDashboardTop("dashboardDiscardTop", rows, "discardAmount", (row) => row.monthlyDiscard, "\uD3D0\uAE30 \uB370\uC774\uD130\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.");
  renderDashboardTop("dashboardOtherTop", rows.map((row) => ({ ...row, otherAmount: vendingOtherAmount(row) })), "otherAmount", (row) => row.monthlyOther, "\uAE30\uD0C0 \uB370\uC774\uD130\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.");
}

function dashboardCard(label, value, tone = "") {
  return '<article class="dashboard-card ' + escapeHtml(tone) + '"><span>' + escapeHtml(label) + '</span><strong>' + escapeHtml(value) + '</strong></article>';
}

function groupVendingDashboardByBranch(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const target = map.get(row.branch) || { branch: row.branch, salesAmount: 0, discardAmount: 0, otherAmount: 0 };
    target.salesAmount += row.salesAmount || 0;
    target.discardAmount += row.discardAmount || 0;
    target.otherAmount += vendingOtherAmount(row);
    map.set(row.branch, target);
  });
  return map;
}

function vendingOtherAmount(row) {
  return (Number(row.monthlyOther) || 0) * (Number(row.price) || 0);
}

function renderDashboardTop(targetId, rows, valueKey, qtyFor, emptyText) {
  const ranked = [...rows]
    .filter((row) => (Number(row[valueKey]) || 0) > 0)
    .sort((a, b) => (Number(b[valueKey]) || 0) - (Number(a[valueKey]) || 0) || a.code.localeCompare(b.code, "ko"))
    .slice(0, 10);
  $(targetId).innerHTML = ranked.map((row, index) => {
    const qty = qtyText(row, qtyFor(row));
    return rankRow(index, row.category + " / " + row.branch, row.code + " / " + row.spec + " / " + qty, cf.format(row[valueKey]));
  }).join("") || '<div class="empty">' + escapeHtml(emptyText) + '</div>';
}

function renderPurpose(forecastMonths, analysisPeriods) {
  if ($("datasetFilter").value !== "consumable") return;
  const analysisMonths = analysisPeriods.length;
  const groups = purposeGroups(state.filteredRows, forecastMonths, analysisPeriods);
  $("purposeBasis").textContent = `품번별 재고 계산은 유지하고 같은 역할의 품목 금액만 합산합니다. ${describePeriodRange(analysisPeriods)} 평균 월금액은 현재 등록 단가 기준이며, ${nf.format(forecastMonths)}개월 예상 금액은 그 평균의 ${nf.format(forecastMonths)}개월분입니다.`;
  $("purposeAverageTitle").textContent = `용도별 ${nf.format(analysisMonths)}개월 평균 월금액`;
  $("purposeAverageHeader").textContent = `${nf.format(analysisMonths)}개월 평균 월금액`;
  $("purposeForecastHeader").textContent = `${nf.format(forecastMonths)}개월 예상 금액`;
  $("purposeCount").textContent = `${nf.format(groups.length)}묶음`;
  renderBars("purposeCurrentAmountList", groups.map((group) => ({ label: group.group, value: group.currentAmount })), "당월 소모 금액 데이터가 없습니다.");
  renderBars("purposeAverageAmountList", groups.map((group) => ({ label: group.group, value: group.averageMonthlyAmount })), "과거 소모 금액 데이터가 없습니다.");
  $("purposeBody").innerHTML = groups.map((group) => `
    <tr>
      <td><strong>${escapeHtml(group.group)}</strong></td>
      <td class="number">${nf.format(group.codes.size)}</td>
      <td class="number">${nf.format(group.consumeBranches.size)}</td>
      <td>${escapeHtml(Array.from(group.items).slice(0, 3).join(", "))}</td>
      <td class="number">${cf.format(group.currentAmount)}</td>
      <td class="number">${cf.format(group.averageMonthlyAmount)}</td>
      <td class="number">${cf.format(group.forecastAmount)}</td>
      <td class="number">${nf.format(group.zeroPriceCodes.size)}</td>
    </tr>
  `).join("") || emptyRow(8);
}

function purposeGroups(rows, forecastMonths, analysisPeriods) {
  const map = new Map();
  rows.forEach((row) => {
    const groupName = row.purposeGroup || "기타/검토 필요";
    const group = map.get(groupName) || {
      group: groupName,
      codes: new Set(),
      consumeBranches: new Set(),
      items: new Set(),
      zeroPriceCodes: new Set(),
      currentAmount: 0,
      averageMonthlyAmount: 0,
      forecastAmount: 0,
    };
    group.codes.add(row.code);
    group.items.add(`${row.category} (${row.code})`);
    group.currentAmount += row.consumeAmount || 0;
    const selectedSet = new Set(analysisPeriods);
    const selectedMonths = (row.monthlyHistory || []).filter((month) => selectedSet.has(month.period));
    group.averageMonthlyAmount += selectedMonths.reduce((total, month) => total + (month.amount || 0), 0) / (selectedMonths.length || 1);
    if (row.monthlyConsume > 0) group.consumeBranches.add(row.branch);
    if (row.price === 0 && (row.monthlyConsume > 0 || selectedMonths.some((month) => (month.consume || 0) > 0))) group.zeroPriceCodes.add(row.code);
    map.set(groupName, group);
  });
  return Array.from(map.values())
    .map((group) => ({ ...group, forecastAmount: group.averageMonthlyAmount * forecastMonths }))
    .filter((group) => group.currentAmount > 0 || group.averageMonthlyAmount > 0 || group.zeroPriceCodes.size > 0)
    .sort((a, b) => b.currentAmount - a.currentAmount || b.averageMonthlyAmount - a.averageMonthlyAmount);
}

function renderExceptions() {
  const rows = [];
  state.filteredRows.forEach((row) => {
    if (row.stockGap !== 0) rows.push({ type: "실재고 차이", row });
    if (row.purchaseQty > 0 && row.planningMonthlyConsume > 0) rows.push({ type: ["vending", "consumable"].includes($("datasetFilter").value) ? "구매필요" : "재고 부족", row });
  });
  $("exceptionCount").textContent = `${nf.format(rows.length)}건`;
  $("exceptionBody").innerHTML = rows.slice(0, 350).map(({ type, row }) => `
    <tr>
      <td><span class="pill bad">${escapeHtml(type)}</span></td>
      <td>${typePill(row)}</td>
      <td>${escapeHtml(row.branch)}</td>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${stockCell(row)}</td>
      <td class="number">${escapeHtml(qtyText(row, row.monthlyConsume))}</td>
      <td class="number">${cf.format(row.price)}</td>
    </tr>
  `).join("") || emptyRow(10);
}

function renderBranchAmounts() {
  const vendingMode = $("datasetFilter").value === "vending";
  const valueKey = vendingMode ? "salesAmount" : "consumeAmount";
  const groups = groupSum(state.filteredRows, "branch", valueKey);
  renderBars("branchAmountList", groups, vendingMode ? "판매 금액 데이터가 없습니다." : "소모 금액 데이터가 없습니다.");
}

function renderCategoryAmounts() {
  const vendingMode = $("datasetFilter").value === "vending";
  const valueKey = vendingMode ? "salesAmount" : "consumeAmount";
  const groups = groupSum(state.filteredRows, "category", valueKey);
  renderBars("categoryAmountList", groups, vendingMode ? "판매 금액 데이터가 없습니다." : "카테고리 금액 데이터가 없습니다.");
}

function renderTopConsume() {
  const vendingMode = $("datasetFilter").value === "vending";
  const rows = [...state.filteredRows].filter((row) => row.monthlyConsume > 0).sort((a, b) => b.monthlyConsume - a.monthlyConsume).slice(0, 10);
  const emptyText = vendingMode ? "판매 데이터가 없습니다." : "소모/판매 데이터가 없습니다.";
  $("topConsumeList").innerHTML = rows.map((row, index) => rankRow(index, `${row.category} / ${row.branch}`, row.spec, qtyText(row, row.monthlyConsume))).join("") || `<div class="empty">${emptyText}</div>`;
}

function renderBars(targetId, groups, emptyText) {
  const max = Math.max(1, ...groups.map((group) => group.value));
  $(targetId).innerHTML = groups.slice(0, 12).map((group) => `
    <div class="bar-row">
      <strong>${escapeHtml(group.label || "미분류")}</strong>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(2, (group.value / max) * 100)}%"></div></div>
      <span class="number">${cf.format(group.value)}</span>
    </div>
  `).join("") || `<div class="empty">${escapeHtml(emptyText)}</div>`;
}

function groupSum(rows, labelKey, valueKey) {
  const map = new Map();
  rows.forEach((row) => map.set(row[labelKey] || "미분류", (map.get(row[labelKey] || "미분류") || 0) + row[valueKey]));
  return Array.from(map, ([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value);
}

function rankRow(index, title, subtitle, value) {
  return `
    <div class="rank-row">
      <span class="rank-index">${index + 1}</span>
      <div class="rank-main"><strong>${escapeHtml(title)}</strong><span>${escapeHtml(subtitle || "")}</span></div>
      <strong class="number">${escapeHtml(value)}</strong>
    </div>
  `;
}

function exportPurchaseXlsx() {
  const dataset = $("datasetFilter").value;
  const safetyMonths = Math.max(1, toNumber($("safetyMonths").value) || 2);
  const rows = purchaseRows().sort((a, b) =>
    b.purchaseAmount - a.purchaseAmount || b.purchaseQty - a.purchaseQty || b.monthlyAverage - a.monthlyAverage || a.code.localeCompare(b.code, "ko"));
  const header = ["품번", "품명", "규격", "단위", "월 평균 판매", `${safetyMonths}개월 안전재고 목표`, "모든 지점 기말재고", "창고 기말재고", "총 보유", "구매필요", "단가", "예상비용"];
  const sheetRows = [header, ...rows.map((row) => [
    row.code,
    row.category,
    row.spec,
    row.unit || "EA",
    row.monthlyAverage,
    row.targetStock,
    row.branchStock,
    row.warehouseStock || 0,
    row.totalStock,
    row.purchaseQty,
    row.price,
    row.purchaseAmount,
  ])];
  const baseName = dataset === "vending" ? "자판기_구매필요" : dataset === "consumable" ? "경상소모품_구매필요" : "구매필요";
  const filename = `${baseName}.xlsx`;
  downloadXlsx(filename, "구매필요", sheetRows);
}

function exportDeliveryXlsx() {
  const selectedBranch = $("branchFilter").value || "전체";
  const rows = state.filteredRows.filter((row) => row.deliveryQty > 0);
  const deliveryWeeks = Math.max(1, toNumber($("deliveryWeeks").value) || 2);
  const header = ["지점", "품번", "품명", "규격", "단위", "월 평균 판매", `${deliveryWeeks}주 목표`, "지점 기말재고", "보낼 수량"];
  const sheetRows = [header, ...rows.map((row) => [
    row.branch,
    row.code,
    row.category,
    row.spec,
    row.unit || "EA",
    row.deliveryMonthlySales,
    row.deliveryTargetStock,
    row.stock,
    row.deliveryQty,
  ])];
  const filename = selectedBranch !== "전체" ? `자판기_${selectedBranch}_배송계획.xlsx` : "자판기_배송계획.xlsx";
  downloadXlsx(filename, "배송계획", sheetRows);
}

function csvCell(value) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
}

function downloadXlsx(filename, sheetName, rows) {
  const files = xlsxFiles(sheetName, rows);
  const zipBytes = createZip(files);
  const blob = new Blob([zipBytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
  URL.revokeObjectURL(link.href);
}

function xlsxFiles(sheetName, rows) {
  const sheetXml = worksheetXml(rows);
  const workbookSheetName = escapeXml(sheetName.slice(0, 31) || "Sheet1");
  return [
    {
      name: "[Content_Types].xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
    },
    {
      name: "_rels/.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    },
    {
      name: "xl/workbook.xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${workbookSheetName}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
    },
    { name: "xl/worksheets/sheet1.xml", content: sheetXml },
  ];
}

function worksheetXml(rows) {
  const columnCount = Math.max(1, ...rows.map((row) => row.length));
  const columnWidths = Array.from({ length: columnCount }, (_, index) => {
    const width = Math.min(42, Math.max(10, ...rows.map((row) => String(row[index] ?? "").length + 2)));
    return `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`;
  }).join("");
  const rowXml = rows.map((row, rowIndex) => `
    <row r="${rowIndex + 1}">
      ${row.map((value, colIndex) => worksheetCellXml(value, rowIndex + 1, colIndex + 1)).join("")}
    </row>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>${columnWidths}</cols>
  <sheetData>${rowXml}
  </sheetData>
</worksheet>`;
}

function worksheetCellXml(value, rowNumber, columnNumber) {
  const ref = `${columnName(columnNumber)}${rowNumber}`;
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${ref}"><v>${value}</v></c>`;
  }
  return `<c r="${ref}" t="inlineStr"><is><t>${escapeXml(value ?? "")}</t></is></c>`;
}

function columnName(number) {
  let name = "";
  let current = number;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    current = Math.floor((current - 1) / 26);
  }
  return name;
}

function escapeXml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function createZip(files) {
  const encoder = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  files.forEach((file) => {
    const nameBytes = encoder.encode(file.name);
    const contentBytes = encoder.encode(file.content);
    const crc = crc32(contentBytes);
    const localHeader = zipLocalHeader(nameBytes, contentBytes, crc);
    localParts.push(localHeader, contentBytes);
    centralParts.push(zipCentralHeader(nameBytes, contentBytes, crc, offset));
    offset += localHeader.length + contentBytes.length;
  });
  const centralSize = centralParts.reduce((total, part) => total + part.length, 0);
  const end = zipEndRecord(files.length, centralSize, offset);
  return concatUint8([...localParts, ...centralParts, end]);
}

function zipLocalHeader(nameBytes, contentBytes, crc) {
  const header = new Uint8Array(30 + nameBytes.length);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x04034b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, 0, true);
  view.setUint32(14, crc, true);
  view.setUint32(18, contentBytes.length, true);
  view.setUint32(22, contentBytes.length, true);
  view.setUint16(26, nameBytes.length, true);
  header.set(nameBytes, 30);
  return header;
}

function zipCentralHeader(nameBytes, contentBytes, crc, offset) {
  const header = new Uint8Array(46 + nameBytes.length);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x02014b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 20, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, 0, true);
  view.setUint32(16, crc, true);
  view.setUint32(20, contentBytes.length, true);
  view.setUint32(24, contentBytes.length, true);
  view.setUint16(28, nameBytes.length, true);
  view.setUint32(42, offset, true);
  header.set(nameBytes, 46);
  return header;
}

function zipEndRecord(fileCount, centralSize, centralOffset) {
  const header = new Uint8Array(22);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x06054b50, true);
  view.setUint16(8, fileCount, true);
  view.setUint16(10, fileCount, true);
  view.setUint32(12, centralSize, true);
  view.setUint32(16, centralOffset, true);
  return header;
}

function concatUint8(parts) {
  const totalLength = parts.reduce((total, part) => total + part.length, 0);
  const output = new Uint8Array(totalLength);
  let offset = 0;
  parts.forEach((part) => {
    output.set(part, offset);
    offset += part.length;
  });
  return output;
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let index = 0; index < bytes.length; index += 1) {
    crc ^= bytes[index];
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function typePill(row) {
  return `<span class="pill ${row.sourceId}">${escapeHtml(row.sourceLabel)}</span>`;
}

function emptyRow(colspan) {
  return `<tr><td colspan="${colspan}" class="empty">표시할 데이터가 없습니다.</td></tr>`;
}

function sum(rows, key) {
  return rows.reduce((total, row) => total + (Number(row[key]) || 0), 0);
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  const normalized = String(value).replace(/[,\s원]/g, "");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeCode(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/\.0$/, "");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function setStatus(message, isError = false) {
  $("loadStatus").textContent = message;
  $("loadStatus").classList.toggle("error", isError);
}
