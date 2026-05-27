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
const OTHER_MEMO_API_URL = "https://script.google.com/macros/s/AKfycbxrvMafHyYAHVAJnv2ev_O9KPdNfa1ZVsiRq65CxzWYorzK8-8ZmbZy1-IfcH7ShFY/exec";

const state = {
  rows: [],
  filteredRows: [],
  branches: [],
  history: {},
  activeTab: "overview",
  discardSelection: { period: null, branch: null },
  salesSelection: { period: null, branch: null },
  otherSelection: { period: null, branch: null },
  otherNotes: loadOtherNotes(),
  memoSyncStatus: OTHER_MEMO_API_URL ? "online" : "local",
  activeOtherNote: null,
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
    setOtherMemoSyncStatus("온라인 메모 연결 실패: 이 브라우저 저장값으로 임시 표시합니다.", true);
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
  $("datasetFilter").addEventListener("change", render);
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
  $("exportPurchase").addEventListener("click", exportPurchaseCsv);
  $("exportDelivery").addEventListener("click", exportDeliveryCsv);
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
    const button = event.target.closest("[data-other-note-key]");
    if (!button) return;
    openOtherMemo(button.dataset.otherNoteKey, button.dataset.otherNoteLabel);
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
      setOtherMemoSyncStatus("메모 저장 실패: 연결 상태를 확인하세요.", true);
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
      setOtherMemoSyncStatus("메모 삭제 실패: 연결 상태를 확인하세요.", true);
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
  $("vendingBranchButtons").addEventListener("click", (event) => {
    const button = event.target.closest("[data-branch]");
    if (!button) return;
    $("branchFilter").value = button.dataset.branch;
    render();
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

async function loadSnapshot() {
  setStatus("로컬 스냅샷 읽는 중...");
  const dataUrl = new URL("./data.json", window.location.href).href;
  try {
    const response = await fetch(`${dataUrl}?v=${Date.now()}`);
    if (!response.ok) {
      throw new Error(`data.json fetch failed: ${response.status}`);
    }
    const payload = await response.json();
    state.rows = Array.isArray(payload.rows) ? payload.rows : [];
    state.history = payload.history || {};
    state.branches = ["전체", ...Array.from(new Set(state.rows.map((row) => row.branch))).sort()];
    fillBranchFilter();
    const generatedAt = payload.generatedAt ? new Date(payload.generatedAt).toLocaleString("ko-KR") : "시간 정보 없음";
    setStatus(`스냅샷 ${generatedAt} / ${nf.format(state.rows.length)}개`);
    await loadSharedOtherNotes();
    render();
  } catch (error) {
    console.error(error);
    setStatus(`data.json을 읽지 못했습니다. GitHub Pages에 index.html과 같은 위치로 data.json을 올렸는지 확인하세요. 확인 주소: ${dataUrl}`, true);
  }
}

async function fetchSheet(spreadsheetId, sheetName) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
  const response = await fetch(url, { credentials: "include" });
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
  $("branchFilter").innerHTML = state.branches.map((branch) => `<option value="${escapeHtml(branch)}">${escapeHtml(branch)}</option>`).join("");
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
  $("safetyLabel").textContent = vendingMode ? "구매 목표(개월)" : consumableMode ? "예측 기간(개월)" : "안전재고";
  $("deliveryWeeksField").classList.toggle("hidden", !vendingMode);
  $("deliveryTab").classList.toggle("hidden", !vendingMode);
  $("discardTab").classList.toggle("hidden", !vendingMode);
  $("otherTab").classList.toggle("hidden", !vendingMode);
  $("purposeTab").classList.toggle("hidden", !consumableMode);
  $("purchaseTab").textContent = vendingMode || consumableMode ? "구매필요" : "구매 추천";
  $("overviewTab").textContent = vendingMode ? "판매" : "현황";
  $("consumeMetricLabel").textContent = vendingMode ? "당월 판매 수량" : consumableMode ? "당월 소모 수량" : "월 소모/판매 수량";
  $("consumeAmountMetricLabel").textContent = vendingMode ? "당월 매출" : "월 소모 금액";
  $("purchaseMetricLabel").textContent = vendingMode && state.activeTab === "delivery" ? "하이랙스 배송 필요 금액" : vendingMode || consumableMode ? "구매필요 금액" : "구매 추천 금액";
  $("overviewTitle").textContent = vendingMode ? "자판기 판매 현황" : consumableMode ? "지점별 소모품 현황" : "현재고 현황";
  $("branchAmountTitle").textContent = vendingMode ? "지점별 당월 매출" : "지점별 당월 소모 금액";
  $("topConsumeTitle").textContent = vendingMode ? "판매 상위 품목" : "소모/판매 상위 품목";
  $("categoryAmountTitle").textContent = vendingMode ? "카테고리별 당월 매출" : "카테고리별 소모 금액";
  $("discardAmountTitle").textContent = vendingMode ? "폐기 금액 상위" : "이동 금액 상위";
  $("purchaseTitle").textContent = vendingMode
    ? `구매필요 계획 (기준: ${analysisLabel} 평균 판매, 목표: ${nf.format(safetyMonths)}개월분)`
    : consumableMode
      ? `구매필요 계획 (기준: ${analysisLabel} 평균 소모, 목표: ${nf.format(safetyMonths)}개월분)`
      : `${nf.format(safetyMonths)}개월 안전재고 기준 구매 추천`;
  $("deliveryTitle").textContent = `하이랙스 창고 배송 계획 (주기: ${nf.format(deliveryWeeks)}주)`;
  $("deliveryBasis").textContent = `당월 판매량을 4주 기준으로 환산해, 현재 지점 재고를 다음 ${nf.format(deliveryWeeks)}주 판매 예상량까지 채우기 위한 배송량입니다. 하이랙스 창고의 보유 수량은 아직 반영하지 않습니다.`;
  const historyDescriptions = [
    describeHistory("경상소모품", state.history.consumableMonths),
    describeHistory("자판기 상품", state.history.vendingMonths),
  ].filter(Boolean);
  $("purchaseBasis").textContent = vendingMode
    ? `품번이 같은 상품만 동일 품목으로 집계합니다. ${describeSelectedHistory("자판기 상품", analysisPeriods)} 판매량을 기준으로, 현재 재고에서 추가로 확보할 구매필요 수량입니다.${branch === "전체" ? " 전체 보기에서는 품번별로 필요한 지점의 수량을 합산합니다." : ""}`
    : consumableMode
      ? `${describeSelectedHistory("경상소모품", analysisPeriods)} 소모량을 기준으로, 각 지점의 현재 재고에서 부족한 ${nf.format(safetyMonths)}개월 예상 소모분을 구매필요로 계산합니다.`
    : historyDescriptions.length
      ? `${historyDescriptions.join(", ")} 소모량을 기준으로 계산합니다. 입고예정수량은 아직 0으로 계산합니다.`
      : "당월 소모량을 기준으로 계산합니다. 입고예정수량은 아직 0으로 계산합니다.";
  setVendingHeaders(vendingMode, consumableMode, analysisMonths, safetyMonths);
  renderBranchFilters(dataset, branch, safetyMonths, deliveryWeeks, analysisPeriods);

  state.filteredRows = state.rows
    .filter((row) => dataset === "all" || row.sourceId === dataset)
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
  const visible = dataset === "vending" || dataset === "consumable";
  panel.classList.toggle("hidden", !visible);
  if (!visible) return;
  const vendingMode = dataset === "vending";
  $("branchFilterTitle").textContent = vendingMode ? "지점별 보기" : "지점별 소모 금액 및 구매필요";
  $("branchFilterDescription").textContent = vendingMode
    ? "지점을 선택하면 아래 판매·폐기, 배송 및 구매필요 계획이 해당 지점 기준으로 표시됩니다."
    : `각 지점의 당월 소모 금액과 ${nf.format(safetyMonths)}개월 예상 소모량 기준 구매필요 금액을 비교할 수 있습니다.`;

  const rows = state.rows.filter((row) => row.sourceId === dataset);
  const branchNames = ["전체", ...Array.from(new Set(rows.map((row) => row.branch))).sort((a, b) => a.localeCompare(b, "ko"))];
  $("vendingBranchButtons").innerHTML = branchNames.map((branch) => {
    const branchRows = branch === "전체" ? rows : rows.filter((row) => row.branch === branch);
    const plannedRows = branchRows.map((row) => withPurchase(row, safetyMonths, deliveryWeeks, analysisPeriods));
    const isActive = selectedBranch === branch;
    return `
      <button class="branch-card ${isActive ? "active" : ""}" data-branch="${escapeHtml(branch)}" type="button" aria-pressed="${isActive}">
        <strong>${escapeHtml(branch)}</strong>
        ${vendingMode
          ? `<span>당월 판매 ${nf.format(sum(branchRows, "monthlyConsume"))}</span><span>매출 ${cf.format(sum(branchRows, "salesAmount"))}</span>`
          : `<span>소모금액 ${cf.format(sum(branchRows, "consumeAmount"))}</span><span>구매필요 ${cf.format(sum(plannedRows, "purchaseAmount"))}</span>`}
      </button>
    `;
  }).join("");
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
  const deliveryWeeklySales = row.monthlyConsume / 4;
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
  const metricKey = $("datasetFilter").value === "vending" && state.activeTab === "delivery" ? "deliveryAmount" : "purchaseAmount";
  $("totalPurchaseAmount").textContent = cf.format(sum(state.filteredRows, metricKey));
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
      <td class="number">${escapeHtml(qtyText(row, row.stock))}</td>
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
    <th>구분</th>
    <th>지점</th>
    <th>품번</th>
    <th>품명</th>
    <th>규격</th>
    <th>단위</th>
    <th class="number">현재고</th>
    <th class="number">판매 수량</th>
    <th class="number">구매단가</th>
    <th class="number">판매단가</th>
    <th class="number">매출</th>
  `;
  $("overviewBody").innerHTML = detailRows.slice(0, 350).map((row) => `
    <tr>
      <td>${typePill(row)}</td>
      <td>${escapeHtml(row.branch)}</td>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${escapeHtml(qtyText(row, row.stock))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.detailSalesQty))}</td>
      <td class="number">${cf.format(row.price)}</td>
      <td class="number">${cf.format(row.salePrice || 0)}</td>
      <td class="number">${cf.format(row.detailSalesAmount)}</td>
    </tr>
  `).join("") || emptyRow(11);
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
      const discardByPeriod = Object.fromEntries((row.monthlyHistory || [])
        .filter((month) => selectedSet.has(month.period))
        .map((month) => [month.period, Number(month.discard) || 0]));
      const selectedDiscardTotal = analysisPeriods.reduce((total, period) => total + (discardByPeriod[period] || 0), 0);
      return {
        ...row,
        discardByPeriod,
        selectedDiscardTotal,
        selectedDiscardAmount: selectedDiscardTotal * row.price,
      };
    })
    .filter((row) => row.selectedDiscardTotal > 0);

  const chartData = buildDiscardChartData(rowsWithDiscard, analysisPeriods);
  renderVendingDiscardMonthChart(chartData);
  renderVendingDiscardChart(chartData);

  const detailRows = rowsWithDiscard
    .map((row) => {
      const detailQty = state.discardSelection.period ? (row.discardByPeriod[state.discardSelection.period] || 0) : row.selectedDiscardTotal;
      return {
        ...row,
        detailDiscardQty: detailQty,
        detailDiscardAmount: detailQty * row.price,
      };
    })
    .filter((row) => row.detailDiscardQty > 0)
    .filter((row) => !state.discardSelection.branch || row.branch === state.discardSelection.branch)
    .sort((a, b) => b.detailDiscardAmount - a.detailDiscardAmount || b.detailDiscardQty - a.detailDiscardQty || a.code.localeCompare(b.code, "ko"));

  const selectedTitle = [
    state.discardSelection.branch,
    state.discardSelection.period ? formatPeriodLabel(state.discardSelection.period) : null,
  ].filter(Boolean).join(" / ");
  $("vendingDiscardHeaderRow").innerHTML = `
    <th>구분</th>
    <th>지점</th>
    <th>품번</th>
    <th>품명</th>
    <th>규격</th>
    <th>단위</th>
    <th class="number">현재고</th>
    <th class="number">폐기 수량</th>
    <th class="number">구매단가</th>
    <th class="number">폐기 금액</th>
  `;
  $("vendingDiscardDetailTitle").textContent = selectedTitle ? `자판기 폐기 세부내역 (${selectedTitle})` : "자판기 폐기 세부내역";
  $("vendingDiscardBasis").textContent = selectedTitle
    ? `${selectedTitle} 조건에 해당하는 폐기만 표시합니다. 폐기 금액은 구매단가 기준입니다.`
    : `${describePeriodRange(analysisPeriods)} 전체 폐기 합계입니다. 그래프의 월/지점 셀을 클릭하면 세부내역이 좁혀집니다.`;
  $("vendingDiscardCount").textContent = `${nf.format(detailRows.length)}건 / 총 ${nf.format(sum(detailRows, "detailDiscardQty"))} EA / ${cf.format(sum(detailRows, "detailDiscardAmount"))}`;
  $("vendingDiscardBody").innerHTML = detailRows.slice(0, 350).map((row) => `
    <tr>
      <td>${typePill(row)}</td>
      <td>${escapeHtml(row.branch)}</td>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${escapeHtml(qtyText(row, row.stock))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.detailDiscardQty))}</td>
      <td class="number">${cf.format(row.price)}</td>
      <td class="number">${cf.format(row.detailDiscardAmount)}</td>
    </tr>
  `).join("") || emptyRow(10);
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
  const vendingMode = $("datasetFilter").value === "vending";
  const aggregateByCode = vendingMode && $("branchFilter").value === "전체";
  const rows = purchaseRows(vendingMode, aggregateByCode).sort(vendingMode
    ? (a, b) => a.code.localeCompare(b.code, "ko") || a.branch.localeCompare(b.branch, "ko")
    : (a, b) => b.purchaseAmount - a.purchaseAmount || b.purchaseQty - a.purchaseQty);
  $("purchaseBranchHeader").textContent = aggregateByCode ? "필요 지점" : "지점";
  $("purchaseBody").innerHTML = rows.slice(0, 350).map((row) => `
    <tr>
      <td>${typePill(row)}</td>
      <td>${escapeHtml(row.branch)}</td>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${escapeHtml(qtyText(row, row.stock))}</td>
      ${vendingMode ? "" : `<td class="number">${escapeHtml(qtyText(row, row.monthlyConsume))}</td>`}
      <td class="number">${escapeHtml(qtyText(row, row.planningMonthlyConsume, df))}</td>
      ${vendingMode ? `<td class="number">${escapeHtml(qtyText(row, row.avgMonthlyDiscard, df))}</td>` : ""}
      <td>${escapeHtml(row.planningBasis)}</td>
      <td class="number">${escapeHtml(qtyText(row, row.recommendedStock))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.purchaseQty))}</td>
      <td class="number">${cf.format(row.price)}</td>
      <td class="number">${cf.format(row.purchaseAmount)}</td>
    </tr>
  `).join("") || emptyRow(14);
}

function purchaseRows(vendingMode, aggregateByCode) {
  const rows = state.filteredRows.filter((row) => row.purchaseQty > 0);
  if (!vendingMode || !aggregateByCode) return rows;
  const map = new Map();
  rows.forEach((row) => {
    const target = map.get(row.code) || {
      ...row,
      branch: "",
      branchCount: 0,
      stock: 0,
      planningMonthlyConsume: 0,
      avgMonthlyDiscard: 0,
      recommendedStock: 0,
      purchaseQty: 0,
      purchaseAmount: 0,
    };
    target.branchCount += 1;
    target.stock += row.stock;
    target.planningMonthlyConsume += row.planningMonthlyConsume;
    target.avgMonthlyDiscard += row.avgMonthlyDiscard || 0;
    target.recommendedStock += row.recommendedStock;
    target.purchaseQty += row.purchaseQty;
    target.purchaseAmount += row.purchaseAmount;
    map.set(row.code, target);
  });
  return Array.from(map.values()).map((row) => ({
    ...row,
    branch: `${nf.format(row.branchCount)}개 지점`,
  }));
}

function renderDelivery() {
  const rows = state.filteredRows.filter((row) => row.deliveryQty > 0).sort((a, b) => a.code.localeCompare(b.code, "ko") || a.branch.localeCompare(b.branch, "ko"));
  $("deliveryBody").innerHTML = rows.map((row) => `
    <tr>
      <td>${typePill(row)}</td>
      <td>${escapeHtml(row.branch)}</td>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${escapeHtml(qtyText(row, row.stock))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.monthlyConsume))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.deliveryTargetStock))}</td>
      <td class="number">${escapeHtml(qtyText(row, row.deliveryQty))}</td>
      <td class="number">${cf.format(row.price)}</td>
      <td class="number">${cf.format(row.deliveryAmount)}</td>
    </tr>
  `).join("") || emptyRow(12);
}

function renderVendingOtherHistory(vendingMode, branch, query, analysisPeriods) {
  const panel = $("vendingOtherPanel");
  panel.classList.toggle("hidden", !vendingMode);
  if (!vendingMode) return;
  const selectedPeriods = analysisPeriods;
  const selectedSet = new Set(selectedPeriods);
  if (!selectedSet.has(state.otherSelection.period)) state.otherSelection.period = null;
  const rows = (state.history.vendingOtherRows || [])
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
    .map((row) => {
      const detailQty = state.otherSelection.period ? (row.otherByPeriod?.[state.otherSelection.period] || 0) : row.selectedTotal;
      return { ...row, detailOtherQty: detailQty, detailOtherAmount: detailQty * (row.price || 0) };
    })
    .filter((row) => row.detailOtherQty > 0)
    .filter((row) => !state.otherSelection.branch || row.branch === state.otherSelection.branch)
    .sort((a, b) => b.detailOtherAmount - a.detailOtherAmount || b.detailOtherQty - a.detailOtherQty || a.branch.localeCompare(b.branch, "ko") || a.code.localeCompare(b.code, "ko"));
  const selectedTitle = [
    state.otherSelection.branch,
    state.otherSelection.period ? formatPeriodLabel(state.otherSelection.period) : null,
  ].filter(Boolean).join(" / ");
  $("vendingOtherTitle").textContent = selectedTitle ? `기타 세부내역 (${selectedTitle})` : "기타 세부내역";
  $("vendingOtherBasis").textContent = selectedTitle
    ? `${selectedTitle} 조건에 해당하는 기타 입력을 구매단가 기준 금액으로 표시합니다. 품목 행을 클릭하면 메모를 남길 수 있습니다.`
    : `${describeSelectedHistory("자판기 상품", selectedPeriods)} 원본 시트의 기타 입력을 구매단가 기준 금액으로 표시합니다. 월 점을 클릭하면 지점별 비율이 원그래프로 표시됩니다.`;
  $("vendingOtherTotalHeader").textContent = state.otherSelection.period ? "기타 수량" : `${selectedPeriods.length}개월 기타 합계`;
  $("vendingOtherCount").textContent = `총 ${nf.format(sum(detailRows, "detailOtherQty"))} EA / ${cf.format(sum(detailRows, "detailOtherAmount"))} / ${nf.format(detailRows.length)}품목`;
  $("vendingOtherBody").innerHTML = detailRows.map((row) => {
    const noteKey = otherNoteKey(row);
    const note = state.otherNotes[noteKey] || "";
    const label = `${row.branch} / ${row.code} / ${row.category}`;
    return `
    <tr class="${state.activeOtherNote?.key === noteKey ? "selected-row" : ""}">
      <td>${escapeHtml(row.branch)}</td>
      <td>${escapeHtml(row.code)}</td>
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.spec)}</td>
      <td>${escapeHtml(row.unit || "EA")}</td>
      <td class="number">${nf.format(row.selectedMonthCount)}</td>
      <td class="number">${escapeHtml(qtyText(row, row.detailOtherQty))}</td>
      <td class="number">${cf.format(row.price || 0)}</td>
      <td class="number">${cf.format(row.detailOtherAmount)}</td>
      <td class="number">${escapeHtml(qtyText(row, row.selectedAverage, df))}</td>
      <td class="memo-text">${escapeHtml(note || "메모 없음")}</td>
      <td><button class="memo-button" data-other-note-key="${escapeHtml(noteKey)}" data-other-note-label="${escapeHtml(label)}" type="button">${note ? "수정" : "작성"}</button></td>
    </tr>
  `;
  }).join("") || emptyRow(12);
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
  $("purchaseConsumeHeader").classList.toggle("hidden", vendingMode);
  $("purchaseConsumeHeader").textContent = "당월 소모";
  $("purchasePlanningHeader").textContent = vendingMode ? `${analysisMonths}개월 평균 판매` : consumableMode ? `${analysisMonths}개월 평균 월소모` : "추천 기준 월평균";
  $("purchaseDiscardHeader").textContent = `${analysisMonths}개월 평균 폐기`;
  $("purchaseDiscardHeader").classList.toggle("hidden", !vendingMode);
  $("purchaseTargetHeader").textContent = vendingMode ? "목표 보유" : consumableMode ? `${safetyMonths}개월 예상 소모량` : "권장 보유";
  $("purchaseQtyHeader").textContent = vendingMode || consumableMode ? "구매필요" : "구매 추천";
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
  renderCategoryAmounts();
  const discardRows = [...state.filteredRows].filter((row) => row.discardAmount > 0).sort((a, b) => b.discardAmount - a.discardAmount).slice(0, 10);
  const emptyText = vendingMode ? "폐기 금액 데이터가 없습니다." : "이동 금액 데이터가 없습니다.";
  $("discardAmountList").innerHTML = discardRows.map((row, index) => rankRow(index, `${row.category} / ${row.branch}`, row.spec, cf.format(row.discardAmount))).join("") || `<div class="empty">${emptyText}</div>`;
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
      <td class="number">${escapeHtml(qtyText(row, row.stock))}</td>
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

function exportPurchaseCsv() {
  const vendingMode = $("datasetFilter").value === "vending";
  const analysisMonths = selectedAnalysisPeriods($("datasetFilter").value).length;
  const safetyMonths = Math.max(1, toNumber($("safetyMonths").value) || 2);
  const aggregateByCode = vendingMode && $("branchFilter").value === "전체";
  const rows = purchaseRows(vendingMode, aggregateByCode);
  const header = vendingMode
    ? ["구분", aggregateByCode ? "필요지점" : "지점", "품번", "품명", "규격", "단위", "현재고", `${analysisMonths}개월평균판매`, `${analysisMonths}개월평균폐기`, "계산기준", "목표보유", "구매필요", "단가", "예상금액"]
    : $("datasetFilter").value === "consumable"
      ? ["구분", "지점", "품번", "품명", "규격", "단위", "현재고", "당월소모", `${analysisMonths}개월평균월소모`, "계산기준", `${safetyMonths}개월예상소모량`, "구매필요", "단가", "예상금액"]
      : ["구분", "지점", "품번", "품명", "규격", "단위", "현재고", "당월소모", "추천기준월평균", "계산기준", "권장보유", "구매추천", "단가", "예상금액"];
  const csv = [header, ...rows.map((row) => [
    row.sourceLabel,
    row.branch,
    row.code,
    row.category,
    row.spec,
    row.unit || "EA",
    row.stock,
    ...(vendingMode ? [] : [row.monthlyConsume]),
    row.planningMonthlyConsume,
    ...(vendingMode ? [row.avgMonthlyDiscard] : []),
    row.planningBasis,
    row.recommendedStock,
    row.purchaseQty,
    row.price,
    row.purchaseAmount,
  ])].map((line) => line.map(csvCell).join(",")).join("\n");
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = vendingMode ? "자판기_구매필요.csv" : $("datasetFilter").value === "consumable" ? "경상소모품_구매필요.csv" : "구매추천.csv";
  link.click();
  URL.revokeObjectURL(link.href);
}

function exportDeliveryCsv() {
  const rows = state.filteredRows.filter((row) => row.deliveryQty > 0);
  const header = ["구분", "지점", "품번", "품명", "규격", "단위", "현재고", "당월판매", "배송주기(주)", "주당판매환산", "배송주기목표재고", "배송필요", "단가", "예상금액"];
  const csv = [header, ...rows.map((row) => [
    row.sourceLabel,
    row.branch,
    row.code,
    row.category,
    row.spec,
    row.unit || "EA",
    row.stock,
    row.monthlyConsume,
    row.deliveryWeeks,
    row.deliveryWeeklySales,
    row.deliveryTargetStock,
    row.deliveryQty,
    row.price,
    row.deliveryAmount,
  ])].map((line) => line.map(csvCell).join(",")).join("\n");
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "자판기_배송계획.csv";
  link.click();
  URL.revokeObjectURL(link.href);
}

function csvCell(value) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
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
