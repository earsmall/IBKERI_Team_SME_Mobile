const NAVER_URLS = {
  kospi: "https://r.jina.ai/http://https://finance.naver.com/sise/sise_index.naver?code=KOSPI",
  kosdaq: "https://r.jina.ai/http://https://finance.naver.com/sise/sise_index.naver?code=KOSDAQ",
  market: "https://r.jina.ai/http://https://finance.naver.com/marketindex/",
};

const GOLD_API_BASE = "https://api.gold-api.com/price";
const SNAPSHOT_KEY = "market_board_alt_snapshot_v1";
const SHEET_DOCUMENT_ID = "1R2BB-k4L6a6QwiD5nmFz-idJoIUGnhQCDzHF0ijwH4c";
const LOAN_SHEET_DOCUMENT_ID = "1fNiuZjbvbH7hjomQqXjAAxt6GE_b_X-_zuQlzE5p8YY";
const EXPORT_SHEET_DOCUMENT_ID = "1poM2hIz5bfig7oTFekq97yAX5_aVYTbSSHXg8fW2TAk";
const SHEET_BASE_URL =
  `https://docs.google.com/spreadsheets/d/${SHEET_DOCUMENT_ID}/gviz/tq`;
const OPENSHEET_BASE_URL = `https://opensheet.elk.sh/${SHEET_DOCUMENT_ID}`;
const LOAN_SHEET_BASE_URL =
  `https://docs.google.com/spreadsheets/d/${LOAN_SHEET_DOCUMENT_ID}/gviz/tq`;
const LOAN_OPENSHEET_BASE_URL = `https://opensheet.elk.sh/${LOAN_SHEET_DOCUMENT_ID}`;
const EXPORT_SHEET_BASE_URL =
  `https://docs.google.com/spreadsheets/d/${EXPORT_SHEET_DOCUMENT_ID}/gviz/tq`;
const EXPORT_OPENSHEET_BASE_URL = `https://opensheet.elk.sh/${EXPORT_SHEET_DOCUMENT_ID}`;

const marketItems = [
  { id: "govt3y", name: "국고채 3년", code: "GOVT03Y", badge: "Rate", unit: "%" },
  { id: "kospi", name: "KOSPI지수", code: "KOSPI", badge: "Index", unit: "" },
  { id: "kosdaq", name: "KOSDAQ", code: "KOSDAQ", badge: "Index", unit: "" },
  { id: "usdkrw", name: "달러/원", code: "USDKRW", badge: "FX", unit: "" },
  { id: "wti", name: "WTI유", code: "OIL", badge: "Commodity", unit: "" },
  { id: "gold", name: "금", code: "XAU", badge: "Metal", unit: "" },
  { id: "silver", name: "은", code: "XAG", badge: "Metal", unit: "" },
  { id: "bitcoin", name: "비트코인", code: "BTC", badge: "Crypto", unit: "" },
];

let smeData = [];
let smeYears = [];
let startupSeries = [];
let loanSeries = [];
let delinquencySeries = [];
let loanYears = [];
let investmentSeries = [];
let investmentStageSeries = [];
let investmentSectorSeries = [];
let investmentSourceSeries = [];
let investmentDates = [];
let exportSeries = [];
let exportCountrySeries = [];
let exportDates = [];
let smeLoadError = "";
let startupLoadError = "";
let loanLoadError = "";
let investmentLoadError = "";
let exportLoadError = "";

function formatClock(date = new Date()) {
  return new Intl.DateTimeFormat("ko-KR", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  }).format(date);
}

function formatNumber(value, digits = 2) {
  return new Intl.NumberFormat("en-US", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  }).format(value);
}

function formatSmeValue(item, value) {
  if (item.title === "기업수") {
    return `${formatNumber(value / 10000, 1)}<span class="sme-value-unit">만개</span>`;
  }

  if (item.title === "종사자수") {
    return `${formatNumber(value / 10000, 1)}<span class="sme-value-unit">만명</span>`;
  }

  if (item.title === "매출액") {
    return `${formatNumber(value / 1000000, 1)}<span class="sme-value-unit">조원</span>`;
  }

  return `${formatNumber(value, 0)}`;
}

function formatManCount(value) {
  return `${formatNumber(value / 10000, 1)}<span class="sme-value-unit">만개</span>`;
}

function formatTotalFootnote(item, totalValue) {
  if (item.title === "기업수") {
    return `전체 ${formatNumber(totalValue / 10000, 1)}만개`;
  }

  if (item.title === "종사자수") {
    return `전체 ${formatNumber(totalValue / 10000, 1)}만명`;
  }

  if (item.title === "매출액") {
    return `전체 ${formatNumber(totalValue / 1000000, 1)}조원`;
  }

  return `전체 ${formatNumber(totalValue, 0)}`;
}

function formatStartupMetric(value, type) {
  if (value === undefined || value === null || Number.isNaN(value)) {
    return "-";
  }

  if (type === "share") {
    return `${formatNumber(value, 1)}%`;
  }

  return `${formatNumber(value / 10000, 1)}<span class="sme-value-unit">만개</span>`;
}

function formatYoYGrowth(current, previous) {
  if (
    current === undefined ||
    current === null ||
    previous === undefined ||
    previous === null ||
    previous === 0
  ) {
    return "";
  }

  const growth = ((current - previous) / previous) * 100;
  if (growth > 0) {
    return `<span class="startup-delta is-up">(전년대비 ▲ ${formatNumber(Math.abs(growth), 1)}%)</span>`;
  }
  if (growth < 0) {
    return `<span class="startup-delta is-down">(전년대비 ▼ ${formatNumber(Math.abs(growth), 1)}%)</span>`;
  }
  return `<span class="startup-delta">(전년대비 0.0%)</span>`;
}

function formatSmeDelta(item, current, previous) {
  if (
    current === undefined ||
    current === null ||
    previous === undefined ||
    previous === null
  ) {
    return "";
  }

  const delta = current - previous;
  const directionClass = delta > 0 ? " is-up" : delta < 0 ? " is-down" : "";

  if (item.title === "기업수") {
    return `<span class="startup-delta${directionClass}">(전년대비 ${delta > 0 ? "▲ " : delta < 0 ? "▼ " : ""}${formatNumber(Math.abs(delta) / 10000, 1)}만개)</span>`;
  }

  if (item.title === "종사자수") {
    return `<span class="startup-delta${directionClass}">(전년대비 ${delta > 0 ? "▲ " : delta < 0 ? "▼ " : ""}${formatNumber(Math.abs(delta) / 10000, 1)}만명)</span>`;
  }

  if (item.title === "매출액") {
    return `<span class="startup-delta${directionClass}">(전년대비 ${delta > 0 ? "▲ " : delta < 0 ? "▼ " : ""}${formatNumber(Math.abs(delta) / 1000000, 1)}조원)</span>`;
  }

  return `<span class="startup-delta${directionClass}">(전년대비 ${delta > 0 ? "▲ " : delta < 0 ? "▼ " : ""}${formatNumber(Math.abs(delta), 0)})</span>`;
}

function setStatus(message) {
  const status = document.getElementById("last-updated");
  if (status) {
    status.textContent = `업데이트 ${formatClock()}`;
  }
}

function fetchText(url) {
  return fetch(url, { cache: "no-store" }).then((response) => {
    if (!response.ok) {
      throw new Error(`요청 실패: ${response.status}`);
    }

    return response.text();
  });
}

function fetchJson(url) {
  return fetch(url, { cache: "no-store" }).then((response) => {
    if (!response.ok) {
      throw new Error(`요청 실패: ${response.status}`);
    }

    return response.json();
  });
}

function mapGvizPayload(payload) {
  const cols = payload.table?.cols ?? [];
  const rows = payload.table?.rows ?? [];

  return rows.map((row) => {
    const cells = row.c ?? [];
    return cols.reduce((accumulator, col, index) => {
      const key = col.label || col.id || `col_${index}`;
      const cell = cells[index];
      accumulator[key] = cell?.v ?? null;
      return accumulator;
    }, {});
  });
}

function parseNumeric(value) {
  if (value === null || value === undefined || value === "") {
    return Number.NaN;
  }

  if (typeof value === "number") {
    return value;
  }

  return Number(String(value).replace(/,/g, "").trim());
}

function loadGoogleSheetViaJsonp(sheetName, baseUrl = SHEET_BASE_URL) {
  return new Promise((resolve, reject) => {
    const callbackName = `__sheetCallback_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
    const script = document.createElement("script");
    const tqx = encodeURIComponent(`out:json;responseHandler:${callbackName}`);
    const sheet = encodeURIComponent(sheetName);
    const timeoutId = window.setTimeout(() => {
      cleanup();
      reject(new Error(`${sheetName} 시트 응답 시간 초과`));
    }, 7000);

    window[callbackName] = (payload) => {
      cleanup();
      try {
        resolve(mapGvizPayload(payload));
      } catch (error) {
        reject(error);
      }
    };

    script.src = sheetName
      ? `${baseUrl}?tqx=${tqx}&sheet=${sheet}`
      : `${baseUrl}?tqx=${tqx}`;
    script.async = true;

    script.onerror = () => {
      cleanup();
      reject(new Error(`${sheetName} 시트 로드 실패`));
    };

    function cleanup() {
      window.clearTimeout(timeoutId);
      delete window[callbackName];
      script.remove();
    }

    document.head.appendChild(script);
  });
}

function loadGoogleSheetViaOpenSheet(sheetName, baseUrl = OPENSHEET_BASE_URL) {
  const sheet = encodeURIComponent(sheetName);
  return fetchJson(sheetName ? `${baseUrl}/${sheet}?raw=true` : `${baseUrl}?raw=true`);
}

async function loadGoogleSheet(sheetName, options = {}) {
  const {
    baseUrl = SHEET_BASE_URL,
    openSheetBaseUrl = OPENSHEET_BASE_URL,
  } = options;
  try {
    return await loadGoogleSheetViaJsonp(sheetName, baseUrl);
  } catch (jsonpError) {
    try {
      return await loadGoogleSheetViaOpenSheet(sheetName, openSheetBaseUrl);
    } catch (openSheetError) {
      throw new Error(
        `${sheetName} 시트 로드 실패 (${jsonpError.message}; 대체 경로 실패: ${openSheetError.message})`,
      );
    }
  }
}

function parseSmeRows(rows) {
  const labelConfig = {
    기업수: { unit: "개", color: "#2c7be5" },
    종사자수: { unit: "명", color: "#4a9bff" },
    매출액: { unit: "백만원", color: "#7fb8ff" },
  };

  const grouped = {};

  rows.forEach((row) => {
    const item = String(row["항목"] || "").trim();
    const year = String(row["시점"] || "").trim();
    const total = parseNumeric(row["전체기업"]);
    const sme = parseNumeric(row["중소기업"]);

    if (!labelConfig[item] || !year || Number.isNaN(total) || Number.isNaN(sme)) {
      return;
    }

    if (!grouped[item]) {
      grouped[item] = {
        title: item,
        unit: labelConfig[item].unit,
        color: labelConfig[item].color,
        years: {},
      };
    }

    grouped[item].years[year] = { total, sme };
  });

  const orderedItems = ["기업수", "종사자수", "매출액"];
  const nextData = orderedItems
    .map((name) => grouped[name])
    .filter(Boolean);

  const nextYears = Object.keys(nextData[0]?.years || {}).sort();

  if (!nextData.length || !nextYears.length) {
    throw new Error("중소기업 데이터를 파싱하지 못했습니다.");
  }

  return { nextData, nextYears };
}

function normalizeLabel(value) {
  return String(value || "").replace(/\s+/g, "").trim();
}

function parseStartupRows(rows) {
  const byYear = {};

  rows.forEach((row) => {
    const year = String(row["시점"] || "").trim();
    const division = normalizeLabel(row["구분"] || "");
    const startupCount = parseNumeric(row["전체 창업기업 수"]);
    const techStartupCount = parseNumeric(row["기술기반업종 창업기업 수"]);

    if (!year || (Number.isNaN(startupCount) && Number.isNaN(techStartupCount))) {
      return;
    }

    if (!byYear[year]) {
      byYear[year] = {};
    }

    if (!division || division.includes("전체")) {
      byYear[year].startupCount = startupCount;
      byYear[year].techStartupCount = techStartupCount;
    }
  });

  const series = Object.entries(byYear)
    .map(([year, values]) => {
      const startupCount = values.startupCount;
      const techStartupCount = values.techStartupCount;
      const techShare =
        values.techShare !== undefined
          ? values.techShare
          : startupCount && techStartupCount
            ? (techStartupCount / startupCount) * 100
            : undefined;

      return {
        year,
        startupCount,
        techStartupCount,
        techShare,
      };
    })
    .filter((row) => row.startupCount !== undefined || row.techStartupCount !== undefined || row.techShare !== undefined)
    .sort((a, b) => Number(a.year) - Number(b.year));

  return series;
}

function parseSheetDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  const stringValue = String(value || "").trim();
  const dateMatch = stringValue.match(/^Date\((\d+),(\d+),(\d+)\)$/);
  if (dateMatch) {
    return new Date(Number(dateMatch[1]), Number(dateMatch[2]), Number(dateMatch[3]));
  }

  const parsedDate = new Date(stringValue);
  if (!Number.isNaN(parsedDate.getTime())) {
    return parsedDate;
  }

  return null;
}

function formatYearMonth(date) {
  return new Intl.DateTimeFormat("ko-KR", {
    year: "numeric",
    month: "long",
  }).format(date);
}

function formatLoanValue(value) {
  if (value === undefined || value === null || Number.isNaN(value)) {
    return "-";
  }

  return `${formatNumber(value, 1)}<span class="sme-value-unit">조원</span>`;
}

function formatRateValue(value) {
  if (value === undefined || value === null || Number.isNaN(value)) {
    return "-";
  }

  return `${formatNumber(value, 2)}<span class="sme-value-unit">%</span>`;
}

function formatRateDelta(current, previous) {
  if (
    current === undefined ||
    current === null ||
    previous === undefined ||
    previous === null
  ) {
    return "";
  }

  const delta = current - previous;
  const directionClass = delta > 0 ? " is-up" : delta < 0 ? " is-down" : "";
  return `<span class="startup-delta${directionClass}">(전년동월대비 ${delta > 0 ? "▲ " : delta < 0 ? "▼ " : ""}${formatNumber(Math.abs(delta), 2)}%p)</span>`;
}

function formatEokValue(value) {
  if (value === undefined || value === null || Number.isNaN(value)) {
    return "-";
  }

  return `${formatNumber(value, 0)}<span class="sme-value-unit">억원</span>`;
}

function formatCountValue(value) {
  if (value === undefined || value === null || Number.isNaN(value)) {
    return "-";
  }

  return `${formatNumber(value, 0)}<span class="sme-value-unit">개</span>`;
}

function formatDateKey(date) {
  if (!date) {
    return "";
  }

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatInvestmentPeriod(date) {
  if (!date) {
    return "-";
  }

  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const suffixMap = {
    3: "1분기",
    6: "2분기",
    9: "3분기",
    12: "연간",
  };

  return `${year}년 ${suffixMap[month] || `${month}월`}`;
}

function formatInvestmentDelta(current, previous, type) {
  if (
    current === undefined ||
    current === null ||
    previous === undefined ||
    previous === null ||
    previous === 0
  ) {
    return "";
  }

  const growth = ((current - previous) / previous) * 100;
  const delta = current - previous;
  const directionClass = growth > 0 ? " is-up" : growth < 0 ? " is-down" : "";
  const directionMark = growth > 0 ? "▲ " : growth < 0 ? "▼ " : "";
  const deltaText =
    type === "count"
      ? `${delta > 0 ? "▲ " : delta < 0 ? "▼ " : ""}${formatNumber(Math.abs(delta), 0)}개`
      : `${delta > 0 ? "▲ " : delta < 0 ? "▼ " : ""}${formatNumber(Math.abs(delta), 0)}억원`;

  return `<span class="startup-delta${directionClass}">(전년대비 ${directionMark}${formatNumber(Math.abs(growth), 1)}%, ${deltaText})</span>`;
}

function formatInvestmentShare(value, total) {
  if (
    value === undefined ||
    value === null ||
    Number.isNaN(value) ||
    total === undefined ||
    total === null ||
    Number.isNaN(total) ||
    total === 0
  ) {
    return "";
  }

  return `<div class="startup-subvalue">(비중: <span class="startup-share-value">${formatNumber((value / total) * 100, 1)}%</span>)</div>`;
}

function formatLoanYoYGrowth(current, previous) {
  if (
    current === undefined ||
    current === null ||
    previous === undefined ||
    previous === null ||
    previous === 0
  ) {
    return "";
  }

  const growth = ((current - previous) / previous) * 100;
  if (growth > 0) {
    return `<span class="startup-delta is-up">(전년동월대비 ▲ ${formatNumber(Math.abs(growth), 1)}%)</span>`;
  }
  if (growth < 0) {
    return `<span class="startup-delta is-down">(전년동월대비 ▼ ${formatNumber(Math.abs(growth), 1)}%)</span>`;
  }
  return `<span class="startup-delta">(전년동월대비 0.0%)</span>`;
}

function formatLoanYoYDelta(current, previous) {
  if (
    current === undefined ||
    current === null ||
    previous === undefined ||
    previous === null
  ) {
    return "";
  }

  const delta = current - previous;
  const directionClass = delta > 0 ? " is-up" : delta < 0 ? " is-down" : "";
  return `<span class="startup-delta${directionClass}">(증감 ${delta > 0 ? "▲ " : delta < 0 ? "▼ " : ""}${formatNumber(Math.abs(delta), 1)}조원)</span>`;
}

function formatLoanYoYGrowthWithDelta(current, previous) {
  if (
    current === undefined ||
    current === null ||
    previous === undefined ||
    previous === null ||
    previous === 0
  ) {
    return "";
  }

  const growth = ((current - previous) / previous) * 100;
  const delta = current - previous;
  const directionClass = growth > 0 ? " is-up" : growth < 0 ? " is-down" : "";
  const directionMark = growth > 0 ? "▲ " : growth < 0 ? "▼ " : "";

  return `<span class="startup-delta${directionClass}">(전년동월대비 ${directionMark}${formatNumber(Math.abs(growth), 1)}%, ${delta > 0 ? "▲ " : delta < 0 ? "▼ " : ""}${formatNumber(Math.abs(delta), 1)}조원)</span>`;
}

function parseLoanRows(rows) {
  return rows
    .map((row) => {
      const pointDate = parseSheetDate(row["시점"] ?? row[""] ?? row.A);
      const balance = parseNumeric(row["은행권 중기대출잔액"]);
      const netIncrease = parseNumeric(row["은행권 중기대출 순증"]);

      if (!pointDate || (Number.isNaN(balance) && Number.isNaN(netIncrease))) {
        return null;
      }

      return {
        date: pointDate,
        year: pointDate.getFullYear(),
        month: pointDate.getMonth() + 1,
        balance,
        netIncrease,
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.date - b.date);
}

function parseDelinquencyRows(rows) {
  return rows
    .map((row) => {
      const pointDate = parseSheetDate(row["시점"] ?? row.A);
      const largeCompanyRate = parseNumeric(row["대기업 연체율"]);
      const smeRate = parseNumeric(row["중소기업 연체율"]);
      const corporateRate = parseNumeric(row["중소법인 연체율"]);
      const solePropRate = parseNumeric(row["개인사업자 연체율"]);

      if (
        !pointDate ||
        [largeCompanyRate, smeRate, corporateRate, solePropRate].every((value) => Number.isNaN(value))
      ) {
        return null;
      }

      return {
        date: pointDate,
        year: pointDate.getFullYear(),
        month: pointDate.getMonth() + 1,
        largeCompanyRate,
        smeRate,
        corporateRate,
        solePropRate,
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.date - b.date);
}

function parseInvestmentRows(rows, valueKeys) {
  return rows
    .map((row) => {
      const pointDate = parseSheetDate(row["시점"] ?? row.A);
      if (!pointDate) {
        return null;
      }

      const parsed = {
        date: pointDate,
        key: formatDateKey(pointDate),
        year: pointDate.getFullYear(),
        month: pointDate.getMonth() + 1,
      };

      valueKeys.forEach((key) => {
        parsed[key] = parseNumeric(row[key]);
      });

      return parsed;
    })
    .filter(Boolean)
    .sort((a, b) => a.date - b.date);
}

function getLoanSelectedYear() {
  const select = document.getElementById("loan-year-select");
  return Number(select?.value || loanYears[loanYears.length - 1] || 0);
}

function getLoanLatestPointForYear(targetYear) {
  const yearRows = loanSeries.filter((item) => item.year === targetYear);
  if (!yearRows.length) {
    return null;
  }

  return yearRows[yearRows.length - 1];
}

function getDelinquencyLatestPointForYear(targetYear) {
  const yearRows = delinquencySeries.filter((item) => item.year === targetYear);
  if (!yearRows.length) {
    return null;
  }

  return yearRows[yearRows.length - 1];
}

function getLoanBalancePointForYear(targetYear) {
  const yearRows = loanSeries.filter((item) => item.year === targetYear);
  if (!yearRows.length) {
    return null;
  }

  const decemberPoint = yearRows.find((item) => item.month === 12);
  return decemberPoint || yearRows[yearRows.length - 1];
}

function getRecentLoanBalancePoints() {
  const selectedYear = getLoanSelectedYear();
  const years = loanYears.filter((year) => year <= selectedYear).slice(-3);
  return years
    .map((year) => getLoanBalancePointForYear(year))
    .filter(Boolean);
}

function getRecentLoanNetIncreasePoints() {
  const selectedYear = getLoanSelectedYear();
  const decemberPoints = loanSeries
    .filter((item) => item.year <= selectedYear && item.month === 12)
    .sort((a, b) => a.year - b.year);

  return decemberPoints.slice(-3);
}

function getRecentLoanNetIncreaseLatestMonthPoints() {
  const selectedYear = getLoanSelectedYear();
  const selectedYearLatestPoint = getLoanLatestPointForYear(selectedYear);

  if (!selectedYearLatestPoint || selectedYearLatestPoint.month === 12) {
    return [];
  }

  return loanSeries
    .filter((item) => item.year <= selectedYear && item.month === selectedYearLatestPoint.month)
    .sort((a, b) => a.year - b.year)
    .slice(-3);
}

function getPreviousYearSameMonthLoanPoint(referenceDate) {
  if (!referenceDate) {
    return null;
  }

  return (
    loanSeries.find(
      (item) =>
        item.year === referenceDate.getFullYear() - 1 &&
        item.month === referenceDate.getMonth() + 1,
    ) || null
  );
}

function getPreviousYearSameMonthDelinquencyPoint(referenceDate) {
  if (!referenceDate) {
    return null;
  }

  return (
    delinquencySeries.find(
      (item) =>
        item.year === referenceDate.getFullYear() - 1 &&
        item.month === referenceDate.getMonth() + 1,
    ) || null
  );
}

function getFilteredDelinquencySeries(windowSize = 5) {
  const selectedYear = getLoanSelectedYear();
  const latestPoint = getDelinquencyLatestPointForYear(selectedYear);
  if (!latestPoint) {
    return [];
  }

  const years = loanYears.filter((year) => year <= selectedYear).slice(-windowSize);

  return years
    .map((year) => {
      const sameMonthPoint = delinquencySeries.find(
        (item) => item.year === year && item.month === latestPoint.month,
      );
      return sameMonthPoint || getDelinquencyLatestPointForYear(year);
    })
    .filter(Boolean);
}

function getInvestmentSelectedKey() {
  const select = document.getElementById("investment-date-select");
  return select?.value || investmentDates[investmentDates.length - 1] || "";
}

function getInvestmentRecord(series, key = getInvestmentSelectedKey()) {
  return series.find((item) => item.key === key) || null;
}

function getPreviousYearInvestmentRecord(series, key = getInvestmentSelectedKey()) {
  const current = getInvestmentRecord(series, key);
  if (!current) {
    return null;
  }

  return (
    series.find(
      (item) =>
        item.year === current.year - 1 &&
        item.month === current.month,
    ) || null
  );
}

function formatInvestmentPointLabel(key = getInvestmentSelectedKey()) {
  const current = getInvestmentRecord(investmentSeries, key);
  return current?.date ? formatInvestmentPeriod(current.date) : "-";
}

function getRecentInvestmentSeries(series, windowSize = 3) {
  const selectedKey = getInvestmentSelectedKey();
  const current = getInvestmentRecord(series, selectedKey);
  if (!current) {
    return [];
  }

  return series
    .filter((item) => item.month === current.month && item.year <= current.year)
    .slice(-windowSize);
}

function parseExportRows(rows, valueKeys) {
  return rows
    .map((row) => {
      const pointDate = parseExportPoint(row["시점"] ?? row.A);
      if (!pointDate) {
        return null;
      }

      const parsed = {
        date: pointDate,
        key: formatDateKey(pointDate),
        year: pointDate.getFullYear(),
        month: pointDate.getMonth() + 1,
      };

      valueKeys.forEach((key) => {
        parsed[key] = parseNumeric(row[key]);
      });

      return parsed;
    })
    .filter(Boolean)
    .sort((a, b) => a.date - b.date);
}

function parseExportCountryRows(rows, valueKeys) {
  const series = parseExportRows(rows, valueKeys);
  const byKey = new Map();

  series.forEach((item) => {
    const currentTotal = valueKeys.reduce((sum, key) => sum + (item[key] || 0), 0);
    const existing = byKey.get(item.key);
    if (!existing) {
      byKey.set(item.key, item);
      return;
    }

    const existingTotal = valueKeys.reduce((sum, key) => sum + (existing[key] || 0), 0);
    if (currentTotal < existingTotal) {
      byKey.set(item.key, item);
    }
  });

  return Array.from(byKey.values()).sort((a, b) => a.date - b.date);
}

function parseExportPoint(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const year = Math.trunc(value);
    if (year >= 1900 && year <= 3000) {
      return new Date(year, 0, 1);
    }
  }

  const stringValue = String(value || "").trim().replace(/\s*p\)\s*$/i, "");
  const yearMatch = stringValue.match(/^(\d{4})$/);
  if (yearMatch) {
    return new Date(Number(yearMatch[1]), 0, 1);
  }

  const quarterMatch = stringValue.match(/^(\d{4})\s*[.\-\/]\s*([1-4])\s*\/\s*4$/);
  if (quarterMatch) {
    const year = Number(quarterMatch[1]);
    const quarter = Number(quarterMatch[2]);
    const monthMap = { 1: 3, 2: 6, 3: 9, 4: 12 };
    return new Date(year, monthMap[quarter] - 1, 1);
  }

  return parseSheetDate(value);
}

function formatExportPeriod(date) {
  if (!date) {
    return "-";
  }

  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  if (month === 1 && day === 1) {
    return `${year}`;
  }
  const suffixMap = {
    3: "1분기",
    6: "2분기",
    9: "3분기",
    12: "4분기",
  };

  return `${year}년 ${suffixMap[month] || `${month}월`}`;
}

function getExportSelectedKey() {
  const select = document.getElementById("export-date-select");
  return select?.value || exportDates[exportDates.length - 1] || "";
}

function getExportRecord(key = getExportSelectedKey()) {
  return exportSeries.find((item) => item.key === key) || null;
}

function getExportCountryRecord(key = getExportSelectedKey()) {
  return exportCountrySeries.find((item) => item.key === key) || null;
}

function getPreviousYearExportRecord(series, key = getExportSelectedKey()) {
  const current = series.find((item) => item.key === key) || null;
  if (!current) {
    return null;
  }

  return (
    series.find(
      (item) =>
        item.year === current.year - 1 &&
        item.month === current.month,
    ) || null
  );
}

function formatExportCompanyCount(value) {
  if (value === undefined || value === null || Number.isNaN(value)) {
    return "-";
  }

  return `${formatNumber(value, 0)}<span class="sme-value-unit">개사</span>`;
}

function formatExportAmount(value) {
  if (value === undefined || value === null || Number.isNaN(value)) {
    return "-";
  }

  return `${formatNumber(value / 100000, 1)}<span class="sme-value-unit">억달러</span>`;
}

function getRecentExportSeries(series, windowSize = 3) {
  const selectedKey = getExportSelectedKey();
  const current = series.find((item) => item.key === selectedKey) || null;
  if (!current) {
    return [];
  }

  return series
    .filter((item) => item.month === current.month && item.year <= current.year)
    .slice(-windowSize);
}

function renderExportBarChart({ title, points, key, type = "amount", colorValue }) {
  if (!points.length) {
    return "";
  }

  const values = points.map((item) => item[key] ?? 0);
  const maxValue = Math.max(...values, 1);

  return `
    <article class="startup-chart-card">
      <div class="startup-chart-head">
        <div class="startup-chart-title">${title}</div>
      </div>
      <div class="startup-bars" style="--bar-count:${points.length};">
        ${points
          .map((item) => {
            const value = item[key] ?? 0;
            const barHeight = Math.max(8, Math.round((value / maxValue) * 120));
            const valueText = type === "count" ? formatExportCompanyCount(value) : formatExportAmount(value);
            return `
              <div class="startup-bar-item">
                <div class="startup-bar-value">${valueText}</div>
                <div class="startup-bar-track">
                  <div class="startup-bar" style="height:${barHeight}px; --bar-color:${colorValue};"></div>
                </div>
                <div class="startup-bar-label">${formatExportPeriod(item.date)}</div>
              </div>
            `;
          })
          .join("")}
      </div>
    </article>
  `;
}

function renderExportShareCharts() {
  const current = getExportRecord();
  if (!current) {
    return "";
  }

  const companyTotal = current["전체 기업 수"] ?? 0;
  const exportTotal = current["전체 수출금액"] ?? 0;
  const knownCompanyTotal =
    (current["중소기업 기업 수"] ?? 0) +
    (current["대기업 기업 수"] ?? 0) +
    (current["중견기업 기업 수"] ?? 0);
  const knownExportTotal =
    (current["중소기업 수출금액"] ?? 0) +
    (current["대기업 수출금액"] ?? 0) +
    (current["중견기업 수출금액"] ?? 0);
  const otherCompanyValue = Math.max(0, companyTotal - knownCompanyTotal);
  const otherExportValue = Math.max(0, exportTotal - knownExportTotal);

  const companyItems = [
    { name: "중소기업", value: current["중소기업 기업 수"], color: "#2c7be5" },
    { name: "대기업", value: current["대기업 기업 수"], color: "#ff7b63" },
    { name: "중견기업", value: current["중견기업 기업 수"], color: "#ffc857" },
    { name: "기타", value: otherCompanyValue, color: "#94a3b8" },
  ].filter((item) => item.value !== undefined && item.value !== null && !Number.isNaN(item.value) && item.value > 0);

  const exportItems = [
    { name: "중소기업", value: current["중소기업 수출금액"], color: "#2c7be5" },
    { name: "대기업", value: current["대기업 수출금액"], color: "#ff7b63" },
    { name: "중견기업", value: current["중견기업 수출금액"], color: "#ffc857" },
    { name: "기타", value: otherExportValue, color: "#94a3b8" },
  ].filter((item) => item.value !== undefined && item.value !== null && !Number.isNaN(item.value) && item.value > 0);

  const buildPie = (items, total, valueFormatter, stackedValue = false) => {
    const segments = [];
    let cursor = 0;
    items.forEach((item) => {
      const share = (item.value / total) * 100;
      const start = cursor;
      const end = cursor + share;
      segments.push(`${item.color} ${start}% ${end}%`);
      cursor = end;
    });

    return `
      <div class="startup-chart-card">
        <div class="investment-pie-layout">
          <div class="investment-pie" style="--pie-fill:${segments.join(", ")};"></div>
          <div class="investment-pie-legend">
            ${items
              .map(
                (item) => `
                  <div class="investment-pie-legend-item">
                    <span class="investment-pie-legend-swatch" style="--legend-color:${item.color};"></span>
                    <span class="investment-pie-legend-name">${item.name}</span>
                    <span class="investment-pie-legend-value${stackedValue ? " investment-pie-legend-value--stacked" : ""}">${valueFormatter(item.value)} (비중: ${formatNumber((item.value / total) * 100, 1)}%)</span>
                  </div>
                `,
              )
              .join("")}
          </div>
        </div>
      </div>
    `;
  };

  return `
    <article class="startup-chart-card">
      <div class="startup-chart-head">
        <div class="startup-chart-title">중소기업 비중</div>
      </div>
      <div class="export-pie-grid">
        <div class="export-pie-section">
          <div class="startup-subvalue">전체 기업 수 대비</div>
          ${buildPie(companyItems, companyTotal, (value) => formatExportCompanyCount(value), true)}
        </div>
        <div class="export-pie-section">
          <div class="startup-subvalue">전체 수출금액 대비</div>
          ${buildPie(exportItems, exportTotal, (value) => formatExportAmount(value))}
        </div>
      </div>
    </article>
  `;
}

function parseIndexPage(markdown, code) {
  const priceMatch = markdown.match(/_([\d,\.]+)_\s+([\d,\.]+)\s+(-?[\d,\.]+)%/);
  const timeRegex = new RegExp(`https://finance\\.naver\\.com/sise/sise_index\\.naver\\?code=${code}#\\)\\s*(\\d{4}\\.\\d{2}\\.\\d{2}\\s+[^\\n]+)`);
  const timeMatch = markdown.match(timeRegex);

  if (!priceMatch) {
    throw new Error(`${code} 지수 파싱 실패`);
  }

  const value = priceMatch[1];
  const delta = Number(priceMatch[2].replace(/,/g, ""));
  const rate = Number(priceMatch[3].replace(/,/g, ""));

  return {
    value,
    changeValue: delta,
    changePercent: rate,
    direction: rate > 0 ? "up" : rate < 0 ? "down" : "steady",
    metaTime: timeMatch ? timeMatch[1].trim() : "네이버 증권 기준",
    metaSource: code,
  };
}

function parseMarketIndex(markdown) {
  const usdMatch = markdown.match(/### 미국 USD\s+([\d,\.]+)\s+원\s+([\d,\.]+)\s+(상승|하락).*?(\d{4}\.\d{2}\.\d{2}\s+\d{2}:\d{2})/s);
  const wtiMatch = markdown.match(/### WTI\s+([\d,\.]+)\s+달러\s+([\d,\.]+)\s+(상승|하락).*?(\d{4}\.\d{2}\.\d{2})/s);
  const govt3yMatch = markdown.match(/\|\s+\[국고채 \(3년\)\][^|]+\|\s+([\d\.]+)\s+\|\s+!\[Image[^\]]*:\s*(상승|하락)\][^|]*\s+([\d\.]+)\s+\|/);

  if (!usdMatch || !wtiMatch || !govt3yMatch) {
    throw new Error("시장지표 파싱 실패");
  }

  return {
    usdkrw: {
      value: usdMatch[1],
      changeValue: Number(usdMatch[2].replace(/,/g, "")),
      changePercent: null,
      direction: usdMatch[3] === "상승" ? "up" : "down",
      metaTime: usdMatch[4],
      metaSource: "USDKRW",
    },
    wti: {
      value: wtiMatch[1],
      changeValue: Number(wtiMatch[2].replace(/,/g, "")),
      changePercent: null,
      direction: wtiMatch[3] === "상승" ? "up" : "down",
      metaTime: wtiMatch[4],
      metaSource: "WTI",
    },
    govt3y: {
      value: govt3yMatch[1],
      changeValue: Number(govt3yMatch[3]),
      changePercent: null,
      direction: govt3yMatch[2] === "상승" ? "up" : "down",
      metaTime: "네이버 증권 시장금리 표",
      metaSource: "KR3YT",
    },
  };
}

function readSnapshot() {
  try {
    return JSON.parse(localStorage.getItem(SNAPSHOT_KEY) || "{}");
  } catch {
    return {};
  }
}

function writeSnapshot(snapshot) {
  localStorage.setItem(SNAPSHOT_KEY, JSON.stringify(snapshot));
}

function parseAltAsset(payload, previousSnapshot) {
  const previousPrice = previousSnapshot?.[payload.symbol]?.price ?? null;
  const currentPrice = Number(payload.price);
  const delta = previousPrice === null ? null : currentPrice - previousPrice;
  const percent = previousPrice ? (delta / previousPrice) * 100 : null;

  return {
    value: formatNumber(currentPrice, payload.symbol === "BTC" ? 2 : 2),
    changeValue: delta,
    changePercent: percent,
    direction: delta === null ? "steady" : delta > 0 ? "up" : delta < 0 ? "down" : "steady",
    metaTime: payload.updatedAtReadable || payload.updatedAt,
    metaSource: `${payload.symbol} · refresh`,
    changeBasis: "직전 새로고침 대비",
    rawPrice: currentPrice,
  };
}

function getChangeClass(direction) {
  if (direction === "up") {
    return "is-up";
  }

  if (direction === "down") {
    return "is-down";
  }

  return "is-steady";
}

function formatChange(item) {
  if (item.changeValue === null || Number.isNaN(item.changeValue)) {
    return item.changeBasis ? "첫 새로고침 기준값" : "변동률 집계 중";
  }

  const sign = item.changeValue > 0 ? "+" : "";
  const changeNumber = Math.abs(item.changeValue) >= 100
    ? formatNumber(item.changeValue, 2)
    : formatNumber(item.changeValue, 2);

  if (item.changePercent === null || Number.isNaN(item.changePercent)) {
    return `${sign}${changeNumber}${item.changeBasis ? ` · ${item.changeBasis}` : ""}`;
  }

  const percentSign = item.changePercent > 0 ? "+" : "";
  return `${sign}${changeNumber} (${percentSign}${formatNumber(item.changePercent, 2)}%)${item.changeBasis ? ` · ${item.changeBasis}` : ""}`;
}

function renderMarketList(data) {
  const marketList = document.getElementById("market-list");

  marketList.innerHTML = marketItems
    .map((item) => {
      const row = data[item.id];
      const changeClass = row ? getChangeClass(row.direction) : "is-steady";
      const value = row ? `${row.value}${item.unit}` : "-";
      const change = row ? formatChange(row) : "데이터 없음";
      const metaTime = row ? row.metaTime : "연결 대기";
      const metaSource = row ? row.metaSource : item.code;

      return `
        <article class="market-item">
          <div class="market-row">
            <div class="market-left">
              <div class="asset-name">
                <span>${item.name}</span>
                <span class="asset-badge">${item.badge}</span>
              </div>
              <div class="asset-meta">
                <span class="asset-dot"></span>
                <span>${metaTime}</span>
                <span>|</span>
                <span>${metaSource}</span>
              </div>
            </div>
            <div class="market-right">
              <div class="asset-value">${value}</div>
              <div class="asset-change ${changeClass}">${change}</div>
            </div>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderSmeData() {
  if (!smeData.length) {
    document.getElementById("sme-grid").innerHTML = `
      <article class="sme-card">
        <div class="sme-title">데이터를 불러오지 못했습니다.</div>
        <div class="sme-footnote">${smeLoadError || "구글 스프레드시트 공개 설정과 시트 형식을 확인해 주세요."}</div>
      </article>
    `;
    document.getElementById("sme-charts").innerHTML = "";
    return;
  }

  const selectedYear = document.getElementById("sme-year-select").value || smeYears[smeYears.length - 1];
  const selectedIndex = smeYears.findIndex((year) => year === selectedYear);
  const previousYear = selectedIndex > 0 ? smeYears[selectedIndex - 1] : null;
  const smeGrid = document.getElementById("sme-grid");

  smeGrid.innerHTML = smeData
    .map((item) => {
      const currentValue = item.years[selectedYear].sme;
      const totalValue = item.years[selectedYear].total;
      const previousValue = previousYear ? item.years[previousYear]?.sme : null;
      const growthText = formatSmeDelta(item, currentValue, previousValue);
      const share = (currentValue / totalValue) * 100;
      const fill = Math.max(6, Math.round(share));

      return `
        <article class="sme-card">
          <div class="sme-card-header">
            <div class="sme-title">${item.title}</div>
          </div>
          <div class="sme-chart-row">
            <div class="sme-donut" style="--fill:${fill}; --donut-color:${item.color};">
              <div class="sme-donut-inner">
                <div class="sme-donut-percent">${formatNumber(share, 1)}%</div>
              </div>
            </div>
            <div class="sme-metric-copy">
              <div class="sme-value">${formatSmeValue(item, currentValue)}${growthText}</div>
              <span class="sme-unit"></span>
              <div class="sme-footnote">${formatTotalFootnote(item, totalValue)}</div>
            </div>
          </div>
        </article>
      `;
    })
    .join("");
}

function getFilteredSmeYears() {
  const selectedYear =
    document.getElementById("sme-year-select")?.value || smeYears[smeYears.length - 1];
  const selectedIndex = smeYears.findIndex((year) => year === selectedYear);
  if (selectedIndex < 0) {
    return smeYears.slice(-3);
  }

  const startIndex = Math.max(0, selectedIndex - 2);
  return smeYears.slice(startIndex, selectedIndex + 1);
}

function renderSmeCharts() {
  const charts = document.getElementById("sme-charts");

  if (!smeData.length) {
    charts.innerHTML = "";
    return;
  }

  const filteredYears = getFilteredSmeYears();

  charts.innerHTML = smeData
    .map((item) => {
      const values = filteredYears.map((year) => item.years[year]?.sme ?? 0);
      const maxValue = Math.max(...values, 1);

      return `
        <article class="startup-chart-card">
          <div class="startup-chart-head">
            <div class="startup-chart-title">${item.title}</div>
          </div>
          <div class="startup-bars" style="--bar-count:${filteredYears.length};">
            ${filteredYears
              .map((year) => {
                const value = item.years[year]?.sme ?? 0;
                const barHeight = Math.max(8, Math.round((value / maxValue) * 120));
                return `
                  <div class="startup-bar-item">
                    <div class="sme-bar-value">${formatSmeValue(item, value)}</div>
                    <div class="startup-bar-track">
                      <div class="startup-bar" style="height:${barHeight}px; --bar-color:${item.color};"></div>
                    </div>
                    <div class="startup-bar-label">${year}</div>
                  </div>
                `;
              })
              .join("")}
          </div>
        </article>
      `;
    })
    .join("");
}

function renderStartupSummary() {
  const summary = document.getElementById("startup-summary");

  if (!startupSeries.length) {
    summary.innerHTML = `
      <article class="startup-summary-card">
        <div class="startup-value">창업 데이터를 불러오지 못했습니다.</div>
        <div class="startup-subvalue">${startupLoadError || "구글 시트에 창업기업수와 기술기반업종 창업기업 수 데이터가 있는지 확인해 주세요."}</div>
      </article>
    `;
    return;
  }

  const selectedYear =
    document.getElementById("startup-year-select")?.value || startupSeries[startupSeries.length - 1]?.year;
  const currentIndex = startupSeries.findIndex((item) => item.year === selectedYear);
  const current = currentIndex >= 0 ? startupSeries[currentIndex] : startupSeries[startupSeries.length - 1];
  const previous = currentIndex > 0 ? startupSeries[currentIndex - 1] : null;
  const shareText = current.techShare !== undefined ? `${formatNumber(current.techShare, 1)}%` : "-";
  const startupGrowth = formatYoYGrowth(current.startupCount, previous?.startupCount);
  const techGrowth = formatYoYGrowth(current.techStartupCount, previous?.techStartupCount);

  summary.innerHTML = `
    <article class="startup-summary-card">
      <div class="startup-summary-grid">
        <div class="startup-metric">
          <div class="startup-kicker">${current.year}년 창업기업수</div>
          <div class="startup-value">${formatManCount(current.startupCount || 0)}${startupGrowth}</div>
        </div>
        <div class="startup-metric">
          <div class="startup-kicker">${current.year}년 기술기반업종 창업기업수</div>
          <div class="startup-value">${formatManCount(current.techStartupCount || 0)}${techGrowth}</div>
          <div class="startup-subvalue"><span class="startup-share-label">전체 창업기업 수 대비 기술기반업종 비중</span>: <span class="startup-share-value">${shareText}</span></div>
        </div>
      </div>
    </article>
  `;
}

function getFilteredStartupSeries(windowSize = 5) {
  const selectedYear =
    document.getElementById("startup-year-select")?.value || startupSeries[startupSeries.length - 1]?.year;
  const selectedIndex = startupSeries.findIndex((item) => item.year === selectedYear);
  if (selectedIndex < 0) {
    return startupSeries.slice(-windowSize);
  }

  const startIndex = Math.max(0, selectedIndex - (windowSize - 1));
  return startupSeries.slice(startIndex, selectedIndex + 1);
}

function renderStartupBarChart({ title, key, type, colorClass, colorValue }) {
  const filteredSeries = getFilteredStartupSeries(3);

  if (!filteredSeries.length) {
    return "";
  }

  const values = filteredSeries.map((item) => item[key] ?? 0);
  const maxValue = Math.max(...values, 1);

  return `
    <article class="startup-chart-card">
      <div class="startup-chart-head">
        <div class="startup-chart-title">${title}</div>
      </div>
      <div class="startup-bars" style="--bar-count:${filteredSeries.length};">
        ${filteredSeries
          .map((item) => {
            const value = item[key] ?? 0;
            const barHeight = Math.max(8, Math.round((value / maxValue) * 120));
            return `
              <div class="startup-bar-item">
                <div class="startup-bar-value">${formatStartupMetric(value, type)}</div>
                <div class="startup-bar-track">
                  <div class="startup-bar ${colorClass}" style="height:${barHeight}px; --bar-color:${colorValue};"></div>
                </div>
                <div class="startup-bar-label">${item.year}</div>
              </div>
            `;
          })
          .join("")}
      </div>
    </article>
  `;
}

function renderStartupLineChart({ title, key }) {
  const filteredSeries = getFilteredStartupSeries(5);

  if (!filteredSeries.length) {
    return "";
  }

  const values = filteredSeries.map((item) => item[key] ?? 0);
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  const range = maxValue - minValue || 1;
  const width = 320;
  const height = 150;
  const paddingX = 22;
  const paddingTop = 26;
  const paddingBottom = 18;
  const usableWidth = width - paddingX * 2;
  const usableHeight = height - paddingTop - paddingBottom;

  const points = filteredSeries.map((item, index) => {
    const x =
      filteredSeries.length === 1
        ? width / 2
        : paddingX + (usableWidth / (filteredSeries.length - 1)) * index;
    const value = item[key] ?? 0;
    const y = paddingTop + (1 - (value - minValue) / range) * usableHeight;
    return { x, y, value, year: item.year };
  });

  const path = points
    .map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`)
    .join(" ");

  return `
    <article class="startup-chart-card">
      <div class="startup-chart-head">
        <div class="startup-chart-title">${title}</div>
      </div>
      <div class="startup-line-chart" style="--bar-count:${filteredSeries.length};">
        <svg class="startup-line-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none" aria-hidden="true">
          <line class="startup-line-grid" x1="${paddingX}" y1="${height - paddingBottom}" x2="${width - paddingX}" y2="${height - paddingBottom}"></line>
          <path class="startup-line-path" d="${path}"></path>
          ${points
            .map(
              (point) => `
                <text class="startup-line-value" x="${point.x}" y="${Math.max(12, point.y - 10)}">${formatNumber(point.value, 1)}%</text>
                <circle class="startup-line-point" cx="${point.x}" cy="${point.y}" r="4"></circle>
              `,
            )
            .join("")}
        </svg>
        <div class="startup-line-years">
          ${points.map((point) => `<div class="startup-line-year">${point.year}</div>`).join("")}
        </div>
      </div>
    </article>
  `;
}

function renderStartupCharts() {
  const charts = document.getElementById("startup-charts");

  if (!startupSeries.length) {
    charts.innerHTML = "";
    return;
  }

  charts.innerHTML = [
    renderStartupBarChart({
      title: "창업기업수",
      key: "startupCount",
      type: "count",
      colorClass: "is-blue",
      colorValue: "#2c7be5",
    }),
    renderStartupBarChart({
      title: "기술기반업종 창업기업 수",
      key: "techStartupCount",
      type: "count",
      colorClass: "is-sky",
      colorValue: "#59a7ff",
    }),
    renderStartupLineChart({
      title: "기술기반업종 비중",
      key: "techShare",
    }),
  ].join("");
}

function renderLoanSummary() {
  const summary = document.getElementById("loan-summary");

  if (!loanSeries.length && !delinquencySeries.length) {
    summary.innerHTML = `
      <article class="startup-summary-card">
        <div class="startup-value">대출 데이터를 불러오지 못했습니다.</div>
        <div class="startup-subvalue">${loanLoadError || "구글 시트 대출 데이터 구성을 확인해 주세요."}</div>
      </article>
    `;
    return;
  }

  const latest = getLoanLatestPointForYear(getLoanSelectedYear());
  if (!latest) {
    summary.innerHTML = "";
    return;
  }

  const previousSameMonth = getPreviousYearSameMonthLoanPoint(latest.date);
  const balanceGrowth = formatLoanYoYGrowth(latest.balance, previousSameMonth?.balance);
  const netIncreaseGrowth = formatLoanYoYGrowthWithDelta(latest.netIncrease, previousSameMonth?.netIncrease);
  const latestDelinquency = getDelinquencyLatestPointForYear(getLoanSelectedYear());
  const previousDelinquency = getPreviousYearSameMonthDelinquencyPoint(latestDelinquency?.date);
  const delinquencyDelta = formatRateDelta(latestDelinquency?.smeRate, previousDelinquency?.smeRate);
  const corporateDelta = formatRateDelta(latestDelinquency?.corporateRate, previousDelinquency?.corporateRate);
  const solePropDelta = formatRateDelta(latestDelinquency?.solePropRate, previousDelinquency?.solePropRate);
  const delinquencyDateLabel = latestDelinquency?.date ? formatYearMonth(latestDelinquency.date) : "-";

  summary.innerHTML = `
    <article class="startup-summary-card">
      <div class="startup-summary-grid">
        <div class="startup-metric">
          <div class="startup-kicker">${formatYearMonth(latest.date)} 은행권 중소기업대출잔액</div>
          <div class="startup-value">${formatLoanValue(latest.balance)}${balanceGrowth}</div>
        </div>
        <div class="startup-metric">
          <div class="startup-kicker">${formatYearMonth(latest.date)} 은행권 중소기업대출 순증</div>
          <div class="startup-value">${formatLoanValue(latest.netIncrease)}${netIncreaseGrowth}</div>
        </div>
      </div>
    </article>
    <article class="startup-summary-card">
      <div class="startup-summary-grid">
        <div class="startup-metric">
          <div class="startup-kicker">${delinquencyDateLabel} 중소기업 대출 연체율</div>
          <div class="startup-value">${formatRateValue(latestDelinquency?.smeRate)}${delinquencyDelta}</div>
        </div>
        <div class="startup-metric">
          <div class="startup-kicker">${delinquencyDateLabel} 중소법인 대출 연체율</div>
          <div class="startup-value">${formatRateValue(latestDelinquency?.corporateRate)}${corporateDelta}</div>
        </div>
        <div class="startup-metric">
          <div class="startup-kicker">${delinquencyDateLabel} 개인사업자 대출 연체율</div>
          <div class="startup-value">${formatRateValue(latestDelinquency?.solePropRate)}${solePropDelta}</div>
        </div>
      </div>
    </article>
  `;
}

function renderLoanBarChart({ title, labelMode, points, valueKey, colorValue }) {
  if (!points.length) {
    return "";
  }

  const values = points.map((item) => item[valueKey] ?? 0);
  const maxValue = Math.max(...values, 1);

  return `
    <article class="startup-chart-card">
      <div class="startup-chart-head">
        <div class="startup-chart-title">${title}</div>
      </div>
      <div class="startup-bars" style="--bar-count:${points.length};">
        ${points
          .map((item) => {
            const value = item[valueKey] ?? 0;
            const barHeight = Math.max(8, Math.round((value / maxValue) * 120));
            const label =
              labelMode === "recent-month"
                ? `${item.year}.${item.month}월`
                : labelMode === "mixed"
                  ? `${item.year}.${item.month === 12 ? "12월" : `${item.month}월`}`
                  : `${item.year}.12월`;
            return `
              <div class="startup-bar-item">
                <div class="startup-bar-value">${formatLoanValue(value)}</div>
                <div class="startup-bar-track">
                  <div class="startup-bar" style="height:${barHeight}px; --bar-color:${colorValue};"></div>
                </div>
                <div class="startup-bar-label">${label}</div>
              </div>
            `;
          })
          .join("")}
      </div>
    </article>
  `;
}

function renderLoanLineChart({ title, seriesConfig, points }) {
  if (!points.length) {
    return "";
  }

  const allValues = points.flatMap((point) =>
    seriesConfig
      .map((series) => point[series.key])
      .filter((value) => value !== undefined && value !== null && !Number.isNaN(value)),
  );

  if (!allValues.length) {
    return "";
  }

  const minValue = Math.min(...allValues);
  const maxValue = Math.max(...allValues);
  const range = maxValue - minValue || 1;
  const width = 320;
  const height = 150;
  const paddingX = 22;
  const paddingTop = 24;
  const paddingBottom = 18;
  const usableWidth = width - paddingX * 2;
  const usableHeight = height - paddingTop - paddingBottom;

  const xForIndex = (index) =>
    points.length === 1
      ? width / 2
      : paddingX + (usableWidth / (points.length - 1)) * index;

  const yForValue = (value) =>
    paddingTop + (1 - (value - minValue) / range) * usableHeight;

  const seriesPoints = seriesConfig.map((series) => {
    const plotted = points
      .map((point, index) => {
        const value = point[series.key];
        if (value === undefined || value === null || Number.isNaN(value)) {
          return null;
        }

        return {
          x: xForIndex(index),
          y: yForValue(value),
          value,
        };
      })
      .filter(Boolean);

    return {
      ...series,
      plotted,
      path: plotted
        .map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`)
        .join(" "),
    };
  });

  return `
    <article class="startup-chart-card">
      <div class="startup-chart-head">
        <div class="startup-chart-title">${title}</div>
        <div class="loan-line-legend">
          ${seriesPoints
            .map(
              (series) => `
                <div class="loan-line-legend-item">
                  <span class="loan-line-legend-swatch" style="--legend-color:${series.color};"></span>
                  <span>${series.label}</span>
                </div>
              `,
            )
            .join("")}
        </div>
      </div>
      <div class="startup-line-chart" style="--bar-count:${points.length};">
        <svg class="startup-line-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none" aria-hidden="true">
          <line class="startup-line-grid" x1="${paddingX}" y1="${height - paddingBottom}" x2="${width - paddingX}" y2="${height - paddingBottom}"></line>
          ${seriesPoints
            .map(
              (series) => `
                <path d="${series.path}" fill="none" stroke="${series.color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>
                ${series.plotted
                  .map(
                    (point) => `
                      <text class="startup-line-value" x="${point.x}" y="${series.label === "개인사업자" ? Math.min(height - 6, point.y + 18) : Math.max(12, point.y - 10)}" style="fill:${series.color};">${formatNumber(point.value, 2)}%</text>
                      <circle cx="${point.x}" cy="${point.y}" r="4" fill="${series.color}" stroke="#ffffff" stroke-width="2"></circle>
                    `,
                  )
                  .join("")}
              `,
            )
            .join("")}
        </svg>
        <div class="startup-line-years">
          ${points.map((point) => `<div class="startup-line-year">${point.year}.${point.month}월</div>`).join("")}
        </div>
      </div>
    </article>
  `;
}

function renderLoanCharts() {
  const charts = document.getElementById("loan-charts");

  if (!loanSeries.length && !delinquencySeries.length) {
    charts.innerHTML = "";
    return;
  }

  const loanBalanceYears = getRecentLoanBalancePoints();
  const loanNetIncreaseYears = getRecentLoanNetIncreasePoints();
  const loanNetIncreaseLatestMonthYears = getRecentLoanNetIncreaseLatestMonthPoints();
  const delinquencyPoints = getFilteredDelinquencySeries(5);

  charts.innerHTML = [
    renderLoanBarChart({
      title: "은행권 중소기업대출잔액",
      labelMode: "mixed",
      points: loanBalanceYears,
      valueKey: "balance",
      colorValue: "#2c7be5",
    }),
    renderLoanBarChart({
      title: "은행권 중소기업대출 순증",
      labelMode: "year-end",
      points: loanNetIncreaseYears,
      valueKey: "netIncrease",
      colorValue: "#59a7ff",
    }),
    renderLoanBarChart({
      title: "은행권 중소기업대출 순증",
      labelMode: "recent-month",
      points: loanNetIncreaseLatestMonthYears,
      valueKey: "netIncrease",
      colorValue: "#88c2ff",
    }),
    renderLoanLineChart({
      title: "중소기업대출 연체율",
      points: delinquencyPoints,
      seriesConfig: [
        { key: "smeRate", label: "중소기업", color: "#ff675f" },
        { key: "largeCompanyRate", label: "대기업", color: "#7fb8ff" },
      ],
    }),
    renderLoanLineChart({
      title: "중소기업대출 연체율: 중소법인 vs 개인사업자",
      points: delinquencyPoints,
      seriesConfig: [
        { key: "corporateRate", label: "중소법인", color: "#ff7b63" },
        { key: "solePropRate", label: "개인사업자", color: "#ffb347" },
      ],
    }),
  ]
    .filter(Boolean)
    .join("");
}

function renderInvestmentSummary() {
  const summary = document.getElementById("investment-summary");

  if (!investmentSeries.length) {
    summary.innerHTML = `
      <article class="startup-summary-card">
        <div class="startup-value">투자 데이터를 불러오지 못했습니다.</div>
        <div class="startup-subvalue">${investmentLoadError || "구글 시트 투자 데이터 구성을 확인해 주세요."}</div>
      </article>
    `;
    return;
  }

  const pointLabel = formatInvestmentPointLabel();
  const currentOverview = getInvestmentRecord(investmentSeries);
  const previousOverview = getPreviousYearInvestmentRecord(investmentSeries);
  const currentStage = getInvestmentRecord(investmentStageSeries);
  const currentSector = getInvestmentRecord(investmentSectorSeries);
  const currentSource = getInvestmentRecord(investmentSourceSeries);
  const benchmarkTotal = currentOverview?.["신규 벤처투자금액"] ?? 0;
  const sourceTotal = (currentSource?.정책금융 ?? 0) + (currentSource?.민간부문 ?? 0);
  const stageTotal =
    (currentStage?.["초기 투자(3년 이내)"] ?? 0) +
    (currentStage?.["중기 투자(3~7년 이내)"] ?? 0) +
    (currentStage?.["후기 투자(7년 초과)"] ?? 0);

  const topSectors = [
    { name: "ICT서비스", value: currentSector?.["ICT서비스"] },
    { name: "바이오·의료", value: currentSector?.["바이오·의료"] },
    { name: "전기·기계·장비", value: currentSector?.["전기·기계·장비"] },
    { name: "ICT제조", value: currentSector?.["ICT제조"] },
    { name: "유통·서비스", value: currentSector?.["유통·서비스"] },
    { name: "화학·소재", value: currentSector?.["화학·소재"] },
    { name: "영상·공연·음반", value: currentSector?.["영상·공연·음반"] },
    { name: "게임", value: currentSector?.게임 },
    { name: "기타", value: currentSector?.기타 },
  ]
    .filter((item) => item.value !== undefined && item.value !== null && !Number.isNaN(item.value))
    .sort((a, b) => b.value - a.value)
    .slice(0, 3);

  summary.innerHTML = `
    <article class="startup-summary-card">
      <div class="startup-summary-grid">
        <div class="startup-metric">
          <div class="startup-kicker">${pointLabel} 신규 벤처투자금액</div>
          <div class="startup-value">${formatEokValue(currentOverview?.["신규 벤처투자금액"])}${formatInvestmentDelta(currentOverview?.["신규 벤처투자금액"], previousOverview?.["신규 벤처투자금액"], "amount")}</div>
        </div>
        <div class="startup-metric">
          <div class="startup-kicker">${pointLabel} 벤처펀드 결정금액</div>
          <div class="startup-value">${formatEokValue(currentOverview?.["벤처펀드 결정금액"])}${formatInvestmentDelta(currentOverview?.["벤처펀드 결정금액"], previousOverview?.["벤처펀드 결정금액"], "amount")}</div>
        </div>
        <div class="startup-metric">
          <div class="startup-kicker">${pointLabel} 벤처펀드 결성 수</div>
          <div class="startup-value">${formatCountValue(currentOverview?.["벤처펀드 결성 수"])}${formatInvestmentDelta(currentOverview?.["벤처펀드 결성 수"], previousOverview?.["벤처펀드 결성 수"], "count")}</div>
        </div>
      </div>
    </article>
    <article class="startup-summary-card">
      <div class="startup-summary-title">&lt; 업력별 투자 &gt;</div>
      <div class="startup-summary-grid startup-summary-grid--three-col">
        <div class="startup-metric investment-metric investment-metric--blue">
          <div class="startup-kicker investment-kicker investment-kicker--blue">초기(3년 이내)</div>
          <div class="startup-value investment-value investment-value--blue investment-value--compact">${formatEokValue(currentStage?.["초기 투자(3년 이내)"])}</div>
          ${formatInvestmentShare(currentStage?.["초기 투자(3년 이내)"], stageTotal)}
        </div>
        <div class="startup-metric investment-metric investment-metric--orange">
          <div class="startup-kicker investment-kicker investment-kicker--orange">중기(3~7년 이내)</div>
          <div class="startup-value investment-value investment-value--orange investment-value--compact">${formatEokValue(currentStage?.["중기 투자(3~7년 이내)"])}</div>
          ${formatInvestmentShare(currentStage?.["중기 투자(3~7년 이내)"], stageTotal)}
        </div>
        <div class="startup-metric investment-metric investment-metric--violet">
          <div class="startup-kicker investment-kicker investment-kicker--violet">후기(7년 초과)</div>
          <div class="startup-value investment-value investment-value--violet investment-value--compact">${formatEokValue(currentStage?.["후기 투자(7년 초과)"])}</div>
          ${formatInvestmentShare(currentStage?.["후기 투자(7년 초과)"], stageTotal)}
        </div>
      </div>
    </article>
    <article class="startup-summary-card">
      <div class="startup-summary-title">&lt; 업종별 투자 &gt;</div>
      <div class="startup-summary-grid startup-summary-grid--three-col">
        ${topSectors
          .map(
            (item, index) => `
                <div class="startup-metric investment-metric investment-metric--${index === 0 ? "blue" : index === 1 ? "orange" : "violet"}">
                  <div class="startup-kicker investment-kicker investment-kicker--${index === 0 ? "blue" : index === 1 ? "orange" : "violet"}${item.name === "전기·기계·장비" ? " investment-kicker--compact" : ""}">${item.name}</div>
                  <div class="startup-value investment-value investment-value--${index === 0 ? "blue" : index === 1 ? "orange" : "violet"} investment-value--compact">${formatEokValue(item.value)}</div>
                  ${formatInvestmentShare(item.value, benchmarkTotal)}
                </div>
            `,
          )
          .join("")}
      </div>
    </article>
    <article class="startup-summary-card">
      <div class="startup-summary-title">&lt; 출자자별 출자금액 &gt;</div>
      <div class="startup-summary-grid startup-summary-grid--two-col">
        <div class="startup-metric investment-metric investment-metric--blue">
          <div class="startup-kicker investment-kicker investment-kicker--blue">정책금융</div>
          <div class="startup-value investment-value investment-value--blue">${formatEokValue(currentSource?.정책금융)}</div>
          ${formatInvestmentShare(currentSource?.정책금융, sourceTotal)}
        </div>
        <div class="startup-metric investment-metric investment-metric--orange">
          <div class="startup-kicker investment-kicker investment-kicker--orange">민간부문</div>
          <div class="startup-value investment-value investment-value--orange">${formatEokValue(currentSource?.민간부문)}</div>
          ${formatInvestmentShare(currentSource?.민간부문, sourceTotal)}
        </div>
      </div>
    </article>
  `;
}

function renderInvestmentBarChart({ title, points, key, type = "amount", colorValue }) {
  if (!points.length) {
    return "";
  }

  const values = points.map((item) => item[key] ?? 0);
  const maxValue = Math.max(...values, 1);

  return `
    <article class="startup-chart-card">
      <div class="startup-chart-head">
        <div class="startup-chart-title">${title}</div>
      </div>
      <div class="startup-bars" style="--bar-count:${points.length};">
        ${points
          .map((item) => {
            const value = item[key] ?? 0;
            const barHeight = Math.max(8, Math.round((value / maxValue) * 120));
            const valueText = type === "count" ? formatCountValue(value) : formatEokValue(value);
            return `
              <div class="startup-bar-item">
                <div class="startup-bar-value">${valueText}</div>
                <div class="startup-bar-track">
                  <div class="startup-bar" style="height:${barHeight}px; --bar-color:${colorValue};"></div>
                </div>
                <div class="startup-bar-label">${formatInvestmentPeriod(item.date)}</div>
              </div>
            `;
          })
          .join("")}
      </div>
    </article>
  `;
}

function renderInvestmentPieChart({ title, items, totalLabel, denominatorValue }) {
  const validItems = items.filter((item) => item.value !== undefined && item.value !== null && !Number.isNaN(item.value) && item.value > 0);
  if (!validItems.length) {
    return "";
  }

  const total = validItems.reduce((sum, item) => sum + item.value, 0);
  const baseTotal = denominatorValue && denominatorValue > 0 ? denominatorValue : total;
  const segments = [];
  let cursor = 0;

  validItems.forEach((item) => {
    const share = (item.value / total) * 100;
    const start = cursor;
    const end = cursor + share;
    segments.push(`${item.color} ${start}% ${end}%`);
    cursor = end;
  });

  return `
    <article class="startup-chart-card">
      <div class="startup-chart-head">
        <div class="startup-chart-title">${title}</div>
      </div>
      <div class="investment-pie-layout">
        <div class="investment-pie" style="--pie-fill:${segments.join(", ")};"></div>
        <div class="investment-pie-legend">
          ${validItems
            .map(
              (item) => `
                <div class="investment-pie-legend-item">
                  <span class="investment-pie-legend-swatch" style="--legend-color:${item.color};"></span>
                  <span class="investment-pie-legend-name">${item.name}</span>
                  <span class="investment-pie-legend-value">${formatNumber(item.value, 0)}억원 (비중: ${formatNumber((item.value / baseTotal) * 100, 1)}%)</span>
                </div>
              `,
            )
            .join("")}
        </div>
      </div>
    </article>
  `;
}

function renderInvestmentCharts() {
  const charts = document.getElementById("investment-charts");

  if (!investmentSeries.length) {
    charts.innerHTML = "";
    return;
  }

  const currentStage = getInvestmentRecord(investmentStageSeries);
  const currentSector = getInvestmentRecord(investmentSectorSeries);
  const currentSource = getInvestmentRecord(investmentSourceSeries);
  const overviewRecent = getRecentInvestmentSeries(investmentSeries, 3);
  const currentOverview = getInvestmentRecord(investmentSeries);
  const benchmarkTotal = currentOverview?.["신규 벤처투자금액"] ?? 0;

  charts.innerHTML = [
    renderInvestmentBarChart({
      title: "신규 벤처투자금액",
      points: overviewRecent,
      key: "신규 벤처투자금액",
      colorValue: "#2c7be5",
    }),
    renderInvestmentBarChart({
      title: "벤처펀드 결성금액",
      points: overviewRecent,
      key: "벤처펀드 결정금액",
      colorValue: "#59a7ff",
    }),
    renderInvestmentPieChart({
      title: "업력별 투자 비중",
      totalLabel: "신규 벤처투자금액 기준",
      denominatorValue: benchmarkTotal,
      items: [
        { name: "초기 투자(3년 이내)", value: currentStage?.["초기 투자(3년 이내)"], color: "#2c7be5" },
        { name: "중기 투자(3~7년 이내)", value: currentStage?.["중기 투자(3~7년 이내)"], color: "#59a7ff" },
        { name: "후기 투자(7년 초과)", value: currentStage?.["후기 투자(7년 초과)"], color: "#8bd3ff" },
      ],
    }),
    renderInvestmentPieChart({
      title: "업종별 투자 비중",
      totalLabel: "신규 벤처투자금액 기준",
      denominatorValue: benchmarkTotal,
      items: [
        { name: "ICT서비스", value: currentSector?.["ICT서비스"], color: "#2c7be5" },
        { name: "바이오·의료", value: currentSector?.["바이오·의료"], color: "#59a7ff" },
        { name: "전기·기계·장비", value: currentSector?.["전기·기계·장비"], color: "#8bd3ff" },
        { name: "ICT제조", value: currentSector?.["ICT제조"], color: "#ff7b63" },
        { name: "유통·서비스", value: currentSector?.["유통·서비스"], color: "#ff9c73" },
        { name: "화학·소재", value: currentSector?.["화학·소재"], color: "#ffc857" },
        { name: "영상·공연·음반", value: currentSector?.["영상·공연·음반"], color: "#9c89ff" },
        { name: "게임", value: currentSector?.게임, color: "#7a6ff0" },
        { name: "기타", value: currentSector?.기타, color: "#94a3b8" },
      ],
    }),
    renderInvestmentPieChart({
      title: "정책금융 출자 비중",
      totalLabel: "신규 벤처투자금액 기준",
      denominatorValue: currentSource?.정책금융,
      items: [
        { name: "모태펀드", value: currentSource?.모태펀드, color: "#2c7be5" },
        { name: "성장금융", value: currentSource?.성장금융, color: "#59a7ff" },
        { name: "산업은행", value: currentSource?.산업은행, color: "#8bd3ff" },
        { name: "기타 정책금융", value: currentSource?.["기타 정책금융"], color: "#b8e4ff" },
      ],
    }),
    renderInvestmentPieChart({
      title: "민간부문 출자 비중",
      totalLabel: "신규 벤처투자금액 기준",
      denominatorValue: benchmarkTotal,
      items: [
        { name: "개인", value: currentSource?.개인, color: "#ff7b63" },
        { name: "일반법인", value: currentSource?.일반법인, color: "#ff9c73" },
        { name: "금융기관(산은 제외)", value: currentSource?.["금융기관(산은 제외)"], color: "#ffc857" },
        { name: "연기금 및 공제회", value: currentSource?.["연기금 및 공제회"], color: "#9c89ff" },
        { name: "VC", value: currentSource?.VC, color: "#7a6ff0" },
        { name: "기타단체 및 외국인", value: currentSource?.["기타단체 및 외국인"], color: "#94a3b8" },
      ],
    }),
  ].join("");
}

function renderExportSummary() {
  const summary = document.getElementById("export-summary");

  if (!exportSeries.length) {
    summary.innerHTML = `
      <article class="startup-summary-card">
        <div class="startup-value">수출 데이터를 불러오지 못했습니다.</div>
        <div class="startup-subvalue">${exportLoadError || "구글 시트 수출 데이터 구성을 확인해 주세요."}</div>
      </article>
    `;
    return;
  }

  const current = getExportRecord();
  const currentCountry = getExportCountryRecord();
  const previous = getPreviousYearExportRecord(exportSeries);
  const previousCountry = getPreviousYearExportRecord(exportCountrySeries);
  const companyShare =
    current?.["전체 기업 수"] && current?.["중소기업 기업 수"] !== undefined
      ? (current["중소기업 기업 수"] / current["전체 기업 수"]) * 100
      : null;
  const exportShare =
    current?.["전체 수출금액"] && current?.["중소기업 수출금액"] !== undefined
      ? (current["중소기업 수출금액"] / current["전체 수출금액"]) * 100
      : null;
  const exportCountries = [
    { name: "미국", key: "미국 수출금액", color: "blue" },
    { name: "중국", key: "중국 수출금액", color: "orange" },
    { name: "일본", key: "일본 수출금액", color: "violet" },
    { name: "베트남", key: "베트남 수출금액", color: "blue" },
    { name: "홍콩", key: "홍콩 수출금액", color: "orange" },
    { name: "대만", key: "대만 수출금액", color: "violet" },
    { name: "싱가포르", key: "싱가포르 수출금액", color: "blue" },
  ];

  const topCountries = exportCountries
    .map((item) => ({
      ...item,
      value: currentCountry?.[item.key],
      previousValue: previousCountry?.[item.key],
    }))
    .filter((item) => item.value !== undefined && item.value !== null && !Number.isNaN(item.value))
    .sort((a, b) => b.value - a.value)
    .slice(0, 5);

  summary.innerHTML = `
    <article class="startup-summary-card">
      <div class="startup-summary-grid startup-summary-grid--two-col">
        <div class="startup-metric">
          <div class="startup-kicker">수출 중소기업 수</div>
          <div class="startup-value">${formatExportCompanyCount(current?.["중소기업 기업 수"])}</div>
          <div class="startup-subvalue">(비중: <span class="startup-share-value">${companyShare === null ? "-" : `${formatNumber(companyShare, 1)}%`}</span>)</div>
          <div class="startup-subvalue">${formatYoYGrowth(current?.["중소기업 기업 수"], previous?.["중소기업 기업 수"])}</div>
        </div>
        <div class="startup-metric">
          <div class="startup-kicker">중소기업 수출금액</div>
          <div class="startup-value">${formatExportAmount(current?.["중소기업 수출금액"])}</div>
          <div class="startup-subvalue">(비중: <span class="startup-share-value">${exportShare === null ? "-" : `${formatNumber(exportShare, 1)}%`}</span>)</div>
          <div class="startup-subvalue">${formatYoYGrowth(current?.["중소기업 수출금액"], previous?.["중소기업 수출금액"])}</div>
        </div>
      </div>
    </article>
    <article class="startup-summary-card">
      <div class="startup-summary-title">&lt; 중소기업 주요 수출국가 &gt;</div>
      <div class="startup-summary-grid">
        ${topCountries
          .map(
            (item) => `
              <div class="startup-metric investment-metric investment-metric--${item.color}">
                <div class="startup-kicker investment-kicker investment-kicker--${item.color}">${item.name}</div>
                <div class="startup-value investment-value investment-value--${item.color}">${formatExportAmount(item.value)}${formatYoYGrowth(item.value, item.previousValue)}</div>
              </div>
            `,
          )
          .join("")}
      </div>
    </article>
  `;
}

function renderExportCharts() {
  const charts = document.getElementById("export-charts");

  if (!exportSeries.length) {
    charts.innerHTML = "";
    return;
  }

  const recentExport = getRecentExportSeries(exportSeries, 3);
  charts.innerHTML = [
    renderExportBarChart({
      title: "수출 중소기업 수",
      points: recentExport,
      key: "중소기업 기업 수",
      type: "count",
      colorValue: "#2c7be5",
    }),
    renderExportBarChart({
      title: "중소기업 수출금액",
      points: recentExport,
      key: "중소기업 수출금액",
      type: "amount",
      colorValue: "#59a7ff",
    }),
    renderExportShareCharts(),
  ].join("");
}

function initStartupYearSelect() {
  const select = document.getElementById("startup-year-select");
  if (!startupSeries.length) {
    select.innerHTML = "";
    return;
  }

  select.innerHTML = startupSeries
    .map((item) => `<option value="${item.year}">${item.year}</option>`)
    .join("");
  select.value = startupSeries[startupSeries.length - 1].year;
}

function initLoanYearSelect() {
  const select = document.getElementById("loan-year-select");
  if (!loanYears.length) {
    select.innerHTML = "";
    return;
  }

  select.innerHTML = loanYears
    .map((year) => `<option value="${year}">${year}</option>`)
    .join("");
  select.value = String(loanYears[loanYears.length - 1]);
  select.onchange = () => {
    renderLoanSummary();
    renderLoanCharts();
  };
}

function initInvestmentDateSelect() {
  const select = document.getElementById("investment-date-select");
  if (!investmentDates.length) {
    select.innerHTML = "";
    return;
  }

  select.innerHTML = investmentDates
    .map((key) => {
      const record = getInvestmentRecord(investmentSeries, key);
      return `<option value="${key}">${record?.date ? formatInvestmentPeriod(record.date) : key}</option>`;
    })
    .join("");
  select.value = investmentDates[investmentDates.length - 1];
  select.onchange = () => {
    renderInvestmentSummary();
    renderInvestmentCharts();
  };
}

function initExportDateSelect() {
  const select = document.getElementById("export-date-select");
  if (!exportDates.length) {
    select.innerHTML = "";
    return;
  }

  select.innerHTML = exportDates
    .map((key) => {
      const record = getExportRecord(key);
      return `<option value="${key}">${formatExportPeriod(record?.date)}</option>`;
    })
    .join("");
  select.value = exportDates[exportDates.length - 1];
  select.onchange = () => {
    renderExportSummary();
    renderExportCharts();
  };
}

function initSmeYearSelect() {
  const yearSelect = document.getElementById("sme-year-select");
  if (!smeYears.length) {
    yearSelect.innerHTML = "";
    return;
  }
  yearSelect.innerHTML = smeYears
    .map((year) => `<option value="${year}">${year}</option>`)
    .join("");
  yearSelect.value = smeYears[smeYears.length - 1];
  yearSelect.onchange = () => {
    renderSmeData();
    renderSmeCharts();
  };
}

async function loadSmeData() {
  try {
    smeLoadError = "";
    startupLoadError = "";
    loanLoadError = "";
    investmentLoadError = "";
    exportLoadError = "";
    const [
      smeSheetRows,
      startupSheetRows,
      loanSheetRows,
      delinquencySheetRows,
      investmentSheetRows,
      investmentStageSheetRows,
      investmentSectorSheetRows,
      investmentSourceSheetRows,
      exportSheetRows,
      exportCountrySheetRows,
    ] = await Promise.all([
      loadGoogleSheet("위상"),
      loadGoogleSheet("창업"),
      loadGoogleSheet("", {
        baseUrl: LOAN_SHEET_BASE_URL,
        openSheetBaseUrl: LOAN_OPENSHEET_BASE_URL,
      }),
      loadGoogleSheet("연체율", {
        baseUrl: LOAN_SHEET_BASE_URL,
        openSheetBaseUrl: LOAN_OPENSHEET_BASE_URL,
      }),
      loadGoogleSheet("투자", {
        baseUrl: LOAN_SHEET_BASE_URL,
        openSheetBaseUrl: LOAN_OPENSHEET_BASE_URL,
      }),
      loadGoogleSheet("업력별투자", {
        baseUrl: LOAN_SHEET_BASE_URL,
        openSheetBaseUrl: LOAN_OPENSHEET_BASE_URL,
      }),
      loadGoogleSheet("업종별투자", {
        baseUrl: LOAN_SHEET_BASE_URL,
        openSheetBaseUrl: LOAN_OPENSHEET_BASE_URL,
      }),
      loadGoogleSheet("출자자별", {
        baseUrl: LOAN_SHEET_BASE_URL,
        openSheetBaseUrl: LOAN_OPENSHEET_BASE_URL,
      }),
      loadGoogleSheet("수출", {
        baseUrl: EXPORT_SHEET_BASE_URL,
        openSheetBaseUrl: EXPORT_OPENSHEET_BASE_URL,
      }).catch((error) => {
        exportLoadError = `오류: ${error.message}`;
        return null;
      }),
      loadGoogleSheet("국가별", {
        baseUrl: EXPORT_SHEET_BASE_URL,
        openSheetBaseUrl: EXPORT_OPENSHEET_BASE_URL,
      }).catch(() => null),
    ]);
    const { nextData, nextYears } = parseSmeRows(smeSheetRows);
    startupSeries = parseStartupRows(startupSheetRows);
    loanSeries = parseLoanRows(loanSheetRows);
    delinquencySeries = parseDelinquencyRows(delinquencySheetRows);
    investmentSeries = parseInvestmentRows(investmentSheetRows, [
      "신규 벤처투자금액",
      "피투자기업 수",
      "기업당 투자금액",
      "벤처펀드 결정금액",
      "벤처펀드 결성 수",
    ]);
    investmentStageSeries = parseInvestmentRows(investmentStageSheetRows, [
      "초기 투자(3년 이내)",
      "중기 투자(3~7년 이내)",
      "후기 투자(7년 초과)",
    ]);
    investmentSectorSeries = parseInvestmentRows(investmentSectorSheetRows, [
      "ICT서비스",
      "바이오·의료",
      "전기·기계·장비",
      "ICT제조",
      "유통·서비스",
      "화학·소재",
      "영상·공연·음반",
      "게임",
      "기타",
    ]);
    investmentSourceSeries = parseInvestmentRows(investmentSourceSheetRows, [
      "정책금융",
      "모태펀드",
      "성장금융",
      "산업은행",
      "기타 정책금융",
      "민간부문",
      "개인",
      "일반법인",
      "금융기관(산은 제외)",
      "연기금 및 공제회",
      "VC",
      "기타단체 및 외국인",
    ]);
    exportSeries = exportSheetRows
      ? parseExportRows(exportSheetRows, [
          "전체 기업 수",
          "대기업 기업 수",
          "중견기업 기업 수",
          "중소기업 기업 수",
          "전체 수출금액",
          "대기업 수출금액",
          "중견기업 수출금액",
          "중소기업 수출금액",
        ])
      : [];
    exportCountrySeries = exportCountrySheetRows
      ? parseExportCountryRows(exportCountrySheetRows, [
          "미국 수출금액",
          "중국 수출금액",
          "일본 수출금액",
          "베트남 수출금액",
          "홍콩 수출금액",
          "대만 수출금액",
          "싱가포르 수출금액",
        ])
      : [];
    loanYears = [...new Set([...loanSeries.map((item) => item.year), ...delinquencySeries.map((item) => item.year)])].sort((a, b) => a - b);
    investmentDates = investmentSeries.map((item) => item.key);
    exportDates = exportSeries.map((item) => item.key);
    smeData = nextData;
    smeYears = nextYears;
    initSmeYearSelect();
    initStartupYearSelect();
    initLoanYearSelect();
    initInvestmentDateSelect();
    initExportDateSelect();
    document.getElementById("startup-year-select").onchange = () => {
      renderStartupSummary();
      renderStartupCharts();
    };
    renderSmeData();
    renderSmeCharts();
    renderStartupSummary();
    renderStartupCharts();
    renderLoanSummary();
    renderLoanCharts();
    renderInvestmentSummary();
    renderInvestmentCharts();
    renderExportSummary();
    renderExportCharts();
  } catch (error) {
    smeData = [];
    smeYears = [];
    startupSeries = [];
    loanSeries = [];
    delinquencySeries = [];
    loanYears = [];
    investmentSeries = [];
    investmentStageSeries = [];
    investmentSectorSeries = [];
    investmentSourceSeries = [];
    investmentDates = [];
    exportSeries = [];
    exportCountrySeries = [];
    exportDates = [];
    smeLoadError = `오류: ${error.message}`;
    startupLoadError = `오류: ${error.message}`;
    loanLoadError = `오류: ${error.message}`;
    investmentLoadError = `오류: ${error.message}`;
    exportLoadError = `오류: ${error.message}`;
    initSmeYearSelect();
    initStartupYearSelect();
    initLoanYearSelect();
    initInvestmentDateSelect();
    initExportDateSelect();
    renderSmeData();
    renderSmeCharts();
    renderStartupSummary();
    renderStartupCharts();
    renderLoanSummary();
    renderLoanCharts();
    renderInvestmentSummary();
    renderInvestmentCharts();
    renderExportSummary();
    renderExportCharts();
  }
}

function bindTabs() {
  const buttons = document.querySelectorAll("[data-tab]");
  const panels = document.querySelectorAll("[data-panel]");

  buttons.forEach((button) => {
    button.addEventListener("click", () => {
      buttons.forEach((item) => item.classList.remove("is-active"));
      panels.forEach((panel) => panel.classList.remove("is-active"));
      button.classList.add("is-active");
      document
        .querySelector(`[data-panel="${button.dataset.tab}"]`)
        .classList.add("is-active");
    });
  });
}

async function loadMarketData() {
  const button = document.getElementById("refresh-button");
  const previousSnapshot = readSnapshot();

  button.disabled = true;
  button.textContent = "갱신 중";
  setStatus("시세 데이터를 다시 불러오는 중입니다.");

  try {
    const [kospiText, kosdaqText, marketText, goldPayload, silverPayload, bitcoinPayload] =
      await Promise.all([
        fetchText(NAVER_URLS.kospi),
        fetchText(NAVER_URLS.kosdaq),
        fetchText(NAVER_URLS.market),
        fetchJson(`${GOLD_API_BASE}/XAU`),
        fetchJson(`${GOLD_API_BASE}/XAG`),
        fetchJson(`${GOLD_API_BASE}/BTC`),
      ]);

    const merged = {
      kospi: parseIndexPage(kospiText, "KOSPI"),
      kosdaq: parseIndexPage(kosdaqText, "KOSDAQ"),
      ...parseMarketIndex(marketText),
      gold: parseAltAsset(goldPayload, previousSnapshot),
      silver: parseAltAsset(silverPayload, previousSnapshot),
      bitcoin: parseAltAsset(bitcoinPayload, previousSnapshot),
    };

    renderMarketList(merged);
    writeSnapshot({
      XAU: { price: merged.gold.rawPrice },
      XAG: { price: merged.silver.rawPrice },
      BTC: { price: merged.bitcoin.rawPrice },
    });
    setStatus("실시간 시세 반영 완료");
  } catch (error) {
    renderMarketList({});
    setStatus(`업데이트 실패: ${error.message}`);
  } finally {
    button.disabled = false;
    button.textContent = "새로고침";
  }
}

function bindRefresh() {
  const button = document.getElementById("refresh-button");
  if (button) {
    button.addEventListener("click", loadMarketData);
  }
}

bindTabs();
bindRefresh();
loadSmeData();
