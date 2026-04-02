const NAVER_URLS = {
  kospi: "https://r.jina.ai/http://https://finance.naver.com/sise/sise_index.naver?code=KOSPI",
  kosdaq: "https://r.jina.ai/http://https://finance.naver.com/sise/sise_index.naver?code=KOSDAQ",
  market: "https://r.jina.ai/http://https://finance.naver.com/marketindex/",
};

const GOLD_API_BASE = "https://api.gold-api.com/price";
const SNAPSHOT_KEY = "market_board_alt_snapshot_v1";
const SHEET_BASE_URL =
  "https://docs.google.com/spreadsheets/d/1R2BB-k4L6a6QwiD5nmFz-idJoIUGnhQCDzHF0ijwH4c/gviz/tq";

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
let smeLoadError = "";
let startupLoadError = "";

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

function loadGoogleSheet(sheetName) {
  return new Promise((resolve, reject) => {
    const callbackName = `__sheetCallback_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
    const script = document.createElement("script");
    const tqx = encodeURIComponent(`out:json;responseHandler:${callbackName}`);
    const sheet = encodeURIComponent(sheetName);

    window[callbackName] = (payload) => {
      cleanup();
      try {
        resolve(mapGvizPayload(payload));
      } catch (error) {
        reject(error);
      }
    };

    script.src = `${SHEET_BASE_URL}?tqx=${tqx}&sheet=${sheet}`;
    script.async = true;

    script.onerror = () => {
      cleanup();
      reject(new Error(`${sheetName} 시트 로드 실패`));
    };

    function cleanup() {
      delete window[callbackName];
      script.remove();
    }

    document.head.appendChild(script);
  });
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
    const total = Number(row["전체기업"]);
    const sme = Number(row["중소기업"]);

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
    const startupCount = Number(row["전체 창업기업 수"]);
    const techStartupCount = Number(row["기술기반업종 창업기업 수"]);

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
  const smeGrid = document.getElementById("sme-grid");

  smeGrid.innerHTML = smeData
    .map((item) => {
      const currentValue = item.years[selectedYear].sme;
      const totalValue = item.years[selectedYear].total;
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
              <div class="sme-value">${formatSmeValue(item, currentValue)}</div>
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
            <div class="startup-chart-subtitle">선택한 기준년도를 포함한 최근 3개년</div>
          </div>
          <div class="startup-bars" style="--bar-count:${filteredYears.length};">
            ${filteredYears
              .map((year) => {
                const value = item.years[year]?.sme ?? 0;
                const barHeight = Math.max(8, Math.round((value / maxValue) * 150));
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

function getFilteredStartupSeries() {
  const selectedYear =
    document.getElementById("startup-year-select")?.value || startupSeries[startupSeries.length - 1]?.year;
  const selectedIndex = startupSeries.findIndex((item) => item.year === selectedYear);
  if (selectedIndex < 0) {
    return startupSeries;
  }

  const startIndex = Math.max(0, selectedIndex - 4);
  return startupSeries.slice(startIndex, selectedIndex + 1);
}

function renderStartupBarChart({ title, subtitle, key, type, colorClass, colorValue }) {
  const filteredSeries = getFilteredStartupSeries();

  if (!filteredSeries.length) {
    return "";
  }

  const values = filteredSeries.map((item) => item[key] ?? 0);
  const maxValue = Math.max(...values, 1);

  return `
    <article class="startup-chart-card">
      <div class="startup-chart-head">
        <div class="startup-chart-title">${title}</div>
        <div class="startup-chart-subtitle">${subtitle}</div>
      </div>
      <div class="startup-bars" style="--bar-count:${filteredSeries.length};">
        ${filteredSeries
          .map((item) => {
            const value = item[key] ?? 0;
            const barHeight = Math.max(8, Math.round((value / maxValue) * 150));
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

function renderStartupLineChart({ title, subtitle, key }) {
  const filteredSeries = getFilteredStartupSeries();

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
        <div class="startup-chart-subtitle">${subtitle}</div>
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
      subtitle: "선택한 기준년도를 포함한 최근 5개년",
      key: "startupCount",
      type: "count",
      colorClass: "is-blue",
      colorValue: "#2c7be5",
    }),
    renderStartupBarChart({
      title: "기술기반업종 창업기업 수",
      subtitle: "선택한 기준년도를 포함한 최근 5개년",
      key: "techStartupCount",
      type: "count",
      colorClass: "is-sky",
      colorValue: "#59a7ff",
    }),
    renderStartupLineChart({
      title: "기술기반업종 비중",
      subtitle: "선택한 기준년도를 포함한 최근 5개년",
      key: "techShare",
    }),
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
    const [smeSheetRows, startupSheetRows] = await Promise.all([
      loadGoogleSheet("위상"),
      loadGoogleSheet("창업"),
    ]);
    const { nextData, nextYears } = parseSmeRows(smeSheetRows);
    startupSeries = parseStartupRows(startupSheetRows);
    smeData = nextData;
    smeYears = nextYears;
    initSmeYearSelect();
    initStartupYearSelect();
    document.getElementById("startup-year-select").onchange = () => {
      renderStartupSummary();
      renderStartupCharts();
    };
    renderSmeData();
    renderSmeCharts();
    renderStartupSummary();
    renderStartupCharts();
  } catch (error) {
    smeData = [];
    smeYears = [];
    startupSeries = [];
    smeLoadError = `오류: ${error.message}`;
    startupLoadError = `오류: ${error.message}`;
    initSmeYearSelect();
    initStartupYearSelect();
    renderSmeData();
    renderSmeCharts();
    renderStartupSummary();
    renderStartupCharts();
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
