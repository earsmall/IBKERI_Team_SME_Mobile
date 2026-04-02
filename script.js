const NAVER_URLS = {
  kospi: "https://r.jina.ai/http://https://finance.naver.com/sise/sise_index.naver?code=KOSPI",
  kosdaq: "https://r.jina.ai/http://https://finance.naver.com/sise/sise_index.naver?code=KOSDAQ",
  market: "https://r.jina.ai/http://https://finance.naver.com/marketindex/",
};

const GOLD_API_BASE = "https://api.gold-api.com/price";
const SNAPSHOT_KEY = "market_board_alt_snapshot_v1";

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

const smeData = [
  {
    title: "기업수",
    unit: "개",
    color: "#2c7be5",
    years: {
      "2021": { total: 7723867, sme: 7713895 },
      "2022": { total: 8053163, sme: 8042726 },
      "2023": { total: 8309696, sme: 8298915 },
    },
  },
  {
    title: "종사자수",
    unit: "명",
    color: "#4a9bff",
    years: {
      "2021": { total: 22865491, sme: 18492614 },
      "2022": { total: 23410899, sme: 18956294 },
      "2023": { total: 23767377, sme: 19117649 },
    },
  },
  {
    title: "매출액",
    unit: "백만원",
    color: "#7fb8ff",
    years: {
      "2021": { total: 64500838, sme: 30171248 },
      "2022": { total: 74944317, sme: 33090291 },
      "2023": { total: 73591237, sme: 33012545 },
    },
  },
];

const smeYears = Object.keys(smeData[0].years);

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

  return `${formatNumber(value, 0)}`;
}

function formatTotalFootnote(item, totalValue) {
  if (item.title === "기업수") {
    return `전체 ${formatNumber(totalValue / 10000, 1)}만개`;
  }

  if (item.title === "종사자수") {
    return `전체 ${formatNumber(totalValue, 0)}명`;
  }

  return `전체 ${formatNumber(totalValue, 0)}백만원`;
}

function setStatus(message) {
  document.getElementById("last-updated").textContent = `업데이트 ${formatClock()}`;
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

function initSmeYearSelect() {
  const yearSelect = document.getElementById("sme-year-select");
  yearSelect.innerHTML = smeYears
    .map((year) => `<option value="${year}">${year}</option>`)
    .join("");
  yearSelect.value = smeYears[smeYears.length - 1];
  yearSelect.addEventListener("change", renderSmeData);
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
  document.getElementById("refresh-button").addEventListener("click", loadMarketData);
}

initSmeYearSelect();
renderSmeData();
bindTabs();
bindRefresh();
renderMarketList({});
loadMarketData();
