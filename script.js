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

const notes = [
  {
    title: "환율과 금리가 먼저 흔들리면",
    body: "원화 약세와 국고채 금리 상승은 위험회피 심리와 물가 부담을 함께 시사할 수 있어 다른 지표보다 우선적으로 읽는 편이 좋습니다.",
  },
  {
    title: "원자재와 귀금속을 같이 보는 이유",
    body: "WTI는 인플레이션과 경기 충격을, 금과 은은 안전자산 선호와 실질금리 변화를 보여줘서 같이 놓고 보면 해석이 빨라집니다.",
  },
  {
    title: "비트코인은 위험선호 보조지표",
    body: "전통시장 지표는 아니지만 변동성이 큰 위험자산 심리를 반영하기 때문에 KOSDAQ과 함께 보면 투자심리의 온도차를 읽기 좋습니다.",
  },
];

const sources = [
  {
    label: "네이버 증권 KOSPI 지수 페이지",
    detail: "KOSPI 실시간 지수 및 전일비",
    url: "https://finance.naver.com/sise/sise_index.naver?code=KOSPI",
  },
  {
    label: "네이버 증권 KOSDAQ 지수 페이지",
    detail: "KOSDAQ 실시간 지수 및 전일비",
    url: "https://finance.naver.com/sise/sise_index.naver?code=KOSDAQ",
  },
  {
    label: "네이버 증권 시장지표",
    detail: "달러/원, WTI, 국고채 3년물",
    url: "https://finance.naver.com/marketindex/",
  },
  {
    label: "Gold API",
    detail: "금, 은, 비트코인 실시간 가격 API",
    url: "https://gold-api.com/",
  },
  {
    label: "r.jina.ai",
    detail: "정적 배포 환경에서 네이버 페이지 텍스트를 읽기 위한 미러",
    url: "https://r.jina.ai/",
  },
];

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

function renderNotes() {
  const noteList = document.getElementById("note-list");
  noteList.innerHTML = notes
    .map(
      (note) => `
        <article class="note-card">
          <strong>${note.title}</strong>
          <p>${note.body}</p>
        </article>
      `,
    )
    .join("");
}

function renderSources() {
  const sourceList = document.getElementById("source-list");
  sourceList.innerHTML = sources
    .map(
      (source) => `
        <li>
          <strong>${source.label}</strong><br />
          ${source.detail}<br />
          <a href="${source.url}" target="_blank" rel="noreferrer">원문 보기</a>
        </li>
      `,
    )
    .join("");
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

renderNotes();
renderSources();
bindTabs();
bindRefresh();
renderMarketList({});
loadMarketData();
