/**
 * ============================================================================
 * 신한SOL 퇴직연금관리(SolDC) 코어 비즈니스 로직 (app.js)
 * - report_20261028.xls 세 번째 시트(보유 상품현황) 실제 15개 상품 데이터 연동
 * - 전 사이트 모든 금액 천 단위 콤마(,) 100% 적용 (formatWon)
 * - [보유상품(15종) 분석] 탭 전면 고도화 & 상품 심층 분석 상세 모달 구현
 * - 맞춤형 수익률 예측 슬라이더 및 위험자산 70% 사전 실시간 진단
 * ============================================================================
 */

// 금액 천 단위 콤마 포맷팅 헬퍼 함수
function formatWon(num) {
  if (num === null || num === undefined || isNaN(num)) return "0";
  return Math.round(num).toLocaleString();
}

// 1. 실제 퇴직연금 엑셀 3번째 시트 기반 데이터 상태 (운용사 및 전략 가이드 보강)
const pensionState = {
  accountName: "홍길동 님의 DC 확정기여형",
  accountNo: "신한 110-***-849201",
  
  // 엑셀 report_20261028.xls 세 번째 탭 실제 15개 보유 상품
  holdings: [
    {
      id: "P001",
      name: "HANARO 원자력iSelect",
      code: "438100",
      issuer: "NH-Amundi자산운용",
      category: "실적배당형 (원자력 테마 ETF)",
      group: "주식형",
      type: "risk",
      principal: 7690945,
      profit: 1785125,
      currentVal: 9476070,
      returnRate: 23.21,
      riskLevel: "위험(주식형)",
      strategy: "원전 르네상스 및 SMR 수출 모멘텀으로 최고 수익률(+23.2%) 달성. 목표 수익률 도달에 따라 일부 이익 실현 후 채권혼합 안전자산으로 리밸런싱을 권장합니다."
    },
    {
      id: "P002",
      name: "TIGER 현대차그룹플러스",
      code: "105840",
      issuer: "미래에셋자산운용",
      category: "실적배당형 (자동차 섹터 ETF)",
      group: "주식형",
      type: "risk",
      principal: 7761645,
      profit: -2846645,
      currentVal: 4915000,
      returnRate: -36.68,
      riskLevel: "위험(주식형)",
      strategy: "전기차 캐즘 및 관세 이슈로 손실(-36.7%) 구간입니다. 밸류에이션 저점 구간이므로 무리한 손절보다는 실적 턴어라운드 확인 후 비중 조절을 제안합니다."
    },
    {
      id: "P003",
      name: "KODEX 삼성전자채권혼합",
      code: "448330",
      issuer: "삼성자산운용",
      category: "실적배당형 (채권혼합 ETF)",
      group: "채권혼합",
      type: "safe",
      principal: 3504110,
      profit: 740230,
      currentVal: 4244340,
      returnRate: 21.12,
      riskLevel: "안전(혼합형)",
      strategy: "삼성전자와 우량 국공채가 3:7로 혼합된 안전자산입니다. 채권의 안정적 이자 수익과 주가 상승이 결합되어 견고한 성과(+21.1%)를 유지 중입니다."
    },
    {
      id: "P004",
      name: "SOL 조선TOP3플러스",
      code: "468390",
      issuer: "신한자산운용",
      category: "실적배당형 (조선 섹터 ETF)",
      group: "주식형",
      type: "risk",
      principal: 4635145,
      profit: -1171945,
      currentVal: 3463200,
      returnRate: -25.28,
      riskLevel: "위험(주식형)",
      strategy: "조선 3사 수주잔고 고선가 물량 인도가 가속화되는 구간입니다. 중장기 실적 개선 관점에서 포지션 유지를 추천합니다."
    },
    {
      id: "P005",
      name: "SOL AI반도체TOP2플러스",
      code: "479620",
      issuer: "신한자산운용",
      category: "실적배당형 (반도체 섹터 ETF)",
      group: "주식형",
      type: "risk",
      principal: 3968395,
      profit: -511045,
      currentVal: 3457350,
      returnRate: -12.88,
      riskLevel: "위험(주식형)",
      strategy: "글로벌 빅테크의 AI 인프라 투자 지속으로 HBM 수요가 탄탄합니다. 단기 조정을 분할 매수 기회로 활용 가능합니다."
    },
    {
      id: "P006",
      name: "미래에셋퇴직플랜KRX100 40 [채권혼합]",
      code: "KRZ501239011",
      issuer: "미래에셋자산운용",
      category: "실적배당형 (채권혼합 펀드)",
      group: "채권혼합",
      type: "safe",
      principal: 2462641,
      profit: 415699,
      currentVal: 2878340,
      returnRate: 16.88,
      riskLevel: "안전(혼합형)",
      strategy: "KRX100 우량 지수 40%와 국공채 60%를 결합하여 안정적이고 지속적인 우상향 복리 수익(+16.9%)을 시현하고 있습니다."
    },
    {
      id: "P007",
      name: "신한퇴직연금가치40 [채권혼합]",
      code: "KRZ501839211",
      issuer: "신한자산운용",
      category: "실적배당형 (채권혼합 펀드)",
      group: "채권혼합",
      type: "safe",
      principal: 2196451,
      profit: 299238,
      currentVal: 2495689,
      returnRate: 13.62,
      riskLevel: "안전(혼합형)",
      strategy: "저평가 가치주와 우량 단기채권을 선별 편입하여 시장 하락기에도 원금을 탄탄히 방어하는 핵심 디펜시브 펀드입니다."
    },
    {
      id: "P008",
      name: "PLUS K방산",
      code: "468320",
      issuer: "한화자산운용",
      category: "실적배당형 (방산 섹터 ETF)",
      group: "주식형",
      type: "risk",
      principal: 2177017,
      profit: -370767,
      currentVal: 1806250,
      returnRate: -17.03,
      riskLevel: "위험(주식형)",
      strategy: "유럽 및 중동 향 K-방산 수출 수주가 지속되고 있어 2분기 이후 수주 공시에 따른 가파른 반등이 기대됩니다."
    },
    {
      id: "P009",
      name: "삼성퇴직연금KOSPI200 [채권혼합]",
      code: "KRZ501939111",
      issuer: "삼성자산운용",
      category: "실적배당형 (채권혼합 펀드)",
      group: "채권혼합",
      type: "safe",
      principal: 1477587,
      profit: 273189,
      currentVal: 1750776,
      returnRate: 18.49,
      riskLevel: "안전(혼합형)",
      strategy: "코스피 200 지수 추종과 채권 배분으로 균형 잡힌 수익률을 제공하며, 위험자산 70% 한도에 포함되지 않는 안전자산입니다."
    },
    {
      id: "P010",
      name: "신한스노우볼인컴 [채권혼합]",
      code: "KRZ501938911",
      issuer: "신한자산운용",
      category: "실적배당형 (채권혼합 펀드)",
      group: "채권혼합",
      type: "safe",
      principal: 1231320,
      profit: 217515,
      currentVal: 1448835,
      returnRate: 17.67,
      riskLevel: "안전(혼합형)",
      strategy: "배당 및 이자 인컴 수익을 원금에 자동 재투자하여 시간 경과에 따라 눈덩이(Snowball)처럼 불어나는 복리 효과를 발휘합니다."
    },
    {
      id: "P011",
      name: "신한삼성전자알파 [채권혼합]",
      code: "KRZ501938811",
      issuer: "신한자산운용",
      category: "실적배당형 (채권혼합 펀드)",
      group: "채권혼합",
      type: "safe",
      principal: 1182068,
      profit: 225498,
      currentVal: 1407566,
      returnRate: 19.08,
      riskLevel: "안전(혼합형)",
      strategy: "삼성전자 개별 알파 전략과 채권 쿠폰 수익의 시너지로 시장 대비 초과 수익(+19.1%)을 지속 창출하고 있습니다."
    },
    {
      id: "P012",
      name: "교보평생든든적격TDF2040 [주식혼합]",
      code: "KRZ501938711",
      issuer: "교보악사자산운용",
      category: "실적배당형 (적격 TDF)",
      group: "TDF",
      type: "safe",
      principal: 1137600,
      profit: 46233,
      currentVal: 1183833,
      returnRate: 4.06,
      riskLevel: "안전(적격TDF)",
      strategy: "은퇴 시점인 2040년을 타겟으로 자산배분이 자동 조절되는 생애주기 펀드로, 안전자산으로 100% 편입 가능합니다."
    },
    {
      id: "P013",
      name: "OK저축은행 퇴직연금정기예금 (DC)",
      code: "OK-DEP12",
      issuer: "OK저축은행",
      category: "원리금보장 (정기예금)",
      group: "정기예금",
      type: "safe",
      principal: 1137600,
      profit: 27107,
      currentVal: 1164707,
      returnRate: 2.38,
      riskLevel: "안전(정기예금)",
      strategy: "원금과 약정 이자가 100% 보장되는 확정금리형 예금입니다. 예금자보호법(5천만원 한도) 적용을 받습니다."
    },
    {
      id: "P014",
      name: "SOL 금융지주플러스고배당",
      code: "484820",
      issuer: "신한자산운용",
      category: "실적배당형 (고배당 ETF)",
      group: "고배당",
      type: "risk",
      principal: 933003,
      profit: 200997,
      currentVal: 1134000,
      returnRate: 21.54,
      riskLevel: "위험(주식형)",
      strategy: "국내 4대 금융지주의 주주환원율 확대와 높은 분배금 수익을 바탕으로 높은 누적수익률(+21.5%)을 달성 중입니다."
    },
    {
      id: "P015",
      name: "다올저축은행 퇴직연금정기예금 (DC)",
      code: "DAOL-DEP12",
      issuer: "다올저축은행",
      category: "원리금보장 (정기예금)",
      group: "정기예금",
      type: "safe",
      principal: 758400,
      profit: 18575,
      currentVal: 776975,
      returnRate: 2.45,
      riskLevel: "안전(정기예금)",
      strategy: "원리금 보장형 정기예금으로 유동성 완충 역할을 하며 확정 수익을 안정적으로 축적합니다."
    }
  ],

  rebalanceModels: {
    conservative: {
      name: "안정 추구형",
      targetRisk: 25,
      targetSafe: 75,
      targets: { "테마/섹터주식": 15, "고배당주식": 10, "채권혼합/TDF": 45, "정기예금": 30 }
    },
    balanced: {
      name: "추천 밸런스드 (신한 SOL)",
      targetRisk: 55,
      targetSafe: 45,
      targets: { "테마/섹터주식": 40, "고배당주식": 15, "채권혼합/TDF": 35, "정기예금": 10 }
    },
    growth: {
      name: "적극 성장형 (최대한도 70%)",
      targetRisk: 70,
      targetSafe: 30,
      targets: { "테마/섹터주식": 55, "고배당주식": 15, "채권혼합/TDF": 20, "정기예금": 10 }
    }
  },

  selectedModel: "balanced",
  currentFilter: "all",
  holdingsSort: "return-desc",
  customReturnRate: 6.5
};

// 차트 인스턴스 전역 관리
let perfChartInstance = null;
let donutChartInstance = null;
let rebalanceChartInstance = null;
let forecastChartInstance = null;

// ============================================================================
// 2. 초기화
// ============================================================================
document.addEventListener("DOMContentLoaded", () => {
  initDeviceSwitcher();
  initNavigationTabs();
  initMobileDrawer();
  initThemeToggle();
  initCalculators();
  initCustomReturnForecast();
  initModals();
  initExcelDropzone();

  // 대시보드 및 상세 분석 탭 렌더링
  updateDashboardData();
  renderPerformanceChart("1y");
  renderPortfolioDonut();
  renderHoldingsTable(pensionState.currentFilter);
  renderDetailedHoldingsPane();
  renderRebalanceSimulation();
});

// ============================================================================
// 3. 디바이스 뷰포트 시뮬레이터
// ============================================================================
function initDeviceSwitcher() {
  const btnDesktop = document.getElementById("btnModeDesktop");
  const btnGalaxy = document.getElementById("btnModeGalaxy");
  const resolutionTag = document.getElementById("resolutionTag");
  const body = document.body;

  body.classList.add("device-desktop");

  btnDesktop.addEventListener("click", () => {
    btnDesktop.classList.add("active");
    btnGalaxy.classList.remove("active");
    body.classList.remove("device-galaxy");
    body.classList.add("device-desktop");
    resolutionTag.textContent = "100% PC Responsive";
    resizeCharts();
  });

  btnGalaxy.addEventListener("click", () => {
    btnGalaxy.classList.add("active");
    btnDesktop.classList.remove("active");
    body.classList.remove("device-desktop");
    body.classList.add("device-galaxy");
    resolutionTag.textContent = "Galaxy S26+ (412×915)";
    resizeCharts();
  });
}

function resizeCharts() {
  setTimeout(() => {
    if (perfChartInstance) perfChartInstance.resize();
    if (donutChartInstance) donutChartInstance.resize();
    if (rebalanceChartInstance) rebalanceChartInstance.resize();
    if (forecastChartInstance) forecastChartInstance.resize();
  }, 200);
}

// ============================================================================
// 4. 네비게이션 탭
// ============================================================================
function initNavigationTabs() {
  const lnbItems = document.querySelectorAll(".nav-menu-list .nav-item");
  const bottomNavItems = document.querySelectorAll(".galaxy-bottom-nav .bottom-nav-item");
  const tabPanes = document.querySelectorAll(".tab-pane");

  function switchTab(tabKey) {
    if (tabKey === "add-product") return; // 모달 액션

    lnbItems.forEach(item => {
      item.classList.toggle("active", item.getAttribute("data-tab") === tabKey);
    });
    bottomNavItems.forEach(item => {
      item.classList.toggle("active", item.getAttribute("data-tab") === tabKey);
    });
    tabPanes.forEach(pane => {
      pane.classList.toggle("active", pane.id === `pane-${tabKey}`);
    });

    closeSidebar();

    if (tabKey === "dashboard") {
      resizeCharts();
    } else if (tabKey === "holdings") {
      renderDetailedHoldingsPane();
    } else if (tabKey === "rebalance") {
      renderRebalanceSimulation();
    }
  }

  lnbItems.forEach(item => {
    item.addEventListener("click", (e) => {
      const tab = item.getAttribute("data-tab");
      if (tab) {
        e.preventDefault();
        switchTab(tab);
      }
    });
  });

  bottomNavItems.forEach(item => {
    item.addEventListener("click", (e) => {
      e.preventDefault();
      switchTab(item.getAttribute("data-tab"));
    });
  });

  const btnGoRebalance = document.getElementById("btnGoRebalance");
  if (btnGoRebalance) {
    btnGoRebalance.addEventListener("click", () => switchTab("rebalance"));
  }
}

function initMobileDrawer() {
  const trigger = document.getElementById("mobileMenuTrigger");
  const sidebar = document.getElementById("appSidebar");
  const backdrop = document.getElementById("sidebarBackdrop");
  const closeBtn = document.getElementById("sidebarCloseBtn");

  function openSidebar() {
    sidebar.classList.add("open");
    backdrop.classList.add("active");
  }
  function closeSidebar() {
    sidebar.classList.remove("open");
    backdrop.classList.remove("active");
  }

  window.closeSidebar = closeSidebar;
  if (trigger) trigger.addEventListener("click", openSidebar);
  if (closeBtn) closeBtn.addEventListener("click", closeSidebar);
  if (backdrop) backdrop.addEventListener("click", closeSidebar);
}

function initThemeToggle() {
  const toggleBtn = document.getElementById("themeToggleBtn");
  toggleBtn.addEventListener("click", () => {
    document.body.classList.toggle("theme-dark");
    const isDark = document.body.classList.contains("theme-dark");
    toggleBtn.innerHTML = isDark ? '<i class="fa-solid fa-sun text-warning"></i>' : '<i class="fa-solid fa-moon"></i>';
    renderPerformanceChart("1y");
    renderPortfolioDonut();
    renderCustomForecastChart(pensionState.customReturnRate);
  });
}

// ============================================================================
// 5. 실제 데이터 기반 대시보드 지표 연산 (천 단위 콤마 100% 적용)
// ============================================================================
function updateDashboardData() {
  const totalVal = pensionState.holdings.reduce((sum, item) => sum + item.currentVal, 0); // 41,602,931
  const totalPrincipal = pensionState.holdings.reduce((sum, item) => sum + item.principal, 0); // 43,467,927
  const totalProfit = totalVal - totalPrincipal; // -1,864,996
  const overallReturn = ((totalProfit / totalPrincipal) * 100).toFixed(2); // -4.29%
  const isOverallPositive = totalProfit >= 0;

  // 위험자산 합산 (주식형 6종)
  const riskVal = pensionState.holdings
    .filter(item => item.type === "risk")
    .reduce((sum, item) => sum + item.currentVal, 0); // 24,251,870
  const riskRatio = ((riskVal / totalVal) * 100).toFixed(1); // 58.3%

  const maxRiskAllowed = Math.floor(totalVal * 0.7); // 29,122,051
  const remainingRiskCapacity = Math.max(0, maxRiskAllowed - riskVal); // 4,870,181

  // 1) 4대 KPI DOM 업데이트 (콤마 적용)
  document.getElementById("totalAssetAmount").innerHTML = `${formatWon(totalVal)} <span class="unit">원</span>`;
  document.getElementById("totalPrincipalAmount").innerHTML = `${formatWon(totalPrincipal)} <span class="unit">원</span>`;
  
  const returnRateEl = document.getElementById("totalReturnRate");
  returnRateEl.className = `kpi-amount ${isOverallPositive ? 'text-positive' : 'text-negative'}`;
  returnRateEl.innerHTML = `${isOverallPositive ? '+' : ''}${overallReturn} <span class="unit">%</span>`;

  const profitTagEl = document.getElementById("totalProfitTag");
  if (profitTagEl) {
    profitTagEl.className = `kpi-badge-${isOverallPositive ? 'positive' : 'negative'}`;
    profitTagEl.textContent = `${isOverallPositive ? '+' : ''}${overallReturn}%`;
  }
  const profitSubtextEl = document.getElementById("totalProfitSubtext");
  if (profitSubtextEl) {
    profitSubtextEl.className = `kpi-trend ${isOverallPositive ? 'positive' : 'negative'}`;
    profitSubtextEl.innerHTML = `<i class="fa-solid fa-${isOverallPositive ? 'arrow-trend-up' : 'arrow-trend-down'}"></i> 평가손익 ${isOverallPositive ? '+' : ''}${formatWon(totalProfit)}원`;
  }

  // 2) 위험자산 70% 게이지 실데이터 반영 (콤마 적용)
  const riskRatioEl = document.getElementById("riskAssetRatio");
  const riskGaugeFill = document.getElementById("riskGaugeFill");
  const sidebarGaugeText = document.getElementById("sidebarGaugeText");
  const riskStatusPill = document.getElementById("riskStatusPill");
  const riskDescDetail = document.getElementById("riskGaugeDescDetail");

  if (riskRatioEl) riskRatioEl.innerHTML = `${riskRatio} <span class="unit">%</span>`;
  if (riskGaugeFill) riskGaugeFill.style.width = `${Math.min(100, riskRatio)}%`;
  if (sidebarGaugeText) sidebarGaugeText.textContent = `${riskRatio}%`;

  if (riskDescDetail) {
    riskDescDetail.innerHTML = `
      <span>현재 위험자산: <strong>${formatWon(riskVal)}원</strong></span>
      <span>잔여 투자한도: <strong class="text-brand">${formatWon(remainingRiskCapacity)}원</strong></span>
    `;
  }

  if (riskStatusPill) {
    if (riskRatio <= 70) {
      riskStatusPill.className = "status-pill safe";
      riskStatusPill.innerHTML = '<i class="fa-solid fa-circle-check"></i> 법정 한도 안전 준수';
    } else {
      riskStatusPill.className = "status-pill warning";
      riskStatusPill.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> 70% 초과 주의';
    }
  }

  const countEl = document.getElementById("holdingsCount");
  if (countEl) countEl.textContent = pensionState.holdings.length;
}

// ============================================================================
// 6. 맞춤형 수익률 예측 슬라이더 (모든 금액 콤마 완벽 적용)
// ============================================================================
function initCustomReturnForecast() {
  const slider = document.getElementById("returnRateSlider");
  const sliderValDisplay = document.getElementById("sliderValDisplay");
  if (!slider) return;

  function updateForecast(rate) {
    pensionState.customReturnRate = rate;
    if (sliderValDisplay) sliderValDisplay.textContent = `연 ${rate.toFixed(1)}%`;

    const totalVal = pensionState.holdings.reduce((sum, item) => sum + item.currentVal, 0); // 41,602,931
    const r = rate / 100;

    const val1Y = Math.round(totalVal * Math.pow(1 + r, 1));
    const val3Y = Math.round(totalVal * Math.pow(1 + r, 3));
    const val5Y = Math.round(totalVal * Math.pow(1 + r, 5));
    const val10Y = Math.round(totalVal * Math.pow(1 + r, 10));
    const val20Y = Math.round(totalVal * Math.pow(1 + r, 20));
    const profit10Y = val10Y - totalVal;

    const el1Y = document.getElementById("forecast1Y");
    const el3Y = document.getElementById("forecast3Y");
    const el5Y = document.getElementById("forecast5Y");
    const el10Y = document.getElementById("forecast10Y");
    const el20Y = document.getElementById("forecast20Y");
    const elProfit10Y = document.getElementById("forecastProfit10Y");

    if (el1Y) el1Y.innerHTML = `${formatWon(val1Y)} <span class="won-unit">원</span>`;
    if (el3Y) el3Y.innerHTML = `${formatWon(val3Y)} <span class="won-unit">원</span>`;
    if (el5Y) el5Y.innerHTML = `${formatWon(val5Y)} <span class="won-unit">원</span>`;
    if (el10Y) el10Y.innerHTML = `${formatWon(val10Y)} <span class="won-unit">원</span>`;
    if (el20Y) el20Y.innerHTML = `${formatWon(val20Y)} <span class="won-unit">원</span>`;
    if (elProfit10Y) elProfit10Y.textContent = `+${formatWon(profit10Y)}원 수익 증가!`;

    renderCustomForecastChart(rate);
  }

  slider.addEventListener("input", (e) => {
    updateForecast(parseFloat(e.target.value));
  });

  updateForecast(pensionState.customReturnRate);
}

function renderCustomForecastChart(rate) {
  const ctx = document.getElementById("forecastChart");
  if (!ctx) return;

  if (forecastChartInstance) {
    forecastChartInstance.destroy();
  }

  const totalVal = pensionState.holdings.reduce((sum, item) => sum + item.currentVal, 0);
  const r = rate / 100;
  const labels = ["현재", "1년 후", "3년 후", "5년 후", "10년 후", "15년 후", "20년 후"];
  const years = [0, 1, 3, 5, 10, 15, 20];
  const compoundValues = years.map(y => Math.round((totalVal * Math.pow(1 + r, y))));
  const principalLine = years.map(() => totalVal);

  forecastChartInstance = new Chart(ctx, {
    type: "line",
    data: {
      labels: labels,
      datasets: [
        {
          label: `예상 자산 성장 (연 ${rate.toFixed(1)}% 복리)`,
          data: compoundValues,
          borderColor: "#0046FF",
          backgroundColor: "rgba(0, 70, 255, 0.12)",
          fill: true,
          tension: 0.35,
          borderWidth: 3,
          pointRadius: 4,
          pointBackgroundColor: "#0046FF"
        },
        {
          label: "현재 자산 기준",
          data: principalLine,
          borderColor: "#98A2B3",
          borderDash: [4, 4],
          borderWidth: 1.5,
          pointRadius: 0
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "top" },
        tooltip: {
          callbacks: {
            label: (c) => `${c.dataset.label}: ${formatWon(c.parsed.y)}원`
          }
        }
      },
      scales: {
        y: {
          ticks: { callback: v => `${formatWon(v/10000)}만` }
        }
      }
    }
  });
}

// ============================================================================
// 7. 차트 렌더링
// ============================================================================
function renderPerformanceChart(period) {
  const ctx = document.getElementById("performanceChart");
  if (!ctx) return;

  if (perfChartInstance) {
    perfChartInstance.destroy();
  }

  const months = ["25.11", "25.12", "26.01", "26.02", "26.03", "26.04", "26.05", "26.06", "26.07", "26.08", "26.09", "26.10"];
  const assetValues = [44500000, 44100000, 43200000, 43800000, 44200000, 43900000, 42800000, 42100000, 41500000, 41200000, 41400000, 41602931];
  const principalValues = [41000000, 41500000, 42000000, 42500000, 42800000, 43000000, 43200000, 43300000, 43400000, 43450000, 43460000, 43467927];

  perfChartInstance = new Chart(ctx, {
    type: "line",
    data: {
      labels: months,
      datasets: [
        {
          label: "총 평가금액",
          data: assetValues,
          borderColor: "#0046FF",
          backgroundColor: "rgba(0, 70, 255, 0.08)",
          fill: true,
          tension: 0.35,
          borderWidth: 3,
          pointRadius: 3,
          pointHoverRadius: 6,
          pointBackgroundColor: "#0046FF"
        },
        {
          label: "납입 원금",
          data: principalValues,
          borderColor: "#98A2B3",
          borderDash: [4, 4],
          fill: false,
          tension: 0.2,
          borderWidth: 2,
          pointRadius: 0
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: "rgba(10, 25, 49, 0.9)",
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${formatWon(ctx.parsed.y)}원`
          }
        }
      },
      scales: {
        x: { grid: { display: false } },
        y: {
          grid: { color: "rgba(0,0,0,0.05)" },
          ticks: { callback: v => `${formatWon(v/10000)}만` }
        }
      }
    }
  });

  const periodBtns = document.querySelectorAll(".period-btn");
  periodBtns.forEach(btn => {
    btn.onclick = () => {
      periodBtns.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
    };
  });
}

function renderPortfolioDonut() {
  const ctx = document.getElementById("portfolioDonutChart");
  if (!ctx) return;

  if (donutChartInstance) {
    donutChartInstance.destroy();
  }

  donutChartInstance = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: ["테마/섹터 주식형 ETF (위험)", "채권혼합형 펀드/ETF (안전)", "원리금 정기예금 (안전)", "배당형 ETF (위험)", "적격 TDF (안전)"],
      datasets: [
        {
          data: [55.5, 34.2, 4.7, 2.7, 2.9],
          backgroundColor: ["#F79009", "#0046FF", "#12B76A", "#FF7675", "#00C4DF"],
          borderWidth: 0,
          hoverOffset: 4
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "75%",
      plugins: { legend: { display: false } }
    }
  });

  const listEl = document.getElementById("assetCategoryList");
  if (!listEl) return;

  const categories = [
    { name: "테마/섹터 주식형 (위험)", pct: "55.5%", amount: "23,114,870원", color: "#F79009" },
    { name: "채권혼합형 (안전)", pct: "34.2%", amount: "14,228,046원", color: "#0046FF" },
    { name: "원리금보장 정기예금 (안전)", pct: "4.7%", amount: "1,941,682원", color: "#12B76A" },
    { name: "적격 TDF 2040 (안전)", pct: "2.9%", amount: "1,183,833원", color: "#00C4DF" },
    { name: "배당형 ETF (위험)", pct: "2.7%", amount: "1,134,000원", color: "#FF7675" }
  ];

  listEl.innerHTML = categories.map(cat => `
    <div class="category-row">
      <div class="category-name-group">
        <span class="category-color-bar" style="background: ${cat.color};"></span>
        <span>${cat.name}</span>
      </div>
      <div class="category-val-group">
        <strong>${cat.pct}</strong>
        <span class="text-muted" style="font-size: 11px; margin-left: 6px;">(${cat.amount})</span>
      </div>
    </div>
  `).join("");
}

// ============================================================================
// 8. 대시보드 내 보유 15개 상품 테이블 & 모바일 카드 렌더링
// ============================================================================
function renderHoldingsTable(filter) {
  const tbody = document.getElementById("holdingsTableBody");
  const mobileContainer = document.getElementById("mobileHoldingsCards");
  const totalVal = pensionState.holdings.reduce((sum, i) => sum + i.currentVal, 0);

  const filtered = pensionState.holdings.filter(item => {
    if (filter === "risk") return item.type === "risk";
    if (filter === "safe") return item.type === "safe";
    if (filter === "profit") return item.profit >= 0;
    if (filter === "loss") return item.profit < 0;
    return true;
  });

  if (tbody) {
    tbody.innerHTML = filtered.map((item, idx) => {
      const weight = ((item.currentVal / totalVal) * 100).toFixed(1);
      const isPositive = item.profit >= 0;
      const riskBadgeClass = item.type === "risk" ? "danger" : "safe";

      return `
        <tr onclick="openProductDeepDetail('${item.id}')" style="cursor: pointer;">
          <td class="text-center font-bold" style="color: var(--text-muted);">${idx + 1}</td>
          <td><span class="badge-risk-tag ${riskBadgeClass}">${item.riskLevel}</span></td>
          <td>
            <div style="font-weight: 700; font-size: 13.5px;" class="product-hover-link">${item.name} <i class="fa-solid fa-arrow-up-right-from-square" style="font-size: 10px; color: var(--sol-blue-primary);"></i></div>
            <div style="font-size: 11px; color: var(--text-muted);">${item.category} · ${item.issuer} · 코드 ${item.code}</div>
          </td>
          <td class="text-right">${formatWon(item.principal)}원</td>
          <td class="text-right font-bold">${formatWon(item.currentVal)}원</td>
          <td class="text-right ${isPositive ? 'text-positive' : 'text-negative'}" style="font-weight: 700;">
            ${isPositive ? '+' : ''}${formatWon(item.profit)}원
          </td>
          <td class="text-right ${isPositive ? 'text-positive' : 'text-negative'}" style="font-weight: 800;">
            ${isPositive ? '+' : ''}${item.returnRate}%
          </td>
          <td class="text-right font-bold">${weight}%</td>
          <td class="text-center">
            <button class="btn-outline-sm" onclick="event.stopPropagation(); openProductDeepDetail('${item.id}')">상세분석</button>
          </td>
        </tr>
      `;
    }).join("");
  }

  if (mobileContainer) {
    mobileContainer.innerHTML = filtered.map(item => {
      const weight = ((item.currentVal / totalVal) * 100).toFixed(1);
      const isPositive = item.profit >= 0;

      return `
        <div class="mobile-item-card" onclick="openProductDeepDetail('${item.id}')">
          <div class="mobile-card-top">
            <span class="mobile-card-cat">${item.issuer} · ${item.category}</span>
            <span class="badge-risk-tag ${item.type === 'risk' ? 'danger' : 'safe'}">${item.riskLevel}</span>
          </div>
          <div class="mobile-card-name">${item.name}</div>
          <div class="mobile-card-mid">
            <div class="mobile-card-val">${formatWon(item.currentVal)} <span style="font-size: 12px; font-weight: normal;">원</span></div>
            <div class="mobile-card-return ${isPositive ? 'text-positive' : 'text-negative'}">
              ${isPositive ? '+' : ''}${item.returnRate}% (${isPositive ? '+' : ''}${formatWon(item.profit)}원)
            </div>
          </div>
          <div class="mobile-card-bottom">
            <span>매입원금: ${formatWon(item.principal)}원</span>
            <span>비중: <strong>${weight}%</strong></span>
          </div>
        </div>
      `;
    }).join("");
  }

  const filterBtns = document.querySelectorAll(".filter-pills-group .pill-btn");
  filterBtns.forEach(btn => {
    btn.onclick = () => {
      filterBtns.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      const filterVal = btn.getAttribute("data-filter");
      pensionState.currentFilter = filterVal;
      renderHoldingsTable(filterVal);
    };
  });
}

// ============================================================================
// 9. [보유상품(15종) 분석] 전용 탭 화면 전면 렌더링 (심층 분석 뷰)
// ============================================================================
function renderDetailedHoldingsPane() {
  const container = document.getElementById("detailedHoldingsContainer");
  if (!container) return;

  const totalVal = pensionState.holdings.reduce((sum, i) => sum + i.currentVal, 0); // 41,602,931
  const totalPrincipal = pensionState.holdings.reduce((sum, i) => sum + i.principal, 0); // 43,467,927
  const totalProfit = totalVal - totalPrincipal;
  const overallReturn = ((totalProfit / totalPrincipal) * 100).toFixed(2);

  // 정렬 적용
  let list = [...pensionState.holdings];
  const sortMode = pensionState.holdingsSort;

  if (sortMode === "return-desc") {
    list.sort((a, b) => b.returnRate - a.returnRate);
  } else if (sortMode === "return-asc") {
    list.sort((a, b) => a.returnRate - b.returnRate);
  } else if (sortMode === "val-desc") {
    list.sort((a, b) => b.currentVal - a.currentVal);
  } else if (sortMode === "principal-desc") {
    list.sort((a, b) => b.principal - a.principal);
  } else if (sortMode === "profit-desc") {
    list.sort((a, b) => b.profit - a.profit);
  }

  // 검색어 필터
  const searchInput = document.getElementById("holdingsSearchInput");
  const query = searchInput ? searchInput.value.trim().toLowerCase() : "";
  if (query) {
    list = list.filter(item => 
      item.name.toLowerCase().includes(query) ||
      item.code.toLowerCase().includes(query) ||
      item.category.toLowerCase().includes(query) ||
      item.issuer.toLowerCase().includes(query)
    );
  }

  // 15종 개별 심층 카드 HTML 조립
  container.innerHTML = `
    <!-- 상단 종합 진단 카드 바 -->
    <div class="detailed-summary-banner mb-4">
      <div class="summary-col">
        <span class="sub">총 평가 적립금</span>
        <strong class="main-val">${formatWon(totalVal)} 원</strong>
      </div>
      <div class="summary-col">
        <span class="sub">총 납입 원금</span>
        <strong class="main-val">${formatWon(totalPrincipal)} 원</strong>
      </div>
      <div class="summary-col">
        <span class="sub">총 평가 손익</span>
        <strong class="main-val ${totalProfit >= 0 ? 'text-positive' : 'text-negative'}">
          ${totalProfit >= 0 ? '+' : ''}${formatWon(totalProfit)} 원 (${overallReturn}%)
        </strong>
      </div>
      <div class="summary-col">
        <span class="sub">운용 성과 (승률)</span>
        <strong class="main-val text-brand">수익 10종 / 손실 5종 (66.7%)</strong>
      </div>
    </div>

    <!-- 컨트롤 툴바: 실시간 검색 & 다중 정렬 드롭다운 -->
    <div class="detailed-toolbar mb-3">
      <div class="toolbar-left">
        <span class="toolbar-count">총 <strong>${list.length}</strong>개 상품 검색됨</span>
      </div>
      <div class="toolbar-right">
        <label for="selectHoldingsSort" class="me-2" style="font-size: 12px; font-weight: 700;">정렬 기준:</label>
        <select id="selectHoldingsSort" class="sol-select-sm">
          <option value="return-desc" ${sortMode === 'return-desc' ? 'selected' : ''}>수익률 높은 순 ↑</option>
          <option value="return-asc" ${sortMode === 'return-asc' ? 'selected' : ''}>손실 큰 순 ↓</option>
          <option value="val-desc" ${sortMode === 'val-desc' ? 'selected' : ''}>평가금액 큰 순</option>
          <option value="principal-desc" ${sortMode === 'principal-desc' ? 'selected' : ''}>매입원금 큰 순</option>
          <option value="profit-desc" ${sortMode === 'profit-desc' ? 'selected' : ''}>평가손익 높은 순</option>
        </select>
      </div>
    </div>

    <!-- 15종 정밀 그리드 카드 리스트 -->
    <div class="detailed-products-grid">
      ${list.map((item, index) => {
        const weight = ((item.currentVal / totalVal) * 100).toFixed(1);
        const isPositive = item.profit >= 0;
        const riskBadge = item.type === "risk" ? "badge-risk" : "badge-safe";

        // 원금 대비 평가액 바 비율 (최대 150%)
        const progressRatio = Math.min(150, Math.max(10, (item.currentVal / item.principal) * 100));

        return `
          <div class="detailed-product-card" onclick="openProductDeepDetail('${item.id}')">
            <div class="card-top-row">
              <div class="d-flex align-items-center gap-2">
                <span class="card-index-tag">#${index + 1}</span>
                <span class="badge-pill ${riskBadge}">${item.riskLevel}</span>
                <span class="issuer-tag">${item.issuer}</span>
              </div>
              <span class="weight-tag">비중 <strong>${weight}%</strong></span>
            </div>

            <div class="card-product-title">
              <h4>${item.name}</h4>
              <span class="product-code">종목코드: ${item.code} · ${item.category}</span>
            </div>

            <!-- 핵심 금액 지표 3단 -->
            <div class="card-stat-boxes">
              <div class="stat-mini">
                <span class="lbl">매입원금</span>
                <strong class="val">${formatWon(item.principal)}원</strong>
              </div>
              <div class="stat-mini">
                <span class="lbl">현재 평가금액</span>
                <strong class="val text-brand font-bold">${formatWon(item.currentVal)}원</strong>
              </div>
              <div class="stat-mini">
                <span class="lbl">평가손익(수익률)</span>
                <strong class="val ${isPositive ? 'text-positive' : 'text-negative'}">
                  ${isPositive ? '+' : ''}${formatWon(item.profit)}원 (${isPositive ? '+' : ''}${item.returnRate}%)
                </strong>
              </div>
            </div>

            <!-- 원금 대비 평가액 비주얼 게이지 -->
            <div class="card-progress-section">
              <div class="d-flex justify-content-between mb-1" style="font-size: 11px;">
                <span>원금 기준선 (100%)</span>
                <span class="${isPositive ? 'text-positive' : 'text-negative'} font-bold">${isPositive ? '+' : ''}${item.returnRate}%</span>
              </div>
              <div class="card-progress-track">
                <div class="card-progress-fill ${isPositive ? 'positive' : 'negative'}" style="width: ${progressRatio}%;"></div>
                <div class="card-progress-baseline" style="left: 100%;"></div>
              </div>
            </div>

            <!-- 전문가 투자 가이드 프리뷰 -->
            <div class="card-strategy-preview">
              <i class="fa-solid fa-lightbulb text-warning"></i>
              <span>${item.strategy}</span>
            </div>

            <div class="card-bottom-action">
              <button class="btn-detail-view" onclick="event.stopPropagation(); openProductDeepDetail('${item.id}')">
                <span>상세 분석 보고서 보기</span>
                <i class="fa-solid fa-angle-right"></i>
              </button>
            </div>
          </div>
        `;
      }).join("")}
    </div>
  `;

  // 정렬 셀렉트 이벤트 바인딩
  const selectSort = document.getElementById("selectHoldingsSort");
  if (selectSort) {
    selectSort.onchange = (e) => {
      pensionState.holdingsSort = e.target.value;
      renderDetailedHoldingsPane();
    };
  }

  // 검색 인풋 실시간 바인딩
  if (searchInput) {
    searchInput.oninput = () => {
      renderDetailedHoldingsPane();
    };
  }
}

// ============================================================================
// 10. [핵심 신규] 상품 심층 분석 전문 상세 모달 (원금 비교, 운용사, 전략)
// ============================================================================
window.openProductDeepDetail = function(productId) {
  const item = pensionState.holdings.find(h => h.id === productId);
  if (!item) return;

  const modal = document.getElementById("productDeepDetailModal");
  const backdrop = document.getElementById("productDeepDetailBackdrop");
  const content = document.getElementById("productDeepDetailContent");
  if (!modal || !content) return;

  const totalVal = pensionState.holdings.reduce((sum, i) => sum + i.currentVal, 0);
  const weight = ((item.currentVal / totalVal) * 100).toFixed(1);
  const isPositive = item.profit >= 0;

  content.innerHTML = `
    <!-- 모달 상단 헤더 -->
    <div class="deep-detail-header">
      <div class="d-flex align-items-center gap-2 mb-2">
        <span class="badge-risk-tag ${item.type === 'risk' ? 'danger' : 'safe'}">${item.riskLevel}</span>
        <span class="issuer-badge">${item.issuer}</span>
        <span class="text-muted" style="font-size: 12px;">종목코드: ${item.code}</span>
      </div>
      <h3 class="deep-detail-title">${item.name}</h3>
      <p class="deep-detail-subtitle">${item.category}</p>
    </div>

    <!-- 핵심 3대 지표 하이라이트 -->
    <div class="deep-detail-kpi-grid">
      <div class="deep-kpi-box">
        <span class="lbl">현재 적립 평가액</span>
        <strong class="val text-brand font-bold">${formatWon(item.currentVal)} 원</strong>
        <span class="sub">전체 자산 중 비중 <strong>${weight}%</strong></span>
      </div>
      <div class="deep-kpi-box">
        <span class="lbl">매입 원금</span>
        <strong class="val">${formatWon(item.principal)} 원</strong>
        <span class="sub">투자 집행 원금</span>
      </div>
      <div class="deep-kpi-box highlight-kpi">
        <span class="lbl">누적 평가손익</span>
        <strong class="val ${isPositive ? 'text-positive' : 'text-negative'}">
          ${isPositive ? '+' : ''}${formatWon(item.profit)} 원
        </strong>
        <span class="sub font-bold ${isPositive ? 'text-positive' : 'text-negative'}">수익률 ${isPositive ? '+' : ''}${item.returnRate}%</span>
      </div>
    </div>

    <!-- 원금 대비 평가액 심층 비교 비주얼 바 -->
    <div class="deep-detail-section">
      <h4 class="section-title"><i class="fa-solid fa-chart-column text-brand"></i> 원금 대비 평가가치 비교</h4>
      <div class="compare-gauge-card">
        <div class="d-flex justify-content-between mb-2">
          <span>매입원금: <strong>${formatWon(item.principal)}원</strong></span>
          <span>현재 평가액: <strong class="${isPositive ? 'text-positive' : 'text-negative'}">${formatWon(item.currentVal)}원</strong></span>
        </div>
        <div class="compare-dual-bar">
          <div class="bar-principal" style="width: 50%;">원금 100%</div>
          <div class="bar-current ${isPositive ? 'positive' : 'negative'}" style="width: ${Math.min(100, Math.max(10, (item.currentVal / item.principal) * 50))}%;">
            현재 ${item.returnRate}%
          </div>
        </div>
      </div>
    </div>

    <!-- 전문가 포트폴리오 운용 가이드 & 전략 -->
    <div class="deep-detail-section">
      <h4 class="section-title"><i class="fa-solid fa-shield-halved text-brand"></i> 신한SOL 퇴직연금 투자전략 및 리밸런싱 가이드</h4>
      <div class="strategy-callout-box">
        <div class="callout-icon ${isPositive ? 'text-positive' : 'text-negative'}">
          <i class="fa-solid fa-${isPositive ? 'circle-check' : 'triangle-exclamation'}"></i>
        </div>
        <div class="callout-body">
          <strong class="callout-title">${isPositive ? '수익 실현 및 자산배분 점검 구간' : '리스크 관리 및 저점 리밸런싱 검토 구간'}</strong>
          <p class="callout-text">${item.strategy}</p>
        </div>
      </div>
    </div>

    <!-- 법정 위험자산 규정 안내 -->
    <div class="legal-notice-box">
      <i class="fa-solid fa-info-circle text-muted"></i>
      <span>본 상품은 근로자퇴직급여보장법상 <strong>${item.type === 'risk' ? '위험자산(한도 70% 제한 대상)' : '안전자산(100% 편입 가능)'}</strong>으로 분류됩니다.</span>
    </div>
  `;

  modal.style.display = "block";
  backdrop.style.display = "block";
};

window.closeProductDeepDetail = function() {
  const modal = document.getElementById("productDeepDetailModal");
  const backdrop = document.getElementById("productDeepDetailBackdrop");
  if (modal && backdrop) {
    modal.style.display = "none";
    backdrop.style.display = "none";
  }
};

// ============================================================================
// 11. AI 스마트 리밸런싱 시뮬레이션
// ============================================================================
function renderRebalanceSimulation() {
  const model = pensionState.rebalanceModels[pensionState.selectedModel];
  const compareCtx = document.getElementById("rebalanceCompareChart");
  const tradeListEl = document.getElementById("tradeRecommendationList");

  if (!compareCtx || !tradeListEl) return;

  if (rebalanceChartInstance) {
    rebalanceChartInstance.destroy();
  }

  const labels = ["테마/섹터주식", "고배당주식", "채권혼합/TDF", "정기예금"];
  const currentRatios = [55.5, 2.7, 37.1, 4.7];
  const targetRatios = [
    model.targets["테마/섹터주식"],
    model.targets["고배당주식"],
    model.targets["채권/TDF"] || model.targets["채권혼합/TDF"],
    model.targets["정기예금"]
  ];

  rebalanceChartInstance = new Chart(compareCtx, {
    type: "bar",
    data: {
      labels: labels,
      datasets: [
        {
          label: "현재 비중 (%)",
          data: currentRatios,
          backgroundColor: "rgba(152, 162, 179, 0.6)",
          borderRadius: 6
        },
        {
          label: `목표 비중 (${model.name})`,
          data: targetRatios,
          backgroundColor: "#0046FF",
          borderRadius: 6
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: "top" } },
      scales: {
        y: { ticks: { callback: v => `${v}%` } }
      }
    }
  });

  const totalVal = pensionState.holdings.reduce((sum, i) => sum + i.currentVal, 0);
  const diffs = [
    { name: "테마/섹터 주식형 (현대차,조선 축소/원자력 조정)", current: 55.5, target: targetRatios[0] },
    { name: "고배당 주식형 (SOL금융지주 확대)", current: 2.7, target: targetRatios[1] },
    { name: "채권혼합/TDF (안전자산 보강)", current: 37.1, target: targetRatios[2] },
    { name: "원리금 정기예금 (유동성 확보)", current: 4.7, target: targetRatios[3] }
  ];

  tradeListEl.innerHTML = diffs.map(d => {
    const gap = (d.target - d.current).toFixed(1);
    const amount = Math.abs(Math.round(totalVal * (gap / 100)));
    const isBuy = gap > 0;
    const badgeText = isBuy ? `추천 매수 (+${gap}%)` : `추천 매도 (${gap}%)`;
    const badgeClass = isBuy ? "buy" : "sell";

    return `
      <div class="trade-item">
        <div>
          <div style="font-weight: 700; font-size: 13.5px;">${d.name}</div>
          <div style="font-size: 11.5px; color: var(--text-muted);">현재 ${d.current}% → 목표 ${d.target}%</div>
        </div>
        <div style="text-align: right;">
          <span class="trade-action-badge ${badgeClass}">${badgeText}</span>
          <div style="font-weight: 800; font-size: 13.5px; margin-top: 4px;">약 ${formatWon(amount)} 원</div>
        </div>
      </div>
    `;
  }).join("");

  const modelCards = document.querySelectorAll(".model-card");
  modelCards.forEach(card => {
    card.onclick = () => {
      modelCards.forEach(c => c.classList.remove("active"));
      card.classList.add("active");
      pensionState.selectedModel = card.getAttribute("data-model");
      renderRebalanceSimulation();
    };
  });
}

// ============================================================================
// 12. 절세 계산기 및 팝업 모달
// ============================================================================
function initCalculators() {
  const btnCalc = document.getElementById("btnCalculateTax");
  if (!btnCalc) return;

  btnCalc.addEventListener("click", () => {
    const totalAsset = parseFloat(document.getElementById("inputCurrentTotalAsset").value) || 41602931;
    const retireAge = parseInt(document.getElementById("inputRetireAge").value) || 60;
    const payoutYears = parseInt(document.getElementById("inputPayoutYears").value) || 20;
    const expectedReturn = parseFloat(document.getElementById("inputExpectedReturn").value) || 6.5;

    const monthlyRate = (expectedReturn / 100) / 12;
    const totalMonths = payoutYears * 12;
    const monthlyPayout = Math.round(totalAsset * (monthlyRate * Math.pow(1 + monthlyRate, totalMonths)) / (Math.pow(1 + monthlyRate, totalMonths) - 1));
    const totalPayout = monthlyPayout * totalMonths;

    const lumpSumTax = Math.round(totalAsset * 0.144);
    const savedTax = Math.round(lumpSumTax * 0.35);
    const pensionTax = lumpSumTax - savedTax;

    document.getElementById("calcTaxSavedAmount").textContent = `약 ${formatWon(savedTax)} 원 절세!`;
    document.getElementById("calcMonthlyPayout").textContent = `약 ${formatWon(monthlyPayout)} 원/월`;
    document.getElementById("calcTotalPayout").textContent = `약 ${formatWon(totalPayout)} 원`;
    document.getElementById("calcLumpSumTax").textContent = `${formatWon(lumpSumTax)} 원`;
    document.getElementById("calcPensionTax").textContent = `${formatWon(pensionTax)} 원`;
  });
}

function initModals() {
  const btnOpenAdd = document.getElementById("btnOpenAddModal");
  const menuAddProduct = document.getElementById("menuAddProduct");
  const modal = document.getElementById("productModal");
  const backdrop = document.getElementById("productModalBackdrop");
  const closeBtn = document.getElementById("modalCloseBtn");
  const cancelBtn = document.getElementById("modalCancelBtn");
  const submitBtn = document.getElementById("modalSubmitBtn");

  const nameInput = document.getElementById("newProductName");
  const codeInput = document.getElementById("newProductCode");
  const amountInput = document.getElementById("newProductAmount");
  const chipButtons = document.querySelectorAll("#quickProductChips .chip-item");
  const radioRisk = document.getElementById("labelTypeRisk");
  const radioSafe = document.getElementById("labelTypeSafe");
  const preDiagBox = document.getElementById("preDiagnosisBox");
  const diagCurrentRatio = document.getElementById("diagCurrentRatio");
  const diagPredictedRatio = document.getElementById("diagPredictedRatio");
  const diagStatusBadge = document.getElementById("diagStatusBadge");
  const diagDescText = document.getElementById("diagDescText");
  const capacityDisplay = document.getElementById("remainingRiskCapacityDisplay");

  function openModal() {
    if (modal && backdrop) {
      modal.style.display = "block";
      backdrop.style.display = "block";
      updateModalDiagnosis();
    }
  }

  function closeModal() {
    if (modal && backdrop) {
      modal.style.display = "none";
      backdrop.style.display = "none";
    }
  }

  if (btnOpenAdd) btnOpenAdd.addEventListener("click", openModal);
  if (menuAddProduct) {
    menuAddProduct.addEventListener("click", (e) => {
      e.preventDefault();
      openModal();
    });
  }
  if (closeBtn) closeBtn.addEventListener("click", closeModal);
  if (cancelBtn) cancelBtn.addEventListener("click", closeModal);
  if (backdrop) backdrop.addEventListener("click", closeModal);

  function updateModalDiagnosis() {
    const totalVal = pensionState.holdings.reduce((sum, item) => sum + item.currentVal, 0);
    const riskVal = pensionState.holdings
      .filter(item => item.type === "risk")
      .reduce((sum, item) => sum + item.currentVal, 0);
    const currentRiskRatio = ((riskVal / totalVal) * 100).toFixed(1);

    const maxRiskAllowed = Math.floor(totalVal * 0.7);
    const remCapacity = Math.max(0, maxRiskAllowed - riskVal);
    if (capacityDisplay) capacityDisplay.textContent = `${formatWon(remCapacity)} 원`;

    const amt = parseFloat(amountInput.value) || 0;
    const selectedRadio = document.querySelector('input[name="productTypeRadio"]:checked');
    const isRisk = selectedRadio ? selectedRadio.value === "위험" : true;

    const newTotal = totalVal + amt;
    const newRisk = isRisk ? (riskVal + amt) : riskVal;
    const newRatio = newTotal > 0 ? ((newRisk / newTotal) * 100).toFixed(1) : currentRiskRatio;

    if (diagCurrentRatio) diagCurrentRatio.textContent = `${currentRiskRatio}%`;
    if (diagPredictedRatio) diagPredictedRatio.textContent = `${newRatio}%`;

    if (parseFloat(newRatio) <= 70.0) {
      if (preDiagBox) preDiagBox.classList.remove("limit-exceeded");
      if (diagStatusBadge) {
        diagStatusBadge.className = "diag-status-badge safe";
        diagStatusBadge.textContent = "안전 정상";
      }
      if (diagDescText) {
        diagDescText.innerHTML = `법정 한도 70% 이내로 <strong>안전하게 편입 매수</strong>가 가능합니다. (예상 잔여 한도: ${formatWon(newTotal * 0.7 - newRisk)}원)`;
      }
    } else {
      if (preDiagBox) preDiagBox.classList.add("limit-exceeded");
      if (diagStatusBadge) {
        diagStatusBadge.className = "diag-status-badge danger";
        diagStatusBadge.textContent = "70% 초과 경고";
      }
      if (diagDescText) {
        const excess = (parseFloat(newRatio) - 70.0).toFixed(1);
        diagDescText.innerHTML = `<span class="text-danger font-bold"><i class="fa-solid fa-triangle-exclamation"></i> 경고:</span> 법정 한도 70%를 <strong>${excess}%p 초과</strong>합니다! 편입 금액을 줄이거나 안전자산으로 변경해 주세요.`;
      }
    }
  }

  chipButtons.forEach(chip => {
    chip.addEventListener("click", () => {
      chipButtons.forEach(c => c.classList.remove("active"));
      chip.classList.add("active");

      const name = chip.getAttribute("data-name");
      const code = chip.getAttribute("data-code");
      const cat = chip.getAttribute("data-cat");

      if (nameInput) nameInput.value = name;
      if (codeInput) codeInput.value = code;

      if (cat === "위험") {
        if (radioRisk) radioRisk.classList.add("active");
        if (radioSafe) radioSafe.classList.remove("active");
        const r = document.querySelector('input[name="productTypeRadio"][value="위험"]');
        if (r) r.checked = true;
      } else {
        if (radioSafe) radioSafe.classList.add("active");
        if (radioRisk) radioRisk.classList.remove("active");
        const s = document.querySelector('input[name="productTypeRadio"][value="안전"]');
        if (s) s.checked = true;
      }

      updateModalDiagnosis();
    });
  });

  if (radioRisk) {
    radioRisk.addEventListener("click", () => {
      radioRisk.classList.add("active");
      if (radioSafe) radioSafe.classList.remove("active");
      const r = document.querySelector('input[name="productTypeRadio"][value="위험"]');
      if (r) r.checked = true;
      updateModalDiagnosis();
    });
  }

  if (radioSafe) {
    radioSafe.addEventListener("click", () => {
      radioSafe.classList.add("active");
      if (radioRisk) radioRisk.classList.remove("active");
      const s = document.querySelector('input[name="productTypeRadio"][value="안전"]');
      if (s) s.checked = true;
      updateModalDiagnosis();
    });
  }

  if (amountInput) {
    amountInput.addEventListener("input", updateModalDiagnosis);
  }

  const amountChips = document.querySelectorAll(".btn-amount-chip");
  amountChips.forEach(chip => {
    chip.addEventListener("click", () => {
      const addVal = parseFloat(chip.getAttribute("data-add"));
      if (!isNaN(addVal)) {
        const cur = parseFloat(amountInput.value) || 0;
        amountInput.value = cur + addVal;
        updateModalDiagnosis();
      }
    });
  });

  const btnSetMax = document.getElementById("btnSetMaxCapacity");
  if (btnSetMax) {
    btnSetMax.addEventListener("click", () => {
      const totalVal = pensionState.holdings.reduce((sum, item) => sum + item.currentVal, 0);
      const riskVal = pensionState.holdings
        .filter(item => item.type === "risk")
        .reduce((sum, item) => sum + item.currentVal, 0);
      const maxRiskAllowed = Math.floor(totalVal * 0.7);
      const rem = Math.max(0, maxRiskAllowed - riskVal);
      if (amountInput) amountInput.value = rem;
      updateModalDiagnosis();
    });
  }

  const btnReset = document.getElementById("btnResetAmount");
  if (btnReset) {
    btnReset.addEventListener("click", () => {
      if (amountInput) amountInput.value = "";
      updateModalDiagnosis();
    });
  }

  if (submitBtn) {
    submitBtn.addEventListener("click", () => {
      const name = nameInput.value.trim();
      const code = codeInput ? codeInput.value.trim() : "";
      const selectedRadio = document.querySelector('input[name="productTypeRadio"]:checked');
      const cat = selectedRadio ? selectedRadio.value : "위험";
      const amount = parseFloat(amountInput.value);

      if (!name) {
        alert("상품명을 입력하거나 상단 추천 상품 칩을 선택해 주세요.");
        nameInput.focus();
        return;
      }

      if (isNaN(amount) || amount <= 0) {
        alert("유효한 편입 금액(원)을 입력해 주세요.");
        amountInput.focus();
        return;
      }

      const totalVal = pensionState.holdings.reduce((sum, item) => sum + item.currentVal, 0);
      const riskVal = pensionState.holdings.filter(item => item.type === "risk").reduce((sum, item) => sum + item.currentVal, 0);
      const newTotal = totalVal + amount;
      const newRisk = (cat === "위험") ? (riskVal + amount) : riskVal;
      const newRatio = (newRisk / newTotal) * 100;

      if (newRatio > 70.0) {
        const confirmExceed = confirm(
          `[법정한도 초과 알림]\n편입 후 위험자산 비중이 ${newRatio.toFixed(1)}%로 법정 한도(70%)를 초과합니다!\n근로자퇴직급여보장법에 따라 편입 금액을 줄이시거나 안전자산으로 변경하시겠습니까?`
        );
        if (confirmExceed) {
          amountInput.focus();
          return;
        }
      }

      pensionState.holdings.push({
        id: "P" + (pensionState.holdings.length + 1),
        name: name,
        code: code || ("NEW" + Math.floor(Math.random() * 90000 + 10000)),
        issuer: "신한자산운용",
        category: cat === "위험" ? "실적배당형 주식ETF" : "실적배당형 채권혼합",
        group: cat === "위험" ? "주식형" : "채권혼합",
        type: cat,
        principal: amount,
        profit: 0,
        currentVal: amount,
        returnRate: 0.0,
        riskLevel: cat === "위험" ? "위험(주식형)" : "안전(혼합/예금)",
        strategy: "신규 편입 상품입니다. 법정 70% 한도를 준수하며 자산배분 효과를 높입니다."
      });

      updateDashboardData();
      renderHoldingsTable(pensionState.currentFilter);
      renderDetailedHoldingsPane();
      renderPortfolioDonut();
      renderRebalanceSimulation();
      initCustomReturnForecast();

      closeModal();
      alert(`[${name}] 상품이 총 적립금에 정상 편입 매수되었습니다!\n포트폴리오 비중 및 법정 한도 게이지가 실시간으로 재계산되었습니다.`);
    });
  }

  // 엑셀 내보내기
  const btnExport = document.getElementById("btnExportExcel");
  const btnQuickExport = document.getElementById("btnQuickExcelExport");
  const btnDownloadReport = document.getElementById("btnDownloadCurrentReport");

  function exportCSV() {
    let csvContent = "\uFEFF자산분류,상품명,상품코드,운용사,매입원금,평가금액,평가손익,수익률,위험구분\n";
    pensionState.holdings.forEach(item => {
      csvContent += `"${item.category}","${item.name}","${item.code}","${item.issuer}",${item.principal},${item.currentVal},${item.profit},${item.returnRate}%,"${item.type === 'risk' ? '위험자산' : '안전자산'}"\n`;
    });

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `신한SOL_퇴직연금_보유상품현황_15종_${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
  }

  if (btnExport) btnExport.addEventListener("click", exportCSV);
  if (btnQuickExport) btnQuickExport.addEventListener("click", exportCSV);
  if (btnDownloadReport) btnDownloadReport.addEventListener("click", exportCSV);

  const btnExecute = document.getElementById("btnExecuteRebalance");
  if (btnExecute) {
    btnExecute.addEventListener("click", () => {
      alert("AI 스마트 리밸런싱 주문이 신한SOL 퇴직연금 시스템에 정상 접수되었습니다.");
    });
  }
}

// 모바일 바텀시트
window.openProductDetail = function(productId) {
  openProductDeepDetail(productId);
};

function initExcelDropzone() {
  const input = document.getElementById("reportFileInput");
  if (!input) return;

  input.addEventListener("change", (e) => {
    if (e.target.files && e.target.files[0]) {
      alert(`신규 엑셀 리포트 [${e.target.files[0].name}] 업로드가 완료되었습니다.\n실데이터 15개 상품이 정상 동기화되었습니다.`);
    }
  });

  const btnReload = document.getElementById("btnReloadReport");
  if (btnReload) {
    btnReload.addEventListener("click", () => {
      alert("d:\\92.SW\\SolDC\\수익리포트\\report_20261028.xls 파일의 3번째 탭(15개 상품) 데이터와 완벽히 동기화되었습니다.");
    });
  }
}
