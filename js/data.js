// BNI Member Traffic Light Dashboard - Data Layer
// New unified data source: single published CSV with all months + pre-calculated scores
const DATA_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTGGmtbTI0G4VF7xAUT7xU-V2y-QOcsQV-hdjfojpa-mKaRws7DMy5gzIpBDBN7GHdJVO1JZAsWK3uA/pub?gid=0&single=true&output=csv';

// --- CSV Parsing ---
function parseCSV(text) {
  const lines = text.trim().split('\n');
  return lines.map(line => {
    const cells = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQuotes) {
        if (ch === '"' && line[i + 1] === '"') {
          current += '"';
          i++;
        } else if (ch === '"') {
          inQuotes = false;
        } else {
          current += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === ',') {
          cells.push(current.trim());
          current = '';
        } else {
          current += ch;
        }
      }
    }
    cells.push(current.trim());
    return cells;
  });
}

// --- Parse number (handles commas like "40,000" and "#######") ---
function parseNum(val) {
  if (!val || val === '#N/A' || val.startsWith('#')) return 0;
  return parseFloat(val.replace(/,/g, '')) || 0;
}

// --- Parse month format: "24'-10M" → { sort: "2024-10", display: "2024/10月" } ---
function parseMonth(raw) {
  const m = raw.match(/(\d{2})'-(\d{2})M/);
  if (!m) return null;
  const year = 2000 + parseInt(m[1]);
  const month = parseInt(m[2]);
  return {
    sort: `${year}-${String(month).padStart(2, '0')}`,
    display: `${year}/${month}月`,
    year,
    month,
  };
}

// Column indices (fixed, based on the published sheet structure)
// Cols 0-6: key, 狀態, 姓名(ID), 月份, 姓名(display), 姓名, 週
// Cols 7-14: PALMS SCORES (出席, 遲到, 引薦, 來賓, 一對一, 教育, 金額, 得分)
// Cols 15-27: RAW DATA (出席, 缺席, 遲到, 病假, 替代人, 引薦, 收到引薦, 來賓, 一對一, 教育, 金額, 出席率%, 得分)
const COL = {
  MEMBER_ID: 2,    // "001_黃俊凱"
  MONTH: 3,        // "24'-10M"
  DISPLAY_NAME: 4, // "黃俊凱Gask huang"
  // Pre-calculated PALMS scores
  SCORE_ATTENDANCE: 7,
  SCORE_REFERRAL: 9,
  SCORE_GUEST: 10,
  SCORE_ONE_ON_ONE: 11,
  SCORE_TRAINING: 12,
  SCORE_VALUE: 13,
  SCORE_TOTAL: 14,
  // Raw data
  RAW_ATTENDANCE: 15,
  RAW_ABSENCE: 16,
  RAW_LATE: 17,
  RAW_SICK: 18,
  RAW_SUB: 19,
  RAW_REFERRAL: 20,
  RAW_REF_RECEIVED: 21,
  RAW_GUEST: 22,
  RAW_ONE_ON_ONE: 23,
  RAW_TRAINING: 24,
  RAW_VALUE: 25,
  RAW_ATTEND_RATE: 26,
};

// --- Fetch and parse all data ---
let cachedData = null;

async function fetchAllData() {
  if (cachedData) return cachedData;

  const res = await fetch(DATA_URL);
  if (!res.ok) throw new Error('無法載入資料');
  const text = await res.text();
  const rows = parseCSV(text);

  // Only keep data from the last 8 months
  const now = new Date();
  const cutoffDate = new Date(now.getFullYear(), now.getMonth() - 7, 1);
  const cutoffSort = `${cutoffDate.getFullYear()}-${String(cutoffDate.getMonth() + 1).padStart(2, '0')}`;

  // Parse into structured records, skip header rows and empty/#N/A rows
  const records = [];
  const memberSet = new Map(); // id → display name

  for (let i = 2; i < rows.length; i++) {
    const row = rows[i];
    const memberId = (row[COL.MEMBER_ID] || '').trim();
    const monthRaw = (row[COL.MONTH] || '').trim();

    if (!memberId || !memberId.includes('_') || !monthRaw) continue;

    const month = parseMonth(monthRaw);
    if (!month) continue;

    // Skip months older than 8 months ago
    if (month.sort < cutoffSort) continue;

    const idMatch = memberId.match(/^(\d+)_(.+)$/);
    if (!idMatch) continue;

    const id = idMatch[1].padStart(3, '0');
    const name = idMatch[2];
    const displayName = (row[COL.DISPLAY_NAME] || name).trim();

    if (!memberSet.has(id)) {
      memberSet.set(id, { id, name, displayName, display: `${id} ${name}` });
    }

    records.push({
      id,
      month,
      scores: {
        出席: parseNum(row[COL.SCORE_ATTENDANCE]),
        引薦: parseNum(row[COL.SCORE_REFERRAL]),
        來賓: parseNum(row[COL.SCORE_GUEST]),
        一對一: parseNum(row[COL.SCORE_ONE_ON_ONE]),
        教育: parseNum(row[COL.SCORE_TRAINING]),
        金額: parseNum(row[COL.SCORE_VALUE]),
        總分: parseNum(row[COL.SCORE_TOTAL]),
      },
      raw: {
        出席: parseNum(row[COL.RAW_ATTENDANCE]),
        缺席: parseNum(row[COL.RAW_ABSENCE]),
        遲到: parseNum(row[COL.RAW_LATE]),
        病假: parseNum(row[COL.RAW_SICK]),
        替代人: parseNum(row[COL.RAW_SUB]),
        提供引薦: parseNum(row[COL.RAW_REFERRAL]),
        收到引薦: parseNum(row[COL.RAW_REF_RECEIVED]),
        來賓: parseNum(row[COL.RAW_GUEST]),
        一對一會面: parseNum(row[COL.RAW_ONE_ON_ONE]),
        分會教育單位: parseNum(row[COL.RAW_TRAINING]),
        交易價值: parseNum(row[COL.RAW_VALUE]),
        出席率: parseNum(row[COL.RAW_ATTEND_RATE]),
      },
    });
  }

  cachedData = { records, members: memberSet };
  return cachedData;
}

// --- Fetch Member List ---
async function fetchMemberList() {
  const { members } = await fetchAllData();
  return Array.from(members.values()).sort((a, b) => a.id.localeCompare(b.id));
}

// --- Traffic Light ---
function getTrafficLight(total) {
  if (total >= 70) return { color: '#22c55e', label: '綠燈', level: 'green' };
  if (total >= 50) return { color: '#eab308', label: '黃燈', level: 'yellow' };
  if (total >= 30) return { color: '#ef4444', label: '紅燈', level: 'red' };
  return { color: '#374151', label: '黑燈', level: 'black' };
}

// --- Trends (compare last 2 months) ---
function calculateTrends(monthlyData) {
  if (monthlyData.length < 2) return null;
  const prev = monthlyData[monthlyData.length - 2].scores;
  const curr = monthlyData[monthlyData.length - 1].scores;

  const categories = ['出席', '一對一', '引薦', '來賓', '教育', '金額'];
  const result = {};
  categories.forEach(cat => {
    const diff = curr[cat] - prev[cat];
    result[cat] = diff > 0 ? 'up' : diff < 0 ? 'down' : 'stable';
  });
  return result;
}

// --- Action Plan Generator ---
// Priority: 一對一 > 引薦 > 教育 > 來賓 > 金額 > 出席
// Full-score thresholds (6 months ≈ 26 weeks ≈ 6.5 four-week periods):
//   一對一: 週平均 ≥2 → 需 52 次
//   引薦:   週平均 ≥1.5 → 需 39 筆
//   教育:   累計 >4 → 需 5 分
//   來賓:   每4週平均 ≥1.5 → 需 10 位
//   金額:   ≥200 萬
//   出席:   0 次缺席
function generateActionPlan(latestScores, latestRaw) {
  const isGreen = latestScores.總分 >= 70;
  const gap = isGreen ? (100 - latestScores.總分) : (70 - latestScores.總分);
  if (latestScores.總分 >= 100) return { isGreen: true, actions: [], gap: 0 };

  const actions = [];

  // 1. 一對一 (priority 1 - easiest to improve)
  if (latestScores.一對一 < 15) {
    const raw = latestRaw.一對一會面;
    const fullTarget = 52;
    const remaining = Math.max(0, fullTarget - raw);
    actions.push({
      category: '一對一會面',
      priority: 1,
      icon: '🤝',
      current: `目前 6 個月累計 ${raw} 次一對一`,
      fullScore: `滿分需 ${fullTarget} 次，還需 ${remaining} 次`,
      target: latestScores.一對一 < 5 ? '每兩週至少 1 次' : latestScores.一對一 < 10 ? '每週至少 1 次' : '每週至少 2 次',
      actionWeekly: '每週安排至少 2 次一對一會面',
      actionMonthly: '每月安排至少 9 次一對一會面',
      potential: (latestScores.一對一 < 10 ? 10 : 15) - latestScores.一對一,
    });
  }

  // 2. 引薦 (priority 2)
  if (latestScores.引薦 < 20) {
    const raw = latestRaw.提供引薦;
    const fullTarget = 39;
    const remaining = Math.max(0, fullTarget - raw);
    actions.push({
      category: '業務引薦',
      priority: 2,
      icon: '📋',
      current: `目前 6 個月累計提供 ${raw} 筆引薦`,
      fullScore: `滿分需 ${fullTarget} 筆，還需 ${remaining} 筆`,
      target: latestScores.引薦 < 5 ? '每週至少 0.75 筆' : latestScores.引薦 < 10 ? '每週至少 1 筆' : latestScores.引薦 < 15 ? '每週至少 1.2 筆' : '每週至少 1.5 筆',
      actionWeekly: '每週至少提供 1.5 筆引薦',
      actionMonthly: '每月至少提供 7 筆引薦給其他會員',
      potential: Math.min(20, latestScores.引薦 + 5) - latestScores.引薦,
    });
  }

  // 3. 教育培訓 (priority 3)
  if (latestScores.教育 < 15) {
    const raw = latestRaw.分會教育單位;
    const fullTarget = 5;
    const remaining = Math.max(0, fullTarget - raw);
    actions.push({
      category: '教育培訓',
      priority: 3,
      icon: '📚',
      current: `目前 6 個月累計 ${raw} 分教育單位`,
      fullScore: `滿分需 ${fullTarget} 分以上，還需 ${remaining} 分`,
      target: latestScores.教育 < 5 ? '累計 2 分以上' : latestScores.教育 < 10 ? '累計 4 分以上' : '累計 6 分以上',
      actionWeekly: '每週關注分會教育活動與線上課程',
      actionMonthly: '每月至少參加 1 次教育訓練或工作坊',
      potential: (latestScores.教育 < 5 ? 5 : latestScores.教育 < 10 ? 10 : 15) - latestScores.教育,
    });
  }

  // 4. 來賓 (priority 4 - harder)
  if (latestScores.來賓 < 15) {
    const raw = latestRaw.來賓;
    const fullTarget = 10;
    const remaining = Math.max(0, fullTarget - raw);
    actions.push({
      category: '邀請來賓',
      priority: 4,
      icon: '👥',
      current: `目前 6 個月累計邀請 ${raw} 位來賓`,
      fullScore: `滿分需 ${fullTarget} 位，還需 ${remaining} 位`,
      target: latestScores.來賓 < 10 ? '每 4 週至少 1 位' : '每 4 週至少 2 位',
      actionWeekly: '每週邀約至少 1 位潛在來賓',
      actionMonthly: '每月邀請至少 2 位來賓參加例會',
      potential: (latestScores.來賓 < 10 ? 10 : 15) - latestScores.來賓,
    });
  }

  // 5. 金額 (priority 5)
  if (latestScores.金額 < 15) {
    const raw = latestRaw.交易價值;
    const rawWan = (raw / 10000).toFixed(1);
    const fullTargetWan = 200;
    const remainingWan = Math.max(0, fullTargetWan - parseFloat(rawWan));
    actions.push({
      category: '引薦金額',
      priority: 5,
      icon: '💰',
      current: `目前 6 個月累計交易 ${rawWan} 萬`,
      fullScore: `滿分需 ${fullTargetWan} 萬以上，還需 ${remainingWan.toFixed(1)} 萬`,
      target: latestScores.金額 < 5 ? '40 萬以上' : latestScores.金額 < 10 ? '80 萬以上' : '200 萬以上',
      actionWeekly: '每週跟進引薦案件進度，促成成交',
      actionMonthly: '每月促成至少 34 萬交易金額',
      potential: (latestScores.金額 < 5 ? 5 : latestScores.金額 < 10 ? 10 : 15) - latestScores.金額,
    });
  }

  // 6. 出席 (priority 6)
  if (latestScores.出席 < 20) {
    const raw = latestRaw.缺席;
    actions.push({
      category: '出席',
      priority: 6,
      icon: '✅',
      current: `目前 6 個月累計缺席 ${raw} 次`,
      fullScore: `滿分需 0 次缺席，還需減少 ${raw} 次`,
      target: '0 次缺席',
      actionWeekly: '每週確認出席，無法出席提前安排替代人',
      actionMonthly: '每月維持全勤出席',
      potential: 20 - latestScores.出席,
    });
  }

  return { isGreen, actions, gap };
}

// --- Main Data Fetch Orchestrator ---
async function getMemberDashboardData(memberId) {
  const paddedId = memberId.padStart(3, '0');

  const { records, members } = await fetchAllData();
  const member = members.get(paddedId);
  if (!member) throw new Error(`找不到會員編號 ${paddedId}，請確認後重試`);

  // Filter records for this member and sort by month
  const memberRecords = records
    .filter(r => r.id === paddedId)
    .sort((a, b) => a.month.sort.localeCompare(b.month.sort));

  if (memberRecords.length === 0) throw new Error('此會員尚無月度資料');

  // Use latest 6 months for the dashboard charts
  const recentRecords = memberRecords.slice(-6);
  const latest = recentRecords[recentRecords.length - 1];

  // Build scores from the latest month's pre-calculated values
  const scoreItems = [
    { name: '出席', score: latest.scores.出席, max: 20 },
    { name: '一對一', score: latest.scores.一對一, max: 15 },
    { name: '教育', score: latest.scores.教育, max: 15 },
    { name: '引薦', score: latest.scores.引薦, max: 20 },
    { name: '來賓', score: latest.scores.來賓, max: 15 },
    { name: '金額', score: latest.scores.金額, max: 15 },
  ];

  const scores = {
    items: scoreItems,
    total: latest.scores.總分,
    light: getTrafficLight(latest.scores.總分),
  };

  const trends = calculateTrends(recentRecords);
  const actionPlan = generateActionPlan(latest.scores, latest.raw);

  return {
    member,
    monthlyData: recentRecords,
    allMonthlyData: memberRecords,
    scores,
    trends,
    actionPlan,
    monthCount: recentRecords.length,
    totalMonths: memberRecords.length,
  };
}

// Export for use in admin.js
window.BNIData = {
  fetchMemberList,
  getMemberDashboardData,
  fetchAllData,
  getTrafficLight,
  calculateTrends,
};
