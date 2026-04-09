// BNI Admin Dashboard - Leadership Team
(function () {
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  let teamData = null;
  let currentSort = { key: 'total', dir: 'desc' };
  let currentFilter = 'all';
  let currentSearch = '';
  let charts = {};

  // =====================
  // DATA PROCESSING
  // =====================

  async function buildTeamData() {
    const { records, members } = await BNIData.fetchAllData();
    const memberList = await BNIData.fetchMemberList();

    // Group records by member
    const memberRecords = {};
    records.forEach(r => {
      if (!memberRecords[r.id]) memberRecords[r.id] = [];
      memberRecords[r.id].push(r);
    });

    // Collect all months
    const allMonthsSet = new Set();
    records.forEach(r => allMonthsSet.add(r.month.sort));
    const allMonths = [...allMonthsSet].sort();
    const recentMonths = allMonths.slice(-6);

    // Build per-member data
    const membersData = [];
    for (const m of memberList) {
      const recs = memberRecords[m.id];
      if (!recs || recs.length === 0) continue;

      recs.sort((a, b) => a.month.sort.localeCompare(b.month.sort));
      const recent = recs.filter(r => recentMonths.includes(r.month.sort));
      if (recent.length === 0) continue;

      const latest = recent[recent.length - 1];
      const prev = recent.length >= 2 ? recent[recent.length - 2] : null;
      const twoMonthsAgo = recent.length >= 3 ? recent[recent.length - 3] : null;

      const latestScores = {
        total: latest.scores.總分,
        attendance: latest.scores.出席,
        oneOnOne: latest.scores.一對一,
        training: latest.scores.教育,
        referral: latest.scores.引薦,
        guest: latest.scores.來賓,
        value: latest.scores.金額,
      };

      const light = BNIData.getTrafficLight(latestScores.total);
      const trendVal = prev ? latestScores.total - prev.scores.總分 : 0;

      // Category changes for alerts
      const categoryChanges = {};
      if (twoMonthsAgo) {
        const cats = ['出席', '一對一', '教育', '引薦', '來賓', '金額'];
        cats.forEach(c => {
          categoryChanges[c] = latest.scores[c] - twoMonthsAgo.scores[c];
        });
      }

      // Consecutive low months
      let lowStreak = 0;
      for (let i = recent.length - 1; i >= 0; i--) {
        if (recent[i].scores.總分 < 50) lowStreak++;
        else break;
      }

      // Find weakest category
      const catEntries = [
        { name: '出席', score: latestScores.attendance, max: 20 },
        { name: '一對一', score: latestScores.oneOnOne, max: 15 },
        { name: '教育', score: latestScores.training, max: 15 },
        { name: '引薦', score: latestScores.referral, max: 20 },
        { name: '來賓', score: latestScores.guest, max: 15 },
        { name: '金額', score: latestScores.value, max: 15 },
      ];
      const weakest = catEntries.reduce((w, c) =>
        (c.score / c.max) < (w.score / w.max) ? c : w
      );

      membersData.push({
        id: m.id,
        name: m.name,
        display: m.display,
        scores: latestScores,
        light,
        trend: trendVal,
        twoMonthDelta: twoMonthsAgo ? latestScores.total - twoMonthsAgo.scores.總分 : null,
        categoryChanges,
        lowStreak,
        weakest,
        monthlyData: recent,
        monthCount: recent.length,
      });
    }

    // Per-month team aggregates
    const monthAggregates = recentMonths.map(monthSort => {
      const monthMembers = [];
      for (const md of membersData) {
        const rec = md.monthlyData.find(r => r.month.sort === monthSort);
        if (rec) monthMembers.push(rec);
      }
      const totals = monthMembers.map(r => r.scores.總分);
      const avg = totals.length ? totals.reduce((a, b) => a + b, 0) / totals.length : 0;

      let green = 0, yellow = 0, red = 0, black = 0;
      totals.forEach(t => {
        if (t >= 70) green++;
        else if (t >= 50) yellow++;
        else if (t >= 30) red++;
        else black++;
      });

      // Per-category averages
      const catAvgs = {};
      ['出席', '一對一', '教育', '引薦', '來賓', '金額'].forEach(cat => {
        const vals = monthMembers.map(r => r.scores[cat]);
        catAvgs[cat] = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
      });

      const display = monthSort.replace(/^(\d{4})-0?(\d+)$/, '$1/$2月');
      return { monthSort, display, avg, green, yellow, red, black, total: totals.length, catAvgs };
    });

    return { membersData, monthAggregates, recentMonths, allMonths };
  }

  function computeAlerts(membersData) {
    const declining = membersData
      .filter(m => m.twoMonthDelta !== null && m.twoMonthDelta <= -10)
      .sort((a, b) => a.twoMonthDelta - b.twoMonthDelta);

    const persistentLow = membersData
      .filter(m => m.lowStreak >= 3)
      .sort((a, b) => b.lowStreak - a.lowStreak);

    const improving = membersData
      .filter(m => m.twoMonthDelta !== null && m.twoMonthDelta >= 10)
      .sort((a, b) => b.twoMonthDelta - a.twoMonthDelta);

    return { declining, persistentLow, improving };
  }

  // =====================
  // RENDER FUNCTIONS
  // =====================

  function renderKPI(membersData, monthAggregates) {
    const container = $('#kpi-row');
    const totals = membersData.map(m => m.scores.total);
    const avg = totals.length ? (totals.reduce((a, b) => a + b, 0) / totals.length).toFixed(1) : 0;
    const avgLight = BNIData.getTrafficLight(Math.round(avg));
    const greenCount = membersData.filter(m => m.scores.total >= 70).length;
    const greenRate = totals.length ? Math.round(greenCount / totals.length * 100) : 0;
    const atRisk = membersData.filter(m => m.scores.total < 50).length;

    // Month-over-month percentage change
    const lastTwo = monthAggregates.slice(-2);
    let monthDeltaPct = 0;
    if (lastTwo.length === 2 && lastTwo[0].avg > 0) {
      monthDeltaPct = ((lastTwo[1].avg - lastTwo[0].avg) / lastTwo[0].avg * 100).toFixed(1);
    }
    const deltaSign = monthDeltaPct > 0 ? '+' : '';
    const deltaClass = monthDeltaPct > 0 ? 'trend-up' : monthDeltaPct < 0 ? 'trend-down' : 'trend-stable';
    const deltaWord = monthDeltaPct > 0 ? '提升' : monthDeltaPct < 0 ? '降低' : '持平';

    // Attendance rate (latest month)
    const latestAgg = monthAggregates[monthAggregates.length - 1];
    const fullAttendance = membersData.filter(m => m.scores.attendance >= 20).length;
    const attendRate = totals.length ? Math.round(fullAttendance / totals.length * 100) : 0;

    container.innerHTML = `
      <div class="kpi-card">
        <div class="kpi-label">團隊平均分</div>
        <div class="kpi-value" style="color:${avgLight.color}">${avg}</div>
        <div class="kpi-sub">${avgLight.label}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">綠燈率</div>
        <div class="kpi-value" style="color:var(--green)">${greenRate}%</div>
        <div class="kpi-sub">${greenCount} / ${totals.length} 人</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">警戒人數</div>
        <div class="kpi-value" style="color:var(--red)">${atRisk}</div>
        <div class="kpi-sub">總分低於 50 分</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">月度變化</div>
        <div class="kpi-value ${deltaClass}">${deltaSign}${monthDeltaPct}%</div>
        <div class="kpi-sub">${deltaWord} vs 上月平均</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">全勤率</div>
        <div class="kpi-value" style="color:var(--green)">${attendRate}%</div>
        <div class="kpi-sub">${fullAttendance} / ${totals.length} 人全勤</div>
      </div>
    `;
  }

  function renderLightDistribution(membersData) {
    const container = $('#light-distribution');
    const legend = $('#light-legend');
    const total = membersData.length;
    const counts = { green: 0, yellow: 0, red: 0, black: 0 };
    membersData.forEach(m => {
      if (m.scores.total >= 70) counts.green++;
      else if (m.scores.total >= 50) counts.yellow++;
      else if (m.scores.total >= 30) counts.red++;
      else counts.black++;
    });

    const segments = [
      { key: 'green', color: 'var(--green)', label: '綠燈' },
      { key: 'yellow', color: 'var(--yellow)', label: '黃燈' },
      { key: 'red', color: 'var(--red)', label: '紅燈' },
      { key: 'black', color: 'var(--black-light)', label: '黑燈' },
    ];

    container.innerHTML = segments
      .filter(s => counts[s.key] > 0)
      .map(s => {
        const pct = (counts[s.key] / total * 100).toFixed(1);
        return `<div class="light-segment" style="width:${pct}%;background:${s.color}">${counts[s.key]}</div>`;
      }).join('');

    legend.innerHTML = segments.map(s =>
      `<span class="legend-item"><span class="legend-dot" style="background:${s.color}"></span>${s.label} ${counts[s.key]} 人</span>`
    ).join('');
  }

  // --- Table ---
  function getScoreClass(score, max) {
    if (score >= max) return 'score-full';
    if (score === 0) return 'score-zero';
    const ratio = score / max;
    if (ratio >= 0.7) return 'score-high';
    if (ratio >= 0.4) return 'score-mid';
    return 'score-low';
  }

  function getBarColor(score, max) {
    if (score >= max) return 'var(--green)';
    if (score === 0) return '#d1d5db';
    const ratio = score / max;
    if (ratio >= 0.7) return '#059669';
    if (ratio >= 0.4) return 'var(--yellow)';
    return 'var(--red)';
  }

  function renderTable(membersData) {
    const tbody = $('#table-body');

    // Filter
    let filtered = membersData;
    if (currentFilter !== 'all') {
      const filterMap = { green: 70, yellow: 50, red: 30, black: 0 };
      filtered = membersData.filter(m => {
        if (currentFilter === 'green') return m.scores.total >= 70;
        if (currentFilter === 'yellow') return m.scores.total >= 50 && m.scores.total < 70;
        if (currentFilter === 'red') return m.scores.total >= 30 && m.scores.total < 50;
        if (currentFilter === 'black') return m.scores.total < 30;
      });
    }
    if (currentSearch) {
      filtered = filtered.filter(m =>
        m.name.includes(currentSearch) || m.id.includes(currentSearch)
      );
    }

    // Sort
    const sorted = [...filtered].sort((a, b) => {
      let va, vb;
      switch (currentSort.key) {
        case 'id': va = a.id; vb = b.id; break;
        case 'name': va = a.name; vb = b.name; break;
        case 'total': va = a.scores.total; vb = b.scores.total; break;
        case 'attendance': va = a.scores.attendance; vb = b.scores.attendance; break;
        case 'oneOnOne': va = a.scores.oneOnOne; vb = b.scores.oneOnOne; break;
        case 'training': va = a.scores.training; vb = b.scores.training; break;
        case 'referral': va = a.scores.referral; vb = b.scores.referral; break;
        case 'guest': va = a.scores.guest; vb = b.scores.guest; break;
        case 'value': va = a.scores.value; vb = b.scores.value; break;
        case 'trend': va = a.trend; vb = b.trend; break;
        default: va = a.scores.total; vb = b.scores.total;
      }
      if (typeof va === 'string') {
        return currentSort.dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
      }
      return currentSort.dir === 'asc' ? va - vb : vb - va;
    });

    function miniBar(score, max) {
      const pct = Math.round(score / max * 100);
      const color = getBarColor(score, max);
      const cls = getScoreClass(score, max);
      return `<div class="mini-bar">
        <span class="mini-bar-num ${cls}">${score}</span>
        <div class="mini-bar-track"><div class="mini-bar-fill" style="width:${pct}%;background:${color}"></div></div>
      </div>`;
    }

    tbody.innerHTML = sorted.map(m => {
      const trendStr = m.trend > 0 ? `<span class="trend-up">▲+${m.trend}</span>`
        : m.trend < 0 ? `<span class="trend-down">▼${m.trend}</span>`
        : `<span class="trend-stable">—</span>`;

      return `<tr data-id="${m.id}">
        <td>${m.id}</td>
        <td><strong>${m.name}</strong></td>
        <td><strong style="color:${m.light.color}">${m.scores.total}</strong></td>
        <td><span class="light-dot" style="background:${m.light.color}"></span></td>
        <td>${miniBar(m.scores.attendance, 20)}</td>
        <td>${miniBar(m.scores.oneOnOne, 15)}</td>
        <td>${miniBar(m.scores.training, 15)}</td>
        <td>${miniBar(m.scores.referral, 20)}</td>
        <td>${miniBar(m.scores.guest, 15)}</td>
        <td>${miniBar(m.scores.value, 15)}</td>
        <td>${trendStr}</td>
      </tr>`;
    }).join('');
  }

  // --- Alerts ---
  function renderAlerts(alerts) {
    // Declining
    const declEl = $('#alert-declining');
    if (alerts.declining.length === 0) {
      declEl.innerHTML = '<div class="alert-empty">目前沒有績效下滑的會員</div>';
    } else {
      declEl.innerHTML = alerts.declining.map(m => {
        const prevScore = m.scores.total - m.twoMonthDelta;
        const declined = Object.entries(m.categoryChanges)
          .filter(([, v]) => v < 0)
          .map(([k, v]) => `${k}${v}`)
          .join('、');
        return `<div class="alert-card alert-card-red">
          <div class="alert-info">
            <div class="alert-name">${m.id} ${m.name}</div>
            <div class="alert-detail">${prevScore} 分 → ${m.scores.total} 分${declined ? '（' + declined + '）' : ''}</div>
            <span class="alert-action alert-action-red">建議安排一對一關懷</span>
          </div>
          <div class="alert-delta" style="color:var(--red)">${m.twoMonthDelta}</div>
        </div>`;
      }).join('');
    }

    // Persistent Low
    const lowEl = $('#alert-low');
    if (alerts.persistentLow.length === 0) {
      lowEl.innerHTML = '<div class="alert-empty">目前沒有持續低分的會員</div>';
    } else {
      lowEl.innerHTML = alerts.persistentLow.map(m =>
        `<div class="alert-card alert-card-black">
          <div class="alert-info">
            <div class="alert-name">${m.id} ${m.name}</div>
            <div class="alert-detail">目前 ${m.scores.total} 分，已連續 ${m.lowStreak} 個月低於 50 分，最弱項目：${m.weakest.name}</div>
            <span class="alert-action alert-action-black">建議主動提供協助與資源</span>
          </div>
          <div class="alert-delta" style="color:var(--black-light)">${m.scores.total}</div>
        </div>`
      ).join('');
    }

    // Improving
    const impEl = $('#alert-improving');
    if (alerts.improving.length === 0) {
      impEl.innerHTML = '<div class="alert-empty">目前沒有大幅進步的會員</div>';
    } else {
      impEl.innerHTML = alerts.improving.map(m => {
        const prevScore = m.scores.total - m.twoMonthDelta;
        return `<div class="alert-card alert-card-green">
          <div class="alert-info">
            <div class="alert-name">${m.id} ${m.name}</div>
            <div class="alert-detail">${prevScore} 分 → ${m.scores.total} 分</div>
            <span class="alert-action alert-action-green">建議會中公開表揚</span>
          </div>
          <div class="alert-delta" style="color:var(--green)">+${m.twoMonthDelta}</div>
        </div>`;
      }).join('');
    }
  }

  // --- Charts ---
  function renderCharts(monthAggregates, membersData) {
    // Destroy old charts
    Object.values(charts).forEach(c => c.destroy());
    charts = {};

    const labels = monthAggregates.map(m => m.display);

    // 4a: Team Average Trend
    charts.avg = new Chart($('#chart-team-avg'), {
      type: 'line',
      data: {
        labels,
        datasets: [{
          label: '團隊平均分',
          data: monthAggregates.map(m => m.avg.toFixed(1)),
          borderColor: '#3b82f6',
          backgroundColor: 'rgba(59,130,246,0.1)',
          fill: true,
          tension: 0.3,
          pointRadius: 5,
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          annotation: false,
          legend: { display: false },
        },
        scales: {
          y: {
            min: 0, max: 100,
            ticks: { stepSize: 10 },
          }
        }
      },
      plugins: [{
        id: 'greenLine',
        afterDraw(chart) {
          const yScale = chart.scales.y;
          const y = yScale.getPixelForValue(70);
          const ctx = chart.ctx;
          ctx.save();
          ctx.strokeStyle = '#22c55e';
          ctx.lineWidth = 2;
          ctx.setLineDash([6, 4]);
          ctx.beginPath();
          ctx.moveTo(chart.chartArea.left, y);
          ctx.lineTo(chart.chartArea.right, y);
          ctx.stroke();
          ctx.fillStyle = '#22c55e';
          ctx.font = '12px sans-serif';
          ctx.fillText('綠燈 70', chart.chartArea.right - 50, y - 6);
          ctx.restore();
        }
      }]
    });

    // 4b: Light Distribution Trend
    charts.lightTrend = new Chart($('#chart-light-trend'), {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: '綠燈', data: monthAggregates.map(m => m.green), backgroundColor: '#22c55e' },
          { label: '黃燈', data: monthAggregates.map(m => m.yellow), backgroundColor: '#eab308' },
          { label: '紅燈', data: monthAggregates.map(m => m.red), backgroundColor: '#ef4444' },
          { label: '黑燈', data: monthAggregates.map(m => m.black), backgroundColor: '#374151' },
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } },
        plugins: { legend: { position: 'bottom' } },
      }
    });

    // 4c: Category Weakness - 各類別團隊平均得分率
    const catMaxMap = { '出席': 20, '一對一': 15, '教育': 15, '引薦': 20, '來賓': 15, '金額': 15 };
    const catNames = Object.keys(catMaxMap);
    const catScoreRate = catNames.map(cat => {
      const max = catMaxMap[cat];
      const key = cat === '出席' ? 'attendance' : cat === '一對一' ? 'oneOnOne' : cat === '教育' ? 'training' : cat === '引薦' ? 'referral' : cat === '來賓' ? 'guest' : 'value';
      const avgScore = membersData.reduce((s, m) => s + m.scores[key], 0) / membersData.length;
      const rate = Math.round(avgScore / max * 100);
      return { cat, rate, avgScore: avgScore.toFixed(1), max };
    }).sort((a, b) => a.rate - b.rate); // 最弱排前面

    charts.weakness = new Chart($('#chart-weakness'), {
      type: 'bar',
      data: {
        labels: catScoreRate.map(d => `${d.cat}（${d.avgScore}/${d.max}）`),
        datasets: [{
          label: '平均得分率 %',
          data: catScoreRate.map(d => d.rate),
          backgroundColor: catScoreRate.map(d => d.rate >= 70 ? '#22c55e' : d.rate >= 40 ? '#eab308' : '#ef4444'),
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: 'y',
        scales: { x: { min: 0, max: 100, ticks: { callback: v => v + '%' } } },
        plugins: { legend: { display: false } },
      }
    });

    // 4d: Score Distribution
    const buckets = Array(10).fill(0);
    membersData.forEach(m => {
      const idx = Math.min(Math.floor(m.scores.total / 10), 9);
      buckets[idx]++;
    });
    const bucketLabels = buckets.map((_, i) => `${i * 10}-${i * 10 + 9}`);
    bucketLabels[9] = '90-100';

    charts.dist = new Chart($('#chart-distribution'), {
      type: 'bar',
      data: {
        labels: bucketLabels,
        datasets: [{
          label: '人數',
          data: buckets,
          backgroundColor: bucketLabels.map((_, i) => {
            if (i >= 7) return '#22c55e';
            if (i >= 5) return '#eab308';
            if (i >= 3) return '#ef4444';
            return '#374151';
          }),
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } },
        plugins: { legend: { display: false } },
      }
    });
  }

  // --- Meeting Prep ---
  function renderMeetingPrep(membersData, alerts, monthAggregates) {
    // Praise
    const praiseEl = $('#meeting-praise');
    const praiseList = [];
    // Top improvers
    alerts.improving.slice(0, 5).forEach(m => {
      const prev = m.scores.total - m.twoMonthDelta;
      praiseList.push(`<div class="meeting-item">
        <span class="meeting-icon">🌟</span>
        <div class="meeting-text">
          <strong>${m.id} ${m.name}</strong>
          <span class="sub">進步 ${m.twoMonthDelta} 分（${prev} → ${m.scores.total}）</span>
        </div>
      </div>`);
    });
    // Perfect or near-perfect
    membersData.filter(m => m.scores.total >= 90).sort((a, b) => b.scores.total - a.scores.total).forEach(m => {
      praiseList.push(`<div class="meeting-item">
        <span class="meeting-icon">🏆</span>
        <div class="meeting-text">
          <strong>${m.id} ${m.name}</strong>
          <span class="sub">總分 ${m.scores.total} 分，表現卓越</span>
        </div>
      </div>`);
    });
    praiseEl.innerHTML = praiseList.length ? praiseList.join('') : '<div class="meeting-empty">本期暫無特別表揚對象</div>';

    // Care
    const careEl = $('#meeting-care');
    const careList = [];
    // Declining first
    alerts.declining.slice(0, 5).forEach(m => {
      careList.push(`<div class="meeting-item">
        <span class="meeting-icon">⚠️</span>
        <div class="meeting-text">
          <strong>${m.id} ${m.name}</strong>
          <span class="sub">下降 ${Math.abs(m.twoMonthDelta)} 分，需優先關懷</span>
        </div>
      </div>`);
    });
    // Persistent low
    alerts.persistentLow.forEach(m => {
      careList.push(`<div class="meeting-item">
        <span class="meeting-icon">🔔</span>
        <div class="meeting-text">
          <strong>${m.id} ${m.name}</strong>
          <span class="sub">連續 ${m.lowStreak} 個月低於 50 分（目前 ${m.scores.total} 分）</span>
        </div>
      </div>`);
    });
    careEl.innerHTML = careList.length ? careList.join('') : '<div class="meeting-empty">目前沒有需要特別關懷的會員</div>';

    // Focus - weakest category
    const focusEl = $('#meeting-focus');
    const catMaxMap2 = { '出席': 20, '一對一': 15, '教育': 15, '引薦': 20, '來賓': 15, '金額': 15 };
    const catNames2 = Object.keys(catMaxMap2);
    const catStats = catNames2.map(cat => {
      const key = cat === '出席' ? 'attendance' : cat === '一對一' ? 'oneOnOne' : cat === '教育' ? 'training' : cat === '引薦' ? 'referral' : cat === '來賓' ? 'guest' : 'value';
      const max = catMaxMap2[cat];
      const avgScore = membersData.reduce((s, m) => s + m.scores[key], 0) / membersData.length;
      const rate = Math.round(avgScore / max * 100);
      return { cat, rate, avgScore: avgScore.toFixed(1), max };
    }).sort((a, b) => a.rate - b.rate); // 最弱排前面

    const weakest = catStats[0];
    const actionMap = {
      '出席': '鼓勵全勤出席，無法出席請安排替代人',
      '一對一': '推動每週至少安排 2 次一對一會面',
      '教育': '舉辦更多分會教育活動或線上課程',
      '引薦': '設定每週引薦目標，互相提醒',
      '來賓': '每月主題邀請來賓活動',
      '金額': '追蹤引薦案件成交進度',
    };

    focusEl.innerHTML = `<div class="meeting-highlight">
      團隊最弱環節：<strong>${weakest.cat}</strong>（平均得分率 ${weakest.rate}%，平均 ${weakest.avgScore} / ${weakest.max}）<br><br>
      建議行動：<strong>${actionMap[weakest.cat]}</strong>
    </div>`;

    // Goal
    const goalEl = $('#meeting-goal');
    const currentAvg = membersData.reduce((s, m) => s + m.scores.total, 0) / membersData.length;
    const belowGreen = membersData.filter(m => m.scores.total < 70 && m.scores.total >= 50);
    const greenCount = membersData.filter(m => m.scores.total >= 70).length;
    const greenRate = Math.round(greenCount / membersData.length * 100);

    // How much each yellow member needs on average to reach green
    const yellowAvgGap = belowGreen.length
      ? Math.ceil(belowGreen.reduce((s, m) => s + (70 - m.scores.total), 0) / belowGreen.length)
      : 0;

    // If all yellow members reach green, new green rate
    const newGreenCount = greenCount + belowGreen.length;
    const newGreenRate = Math.round(newGreenCount / membersData.length * 100);

    goalEl.innerHTML = `<div class="meeting-highlight">
      目前團隊平均：<strong>${currentAvg.toFixed(1)} 分</strong>，綠燈率 <strong>${greenRate}%</strong>（${greenCount} 人）<br><br>
      黃燈會員有 <strong>${belowGreen.length} 位</strong>，距離綠燈最近，是提升重點對象。<br><br>
      若這 ${belowGreen.length} 位黃燈會員平均各提升 <strong>${yellowAvgGap} 分</strong> 即可達到綠燈，團隊綠燈率將提升至 <strong>${newGreenRate}%</strong>（${newGreenCount} 人）。
    </div>`;
  }

  // --- Modal ---
  let modalChart = null;

  async function openMemberModal(memberId) {
    const modal = $('#member-modal');
    const body = $('#modal-body');
    modal.classList.remove('hidden');
    body.innerHTML = '<div style="text-align:center;padding:40px"><div class="spinner"></div><p>載入中...</p></div>';

    try {
      const data = await BNIData.getMemberDashboardData(memberId);
      const s = data.scores;

      body.innerHTML = `
        <div class="modal-member-header">
          <h2>${data.member.display}</h2>
        </div>
        <div class="modal-traffic">
          <div class="modal-traffic-score" style="color:${s.light.color}">${s.total}</div>
          <div class="modal-traffic-label" style="color:${s.light.color}">${s.light.label}</div>
        </div>
        <div class="modal-metrics">
          ${s.items.map(item => `
            <div class="modal-metric">
              <div class="modal-metric-name">${item.name}</div>
              <div class="modal-metric-score" style="color:${item.score >= item.max ? 'var(--green)' : item.score === 0 ? '#ccc' : 'var(--text)'}">${item.score}</div>
              <div class="modal-metric-max">/ ${item.max}</div>
            </div>
          `).join('')}
        </div>
        <div class="modal-chart-container"><canvas id="modal-chart"></canvas></div>
      `;

      // Render trend chart in modal
      if (modalChart) { modalChart.destroy(); modalChart = null; }
      const labels = data.monthlyData.map(r => r.month.display);
      modalChart = new Chart($('#modal-chart'), {
        type: 'line',
        data: {
          labels,
          datasets: [{
            label: '總分',
            data: data.monthlyData.map(r => r.scores.總分),
            borderColor: '#3b82f6',
            backgroundColor: 'rgba(59,130,246,0.1)',
            fill: true,
            tension: 0.3,
            pointRadius: 5,
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: { y: { min: 0, max: 100 } },
          plugins: { legend: { display: false } },
        }
      });
    } catch (e) {
      body.innerHTML = `<p style="color:var(--red);text-align:center">載入失敗：${e.message}</p>`;
    }
  }

  function closeMemberModal() {
    $('#member-modal').classList.add('hidden');
    if (modalChart) { modalChart.destroy(); modalChart = null; }
  }

  // =====================
  // EVENT HANDLERS
  // =====================

  function setupEvents() {
    // Tab switching
    $$('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        $$('.tab-btn').forEach(b => b.classList.remove('active'));
        $$('.tab-content').forEach(c => c.classList.remove('active'));
        btn.classList.add('active');
        $(`#tab-${btn.dataset.tab}`).classList.add('active');
        window.scrollTo(0, 0);

        // Lazy render charts
        if (btn.dataset.tab === 'analytics' && Object.keys(charts).length === 0) {
          renderCharts(teamData.monthAggregates, teamData.membersData);
        }
      });
    });

    // Table sort
    $$('.member-table th[data-sort]').forEach(th => {
      th.addEventListener('click', () => {
        const key = th.dataset.sort;
        if (currentSort.key === key) {
          currentSort.dir = currentSort.dir === 'desc' ? 'asc' : 'desc';
        } else {
          currentSort = { key, dir: 'desc' };
        }
        // Update header styles
        $$('.member-table th').forEach(h => h.classList.remove('sort-active', 'sort-asc', 'sort-desc'));
        th.classList.add('sort-active', currentSort.dir === 'asc' ? 'sort-asc' : 'sort-desc');
        renderTable(teamData.membersData);
      });
    });

    // Filter buttons
    $$('.filter-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        $$('.filter-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentFilter = btn.dataset.filter;
        renderTable(teamData.membersData);
      });
    });

    // Search
    $('#search-input').addEventListener('input', (e) => {
      currentSearch = e.target.value.trim();
      renderTable(teamData.membersData);
    });

    // Table row click → modal
    $('#table-body').addEventListener('click', (e) => {
      const row = e.target.closest('tr[data-id]');
      if (row) openMemberModal(row.dataset.id);
    });

    // Modal close
    $('#modal-close').addEventListener('click', closeMemberModal);
    $('#member-modal').addEventListener('click', (e) => {
      if (e.target === $('#member-modal')) closeMemberModal();
    });
  }

  // =====================
  // INIT
  // =====================

  document.addEventListener('DOMContentLoaded', async () => {
    try {
      teamData = await buildTeamData();
      const alerts = computeAlerts(teamData.membersData);

      // Render all sections
      renderKPI(teamData.membersData, teamData.monthAggregates);
      renderLightDistribution(teamData.membersData);
      renderTable(teamData.membersData);
      renderAlerts(alerts);
      renderMeetingPrep(teamData.membersData, alerts, teamData.monthAggregates);

      // Setup events
      setupEvents();

      // Show dashboard
      $('#loading-screen').classList.add('hidden');
      $('#admin-dashboard').classList.remove('hidden');
    } catch (e) {
      $('#loading-screen').innerHTML = `
        <div class="loading-content">
          <p style="color:var(--red)">載入失敗：${e.message}</p>
          <p style="margin-top:8px"><a href="" onclick="location.reload()">重新載入</a></p>
        </div>
      `;
    }
  });
})();
