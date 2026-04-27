const columns = {
  customerCode: "קוד לקוח",
  created: "נוצר בתאריך",
  owner: "מנהל לקוח",
  customerStatus: "סטטוס לקוח",
  source: "מקור הגעה",
  potential: "פוטנציאל לקוח",
  leadStatus: "סטטוס ליד",
  saleStatus: "סטטוס מכירה",
  quoteSent: "האם נשלחה הצעת מחיר",
  quoteStatus: "סטטוס הצעת מחיר",
  closeReason: "סיבת סגירה",
  notes: "מידע נוסף",
  email: "דואר אלקטרוני",
};

const state = {
  rows: [],
  filtered: [],
  headers: [],
  sourceName: "",
  activeView: "overview",
};

const el = {
  input: document.querySelector("#csvInput"),
  fileName: document.querySelector("#fileName"),
  loadSample: document.querySelector("#loadSampleBtn"),
  exportBtn: document.querySelector("#exportBtn"),
  exportPriorityBtn: document.querySelector("#exportPriorityBtn"),
  copyReportBtn: document.querySelector("#copyReportBtn"),
  resetBtn: document.querySelector("#resetBtn"),
  empty: document.querySelector("#emptyState"),
  metrics: document.querySelector("#metrics"),
  insightsPanel: document.querySelector("#insightsPanel"),
  insightsList: document.querySelector("#insightsList"),
  filtersPanel: document.querySelector("#filtersPanel"),
  commandCenter: document.querySelector("#commandCenter"),
  commandCenterCards: document.querySelector("#commandCenterCards"),
  viewSummary: document.querySelector("#viewSummary"),
  analytics: document.querySelector("#analytics"),
  resultCount: document.querySelector("#resultCount"),
  search: document.querySelector("#searchInput"),
  owner: document.querySelector("#ownerFilter"),
  source: document.querySelector("#sourceFilter"),
  potential: document.querySelector("#potentialFilter"),
  leadStatus: document.querySelector("#leadStatusFilter"),
  saleStatus: document.querySelector("#saleStatusFilter"),
  closeReason: document.querySelector("#closeReasonFilter"),
  fromDate: document.querySelector("#fromDate"),
  toDate: document.querySelector("#toDate"),
  leadStatusChart: document.querySelector("#leadStatusChart"),
  sourceChart: document.querySelector("#sourceChart"),
  funnelChart: document.querySelector("#funnelChart"),
  dailyTrend: document.querySelector("#dailyTrend"),
  closeReasonChart: document.querySelector("#closeReasonChart"),
  duplicateTable: document.querySelector("#duplicateTable"),
  sourceQualityTable: document.querySelector("#sourceQualityTable"),
  hourChart: document.querySelector("#hourChart"),
  weekdayChart: document.querySelector("#weekdayChart"),
  sourceStatusMatrix: document.querySelector("#sourceStatusMatrix"),
  sourceRecommendationsTable: document.querySelector("#sourceRecommendationsTable"),
  ownerPerformanceTable: document.querySelector("#ownerPerformanceTable"),
  ownerFocusTable: document.querySelector("#ownerFocusTable"),
  keywordCloud: document.querySelector("#keywordCloud"),
  staleLeadsTable: document.querySelector("#staleLeadsTable"),
  executiveReport: document.querySelector("#executiveReport"),
  actionPlan: document.querySelector("#actionPlan"),
  scoreBucketsChart: document.querySelector("#scoreBucketsChart"),
  targetInterest: document.querySelector("#targetInterestInput"),
  targetQuote: document.querySelector("#targetQuoteInput"),
  targetIrrelevant: document.querySelector("#targetIrrelevantInput"),
  goalTracker: document.querySelector("#goalTracker"),
  dataQualityCards: document.querySelector("#dataQualityCards"),
  agingBuckets: document.querySelector("#agingBuckets"),
  potentialQualityTable: document.querySelector("#potentialQualityTable"),
  priorityLeadsTable: document.querySelector("#priorityLeadsTable"),
  groupBy: document.querySelector("#groupBySelect"),
  groupTable: document.querySelector("#groupTable"),
  leadsTable: document.querySelector("#leadsTable"),
  reportTabs: [...document.querySelectorAll(".report-tab")],
  reportViews: [...document.querySelectorAll(".report-view")],
};

const viewDescriptions = {
  overview: "מבט על על כל המערכת: סטטוסים, יעדים, ציון איכות והמלצות פעולה.",
  acquisition: "מראה מאיפה הלידים מגיעים, מתי הם מגיעים, ואילו ערוצים מייצרים איכות מול רעש.",
  team: "משקף ביצועי צוות, עומסים, לידים חמים ואיפה כדאי לחלק משימות מחדש.",
  quality: "מאתר חסמים, בעיות דאטה, גיל לידים ואותות שחוזרים בשיחות.",
  operations: "מרכז את המשימות המיידיות, הלידים התקועים, טבלאות עבודה והרשומות עצמן.",
};

const filterMap = [
  [el.owner, columns.owner],
  [el.source, columns.source],
  [el.potential, columns.potential],
  [el.leadStatus, columns.leadStatus],
  [el.saleStatus, columns.saleStatus],
  [el.closeReason, columns.closeReason],
];

el.input.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  const text = await file.text();
  loadCsv(text, file.name);
});

el.loadSample.addEventListener("click", async () => {
  if (window.EMBEDDED_SAMPLE_CSV) {
    loadCsv(window.EMBEDDED_SAMPLE_CSV, "sample-leads.csv");
    return;
  }

  try {
    const response = await fetch("data/sample-leads.csv", { cache: "no-store" });
    if (response.ok) {
      loadCsv(await response.text(), "sample-leads.csv");
      return;
    }
  } catch (error) {
    // GitHub Pages can still run from a single HTML file without a data folder.
  }

  alert("לא נמצא קובץ דוגמה. אפשר להעלות קובץ CSV ידני או לבדוק שחבילת המערכת מלאה.");
});

el.resetBtn.addEventListener("click", () => {
  [el.search, el.fromDate, el.toDate].forEach((input) => {
    input.value = "";
  });
  filterMap.forEach(([select]) => {
    select.value = "__all";
  });
  applyFilters();
});

el.exportBtn.addEventListener("click", exportSummary);
el.exportPriorityBtn.addEventListener("click", exportPriorityLeads);
el.copyReportBtn.addEventListener("click", copyExecutiveReport);

[el.search, el.fromDate, el.toDate, el.groupBy, el.targetInterest, el.targetQuote, el.targetIrrelevant].forEach((input) => {
  input.addEventListener("input", applyFilters);
});

filterMap.forEach(([select]) => {
  select.addEventListener("change", applyFilters);
});

el.reportTabs.forEach((button) => {
  button.addEventListener("click", () => {
    setActiveView(button.dataset.view);
  });
});

function loadCsv(text, name) {
  const parsed = parseCsv(text);
  if (parsed.length < 2) {
    alert("הקובץ לא מכיל מספיק נתונים לניתוח.");
    return;
  }

  state.headers = parsed[0].map(clean);
  state.rows = parsed.slice(1).map((cells, index) => {
    const row = { _rowNumber: index + 1 };
    state.headers.forEach((header, headerIndex) => {
      row[header] = clean(cells[headerIndex]);
    });
    row._createdDate = parseHebrewDate(row[columns.created]);
    row._search = state.headers.map((header) => row[header]).join(" ").toLowerCase();
    return row;
  });

  state.sourceName = name;
  el.fileName.textContent = name;
  el.empty.hidden = true;
  el.metrics.hidden = false;
  el.insightsPanel.hidden = false;
  el.filtersPanel.hidden = false;
  el.commandCenter.hidden = false;
  el.analytics.hidden = false;
  buildControls();
  setActiveView(state.activeView);
  applyFilters();
}

function buildControls() {
  filterMap.forEach(([select, column]) => {
    fillSelect(select, uniqueValues(state.rows, column));
  });

  const groupColumns = [
    columns.leadStatus,
    columns.saleStatus,
    columns.source,
    columns.owner,
    columns.potential,
    columns.closeReason,
    columns.quoteSent,
    columns.quoteStatus,
  ].filter((column) => state.headers.includes(column));

  el.groupBy.innerHTML = groupColumns
    .map((column) => `<option value="${escapeHtml(column)}">${escapeHtml(column)}</option>`)
    .join("");

  const timestamps = state.rows.map((row) => row._createdDate?.getTime()).filter(Boolean);
  if (timestamps.length) {
    el.fromDate.min = toDateInput(new Date(Math.min(...timestamps)));
    el.fromDate.max = toDateInput(new Date(Math.max(...timestamps)));
    el.toDate.min = el.fromDate.min;
    el.toDate.max = el.fromDate.max;
  }
}

function setActiveView(view) {
  state.activeView = viewDescriptions[view] ? view : "overview";

  el.reportTabs.forEach((button) => {
    const active = button.dataset.view === state.activeView;
    button.classList.toggle("active", active);
    button.classList.toggle("ghost", !active);
    button.setAttribute("aria-pressed", active ? "true" : "false");
  });

  el.reportViews.forEach((panel) => {
    panel.classList.toggle("active", panel.dataset.view === state.activeView);
  });

  updateViewSummary();
}

function applyFilters() {
  const search = el.search.value.trim().toLowerCase();
  const from = el.fromDate.value ? new Date(`${el.fromDate.value}T00:00:00`) : null;
  const to = el.toDate.value ? new Date(`${el.toDate.value}T23:59:59`) : null;

  state.filtered = state.rows.filter((row) => {
    if (search && !row._search.includes(search)) return false;
    if (from && (!row._createdDate || row._createdDate < from)) return false;
    if (to && (!row._createdDate || row._createdDate > to)) return false;

    return filterMap.every(([select, column]) => {
      return select.value === "__all" || normalize(row[column]) === select.value;
    });
  });

  render();
}

function render() {
  renderMetrics();
  renderInsights();
  el.resultCount.textContent = `${formatNumber(state.filtered.length)} מתוך ${formatNumber(state.rows.length)} לידים`;
  renderCommandCenter();
  renderBarChart(el.leadStatusChart, countBy(state.filtered, columns.leadStatus));
  renderBarChart(el.sourceChart, countBy(state.filtered, columns.source));
  renderFunnel();
  renderDailyTrend();
  renderBarChart(el.closeReasonChart, countCloseReasonsForIrrelevant(state.filtered));
  renderDuplicateTable();
  renderSourceQualityTable();
  renderSourceRecommendationsTable();
  renderHourChart();
  renderWeekdayChart();
  renderSourceStatusMatrix();
  renderOwnerPerformanceTable();
  renderOwnerFocusTable();
  renderKeywordCloud();
  renderStaleLeadsTable();
  renderExecutiveReport();
  renderActionPlan();
  renderScoreBuckets();
  renderGoalTracker();
  renderDataQualityCards();
  renderAgingBuckets();
  renderPotentialQualityTable();
  renderPriorityLeadsTable();
  renderGroupTable();
  renderLeadsTable();
  updateViewSummary();
}

function renderCommandCenter() {
  const total = state.filtered.length;
  if (!total) {
    el.commandCenterCards.innerHTML = `
      <article class="command-card">
        <strong>מרכז הבקרה מוכן</strong>
        <span>לחץ על "הפעל ניתוח סטטיסטי" או העלה קובץ CSV כדי להציג KPI, דוחות ומסקנות.</span>
      </article>
      <article class="command-card">
        <strong>מה תקבל כאן?</strong>
        <span>מסך מרכזי עם סיכום מצב, מעבר בין דוחות, מסקנות שיווק, ביצועי צוות וסיכונים.</span>
      </article>`;
    updateViewSummary();
    return;
  }

  const averageScore = getAverageLeadScore(state.filtered);
  const bestSource = buildGroupStats(state.filtered, columns.source)
    .filter((group) => group.total >= 2)
    .sort((a, b) => b.score - a.score || b.total - a.total)[0];
  const mainRisk = getMainRiskText();
  const actions = buildActionPlan();
  const openLeads = state.filtered.filter(isOpenLead).length;

  const cards = [
    ["מצב כללי", `${formatNumber(total)} לידים בסינון הנוכחי עם ציון איכות ממוצע ${averageScore}/100.`],
    ["מוקד הזדמנות", bestSource ? `${bestSource.label} בולט כמקור איכותי עם ${bestSource.interestRate}% יחס עניין.` : "אין עדיין מספיק עומק להשוואת מקורות."],
    ["מוקד סיכון", mainRisk],
    ["פוקוס ניהולי", `${formatNumber(openLeads)} לידים עדיין פתוחים. ${actions[0] || "כרגע אין עומס חריג שמצריך פעולה מיידית."}`],
  ];

  el.commandCenterCards.innerHTML = cards
    .map(([title, body]) => `<article class="command-card"><strong>${escapeHtml(title)}</strong><span>${escapeHtml(body)}</span></article>`)
    .join("");
}

function updateViewSummary() {
  if (!el.viewSummary) return;

  const total = state.filtered.length;
  const summary = viewDescriptions[state.activeView] || viewDescriptions.overview;
  if (!total) {
    el.viewSummary.textContent = summary;
    return;
  }

  if (state.activeView === "acquisition") {
    const bestSource = buildGroupStats(state.filtered, columns.source)
      .filter((group) => group.total >= 2)
      .sort((a, b) => b.score - a.score || b.total - a.total)[0];
    el.viewSummary.textContent = bestSource
      ? `${summary} כרגע ${bestSource.label} מוביל באיכות.`
      : summary;
    return;
  }

  if (state.activeView === "team") {
    const stale = getPriorityLeads(state.filtered).filter((item) => item.age >= 3).length;
    el.viewSummary.textContent = `${summary} יש ${formatNumber(stale)} לידים דחופים שדורשים חלוקת טיפול.`;
    return;
  }

  if (state.activeView === "quality") {
    const mainRisk = getMainRiskText();
    el.viewSummary.textContent = `${summary} כרגע בולט: ${mainRisk}`;
    return;
  }

  if (state.activeView === "operations") {
    const priority = getPriorityLeads(state.filtered).length;
    el.viewSummary.textContent = `${summary} ${formatNumber(priority)} לידים נמצאו לטיפול מהיר.`;
    return;
  }

  el.viewSummary.textContent = summary;
}

function renderMetrics() {
  const total = state.filtered.length;
  const interested = state.filtered.filter(isInterested).length;
  const irrelevant = state.filtered.filter(isIrrelevant).length;
  const otherProcess = state.filtered.filter(isOtherProcessWithoutCloseReason).length;
  const quotes = state.filtered.filter(hasQuote).length;
  const conversion = total ? Math.round((interested / total) * 100) : 0;

  const metrics = [
    ["סה״כ לידים", formatNumber(total)],
    ["מעוניינים / בטיפול", formatNumber(interested)],
    ["לא רלוונטיים", formatNumber(irrelevant)],
    ["עברו תהליך אחר", formatNumber(otherProcess)],
    ["הצעות מחיר", formatNumber(quotes)],
    ["יחס עניין", `${conversion}%`],
  ];

  el.metrics.innerHTML = metrics
    .map(([label, value]) => `<article class="metric"><span>${label}</span><strong>${value}</strong></article>`)
    .join("");
}

function renderInsights() {
  const total = state.filtered.length;
  const sourceStats = buildGroupStats(state.filtered, columns.source).filter((group) => group.total >= 2);
  const ownerStats = buildGroupStats(state.filtered, columns.owner).filter((group) => group.total >= 2);
  const closeReasons = Object.entries(countCloseReasonsForIrrelevant(state.filtered)).sort((a, b) => b[1] - a[1]);
  const otherProcessCount = state.filtered.filter(isOtherProcessWithoutCloseReason).length;
  const duplicateCount = getDuplicateCustomerCodes(state.filtered).reduce((sum, item) => sum + item.count - 1, 0);
  const bestHour = getBestHour(state.filtered);
  const staleCount = state.filtered.filter((row) => row._createdDate && isOpenLead(row) && daysBetween(row._createdDate, getLatestDate(state.rows) || new Date()) >= 3).length;
  const insights = [];

  if (!total) {
    el.insightsList.innerHTML = `<article class="insight-card">אין מספיק נתונים בסינון הנוכחי. נסה להרחיב את טווח התאריכים או לנקות חלק מהסינונים.</article>`;
    return;
  }

  const bestSource = sourceStats.sort((a, b) => b.score - a.score || b.total - a.total)[0];
  const weakestSource = sourceStats.sort((a, b) => a.score - b.score || b.total - a.total)[0];
  const bestOwner = ownerStats.sort((a, b) => b.interestRate - a.interestRate || b.total - a.total)[0];
  const topCloseReason = closeReasons[0];

  if (bestSource) insights.push(["המקור החזק ביותר", `${bestSource.label} מוביל בציון איכות ${bestSource.score}/100 עם ${percent(bestSource.interested, bestSource.total)} יחס עניין.`]);
  if (weakestSource && weakestSource.label !== bestSource?.label) insights.push(["מקור שדורש בדיקה", `${weakestSource.label} מציג ציון איכות ${weakestSource.score}/100. כדאי לבדוק מסרים, קהל יעד או סינון מוקדם.`]);
  if (bestOwner) insights.push(["ביצועי נציגים", `${bestOwner.label} מציג/ה יחס עניין של ${percent(bestOwner.interested, bestOwner.total)} מתוך ${formatNumber(bestOwner.total)} לידים.`]);
  if (topCloseReason) insights.push(["חסם מרכזי", `בקרב לידים לא רלוונטיים, ${topCloseReason[0]} חוזרת ${formatNumber(topCloseReason[1])} פעמים.`]);
  if (bestHour) insights.push(["חלון זמן חזק", `השעה ${String(bestHour.hour).padStart(2, "0")}:00 בולטת עם ${formatNumber(bestHour.count)} לידים.`]);
  if (duplicateCount) insights.push(["איכות רשומות", `נמצאו ${formatNumber(duplicateCount)} רשומות כפולות לפי קוד לקוח. כדאי לאחד כפילויות לפני קבלת החלטות.`]);
  if (otherProcessCount) insights.push(["תהליך אחר", `${formatNumber(otherProcessCount)} לידים ללא סיבת סגירה מסווגים כתהליך אחר ולא כחסם.`]);
  if (staleCount) insights.push(["לידים תקועים", `${formatNumber(staleCount)} לידים פתוחים כבר בני 3 ימים ומעלה ודורשים טיפול.`]);

  el.insightsList.innerHTML = insights
    .map(([title, body]) => `<article class="insight-card"><strong>${escapeHtml(title)}</strong><span>${escapeHtml(body)}</span></article>`)
    .join("");
}

function renderBarChart(container, counts) {
  const entries = Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  const max = Math.max(1, ...entries.map(([, value]) => value));

  container.innerHTML = entries.length
    ? entries
        .map(([label, value]) => {
          const width = Math.max(5, Math.round((value / max) * 100));
          return `
            <div class="bar-row">
              <span class="bar-label" title="${escapeHtml(label)}">${escapeHtml(label)}</span>
              <span class="bar-track"><span class="bar-fill" style="width:${width}%"></span></span>
              <span class="bar-value">${formatNumber(value)}</span>
            </div>`;
        })
        .join("")
    : `<p class="muted">אין נתונים להצגה בסינון הנוכחי.</p>`;
}

function renderFunnel() {
  const total = state.filtered.length;
  const interested = state.filtered.filter(isInterested).length;
  const quotes = state.filtered.filter(hasQuote).length;
  const followUp = state.filtered.filter((row) => /מעקב|המשך טיפול|פגישה/.test(`${normalize(row[columns.leadStatus])} ${normalize(row[columns.saleStatus])}`)).length;
  const irrelevant = state.filtered.filter(isIrrelevant).length;
  const steps = [
    ["לידים שנכנסו", total],
    ["עניין / טיפול", interested],
    ["מעקב פעיל", followUp],
    ["הצעות מחיר", quotes],
    ["לא רלוונטיים", irrelevant],
  ];
  const max = Math.max(1, total);

  el.funnelChart.innerHTML = steps
    .map(([label, value]) => {
      const width = Math.max(8, Math.round((value / max) * 100));
      return `
        <div class="funnel-step" style="width:${width}%">
          <span>${escapeHtml(label)}</span>
          <strong>${formatNumber(value)}</strong>
          <small>${percent(value, total)}</small>
        </div>`;
    })
    .join("");
}

function renderDailyTrend() {
  const counts = countByDate(state.filtered);
  const entries = Object.entries(counts).sort((a, b) => a[0].localeCompare(b[0]));
  const max = Math.max(1, ...entries.map(([, value]) => value));

  el.dailyTrend.innerHTML = entries.length
    ? entries
        .map(([date, value]) => {
          const height = Math.max(12, Math.round((value / max) * 140));
          return `
            <div class="trend-day" title="${formatDateLabel(date)}: ${formatNumber(value)}">
              <span class="trend-bar" style="height:${height}px"></span>
              <small>${formatShortDate(date)}</small>
              <strong>${formatNumber(value)}</strong>
            </div>`;
        })
        .join("")
    : `<p class="muted">אין תאריכים תקינים להצגת מגמה.</p>`;
}

function renderDuplicateTable() {
  const duplicates = getDuplicateCustomerCodes(state.filtered).slice(0, 12);
  el.duplicateTable.innerHTML = duplicates.length
    ? duplicates
        .map((item) => `
          <tr>
            <td>${escapeHtml(item.code)}</td>
            <td>${formatNumber(item.count)}</td>
            <td>${escapeHtml(item.owners.join(", "))}</td>
          </tr>`)
        .join("")
    : `<tr><td colspan="3" class="muted">לא נמצאו קודי לקוח כפולים בסינון הנוכחי.</td></tr>`;
}

function renderSourceQualityTable() {
  const stats = buildGroupStats(state.filtered, columns.source).sort((a, b) => b.score - a.score || b.total - a.total);
  el.sourceQualityTable.innerHTML = stats.length
    ? stats
        .map((group) => `
          <tr>
            <td><span class="badge">${escapeHtml(group.label)}</span></td>
            <td>${formatNumber(group.total)}</td>
            <td>${formatNumber(group.interested)}</td>
            <td>${formatNumber(group.irrelevant)}</td>
            <td>${formatNumber(group.quotes)}</td>
            <td>${group.interestRate}%</td>
            <td><span class="score">${group.score}</span></td>
          </tr>`)
        .join("")
    : `<tr><td colspan="7" class="muted">אין נתונים להצגה.</td></tr>`;
}

function renderSourceRecommendationsTable() {
  const rows = buildSourceRecommendations(state.filtered);
  el.sourceRecommendationsTable.innerHTML = rows.length
    ? rows
        .map((item) => `
          <tr>
            <td><span class="badge">${escapeHtml(item.label)}</span></td>
            <td>${formatNumber(item.total)}</td>
            <td><span class="score">${item.score}</span></td>
            <td>${escapeHtml(item.highlight)}</td>
            <td>${escapeHtml(item.recommendation)}</td>
          </tr>`)
        .join("")
    : `<tr><td colspan="5" class="muted">אין מספיק מקורות כדי לגבש המלצה.</td></tr>`;
}

function renderHourChart() {
  const counts = Array.from({ length: 24 }, (_, hour) => ({ hour, value: 0 }));
  state.filtered.forEach((row) => {
    if (row._createdDate) counts[row._createdDate.getHours()].value += 1;
  });
  const active = counts.filter((item) => item.value > 0);
  const max = Math.max(1, ...counts.map((item) => item.value));

  el.hourChart.innerHTML = active.length
    ? counts
        .map((item) => {
          const height = item.value ? Math.max(10, Math.round((item.value / max) * 118)) : 3;
          return `
            <div class="hour-cell" title="${String(item.hour).padStart(2, "0")}:00 - ${formatNumber(item.value)} לידים">
              <span class="hour-stick" style="height:${height}px"></span>
              <small>${String(item.hour).padStart(2, "0")}</small>
            </div>`;
        })
        .join("")
    : `<p class="muted">אין שעות תקינות להצגה.</p>`;
}

function renderWeekdayChart() {
  const names = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"];
  const counts = names.map((name, index) => ({ name, index, value: 0, interested: 0 }));
  state.filtered.forEach((row) => {
    if (!row._createdDate) return;
    const day = row._createdDate.getDay();
    counts[day].value += 1;
    counts[day].interested += isInterested(row) ? 1 : 0;
  });
  const max = Math.max(1, ...counts.map((item) => item.value));

  el.weekdayChart.innerHTML = counts
    .map((item) => {
      const width = item.value ? Math.max(8, Math.round((item.value / max) * 100)) : 2;
      return `
        <div class="weekday-row">
          <span>${item.name}</span>
          <div class="weekday-track">
            <b style="width:${width}%"></b>
          </div>
          <strong>${formatNumber(item.value)}</strong>
          <small>${percent(item.interested, item.value)} עניין</small>
        </div>`;
    })
    .join("");
}

function renderSourceStatusMatrix() {
  const sources = topKeys(countBy(state.filtered, columns.source), 8);
  const statuses = topKeys(countBy(state.filtered, columns.leadStatus), 8);

  if (!sources.length || !statuses.length) {
    el.sourceStatusMatrix.innerHTML = `<tbody><tr><td class="muted">אין נתונים להצגת מטריצה.</td></tr></tbody>`;
    return;
  }

  const rows = sources.map((source) => {
    const cells = statuses.map((status) => {
      const count = state.filtered.filter((row) => normalize(row[columns.source]) === source && normalize(row[columns.leadStatus]) === status).length;
      return `<td class="${count ? "heat-cell" : ""}" style="--heat:${Math.min(1, count / Math.max(1, state.filtered.length / 8))}">${count ? formatNumber(count) : "-"}</td>`;
    });
    return `<tr><th>${escapeHtml(source)}</th>${cells.join("")}</tr>`;
  });

  el.sourceStatusMatrix.innerHTML = `
    <thead><tr><th>מקור / סטטוס</th>${statuses.map((status) => `<th>${escapeHtml(status)}</th>`).join("")}</tr></thead>
    <tbody>${rows.join("")}</tbody>`;
}

function renderOwnerPerformanceTable() {
  const duplicateCustomerCodes = new Set(getDuplicateCustomerCodes(state.filtered).map((item) => item.code));
  const stats = buildGroupStats(state.filtered, columns.owner)
    .map((group) => {
      const duplicateCount = state.filtered.filter((row) => {
        const code = clean(row[columns.customerCode]).trim();
        return normalize(row[columns.owner]) === group.label && duplicateCustomerCodes.has(code);
      }).length;
      return { ...group, duplicateCount, score: Math.max(0, group.score - Math.min(20, duplicateCount * 2)) };
    })
    .sort((a, b) => b.score - a.score || b.total - a.total);

  el.ownerPerformanceTable.innerHTML = stats.length
    ? stats
        .map((group) => `
          <tr>
            <td><span class="badge">${escapeHtml(group.label)}</span></td>
            <td>${formatNumber(group.total)}</td>
            <td>${formatNumber(group.interested)}</td>
            <td>${formatNumber(group.irrelevant)}</td>
            <td>${formatNumber(group.quotes)}</td>
            <td>${formatNumber(group.duplicateCount)}</td>
            <td><span class="score">${group.score}</span></td>
          </tr>`)
        .join("")
    : `<tr><td colspan="7" class="muted">אין נתונים זמינים.</td></tr>`;
}

function renderOwnerFocusTable() {
  const today = getLatestDate(state.rows) || new Date();
  const rows = buildOwnerFocusStats(state.filtered, today);

  el.ownerFocusTable.innerHTML = rows.length
    ? rows
        .map((item) => `
          <tr>
            <td><span class="badge">${escapeHtml(item.label)}</span></td>
            <td>${formatNumber(item.open)}</td>
            <td>${formatNumber(item.hot)}</td>
            <td>${formatNumber(item.stale)}</td>
            <td><span class="score ${item.averageScore >= 80 ? "hot" : ""}">${item.averageScore}</span></td>
            <td>${escapeHtml(item.conclusion)}</td>
          </tr>`)
        .join("")
    : `<tr><td colspan="6" class="muted">אין מספיק נתונים לבניית תמונת עומס צוותית.</td></tr>`;
}

function renderKeywordCloud() {
  const signals = extractConversationSignals(state.filtered);
  const max = Math.max(1, ...signals.map((item) => item.count));

  el.keywordCloud.innerHTML = signals.length
    ? signals
        .map((item) => {
          const width = Math.max(8, Math.round((item.count / max) * 100));
          return `
            <article class="signal-item">
              <div>
                <strong>${escapeHtml(item.label)}</strong>
                <span>${escapeHtml(item.recommendation)}</span>
              </div>
              <b>${formatNumber(item.count)}</b>
              <small><i style="width:${width}%"></i></small>
            </article>`;
        })
        .join("")
    : `<p class="muted">אין מספיק הערות כדי לזהות אותות שיחה.</p>`;
}

function renderStaleLeadsTable() {
  const today = getLatestDate(state.rows) || new Date();
  const stale = state.filtered
    .filter((row) => row._createdDate && isOpenLead(row))
    .map((row) => ({ row, age: daysBetween(row._createdDate, today) }))
    .filter((item) => item.age >= 3)
    .sort((a, b) => b.age - a.age)
    .slice(0, 15);

  el.staleLeadsTable.innerHTML = stale.length
    ? stale
        .map(({ row, age }) => `
          <tr>
            <td>${escapeHtml(row[columns.created] || "-")}</td>
            <td>${escapeHtml(row[columns.customerCode] || "-")}</td>
            <td>${escapeHtml(normalize(row[columns.leadStatus]))}</td>
            <td><span class="badge warn">${formatNumber(age)} ימים</span></td>
          </tr>`)
        .join("")
    : `<tr><td colspan="4" class="muted">אין לידים פתוחים מעל 3 ימים בסינון הנוכחי.</td></tr>`;
}

function renderExecutiveReport() {
  const total = state.filtered.length;
  if (!total) {
    el.executiveReport.innerHTML = `<article class="report-card">אין מספיק נתונים לבניית דוח מנהלים.</article>`;
    return;
  }

  const scores = state.filtered.map((row) => calculateLeadScore(row).score);
  const averageScore = Math.round(scores.reduce((sum, score) => sum + score, 0) / scores.length);
  const hotLeads = state.filtered.filter((row) => calculateLeadScore(row).score >= 75).length;
  const stale = getPriorityLeads(state.filtered).filter((item) => item.age >= 3).length;
  const sourceStats = buildGroupStats(state.filtered, columns.source).filter((group) => group.total >= 2);
  const bestSource = [...sourceStats].sort((a, b) => b.score - a.score || b.total - a.total)[0];
  const weakSource = [...sourceStats].sort((a, b) => a.score - b.score || b.total - a.total)[0];
  const invalidCustomerCodes = state.filtered.filter((row) => !isValidCustomerCode(row[columns.customerCode])).length;
  const closeReasons = Object.entries(countCloseReasonsForIrrelevant(state.filtered)).sort((a, b) => b[1] - a[1]);
  const mainBlocker = closeReasons[0];

  const cards = [
    ["בריאות כללית", `ציון איכות ממוצע ${averageScore}/100. יש ${formatNumber(hotLeads)} לידים חמים מתוך ${formatNumber(total)}.`],
    ["מיקוד מכירה", stale ? `${formatNumber(stale)} לידים פתוחים דורשים חזרה מהירה.` : "אין עומס חריג של לידים פתוחים לפי הסינון הנוכחי."],
    ["מקור להשקעה", bestSource ? `${bestSource.label} מוביל באיכות עם ציון ${bestSource.score}/100.` : "אין מספיק נתונים להשוואת מקורות."],
    ["מקור לבדיקה", weakSource && weakSource.label !== bestSource?.label ? `${weakSource.label} חלש יחסית, עם ציון ${weakSource.score}/100.` : "אין מקור חלש מובהק בסינון הנוכחי."],
    ["חסם מרכזי", mainBlocker ? `${mainBlocker[0]} היא סיבת הסגירה הבולטת ללא רלוונטיים.` : "אין חסם מרכזי מובהק מתוך סיבות הסגירה."],
    ["איכות נתונים", invalidCustomerCodes ? `${formatNumber(invalidCustomerCodes)} קודי לקוח נראים לא תקינים ודורשים בדיקה.` : "לא זוהתה בעיית קודי לקוח משמעותית."],
  ];

  el.executiveReport.innerHTML = cards
    .map(([title, body]) => `<article class="report-card"><strong>${escapeHtml(title)}</strong><span>${escapeHtml(body)}</span></article>`)
    .join("");
}

function renderActionPlan() {
  const actions = buildActionPlan();
  el.actionPlan.innerHTML = actions.length
    ? actions
        .map((action, index) => `
          <article class="action-item">
            <strong>${index + 1}</strong>
            <span>${escapeHtml(action)}</span>
          </article>`)
        .join("")
    : `<article class="action-item"><strong>1</strong><span>אין מספיק נתונים לבניית תוכנית פעולה.</span></article>`;
}

function buildActionPlan() {
  const total = state.filtered.length;
  if (!total) return [];

  const plan = [];
  const priority = getPriorityLeads(state.filtered);
  const qualityStats = getDataQualityStats(state.filtered);
  const interestRate = Math.round((state.filtered.filter(isInterested).length / total) * 100);
  const quoteRate = Math.round((state.filtered.filter(hasQuote).length / total) * 100);
  const irrelevantRate = Math.round((state.filtered.filter(isIrrelevant).length / total) * 100);
  const targetInterest = Number(el.targetInterest.value || 0);
  const targetQuote = Number(el.targetQuote.value || 0);
  const targetIrrelevant = Number(el.targetIrrelevant.value || 0);
  const sourceStats = buildGroupStats(state.filtered, columns.source).filter((group) => group.total >= 2);
  const weakSource = [...sourceStats].sort((a, b) => a.score - b.score || b.total - a.total)[0];
  const bestSource = [...sourceStats].sort((a, b) => b.score - a.score || b.total - a.total)[0];

  if (priority.length) {
    plan.push(`לטפל קודם ב-${formatNumber(Math.min(priority.length, 10))} הלידים הראשונים בטבלת "לידים לטיפול מיידי".`);
  }
  if (interestRate < targetInterest) {
    plan.push(`יחס העניין נמוך מהיעד ב-${targetInterest - interestRate}%. כדאי לבדוק מסרים ומקורות חלשים.`);
  }
  if (quoteRate < targetQuote) {
    plan.push(`יחס הצעות המחיר נמוך מהיעד ב-${targetQuote - quoteRate}%. מומלץ להגדיר קריטריון ברור למעבר להצעה.`);
  }
  if (irrelevantRate > targetIrrelevant) {
    plan.push(`אחוז הלא רלוונטיים גבוה מהיעד ב-${irrelevantRate - targetIrrelevant}%. כדאי לשפר סינון מוקדם בקמפיינים.`);
  }
  if (qualityStats.health < 85) {
    plan.push(`ציון איכות הדאטה הוא ${qualityStats.health}/100. להתחיל מניקוי טלפונים, מקורות חסרים וכפילויות.`);
  }
  if (weakSource && bestSource && weakSource.label !== bestSource.label) {
    plan.push(`להשוות בין ${bestSource.label} לבין ${weakSource.label}: מקור אחד חזק והשני דורש בדיקה.`);
  }
  if (!plan.length) {
    plan.push("המדדים המרכזיים עומדים ביעד. מומלץ להגדיל נפח לידים מהמקורות החזקים.");
  }

  return plan.slice(0, 6);
}

function renderScoreBuckets() {
  const buckets = {
    "חם מאוד 80-100": 0,
    "טוב 60-79": 0,
    "בינוני 40-59": 0,
    "נמוך 0-39": 0,
  };

  state.filtered.forEach((row) => {
    const score = calculateLeadScore(row).score;
    if (score >= 80) buckets["חם מאוד 80-100"] += 1;
    else if (score >= 60) buckets["טוב 60-79"] += 1;
    else if (score >= 40) buckets["בינוני 40-59"] += 1;
    else buckets["נמוך 0-39"] += 1;
  });

  renderBarChart(el.scoreBucketsChart, buckets);
}

function renderGoalTracker() {
  const total = state.filtered.length;
  const interestRate = total ? Math.round((state.filtered.filter(isInterested).length / total) * 100) : 0;
  const quoteRate = total ? Math.round((state.filtered.filter(hasQuote).length / total) * 100) : 0;
  const irrelevantRate = total ? Math.round((state.filtered.filter(isIrrelevant).length / total) * 100) : 0;
  const targets = [
    {
      label: "יחס עניין",
      current: interestRate,
      target: Number(el.targetInterest.value || 0),
      type: "minimum",
    },
    {
      label: "הצעות מחיר",
      current: quoteRate,
      target: Number(el.targetQuote.value || 0),
      type: "minimum",
    },
    {
      label: "לא רלוונטיים",
      current: irrelevantRate,
      target: Number(el.targetIrrelevant.value || 0),
      type: "maximum",
    },
  ];

  el.goalTracker.innerHTML = targets
    .map((item) => {
      const success = item.type === "minimum" ? item.current >= item.target : item.current <= item.target;
      const progress = item.type === "minimum"
        ? Math.min(100, item.target ? Math.round((item.current / item.target) * 100) : 100)
        : Math.min(100, item.current ? Math.round((item.target / item.current) * 100) : 100);
      const gap = item.type === "minimum" ? item.current - item.target : item.target - item.current;
      return `
        <article class="goal-card ${success ? "ok" : "warn"}">
          <div>
            <strong>${escapeHtml(item.label)}</strong>
            <span>${item.current}% מול יעד ${item.target}%</span>
          </div>
          <div class="goal-progress"><b style="width:${Math.max(4, progress)}%"></b></div>
          <small>${success ? "עומד ביעד" : `פער של ${Math.abs(gap)}%`}</small>
        </article>`;
    })
    .join("");
}

function renderDataQualityCards() {
  const stats = getDataQualityStats(state.filtered);
  const cards = [
    ["ציון איכות דאטה", `${stats.health}/100`, stats.health >= 85 ? "ok" : "warn"],
    ["קודי לקוח לא תקינים", formatNumber(stats.invalidCustomerCodes), stats.invalidCustomerCodes ? "warn" : "ok"],
    ["מקור חסר", formatNumber(stats.missingSource), stats.missingSource ? "warn" : "ok"],
    ["מנהל חסר", formatNumber(stats.missingOwner), stats.missingOwner ? "warn" : "ok"],
    ["סטטוס חסר", formatNumber(stats.missingStatus), stats.missingStatus ? "warn" : "ok"],
    ["תאריך לא תקין", formatNumber(stats.invalidDates), stats.invalidDates ? "warn" : "ok"],
  ];

  el.dataQualityCards.innerHTML = cards
    .map(([label, value, stateName]) => `
      <article class="quality-card ${stateName}">
        <span>${escapeHtml(label)}</span>
        <strong>${escapeHtml(value)}</strong>
      </article>`)
        .join("");
}

function renderAgingBuckets() {
  const today = getLatestDate(state.rows) || new Date();
  const buckets = getAgingBucketStats(state.filtered, today);

  el.agingBuckets.innerHTML = buckets.length
    ? buckets
        .map((item) => `
          <article class="aging-card ${item.daysFrom >= 4 ? "warn" : "ok"}">
            <span>${escapeHtml(item.label)}</span>
            <strong>${formatNumber(item.total)}</strong>
            <small>${escapeHtml(item.detail)}</small>
          </article>`)
        .join("")
    : `<article class="aging-card"><span>אין נתונים</span><strong>0</strong><small>אין תאריכים תקינים לחישוב גיל לידים.</small></article>`;
}

function renderPotentialQualityTable() {
  const groups = new Map();
  state.filtered.forEach((row) => {
    const key = normalize(row[columns.potential]);
    if (!groups.has(key)) {
      groups.set(key, { total: 0, interested: 0, irrelevant: 0, other: 0, quotes: 0, scoreSum: 0 });
    }
    const group = groups.get(key);
    group.total += 1;
    group.interested += isInterested(row) ? 1 : 0;
    group.irrelevant += isIrrelevant(row) ? 1 : 0;
    group.other += isOtherProcessWithoutCloseReason(row) ? 1 : 0;
    group.quotes += hasQuote(row) ? 1 : 0;
    group.scoreSum += calculateLeadScore(row).score;
  });

  el.potentialQualityTable.innerHTML = [...groups.entries()]
    .sort((a, b) => (b[1].scoreSum / b[1].total) - (a[1].scoreSum / a[1].total))
    .map(([label, group]) => `
      <tr>
        <td><span class="badge">${escapeHtml(label)}</span></td>
        <td>${formatNumber(group.total)}</td>
        <td>${formatNumber(group.interested)}</td>
        <td>${formatNumber(group.irrelevant)}</td>
        <td>${formatNumber(group.other)}</td>
        <td>${formatNumber(group.quotes)}</td>
        <td><span class="score">${Math.round(group.scoreSum / group.total)}</span></td>
      </tr>`)
    .join("");
}

function renderPriorityLeadsTable() {
  const priority = getPriorityLeads(state.filtered).slice(0, 18);
  el.priorityLeadsTable.innerHTML = priority.length
    ? priority
        .map(({ row, score, reasons, action }) => `
          <tr>
            <td><span class="score ${score >= 80 ? "hot" : ""}">${score}</span></td>
            <td>${escapeHtml(action)}</td>
            <td>${escapeHtml(row[columns.created] || "-")}</td>
            <td>${escapeHtml(row[columns.customerCode] || "-")}</td>
            <td>${escapeHtml(normalize(row[columns.owner]))}</td>
            <td>${escapeHtml(normalize(row[columns.source]))}</td>
            <td>${escapeHtml(normalize(row[columns.leadStatus]))}</td>
            <td>${escapeHtml(reasons.join(", "))}</td>
          </tr>`)
        .join("")
    : `<tr><td colspan="8" class="muted">אין לידים פתוחים או חמים לטיפול מיידי בסינון הנוכחי.</td></tr>`;
}

function renderGroupTable() {
  const column = el.groupBy.value || columns.leadStatus;
  const groups = new Map();

  state.filtered.forEach((row) => {
    const key = normalize(row[column]);
    if (!groups.has(key)) {
      groups.set(key, { total: 0, interested: 0, irrelevant: 0, quotes: 0 });
    }
    const group = groups.get(key);
    group.total += 1;
    group.interested += isInterested(row) ? 1 : 0;
    group.irrelevant += isIrrelevant(row) ? 1 : 0;
    group.quotes += hasQuote(row) ? 1 : 0;
  });

  el.groupTable.innerHTML = [...groups.entries()]
    .sort((a, b) => b[1].total - a[1].total)
    .map(([label, group]) => {
      const ratio = group.total ? Math.round((group.interested / group.total) * 100) : 0;
      return `
        <tr>
          <td><span class="badge">${escapeHtml(label)}</span></td>
          <td>${formatNumber(group.total)}</td>
          <td>${formatNumber(group.interested)}</td>
          <td>${formatNumber(group.irrelevant)}</td>
          <td>${formatNumber(group.quotes)}</td>
          <td>${ratio}%</td>
        </tr>`;
    })
    .join("");
}

function buildGroupStats(rows, column) {
  const groups = new Map();
  rows.forEach((row) => {
    const key = normalize(row[column]);
    if (!groups.has(key)) {
      groups.set(key, { label: key, total: 0, interested: 0, irrelevant: 0, quotes: 0 });
    }
    const group = groups.get(key);
    group.total += 1;
    group.interested += isInterested(row) ? 1 : 0;
    group.irrelevant += isIrrelevant(row) ? 1 : 0;
    group.quotes += hasQuote(row) ? 1 : 0;
  });

  return [...groups.values()].map((group) => {
    const interestRate = group.total ? Math.round((group.interested / group.total) * 100) : 0;
    const irrelevantRate = group.total ? Math.round((group.irrelevant / group.total) * 100) : 0;
    const quoteRate = group.total ? Math.round((group.quotes / group.total) * 100) : 0;
    return {
      ...group,
      interestRate,
      irrelevantRate,
      quoteRate,
      score: Math.max(0, Math.min(100, Math.round(interestRate * 0.7 + quoteRate * 0.4 - irrelevantRate * 0.35))),
    };
  });
}

function buildSourceRecommendations(rows) {
  return buildGroupStats(rows, columns.source)
    .filter((group) => group.total >= 3)
    .sort((a, b) => b.total - a.total)
    .map((group) => {
      let highlight = `יחס עניין ${group.interestRate}% והצעות ${group.quoteRate}%.`;
      let recommendation = "להמשיך להזרים נפח ולבדוק אם אפשר להגדיל תקציב או חשיפה.";

      if (group.irrelevantRate >= 45) {
        highlight = `אחוז לא רלוונטיים גבוה (${group.irrelevantRate}%).`;
        recommendation = "לחדד קהל יעד, טפסים וסינון ראשוני כדי להפחית רעש.";
      } else if (group.quoteRate <= 8 && group.interestRate >= 30) {
        highlight = `יש עניין, אבל מעט הצעות מחיר (${group.quoteRate}%).`;
        recommendation = "לבדוק האם הלידים נעצרים לפני שלב הצעת המחיר ולחזק המשך טיפול.";
      } else if (group.score >= 55) {
        highlight = `המקור שומר על איכות יציבה עם ציון ${group.score}.`;
        recommendation = "להגדיל השקעה בזהירות ולשכפל מסרים שעובדים.";
      }

      return {
        ...group,
        highlight,
        recommendation,
      };
    });
}

function buildOwnerFocusStats(rows, today) {
  return buildGroupStats(rows, columns.owner)
    .filter((group) => group.total >= 2)
    .map((group) => {
      const ownerRows = rows.filter((row) => normalize(row[columns.owner]) === group.label);
      const open = ownerRows.filter(isOpenLead).length;
      const hot = ownerRows.filter((row) => calculateLeadScore(row).score >= 75).length;
      const stale = ownerRows.filter((row) => row._createdDate && isOpenLead(row) && daysBetween(row._createdDate, today) >= 3).length;
      const averageScore = getAverageLeadScore(ownerRows);

      let conclusion = "איזון סביר בין עומס לאיכות.";
      if (stale >= 3) {
        conclusion = "יש הצטברות לידים תקועים, כדאי לפרוק עומס או להדק מעקב.";
      } else if (hot >= 3) {
        conclusion = "מחזיק כמה לידים חמים, כדאי לשמור לו חלון טיפול ממוקד.";
      } else if (open >= Math.max(6, hot + 4)) {
        conclusion = "עומס פתוח יחסית גבוה לעומת ההתקדמות, שווה לבדוק קצב טיפול.";
      }

      return {
        label: group.label,
        open,
        hot,
        stale,
        averageScore,
        conclusion,
      };
    })
    .sort((a, b) => b.stale - a.stale || b.hot - a.hot || b.open - a.open);
}

function getDuplicateCustomerCodes(rows) {
  const codes = new Map();
  rows.forEach((row) => {
    const code = clean(row[columns.customerCode]).trim();
    if (!code) return;
    if (!codes.has(code)) {
      codes.set(code, { code, count: 0, owners: new Set() });
    }
    const item = codes.get(code);
    item.count += 1;
    item.owners.add(normalize(row[columns.owner]));
  });

  return [...codes.values()]
    .filter((item) => item.count > 1)
    .map((item) => ({ ...item, owners: [...item.owners].slice(0, 4) }))
    .sort((a, b) => b.count - a.count || a.code.localeCompare(b.code, "he"));
}

function topKeys(counts, limit) {
  return Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, limit)
    .map(([key]) => key);
}

function extractKeywords(rows) {
  const stopWords = new Set([
    "לא", "כן", "את", "של", "עם", "על", "זה", "הוא", "היא", "אני", "אנחנו", "הם", "הן",
    "אין", "יש", "היה", "הייתה", "יותר", "פחות", "גם", "כל", "או", "אם", "כי", "אל",
    "לו", "לה", "לי", "בו", "בה", "מה", "מי", "אז", "רק", "כבר", "מאוד", "צריך",
    "אמר", "אמרה", "אמרו", "ביקש", "ביקשה", "שוב", "ליד", "לקוח", "לקוחה",
  ]);
  const counts = {};

  rows.forEach((row) => {
    const text = clean(row[columns.notes])
      .replace(/[^\u0590-\u05FFa-zA-Z0-9\s]/g, " ")
      .toLowerCase();
    text.split(/\s+/).forEach((word) => {
      const value = word.trim();
      if (value.length < 3 || stopWords.has(value) || /^\d+$/.test(value)) return;
      counts[value] = (counts[value] || 0) + 1;
    });
  });

  return Object.entries(counts)
    .map(([word, count]) => ({ word, count }))
    .filter((item) => item.count >= 2)
    .sort((a, b) => b.count - a.count || a.word.localeCompare(b.word, "he"));
}

function extractConversationSignals(rows) {
  const definitions = [
    {
      label: "אין מענה / לא עונה",
      pattern: /אין מענה|לא עונה|לא ענו|ניסיון נוסף/,
      recommendation: "לתעדף לניסיון חזרה נוסף בשעות פעילות חזקות.",
    },
    {
      label: "ביקש הצעת מחיר",
      pattern: /הצעת מחיר|הצעה|מחיר|הצעת/,
      recommendation: "לוודא שנשלחה הצעה ולעשות follow-up קצר.",
    },
    {
      label: "נקבעה פגישה / הדגמה",
      pattern: /פגישה|הדגמה|דמו|שיחה עם מנהל/,
      recommendation: "להכין סיכום צורך לפני הפגישה ולסגור שלב הבא.",
    },
    {
      label: "ממתין לאישור",
      pattern: /ממתין|ממתינה|אישור|שותף|מנהל/,
      recommendation: "לשלוח תזכורת עם ערך ברור וסיבה להתקדם.",
    },
    {
      label: "בעיית תקציב",
      pattern: /תקציב|יקר|עלות|מחיר גבוה|לא מתאים תקציבית/,
      recommendation: "להציע מסלול כניסה או להסביר ROI קצר.",
    },
    {
      label: "קוד לקוח שגוי",
      pattern: /קוד לקוח שגוי|מספר לא|לא מחובר/,
      recommendation: "לסמן לניקוי דאטה ולא להשאיר ברשימת follow-up.",
    },
    {
      label: "מחפש עבודה",
      pattern: /מחפש עבודה|משרה|קורות חיים/,
      recommendation: "לסווג כלא רלוונטי ולהחריג ממדדי מכירה.",
    },
    {
      label: "לקוח חם",
      pattern: /לקוח חם|מעוניין|רוצה|ביקש שיחזרו|חזרה עד סוף היום/,
      recommendation: "לתעדף לשיחה מהירה באותו יום.",
    },
  ];

  return definitions
    .map((definition) => {
      const count = rows.filter((row) => {
        const text = `${normalize(row[columns.notes])} ${normalize(row[columns.leadStatus])} ${normalize(row[columns.saleStatus])} ${normalize(row[columns.closeReason])}`;
        return definition.pattern.test(text);
      }).length;
      return { ...definition, count };
    })
    .filter((item) => item.count > 0)
    .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, "he"));
}

function isOpenLead(row) {
  if (isIrrelevant(row)) return false;
  const statusText = `${normalize(row[columns.leadStatus])} ${normalize(row[columns.saleStatus])} ${normalize(row[columns.closeReason])}`;
  return /מעוניין|המשך טיפול|מעקב|פגישה|אין מענה|לחזור|ממתין|בטיפול/.test(statusText);
}

function getLatestDate(rows) {
  const timestamps = rows.map((row) => row._createdDate?.getTime()).filter(Boolean);
  return timestamps.length ? new Date(Math.max(...timestamps)) : null;
}

function getBestHour(rows) {
  const counts = Array.from({ length: 24 }, (_, hour) => ({ hour, count: 0 }));
  rows.forEach((row) => {
    if (row._createdDate) counts[row._createdDate.getHours()].count += 1;
  });
  const best = counts.sort((a, b) => b.count - a.count)[0];
  return best?.count ? best : null;
}

function calculateLeadScore(row) {
  const reasons = [];
  let score = 45;

  if (isInterested(row)) {
    score += 24;
    reasons.push("עניין/טיפול");
  }

  if (hasQuote(row)) {
    score += 14;
    reasons.push("יש הצעת מחיר");
  }

  const potential = normalize(row[columns.potential]);
  if (/basic|premium|גבוה|חם/i.test(potential)) {
    score += 10;
    reasons.push("פוטנציאל חיובי");
  } else if (/לא רלוונטי/.test(potential)) {
    score -= 28;
    reasons.push("פוטנציאל לא רלוונטי");
  } else if (/לא ידוע/.test(potential)) {
    score -= 4;
  }

  if (isIrrelevant(row)) {
    score -= 38;
    reasons.push("לא רלוונטי");
  }

  const text = `${normalize(row[columns.leadStatus])} ${normalize(row[columns.saleStatus])} ${normalize(row[columns.notes])}`;
  if (/פגישה|מעוניין|הצעה|מייל|שלחתי|ממתין|אתר/.test(text)) {
    score += 8;
    reasons.push("אות חיובי בהערות");
  }
  if (/אין מענה|נעלם|קוד לקוח שגוי|מחפש עבודה|לא עונה/.test(text)) {
    score -= 12;
    reasons.push("סיכון בהערות");
  }

  if (!isValidCustomerCode(row[columns.customerCode])) {
    score -= 18;
    reasons.push("קוד לקוח לבדיקה");
  }

  if (!row._createdDate) {
    score -= 6;
    reasons.push("תאריך חסר");
  }

  if (normalize(row[columns.source]) === "לא צוין") {
    score -= 5;
  }

  return {
    score: Math.max(0, Math.min(100, Math.round(score))),
    reasons: reasons.length ? reasons : ["ליד רגיל"],
  };
}

function getPriorityLeads(rows) {
  const latest = getLatestDate(state.rows) || new Date();
  return rows
    .filter((row) => !isIrrelevant(row) && (isOpenLead(row) || isInterested(row) || hasQuote(row)))
    .map((row) => {
      const quality = calculateLeadScore(row);
      const age = row._createdDate ? daysBetween(row._createdDate, latest) : 0;
      const urgency = Math.min(24, age * 4) + (hasQuote(row) ? 10 : 0) + (isOpenLead(row) ? 8 : 0);
      return {
        row,
        age,
        score: quality.score,
        priority: quality.score + urgency,
        reasons: [...quality.reasons, age ? `${age} ימים` : "חדש"],
        action: getRecommendedAction(row, age, quality.score),
      };
    })
    .sort((a, b) => b.priority - a.priority || b.score - a.score);
}

function getRecommendedAction(row, age, score) {
  if (hasQuote(row)) return "לחזור על הצעת מחיר";
  if (age >= 5) return "חזרה דחופה היום";
  if (score >= 80) return "לתעדף לשיחה";
  if (/אין מענה|לא עונה/.test(normalize(row[columns.notes]))) return "ניסיון התקשרות נוסף";
  if (/ממתין|שלחתי|מייל/.test(normalize(row[columns.notes]))) return "מעקב אחרי מייל";
  return "להמשיך טיפול";
}

function isValidCustomerCode(value) {
  const cleaned = clean(value).replace(/\s+/g, "");
  return cleaned.length >= 3;
}

function isMissingValue(value) {
  const normalized = normalize(value);
  return normalized === "לא צוין";
}

function getDataQualityStats(rows) {
  const total = Math.max(1, rows.length);
  const duplicateExtras = getDuplicateCustomerCodes(rows).reduce((sum, item) => sum + item.count - 1, 0);
  const stats = rows.reduce((acc, row) => {
    acc.invalidCustomerCodes += isValidCustomerCode(row[columns.customerCode]) ? 0 : 1;
    acc.missingSource += isMissingValue(row[columns.source]) ? 1 : 0;
    acc.missingOwner += isMissingValue(row[columns.owner]) ? 1 : 0;
    acc.missingStatus += isMissingValue(row[columns.leadStatus]) ? 1 : 0;
    acc.invalidDates += row._createdDate ? 0 : 1;
    return acc;
  }, {
    invalidCustomerCodes: 0,
    missingSource: 0,
    missingOwner: 0,
    missingStatus: 0,
    invalidDates: 0,
  });

  stats.duplicateExtras = duplicateExtras;
  stats.totalIssues = stats.invalidCustomerCodes + stats.missingSource + stats.missingOwner + stats.missingStatus + stats.invalidDates + duplicateExtras;
  stats.health = Math.max(0, Math.min(100, 100 - Math.round((stats.totalIssues / (total * 2)) * 100)));
  return stats;
}

function getAverageLeadScore(rows) {
  if (!rows.length) return 0;
  const total = rows.reduce((sum, row) => sum + calculateLeadScore(row).score, 0);
  return Math.round(total / rows.length);
}

function getAgingBucketStats(rows, today) {
  const buckets = [
    { label: "0-1 ימים", daysFrom: 0, daysTo: 1, total: 0, open: 0, interested: 0 },
    { label: "2-3 ימים", daysFrom: 2, daysTo: 3, total: 0, open: 0, interested: 0 },
    { label: "4-7 ימים", daysFrom: 4, daysTo: 7, total: 0, open: 0, interested: 0 },
    { label: "8+ ימים", daysFrom: 8, daysTo: Infinity, total: 0, open: 0, interested: 0 },
  ];

  rows.forEach((row) => {
    if (!row._createdDate) return;
    const age = daysBetween(row._createdDate, today);
    const bucket = buckets.find((item) => age >= item.daysFrom && age <= item.daysTo);
    if (!bucket) return;
    bucket.total += 1;
    bucket.open += isOpenLead(row) ? 1 : 0;
    bucket.interested += isInterested(row) ? 1 : 0;
  });

  return buckets.map((bucket) => ({
    ...bucket,
    detail: bucket.total
      ? `${formatNumber(bucket.open)} פתוחים, ${formatNumber(bucket.interested)} מעוניינים / בטיפול`
      : "אין לידים בקבוצת הזמן הזו",
  }));
}

function getMainRiskText() {
  const total = state.filtered.length;
  if (!total) return "אין נתונים מספקים לזיהוי סיכון מרכזי.";

  const qualityStats = getDataQualityStats(state.filtered);
  const closeReasons = Object.entries(countCloseReasonsForIrrelevant(state.filtered)).sort((a, b) => b[1] - a[1]);
  const stale = getPriorityLeads(state.filtered).filter((item) => item.age >= 3).length;

  if (stale >= 6) {
    return `${formatNumber(stale)} לידים דחופים כבר בני 3 ימים ומעלה.`;
  }

  if (qualityStats.invalidCustomerCodes >= 6) {
    return `${formatNumber(qualityStats.invalidCustomerCodes)} קודי לקוח לא תקינים פוגעים ביכולת המעקב.`;
  }

  if (closeReasons[0]) {
    return `${closeReasons[0][0]} היא סיבת הסגירה המובילה בקרב לא רלוונטיים.`;
  }

  return "לא זוהה כרגע סיכון חריג אחד שמוביל על השאר.";
}

function daysBetween(start, end) {
  const ms = new Date(end.getFullYear(), end.getMonth(), end.getDate()) -
    new Date(start.getFullYear(), start.getMonth(), start.getDate());
  return Math.max(0, Math.round(ms / 86400000));
}

function renderLeadsTable() {
  const visibleHeaders = [
    columns.customerCode,
    columns.created,
    columns.owner,
    columns.source,
    columns.potential,
    columns.leadStatus,
    columns.saleStatus,
    columns.closeReason,
    columns.notes,
  ].filter((header) => state.headers.includes(header));

  const head = `<thead><tr><th>ציון איכות</th>${visibleHeaders.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>`;
  const body = state.filtered
    .slice(0, 120)
    .map((row) => {
      const score = calculateLeadScore(row).score;
      return `<tr><td><span class="score ${score >= 80 ? "hot" : ""}">${score}</span></td>${visibleHeaders.map((header) => `<td>${escapeHtml(row[header] || "-")}</td>`).join("")}</tr>`;
    })
    .join("");

  el.leadsTable.innerHTML = `${head}<tbody>${body}</tbody>`;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"' && inQuotes && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === "," && !inQuotes) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell);
      if (row.some((value) => value.trim() !== "")) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  row.push(cell);
  if (row.some((value) => value.trim() !== "")) rows.push(row);
  return rows;
}

function fillSelect(select, values) {
  select.innerHTML = [
    `<option value="__all">הכל</option>`,
    ...values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`),
  ].join("");
}

function uniqueValues(rows, column) {
  return [...new Set(rows.map((row) => normalize(row[column])))]
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, "he"));
}

function countBy(rows, column) {
  return rows.reduce((acc, row) => {
    const key = normalize(row[column]);
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function countCloseReasonsForIrrelevant(rows) {
  return rows.reduce((acc, row) => {
    if (!isIrrelevant(row) || !hasRealCloseReason(row)) return acc;
    const key = normalize(row[columns.closeReason]);
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function countByDate(rows) {
  return rows.reduce((acc, row) => {
    if (!row._createdDate) return acc;
    const key = toDateInput(row._createdDate);
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function countByHour(rows) {
  const counts = Array.from({ length: 24 }, (_, hour) => ({ hour, count: 0 }));
  rows.forEach((row) => {
    if (row._createdDate) counts[row._createdDate.getHours()].count += 1;
  });
  return counts;
}

function countByWeekday(rows) {
  const names = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"];
  const counts = names.map((name, index) => ({ name, index, value: 0, interested: 0 }));
  rows.forEach((row) => {
    if (!row._createdDate) return;
    const day = row._createdDate.getDay();
    counts[day].value += 1;
    counts[day].interested += isInterested(row) ? 1 : 0;
  });
  return counts;
}

function clean(value) {
  return String(value ?? "").replace(/^\uFEFF/, "").trim();
}

function normalize(value) {
  const cleaned = clean(value);
  return cleaned && cleaned !== "-" ? cleaned : "לא צוין";
}

function isInterested(row) {
  const lead = normalize(row[columns.leadStatus]);
  const sale = normalize(row[columns.saleStatus]);
  return /מעוניין|המשך טיפול|מעקב|פגישה|סגירה|נשלח|הצעה/.test(`${lead} ${sale}`);
}

function isIrrelevant(row) {
  return /לא רלוונטי|קוד לקוח שגוי|מחפש עבודה/.test(
    `${normalize(row[columns.leadStatus])} ${normalize(row[columns.closeReason])} ${normalize(row[columns.potential])}`
  );
}

function hasRealCloseReason(row) {
  const reason = normalize(row[columns.closeReason]);
  return reason !== "לא צוין";
}

function isOtherProcessWithoutCloseReason(row) {
  return !hasRealCloseReason(row);
}

function hasQuote(row) {
  return !["לא צוין", "לא", "-"].includes(normalize(row[columns.quoteSent])) ||
    !["לא צוין", "-"].includes(normalize(row[columns.quoteStatus]));
}

function parseHebrewDate(value) {
  const match = clean(value).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!match) return null;
  const [, day, month, year, hour = "0", minute = "0"] = match;
  return new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute));
}

function toDateInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatShortDate(dateKey) {
  const [, month, day] = dateKey.split("-");
  return `${day}/${month}`;
}

function formatDateLabel(dateKey) {
  const [year, month, day] = dateKey.split("-");
  return `${day}/${month}/${year}`;
}

function percent(value, total) {
  return total ? `${Math.round((value / total) * 100)}%` : "0%";
}

function exportSummary() {
  if (!state.filtered.length) {
    alert("אין נתונים לייצוא.");
    return;
  }

  const rows = [["פרמטר", "ערך", "כמות"]];
  const latest = getLatestDate(state.rows) || new Date();
  [
    ["סטטוס ליד", countBy(state.filtered, columns.leadStatus)],
    ["סטטוס מכירה", countBy(state.filtered, columns.saleStatus)],
    ["מקור הגעה", countBy(state.filtered, columns.source)],
    ["מנהל לקוח", countBy(state.filtered, columns.owner)],
    ["פוטנציאל לקוח", countBy(state.filtered, columns.potential)],
    ["סיבת סגירה ללא רלוונטיים", countCloseReasonsForIrrelevant(state.filtered)],
  ].forEach(([label, counts]) => {
    Object.entries(counts)
      .sort((a, b) => b[1] - a[1])
      .forEach(([value, count]) => rows.push([label, value, count]));
  });

  rows.push(["עברו תהליך אחר", "ללא סיבת סגירה", state.filtered.filter(isOtherProcessWithoutCloseReason).length]);

  rows.push([]);
  rows.push(["איכות מקור", "סה״כ", "מעוניינים", "לא רלוונטיים", "הצעות", "יחס עניין", "ציון איכות"]);
  buildGroupStats(state.filtered, columns.source)
    .sort((a, b) => b.score - a.score || b.total - a.total)
    .forEach((group) => rows.push([group.label, group.total, group.interested, group.irrelevant, group.quotes, `${group.interestRate}%`, group.score]));

  rows.push([]);
  rows.push(["מגמה יומית", "כמות"]);
  Object.entries(countByDate(state.filtered))
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([date, count]) => rows.push([formatDateLabel(date), count]));

  rows.push([]);
  rows.push(["קודי לקוח כפולים", "כמות", "מנהלי לקוח"]);
  getDuplicateCustomerCodes(state.filtered).forEach((item) => rows.push([item.code, item.count, item.owners.join(", ")]));

  rows.push([]);
  rows.push(["שעה", "כמות לידים"]);
  countByHour(state.filtered).forEach((item) => rows.push([`${String(item.hour).padStart(2, "0")}:00`, item.count]));

  rows.push([]);
  rows.push(["יום בשבוע", "כמות", "מעוניינים"]);
  countByWeekday(state.filtered).forEach((item) => rows.push([item.name, item.value, item.interested]));

  rows.push([]);
  rows.push(["מנהל לקוח", "סה״כ", "מעוניינים", "לא רלוונטיים", "הצעות", "ציון"]);
  buildGroupStats(state.filtered, columns.owner)
    .sort((a, b) => b.score - a.score || b.total - a.total)
    .forEach((group) => rows.push([group.label, group.total, group.interested, group.irrelevant, group.quotes, group.score]));

  rows.push([]);
  rows.push(["המלצות לפי מקור", "נפח", "ציון איכות", "מה בולט", "המלצה"]);
  buildSourceRecommendations(state.filtered).forEach((group) => rows.push([group.label, group.total, group.score, group.highlight, group.recommendation]));

  rows.push([]);
  rows.push(["פוקוס מנהלי לקוח", "פתוחים", "חמים", "תקועים", "ציון ממוצע", "מסקנה"]);
  buildOwnerFocusStats(state.filtered, latest).forEach((item) => rows.push([item.label, item.open, item.hot, item.stale, item.averageScore, item.conclusion]));

  rows.push([]);
  rows.push(["אותות מהערות", "כמות", "המלצה"]);
  extractConversationSignals(state.filtered).forEach((item) => rows.push([item.label, item.count, item.recommendation]));

  rows.push([]);
  rows.push(["לידים תקועים", "קוד לקוח", "תאריך", "סטטוס", "גיל בימים"]);
  state.filtered.filter((row) => row._createdDate && isOpenLead(row))
    .map((row) => ({ row, age: daysBetween(row._createdDate, latest) }))
    .filter((item) => item.age >= 3)
    .sort((a, b) => b.age - a.age)
    .forEach(({ row, age }) => rows.push(["ליד פתוח", row[columns.customerCode], row[columns.created], normalize(row[columns.leadStatus]), age]));

  rows.push([]);
  rows.push(["לידים לטיפול מיידי", "ציון", "פעולה מומלצת", "קוד לקוח", "תאריך", "מנהל", "מקור", "סטטוס", "סיבה לדירוג"]);
  getPriorityLeads(state.filtered).slice(0, 40).forEach(({ row, score, action, reasons }) => {
    rows.push(["לטיפול", score, action, row[columns.customerCode], row[columns.created], normalize(row[columns.owner]), normalize(row[columns.source]), normalize(row[columns.leadStatus]), reasons.join(", ")]);
  });

  rows.push([]);
  rows.push(["ציוני איכות לכל הלידים", "ציון", "קוד לקוח", "מנהל", "מקור", "סטטוס", "סיבות"]);
  state.filtered.forEach((row) => {
    const quality = calculateLeadScore(row);
    rows.push(["ליד", quality.score, row[columns.customerCode], normalize(row[columns.owner]), normalize(row[columns.source]), normalize(row[columns.leadStatus]), quality.reasons.join(", ")]);
  });

  rows.push([]);
  rows.push(["יעדים מול ביצועים", "ביצוע", "יעד"]);
  rows.push(["יחס עניין", `${percent(state.filtered.filter(isInterested).length, state.filtered.length)}`, `${Number(el.targetInterest.value || 0)}%`]);
  rows.push(["הצעות מחיר", `${percent(state.filtered.filter(hasQuote).length, state.filtered.length)}`, `${Number(el.targetQuote.value || 0)}%`]);
  rows.push(["לא רלוונטיים", `${percent(state.filtered.filter(isIrrelevant).length, state.filtered.length)}`, `${Number(el.targetIrrelevant.value || 0)}%`]);

  rows.push([]);
  const qualityStats = getDataQualityStats(state.filtered);
  rows.push(["איכות נתונים", "ערך"]);
  rows.push(["ציון איכות דאטה", qualityStats.health]);
  rows.push(["קודי לקוח לא תקינים", qualityStats.invalidCustomerCodes]);
  rows.push(["מקור חסר", qualityStats.missingSource]);
  rows.push(["מנהל חסר", qualityStats.missingOwner]);
  rows.push(["סטטוס חסר", qualityStats.missingStatus]);
  rows.push(["תאריך לא תקין", qualityStats.invalidDates]);
  rows.push(["כפילויות עודפות", qualityStats.duplicateExtras]);

  rows.push([]);
  rows.push(["גיל לידים", "סה״כ", "פתוחים", "מעוניינים"]);
  getAgingBucketStats(state.filtered, latest).forEach((item) => rows.push([item.label, item.total, item.open, item.interested]));

  rows.push([]);
  rows.push(["איכות לפי פוטנציאל", "סה״כ", "מעוניינים", "לא רלוונטיים", "תהליך אחר", "הצעות", "ציון ממוצע"]);
  const potentialGroups = new Map();
  state.filtered.forEach((row) => {
    const key = normalize(row[columns.potential]);
    if (!potentialGroups.has(key)) potentialGroups.set(key, { total: 0, interested: 0, irrelevant: 0, other: 0, quotes: 0, scoreSum: 0 });
    const group = potentialGroups.get(key);
    group.total += 1;
    group.interested += isInterested(row) ? 1 : 0;
    group.irrelevant += isIrrelevant(row) ? 1 : 0;
    group.other += isOtherProcessWithoutCloseReason(row) ? 1 : 0;
    group.quotes += hasQuote(row) ? 1 : 0;
    group.scoreSum += calculateLeadScore(row).score;
  });
  [...potentialGroups.entries()].forEach(([label, group]) => rows.push([label, group.total, group.interested, group.irrelevant, group.other, group.quotes, Math.round(group.scoreSum / group.total)]));

  downloadCsv(rows, "sales-statistics-summary.csv");
}

function exportPriorityLeads() {
  const priority = getPriorityLeads(state.filtered);
  if (!priority.length) {
    alert("אין לידים לטיפול מיידי בסינון הנוכחי.");
    return;
  }

  const rows = [[
    "ציון",
    "פעולה מומלצת",
    "עסק",
    "תאריך",
    "מנהל",
    "מקור",
    "סטטוס",
    "סיבה לדירוג",
  ]];

  priority.forEach(({ row, score, action, reasons }) => {
    rows.push([
      score,
      action,
      row[columns.customerCode],
      row[columns.created],
      normalize(row[columns.owner]),
      normalize(row[columns.source]),
      normalize(row[columns.leadStatus]),
      reasons.join(", "),
    ]);
  });

  downloadCsv(rows, "priority-leads.csv");
}

async function copyExecutiveReport() {
  const text = buildExecutiveReportText();
  if (!text) {
    alert("אין דוח להעתקה. הפעל ניתוח קודם.");
    return;
  }

  try {
    await navigator.clipboard.writeText(text);
    alert("דוח המנהלים הועתק ללוח.");
  } catch (error) {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    textarea.remove();
    alert("דוח המנהלים הועתק ללוח.");
  }
}

function buildExecutiveReportText() {
  if (!state.filtered.length) return "";
  const lines = [
    "דוח מנהלים - מערכת לניתוח סטטיסטי של מכירות",
    `סה״כ לידים: ${formatNumber(state.filtered.length)}`,
    `מעוניינים / בטיפול: ${formatNumber(state.filtered.filter(isInterested).length)}`,
    `לא רלוונטיים: ${formatNumber(state.filtered.filter(isIrrelevant).length)}`,
    `הצעות מחיר: ${formatNumber(state.filtered.filter(hasQuote).length)}`,
    "",
    "תוכנית פעולה:",
    ...buildActionPlan().map((item, index) => `${index + 1}. ${item}`),
  ];
  return lines.join("\n");
}

function downloadCsv(rows, filename) {
  const csv = rows.map((row) => row.map(csvEscape).join(",")).join("\n");
  const blob = new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
  URL.revokeObjectURL(link.href);
}

function csvEscape(value) {
  return `"${String(value).replaceAll('"', '""')}"`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatNumber(value) {
  return new Intl.NumberFormat("he-IL").format(value);
}


