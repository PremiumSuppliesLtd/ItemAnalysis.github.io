// ============================================
// ITEM ANALYSIS DASHBOARD - Main Application Logic
// ============================================

// ===== GLOBALS =====
let rawData = [];
let processedData = {};
let currentTab = 'overview';
let yearRangeFrom = null;
let yearRangeTo = null;

// Chart instances
let chartInstances = {};

// Reorder data
let reorderData = [];
let purchaseData = [];
let reorderProcessed = null;

// Month parsing utilities for "YYYY-M" format (e.g. "2025-11", "2024-1", "2022")
const _MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function parseMonthStr(str) {
    if (!str) return { year: 0, month: 1 };
    const s = str.toString().trim();
    const parts = s.split('-');
    const year = parseInt(parts[0]) || 0;
    const month = parts.length > 1 ? (parseInt(parts[1]) || 1) : 1;
    return { year, month };
}

function monthSortKey(str) {
    const { year, month } = parseMonthStr(str);
    return year * 100 + month;
}

function monthsDiffCalc(later, earlier) {
    const l = parseMonthStr(later);
    const e = parseMonthStr(earlier);
    return (l.year - e.year) * 12 + (l.month - e.month);
}

function formatMonthLabel(str) {
    const { year, month } = parseMonthStr(str);
    if (year && month >= 1 && month <= 12) {
        return _MONTH_NAMES[month - 1] + ' ' + year;
    }
    return str;
}

// Premium color palette
const COLORS = {
    primary: '#6366f1',
    secondary: '#8b5cf6',
    success: '#10b981',
    danger: '#ef4444',
    warning: '#f59e0b',
    info: '#3b82f6',
    rose: '#f43f5e',
    cyan: '#06b6d4',
    amber: '#fbbf24',
    emerald: '#34d399',
    purple: '#a78bfa',
    pink: '#f472b6',
    palette: [
        '#6366f1', '#8b5cf6', '#06b6d4', '#10b981', '#f59e0b',
        '#ef4444', '#f43f5e', '#3b82f6', '#a78bfa', '#f472b6',
        '#14b8a6', '#eab308', '#ec4899', '#0ea5e9', '#84cc16'
    ]
};

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    initNavigation();
    initSalesSubNav();
    initEventListeners();
    fetchData();
});

// ===== THEME MANAGEMENT =====
function initTheme() {
    const saved = localStorage.getItem('item_analysis_theme');
    if (saved === 'light') {
        document.body.classList.add('light-theme');
        document.getElementById('themeBtn').textContent = '🌙';
    }
}

function toggleTheme() {
    const isLight = document.body.classList.toggle('light-theme');
    document.getElementById('themeBtn').textContent = isLight ? '🌙' : '☀️';
    localStorage.setItem('item_analysis_theme', isLight ? 'light' : 'dark');
    // Re-render charts with new colors
    if (rawData.length > 0) {
        renderCurrentTab();
    }
}

// ===== NAVIGATION =====
function initNavigation() {
    document.querySelectorAll('.report-tab').forEach(tab => {
        tab.addEventListener('click', (e) => {
            const targetTab = e.currentTarget.getAttribute('data-tab');
            switchTab(targetTab);
        });
    });
}

function switchTab(tabName) {
    if (tabName === currentTab) return;
    currentTab = tabName;

    // Update tab buttons
    document.querySelectorAll('.report-tab').forEach(t => t.classList.remove('active'));
    document.querySelector(`.report-tab[data-tab="${tabName}"]`).classList.add('active');

    // Show/hide tab content
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active-tab'));
    const target = document.getElementById(`tab-${tabName}`);
    if (target) {
        target.classList.add('active-tab');
        // Trigger animations
        target.style.animation = 'none';
        target.offsetHeight;
        target.style.animation = '';
    }

    renderCurrentTab();
}

// ===== CSS for tab visibility =====
const tabStyle = document.createElement('style');
tabStyle.textContent = `
    .tab-content { display: none; }
    .tab-content.active-tab { display: block; animation: tabFadeIn 0.4s ease-out; }
    @keyframes tabFadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
`;
document.head.appendChild(tabStyle);

// ===== EVENT LISTENERS =====
function initEventListeners() {
    document.getElementById('themeBtn').addEventListener('click', toggleTheme);
    document.getElementById('refreshBtn').addEventListener('click', () => fetchData(true));

    // Global year range filter
    const yearFrom = document.getElementById('yearFrom');
    const yearTo = document.getElementById('yearTo');
    if (yearFrom) yearFrom.addEventListener('change', onYearRangeChange);
    if (yearTo) yearTo.addEventListener('change', onYearRangeChange);

    // Dropped tab filters
    const droppedSearch = document.getElementById('droppedCustomerSearch');
    const droppedMonths = document.getElementById('droppedMonthsFilter');
    const droppedSort = document.getElementById('droppedSortBy');
    if (droppedSearch) droppedSearch.addEventListener('input', debounce(renderDropped, 300));
    if (droppedMonths) droppedMonths.addEventListener('change', renderDropped);
    if (droppedSort) droppedSort.addEventListener('change', renderDropped);

    // Heatmap filters
    const heatCust = document.getElementById('heatmapCustomerSearch');
    const heatItem = document.getElementById('heatmapItemSearch');
    const heatLimit = document.getElementById('heatmapLimit');
    if (heatCust) heatCust.addEventListener('input', debounce(renderHeatmap, 300));
    if (heatItem) heatItem.addEventListener('input', debounce(renderHeatmap, 300));
    if (heatLimit) heatLimit.addEventListener('change', renderHeatmap);

    // Declining filters
    const declSearch = document.getElementById('decliningCustomerSearch');
    const declThreshold = document.getElementById('decliningThreshold');
    if (declSearch) declSearch.addEventListener('input', debounce(renderDeclining, 300));
    if (declThreshold) declThreshold.addEventListener('change', renderDeclining);

    // Top Items filters
    const topCust = document.getElementById('topItemsCustomerSelect');
    const topLimit = document.getElementById('topItemsLimit');
    if (topCust) topCust.addEventListener('change', renderTopItems);
    if (topLimit) topLimit.addEventListener('change', renderTopItems);

    // Trends filters
    const trendCust = document.getElementById('trendCustomerSelect');
    const trendItem = document.getElementById('trendItemSelect');
    if (trendCust) trendCust.addEventListener('change', renderTrends);
    if (trendItem) trendItem.addEventListener('change', renderTrends);

    // Deep Dive
    const deepBtn = document.getElementById('deepDiveBtn');
    const deepSearch = document.getElementById('deepDiveItemSearch');
    if (deepBtn) deepBtn.addEventListener('click', renderDeepDive);
    if (deepSearch) deepSearch.addEventListener('keypress', (e) => { if (e.key === 'Enter') renderDeepDive(); });

    // Reorder filters
    const reorderSearch = document.getElementById('reorderItemSearch');
    const reorderSupplier = document.getElementById('reorderSupplierFilter');
    const reorderUrgency = document.getElementById('reorderUrgencyFilter');
    const reorderView = document.getElementById('reorderViewMode');
    if (reorderSearch) reorderSearch.addEventListener('input', debounce(renderReorderTable, 300));
    if (reorderSupplier) reorderSupplier.addEventListener('change', renderReorderTable);
    if (reorderUrgency) reorderUrgency.addEventListener('change', renderReorderTable);
    if (reorderView) reorderView.addEventListener('change', renderReorderTable);

    // Sales Intel filters
    const priceSearch = document.getElementById('priceSearchItem');
    const priceSort = document.getElementById('priceSortBy');
    const salesCustSelect = document.getElementById('salesCustSelect');
    if (priceSearch) priceSearch.addEventListener('input', debounce(renderSalesPriceAnalytics, 300));
    if (priceSort) priceSort.addEventListener('change', renderSalesPriceAnalytics);
    if (salesCustSelect) salesCustSelect.addEventListener('change', renderSalesCustomerRevenue);

    // Stock Builder search
    const stockSearch = document.getElementById('stockBuilderSearch');
    if (stockSearch) stockSearch.addEventListener('input', debounce(renderStockBuilder, 400));
}

// ===== DATA FETCHING =====
async function fetchData(forceRefresh = false) {
    const loader = document.getElementById('globalLoader');

    try {
        loader.classList.remove('hidden');

        const url = typeof ITEM_ANALYSIS_URL !== 'undefined' ? ITEM_ANALYSIS_URL.trim() : '';
        if (!url || url === 'PASTE_YOUR_ITEM_ANALYSIS_SCRIPT_URL_HERE') {
            showEmptyState('Please configure your Google Apps Script URL in index.html');
            return;
        }

        const response = await fetch(url);
        if (!response.ok) throw new Error('Network error: ' + response.status);

        const result = await response.json();
        if (result.status === 'error') throw new Error(result.message);

        rawData = result.data || [];
        if (rawData.length === 0) {
            showEmptyState('No data found in the Item Analysis sheet');
            return;
        }

        populateYearDropdowns();
        processData();
        populateDropdowns();
        renderCurrentTab();

    } catch (err) {
        console.error('Fetch error:', err);
        showEmptyState('Error: ' + err.message);
    } finally {
        setTimeout(() => {
            loader.classList.add('hidden');
        }, 500);
    }

    // Fetch reorder data in parallel (non-blocking)
    fetchReorderData();
    // Fetch sales data in parallel (non-blocking)
    fetchSalesData();
}

function showEmptyState(msg) {
    const loader = document.getElementById('globalLoader');
    loader.classList.add('hidden');
    document.querySelectorAll('.empty-msg').forEach(el => el.textContent = msg);
}

// ===== YEAR RANGE FILTER =====
function populateYearDropdowns() {
    if (rawData.length === 0) return;
    const headers = Object.keys(rawData[0]);
    const monthKey = headers.find(h => h.toUpperCase() === 'MONTH') || headers[1];
    const years = [...new Set(rawData.map(row => {
        const m = (row[monthKey] || '').toString().trim();
        return parseMonthStr(m).year;
    }))].filter(y => y > 0).sort((a, b) => a - b);

    const fromSelect = document.getElementById('yearFrom');
    const toSelect = document.getElementById('yearTo');
    if (!fromSelect || !toSelect) return;

    const prevFrom = fromSelect.value;
    const prevTo = toSelect.value;

    fromSelect.innerHTML = years.map(y => `<option value="${y}">${y}</option>`).join('');
    toSelect.innerHTML = years.map(y => `<option value="${y}">${y}</option>`).join('');

    // Restore or set defaults
    if (prevFrom && years.includes(parseInt(prevFrom))) {
        fromSelect.value = prevFrom;
    } else {
        fromSelect.value = years[0];
    }
    if (prevTo && years.includes(parseInt(prevTo))) {
        toSelect.value = prevTo;
    } else {
        toSelect.value = years[years.length - 1];
    }

    yearRangeFrom = parseInt(fromSelect.value) || null;
    yearRangeTo = parseInt(toSelect.value) || null;
}

function onYearRangeChange() {
    const fromVal = document.getElementById('yearFrom')?.value;
    const toVal = document.getElementById('yearTo')?.value;
    yearRangeFrom = fromVal ? parseInt(fromVal) : null;
    yearRangeTo = toVal ? parseInt(toVal) : null;

    if (rawData.length > 0) {
        processData();
        populateDropdowns();
        renderCurrentTab();
    }
}

// ===== DATA PROCESSING =====
function processData() {
    const headers = Object.keys(rawData[0]);

    // Dynamically identify header keys
    const itemKey = headers.find(h => h.toUpperCase().includes('ITEM_CODE') || h.toUpperCase().includes('ITEM CODE')) || headers[0];
    const monthKey = headers.find(h => h.toUpperCase() === 'MONTH') || headers[1];
    const custKey = headers.find(h => h.toUpperCase().includes('CUST_NAME') || h.toUpperCase().includes('CUST NAME')) || headers[2];
    const qtyKey = headers.find(h => h.toUpperCase().includes('TOTAL QTY') || h.toUpperCase().includes('TOTAL_QTY') || h.toUpperCase().includes('QTY')) || headers[3];
    const monthYearKey = headers.find(h => h.toUpperCase().includes('MONTHYEAR') || h.toUpperCase().includes('MONTH_YEAR')) || headers[4];

    // Normalize data
    let normalized = rawData.map(row => ({
        itemCode: (row[itemKey] || '').toString().trim(),
        month: (row[monthKey] || '').toString().trim(),
        customer: (row[custKey] || '').toString().trim(),
        qty: parseFloat((row[qtyKey] || '0').toString().replace(/,/g, '')) || 0,
        monthYear: (row[monthYearKey] || '').toString().trim()
    })).filter(r => r.itemCode && r.customer && r.month);

    // Apply global year range filter
    if (yearRangeFrom) {
        normalized = normalized.filter(r => parseMonthStr(r.month).year >= yearRangeFrom);
    }
    if (yearRangeTo) {
        normalized = normalized.filter(r => parseMonthStr(r.month).year <= yearRangeTo);
    }

    // Get all months present in data (sorted)
    const allMonths = [...new Set(normalized.map(r => r.month))];
    allMonths.sort((a, b) => monthSortKey(a) - monthSortKey(b));

    // Get current month (latest month in data)
    const latestMonth = allMonths[allMonths.length - 1];
    const latestMonthSortKey = monthSortKey(latestMonth);

    // Unique sets
    const uniqueCustomers = [...new Set(normalized.map(r => r.customer))].sort();
    const uniqueItems = [...new Set(normalized.map(r => r.itemCode))].sort();

    // Build lookup: customer -> item -> {months: {MONTH: qty}, totalQty, monthCount}
    const customerItemMap = {};
    normalized.forEach(row => {
        if (!customerItemMap[row.customer]) customerItemMap[row.customer] = {};
        if (!customerItemMap[row.customer][row.itemCode]) {
            customerItemMap[row.customer][row.itemCode] = { months: {}, totalQty: 0, monthCount: 0 };
        }
        const entry = customerItemMap[row.customer][row.itemCode];
        entry.months[row.month] = (entry.months[row.month] || 0) + row.qty;
        entry.totalQty += row.qty;
    });

    // Calculate month counts
    Object.values(customerItemMap).forEach(items => {
        Object.values(items).forEach(item => {
            item.monthCount = Object.keys(item.months).length;
        });
    });

    // Build monthly totals
    const monthlyTotals = {};
    allMonths.forEach(m => { monthlyTotals[m] = 0; });
    normalized.forEach(r => {
        monthlyTotals[r.month] = (monthlyTotals[r.month] || 0) + r.qty;
    });

    // Build item totals
    const itemTotals = {};
    normalized.forEach(r => {
        itemTotals[r.itemCode] = (itemTotals[r.itemCode] || 0) + r.qty;
    });

    // Build customer totals
    const customerTotals = {};
    normalized.forEach(r => {
        customerTotals[r.customer] = (customerTotals[r.customer] || 0) + r.qty;
    });

    // Dropped Items Analysis
    const droppedItems = [];
    Object.entries(customerItemMap).forEach(([customer, items]) => {
        Object.entries(items).forEach(([itemCode, data]) => {
            const purchaseMonths = Object.keys(data.months);
            const lastPurchaseMonth = purchaseMonths.sort((a, b) =>
                monthSortKey(b) - monthSortKey(a)
            )[0];

            const monthsSinceLastPurchase = monthsDiffCalc(latestMonth, lastPurchaseMonth);

            if (monthsSinceLastPurchase > 0 && data.monthCount >= 1) {
                const avgQty = data.totalQty / data.monthCount;
                droppedItems.push({
                    customer,
                    itemCode,
                    lastMonth: lastPurchaseMonth,
                    lastQty: data.months[lastPurchaseMonth],
                    avgQtyPerMonth: avgQty,
                    monthsInactive: monthsSinceLastPurchase,
                    totalQty: data.totalQty,
                    monthCount: data.monthCount
                });
            }
        });
    });

    // Declining Analysis
    const decliningPairs = [];
    Object.entries(customerItemMap).forEach(([customer, items]) => {
        Object.entries(items).forEach(([itemCode, data]) => {
            const monthEntries = Object.entries(data.months)
                .sort((a, b) => monthSortKey(a[0]) - monthSortKey(b[0]));

            if (monthEntries.length >= 2) {
                let peakQty = 0, peakMonth = '';
                monthEntries.forEach(([m, q]) => {
                    if (q > peakQty) { peakQty = q; peakMonth = m; }
                });

                const latestEntry = monthEntries[monthEntries.length - 1];
                const latestQty = latestEntry[1];
                const latestM = latestEntry[0];

                if (monthSortKey(latestM) > monthSortKey(peakMonth) && latestQty < peakQty) {
                    const declinePct = ((peakQty - latestQty) / peakQty) * 100;
                    decliningPairs.push({
                        customer,
                        itemCode,
                        peakQty,
                        peakMonth,
                        latestQty,
                        latestMonth: latestM,
                        declinePct,
                        monthEntries
                    });
                }
            }
        });
    });

    // Active pairs (bought in latest month)
    let activePairs = 0;
    Object.entries(customerItemMap).forEach(([customer, items]) => {
        Object.entries(items).forEach(([itemCode, data]) => {
            if (data.months[latestMonth]) activePairs++;
        });
    });

    // Store processed data
    processedData = {
        normalized,
        allMonths,
        latestMonth,
        latestMonthSortKey,
        uniqueCustomers,
        uniqueItems,
        customerItemMap,
        monthlyTotals,
        itemTotals,
        customerTotals,
        droppedItems,
        decliningPairs,
        activePairs,
        totalQty: normalized.reduce((sum, r) => sum + r.qty, 0)
    };
}

// ===== DROPDOWN POPULATION =====
function populateDropdowns() {
    const { uniqueCustomers, uniqueItems } = processedData;

    // Top Items customer dropdown
    populateSelect('topItemsCustomerSelect', uniqueCustomers, 'ALL', 'All Customers');
    // Trends customer dropdown
    populateSelect('trendCustomerSelect', uniqueCustomers, 'ALL', 'All Customers (Total)');
    // Trends item dropdown
    populateSelect('trendItemSelect', uniqueItems, 'ALL', 'All Items (Total)');

    // Deep Dive datalist
    const datalist = document.getElementById('deepDiveItemList');
    if (datalist) {
        datalist.innerHTML = uniqueItems.map(i => `<option value="${escapeHtml(i)}">`).join('');
    }
}

function populateSelect(id, options, defaultValue, defaultLabel) {
    const select = document.getElementById(id);
    if (!select) return;
    select.innerHTML = `<option value="${defaultValue}">${defaultLabel}</option>`;
    options.forEach(opt => {
        select.innerHTML += `<option value="${escapeHtml(opt)}">${escapeHtml(opt)}</option>`;
    });
}

// ===== RENDER ROUTER =====
function renderCurrentTab() {
    if (rawData.length === 0 && currentTab !== 'salesintel') return;
    switch (currentTab) {
        case 'overview': renderOverview(); break;
        case 'dropped': renderDropped(); break;
        case 'heatmap': renderHeatmap(); break;
        case 'declining': renderDeclining(); break;
        case 'topitems': renderTopItems(); break;
        case 'trends': renderTrends(); break;
        case 'deepdive': /* Wait for user action */ break;
        case 'reorder': renderReorder(); break;
        case 'salesintel': renderSalesIntel(); break;
    }
}

// ===== OVERVIEW TAB =====
function renderOverview() {
    const { uniqueCustomers, uniqueItems, totalQty, allMonths, droppedItems, decliningPairs, activePairs, monthlyTotals, customerTotals, itemTotals, latestMonth } = processedData;

    // Hero stats
    animateCounter('heroTotalCustomers', uniqueCustomers.length);
    animateCounter('heroTotalItems', uniqueItems.length);
    animateCounter('heroTotalQty', totalQty);
    setText('heroMonthsCovered', allMonths.length);

    // KPIs
    animateCounter('kpiActiveCustomers', uniqueCustomers.length);
    animateCounter('kpiUniqueItems', uniqueItems.length);
    animateCounter('kpiActivePairs', activePairs);
    animateCounter('kpiDropped', droppedItems.length);
    animateCounter('kpiDeclining', decliningPairs.length);
    setText('kpiMonths', `${allMonths[0]}–${allMonths[allMonths.length - 1]}`);

    // Monthly Volume Chart
    renderChart('chartMonthlyVolume', {
        type: 'bar',
        data: {
            labels: allMonths.map(m => formatMonthLabel(m)),
            datasets: [{
                label: 'Total Qty',
                data: allMonths.map(m => monthlyTotals[m] || 0),
                backgroundColor: allMonths.map((_, i) => {
                    const colors = COLORS.palette;
                    return colors[i % colors.length] + '80';
                }),
                borderColor: allMonths.map((_, i) => COLORS.palette[i % COLORS.palette.length]),
                borderWidth: 2,
                borderRadius: 8,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(15, 23, 42, 0.9)',
                    titleColor: '#f1f5f9',
                    bodyColor: '#94a3b8',
                    borderColor: 'rgba(99, 102, 241, 0.3)',
                    borderWidth: 1,
                    padding: 12,
                    cornerRadius: 8,
                    callbacks: {
                        label: ctx => `Qty: ${formatNumber(ctx.parsed.y)}`
                    }
                }
            },
            scales: {
                x: {
                    grid: { color: 'rgba(255,255,255,0.04)' },
                    ticks: { color: getComputedStyle(document.body).getPropertyValue('--text-muted').trim() }
                },
                y: {
                    grid: { color: 'rgba(255,255,255,0.04)' },
                    ticks: {
                        color: getComputedStyle(document.body).getPropertyValue('--text-muted').trim(),
                        callback: v => formatNumber(v)
                    }
                }
            }
        }
    });

    // Top 10 Customers Doughnut
    const sortedCustomers = Object.entries(customerTotals).sort((a, b) => b[1] - a[1]).slice(0, 10);
    renderChart('chartTopCustomers', {
        type: 'doughnut',
        data: {
            labels: sortedCustomers.map(c => truncate(c[0], 20)),
            datasets: [{
                data: sortedCustomers.map(c => c[1]),
                backgroundColor: COLORS.palette.slice(0, 10).map(c => c + 'CC'),
                borderColor: COLORS.palette.slice(0, 10),
                borderWidth: 2,
                hoverOffset: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        color: getComputedStyle(document.body).getPropertyValue('--text-secondary').trim(),
                        padding: 12,
                        font: { size: 11, family: 'Outfit' },
                        usePointStyle: true,
                        pointStyleWidth: 10
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(15, 23, 42, 0.9)',
                    titleColor: '#f1f5f9',
                    bodyColor: '#94a3b8',
                    padding: 12,
                    cornerRadius: 8,
                    callbacks: {
                        label: ctx => ` ${ctx.label}: ${formatNumber(ctx.parsed)}`
                    }
                }
            },
            cutout: '65%'
        }
    });

    // Top 15 Items (Horizontal Bar)
    const sortedItems = Object.entries(itemTotals).sort((a, b) => b[1] - a[1]).slice(0, 15);
    renderChart('chartTopItems', {
        type: 'bar',
        data: {
            labels: sortedItems.map(i => truncate(i[0], 25)),
            datasets: [{
                label: 'Total Qty',
                data: sortedItems.map(i => i[1]),
                backgroundColor: COLORS.palette.map(c => c + '70'),
                borderColor: COLORS.palette,
                borderWidth: 1.5,
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(15, 23, 42, 0.9)',
                    padding: 12,
                    cornerRadius: 8,
                    callbacks: { label: ctx => `Qty: ${formatNumber(ctx.parsed.x)}` }
                }
            },
            scales: {
                x: {
                    grid: { color: 'rgba(255,255,255,0.04)' },
                    ticks: { color: getComputedStyle(document.body).getPropertyValue('--text-muted').trim(), callback: v => formatNumber(v) }
                },
                y: {
                    grid: { display: false },
                    ticks: { color: getComputedStyle(document.body).getPropertyValue('--text-secondary').trim(), font: { size: 11 } }
                }
            }
        }
    });

    // Activity Status Pie
    const activeCount = processedData.activePairs;
    const droppedCount = processedData.droppedItems.length;
    const totalPairs = Object.entries(processedData.customerItemMap).reduce((sum, [_, items]) => sum + Object.keys(items).length, 0);
    const stableCount = totalPairs - activeCount - droppedCount;

    renderChart('chartActivityStatus', {
        type: 'doughnut',
        data: {
            labels: ['Active (Current Month)', 'Dropped (Stopped)', 'Stable (Historical)'],
            datasets: [{
                data: [activeCount, droppedCount, Math.max(stableCount, 0)],
                backgroundColor: [COLORS.success + 'CC', COLORS.danger + 'CC', COLORS.info + 'CC'],
                borderColor: [COLORS.success, COLORS.danger, COLORS.info],
                borderWidth: 2,
                hoverOffset: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: getComputedStyle(document.body).getPropertyValue('--text-secondary').trim(),
                        padding: 16,
                        font: { size: 12, family: 'Outfit' },
                        usePointStyle: true
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(15, 23, 42, 0.9)',
                    padding: 12,
                    cornerRadius: 8
                }
            },
            cutout: '55%'
        }
    });
}

// ===== DROPPED ITEMS TAB =====
function renderDropped() {
    const { droppedItems, latestMonth } = processedData;
    const searchVal = (document.getElementById('droppedCustomerSearch')?.value || '').toLowerCase();
    const minMonths = parseInt(document.getElementById('droppedMonthsFilter')?.value || '2');
    const sortBy = document.getElementById('droppedSortBy')?.value || 'months_desc';

    let filtered = droppedItems.filter(d => d.monthsInactive >= minMonths);
    if (searchVal) {
        filtered = filtered.filter(d => d.customer.toLowerCase().includes(searchVal));
    }

    // Sort
    switch (sortBy) {
        case 'months_desc': filtered.sort((a, b) => b.monthsInactive - a.monthsInactive); break;
        case 'items_desc': filtered.sort((a, b) => b.avgQtyPerMonth - a.avgQtyPerMonth); break;
        case 'customer_asc': filtered.sort((a, b) => a.customer.localeCompare(b.customer)); break;
    }

    // Update badge
    setText('droppedBadge', `${filtered.length} items dropped`);

    // Render table
    const tbody = document.getElementById('droppedBody');
    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="empty-msg">No dropped items found matching your filters</td></tr>';
        return;
    }

    tbody.innerHTML = filtered.slice(0, 200).map((d, i) => {
        const statusClass = d.monthsInactive >= 6 ? 'stopped' : d.monthsInactive >= 3 ? 'declining' : 'new';
        const statusText = d.monthsInactive >= 6 ? '🔴 Lost' : d.monthsInactive >= 3 ? '🟡 At Risk' : '🟠 Recently Stopped';
        const monthBadgeClass = d.monthsInactive >= 6 ? 'critical' : d.monthsInactive >= 3 ? 'warning' : 'ok';

        return `<tr style="--row-index: ${i}">
            <td class="text-center">${i + 1}</td>
            <td class="customer-name">${escapeHtml(d.customer)}</td>
            <td class="item-name">${escapeHtml(d.itemCode)}</td>
            <td>${formatMonthLabel(d.lastMonth)}</td>
            <td class="text-right">${formatNumber(d.lastQty)}</td>
            <td class="text-right">${d.avgQtyPerMonth.toFixed(1)}</td>
            <td class="text-center"><span class="months-badge ${monthBadgeClass}">${d.monthsInactive}</span></td>
            <td><span class="status-badge ${statusClass}">${statusText}</span></td>
        </tr>`;
    }).join('');
}

// ===== HEATMAP TAB =====
function renderHeatmap() {
    const { customerItemMap, allMonths, latestMonth } = processedData;
    const custSearch = (document.getElementById('heatmapCustomerSearch')?.value || '').toLowerCase();
    const itemSearch = (document.getElementById('heatmapItemSearch')?.value || '').toLowerCase();
    const limit = parseInt(document.getElementById('heatmapLimit')?.value || '50');

    // Display months in reverse chronological order (latest first)
    const displayMonths = [...allMonths].reverse();

    // Build rows: each row is a Customer + Item pair
    let rows = [];
    Object.entries(customerItemMap).forEach(([customer, items]) => {
        if (custSearch && !customer.toLowerCase().includes(custSearch)) return;
        Object.entries(items).forEach(([itemCode, data]) => {
            if (itemSearch && !itemCode.toLowerCase().includes(itemSearch)) return;
            rows.push({ customer, itemCode, months: data.months, totalQty: data.totalQty });
        });
    });

    // Sort by total qty desc
    rows.sort((a, b) => b.totalQty - a.totalQty);
    rows = rows.slice(0, limit);

    // Find max qty for heat levels
    let maxQty = 0;
    rows.forEach(r => {
        Object.values(r.months).forEach(q => { if (q > maxQty) maxQty = q; });
    });

    // Render header (reverse chronological)
    const thead = document.getElementById('heatmapHead');
    thead.innerHTML = `<tr>
        <th>Customer / Item</th>
        ${displayMonths.map(m => `<th>${formatMonthLabel(m)}</th>`).join('')}
        <th>Total</th>
    </tr>`;

    // Render body
    const tbody = document.getElementById('heatmapBody');
    if (rows.length === 0) {
        tbody.innerHTML = `<tr><td colspan="${allMonths.length + 2}" class="empty-msg">No data matching filters</td></tr>`;
        return;
    }

    tbody.innerHTML = rows.map((row, i) => {
        const cells = displayMonths.map(m => {
            const qty = row.months[m] || 0;
            const level = qty === 0 ? 0 : Math.min(5, Math.ceil((qty / maxQty) * 5));
            const mKey = monthSortKey(m);
            const latestKey = monthSortKey(processedData.latestMonth);
            const isGap = qty === 0 && mKey <= latestKey;
            const hasEarlierPurchase = allMonths.filter(am =>
                monthSortKey(am) < mKey && row.months[am]
            ).length > 0;
            const hasLaterPurchase = allMonths.filter(am =>
                monthSortKey(am) > mKey && row.months[am]
            ).length > 0;
            const isMiddleGap = isGap && hasEarlierPurchase && hasLaterPurchase;
            const isTrailingGap = isGap && hasEarlierPurchase && !hasLaterPurchase && mKey <= latestKey;

            const cellClass = (isMiddleGap || isTrailingGap) ? 'current-gap' : `level-${level}`;
            return `<td><span class="heatmap-cell ${cellClass}" title="${escapeHtml(row.customer)} - ${escapeHtml(row.itemCode)}: ${qty > 0 ? formatNumber(qty) : 'No purchase'}">${qty > 0 ? formatNumber(qty) : (isMiddleGap || isTrailingGap ? '✗' : '—')}</span></td>`;
        }).join('');

        return `<tr style="--row-index: ${i}">
            <td title="${escapeHtml(row.customer)} | ${escapeHtml(row.itemCode)}">
                <strong>${escapeHtml(truncate(row.customer, 18))}</strong>
                <br><span style="font-size:10px;color:var(--text-muted)">${escapeHtml(truncate(row.itemCode, 25))}</span>
            </td>
            ${cells}
            <td><strong style="color:var(--accent-primary)">${formatNumber(row.totalQty)}</strong></td>
        </tr>`;
    }).join('');
}

// ===== DECLINING TAB =====
function renderDeclining() {
    const { decliningPairs } = processedData;
    const search = (document.getElementById('decliningCustomerSearch')?.value || '').toLowerCase();
    const threshold = parseInt(document.getElementById('decliningThreshold')?.value || '50');

    let filtered = decliningPairs.filter(d => d.declinePct >= threshold);
    if (search) {
        filtered = filtered.filter(d => d.customer.toLowerCase().includes(search));
    }
    filtered.sort((a, b) => b.declinePct - a.declinePct);

    setText('decliningBadge', `${filtered.length} pairs`);

    const tbody = document.getElementById('decliningBody');
    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" class="empty-msg">No declining patterns found matching your criteria</td></tr>';
        return;
    }

    tbody.innerHTML = filtered.slice(0, 200).map((d, i) => {
        const trendClass = d.declinePct >= 75 ? 'trend-down' : d.declinePct >= 50 ? 'trend-down' : 'trend-flat';
        const sparkline = d.monthEntries.map(([m, q]) => q).join(', ');

        return `<tr style="--row-index: ${i}">
            <td class="text-center">${i + 1}</td>
            <td class="customer-name">${escapeHtml(d.customer)}</td>
            <td class="item-name">${escapeHtml(d.itemCode)}</td>
            <td class="text-right" style="font-weight:600;color:var(--success)">${formatNumber(d.peakQty)}</td>
            <td>${formatMonthLabel(d.peakMonth)}</td>
            <td class="text-right" style="font-weight:600;color:var(--danger)">${formatNumber(d.latestQty)}</td>
            <td>${formatMonthLabel(d.latestMonth)}</td>
            <td class="text-right"><span class="${trendClass}" style="font-weight:700">-${d.declinePct.toFixed(1)}%</span></td>
            <td><span class="trend-arrow trend-down">▼ ${d.declinePct >= 75 ? 'Critical' : d.declinePct >= 50 ? 'Severe' : 'Moderate'}</span></td>
        </tr>`;
    }).join('');
}

// ===== TOP ITEMS TAB =====
function renderTopItems() {
    const { customerItemMap, allMonths, latestMonth } = processedData;
    const selectedCustomer = document.getElementById('topItemsCustomerSelect')?.value || 'ALL';
    const limit = parseInt(document.getElementById('topItemsLimit')?.value || '20');

    let pairs = [];
    Object.entries(customerItemMap).forEach(([customer, items]) => {
        if (selectedCustomer !== 'ALL' && customer !== selectedCustomer) return;
        Object.entries(items).forEach(([itemCode, data]) => {
            const sortedMonths = Object.keys(data.months).sort((a, b) => monthSortKey(a) - monthSortKey(b));
            const firstMonth = sortedMonths[0];
            const lastMonth = sortedMonths[sortedMonths.length - 1];
            const trend = sortedMonths.length >= 2 ?
                (data.months[lastMonth] >= data.months[firstMonth] ? 'up' : 'down') : 'flat';

            pairs.push({
                itemCode,
                customer,
                totalQty: data.totalQty,
                monthCount: data.monthCount,
                avgQtyPerMonth: data.totalQty / data.monthCount,
                trend
            });
        });
    });

    pairs.sort((a, b) => b.totalQty - a.totalQty);
    const topPairs = pairs.slice(0, limit);

    // Chart
    const chartData = topPairs.slice(0, 15);
    renderChart('chartTopItemsByCustomer', {
        type: 'bar',
        data: {
            labels: chartData.map(p => truncate(p.itemCode, 20)),
            datasets: [{
                label: 'Total Qty',
                data: chartData.map(p => p.totalQty),
                backgroundColor: chartData.map((_, i) => COLORS.palette[i % COLORS.palette.length] + '80'),
                borderColor: chartData.map((_, i) => COLORS.palette[i % COLORS.palette.length]),
                borderWidth: 2,
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8 }
            },
            scales: {
                x: {
                    grid: { color: 'rgba(255,255,255,0.04)' },
                    ticks: { color: getComputedStyle(document.body).getPropertyValue('--text-muted').trim(), maxRotation: 45, font: { size: 10 } }
                },
                y: {
                    grid: { color: 'rgba(255,255,255,0.04)' },
                    ticks: { color: getComputedStyle(document.body).getPropertyValue('--text-muted').trim(), callback: v => formatNumber(v) }
                }
            }
        }
    });

    // Table
    const tbody = document.getElementById('topItemsBody');
    if (topPairs.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="empty-msg">No data found</td></tr>';
        return;
    }

    tbody.innerHTML = topPairs.map((p, i) => {
        const trendIcon = p.trend === 'up' ? '📈' : p.trend === 'down' ? '📉' : '➡️';
        const trendClass = p.trend === 'up' ? 'trend-up' : p.trend === 'down' ? 'trend-down' : 'trend-flat';

        return `<tr style="--row-index: ${i}">
            <td class="text-center">${i + 1}</td>
            <td class="item-name">${escapeHtml(p.itemCode)}</td>
            <td class="customer-name">${escapeHtml(p.customer)}</td>
            <td class="text-right" style="font-weight:700">${formatNumber(p.totalQty)}</td>
            <td class="text-right">${p.monthCount}</td>
            <td class="text-right">${p.avgQtyPerMonth.toFixed(1)}</td>
            <td><span class="trend-arrow ${trendClass}">${trendIcon} ${p.trend === 'up' ? 'Growing' : p.trend === 'down' ? 'Declining' : 'Stable'}</span></td>
        </tr>`;
    }).join('');
}

// ===== TRENDS TAB =====
function renderTrends() {
    const { normalized, allMonths } = processedData;
    const selectedCustomer = document.getElementById('trendCustomerSelect')?.value || 'ALL';
    const selectedItem = document.getElementById('trendItemSelect')?.value || 'ALL';

    let filtered = normalized;
    if (selectedCustomer !== 'ALL') filtered = filtered.filter(r => r.customer === selectedCustomer);
    if (selectedItem !== 'ALL') filtered = filtered.filter(r => r.itemCode === selectedItem);

    // Aggregate by month
    const monthAgg = {};
    const monthItems = {};
    allMonths.forEach(m => { monthAgg[m] = 0; monthItems[m] = new Set(); });
    filtered.forEach(r => {
        monthAgg[r.month] = (monthAgg[r.month] || 0) + r.qty;
        if (!monthItems[r.month]) monthItems[r.month] = new Set();
        monthItems[r.month].add(r.itemCode);
    });

    const chartTitle = document.getElementById('trendChartTitle');
    if (chartTitle) {
        let title = 'Monthly Purchase Trend';
        if (selectedCustomer !== 'ALL') title += ` — ${truncate(selectedCustomer, 30)}`;
        if (selectedItem !== 'ALL') title += ` — ${truncate(selectedItem, 25)}`;
        chartTitle.textContent = title;
    }

    // Line Chart
    const values = allMonths.map(m => monthAgg[m] || 0);
    renderChart('chartTrendLine', {
        type: 'line',
        data: {
            labels: allMonths.map(m => formatMonthLabel(m)),
            datasets: [{
                label: 'Total Qty',
                data: values,
                borderColor: COLORS.primary,
                backgroundColor: COLORS.primary + '20',
                fill: true,
                tension: 0.4,
                pointBackgroundColor: COLORS.primary,
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
                pointRadius: 5,
                pointHoverRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(15,23,42,0.9)',
                    padding: 12,
                    cornerRadius: 8,
                    callbacks: { label: ctx => `Qty: ${formatNumber(ctx.parsed.y)}` }
                }
            },
            scales: {
                x: {
                    grid: { color: 'rgba(255,255,255,0.04)' },
                    ticks: { color: getComputedStyle(document.body).getPropertyValue('--text-muted').trim() }
                },
                y: {
                    grid: { color: 'rgba(255,255,255,0.04)' },
                    ticks: { color: getComputedStyle(document.body).getPropertyValue('--text-muted').trim(), callback: v => formatNumber(v) },
                    beginAtZero: true
                }
            },
            interaction: { intersect: false, mode: 'index' }
        }
    });

    // Trend Table
    const tbody = document.getElementById('trendBody');
    let prevQty = 0;
    tbody.innerHTML = allMonths.map((m, i) => {
        const qty = monthAgg[m] || 0;
        const items = monthItems[m] ? monthItems[m].size : 0;
        const change = i === 0 ? 0 : qty - prevQty;
        const changePct = prevQty > 0 ? ((change / prevQty) * 100).toFixed(1) : '—';
        const trendClass = change > 0 ? 'trend-up' : change < 0 ? 'trend-down' : 'trend-flat';
        const trendIcon = change > 0 ? '▲' : change < 0 ? '▼' : '—';
        prevQty = qty;

        return `<tr style="--row-index: ${i}">
            <td style="font-weight:600">${formatMonthLabel(m)}</td>
            <td class="text-right" style="font-weight:700">${formatNumber(qty)}</td>
            <td class="text-right">${items}</td>
            <td class="text-right"><span class="${trendClass}" style="font-weight:600">${change > 0 ? '+' : ''}${formatNumber(change)} ${changePct !== '—' ? `(${changePct}%)` : ''}</span></td>
            <td><span class="trend-arrow ${trendClass}">${trendIcon}</span></td>
        </tr>`;
    }).join('');
}

// ===== DEEP DIVE TAB =====
function renderDeepDive() {
    const searchVal = (document.getElementById('deepDiveItemSearch')?.value || '').trim();
    if (!searchVal) return;

    const { normalized, allMonths, customerItemMap, latestMonth } = processedData;

    // Find matching item
    const itemData = normalized.filter(r => r.itemCode.toLowerCase() === searchVal.toLowerCase());
    if (itemData.length === 0) {
        document.getElementById('deepDiveResults').style.display = 'none';
        document.getElementById('deepDiveEmpty').style.display = 'flex';
        document.getElementById('deepDiveEmpty').querySelector('.empty-state-text').textContent = `No data found for item "${searchVal}"`;
        return;
    }

    document.getElementById('deepDiveResults').style.display = 'block';
    document.getElementById('deepDiveEmpty').style.display = 'none';

    const itemCode = itemData[0].itemCode;

    // Aggregate
    const totalQty = itemData.reduce((s, r) => s + r.qty, 0);
    const customers = [...new Set(itemData.map(r => r.customer))];
    const monthsWithData = [...new Set(itemData.map(r => r.month))];

    // Monthly totals for this item
    const monthlyQty = {};
    allMonths.forEach(m => { monthlyQty[m] = 0; });
    itemData.forEach(r => { monthlyQty[r.month] += r.qty; });

    // Customer breakdown
    const customerBreakdown = {};
    itemData.forEach(r => {
        if (!customerBreakdown[r.customer]) {
            customerBreakdown[r.customer] = { totalQty: 0, months: {} };
        }
        customerBreakdown[r.customer].totalQty += r.qty;
        customerBreakdown[r.customer].months[r.month] = (customerBreakdown[r.customer].months[r.month] || 0) + r.qty;
    });

    // Summary card
    const summaryDiv = document.getElementById('deepDiveSummary');
    summaryDiv.innerHTML = `
        <div class="customer-detail-header">
            <div class="customer-detail-avatar">📦</div>
            <div>
                <h3 style="font-size:18px;font-weight:700">${escapeHtml(itemCode)}</h3>
                <p style="font-size:13px;color:var(--text-muted)">Detailed item analysis across ${customers.length} customers</p>
            </div>
        </div>
        <div class="customer-detail-meta">
            <div class="customer-meta-item">
                <span class="customer-meta-label">Total Qty (YTD)</span>
                <span class="customer-meta-value" style="color:var(--accent-primary)">${formatNumber(totalQty)}</span>
            </div>
            <div class="customer-meta-item">
                <span class="customer-meta-label">Unique Customers</span>
                <span class="customer-meta-value" style="color:var(--info)">${customers.length}</span>
            </div>
            <div class="customer-meta-item">
                <span class="customer-meta-label">Months Active</span>
                <span class="customer-meta-value" style="color:var(--success)">${monthsWithData.length} / ${allMonths.length}</span>
            </div>
            <div class="customer-meta-item">
                <span class="customer-meta-label">Avg Qty/Month</span>
                <span class="customer-meta-value" style="color:var(--cyan)">${(totalQty / Math.max(monthsWithData.length, 1)).toFixed(1)}</span>
            </div>
        </div>
    `;

    // Trend Chart
    renderChart('chartDeepDiveTrend', {
        type: 'line',
        data: {
            labels: allMonths.map(m => formatMonthLabel(m)),
            datasets: [{
                label: itemCode,
                data: allMonths.map(m => monthlyQty[m]),
                borderColor: COLORS.cyan,
                backgroundColor: COLORS.cyan + '20',
                fill: true,
                tension: 0.4,
                pointBackgroundColor: COLORS.cyan,
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
                pointRadius: 5
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false }, tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8 } },
            scales: {
                x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: getComputedStyle(document.body).getPropertyValue('--text-muted').trim() } },
                y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: getComputedStyle(document.body).getPropertyValue('--text-muted').trim() }, beginAtZero: true }
            }
        }
    });

    // Customer Pie
    const sortedCusts = Object.entries(customerBreakdown).sort((a, b) => b[1].totalQty - a[1].totalQty).slice(0, 10);
    renderChart('chartDeepDiveCustomers', {
        type: 'doughnut',
        data: {
            labels: sortedCusts.map(c => truncate(c[0], 20)),
            datasets: [{
                data: sortedCusts.map(c => c[1].totalQty),
                backgroundColor: COLORS.palette.slice(0, 10).map(c => c + 'CC'),
                borderColor: COLORS.palette.slice(0, 10),
                borderWidth: 2,
                hoverOffset: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'right', labels: { color: getComputedStyle(document.body).getPropertyValue('--text-secondary').trim(), padding: 10, font: { size: 11 }, usePointStyle: true } },
                tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8 }
            },
            cutout: '60%'
        }
    });

    // Customer Table
    const tbody = document.getElementById('deepDiveBody');
    const custEntries = Object.entries(customerBreakdown).sort((a, b) => b[1].totalQty - a[1].totalQty);

    tbody.innerHTML = custEntries.map(([cust, data], i) => {
        const custMonths = Object.keys(data.months).sort((a, b) => monthSortKey(a) - monthSortKey(b));
        const firstPurchase = custMonths[0];
        const lastPurchase = custMonths[custMonths.length - 1];
        const monthsActive = custMonths.length;
        const isActive = lastPurchase === processedData.latestMonth;
        const monthsSinceLastBuy = monthsDiffCalc(processedData.latestMonth, lastPurchase);

        const statusClass = isActive ? 'active' : monthsSinceLastBuy >= 3 ? 'stopped' : 'declining';
        const statusText = isActive ? '✅ Active' : monthsSinceLastBuy >= 3 ? `🔴 Stopped (${monthsSinceLastBuy}mo)` : `🟡 Inactive (${monthsSinceLastBuy}mo)`;

        return `<tr style="--row-index: ${i}">
            <td class="text-center">${i + 1}</td>
            <td class="customer-name">${escapeHtml(cust)}</td>
            <td class="text-right" style="font-weight:700">${formatNumber(data.totalQty)}</td>
            <td>${formatMonthLabel(firstPurchase)}</td>
            <td>${formatMonthLabel(lastPurchase)}</td>
            <td class="text-center">${monthsActive}</td>
            <td><span class="status-badge ${statusClass}">${statusText}</span></td>
        </tr>`;
    }).join('');
}

// ===== CHART UTILITIES =====
function renderChart(canvasId, config) {
    if (chartInstances[canvasId]) {
        chartInstances[canvasId].destroy();
    }
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    chartInstances[canvasId] = new Chart(ctx.getContext('2d'), config);
}

// ===== UTILITY FUNCTIONS =====
function formatNumber(n) {
    if (typeof n !== 'number' || isNaN(n)) return '0';
    if (Math.abs(n) >= 1000000) return (n / 1000000).toFixed(1) + 'M';
    if (Math.abs(n) >= 1000) return n.toLocaleString();
    return n.toString();
}

function truncate(str, maxLen) {
    if (!str) return '';
    return str.length > maxLen ? str.substring(0, maxLen) + '…' : str;
}

function escapeHtml(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
}

function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

function animateCounter(id, target) {
    const el = document.getElementById(id);
    if (!el) return;

    const duration = 1200;
    const startTime = performance.now();
    const startValue = 0;

    function update(currentTime) {
        const elapsed = currentTime - startTime;
        const progress = Math.min(elapsed / duration, 1);
        const eased = 1 - Math.pow(1 - progress, 3); // ease-out cubic
        const current = Math.round(startValue + (target - startValue) * eased);
        el.textContent = formatNumber(current);
        if (progress < 1) requestAnimationFrame(update);
    }

    requestAnimationFrame(update);
}

function debounce(fn, delay) {
    let timer;
    return function (...args) {
        clearTimeout(timer);
        timer = setTimeout(() => fn.apply(this, args), delay);
    };
}

// =========================================================
// SMART REORDER INTELLIGENCE MODULE
// =========================================================

async function fetchReorderData() {
    const reorderUrl = typeof REORDER_LEVEL_URL !== 'undefined' ? REORDER_LEVEL_URL.trim() : '';
    const data1Url = typeof DATA1_URL !== 'undefined' ? DATA1_URL.trim() : '';

    // Both URLs must be configured
    if (!reorderUrl || reorderUrl.includes('PASTE_YOUR') || !data1Url || data1Url.includes('PASTE_YOUR')) {
        console.warn('Smart Reorder: DATA1_URL or REORDER_LEVEL_URL not configured in index.html');
        return;
    }

    try {
        // Fetch both spreadsheets in parallel (each is a separate Google Spreadsheet)
        const [reorderRes, purchaseRes] = await Promise.all([
            fetch(reorderUrl),
            fetch(data1Url)
        ]);

        const reorderResult = await reorderRes.json();
        const purchaseResult = await purchaseRes.json();

        if (reorderResult.status === 'success') {
            reorderData = reorderResult.data || [];
        }
        if (purchaseResult.status === 'success') {
            purchaseData = purchaseResult.data || [];
        }

        processReorderData();

        // If currently on reorder tab, render it
        if (currentTab === 'reorder') {
            renderReorder();
        }
    } catch (err) {
        console.error('Error fetching reorder data:', err);
    }
}

function processReorderData() {
    if (reorderData.length === 0) return;

    const rHeaders = Object.keys(reorderData[0]);
    console.log('Reorder Level Headers:', rHeaders);

    // Dynamically identify Reorder_Level keys - handle both expected and actual formats
    // Actual API returns: qty, COST, ITEMNAM
    // Expected format: ITEM CODE, REORDER, Current Stock, TO ORDER, ITEM NAME, PACKING
    const rItemCodeKey = rHeaders.find(h => h.toUpperCase().replace(/[\s_]/g, '').includes('ITEMCODE'));
    const rReorderKey = rHeaders.find(h => h.toUpperCase().includes('REORDER'));
    const rStockKey = rHeaders.find(h => h.toUpperCase().includes('CURRENT') || h.toUpperCase().includes('STOCK'));
    const rToOrderKey = rHeaders.find(h => h.toUpperCase().includes('TO ORDER') || h.toUpperCase().includes('TOORDER') || h.toUpperCase().replace(/[\s_]/g, '').includes('TOORDER'));
    const rItemNameKey = rHeaders.find(h => h.toUpperCase().includes('ITEM NAME') || h.toUpperCase().includes('ITEMNAME') || h.toUpperCase().replace(/[\s_]/g, '').includes('ITEMNAME') || h.toUpperCase().includes('ITEMNAM'));
    const rPackingKey = rHeaders.find(h => h.toUpperCase().includes('PACKING') || h.toUpperCase().includes('PACK'));
    const rQtyKey = rHeaders.find(h => h.toLowerCase() === 'qty');
    const rCostKey = rHeaders.find(h => h.toUpperCase() === 'COST');

    // Detect data format: does it have the full expected columns or just qty/COST/ITEMNAM?
    const hasFullFormat = !!(rItemCodeKey && rToOrderKey);

    let reorderItems;

    if (hasFullFormat) {
        // Full format: ITEM CODE, REORDER, Current Stock, TO ORDER, ITEM NAME, PACKING
        reorderItems = reorderData.map(row => {
            const itemCode = (row[rItemCodeKey] || '').toString().trim();
            const reorderLevel = parseFloat((row[rReorderKey] || '0').toString().replace(/,/g, '')) || 0;
            const currentStock = parseFloat((row[rStockKey] || '0').toString().replace(/,/g, '')) || 0;
            const toOrder = parseFloat((row[rToOrderKey] || '0').toString().replace(/,/g, '')) || 0;
            const itemName = (row[rItemNameKey] || '').toString().trim();
            const packing = (row[rPackingKey] || '').toString().trim();
            return { itemCode, reorderLevel, currentStock, toOrder, itemName, packing };
        }).filter(r => r.itemCode && (r.currentStock < r.reorderLevel || r.toOrder < 0));
    } else {
        // Simplified format: qty, COST, ITEMNAM (actual API response)
        // Use ITEMNAM as identifier, qty as current stock
        reorderItems = reorderData.map(row => {
            const itemName = (row[rItemNameKey] || '').toString().trim();
            const currentStock = parseFloat((row[rQtyKey || 'qty'] || '0').toString().replace(/,/g, '')) || 0;
            const cost = parseFloat((row[rCostKey || 'COST'] || '0').toString().replace(/,/g, '')) || 0;
            if (!itemName) return null;
            // Use item name as the identifier (code) for matching
            return {
                itemCode: itemName,
                reorderLevel: 0,
                currentStock,
                toOrder: 0, // will be calculated later
                itemName,
                packing: '',
                cost
            };
        }).filter(r => r !== null && r.itemName);
    }

    // Build purchase data lookup: item_code -> supplier info
    // Data1 headers: item_code, ITEM NAME, Packing, Unit QTY, tran_date, Supplier Name, Price Per PC, Price Per CTN, Master.Description2
    const supplierMap = {}; // desc2 (Master.Description2) -> [{supplier, pricePC, priceCTN, itemName, packing, qty, date}]
    const allSupplierItems = {}; // supplier -> Set of all item codes they supply
    const itemNameToDesc2 = {}; // trimmed ITEM NAME -> Master.Description2 (for reverse lookup)
    const itemPurchaseQty = {}; // desc2 -> total purchase qty (for estimating reorder needs)

    if (purchaseData.length > 0) {
        const pHeaders = Object.keys(purchaseData[0]);
        const pItemCodeKey = pHeaders.find(h => h.toLowerCase().includes('item_code') || h.toLowerCase() === 'item code') || pHeaders[0];
        const pItemNameKey = pHeaders.find(h => h.toUpperCase().includes('ITEM NAME') || h.toUpperCase().replace(/[\s_]/g, '').includes('ITEMNAME')) || pHeaders[1];
        const pPackingKey = pHeaders.find(h => h.toUpperCase().includes('PACKING') || h.toUpperCase().includes('PACK')) || pHeaders[2];
        const pQtyKey = pHeaders.find(h => h.toUpperCase().includes('UNIT QTY') || h.toUpperCase().includes('UNIT_QTY') || h.toUpperCase().includes('QTY')) || pHeaders[3];
        const pDateKey = pHeaders.find(h => h.toUpperCase().includes('TRAN_DATE') || h.toUpperCase().includes('DATE')) || pHeaders[4];
        const pSupplierKey = pHeaders.find(h => h.toUpperCase().includes('SUPPLIER')) || pHeaders[5];
        const pPricePCKey = pHeaders.find(h => h.toUpperCase().includes('PRICE PER PC') || h.toUpperCase().replace(/[\s_]/g, '').includes('PRICEPERPC')) || pHeaders[6];
        const pPriceCTNKey = pHeaders.find(h => h.toUpperCase().includes('PRICE PER CTN') || h.toUpperCase().replace(/[\s_]/g, '').includes('PRICEPERCTN')) || pHeaders[7];
        const pDesc2Key = pHeaders.find(h => h.toUpperCase().includes('DESCRIPTION2') || h.toUpperCase().includes('MASTER.DESCRIPTION2')) || pHeaders[8];

        purchaseData.forEach(row => {
            const desc2 = (row[pDesc2Key] || '').toString().trim();
            const supplier = (row[pSupplierKey] || '').toString().trim();
            const pricePC = parseFloat((row[pPricePCKey] || '0').toString().replace(/,/g, '')) || 0;
            const priceCTN = parseFloat((row[pPriceCTNKey] || '0').toString().replace(/,/g, '')) || 0;
            const itemName = (row[pItemNameKey] || '').toString().trim();
            const packing = (row[pPackingKey] || '').toString().trim();
            const qty = parseFloat((row[pQtyKey] || '0').toString().replace(/,/g, '')) || 0;
            const date = (row[pDateKey] || '').toString().trim();
            const itemCode = (row[pItemCodeKey] || '').toString().trim();

            if (!desc2 || !supplier) return;

            if (!supplierMap[desc2]) supplierMap[desc2] = [];
            supplierMap[desc2].push({ supplier, pricePC, priceCTN, itemName, packing, qty, date, itemCode });

            // Build reverse lookup: full item name -> desc2
            if (itemName) {
                itemNameToDesc2[itemName] = desc2;
            }

            // Accumulate total purchase qty per desc2
            itemPurchaseQty[desc2] = (itemPurchaseQty[desc2] || 0) + qty;

            // Track all items per supplier
            if (!allSupplierItems[supplier]) allSupplierItems[supplier] = new Set();
            allSupplierItems[supplier].add(desc2);
        });
    }

    // Get monthly sales data from the item analysis data 
    // Item Analysis uses Master.ITEM_CODE which corresponds to ITEM CODE in Reorder_Level and Master.Description2 in Data1
    const monthlySales = {}; // itemCode -> { month -> totalQty }
    if (rawData.length > 0 && processedData.normalized) {
        processedData.normalized.forEach(row => {
            const code = row.itemCode;
            if (!monthlySales[code]) monthlySales[code] = {};
            monthlySales[code][row.month] = (monthlySales[code][row.month] || 0) + row.qty;
        });
    }

    // Process each reorder item
    const enrichedItems = reorderItems.map(item => {
        // Resolve the desc2 key for supplier lookup
        // In simplified format, item.itemCode = item.itemName (ITEMNAM from Reorder sheet)
        // In full format, item.itemCode = actual item code
        let lookupKey = item.itemCode;
        let matchedDesc2 = null;
        let matchedPacking = item.packing;

        if (!hasFullFormat) {
            // Try to match reorder ITEMNAM against Data1's ITEM NAME -> desc2
            // Direct match first
            matchedDesc2 = itemNameToDesc2[item.itemName];
            
            // If no direct match, try trimmed comparison
            if (!matchedDesc2) {
                const trimmedName = item.itemName.trim();
                for (const [fullName, desc2] of Object.entries(itemNameToDesc2)) {
                    if (fullName.trim() === trimmedName) {
                        matchedDesc2 = desc2;
                        break;
                    }
                }
            }

            if (matchedDesc2) {
                lookupKey = matchedDesc2;
                // Get packing from purchase data if not available
                const purchaseRows = supplierMap[matchedDesc2] || [];
                if (purchaseRows.length > 0 && !item.packing) {
                    matchedPacking = purchaseRows[0].packing;
                }
            }
        }

        // Find supplier info using the resolved key
        const suppliers = supplierMap[lookupKey] || [];

        // Find best supplier (lowest price per PC, excluding zero)
        let bestSupplier = null;
        let bestPricePC = Infinity;
        let bestPriceCTN = 0;
        let bestUnitQty = 1; // Default multiplier
        let bestPurchaseDate = '';
        const supplierPrices = {}; // supplier -> { prices, priceCTNs, dates, unitQtys }

        suppliers.forEach(s => {
            if (!supplierPrices[s.supplier]) {
                supplierPrices[s.supplier] = { prices: [], priceCTNs: [], dates: [], unitQtys: [] };
            }
            if (s.pricePC > 0) supplierPrices[s.supplier].prices.push(s.pricePC);
            if (s.priceCTN > 0) supplierPrices[s.supplier].priceCTNs.push(s.priceCTN);
            if (s.qty > 0) supplierPrices[s.supplier].unitQtys.push(s.qty);
            supplierPrices[s.supplier].dates.push(s.date);
        });

        Object.entries(supplierPrices).forEach(([supplier, data]) => {
            if (data.prices.length > 0) {
                // Use the most recent price (last entry)
                const latestPrice = data.prices[data.prices.length - 1];
                const latestDate = data.dates[data.dates.length - 1];
                if (latestPrice < bestPricePC) {
                    bestPricePC = latestPrice;
                    bestSupplier = supplier;
                    bestPriceCTN = data.priceCTNs.length > 0 ? data.priceCTNs[data.priceCTNs.length - 1] : 0;
                    bestUnitQty = data.unitQtys.length > 0 ? data.unitQtys[data.unitQtys.length - 1] : 1;
                    bestPurchaseDate = latestDate;
                }
            }
        });

        // Calculate average monthly purchase qty from purchase data
        const totalPurchased = itemPurchaseQty[lookupKey] || 0;
        // Rough estimate: assume purchases span ~8 months
        const estMonthlyUsage = totalPurchased > 0 ? totalPurchased / 8 : 0;

        // Calculate average monthly sales from item analysis data
        const salesData = monthlySales[item.itemCode] || (matchedDesc2 ? monthlySales[matchedDesc2] || {} : {});
        const salesMonths = Object.keys(salesData);
        let lastMonthSales = 0; // Keeping for object backward compatibility

        // Find last year's same month sales
        let lastYearSameMonthSales = 0;
        if (processedData.latestMonth) {
            const latestRef = parseMonthStr(processedData.latestMonth);
            const lyYear = latestRef.year - 1;
            const lyMonth = latestRef.month;
            
            const lyMonthKey = salesMonths.find(m => {
                const parsed = parseMonthStr(m);
                return parsed.year === lyYear && parsed.month === lyMonth;
            });

            if (lyMonthKey && salesData[lyMonthKey]) {
                lastYearSameMonthSales = salesData[lyMonthKey];
            }
        }

        // We use last year's same month as the main reference per user instruction
        let avgMonthlySales = lastYearSameMonthSales;
        let recommendedQty = lastYearSameMonthSales;
        
        let deficitRaw = hasFullFormat ? (item.currentStock < item.reorderLevel ? item.reorderLevel - item.currentStock : Math.abs(item.toOrder)) : Math.max(0, Math.ceil(lastYearSameMonthSales - item.currentStock));
        let deficit = deficitRaw;
        
        let currentStock = item.currentStock;
        let reorderLevel = item.reorderLevel;

        // Determine urgency using PCS ratios since it is identical
        let urgency = 'medium';
        if (hasFullFormat) {
            const stockRatio = item.reorderLevel > 0 ? item.currentStock / item.reorderLevel : 1;
            if (stockRatio <= 0.2 || item.currentStock === 0) urgency = 'critical';
            else if (stockRatio <= 0.5) urgency = 'high';
        } else {
            // For simplified format, use stock vs usage
            if (item.currentStock === 0) urgency = 'critical';
            else if (lastYearSameMonthSales > 0 && item.currentStock < lastYearSameMonthSales * 0.5) urgency = 'critical';
            else if (lastYearSameMonthSales > 0 && item.currentStock < lastYearSameMonthSales) urgency = 'high';
        }

        return {
            ...item,
            itemCode: matchedDesc2 || item.itemCode,
            packing: matchedPacking || item.packing,
            bestSupplier: bestSupplier || 'Unknown',
            bestPricePC: bestPricePC === Infinity ? 0 : bestPricePC,
            bestPriceCTN,
            lastPurchaseDate: bestPurchaseDate,
            allSuppliers: Object.keys(supplierPrices),
            avgMonthlySales,
            lastMonthSales,
            recommendedQty: Math.max(0, recommendedQty),
            currentStock: currentStock,
            reorderLevel: reorderLevel,
            deficit: deficit,
            unitQty: bestUnitQty,
            urgency
        };
    }).filter(item => item.bestSupplier !== 'Unknown' || item.currentStock > 0);

    // Sort by urgency then deficit
    const urgencyOrder = { critical: 0, high: 1, medium: 2 };
    enrichedItems.sort((a, b) => {
        if (urgencyOrder[a.urgency] !== urgencyOrder[b.urgency]) {
            return urgencyOrder[a.urgency] - urgencyOrder[b.urgency];
        }
        return b.deficit - a.deficit;
    });

    // Build supplier bundles
    // For each supplier that has items in the reorder list, find other items they also supply
    const supplierBundles = {};
    const reorderItemCodes = new Set(enrichedItems.map(i => i.itemCode));

    enrichedItems.forEach(item => {
        if (item.bestSupplier === 'Unknown') return;
        if (!supplierBundles[item.bestSupplier]) {
            supplierBundles[item.bestSupplier] = {
                neededItems: [],
                extraItems: []
            };
        }
        supplierBundles[item.bestSupplier].neededItems.push(item);
    });

    // For each supplier, find extra items they supply that are NOT in reorder list
    Object.entries(supplierBundles).forEach(([supplier, bundle]) => {
        const allItems = allSupplierItems[supplier] || new Set();
        allItems.forEach(itemCode => {
            if (!reorderItemCodes.has(itemCode)) {
                // Get item details from purchase data
                const purchases = supplierMap[itemCode] || [];
                const fromThisSupplier = purchases.filter(p => p.supplier === supplier);
                if (fromThisSupplier.length > 0) {
                    const latest = fromThisSupplier[fromThisSupplier.length - 1];
                    bundle.extraItems.push({
                        itemCode,
                        itemName: latest.itemName,
                        packing: latest.packing,
                        pricePC: latest.pricePC,
                        priceCTN: latest.priceCTN,
                        date: latest.date || ''
                    });
                }
            }
        });
    });

    // Get unique suppliers
    const uniqueSuppliers = [...new Set(enrichedItems.map(i => i.bestSupplier).filter(s => s !== 'Unknown'))].sort();

    reorderProcessed = {
        enrichedItems,
        supplierBundles,
        uniqueSuppliers,
        totalItems: enrichedItems.length,
        urgentCount: enrichedItems.filter(i => i.urgency === 'critical').length,
        bundleOps: Object.values(supplierBundles).filter(b => b.extraItems.length > 0).length
    };
}

let currentReorderView = 'cards';
function setReorderView(mode) {
    currentReorderView = mode;
    if (mode === 'cards') {
        document.getElementById('reorderCardsView').style.display = 'block';
        document.getElementById('reorderListView').style.display = 'none';
        document.getElementById('btnViewCards').style.background = 'var(--accent-glow)';
        document.getElementById('btnViewCards').style.color = 'var(--accent-primary)';
        document.getElementById('btnViewList').style.background = 'transparent';
        document.getElementById('btnViewList').style.color = 'var(--text-secondary)';
    } else {
        document.getElementById('reorderCardsView').style.display = 'none';
        document.getElementById('reorderListView').style.display = 'block';
        document.getElementById('btnViewCards').style.background = 'transparent';
        document.getElementById('btnViewCards').style.color = 'var(--text-secondary)';
        document.getElementById('btnViewList').style.background = 'var(--accent-glow)';
        document.getElementById('btnViewList').style.color = 'var(--accent-primary)';
    }
}

function toggleSupplierCard(id) {
    const body = document.getElementById('body-' + id);
    const icon = document.getElementById('icon-' + id);
    if (body.style.display === 'none') {
        body.style.display = 'block';
        icon.style.transform = 'rotate(-180deg)';
    } else {
        body.style.display = 'none';
        icon.style.transform = 'rotate(0deg)';
    }
}

function renderReorder() {
    if (!reorderProcessed || reorderProcessed.enrichedItems.length === 0) {
        if (reorderData.length === 0) {
            document.getElementById('supplierActionPlans').innerHTML = '<div class="empty-state"><p class="empty-state-text">Loading reorder data... Please wait.</p></div>';
            document.getElementById('simpleOrderBody').innerHTML = '<tr><td colspan="8" class="empty-msg">Loading reorder data... Please wait.</td></tr>';
        } else {
            document.getElementById('supplierActionPlans').innerHTML = '<div class="empty-state"><p class="empty-state-text">Everything is fully stocked! No items need to be ordered.</p></div>';
            document.getElementById('simpleOrderBody').innerHTML = '<tr><td colspan="8" class="empty-msg">No items need to be ordered.</td></tr>';
        }
        return;
    }

    const { enrichedItems } = reorderProcessed;
    const searchVal = (document.getElementById('planSearch')?.value || '').toLowerCase();

    // Calculate plan
    let totalItemsNeeded = 0;
    let totalEstCost = 0;
    
    // We only care about needed items for the action plan
    const actionPlan = {};

    enrichedItems.forEach(item => {
        if (item.recommendedQty <= 0) return; // Only process items that actually need ordering
        
        if (searchVal && !item.itemCode.toLowerCase().includes(searchVal) && 
            !item.itemName.toLowerCase().includes(searchVal) && 
            !item.bestSupplier.toLowerCase().includes(searchVal)) {
            return;
        }

        if (!actionPlan[item.bestSupplier]) {
            actionPlan[item.bestSupplier] = { items: [], totalCost: 0, totalCartons: 0 };
        }
        
        // Calculate item cost
        const estCost = item.recommendedQty * item.bestPricePC; // bestPricePC here is actually bestPriceCTN
        
        actionPlan[item.bestSupplier].items.push({
            ...item,
            estCost
        });
        
        actionPlan[item.bestSupplier].totalCost += estCost;
        actionPlan[item.bestSupplier].totalCartons += item.recommendedQty;
        
        totalItemsNeeded++;
        totalEstCost += estCost;
    });

    // Update KPIs
    const activeSuppliers = Object.keys(actionPlan).length;
    animateCounter('planSuppliers', activeSuppliers);
    animateCounter('planItems', totalItemsNeeded);
    
    setText('planEstCost', formatMoney(totalEstCost));
    setText('reorderTotalCostBadge', 'Total: ' + formatMoney(totalEstCost));

    renderSupplierActionCards(actionPlan);
    renderSimpleTable(actionPlan);
}

function formatMoney(val) {
    return typeof val === 'number' ? val.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '0.00';
}

function renderSupplierActionCards(actionPlan) {
    const container = document.getElementById('supplierActionPlans');
    if (!container) return;
    
    const suppliers = Object.entries(actionPlan)
        .sort((a, b) => b[1].totalCost - a[1].totalCost); // Sort by highest cost supplier first
        
    if (suppliers.length === 0) {
        container.innerHTML = '<div class="empty-state"><p class="empty-state-text">No matching items found for your search.</p></div>';
        return;
    }

    container.innerHTML = suppliers.map(([supplier, data]) => {
        const itemRows = data.items.sort((a, b) => b.recommendedQty - a.recommendedQty).map(item => `
            <tr style="border-bottom: 1px solid var(--border);">
                <td style="padding: 16px 0;">
                    <div style="font-weight: 600; color: var(--text-primary); font-size: 15px; margin-bottom: 4px;">${escapeHtml(item.itemName || item.itemCode)}</div>
                    <div style="font-size: 13px; color: var(--text-muted); margin-bottom: 4px;">
                        <span style="display:inline-block; margin-right: 12px;"><strong>Code:</strong> ${escapeHtml(item.itemCode)}</span>
                        <span style="display:inline-block; margin-right: 12px; color: ${item.currentStock === 0 ? 'var(--danger)' : 'var(--text-secondary)'}"><strong>Current Stock:</strong> ${formatNumber(item.currentStock)} PCS</span>
                        <span style="display:inline-block;"><strong>Usually Sell:</strong> ${item.avgMonthlySales} PCS / mo</span>
                    </div>
                    <div style="font-size: 12px; color: var(--text-secondary);">
                        <span style="display:inline-block;">🗓️ <strong>Last Bought:</strong> ${item.lastPurchaseDate ? escapeHtml(item.lastPurchaseDate) : 'Unknown'}</span>
                    </div>
                </td>
                <td style="padding: 16px 0; text-align: right; vertical-align: middle; width: 140px;">
                    <span style="background: rgba(59,130,246,0.1); color: var(--primary); padding: 6px 16px; border-radius: 20px; font-weight: 700; font-size: 14px; display: inline-block;">
                        Order ${formatNumber(item.recommendedQty)} PCS
                    </span>
                    <div style="font-size: 13px; color: var(--text-muted); margin-top: 6px;">
                        @ ${formatMoney(item.bestPricePC)} / PC
                    </div>
                    <div style="font-size: 14px; font-weight: 600; color: var(--text-primary); margin-top: 4px;">
                        ${formatMoney(item.estCost)}
                    </div>
                </td>
            </tr>
        `).join('');

        const safeId = 'supplier-' + supplier.replace(/[^a-zA-Z0-9]/g, '-');
        return `
        <div class="supplier-action-card" style="background: var(--bg-card); border-radius: var(--radius-lg); margin-bottom: 20px; box-shadow: var(--shadow-sm); overflow: hidden; border: 1px solid var(--border-color); transition: var(--transition);">
            <div class="supplier-card-header" onclick="toggleSupplierCard('${safeId}')" style="background: linear-gradient(135deg, var(--bg-card), var(--bg-card-hover)); padding: 20px 24px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border-color); cursor: pointer; transition: var(--transition);">
               <div style="display: flex; align-items: center; gap: 16px;">
                   <div style="font-size: 24px; background: var(--info-bg); width: 48px; height: 48px; display: flex; align-items: center; justify-content: center; border-radius: var(--radius-md); box-shadow: inset 0 2px 4px rgba(0,0,0,0.1);">🏭</div>
                   <div>
                       <h3 style="margin: 0; font-size: 18px; font-weight: 700; color: var(--text-primary); display: flex; align-items: center; gap: 8px;">
                         ${escapeHtml(supplier)}
                       </h3>
                       <span style="font-size: 13px; color: var(--text-muted); margin-top: 6px; display: flex; align-items: center; gap: 8px;">
                         <span style="background: var(--accent-glow); color: var(--accent-primary); padding: 4px 10px; border-radius: 20px; font-weight: 700;">${data.items.length} items</span>
                         <span>(${formatNumber(data.totalCartons)} total pcs)</span>
                       </span>
                   </div>
               </div>
               <div style="text-align: right; display: flex; align-items: center; gap: 24px;">
                   <div>
                       <div style="font-size: 12px; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px;">Estimated Subtotal</div>
                       <div style="font-size: 20px; font-weight: 700; color: var(--success); margin-top: 2px;">${formatMoney(data.totalCost)}</div>
                   </div>
                   <div id="icon-${safeId}" style="transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1); font-size: 20px; color: var(--text-muted); width: 32px; height: 32px; display: flex; align-items: center; justify-content: center; background: var(--bg-secondary); border-radius: 50%;">▼</div>
               </div>
            </div>
            <div id="body-${safeId}" style="padding: 0 24px; display: none; background: var(--bg-secondary);">
                <table style="width: 100%; border-collapse: collapse;">
                    <tbody>
                        ${itemRows}
                    </tbody>
                </table>
            </div>
        </div>`;
    }).join('');
}

function renderSimpleTable(actionPlan) {
    const tbody = document.getElementById('simpleOrderBody');
    if (!tbody) return;
    
    const allItems = [];
    Object.entries(actionPlan).forEach(([supplier, data]) => {
        data.items.forEach(item => {
            allItems.push({ supplier, ...item });
        });
    });
    
    if (allItems.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="empty-msg">No items to order.</td></tr>';
        return;
    }
    
    tbody.innerHTML = allItems.map(item => `
        <tr>
            <td style="font-weight:500;">${escapeHtml(item.supplier)}</td>
            <td style="color:var(--text-secondary);">${escapeHtml(item.itemCode)}</td>
            <td>${escapeHtml(item.itemName || item.itemCode)}</td>
            <td style="font-size: 12px; color: var(--text-muted);">${item.lastPurchaseDate ? escapeHtml(item.lastPurchaseDate) : 'Unknown'}</td>
            <td class="text-center" style="color:${item.currentStock === 0 ? 'var(--danger)' : 'var(--text-primary)'}; font-weight: 600;">${formatNumber(item.currentStock)}</td>
            <td class="text-right" style="color:var(--primary); font-weight: 700;">${formatNumber(item.recommendedQty)}</td>
            <td class="text-right" style="color:var(--text-muted);">${formatMoney(item.bestPricePC)}</td>
            <td class="text-right" style="font-weight:600;">${formatMoney(item.estCost)}</td>
        </tr>
    `).join('');
}

// =========================================================
// SALES INTELLIGENCE MODULE (Data2)
// =========================================================

let salesRawData = [];
let salesProcessed = null;
let currentSalesSubTab = 'revenue';
let salesDataFetched = false;

// ===== SALES SUB-TAB NAVIGATION =====
function initSalesSubNav() {
    document.querySelectorAll('.sales-sub-tab').forEach(tab => {
        tab.addEventListener('click', (e) => {
            const subtab = e.currentTarget.getAttribute('data-subtab');
            switchSalesSubTab(subtab);
        });
    });
}

function switchSalesSubTab(subtab) {
    if (subtab === currentSalesSubTab) return;
    currentSalesSubTab = subtab;

    document.querySelectorAll('.sales-sub-tab').forEach(t => t.classList.remove('active'));
    document.querySelector(`.sales-sub-tab[data-subtab="${subtab}"]`)?.classList.add('active');

    document.querySelectorAll('.sales-sub-content').forEach(c => c.classList.remove('active-sub'));
    const target = document.getElementById(`sales-sub-${subtab}`);
    if (target) {
        target.classList.add('active-sub');
        target.style.animation = 'none';
        target.offsetHeight;
        target.style.animation = '';
    }

    renderSalesCurrentSub();
}

// ===== FETCH SALES DATA =====
async function fetchSalesData() {
    const url = typeof DATA2_URL !== 'undefined' ? DATA2_URL.trim() : '';
    if (!url || url.includes('PASTE_YOUR')) {
        console.warn('Sales Intel: DATA2_URL not configured in index.html');
        return;
    }

    try {
        const response = await fetch(url);
        if (!response.ok) throw new Error('Network error: ' + response.status);
        const result = await response.json();
        if (result.status === 'error') throw new Error(result.message);

        salesRawData = result.data || [];
        if (salesRawData.length === 0) return;

        salesDataFetched = true;
        processSalesData();

        // Populate customer dropdown for Customer Analysis
        if (salesProcessed) {
            populateSelect('salesCustSelect', salesProcessed.uniqueCustomers, 'ALL', 'All Customers');
        }

        if (currentTab === 'salesintel') {
            renderSalesIntel();
        }
    } catch (err) {
        console.error('Error fetching sales data:', err);
    }
}

// ===== PROCESS SALES DATA =====
function processSalesData() {
    if (salesRawData.length === 0) return;

    const headers = Object.keys(salesRawData[0]);

    // Dynamically identify keys based on headers
    const itemCodeKey = headers.find(h => h.toLowerCase().includes('item_code') && !h.toLowerCase().includes('master')) || headers[1];
    const unitPriceKey = headers.find(h => h.toLowerCase().includes('unit_price')) || headers[2];
    const itemDesKey = headers.find(h => h.toLowerCase().includes('item_des')) || headers[3];
    const packingKey = headers.find(h => h.toLowerCase().includes('packing')) || headers[4];
    const totalUnitsKey = headers.find(h => h.toLowerCase().includes('total_units')) || headers[5];
    const netAmtKey = headers.find(h => h.toLowerCase().includes('net_amt')) || headers[6];
    const custNameKey = headers.find(h => h.toLowerCase().includes('cust_name')) || headers[7];
    const yearKey = headers.find(h => h.toLowerCase().includes('year') && h.toLowerCase().includes('copy')) || headers.find(h => h.toLowerCase().includes('year') && !h.toLowerCase().includes('year.1') && !h.toUpperCase().startsWith('YEAR.')) || headers[8];
    const monthKey = headers.find(h => h.toUpperCase() === 'MONTH') || headers[11];
    const entryNoKey = headers.find(h => h.toLowerCase().includes('entry_no')) || headers[0];

    // Normalize
    let normalized = salesRawData.map(row => {
        const unitPrice = parseFloat((row[unitPriceKey] || '0').toString().replace(/,/g, '')) || 0;
        const totalUnits = parseFloat((row[totalUnitsKey] || '0').toString().replace(/,/g, '')) || 0;
        // Calculate revenue from unit_price * total_units (ignore net_amt column)
        const netAmt = unitPrice * totalUnits;
        const year = parseInt((row[yearKey] || '0').toString()) || 0;
        const month = parseInt((row[monthKey] || '0').toString()) || 0;

        return {
            entryNo: (row[entryNoKey] || '').toString().trim(),
            itemCode: (row[itemCodeKey] || '').toString().trim(),
            unitPrice,
            itemDes: (row[itemDesKey] || '').toString().trim(),
            packing: (row[packingKey] || '').toString().trim(),
            totalUnits,
            netAmt,
            customer: (row[custNameKey] || '').toString().trim(),
            year,
            month,
            yearMonth: year && month ? `${year}-${month}` : ''
        };
    }).filter(r => r.itemCode && r.customer && r.year > 0);

    // Aggregate totals
    const totalRevenue = normalized.reduce((s, r) => s + r.netAmt, 0);
    const totalUnits = normalized.reduce((s, r) => s + r.totalUnits, 0);
    const uniqueItems = [...new Set(normalized.map(r => r.itemCode))].sort();
    const uniqueCustomers = [...new Set(normalized.map(r => r.customer))].sort();
    const totalTransactions = new Set(normalized.map(r => r.entryNo)).size;
    const avgTxnValue = totalTransactions > 0 ? totalRevenue / totalTransactions : 0;
    const weightedAvgPrice = totalUnits > 0 ? totalRevenue / totalUnits : 0;

    // All year-months sorted
    const allYearMonths = [...new Set(normalized.map(r => r.yearMonth))].filter(Boolean);
    allYearMonths.sort((a, b) => monthSortKey(a) - monthSortKey(b));

    // Monthly aggregates
    const monthlyRevenue = {};
    const monthlyUnits = {};
    const monthlyCount = {};
    allYearMonths.forEach(ym => { monthlyRevenue[ym] = 0; monthlyUnits[ym] = 0; monthlyCount[ym] = 0; });
    normalized.forEach(r => {
        if (!r.yearMonth) return;
        monthlyRevenue[r.yearMonth] = (monthlyRevenue[r.yearMonth] || 0) + r.netAmt;
        monthlyUnits[r.yearMonth] = (monthlyUnits[r.yearMonth] || 0) + r.totalUnits;
        monthlyCount[r.yearMonth] = (monthlyCount[r.yearMonth] || 0) + 1;
    });

    // Revenue by customer
    const customerRevenue = {};
    normalized.forEach(r => {
        if (!customerRevenue[r.customer]) customerRevenue[r.customer] = { revenue: 0, units: 0, entries: new Set(), months: {} };
        customerRevenue[r.customer].revenue += r.netAmt;
        customerRevenue[r.customer].units += r.totalUnits;
        customerRevenue[r.customer].entries.add(r.entryNo);
        if (!customerRevenue[r.customer].months[r.yearMonth]) customerRevenue[r.customer].months[r.yearMonth] = { revenue: 0, units: 0 };
        customerRevenue[r.customer].months[r.yearMonth].revenue += r.netAmt;
        customerRevenue[r.customer].months[r.yearMonth].units += r.totalUnits;
    });

    // Revenue by item
    const itemRevenue = {};
    normalized.forEach(r => {
        if (!itemRevenue[r.itemCode]) itemRevenue[r.itemCode] = { revenue: 0, units: 0, description: r.itemDes, prices: [], months: {} };
        itemRevenue[r.itemCode].revenue += r.netAmt;
        itemRevenue[r.itemCode].units += r.totalUnits;
        if (r.unitPrice > 0) itemRevenue[r.itemCode].prices.push(r.unitPrice);
        if (!itemRevenue[r.itemCode].months[r.yearMonth]) itemRevenue[r.itemCode].months[r.yearMonth] = { revenue: 0, units: 0, avgPrice: 0, prices: [] };
        itemRevenue[r.itemCode].months[r.yearMonth].revenue += r.netAmt;
        itemRevenue[r.itemCode].months[r.yearMonth].units += r.totalUnits;
        if (r.unitPrice > 0) itemRevenue[r.itemCode].months[r.yearMonth].prices.push(r.unitPrice);
    });

    // All years
    const allYears = [...new Set(normalized.map(r => r.year))].filter(y => y > 0).sort();

    // Yearly revenue
    const yearlyRevenue = {};
    normalized.forEach(r => {
        yearlyRevenue[r.year] = (yearlyRevenue[r.year] || 0) + r.netAmt;
    });

    salesProcessed = {
        normalized,
        totalRevenue,
        totalUnits,
        uniqueItems,
        uniqueCustomers,
        totalTransactions,
        avgTxnValue,
        weightedAvgPrice,
        allYearMonths,
        monthlyRevenue,
        monthlyUnits,
        monthlyCount,
        customerRevenue,
        itemRevenue,
        allYears,
        yearlyRevenue
    };
}

// ===== RENDER ROUTER =====
function renderSalesIntel() {
    if (!salesProcessed) {
        if (!salesDataFetched) {
            fetchSalesData();
        }
        return;
    }

    // Update hero stats
    setText('salesHeroRevenue', formatSalesMoney(salesProcessed.totalRevenue));
    setText('salesHeroUnits', formatNumber(salesProcessed.totalUnits));
    setText('salesHeroItems', formatNumber(salesProcessed.uniqueItems.length));
    setText('salesHeroCustomers', formatNumber(salesProcessed.uniqueCustomers.length));

    renderSalesCurrentSub();
}

function renderSalesCurrentSub() {
    if (!salesProcessed) return;
    switch (currentSalesSubTab) {
        case 'revenue': renderSalesRevenueOverview(); break;
        case 'topgen': renderSalesTopGenerators(); break;
        case 'priceanalytics': renderSalesPriceAnalytics(); break;
        case 'custrevenue': renderSalesCustomerRevenue(); break;
        case 'periodcomp': renderSalesPeriodComparison(); break;
        case 'stockbuilder': renderStockBuilder(); break;
    }
}

function formatSalesMoney(val) {
    if (typeof val !== 'number' || isNaN(val)) return '0.00';
    if (Math.abs(val) >= 1000000) return (val / 1000000).toFixed(2) + 'M';
    if (Math.abs(val) >= 1000) return val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return val.toFixed(2);
}

// ===== 1. REVENUE OVERVIEW =====
function renderSalesRevenueOverview() {
    const { totalRevenue, totalUnits, weightedAvgPrice, totalTransactions, avgTxnValue, allYearMonths, monthlyRevenue, monthlyUnits, customerRevenue } = salesProcessed;

    // KPIs
    setText('kpiTotalRevenue', formatSalesMoney(totalRevenue));
    setText('kpiAvgPrice', formatSalesMoney(weightedAvgPrice));
    setText('kpiTransactions', formatNumber(totalTransactions));
    setText('kpiAvgTxn', formatSalesMoney(avgTxnValue));

    const mutedColor = getComputedStyle(document.body).getPropertyValue('--text-muted').trim();

    // Monthly Revenue Trend
    renderChart('chartSalesRevenueTrend', {
        type: 'line',
        data: {
            labels: allYearMonths.map(m => formatMonthLabel(m)),
            datasets: [{
                label: 'Revenue',
                data: allYearMonths.map(m => monthlyRevenue[m] || 0),
                borderColor: '#f59e0b',
                backgroundColor: 'rgba(245, 158, 11, 0.1)',
                fill: true,
                tension: 0.4,
                pointBackgroundColor: '#f59e0b',
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
                pointRadius: 4,
                pointHoverRadius: 7
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8, callbacks: { label: ctx => `Revenue: ${formatSalesMoney(ctx.parsed.y)}` } }
            },
            scales: {
                x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor } },
                y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatSalesMoney(v) }, beginAtZero: true }
            },
            interaction: { intersect: false, mode: 'index' }
        }
    });

    // Revenue by Top 10 Customers
    const sortedCusts = Object.entries(customerRevenue).sort((a, b) => b[1].revenue - a[1].revenue).slice(0, 10);
    renderChart('chartSalesRevenueByCustomer', {
        type: 'doughnut',
        data: {
            labels: sortedCusts.map(c => truncate(c[0], 20)),
            datasets: [{
                data: sortedCusts.map(c => c[1].revenue),
                backgroundColor: COLORS.palette.slice(0, 10).map(c => c + 'CC'),
                borderColor: COLORS.palette.slice(0, 10),
                borderWidth: 2,
                hoverOffset: 8
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { position: 'right', labels: { color: mutedColor, padding: 10, font: { size: 11, family: 'Outfit' }, usePointStyle: true, pointStyleWidth: 10 } },
                tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8, callbacks: { label: ctx => ` ${ctx.label}: ${formatSalesMoney(ctx.parsed)}` } }
            },
            cutout: '65%'
        }
    });

    // Monthly Units Sold
    renderChart('chartSalesUnitsTrend', {
        type: 'bar',
        data: {
            labels: allYearMonths.map(m => formatMonthLabel(m)),
            datasets: [{
                label: 'Units Sold',
                data: allYearMonths.map(m => monthlyUnits[m] || 0),
                backgroundColor: allYearMonths.map((_, i) => COLORS.palette[i % COLORS.palette.length] + '70'),
                borderColor: allYearMonths.map((_, i) => COLORS.palette[i % COLORS.palette.length]),
                borderWidth: 1.5,
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false }, tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8, callbacks: { label: ctx => `Units: ${formatNumber(ctx.parsed.y)}` } } },
            scales: {
                x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor } },
                y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatNumber(v) } }
            }
        }
    });

    // Avg Price Trend
    const avgPriceByMonth = allYearMonths.map(m => {
        const rev = monthlyRevenue[m] || 0;
        const units = monthlyUnits[m] || 0;
        return units > 0 ? rev / units : 0;
    });
    renderChart('chartSalesAvgPriceTrend', {
        type: 'line',
        data: {
            labels: allYearMonths.map(m => formatMonthLabel(m)),
            datasets: [{
                label: 'Avg Price',
                data: avgPriceByMonth,
                borderColor: '#10b981',
                backgroundColor: 'rgba(16, 185, 129, 0.1)',
                fill: true,
                tension: 0.4,
                pointBackgroundColor: '#10b981',
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
                pointRadius: 4,
                pointHoverRadius: 7
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false }, tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8, callbacks: { label: ctx => `Avg Price: ${formatSalesMoney(ctx.parsed.y)}` } } },
            scales: {
                x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor } },
                y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatSalesMoney(v) } }
            },
            interaction: { intersect: false, mode: 'index' }
        }
    });
}

// ===== 2. TOP GENERATORS =====
function renderSalesTopGenerators() {
    const { itemRevenue, totalRevenue, customerRevenue } = salesProcessed;
    const mutedColor = getComputedStyle(document.body).getPropertyValue('--text-muted').trim();

    // Sort items by revenue
    const sortedItems = Object.entries(itemRevenue).sort((a, b) => b[1].revenue - a[1].revenue);

    // Pareto analysis
    const totalItemCount = sortedItems.length;
    const top20Count = Math.ceil(totalItemCount * 0.2);
    const top20Revenue = sortedItems.slice(0, top20Count).reduce((s, [, d]) => s + d.revenue, 0);
    const top20Pct = totalRevenue > 0 ? ((top20Revenue / totalRevenue) * 100).toFixed(1) : 0;
    const top5Revenue = sortedItems.slice(0, 5).reduce((s, [, d]) => s + d.revenue, 0);
    const bottom50Count = Math.ceil(totalItemCount * 0.5);
    const bottom50Revenue = sortedItems.slice(-bottom50Count).reduce((s, [, d]) => s + d.revenue, 0);
    const bottom50Pct = totalRevenue > 0 ? ((bottom50Revenue / totalRevenue) * 100).toFixed(1) : 0;

    setText('paretoInsightText', `Top ${top20Count} items (20%) generate ${top20Pct}% of total revenue — classic Pareto distribution`);
    setText('paretoTop20Pct', `${top20Pct}%`);
    setText('paretoTop5Revenue', formatSalesMoney(top5Revenue));
    setText('paretoBottom50', `${bottom50Pct}%`);

    // Top 15 Items by Revenue (Horizontal Bar)
    const top15Items = sortedItems.slice(0, 15);
    renderChart('chartTopItemsRevenue', {
        type: 'bar',
        data: {
            labels: top15Items.map(([code, d]) => truncate(d.description || code, 25)),
            datasets: [{
                label: 'Revenue',
                data: top15Items.map(([, d]) => d.revenue),
                backgroundColor: COLORS.palette.map(c => c + '70'),
                borderColor: COLORS.palette,
                borderWidth: 1.5,
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: { legend: { display: false }, tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8, callbacks: { label: ctx => `Revenue: ${formatSalesMoney(ctx.parsed.x)}` } } },
            scales: {
                x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatSalesMoney(v) } },
                y: { grid: { display: false }, ticks: { color: mutedColor, font: { size: 11 } } }
            }
        }
    });

    // Top 15 Customers by Revenue
    const sortedCustomers = Object.entries(customerRevenue).sort((a, b) => b[1].revenue - a[1].revenue).slice(0, 15);
    renderChart('chartTopCustomersRevenue', {
        type: 'bar',
        data: {
            labels: sortedCustomers.map(([name]) => truncate(name, 25)),
            datasets: [{
                label: 'Revenue',
                data: sortedCustomers.map(([, d]) => d.revenue),
                backgroundColor: COLORS.palette.map(c => c + '70'),
                borderColor: COLORS.palette,
                borderWidth: 1.5,
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: { legend: { display: false }, tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8, callbacks: { label: ctx => `Revenue: ${formatSalesMoney(ctx.parsed.x)}` } } },
            scales: {
                x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatSalesMoney(v) } },
                y: { grid: { display: false }, ticks: { color: mutedColor, font: { size: 11 } } }
            }
        }
    });

    // Top Generators Table
    const tbody = document.getElementById('topGenBody');
    let cumRevenue = 0;
    tbody.innerHTML = sortedItems.slice(0, 100).map(([code, data], i) => {
        cumRevenue += data.revenue;
        const share = totalRevenue > 0 ? ((data.revenue / totalRevenue) * 100) : 0;
        const cumPct = totalRevenue > 0 ? ((cumRevenue / totalRevenue) * 100) : 0;
        const avgPrice = data.units > 0 ? data.revenue / data.units : 0;

        return `<tr style="--row-index: ${i}">
            <td class="text-center">${i + 1}</td>
            <td class="item-name">${escapeHtml(code)}</td>
            <td style="color:var(--text-secondary);max-width:200px;overflow:hidden;text-overflow:ellipsis">${escapeHtml(truncate(data.description || '—', 30))}</td>
            <td class="text-right" style="font-weight:700;color:var(--warning)">${formatSalesMoney(data.revenue)}</td>
            <td class="text-right">${formatNumber(data.units)}</td>
            <td class="text-right" style="color:var(--success)">${formatSalesMoney(avgPrice)}</td>
            <td class="text-right">
                <div class="rev-share-bar">
                    <div class="rev-share-bar-bg"><div class="rev-share-bar-fill" style="width:${Math.min(share * 2, 100)}%"></div></div>
                    <span class="rev-share-pct">${share.toFixed(1)}%</span>
                </div>
            </td>
            <td>
                <div class="rev-share-bar">
                    <div class="rev-share-bar-bg"><div class="rev-share-bar-fill cumulative" style="width:${Math.min(cumPct, 100)}%"></div></div>
                    <span class="rev-share-pct">${cumPct.toFixed(1)}%</span>
                </div>
            </td>
        </tr>`;
    }).join('');
}

// ===== 3. PRICE ANALYTICS =====
function renderSalesPriceAnalytics() {
    const { itemRevenue } = salesProcessed;
    const searchVal = (document.getElementById('priceSearchItem')?.value || '').toLowerCase();
    const sortBy = document.getElementById('priceSortBy')?.value || 'variation_desc';
    const mutedColor = getComputedStyle(document.body).getPropertyValue('--text-muted').trim();

    // Build price analytics data
    let priceData = Object.entries(itemRevenue).map(([code, data]) => {
        const prices = data.prices.filter(p => p > 0);
        if (prices.length === 0) return null;
        const minPrice = Math.min(...prices);
        const maxPrice = Math.max(...prices);
        const avgPrice = prices.reduce((s, p) => s + p, 0) / prices.length;
        const variation = avgPrice > 0 ? ((maxPrice - minPrice) / avgPrice) * 100 : 0;

        return {
            code,
            description: data.description || '—',
            minPrice,
            maxPrice,
            avgPrice,
            variation,
            units: data.units,
            revenue: data.revenue
        };
    }).filter(d => d !== null);

    // Filter
    if (searchVal) {
        priceData = priceData.filter(d => d.code.toLowerCase().includes(searchVal) || d.description.toLowerCase().includes(searchVal));
    }

    // Sort
    switch (sortBy) {
        case 'variation_desc': priceData.sort((a, b) => b.variation - a.variation); break;
        case 'avgprice_desc': priceData.sort((a, b) => b.avgPrice - a.avgPrice); break;
        case 'revenue_desc': priceData.sort((a, b) => b.revenue - a.revenue); break;
    }

    // Price Variation Chart (Top 10)
    const top10Variation = [...priceData].sort((a, b) => b.variation - a.variation).slice(0, 10);
    renderChart('chartPriceVariation', {
        type: 'bar',
        data: {
            labels: top10Variation.map(d => truncate(d.description || d.code, 20)),
            datasets: [
                {
                    label: 'Min Price',
                    data: top10Variation.map(d => d.minPrice),
                    backgroundColor: 'rgba(16, 185, 129, 0.7)',
                    borderColor: '#10b981',
                    borderWidth: 1.5,
                    borderRadius: 4,
                    borderSkipped: false
                },
                {
                    label: 'Avg Price',
                    data: top10Variation.map(d => d.avgPrice),
                    backgroundColor: 'rgba(59, 130, 246, 0.7)',
                    borderColor: '#3b82f6',
                    borderWidth: 1.5,
                    borderRadius: 4,
                    borderSkipped: false
                },
                {
                    label: 'Max Price',
                    data: top10Variation.map(d => d.maxPrice),
                    backgroundColor: 'rgba(239, 68, 68, 0.7)',
                    borderColor: '#ef4444',
                    borderWidth: 1.5,
                    borderRadius: 4,
                    borderSkipped: false
                }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: mutedColor, font: { size: 11, family: 'Outfit' }, usePointStyle: true } },
                tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8, callbacks: { label: ctx => ` ${ctx.dataset.label}: ${formatSalesMoney(ctx.parsed.y)}` } }
            },
            scales: {
                x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, maxRotation: 45, font: { size: 10 } } },
                y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatSalesMoney(v) } }
            }
        }
    });

    // Price vs Volume Scatter
    const topVolume = [...priceData].sort((a, b) => b.units - a.units).slice(0, 30);
    renderChart('chartPriceVsVolume', {
        type: 'bubble',
        data: {
            datasets: [{
                label: 'Items',
                data: topVolume.map(d => ({ x: d.units, y: d.avgPrice, r: Math.min(Math.max(Math.sqrt(d.revenue / 1000) * 2, 4), 25), label: d.code })),
                backgroundColor: 'rgba(99, 102, 241, 0.5)',
                borderColor: '#6366f1',
                borderWidth: 1.5
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8,
                    callbacks: {
                        title: ctx => ctx[0]?.raw?.label || '',
                        label: ctx => [`Volume: ${formatNumber(ctx.parsed.x)}`, `Avg Price: ${formatSalesMoney(ctx.parsed.y)}`]
                    }
                }
            },
            scales: {
                x: { title: { display: true, text: 'Total Units Sold', color: mutedColor, font: { size: 12 } }, grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatNumber(v) } },
                y: { title: { display: true, text: 'Average Price', color: mutedColor, font: { size: 12 } }, grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatSalesMoney(v) } }
            }
        }
    });

    // Price Table
    const tbody = document.getElementById('priceAnalyticsBody');
    if (priceData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" class="empty-msg">No items match your search</td></tr>';
        return;
    }

    tbody.innerHTML = priceData.slice(0, 150).map((d, i) => {
        const varClass = d.variation >= 50 ? 'high' : d.variation >= 20 ? 'medium' : 'low';
        return `<tr style="--row-index: ${i}">
            <td class="text-center">${i + 1}</td>
            <td class="item-name">${escapeHtml(d.code)}</td>
            <td style="color:var(--text-secondary);max-width:200px;overflow:hidden;text-overflow:ellipsis">${escapeHtml(truncate(d.description, 30))}</td>
            <td class="text-right" style="color:var(--success)">${formatSalesMoney(d.minPrice)}</td>
            <td class="text-right" style="color:var(--danger)">${formatSalesMoney(d.maxPrice)}</td>
            <td class="text-right" style="font-weight:600">${formatSalesMoney(d.avgPrice)}</td>
            <td class="text-right"><span class="price-var-badge ${varClass}">${d.variation.toFixed(1)}%</span></td>
            <td class="text-right">${formatNumber(d.units)}</td>
            <td class="text-right" style="font-weight:700;color:var(--warning)">${formatSalesMoney(d.revenue)}</td>
        </tr>`;
    }).join('');
}

// ===== 4. CUSTOMER REVENUE ANALYSIS =====
function renderSalesCustomerRevenue() {
    const { customerRevenue, allYearMonths, allYears } = salesProcessed;
    const selectedCust = document.getElementById('salesCustSelect')?.value || 'ALL';
    const mutedColor = getComputedStyle(document.body).getPropertyValue('--text-muted').trim();

    // Update chart title
    const chartTitle = document.getElementById('custRevenueChartTitle');
    if (chartTitle) {
        chartTitle.textContent = selectedCust === 'ALL' ? 'Overall Revenue Trend' : `Revenue Trend — ${truncate(selectedCust, 30)}`;
    }

    // Revenue trend chart
    if (selectedCust === 'ALL') {
        // Show top 5 customers as separate lines
        const top5 = Object.entries(customerRevenue).sort((a, b) => b[1].revenue - a[1].revenue).slice(0, 5);
        renderChart('chartCustRevenueTrend', {
            type: 'line',
            data: {
                labels: allYearMonths.map(m => formatMonthLabel(m)),
                datasets: top5.map(([name, data], i) => ({
                    label: truncate(name, 20),
                    data: allYearMonths.map(m => data.months[m]?.revenue || 0),
                    borderColor: COLORS.palette[i],
                    backgroundColor: COLORS.palette[i] + '15',
                    fill: false,
                    tension: 0.4,
                    pointRadius: 3,
                    pointHoverRadius: 6,
                    borderWidth: 2
                }))
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { labels: { color: mutedColor, font: { size: 11, family: 'Outfit' }, usePointStyle: true } },
                    tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8, mode: 'index', callbacks: { label: ctx => ` ${ctx.dataset.label}: ${formatSalesMoney(ctx.parsed.y)}` } }
                },
                scales: {
                    x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor } },
                    y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatSalesMoney(v) }, beginAtZero: true }
                },
                interaction: { intersect: false, mode: 'index' }
            }
        });
    } else {
        const data = customerRevenue[selectedCust];
        if (data) {
            renderChart('chartCustRevenueTrend', {
                type: 'bar',
                data: {
                    labels: allYearMonths.map(m => formatMonthLabel(m)),
                    datasets: [{
                        label: 'Revenue',
                        data: allYearMonths.map(m => data.months[m]?.revenue || 0),
                        backgroundColor: 'rgba(245, 158, 11, 0.6)',
                        borderColor: '#f59e0b',
                        borderWidth: 1.5,
                        borderRadius: 6,
                        borderSkipped: false
                    }]
                },
                options: {
                    responsive: true, maintainAspectRatio: false,
                    plugins: { legend: { display: false }, tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8, callbacks: { label: ctx => `Revenue: ${formatSalesMoney(ctx.parsed.y)}` } } },
                    scales: {
                        x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor } },
                        y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatSalesMoney(v) }, beginAtZero: true }
                    }
                }
            });
        }
    }

    // Customer table
    const tbody = document.getElementById('custRevenueBody');
    const sortedCustomers = Object.entries(customerRevenue).sort((a, b) => b[1].revenue - a[1].revenue);

    if (sortedCustomers.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="empty-msg">No customer data available</td></tr>';
        return;
    }

    tbody.innerHTML = sortedCustomers.slice(0, 100).map(([name, data], i) => {
        const txnCount = data.entries.size;
        const avgTxn = txnCount > 0 ? data.revenue / txnCount : 0;
        const avgItemsPerTxn = txnCount > 0 ? data.units / txnCount : 0;

        // Growth: compare last year vs previous year
        let growth = null;
        if (allYears.length >= 2) {
            const lastYear = allYears[allYears.length - 1];
            const prevYear = allYears[allYears.length - 2];
            const lastYearRev = Object.entries(data.months)
                .filter(([m]) => parseMonthStr(m).year === lastYear)
                .reduce((s, [, d]) => s + d.revenue, 0);
            const prevYearRev = Object.entries(data.months)
                .filter(([m]) => parseMonthStr(m).year === prevYear)
                .reduce((s, [, d]) => s + d.revenue, 0);
            if (prevYearRev > 0) {
                growth = ((lastYearRev - prevYearRev) / prevYearRev) * 100;
            }
        }

        const growthClass = growth === null ? 'neutral' : growth >= 0 ? 'positive' : 'negative';
        const growthText = growth === null ? 'N/A' : `${growth >= 0 ? '▲' : '▼'} ${Math.abs(growth).toFixed(1)}%`;

        return `<tr style="--row-index: ${i}">
            <td class="text-center">${i + 1}</td>
            <td class="customer-name">${escapeHtml(name)}</td>
            <td class="text-right" style="font-weight:700;color:var(--warning)">${formatSalesMoney(data.revenue)}</td>
            <td class="text-right">${formatNumber(data.units)}</td>
            <td class="text-right">${formatNumber(txnCount)}</td>
            <td class="text-right" style="color:var(--accent-primary)">${formatSalesMoney(avgTxn)}</td>
            <td class="text-right">${avgItemsPerTxn.toFixed(1)}</td>
            <td><span class="growth-indicator ${growthClass}">${growthText}</span></td>
        </tr>`;
    }).join('');
}

// ===== 5. PERIOD COMPARISON =====
function renderSalesPeriodComparison() {
    const { allYears, normalized, allYearMonths, monthlyRevenue } = salesProcessed;
    const mutedColor = getComputedStyle(document.body).getPropertyValue('--text-muted').trim();

    if (allYears.length < 2) {
        document.getElementById('periodCompBody').innerHTML = '<tr><td colspan="6" class="empty-msg">Need at least 2 years of data for comparison</td></tr>';
        return;
    }

    // Use the last 2 years for comparison
    const year2 = allYears[allYears.length - 1];
    const year1 = allYears[allYears.length - 2];

    setText('periodYear1Col', year1.toString());
    setText('periodYear2Col', year2.toString());

    // Build monthly data for each year
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const year1Monthly = {};
    const year2Monthly = {};

    for (let m = 1; m <= 12; m++) {
        const key1 = `${year1}-${m}`;
        const key2 = `${year2}-${m}`;
        year1Monthly[m] = monthlyRevenue[key1] || 0;
        year2Monthly[m] = monthlyRevenue[key2] || 0;
    }

    // YoY Revenue Chart
    renderChart('chartYoYRevenue', {
        type: 'bar',
        data: {
            labels: monthNames,
            datasets: [
                {
                    label: year1.toString(),
                    data: monthNames.map((_, i) => year1Monthly[i + 1]),
                    backgroundColor: 'rgba(99, 102, 241, 0.6)',
                    borderColor: '#6366f1',
                    borderWidth: 1.5,
                    borderRadius: 4,
                    borderSkipped: false
                },
                {
                    label: year2.toString(),
                    data: monthNames.map((_, i) => year2Monthly[i + 1]),
                    backgroundColor: 'rgba(245, 158, 11, 0.6)',
                    borderColor: '#f59e0b',
                    borderWidth: 1.5,
                    borderRadius: 4,
                    borderSkipped: false
                }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: mutedColor, font: { size: 12, family: 'Outfit' }, usePointStyle: true } },
                tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8, callbacks: { label: ctx => ` ${ctx.dataset.label}: ${formatSalesMoney(ctx.parsed.y)}` } }
            },
            scales: {
                x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor } },
                y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatSalesMoney(v) } }
            }
        }
    });

    // MoM Growth Rate Chart
    const growthRates = allYearMonths.map((m, i) => {
        if (i === 0) return 0;
        const prev = monthlyRevenue[allYearMonths[i - 1]] || 0;
        const curr = monthlyRevenue[m] || 0;
        return prev > 0 ? ((curr - prev) / prev) * 100 : 0;
    });

    renderChart('chartMoMGrowth', {
        type: 'bar',
        data: {
            labels: allYearMonths.map(m => formatMonthLabel(m)),
            datasets: [{
                label: 'Growth %',
                data: growthRates,
                backgroundColor: growthRates.map(g => g >= 0 ? 'rgba(16, 185, 129, 0.6)' : 'rgba(239, 68, 68, 0.6)'),
                borderColor: growthRates.map(g => g >= 0 ? '#10b981' : '#ef4444'),
                borderWidth: 1.5,
                borderRadius: 4,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8, callbacks: { label: ctx => `Growth: ${ctx.parsed.y >= 0 ? '+' : ''}${ctx.parsed.y.toFixed(1)}%` } }
            },
            scales: {
                x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, maxRotation: 45, font: { size: 10 } } },
                y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => v + '%' } }
            }
        }
    });

    // Period Comparison Table
    const tbody = document.getElementById('periodCompBody');
    tbody.innerHTML = monthNames.map((name, i) => {
        const m = i + 1;
        const rev1 = year1Monthly[m];
        const rev2 = year2Monthly[m];
        const change = rev2 - rev1;
        const growthPct = rev1 > 0 ? ((change / rev1) * 100) : (rev2 > 0 ? 100 : 0);
        const trendClass = change > 0 ? 'trend-up' : change < 0 ? 'trend-down' : 'trend-flat';
        const trendIcon = change > 0 ? '▲' : change < 0 ? '▼' : '—';

        return `<tr style="--row-index: ${i}">
            <td style="font-weight:600">${name}</td>
            <td class="text-right">${formatSalesMoney(rev1)}</td>
            <td class="text-right" style="font-weight:700">${formatSalesMoney(rev2)}</td>
            <td class="text-right"><span class="${trendClass}" style="font-weight:600">${change >= 0 ? '+' : ''}${formatSalesMoney(change)}</span></td>
            <td class="text-right"><span class="${trendClass}" style="font-weight:700">${growthPct >= 0 ? '+' : ''}${growthPct.toFixed(1)}%</span></td>
            <td><span class="trend-arrow ${trendClass}">${trendIcon} ${Math.abs(growthPct) > 50 ? 'Significant' : Math.abs(growthPct) > 20 ? 'Moderate' : 'Stable'}</span></td>
        </tr>`;
    }).join('');
}

// ===== EXPORT DROPPED ITEMS TO EXCEL =====
function exportDroppedToExcel() {
    if (!processedData || !processedData.droppedItems) {
        alert('No data to export');
        return;
    }

    const searchVal = (document.getElementById('droppedCustomerSearch')?.value || '').toLowerCase();
    const minMonths = parseInt(document.getElementById('droppedMonthsFilter')?.value || '2');
    const sortBy = document.getElementById('droppedSortBy')?.value || 'months_desc';

    // Get the same filtered & sorted data shown in the table
    let filtered = processedData.droppedItems.filter(d => d.monthsInactive >= minMonths);
    if (searchVal) filtered = filtered.filter(d => d.customer.toLowerCase().includes(searchVal));
    switch (sortBy) {
        case 'months_desc': filtered.sort((a, b) => b.monthsInactive - a.monthsInactive); break;
        case 'items_desc': filtered.sort((a, b) => b.avgQtyPerMonth - a.avgQtyPerMonth); break;
        case 'customer_asc': filtered.sort((a, b) => a.customer.localeCompare(b.customer)); break;
    }

    if (filtered.length === 0) { alert('No dropped items to export'); return; }

    // Group by customer
    const grouped = {};
    filtered.forEach(d => {
        if (!grouped[d.customer]) grouped[d.customer] = [];
        grouped[d.customer].push(d);
    });

    // ===== STYLES =====
    const titleStyle = {
        font: { bold: true, sz: 16, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '1E293B' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: { bottom: { style: 'thin', color: { rgb: '475569' } } }
    };
    const subtitleStyle = {
        font: { sz: 11, color: { rgb: 'CBD5E1' } },
        fill: { fgColor: { rgb: '1E293B' } },
        alignment: { horizontal: 'center', vertical: 'center' }
    };
    const customerHeaderStyle = {
        font: { bold: true, sz: 13, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '6366F1' } },
        alignment: { horizontal: 'left', vertical: 'center' },
        border: { bottom: { style: 'medium', color: { rgb: '4F46E5' } } }
    };
    const colHeaderStyle = {
        font: { bold: true, sz: 10, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '334155' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: { bottom: { style: 'thin', color: { rgb: '475569' } } }
    };
    const dataStyleEven = {
        font: { sz: 10 },
        fill: { fgColor: { rgb: 'F8FAFC' } },
        alignment: { vertical: 'center' },
        border: { bottom: { style: 'thin', color: { rgb: 'E2E8F0' } } }
    };
    const dataStyleOdd = {
        font: { sz: 10 },
        fill: { fgColor: { rgb: 'FFFFFF' } },
        alignment: { vertical: 'center' },
        border: { bottom: { style: 'thin', color: { rgb: 'E2E8F0' } } }
    };
    const statusLost = { font: { bold: true, sz: 10, color: { rgb: 'DC2626' } }, fill: { fgColor: { rgb: 'FEE2E2' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: { bottom: { style: 'thin', color: { rgb: 'E2E8F0' } } } };
    const statusAtRisk = { font: { bold: true, sz: 10, color: { rgb: 'D97706' } }, fill: { fgColor: { rgb: 'FEF3C7' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: { bottom: { style: 'thin', color: { rgb: 'E2E8F0' } } } };
    const statusRecent = { font: { bold: true, sz: 10, color: { rgb: 'EA580C' } }, fill: { fgColor: { rgb: 'FFEDD5' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: { bottom: { style: 'thin', color: { rgb: 'E2E8F0' } } } };
    const numStyle = (base) => ({ ...base, alignment: { ...base.alignment, horizontal: 'right' } });

    const COL_HEADERS = ['#', 'Item Code', 'Last Purchase', 'Last Qty', 'Avg Qty/Month', 'Months Inactive', 'Status'];
    const colCount = COL_HEADERS.length;

    // Build rows
    const wsData = [];
    const merges = [];

    // Row 0: Title
    const titleRow = [{ v: '📊 Dropped Items Report', s: titleStyle }];
    for (let c = 1; c < colCount; c++) titleRow.push({ v: '', s: titleStyle });
    wsData.push(titleRow);
    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } });

    // Row 1: Subtitle (date + summary)
    const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    const subRow = [{ v: `Generated: ${today}  |  ${filtered.length} items across ${Object.keys(grouped).length} customers`, s: subtitleStyle }];
    for (let c = 1; c < colCount; c++) subRow.push({ v: '', s: subtitleStyle });
    wsData.push(subRow);
    merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: colCount - 1 } });

    // Row 2: Empty spacer
    wsData.push([]);

    // For each customer group
    const customers = Object.keys(grouped).sort((a, b) => a.localeCompare(b));
    customers.forEach(customer => {
        const items = grouped[customer];
        const rowIdx = wsData.length;

        // Customer header row (merged across all columns)
        const custRow = [{ v: `👤 ${customer}  (${items.length} dropped item${items.length > 1 ? 's' : ''})`, s: customerHeaderStyle }];
        for (let c = 1; c < colCount; c++) custRow.push({ v: '', s: customerHeaderStyle });
        wsData.push(custRow);
        merges.push({ s: { r: rowIdx, c: 0 }, e: { r: rowIdx, c: colCount - 1 } });

        // Column headers
        wsData.push(COL_HEADERS.map(h => ({ v: h, s: colHeaderStyle })));

        // Data rows
        items.forEach((d, idx) => {
            const baseStyle = idx % 2 === 0 ? dataStyleEven : dataStyleOdd;
            const statusText = d.monthsInactive >= 6 ? '🔴 Lost' : d.monthsInactive >= 3 ? '🟡 At Risk' : '🟠 Recently Stopped';
            const sStyle = d.monthsInactive >= 6 ? statusLost : d.monthsInactive >= 3 ? statusAtRisk : statusRecent;

            wsData.push([
                { v: idx + 1, s: { ...baseStyle, alignment: { ...baseStyle.alignment, horizontal: 'center' } } },
                { v: d.itemCode, s: { ...baseStyle, font: { ...baseStyle.font, bold: true } } },
                { v: formatMonthLabel(d.lastMonth), s: baseStyle },
                { v: d.lastQty, s: numStyle(baseStyle), t: 'n' },
                { v: Math.round(d.avgQtyPerMonth * 10) / 10, s: numStyle(baseStyle), t: 'n' },
                { v: d.monthsInactive, s: { ...numStyle(baseStyle), font: { ...baseStyle.font, bold: true, color: { rgb: d.monthsInactive >= 6 ? 'DC2626' : d.monthsInactive >= 3 ? 'D97706' : 'EA580C' } } } , t: 'n' },
                { v: statusText, s: sStyle }
            ]);
        });

        // Blank row separator
        wsData.push([]);
    });

    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!merges'] = merges;

    // Column widths
    ws['!cols'] = [
        { wch: 5 },   // #
        { wch: 18 },  // Item Code
        { wch: 16 },  // Last Purchase
        { wch: 11 },  // Last Qty
        { wch: 14 },  // Avg Qty/Month
        { wch: 16 },  // Months Inactive
        { wch: 22 }   // Status
    ];

    // Row heights
    ws['!rows'] = [{ hpt: 32 }, { hpt: 22 }]; // Title + subtitle

    // Create workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Dropped Items');
    XLSX.writeFile(wb, 'Dropped_Items_Report_' + new Date().toISOString().slice(0, 10) + '.xlsx');
}

// ===== 6. STOCK BUILDER =====
function renderStockBuilder() {
    if (!salesProcessed) return;

    const searchVal = (document.getElementById('stockBuilderSearch')?.value || '').trim().toLowerCase();
    const resultsDiv = document.getElementById('stockBuilderResults');
    const chartsDiv = document.getElementById('stockBuilderCharts');

    if (searchVal.length < 2) {
        if (resultsDiv) resultsDiv.style.display = 'block';
        if (chartsDiv) chartsDiv.style.display = 'none';
        return;
    }

    // Find matching items (match by item_des or item_code)
    const { normalized, allYearMonths } = salesProcessed;
    const matchedRows = normalized.filter(r =>
        r.itemDes.toLowerCase().includes(searchVal) || r.itemCode.toLowerCase().includes(searchVal)
    );

    if (matchedRows.length === 0) {
        if (resultsDiv) {
            resultsDiv.innerHTML = `<div class="empty-state"><div class="empty-state-icon">🔍</div><p class="empty-state-text">No items found matching "${escapeHtml(searchVal)}"</p></div>`;
            resultsDiv.style.display = 'block';
        }
        if (chartsDiv) chartsDiv.style.display = 'none';
        return;
    }

    // Determine the primary item matched (use the most common item_code)
    const itemCodeCounts = {};
    matchedRows.forEach(r => {
        itemCodeCounts[r.itemCode] = (itemCodeCounts[r.itemCode] || 0) + 1;
    });
    const primaryCode = Object.entries(itemCodeCounts).sort((a, b) => b[1] - a[1])[0][0];
    const primaryRows = matchedRows.filter(r => r.itemCode === primaryCode);
    const primaryDesc = primaryRows[0]?.itemDes || primaryCode;

    // Aggregate qty per month
    const monthlyQty = {};
    const monthlyRevenue = {};
    const monthlyCustomers = {};
    const monthlyPrices = {};
    allYearMonths.forEach(ym => { monthlyQty[ym] = 0; monthlyRevenue[ym] = 0; monthlyCustomers[ym] = new Set(); monthlyPrices[ym] = []; });
    primaryRows.forEach(r => {
        if (!r.yearMonth) return;
        monthlyQty[r.yearMonth] = (monthlyQty[r.yearMonth] || 0) + r.totalUnits;
        monthlyRevenue[r.yearMonth] = (monthlyRevenue[r.yearMonth] || 0) + r.netAmt;
        if (!monthlyCustomers[r.yearMonth]) monthlyCustomers[r.yearMonth] = new Set();
        monthlyCustomers[r.yearMonth].add(r.customer);
        if (!monthlyPrices[r.yearMonth]) monthlyPrices[r.yearMonth] = [];
        if (r.unitPrice > 0) monthlyPrices[r.yearMonth].push(r.unitPrice);
    });

    // Filter to months that have data
    const activeMonths = allYearMonths.filter(m => monthlyQty[m] > 0);

    // Aggregate qty per customer
    const custQty = {};
    primaryRows.forEach(r => {
        if (!custQty[r.customer]) custQty[r.customer] = { qty: 0, revenue: 0, prices: [] };
        custQty[r.customer].qty += r.totalUnits;
        custQty[r.customer].revenue += r.netAmt;
        if (r.unitPrice > 0) custQty[r.customer].prices.push(r.unitPrice);
    });

    const totalQty = primaryRows.reduce((s, r) => s + r.totalUnits, 0);
    const mutedColor = getComputedStyle(document.body).getPropertyValue('--text-muted').trim();

    // Hide empty state, show charts
    if (resultsDiv) resultsDiv.style.display = 'none';
    if (chartsDiv) chartsDiv.style.display = 'block';

    // Update titles
    setText('stockBuilderChartTitle', `Monthly Qty Sold — ${truncate(primaryDesc, 40)} (${primaryCode})`);
    setText('stockCustTableTitle', `Qty Sold per Customer — ${truncate(primaryDesc, 30)}`);
    setText('stockMonthTableTitle', `Qty Sold per Month — ${truncate(primaryDesc, 30)}`);

    // Monthly Qty Chart
    renderChart('chartStockBuilderMonthly', {
        type: 'bar',
        data: {
            labels: activeMonths.map(m => formatMonthLabel(m)),
            datasets: [{
                label: 'Qty Sold',
                data: activeMonths.map(m => monthlyQty[m] || 0),
                backgroundColor: 'rgba(99, 102, 241, 0.6)',
                borderColor: '#6366f1',
                borderWidth: 1.5,
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(15,23,42,0.9)', padding: 12, cornerRadius: 8,
                    callbacks: { label: ctx => `Qty: ${formatNumber(ctx.parsed.y)}` }
                }
            },
            scales: {
                x: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, maxRotation: 45, font: { size: 10 } } },
                y: { grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: mutedColor, callback: v => formatNumber(v) }, beginAtZero: true }
            }
        }
    });

    // Customer Table
    const sortedCusts = Object.entries(custQty).sort((a, b) => b[1].qty - a[1].qty);
    const custBody = document.getElementById('stockCustBody');
    if (custBody) {
        custBody.innerHTML = sortedCusts.map(([name, data], i) => {
            const avgPrice = data.prices.length > 0 ? data.prices.reduce((s, p) => s + p, 0) / data.prices.length : 0;
            const share = totalQty > 0 ? ((data.qty / totalQty) * 100) : 0;
            return `<tr style="--row-index: ${i}">
                <td class="text-center">${i + 1}</td>
                <td class="customer-name">${escapeHtml(name)}</td>
                <td class="text-right" style="font-weight:700;color:var(--accent-primary)">${formatNumber(data.qty)}</td>
                <td class="text-right" style="color:var(--success)">${formatSalesMoney(avgPrice)}</td>
                <td class="text-right" style="font-weight:600;color:var(--warning)">${formatSalesMoney(data.revenue)}</td>
                <td>
                    <div class="rev-share-bar">
                        <div class="rev-share-bar-bg"><div class="rev-share-bar-fill" style="width:${Math.min(share, 100)}%"></div></div>
                        <span class="rev-share-pct">${share.toFixed(1)}%</span>
                    </div>
                </td>
            </tr>`;
        }).join('');
    }

    // Month Table
    const monthBody = document.getElementById('stockMonthBody');
    if (monthBody) {
        let prevQty = 0;
        monthBody.innerHTML = activeMonths.map((m, i) => {
            const qty = monthlyQty[m] || 0;
            const rev = monthlyRevenue[m] || 0;
            const custs = monthlyCustomers[m]?.size || 0;
            const prices = monthlyPrices[m] || [];
            const avgP = prices.length > 0 ? prices.reduce((s, p) => s + p, 0) / prices.length : 0;

            const trend = i === 0 ? 0 : qty - prevQty;
            const trendClass = trend > 0 ? 'trend-up' : trend < 0 ? 'trend-down' : 'trend-flat';
            const trendIcon = trend > 0 ? '▲' : trend < 0 ? '▼' : '—';
            prevQty = qty;

            return `<tr style="--row-index: ${i}">
                <td style="font-weight:600">${formatMonthLabel(m)}</td>
                <td class="text-right" style="font-weight:700;color:var(--accent-primary)">${formatNumber(qty)}</td>
                <td class="text-right" style="color:var(--success)">${formatSalesMoney(avgP)}</td>
                <td class="text-right" style="color:var(--warning)">${formatSalesMoney(rev)}</td>
                <td class="text-right">${formatNumber(custs)}</td>
                <td><span class="trend-arrow ${trendClass}">${trendIcon} ${trend !== 0 ? Math.abs(trend) : ''}</span></td>
            </tr>`;
        }).join('');
    }
}
