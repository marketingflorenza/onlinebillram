// ================================================================
// 1. CONFIGURATION
// ================================================================
const CONFIG = {
    // <<< UPDATED VALUES
    API_BASE_URL: 'https://backend-api-ram.vercel.app/api',
    SHEET_ID: '1LFoETe4YdPIxxj27bXxNRip9dKJ-sL7bbDv820ExEtk',
    SHEET_NAME_SUMMARY: 'SUM',
};

// ================================================================
// 2. UI ELEMENTS
// ================================================================
const ui = {
    funnelStatsGrid: document.getElementById('funnelStatsGrid'),
    adsStatsGrid: document.getElementById('adsStatsGrid'),
    salesOverviewStatsGrid: document.getElementById('salesOverviewStatsGrid'),
    salesRevenueStatsGrid: document.getElementById('salesRevenueStatsGrid'),
    salesBillStatsGrid: document.getElementById('salesBillStatsGrid'),
    campaignsTableBody: document.getElementById('campaignsTableBody'),
    campaignsTableHeader: document.getElementById('campaignsTableHeader'),
    errorMessage: document.getElementById('errorMessage'),
    loading: document.getElementById('loading'),
    refreshBtn: document.getElementById('refreshBtn'),
    startDate: document.getElementById('startDate'),
    endDate: document.getElementById('endDate'),
    compareToggle: document.getElementById('compareToggle'),
    compareControls: document.getElementById('compareControls'),
    compareStartDate: document.getElementById('compareStartDate'),
    compareEndDate: document.getElementById('compareEndDate'),
    modal: document.getElementById('detailsModal'),
    modalTitle: document.getElementById('modalTitle'),
    modalBody: document.getElementById('modalBody'),
    modalCloseBtn: document.querySelector('#detailsModal .modal-close-btn'),
    campaignSearchInput: document.getElementById('campaignSearchInput'),
    adSearchInput: document.getElementById('adSearchInput'),
    categoryRevenueChart: document.getElementById('categoryRevenueChart'),
    categoryDetailTableBody: document.getElementById('categoryDetailTableBody'),
    channelTableBody: document.getElementById('channelTableBody'),
    upsellPathsTableBody: document.getElementById('upsellPathsTableBody'),
};

// ================================================================
// 3. GLOBAL STATE
// ================================================================
let charts = {};
let latestCampaignData = [];
let latestCategoryDetails = [];
let latestUpsellPaths = [];
let latestFilteredSalesRows = [];
let currentPopupAds = [];
let currentSort = { key: 'spend', direction: 'desc' };
let allSalesDataCache = [];
let latestComparisonData = null;

// ================================================================
// 4. HELPER FUNCTIONS
// ================================================================
function showError(message) { ui.errorMessage.innerHTML = message; ui.errorMessage.classList.add('show'); }
function hideError() { ui.errorMessage.classList.remove('show'); }
const formatCurrency = (num) => `฿${parseFloat(num || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const formatCurrencyShort = (num) => `฿${parseInt(num || 0).toLocaleString('en-US')}`;
const formatNumber = (num) => parseInt(num || 0).toLocaleString('en-US');
const toNumber = (val) => {
    if (val === null || val === undefined || val === '') return 0;
    if (typeof val === 'number') return isFinite(val) ? val : 0;
    const n = parseFloat(String(val).replace(/[^0-9.-]+/g, ""));
    return isNaN(n) ? 0 : n;
};
function parseGvizDate(gvizDate) {
    if (!gvizDate) return null;
    const match = gvizDate.match(/Date\((\d+),(\d+),(\d+)/);
    if (match) {
        return new Date(parseInt(match[1]), parseInt(match[2]), parseInt(match[3]));
    }
    const d = new Date(gvizDate);
    return isNaN(d) ? null : d;
}
function parseCategories(categoryStr) {
    if (!categoryStr || typeof categoryStr !== 'string') return [];
    return categoryStr.split(',').map(c => c.trim()).filter(Boolean);
}
const isNewCustomer = (row) => String(row['ลูกค้าใหม่'] || '').trim().toLowerCase() === 'true' || String(row['ลูกค้าใหม่'] || '').trim() === '✔' || String(row['ลูกค้าใหม่'] || '').trim() === '1';
function calculateGrowth(current, previous) {
    if (previous === 0) {
        return current > 0 ? { percent: '∞', class: 'positive' } : { percent: '0.0%', class: '' };
    }
    const percentage = ((current - previous) / previous) * 100;
    return {
        percent: `${percentage > 0 ? '+' : ''}${percentage.toFixed(1)}%`,
        class: percentage > 0 ? 'positive' : (percentage < 0 ? 'negative' : '')
    };
}

// ================================================================
// 5. DATA FETCHING
// ================================================================
async function fetchAdsData(startDate, endDate) {
    const since = startDate.split('-').reverse().join('-');
    const until = endDate.split('-').reverse().join('-');
    // <<< CORRECTED ENDPOINT NAME
    const apiUrl = `${CONFIG.API_BASE_URL}/databillRam?since=${since}&until=${until}`;
    try {
        const response = await fetch(apiUrl);
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Ads API error (${response.status}) - ${errorText}`);
        }
        return response.json();
    } catch (error) {
         if (error instanceof TypeError && error.message.includes('fetch')) {
             throw new Error(`<b>Error: Failed to fetch API data.</b><br>This is likely a CORS issue. Please run this file on a local server (like VS Code's "Live Server" extension) to resolve.`);
         }
        throw error;
    }
}

async function fetchSalesData() {
    if (allSalesDataCache.length > 0) return allSalesDataCache;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/gviz/tq?tqx=out:json&sheet=${CONFIG.SHEET_NAME_SUMMARY}`;
    const response = await fetch(sheetUrl);
    const text = await response.text();
    const jsonStr = text.substring(text.indexOf('{'), text.lastIndexOf('}') + 1);
    const gvizData = JSON.parse(jsonStr);
    const cols = gvizData.table.cols.map(c => c.label || c.id || '');
    const salesData = gvizData.table.rows.map(r => {
        const obj = {};
        cols.forEach((col, i) => obj[col] = r.c && r.c[i] ? r.c[i].v : null);
        return obj;
    });
    allSalesDataCache = salesData;
    return salesData;
}

// ================================================================
// 6. DATA PROCESSING
// ================================================================
function linkP1AndUpP1(rows) {
    const p1Lookup = new Map();
    rows.forEach(row => {
        const phone = row['เบอร์ติดต่อ'];
        const p1Value = toNumber(row['P1']);
        const date = parseGvizDate(row['วันที่']);
        if (phone && p1Value > 0 && date) {
            const existing = p1Lookup.get(phone);
            if (!existing || date < existing.p1Date) {
                p1Lookup.set(phone, { p1Date: date, p1Categories: row['หมวดหมู่'] });
            }
        }
    });

    return rows.map(row => {
        const phone = row['เบอร์ติดต่อ'];
        const upP1Value = toNumber(row['ยอดอัพ P1']);
        const date = parseGvizDate(row['วันที่']);
        if (phone && upP1Value > 0 && date) {
            const p1Origin = p1Lookup.get(phone);
            if (p1Origin && date >= p1Origin.p1Date) {
                return { ...row, linkedP1Categories: p1Origin.p1Categories };
            }
        }
        return row;
    });
}

function calculateUpsellPaths(linkedRows) {
    const paths = {};
    linkedRows.forEach(row => {
        if (toNumber(row['ยอดอัพ P1']) > 0 && row.linkedP1Categories) {
            const fromCats = parseCategories(row.linkedP1Categories);
            const toCats = parseCategories(row['หมวดหมู่']);
            const upP1Revenue = toNumber(row['ยอดอัพ P1']);

            if (fromCats.length > 0 && toCats.length > 0) {
                const revenuePortion = upP1Revenue / (fromCats.length * toCats.length);
                fromCats.forEach(fromCat => {
                    toCats.forEach(toCat => {
                        const key = `${fromCat} -> ${toCat}`;
                        if (!paths[key]) {
                            paths[key] = { from: fromCat, to: toCat, count: 0, totalUpP1Revenue: 0, transactions: [] };
                        }
                        paths[key].count++;
                        paths[key].totalUpP1Revenue += revenuePortion;
                        paths[key].transactions.push(row);
                    });
                });
            }
        }
    });
    return Object.values(paths).sort((a, b) => b.totalUpP1Revenue - a.totalUpP1Revenue);
}

function calculateCategoryDetails(filteredRows) {
    const categoryMap = {};
    filteredRows.forEach(row => {
        const p1 = toNumber(row['P1']);
        const upP1 = toNumber(row['ยอดอัพ P1']);
        const upP2 = toNumber(row['ยอดอัพ P2']);
        const rowRevenue = p1 + upP1 + upP2;

        if (rowRevenue > 0) {
            const categories = parseCategories(row['หมวดหมู่']);
            if (categories.length > 0) {
                const p1Portion = p1 / categories.length;
                const upP1Portion = upP1 / categories.length;
                const upP2Portion = upP2 / categories.length;

                categories.forEach(catName => {
                    if (!categoryMap[catName]) {
                        categoryMap[catName] = {
                            name: catName,
                            p1Revenue: 0, upP1Revenue: 0, upP2Revenue: 0,
                            p1Bills: 0, upP1Bills: 0, upP2Bills: 0,
                            newCustomers: 0, totalRevenue: 0,
                            transactions: []
                        };
                    }
                    const category = categoryMap[catName];
                    category.p1Revenue += p1Portion;
                    category.upP1Revenue += upP1Portion;
                    category.upP2Revenue += upP2Portion;
                    category.totalRevenue += (p1Portion + upP1Portion + upP2Portion);

                    if (p1 > 0) category.p1Bills++;
                    if (upP1 > 0) category.upP1Bills++;
                    if (upP2 > 0) category.upP2Bills++;
                    if (isNewCustomer(row)) {
                        category.newCustomers++;
                    }
                    category.transactions.push(row);
                });
            }
        }
    });
    return Object.values(categoryMap).sort((a, b) => b.totalRevenue - a.totalRevenue);
}

function processSalesDataForPeriod(allSalesRows, startDate, endDate) {
    const filteredRows = allSalesRows.filter(row => {
        const d = parseGvizDate(row['วันที่']);
        return d && d >= startDate && d <= endDate;
    });
    
    const summary = { totalBills: 0, totalCustomers: 0, totalRevenue: 0, newCustomers: 0, oldCustomers: 0, p1Revenue: 0, upP1Revenue: 0, upP2Revenue: 0, p1Bills: 0, p2Leads: 0, upP1Bills: 0, upP2Bills: 0 };
    const channelBreakdown = {};
    filteredRows.forEach(row => {
        const p1 = toNumber(row['P1']);
        const upP1 = toNumber(row['ยอดอัพ P1']);
        const upP2 = toNumber(row['ยอดอัพ P2']);
        const p2 = row['P2'];
        const rowRevenue = p1 + upP1 + upP2;

        if (rowRevenue > 0) summary.totalBills++;
        if (p1 > 0) summary.p1Bills++;
        if (upP1 > 0) summary.upP1Bills++;
        if (upP2 > 0) summary.upP2Bills++;
        if (p2 !== null && p2 !== '') summary.p2Leads++;

        summary.p1Revenue += p1;
        summary.upP1Revenue += upP1;
        summary.upP2Revenue += upP2;
        summary.totalRevenue += rowRevenue;
        
        if (isNewCustomer(row)) {
            summary.newCustomers++;
        } else if (rowRevenue > 0) {
            summary.oldCustomers++;
        }

        const channel = row['ช่องทาง'];
        if (channel) {
            if (!channelBreakdown[channel]) {
                channelBreakdown[channel] = { p1: 0, p2: 0, upP2: 0, newCustomers: 0, revenue: 0 };
            }
            if (p1 > 0) channelBreakdown[channel].p1++;
            if (p2) channelBreakdown[channel].p2++;
            if (upP2 > 0) channelBreakdown[channel].upP2++;
            if (isNewCustomer(row)) channelBreakdown[channel].newCustomers++;
            channelBreakdown[channel].revenue += rowRevenue;
        }
    });
    summary.totalCustomers = summary.newCustomers + summary.oldCustomers;
    
    const linkedRows = linkP1AndUpP1(filteredRows);
    const upsellPaths = calculateUpsellPaths(linkedRows);
    const categoryDetails = calculateCategoryDetails(filteredRows);
    return { summary, categoryDetails, filteredRows, channelBreakdown, upsellPaths };
}

// ================================================================
// 7. RENDERING & POPUP FUNCTIONS
// ================================================================
function renderFunnelOverview(adsTotals, salesSummary, comparisonAdsTotals = null, comparisonSalesSummary = null) {
    const createStatCard = (label, currentVal, prevVal, isCurrency = false, isROAS = false) => {
        const displayVal = isROAS ? `${currentVal.toFixed(2)}x` : (isCurrency ? formatCurrency(currentVal) : formatNumber(currentVal));
        let comparisonHtml = '';
        if (comparisonAdsTotals || comparisonSalesSummary) {
            const growth = calculateGrowth(currentVal, prevVal);
            const prevDisplay = isROAS ? `${prevVal.toFixed(2)}x` : (isCurrency ? formatCurrency(prevVal) : formatNumber(prevVal));
            comparisonHtml = `
                <span class="growth-indicator ${growth.class}">${growth.percent}</span>
                <div class="stat-comparison">vs ${prevDisplay}</div>
            `;
        }
        return `<div class="stat-card">
                    <div class="stat-number">
                        <span>${displayVal}</span>
                        ${comparisonHtml}
                    </div>
                    <div class="stat-label">${label}</div>
                </div>`;
    };

    const spend = adsTotals.spend || 0;
    const revenue = salesSummary.totalRevenue || 0;
    const roas = spend > 0 ? revenue / spend : 0;
    const purchases = adsTotals.purchases || 0;
    const cpa = purchases > 0 ? spend / purchases : 0;
    
    const prevSpend = comparisonAdsTotals?.spend || 0;
    const prevRevenue = comparisonSalesSummary?.summary.totalRevenue || 0;
    const prevRoas = prevSpend > 0 ? prevRevenue / prevSpend : 0;
    const prevPurchases = comparisonAdsTotals?.purchases || 0;
    const prevCpa = prevPurchases > 0 ? prevSpend / prevPurchases : 0;


    ui.funnelStatsGrid.innerHTML = [
        createStatCard('Ad Spend', spend, prevSpend, true),
        createStatCard('Total Revenue', revenue, prevRevenue, true),
        createStatCard('ROAS', roas, prevRoas, false, true),
        createStatCard('Purchases', purchases, prevPurchases),
        createStatCard('Cost Per Purchase', cpa, prevCpa, true),
    ].join('');
}

function renderAdsOverview(totals) {
    const createStatCard = (label, value) => `<div class="stat-card"><div class="stat-number">${value}</div><div class="stat-label">${label}</div></div>`;
    ui.adsStatsGrid.innerHTML = [
        createStatCard('Impressions', formatNumber(totals.impressions)),
        createStatCard('Messaging Started', formatNumber(totals.messaging_conversations)),
        createStatCard('Avg. CPM', formatCurrency(totals.cpm)),
        createStatCard('Avg. CTR', `${parseFloat(totals.ctr || 0).toFixed(2)}%`)
    ].join('');
}

function renderSalesOverview(summary, comparisonSummary = null) {
    const createStatCard = (label, currentVal, prevVal, isCurrency = false) => {
        const displayVal = isCurrency ? formatCurrency(currentVal) : formatNumber(currentVal);
        let comparisonHtml = '';
        if (comparisonSummary) {
            const growth = calculateGrowth(currentVal, prevVal);
            const prevDisplay = isCurrency ? formatCurrency(prevVal) : formatNumber(prevVal);
            comparisonHtml = `
                <span class="growth-indicator ${growth.class}">${growth.percent}</span>
                <div class="stat-comparison">vs ${prevDisplay}</div>
            `;
        }
        return `<div class="stat-card">
                    <div class="stat-number">
                        <span>${displayVal}</span>
                        ${comparisonHtml}
                    </div>
                    <div class="stat-label">${label}</div>
                </div>`;
    };
    
    ui.salesOverviewStatsGrid.innerHTML = [
        createStatCard('Total Bills', summary.totalBills, comparisonSummary?.totalBills || 0),
        createStatCard('Total Sales Revenue', summary.totalRevenue, comparisonSummary?.totalRevenue || 0, true),
        createStatCard('Total Customers', summary.totalCustomers, comparisonSummary?.totalCustomers || 0),
        createStatCard('New Customers', summary.newCustomers, comparisonSummary?.newCustomers || 0),
    ].join('');
}

function renderSalesRevenueBreakdown(summary, comparisonSummary = null) {
    const createStatCard = (label, currentVal, prevVal) => {
        const displayVal = formatCurrency(currentVal);
        let comparisonHtml = '';
        if (comparisonSummary) {
            const growth = calculateGrowth(currentVal, prevVal);
            const prevDisplay = formatCurrency(prevVal);
            comparisonHtml = `
                <span class="growth-indicator ${growth.class}">${growth.percent}</span>
                <div class="stat-comparison">vs ${prevDisplay}</div>
            `;
        }
        return `<div class="stat-card">
                    <div class="stat-number">
                        <span>${displayVal}</span>
                        ${comparisonHtml}
                    </div>
                    <div class="stat-label">${label}</div>
                </div>`;
    };
    ui.salesRevenueStatsGrid.innerHTML = [
        createStatCard('P1 Revenue', summary.p1Revenue, comparisonSummary?.p1Revenue || 0),
        createStatCard('UP P1 Revenue', summary.upP1Revenue, comparisonSummary?.upP1Revenue || 0),
        createStatCard('UP P2 Revenue', summary.upP2Revenue, comparisonSummary?.upP2Revenue || 0),
    ].join('');
    charts.revenue.data.datasets[0].data = [summary.p1Revenue, summary.upP1Revenue, summary.upP2Revenue];
    charts.customer.data.datasets[0].data = [summary.newCustomers, summary.oldCustomers];
    charts.revenue.update();
    charts.customer.update();
}

function renderSalesBillStats(summary, comparisonSummary = null) {
    const createStatCard = (label, currentVal, prevVal, isRate = false) => {
        const displayVal = isRate ? `${currentVal.toFixed(1)}%` : formatNumber(currentVal);
        let comparisonHtml = '';
        if (comparisonSummary) {
            const growth = calculateGrowth(currentVal, prevVal);
            const prevDisplay = isRate ? `${prevVal.toFixed(1)}%` : formatNumber(prevVal);
            comparisonHtml = `
                <span class="growth-indicator ${growth.class}">${growth.percent}</span>
                <div class="stat-comparison">vs ${prevDisplay}</div>
            `;
        }
         return `<div class="stat-card">
                    <div class="stat-number">
                        <span>${displayVal}</span>
                        ${comparisonHtml}
                    </div>
                    <div class="stat-label">${label}</div>
                </div>`;
    };
    const p1ToUpP1Rate = summary.p1Bills > 0 ? (summary.upP1Bills / summary.p1Bills) * 100 : 0;
    const p2ConversionRate = summary.p2Leads > 0 ? (summary.upP2Bills / summary.p2Leads) * 100 : 0;
    
    let prevP1ToUpP1Rate = 0, prevP2ConversionRate = 0;
    if(comparisonSummary){
        prevP1ToUpP1Rate = comparisonSummary.p1Bills > 0 ? (comparisonSummary.upP1Bills / comparisonSummary.p1Bills) * 100 : 0;
        prevP2ConversionRate = comparisonSummary.p2Leads > 0 ? (comparisonSummary.upP2Bills / comparisonSummary.p2Leads) * 100 : 0;
    }

    ui.salesBillStatsGrid.innerHTML = [
        createStatCard('P1 Bills', summary.p1Bills, comparisonSummary?.p1Bills || 0),
        createStatCard('P2 Leads', summary.p2Leads, comparisonSummary?.p2Leads || 0),
        createStatCard('UP P1 Bills', summary.upP1Bills, comparisonSummary?.upP1Bills || 0),
        createStatCard('UP P2 Bills', summary.upP2Bills, comparisonSummary?.upP2Bills || 0),
        createStatCard('P1 → UP P1 Rate', p1ToUpP1Rate, prevP1ToUpP1Rate, true),
        createStatCard('P2 Conversion Rate', p2ConversionRate, prevP2ConversionRate, true),
    ].join('');
}

function renderChannelTable(channelData) {
    const tableBody = ui.channelTableBody;
    if (!channelData || Object.keys(channelData).length === 0) {
        tableBody.innerHTML = `<tr><td colspan="6" style="text-align:center;">ไม่พบข้อมูลช่องทาง</td></tr>`;
        return;
    }

    const renderClickableNumber = (count, channel, metric) => {
        if (count > 0) {
            const safeChannel = channel.replace(/'/g, "\\'");
            return `<span class="clickable-cell" onclick="showChannelDetailsPopup('${safeChannel}', '${metric}')">${count.toLocaleString()}</span>`;
        }
        return count.toLocaleString();
    };
    
    const renderClickableCurrency = (amount, channel, metric) => {
        const roundedAmount = Math.round(amount);
        if (roundedAmount > 0) {
            const safeChannel = channel.replace(/'/g, "\\'");
            return `<span class="clickable-cell" onclick="showChannelDetailsPopup('${safeChannel}', '${metric}')">฿${roundedAmount.toLocaleString()}</span>`;
        }
        return `฿${roundedAmount.toLocaleString()}`;
    };

    const totals = { p1: 0, p2: 0, upP2: 0, newCustomers: 0, revenue: 0 };
    
    const sortedChannels = Object.keys(channelData).sort((a, b) => (channelData[b].revenue || 0) - (channelData[a].revenue || 0));

    let tableHtml = sortedChannels.map(channel => {
        const data = channelData[channel];
        
        totals.p1 += data.p1;
        totals.p2 += data.p2;
        totals.upP2 += data.upP2;
        totals.newCustomers += data.newCustomers;
        totals.revenue += data.revenue;

        return `
            <tr>
                <td><strong>${channel}</strong></td>
                <td>${renderClickableNumber(data.p1, channel, 'P1_BILLS')}</td>
                <td>${renderClickableNumber(data.p2, channel, 'P2_LEADS')}</td>
                <td>${renderClickableNumber(data.upP2, channel, 'UP_P2_BILLS')}</td>
                <td>${renderClickableNumber(data.newCustomers, channel, 'NEW_CUSTOMERS')}</td>
                <td class="revenue-cell">${renderClickableCurrency(data.revenue, channel, 'REVENUE')}</td>
            </tr>
        `;
    }).join('');
    
    tableHtml += `
        <tr style="font-weight: bold; border-top: 2px solid var(--neon-cyan);">
            <td>รวมทั้งหมด</td>
            <td>${totals.p1.toLocaleString()}</td>
            <td>${totals.p2.toLocaleString()}</td>
            <td>${totals.upP2.toLocaleString()}</td>
            <td>${totals.newCustomers.toLocaleString()}</td>
            <td class="revenue-cell">฿${Math.round(totals.revenue).toLocaleString()}</td>
        </tr>
    `;

    tableBody.innerHTML = tableHtml;
}

function renderCampaignsTable(campaigns) {
    if (!campaigns || campaigns.length === 0) {
        ui.campaignsTableBody.innerHTML = `<tr><td colspan="7" style="text-align:center;">No campaign data found</td></tr>`;
        return;
    }
    ui.campaignsTableBody.innerHTML = campaigns.map(c => {
        const insights = c.insights || { spend: 0, impressions: 0, purchases: 0, messaging_conversations: 0, cpm: 0 };
        return `
            <tr>
                <td><a href="#" onclick="showAdDetails('${c.id}'); return false;"><strong>${c.name || 'N/A'}</strong></a></td>
                <td><span style="color:${c.status === 'ACTIVE' ? 'var(--color-positive)' : 'var(--text-secondary)'}">${c.status || 'N/A'}</span></td>
                <td class="revenue-cell">${formatCurrency(insights.spend)}</td>
                <td>${formatNumber(insights.impressions)}</td>
                <td>${formatNumber(insights.purchases)}</td>
                <td>${formatNumber(insights.messaging_conversations)}</td>
                <td>${formatCurrency(insights.cpm)}</td>
            </tr>
        `;
    }).join('');
}

function renderCategoryChart(categoryData) {
    const chart = charts.categoryRevenue;
    const topData = categoryData.slice(0, 15);
    chart.data.labels = topData.map(d => d.name);
    chart.data.datasets[0].data = topData.map(d => d.totalRevenue);
    chart.update();
}

function renderCategoryDetailTable(categoryDetails) {
    const rankClasses = ['gold', 'silver', 'bronze'];
    ui.categoryDetailTableBody.innerHTML = categoryDetails.map((cat, index) => {
        const safeCategoryName = cat.name.replace(/'/g, "\\'");
        return `
        <tr class="clickable-row" onclick="showCategoryDetailsPopup('${safeCategoryName}', 'ALL')">
            <td class="rank-column"><span class="rank-badge ${index < 3 ? rankClasses[index] : ''}">${index + 1}</span></td>
            <td data-label="Category"><strong>${cat.name}</strong></td>
            <td data-label="P1 Bills">
                <span class="clickable-cell" onclick="event.stopPropagation(); showCategoryDetailsPopup('${safeCategoryName}', 'P1')">${formatNumber(cat.p1Bills)}</span>
                <small class="sub-revenue">${formatCurrencyShort(cat.p1Revenue)}</small>
            </td>
            <td data-label="UP P1 Bills">
                <span class="clickable-cell" onclick="event.stopPropagation(); showCategoryDetailsPopup('${safeCategoryName}', 'UP_P1')">${formatNumber(cat.upP1Bills)}</span>
                <small class="sub-revenue">${formatCurrencyShort(cat.upP1Revenue)}</small>
            </td>
            <td data-label="UP P2 Bills">
                <span class="clickable-cell" onclick="event.stopPropagation(); showCategoryDetailsPopup('${safeCategoryName}', 'UP_P2')">${formatNumber(cat.upP2Bills)}</span>
                <small class="sub-revenue">${formatCurrencyShort(cat.upP2Revenue)}</small>
            </td>
            <td data-label="New Customers"><span class="clickable-cell" onclick="event.stopPropagation(); showCategoryDetailsPopup('${safeCategoryName}', 'NEW_CUSTOMER')">${formatNumber(cat.newCustomers)}</span></td>
            <td data-label="Total Revenue" class="revenue-cell">${formatCurrency(cat.totalRevenue)}</td>
        </tr>
    `}).join('');
}

function renderUpsellPaths(paths) {
    if (!paths || paths.length === 0) {
        ui.upsellPathsTableBody.innerHTML = `<tr><td colspan="6" style="text-align:center;">ไม่พบข้อมูล Upsell ในช่วงเวลานี้</td></tr>`;
        return;
    }
    const rankClasses = ['gold', 'silver', 'bronze'];
    ui.upsellPathsTableBody.innerHTML = paths.map((path, index) => {
        const pathKey = `${path.from} -> ${path.to}`.replace(/'/g, "\\'");
        return `
            <tr>
                <td class="rank-column"><span class="rank-badge ${index < 3 ? rankClasses[index] : ''}">${index+1}</span></td>
                <td>${path.from}</td>
                <td>${path.to}</td>
                <td>${formatNumber(path.count)}</td>
                <td class="revenue-cell">${formatCurrency(path.totalUpP1Revenue)}</td>
                <td><button class="btn" style="padding: 4px 12px; font-size: 0.8em;" onclick="showUpsellPathDetails('${pathKey}')">ดู</button></td>
            </tr>
        `;
    }).join('');
}

function showChannelDetailsPopup(channelName, metricType) {
    let filteredTransactions = [];
    let title = '';
    const allTransactions = latestFilteredSalesRows.filter(row => row['ช่องทาง'] === channelName);

    switch (metricType) {
        case 'P1_BILLS':
            title = `P1 Bills from: ${channelName}`;
            filteredTransactions = allTransactions.filter(row => toNumber(row['P1']) > 0);
            break;
        case 'P2_LEADS':
            title = `P2 Leads from: ${channelName}`;
            filteredTransactions = allTransactions.filter(row => row['P2'] !== null && row['P2'] !== '');
            break;
        case 'UP_P2_BILLS':
            title = `UP P2 Bills from: ${channelName}`;
            filteredTransactions = allTransactions.filter(row => toNumber(row['ยอดอัพ P2']) > 0);
            break;
        case 'NEW_CUSTOMERS':
            title = `New Customers from: ${channelName}`;
            filteredTransactions = allTransactions.filter(row => isNewCustomer(row));
            break;
        case 'REVENUE':
            title = `All Revenue Bills from: ${channelName}`;
            filteredTransactions = allTransactions.filter(row => (toNumber(row['P1']) + toNumber(row['ยอดอัพ P1']) + toNumber(row['ยอดอัพ P2'])) > 0);
            break;
        default:
            title = `Details for ${channelName}`;
            filteredTransactions = allTransactions;
    }

    ui.modalTitle.textContent = title;
    ui.adSearchInput.style.display = 'none';

    if (filteredTransactions.length === 0) {
        ui.modalBody.innerHTML = '<p style="text-align:center; grid-column: 1 / -1;">No matching transactions found.</p>';
    } else {
        const tableRows = filteredTransactions
            .sort((a,b) => parseGvizDate(b['วันที่']) - parseGvizDate(a['วันที่']))
            .map((row, index) => {
                const p1 = toNumber(row['P1']);
                const upP1 = toNumber(row['ยอดอัพ P1']);
                const upP2 = toNumber(row['ยอดอัพ P2']);
                let billTypes = [];
                if (p1 > 0) billTypes.push('P1');
                if (upP1 > 0) billTypes.push('UP P1');
                if (upP2 > 0) billTypes.push('UP P2');
                if (row['P2']) billTypes.push('P2 Lead');
                return `
                <tr>
                    <td>${index + 1}</td>
                    <td>${new Date(parseGvizDate(row['วันที่'])).toLocaleDateString('th-TH')}</td>
                    <td>${row['ชื่อลูกค้า'] || 'N/A'}</td>
                    <td>${row['หมวดหมู่'] || 'N/A'}</td>
                    <td>${billTypes.join(', ') || 'N/A'}</td>
                    <td class="revenue-cell">${formatCurrency(p1+upP1+upP2)}</td>
                </tr>
                `;
        }).join('');
        ui.modalBody.classList.add('table-view');
        ui.modalBody.innerHTML = `
            <div class="top-categories-table">
                <table>
                    <thead>
                        <tr>
                            <th>ลำดับ</th>
                            <th>Date</th>
                            <th>Customer Name</th>
                            <th>รายการ</th>
                            <th>ประเภท</th>
                            <th>Total Revenue</th>
                        </tr>
                    </thead>
                    <tbody>${tableRows}</tbody>
                </table>
            </div>
        `;
    }
    ui.modal.classList.add('show');
}

function showCategoryDetailsPopup(categoryName, filterType = 'ALL') {
    const categoryData = latestCategoryDetails.find(cat => cat.name === categoryName);
    if (!categoryData) return;

    let filteredTransactions = categoryData.transactions;
    let title = `All Transactions for: ${categoryName}`;

    switch (filterType) {
        case 'P1':
            title = `P1 Bills for: ${categoryName}`;
            filteredTransactions = categoryData.transactions.filter(row => toNumber(row['P1']) > 0);
            break;
        case 'UP_P1':
            title = `UP P1 Bills for: ${categoryName}`;
            filteredTransactions = categoryData.transactions.filter(row => toNumber(row['ยอดอัพ P1']) > 0);
            break;
        case 'UP_P2':
            title = `UP P2 Bills for: ${categoryName}`;
            filteredTransactions = categoryData.transactions.filter(row => toNumber(row['ยอดอัพ P2']) > 0);
            break;
        case 'NEW_CUSTOMER':
            title = `New Customers for: ${categoryName}`;
            filteredTransactions = categoryData.transactions.filter(row => isNewCustomer(row));
            break;
    }

    ui.modalTitle.textContent = title;
    ui.adSearchInput.style.display = 'none';
    
    if (filteredTransactions.length === 0) {
        ui.modalBody.innerHTML = '<p style="text-align:center; grid-column: 1 / -1;">No matching transactions found.</p>';
    } else {
        const tableRows = filteredTransactions
            .sort((a,b) => parseGvizDate(b['วันที่']) - parseGvizDate(a['วันที่']))
            .map((row, index) => {
                const p1 = toNumber(row['P1']);
                const upP1 = toNumber(row['ยอดอัพ P1']);
                const upP2 = toNumber(row['ยอดอัพ P2']);
                
                let billTypes = [];
                if (p1 > 0) billTypes.push('P1');
                if (upP1 > 0) billTypes.push('UP P1');
                if (upP2 > 0) billTypes.push('UP P2');

                return `
                <tr>
                    <td>${index + 1}</td>
                    <td>${new Date(parseGvizDate(row['วันที่'])).toLocaleDateString('th-TH')}</td>
                    <td>${row['ชื่อลูกค้า'] || 'N/A'}</td>
                    <td>${row['หมวดหมู่'] || 'N/A'}</td>
                    <td>${billTypes.join(', ') || 'N/A'}</td>
                    <td class="revenue-cell">${formatCurrency(p1+upP1+upP2)}</td>
                </tr>
                `;
        }).join('');
        ui.modalBody.classList.add('table-view');
        ui.modalBody.innerHTML = `
            <div class="top-categories-table">
                <table>
                    <thead>
                        <tr>
                            <th>ลำดับ</th>
                            <th>Date</th>
                            <th>Customer Name</th>
                            <th>รายการ</th>
                            <th>ประเภทบิล</th>
                            <th>Total Revenue</th>
                        </tr>
                    </thead>
                    <tbody>${tableRows}</tbody>
                </table>
            </div>
        `;
    }
    ui.modal.classList.add('show');
}

function showUpsellPathDetails(pathKey) {
    const pathData = latestUpsellPaths.find(p => `${p.from} -> ${p.to}` === pathKey);
    if (!pathData || !pathData.transactions) { return; }

    const transactions = pathData.transactions;
    ui.modalTitle.textContent = `Upsell Details: ${pathKey}`;
    ui.adSearchInput.style.display = 'none';

    if (transactions.length === 0) {
        ui.modalBody.innerHTML = '<p style="text-align:center; grid-column: 1 / -1;">No transaction details found.</p>';
    } else {
        const tableRows = transactions
            .sort((a, b) => parseGvizDate(b['วันที่']) - parseGvizDate(a['วันที่']))
            .map(row => {
                const upP1 = toNumber(row['ยอดอัพ P1']);
                return `
                <tr>
                    <td>${new Date(parseGvizDate(row['วันที่'])).toLocaleDateString('th-TH')}</td>
                    <td>${row['ชื่อลูกค้า'] || 'N/A'}</td>
                    <td>${row['หมวดหมู่'] || 'N/A'}</td>
                    <td class="revenue-cell">${formatCurrency(upP1)}</td>
                </tr>
                `;
            }).join('');

        ui.modalBody.classList.add('table-view');
        ui.modalBody.innerHTML = `
            <div class="top-categories-table">
                <table>
                    <thead>
                        <tr>
                            <th>Date</th>
                            <th>Customer Name</th>
                            <th>UP P1 Categories</th>
                            <th>UP P1 Revenue</th>
                        </tr>
                    </thead>
                    <tbody>${tableRows}</tbody>
                </table>
            </div>
        `;
    }
    ui.modal.classList.add('show');
}

function sortAndRenderCampaigns() {
    const { key, direction } = currentSort;
    const searchTerm = ui.campaignSearchInput.value.toLowerCase();
    
    if (!latestCampaignData) {
        latestCampaignData = [];
    }

    let filteredData = latestCampaignData.filter(campaign => 
        campaign.name.toLowerCase().includes(searchTerm)
    );

    filteredData.sort((a, b) => {
        let valA, valB;
        if (key === 'name' || key === 'status') {
            valA = a[key]?.toLowerCase() || '';
            valB = b[key]?.toLowerCase() || '';
        } else {
            valA = parseFloat(a.insights?.[key] || 0);
            valB = parseFloat(b.insights?.[key] || 0);
        }
        if (valA < valB) return direction === 'asc' ? -1 : 1;
        if (valA > valB) return direction === 'asc' ? 1 : -1;
        return 0;
    });

    document.querySelectorAll('#campaignsTableHeader .sort-link').forEach(link => {
        link.classList.remove('asc', 'desc');
        if (link.dataset.sort === key) {
            link.classList.add(direction);
        }
    });

    renderCampaignsTable(filteredData);
}

function showAdDetails(campaignId) {
    const campaign = latestCampaignData.find(c => c.id === campaignId);
    if (!campaign) return;
    ui.modalTitle.textContent = `Ads in: ${campaign.name}`;
    ui.adSearchInput.value = '';
    currentPopupAds = campaign.ads || [];
    renderPopupAds(currentPopupAds);
    ui.modalBody.classList.remove('table-view');
    ui.adSearchInput.style.display = 'block';
    ui.modal.classList.add('show');
}

function renderPopupAds(ads) {
     if (!ads || ads.length === 0) {
        ui.modalBody.innerHTML = `<p style="text-align: center; grid-column: 1 / -1;">No ads found for this campaign.</p>`;
    } else {
        ui.modalBody.innerHTML = ads
            .sort((a,b) => (b.insights.spend || 0) - (a.insights.spend || 0))
            .map(ad => {
                const insights = ad.insights || { spend: 0, impressions: 0, purchases: 0, messaging_conversations: 0, cpm: 0 };
                return `
                <div class="ad-card">
                    <div class="ad-card-image">
                        <img src="${ad.thumbnail_url}" alt="Ad thumbnail" onerror="this.src='https://placehold.co/120x120/0d0c1d/a0a0b0?text=No+Image'">
                    </div>
                    <div class="ad-card-details">
                        <h4>${ad.name}</h4>
                        <div class="ad-card-stats">
                            <div>Spend: <span>${formatCurrency(insights.spend)}</span></div>
                            <div>Impressions: <span>${formatNumber(insights.impressions)}</span></div>
                            <div>Purchases: <span>${formatNumber(insights.purchases)}</span></div>
                            <div>Messaging: <span>${formatNumber(insights.messaging_conversations)}</span></div>
                            <div>CPM: <span>${formatCurrency(insights.cpm)}</span></div>
                        </div>
                    </div>
                </div>
            `}).join('');
    }
}

function initializeModal() {
    const closeModal = () => {
        ui.modal.classList.remove('show');
        ui.modalBody.classList.remove('table-view'); 
    }
    ui.modalCloseBtn.addEventListener('click', closeModal);
    ui.modal.addEventListener('click', (event) => {
        if (event.target === ui.modal) closeModal();
    });
}

function initializeCharts() {
    const textColor = '#e0e0e0';
    const gridColor = 'rgba(224, 224, 224, 0.1)';
    const categoryColors = ['#3B82F6', '#EC4899', '#84CC16', '#F59E0B', '#10B981', '#6366F1', '#D946EF', '#F97316', '#06B6D4', '#EAB308'].map(c => c + 'CC');
    
    charts.dailySpend = new Chart(document.getElementById('dailySpendChart').getContext('2d'), {
        type: 'line', data: { labels: [], datasets: [{ label: 'Spend (THB)', data: [], borderColor: '#00f2fe', backgroundColor: 'rgba(0, 242, 254, 0.1)', fill: true, tension: 0.3 }] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: textColor }, grid: { color: gridColor } }, y: { beginAtZero: true, ticks: { color: textColor, callback: v => '฿' + v.toLocaleString() }, grid: { color: gridColor } } } }
    });

    charts.revenue = new Chart(document.getElementById('revenueChart').getContext('2d'), {
        type: 'bar', data: { labels: ['P1', 'UP P1', 'UP P2'], datasets: [{ label: 'Sales (THB)', data: [], backgroundColor: ['#3B82F6', '#EC4899', '#84CC16'] }] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { color: textColor }, grid: { color: gridColor } }, x: { ticks: { color: textColor }, grid: { color: 'transparent' } } } }
    });

    charts.customer = new Chart(document.getElementById('customerChart').getContext('2d'), {
        type: 'doughnut', data: { labels:['New Customers','Old Customers'], datasets:[{ data:[], backgroundColor: ['#F59E0B', '#10B981'], borderColor: '#0d0c1d' }] },
        options: { responsive:true, maintainAspectRatio:false, plugins: { legend: { position: 'right', labels: { color: textColor } } } }
    });

    charts.categoryRevenue = new Chart(ui.categoryRevenueChart.getContext('2d'), {
        type: 'bar',
        data: { labels: [], datasets: [{ label: 'Revenue (THB)', data: [], backgroundColor: categoryColors }] },
        options: {
            responsive: true, maintainAspectRatio: false, indexAxis: 'y',
            plugins: { legend: { display: false } },
            scales: {
                x: { beginAtZero: true, ticks: { color: textColor, callback: v => '฿' + (v / 1000) + 'K' }, grid: { color: gridColor } },
                y: { ticks: { color: textColor }, grid: { color: 'transparent' } }
            }
        }
    });
}

// ================================================================
// 8. MAIN LOGIC & EVENT LISTENERS
// ================================================================
async function main() {
    ui.loading.classList.add('show');
    hideError();
    try {
        const startDateStr = ui.startDate.value;
        const endDateStr = ui.endDate.value;
         if (!startDateStr || !endDateStr) {
            setDefaultDates();
            return main(); // Re-run main after setting dates
        }
        const isCompareMode = ui.compareToggle.checked;
        
        const fetchPromises = [
            fetchAdsData(startDateStr, endDateStr),
            fetchSalesData()
        ];

        if (isCompareMode) {
            const compareStartDateStr = ui.compareStartDate.value;
            const compareEndDateStr = ui.compareEndDate.value;
            if (compareStartDateStr && compareEndDateStr) {
                fetchPromises.push(fetchAdsData(compareStartDateStr, compareEndDateStr));
            }
        }
        
        const results = await Promise.all(fetchPromises);

        const adsResponse = results[0];
        const allSalesRows = results[1];
        const comparisonAdsResponse = results.length > 2 ? results[2] : null;

         if (!adsResponse || !adsResponse.success) {
             throw new Error(adsResponse.error || 'Unknown API error from main Ads fetch.');
        }

        const currentStartDate = new Date(startDateStr + 'T00:00:00');
        const currentEndDate = new Date(endDateStr + 'T23:59:59');
        const salesData = processSalesDataForPeriod(allSalesRows, currentStartDate, currentEndDate);
        latestCategoryDetails = salesData.categoryDetails;
        latestUpsellPaths = salesData.upsellPaths;
        latestFilteredSalesRows = salesData.filteredRows;
        
        let comparisonSalesData = null;
        if (isCompareMode && comparisonAdsResponse && comparisonAdsResponse.success) {
            const compareStartDate = new Date(ui.compareStartDate.value + 'T00:00:00');
            const compareEndDate = new Date(ui.compareEndDate.value + 'T23:59:59');
            comparisonSalesData = processSalesDataForPeriod(allSalesRows, compareStartDate, compareEndDate);
            latestComparisonData = comparisonSalesData;
        } else {
            latestComparisonData = null;
        }
        
        latestCampaignData = adsResponse.data.campaigns || [];
        
        renderFunnelOverview(adsResponse.totals, salesData.summary, comparisonAdsResponse?.totals, comparisonSalesData);
        renderAdsOverview(adsResponse.totals);
        renderSalesOverview(salesData.summary, comparisonSalesData?.summary);
        renderSalesRevenueBreakdown(salesData.summary, comparisonSalesData?.summary);
        renderSalesBillStats(salesData.summary, comparisonSalesData?.summary);
        
        sortAndRenderCampaigns();
        renderCategoryChart(salesData.categoryDetails);
        renderCategoryDetailTable(salesData.categoryDetails);
        renderChannelTable(salesData.channelBreakdown);
        renderUpsellPaths(salesData.upsellPaths);
        
        if (adsResponse.data.dailySpend) {
            charts.dailySpend.data.labels = adsResponse.data.dailySpend.map(d => `${new Date(d.date).getUTCDate()}/${new Date(d.date).getUTCMonth() + 1}`);
            charts.dailySpend.data.datasets[0].data = adsResponse.data.dailySpend.map(d => d.spend);
            charts.dailySpend.update();
        }

    } catch (err) {
        showError(`${err.message || 'An unexpected error occurred.'}`);
        console.error(err);
    } finally {
        ui.loading.classList.remove('show');
    }
}

function setDefaultDates() {
    const today = new Date();
    const firstDayThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    const lastDayLastMonth = new Date(today.getFullYear(), today.getMonth(), 0);
    const firstDayLastMonth = new Date(lastDayLastMonth.getFullYear(), lastDayLastMonth.getMonth(), 1);
    
    const toInputFormat = (date) => date.toISOString().split('T')[0];

    ui.endDate.value = toInputFormat(today);
    ui.startDate.value = toInputFormat(firstDayThisMonth);
    ui.compareEndDate.value = toInputFormat(lastDayLastMonth);
    ui.compareStartDate.value = toInputFormat(firstDayLastMonth);
}

document.addEventListener('DOMContentLoaded', () => {
    initializeCharts();
    initializeModal();
    setDefaultDates();
    main();

    ui.refreshBtn.addEventListener('click', main);
    const dateInputs = [ui.startDate, ui.endDate, ui.compareStartDate, ui.compareEndDate, ui.compareToggle];
    dateInputs.forEach(input => input.addEventListener('change', () => {
        if (ui.startDate.value && ui.endDate.value) {
            main();
        }
    }));

    ui.compareToggle.addEventListener('change', () => {
         ui.compareControls.classList.toggle('show', ui.compareToggle.checked);
         main();
    });

    ui.campaignSearchInput.addEventListener('input', sortAndRenderCampaigns);
    ui.adSearchInput.addEventListener('input', (e) => {
        const searchTerm = e.target.value.toLowerCase();
        const filteredAds = currentPopupAds.filter(ad => ad.name.toLowerCase().includes(searchTerm));
        renderPopupAds(filteredAds);
    });
    ui.campaignsTableHeader.addEventListener('click', (e) => {
        e.preventDefault();
        const link = e.target.closest('.sort-link');
        if (!link) return;
        const sortKey = link.dataset.sort;
        if (currentSort.key === sortKey) {
            currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
        } else {
            currentSort.key = sortKey;
            currentSort.direction = 'desc';
        }
        sortAndRenderCampaigns();
    });
});
