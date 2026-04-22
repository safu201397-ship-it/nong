document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const uploadView = document.getElementById('upload-view');
    const loaderView = document.getElementById('loader-view');
    const dashboardView = document.getElementById('dashboard-view');
    const btnReupload = document.getElementById('btn-reupload');
    const btnExportDashboard = document.getElementById('btn-export');
    const btnExportTable = document.getElementById('btn-export-table');
    const btnExportPivot = document.getElementById('btn-export-pivot');
    const btnExportFactorCount = document.getElementById('btn-export-factor-count');

    // Chart instances
    let trendChart, reasonChart, factorChart, factorCountChart, productChart;
    let lastRenderData = {}; // Stores latest parsed data for redraw

    // Global Color Mapping System
    window.chartColors = { reasons: {}, factors: {} };
    const P_COLORS = ['#EF4444', '#F97316', '#F59E0B', '#10B981', '#3B82F6', '#8B5CF6', '#EC4899', '#14B8A6', '#84CC16', '#F43F5E', '#A855F7', '#EAB308'];
    let reasonCIdx = 0; let factorCIdx = 0;

    function getReasonColor(name) {
        if (!name || name === '-') return '#94A3B8';
        if (!window.chartColors.reasons[name]) {
            window.chartColors.reasons[name] = P_COLORS[reasonCIdx % P_COLORS.length];
            reasonCIdx++;
        }
        return window.chartColors.reasons[name];
    }
    
    function getFactorColor(name) {
        if (!name || name === '-' || name === '未標明/其他') return '#64748B';
        if (!window.chartColors.factors[name]) {
            window.chartColors.factors[name] = P_COLORS[factorCIdx % P_COLORS.length];
            factorCIdx++;
        }
        return window.chartColors.factors[name];
    }

    // View Switching
    function showView(viewId) {
        document.querySelectorAll('.view-section').forEach(el => el.classList.remove('active'));
        document.getElementById(viewId).classList.add('active');
    }

    // Initialize File Drag and Drop
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFile(e.target.files[0]);
        }
    });

    function handleFile(file) {
        showView('loader-view');
        
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            processWorkbook(workbook);
        };
        reader.onerror = () => {
            alert('檔案讀取失敗！');
            showView('upload-view');
        }
        reader.readAsArrayBuffer(file);
    }

    function processWorkbook(workbook) {
        try {
            // Find the best sheet to analyze (Prefer NG明細 or 整理)
            const sheetNames = workbook.SheetNames;
            let targetSheetName = sheetNames.find(s => s.includes('NG明細'));
            if (!targetSheetName) {
                targetSheetName = sheetNames.find(s => s.includes('整理'));
            }
            if (!targetSheetName) {
                targetSheetName = sheetNames[0]; // fallback
            }

            const sheet = workbook.Sheets[targetSheetName];
            // Convert to JSON array, use first row as header
            const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: null });
            
            // 必須先顯示畫面，確保 ECharts 能抓到正確的寬高再進行渲染
            showView('dashboard-view');
            btnReupload.classList.remove('hidden');
            btnExportDashboard.classList.remove('hidden');
            
            // 解析資料並畫圖
            analyzeData(jsonData);
            
            // Resize charts if window resizes
            window.addEventListener('resize', () => {
                if(trendChart) trendChart.resize();
                if(reasonChart) reasonChart.resize();
                if(factorChart) factorChart.resize();
                if(factorCountChart) factorCountChart.resize();
                if(productChart) productChart.resize();
            });

        } catch (error) {
            console.error(error);
            alert('資料解析錯誤，請確認 Excel 報表格式是否正確。');
            showView('upload-view');
        }
    }

    // Helper: Excel serial date to JS Date formatting
    function formatExcelDate(excelDate) {
        // If it's a number (Excel date format)
        if (typeof excelDate === 'number') {
            const date = new Date((excelDate - (25567 + 2)) * 86400 * 1000);
            return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
        }
        // If it's a string containing T or timestamp
        if (typeof excelDate === 'string') {
            try {
                const date = new Date(excelDate);
                if (!isNaN(date.getTime())) {
                    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
                }
            } catch (e) {}
        }
        return excelDate;
    }

    function analyzeData(data) {
        let totalProd = 0;
        let totalNG = 0;
        
        let dailyStats = {}; // { 'YYYY-MM-DD': { prod: 0, ng: 0 } }
        let reasonsStats = {}; // { 'reason': count }
        let factorStats = {}; // { 'factor': count }
        let productFactorStats = {}; // { 'Product': { total: 0, factors: {} } }
        let allFactorsSet = new Set();
        let validRecordsForTable = []; // for data table
        let pivotStats = {}; // { 'Product': { 'Reason': { 'Factor': count } } }

        data.forEach(row => {
            // Try to identify columns flexibly
            let dateVal = row['日期'] || row['每日生產'];
            let productVal = row['產品'] || row['Product'];
            let prodNum = parseFloat(row['生產數']) || 0;
            let ngNum = parseFloat(row['不良包數']) || 0;
            let reasonVal = row['NG原因'] || row['原因'];
            let factorVal = row['影響因素'];
            let intervalVal = row['NG區間'] || row['區間'] || row['區間 / 批次'] || '-';

            // For '整理' sheet which might not have raw absolute numbers but has rates
            if (!row['生產數'] && row['不良率'] !== undefined) {
                 // Assume each row is a record of 1 issue
                 ngNum = 1;
            }

            // Accumulate KPI
            totalProd += prodNum;
            totalNG += ngNum;

            // Daily Stats
            if (dateVal) {
                const dateStr = formatExcelDate(dateVal);
                if (dateStr && String(dateStr) !== "undefined" && dateStr !== null) {
                    if (!dailyStats[dateStr]) dailyStats[dateStr] = { prod: 0, ng: 0 };
                    dailyStats[dateStr].prod += prodNum;
                    dailyStats[dateStr].ng += ngNum;
                }
            }

            // Reason Stats
            if (reasonVal) {
                reasonsStats[reasonVal] = (reasonsStats[reasonVal] || 0) + (ngNum > 0 ? ngNum : 1);
            }

            // Factor Stats
            if (factorVal) {
                factorStats[factorVal] = (factorStats[factorVal] || 0) + (ngNum > 0 ? ngNum : 1);
            }

            // Product & Factor Stacked Stats
            if (productVal) {
                if (!productFactorStats[productVal]) {
                    productFactorStats[productVal] = { total: 0, factors: {} };
                }
                const currentFactor = factorVal || '未標明/其他';
                allFactorsSet.add(currentFactor);
                
                const valToAdd = (ngNum > 0 ? ngNum : 1);
                productFactorStats[productVal].total += valToAdd;
                productFactorStats[productVal].factors[currentFactor] = (productFactorStats[productVal].factors[currentFactor] || 0) + valToAdd;
            }

            // Pivot Stats
            if (ngNum > 0 && productVal) {
                const r = reasonVal || '-';
                const f = factorVal || '-';
                if (!pivotStats[productVal]) pivotStats[productVal] = {};
                if (!pivotStats[productVal][r]) pivotStats[productVal][r] = {};
                pivotStats[productVal][r][f] = (pivotStats[productVal][r][f] || 0) + ngNum;
            }

            // Add to Data Table if it's an NG event
            if (ngNum > 0 || row['不良率'] !== undefined) {
                validRecordsForTable.push({
                    date: dateVal ? formatExcelDate(dateVal) : '-',
                    product: productVal || '-',
                    interval: intervalVal,
                    ngNum: ngNum,
                    prodNum: prodNum,
                    reason: reasonVal || '-',
                    factor: factorVal || '-'
                });
            }
        });

        // 1. Update KPIs
        const totalProdFormatted = totalProd > 0 ? totalProd.toLocaleString() : "無生產總數";
        const totalNGFormatted = totalNG > 0 ? totalNG.toLocaleString() : Object.keys(productFactorStats).reduce((sum, key) => sum + productFactorStats[key].total, 0).toLocaleString(); // fallback to count
        
        // Calculate average rate
        let avgRateStr = "0.00%";
        if (totalProd > 0) {
            avgRateStr = ((totalNG / totalProd) * 100).toFixed(2) + "%";
        }

        document.getElementById('kpi-total-prod').innerText = totalProdFormatted;
        document.getElementById('kpi-total-ng').innerText = totalNGFormatted;
        document.getElementById('kpi-avg-rate').innerText = avgRateStr;

        // Prepare data for charts
        // --- Trend ---
        const sortedDates = Object.keys(dailyStats).sort();
        const trendDates = [];
        const trendProd = [];
        const trendRate = [];
        
        sortedDates.forEach(date => {
            trendDates.push(date);
            const p = dailyStats[date].prod;
            const n = dailyStats[date].ng;
            trendProd.push(p);
            trendRate.push(p > 0 ? ((n / p) * 100).toFixed(2) : 0);
        });

        // --- Pie (Reasons) ---
        const reasonData = Object.keys(reasonsStats).map(k => ({ value: reasonsStats[k], name: String(k) }))
            .sort((a,b) => b.value - a.value).slice(0, 7); // Top 7

        // --- Bar (Factors) ---
        const factorDataEntries = Object.keys(factorStats).map(k => ({ name: k, value: factorStats[k] }))
            .sort((a,b) => a.value - b.value); // Ascending
        const factorNames = factorDataEntries.map(e => e.name);
        const factorValues = factorDataEntries.map(e => e.value);

        // --- Bar (Products Stacked by Factor) ---
        const prodDataEntries = Object.keys(productFactorStats).map(k => ({ 
            name: k, 
            total: productFactorStats[k].total,
            factors: productFactorStats[k].factors
        })).sort((a,b) => a.total - b.total); // Ascending total
        const prodNames = prodDataEntries.map(e => e.name);
        const allFactorsArray = Array.from(allFactorsSet);

        // Store for redraws
        lastRenderData = {
            trendDates, trendProd, trendRate,
            reasonData, factorNames, factorValues,
            prodNames, allFactorsArray, prodDataEntries,
            productFactorStats
        };

        // Render Charts!
        renderTrendChart(trendDates, trendProd, trendRate);
        updateColorsAndRedraw(); // renders Pie, Factor, Product charts

        renderPivotTable(pivotStats);
        renderFactorCountTable(productFactorStats);
        renderDataTable(validRecordsForTable);
    }

    function updateColorsAndRedraw() {
        renderReasonPieChart(lastRenderData.reasonData);
        renderFactorBarChart(lastRenderData.factorNames, lastRenderData.factorValues);
        renderFactorCountChart(lastRenderData.productFactorStats);
        renderProductBarChart(lastRenderData.prodNames, lastRenderData.allFactorsArray, lastRenderData.prodDataEntries);
    }

    // ECharts generic dark theme utility styles
    const chartTextColor = '#e2e8f0';
    const chartLineColor = 'rgba(255, 255, 255, 0.1)';
    const splitLineColor = 'rgba(255, 255, 255, 0.05)';
    const tooltipBg = 'rgba(15, 23, 42, 0.9)';

    function renderTrendChart(dates, prods, rates) {
        if(trendChart) trendChart.dispose();
        trendChart = echarts.init(document.getElementById('trend-chart'));

        const option = {
            toolbox: {
                feature: {
                    saveAsImage: { name: '每日生產不良率趨勢', title: '儲存圖片', backgroundColor: '#131A2A' }
                }
            },
            tooltip: {
                trigger: 'axis',
                backgroundColor: tooltipBg,
                borderColor: '#334155',
                textStyle: { color: chartTextColor },
                axisPointer: { type: 'cross' }
            },
            legend: {
                data: ['生產數量', '不良率 (%)'],
                textStyle: { color: chartTextColor }
            },
            grid: { left: '3%', right: '4%', bottom: '5%', containLabel: true },
            xAxis: [
                {
                    type: 'category',
                    data: dates,
                    axisPointer: { type: 'shadow' },
                    axisLabel: { color: chartTextColor },
                    axisLine: { lineStyle: { color: chartLineColor } }
                }
            ],
            yAxis: [
                {
                    type: 'value',
                    name: '數量',
                    min: 0,
                    axisLabel: { color: chartTextColor },
                    splitLine: { lineStyle: { color: splitLineColor } },
                    nameTextStyle: { color: chartTextColor }
                },
                {
                    type: 'value',
                    name: '不良率 (%)',
                    min: 0,
                    axisLabel: { formatter: '{value} %', color: '#F59E0B' },
                    splitLine: { show: false },
                    nameTextStyle: { color: '#F59E0B' }
                }
            ],
            series: [
                {
                    name: '生產數量',
                    type: 'bar',
                    itemStyle: {
                        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                            { offset: 0, color: '#3B82F6' },
                            { offset: 1, color: '#1E3A8A' }
                        ]),
                        borderRadius: [4, 4, 0, 0]
                    },
                    data: prods
                },
                {
                    name: '不良率 (%)',
                    type: 'line',
                    yAxisIndex: 1,
                    itemStyle: { color: '#F59E0B' },
                    lineStyle: { width: 3, shadowColor: 'rgba(245,158,11,0.5)', shadowBlur: 10 },
                    symbol: 'emptyCircle',
                    symbolSize: 8,
                    data: rates
                }
            ]
        };
        trendChart.setOption(option);
    }

    function renderReasonPieChart(data) {
        if(reasonChart) reasonChart.dispose();
        reasonChart = echarts.init(document.getElementById('reason-chart'));

        const option = {
            toolbox: {
                feature: {
                    saveAsImage: { name: '主要NG原因佔比', title: '儲存圖片', backgroundColor: '#131A2A' }
                }
            },
            tooltip: {
                trigger: 'item',
                backgroundColor: tooltipBg,
                borderColor: '#334155',
                textStyle: { color: chartTextColor },
                formatter: '{b}: {c} ({d}%)' // b是name c是value d是百分比
            },
            legend: {
                orient: 'vertical',
                right: '5%',
                top: 'middle',
                textStyle: { color: chartTextColor }
            },
            series: [
                {
                    name: 'NG原因',
                    type: 'pie',
                    radius: ['40%', '70%'],
                    center: ['40%', '50%'],
                    avoidLabelOverlap: false,
                    itemStyle: {
                        borderRadius: 10,
                        borderColor: '#0f172a', /* border color matches background to create gap */
                        borderWidth: 2
                    },
                    label: { show: false },
                    emphasis: {
                        label: {
                            show: true,
                            fontSize: '14',
                            fontWeight: 'bold',
                            color: '#fff'
                        }
                    },
                    labelLine: { show: false },
                    data: data.map(item => ({
                        name: item.name,
                        value: item.value,
                        itemStyle: { color: getReasonColor(item.name) }
                    }))
                }
            ]
        };
        reasonChart.setOption(option);
    }

    function renderFactorBarChart(names, values) {
        if(factorChart) factorChart.dispose();
        factorChart = echarts.init(document.getElementById('factor-chart'));

        const option = {
            toolbox: {
                feature: {
                    saveAsImage: { name: '影響因素分析', title: '儲存圖片', backgroundColor: '#131A2A' }
                }
            },
            tooltip: {
                trigger: 'axis',
                backgroundColor: tooltipBg,
                borderColor: '#334155',
                textStyle: { color: chartTextColor },
                axisPointer: { type: 'shadow' }
            },
            grid: { left: '3%', right: '10%', bottom: '3%', top: '10%', containLabel: true },
            xAxis: {
                type: 'value',
                axisLabel: { color: chartTextColor },
                splitLine: { lineStyle: { color: splitLineColor } },
            },
            yAxis: {
                type: 'category',
                data: names,
                axisLabel: { color: chartTextColor, fontWeight: 'bold' },
                axisLine: { lineStyle: { color: chartLineColor } }
            },
            series: [
                {
                    name: '發生次數',
                    type: 'bar',
                    data: values.map((val, i) => ({
                        value: val,
                        itemStyle: { color: getFactorColor(names[i]) }
                    })),
                    label: { show: true, position: 'right', color: chartTextColor },
                    itemStyle: {
                        borderRadius: [0, 4, 4, 0]
                    }
                }
            ]
        };
        factorChart.setOption(option);
    }

    function renderProductBarChart(names, factors, dataEntries) {
        const chartContainer = document.getElementById('product-chart');
        // 加大圖表高度以容納多列圖例
        chartContainer.style.minHeight = '550px';
        
        if(productChart) productChart.dispose();
        productChart = echarts.init(chartContainer);

        const seriesData = factors.map(factor => {
            return {
                name: factor,
                type: 'bar',
                stack: 'total',
                itemStyle: { color: getFactorColor(factor) },
                emphasis: { focus: 'series' },
                // Only show label if the value is big enough or if it's > 0
                label: {
                    show: true,
                    formatter: (params) => params.value > 0 ? params.value : '',
                    color: '#fff',
                    fontWeight: 600
                },
                data: dataEntries.map(entry => entry.factors[factor] || 0)
            };
        });

        const option = {
            toolbox: {
                feature: {
                    saveAsImage: { name: '各產品NG次數排名(堆疊明細)', title: '儲存圖片', backgroundColor: '#131A2A' }
                }
            },
            tooltip: {
                trigger: 'axis',
                backgroundColor: tooltipBg,
                borderColor: '#334155',
                textStyle: { color: chartTextColor },
                axisPointer: { type: 'shadow' }
            },
            legend: {
                data: factors,
                top: 0,
                textStyle: { color: chartTextColor }
            },
            grid: { left: '3%', right: '10%', bottom: '3%', top: 120, containLabel: true },
            xAxis: {
                type: 'value',
                axisLabel: { color: chartTextColor },
                splitLine: { lineStyle: { color: splitLineColor } },
                axisLine: { show: false }
            },
            yAxis: {
                type: 'category',
                data: names,
                axisLabel: { color: chartTextColor, fontWeight: 'bold' },
                axisLine: { lineStyle: { color: chartLineColor } }
            },
            series: seriesData
        };
        productChart.setOption(option);
    }

    function renderFactorCountChart(productFactorStats) {
        if(factorCountChart) factorCountChart.dispose();
        factorCountChart = echarts.init(document.getElementById('factor-count-chart'));

        const entries = Object.keys(productFactorStats).map(p => ({
            name: p,
            count: Object.keys(productFactorStats[p].factors).length
        })).sort((a, b) => a.count - b.count); // Ascending for horizontal bar

        const names = entries.map(e => e.name);
        const values = entries.map(e => e.count);

        const option = {
            toolbox: {
                feature: {
                    saveAsImage: { name: '產品受影響因素種類數', title: '儲存圖片', backgroundColor: '#131A2A' }
                }
            },
            tooltip: {
                trigger: 'axis',
                backgroundColor: tooltipBg,
                borderColor: '#334155',
                textStyle: { color: chartTextColor },
                axisPointer: { type: 'shadow' }
            },
            grid: { left: '3%', right: '10%', bottom: '3%', top: '5%', containLabel: true },
            xAxis: {
                type: 'value',
                axisLabel: { color: chartTextColor },
                splitLine: { lineStyle: { color: splitLineColor } },
                axisLine: { show: false }
            },
            yAxis: {
                type: 'category',
                data: names,
                axisLabel: { color: chartTextColor, fontWeight: 'bold' },
                axisLine: { lineStyle: { color: chartLineColor } }
            },
            series: [{
                name: '種類數',
                type: 'bar',
                data: values,
                label: { show: true, position: 'right', color: chartTextColor },
                itemStyle: {
                    color: new echarts.graphic.LinearGradient(1, 0, 0, 0, [
                        { offset: 0, color: '#3B82F6' },
                        { offset: 1, color: '#2563EB' }
                    ]),
                    borderRadius: [0, 4, 4, 0]
                }
            }]
        };
        factorCountChart.setOption(option);
    }

    function renderPivotTable(pivotStats) {
        const tbody = document.getElementById('pivot-table-body');
        tbody.innerHTML = '';
        
        Object.keys(pivotStats).forEach(product => {
            // Sort reasons by count descending or just alphabetically
            Object.keys(pivotStats[product]).forEach(reason => {
                Object.keys(pivotStats[product][reason]).forEach(factor => {
                    const count = pivotStats[product][reason][factor];
                    
                    let tr = document.createElement('tr');
                    
                    // Create Reason color picker
                    let reasonColorInput = '<span class="text-muted">無</span>';
                    if (reason !== '-') {
                        reasonColorInput = `<input type="color" value="${getReasonColor(reason)}" data-type="reason" data-name="${reason}">`;
                    }
                    
                    // Create Factor color picker
                    let factorColorInput = '<span class="text-muted">無</span>';
                    if (factor !== '-' && factor !== '未標明/其他') {
                        factorColorInput = `<input type="color" value="${getFactorColor(factor)}" data-type="factor" data-name="${factor}">`;
                    }

                    tr.innerHTML = `
                        <td><strong>${product}</strong></td>
                        <td>${reason}</td>
                        <td>${factor}</td>
                        <td class="text-warning"><strong>${count}</strong></td>
                        <td>${reasonColorInput}</td>
                        <td>${factorColorInput}</td>
                    `;
                    tbody.appendChild(tr);
                });
            });
        });
        
        // Attach event listeners to color pickers
        document.querySelectorAll('#pivot-table-body input[type="color"]').forEach(input => {
            input.addEventListener('input', (e) => {
                const type = e.target.getAttribute('data-type');
                const name = e.target.getAttribute('data-name');
                const newColor = e.target.value;
                
                if (type === 'reason') {
                    window.chartColors.reasons[name] = newColor;
                } else if (type === 'factor') {
                    window.chartColors.factors[name] = newColor;
                }
                
                // Redraw ECharts instantly
                updateColorsAndRedraw();
                
                // Sync other same color pickers in the table
                document.querySelectorAll(`#pivot-table-body input[data-name="${name}"][data-type="${type}"]`).forEach(inp => {
                    if (inp !== e.target) inp.value = newColor;
                });
            });
        });
    }

    function renderDataTable(records) {
        const tbody = document.getElementById('detail-table-body');
        tbody.innerHTML = ''; // clear

        // Limit to latest or reverse so newest is on top if sorted by date
        // Create rows
        records.forEach(rec => {
            let tr = document.createElement('tr');
            
            // Format colors based on NG
            let ngClass = rec.ngNum > 50 ? 'text-danger' : 
                         (rec.ngNum > 10 ? 'text-warning' : '');

            tr.innerHTML = `
                <td>${rec.date}</td>
                <td><strong>${rec.product}</strong></td>
                <td>${rec.interval}</td>
                <td class="${ngClass}"><strong>${rec.ngNum}</strong></td>
                <td>${rec.prodNum}</td>
                <td>${rec.reason}</td>
                <td>${rec.factor}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    function renderFactorCountTable(productFactorStats) {
        const tbody = document.getElementById('factor-count-table-body');
        if (!tbody) return;
        tbody.innerHTML = '';
        
        const entries = Object.keys(productFactorStats).map(p => {
            // "未標明/其他" excluded? If we want exactly what the chart shows, count the property length.
            // Often there might be a factor called "-" or "未標明/其他". Let's count all keys.
            return {
                name: p,
                count: Object.keys(productFactorStats[p].factors).length
            };
        }).sort((a, b) => b.count - a.count);

        entries.forEach(item => {
            let tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="text-align: center;"><strong>${item.name}</strong></td>
                <td class="text-warning" style="font-size: 1.2rem; text-align: center;"><strong>${item.count}</strong></td>
            `;
            tbody.appendChild(tr);
        });
    }

    // --- Export as Image functionality using html2canvas ---
    
    // 1. Export Dashboard
    if(btnExportDashboard) {
        btnExportDashboard.addEventListener('click', () => {
            const dashboardElem = document.getElementById('dashboard-view');
            
            alert('正在為您擷取儀錶板，這可能需要幾秒鐘...');
            btnExportDashboard.disabled = true;
            
            html2canvas(dashboardElem, {
                backgroundColor: '#0B0F19', // Match body bg
                scale: 2 // High Resolution
            }).then(canvas => {
                let link = document.createElement('a');
                link.download = 'NG_Dashboard_Report.png';
                link.href = canvas.toDataURL('image/png');
                link.click();
                btnExportDashboard.disabled = false;
            }).catch(err => {
                console.error(err);
                alert('擷取圖片失敗！');
                btnExportDashboard.disabled = false;
            });
        });
    }

    // 2. Export Data Table Only
    if(btnExportTable) {
        btnExportTable.addEventListener('click', () => {
            const tableElem = document.getElementById('capture-table-area');
            
            btnExportTable.disabled = true;
            btnExportTable.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> 處理中...';

            html2canvas(tableElem, {
                backgroundColor: '#131A2A', // Match card bg
                scale: 2
            }).then(canvas => {
                let link = document.createElement('a');
                link.download = 'NG_Detailed_Table.png';
                link.href = canvas.toDataURL('image/png');
                link.click();
                btnExportTable.innerHTML = '<i class="fa-solid fa-image"></i> 將表格存為圖片';
                btnExportTable.disabled = false;
            }).catch(err => {
                console.error(err);
                alert('擷取圖片失敗！');
                btnExportTable.innerHTML = '<i class="fa-solid fa-image"></i> 將表格存為圖片';
                btnExportTable.disabled = false;
            });
        });
    }

    // 3. Export Pivot Table
    if(btnExportPivot) {
        btnExportPivot.addEventListener('click', () => {
            const tableElem = document.getElementById('capture-pivot-area');
            btnExportPivot.disabled = true;
            btnExportPivot.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> 處理中...';

            html2canvas(tableElem, {
                backgroundColor: '#131A2A',
                scale: 2
            }).then(canvas => {
                let link = document.createElement('a');
                link.download = 'NG_Pivot_Colors_Table.png';
                link.href = canvas.toDataURL('image/png');
                link.click();
                btnExportPivot.innerHTML = '<i class="fa-solid fa-image"></i> 將選色表存為圖片';
                btnExportPivot.disabled = false;
            }).catch(err => {
                console.error(err);
                alert('擷取圖片失敗！');
                btnExportPivot.innerHTML = '<i class="fa-solid fa-image"></i> 將選色表存為圖片';
                btnExportPivot.disabled = false;
            });
        });
    }

    // 4. Export Factor Count Table
    if(btnExportFactorCount) {
        btnExportFactorCount.addEventListener('click', () => {
            const tableElem = document.getElementById('capture-factor-count-area');
            btnExportFactorCount.disabled = true;
            btnExportFactorCount.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> 處理中...';

            html2canvas(tableElem, {
                backgroundColor: '#131A2A',
                scale: 2
            }).then(canvas => {
                let link = document.createElement('a');
                link.download = 'NG_Factor_Count_Table.png';
                link.href = canvas.toDataURL('image/png');
                link.click();
                btnExportFactorCount.innerHTML = '<i class="fa-solid fa-image"></i> 將特徵表存為圖片';
                btnExportFactorCount.disabled = false;
            }).catch(err => {
                console.error(err);
                alert('擷取圖片失敗！');
                btnExportFactorCount.innerHTML = '<i class="fa-solid fa-image"></i> 將特徵表存為圖片';
                btnExportFactorCount.disabled = false;
            });
        });
    }
});
