document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const uploadView = document.getElementById('upload-view');
    const loaderView = document.getElementById('loader-view');
    const dashboardView = document.getElementById('dashboard-view');
    const btnReupload = document.getElementById('btn-reupload');
    const btnExportDashboard = document.getElementById('btn-export');
    const btnExportTable = document.getElementById('btn-export-table');

    // Chart instances
    let trendChart, reasonChart, factorChart, productChart;

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
        let productStats = {}; // { 'Product': count }
        let validRecordsForTable = []; // for data table

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

            // Product Stats
            if (productVal) {
                productStats[productVal] = (productStats[productVal] || 0) + (ngNum > 0 ? ngNum : 1);
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
        const totalNGFormatted = totalNG > 0 ? totalNG.toLocaleString() : Object.keys(productStats).reduce((a,b)=>a+productStats[b], 0).toLocaleString(); // fallback to count
        
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

        // --- Bar (Products) ---
        const prodDataEntries = Object.keys(productStats).map(k => ({ name: k, value: productStats[k] }))
            .sort((a,b) => a.value - b.value); // Ascending for horizontal bar
        const prodNames = prodDataEntries.map(e => e.name);
        const prodValues = prodDataEntries.map(e => e.value);

        // Render Charts!
        renderTrendChart(trendDates, trendProd, trendRate);
        renderReasonPieChart(reasonData);
        renderFactorBarChart(factorNames, factorValues);
        renderProductBarChart(prodNames, prodValues);

        renderDataTable(validRecordsForTable);
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
                    saveAsImage: { name: '每日生產不良率趨勢', title: '儲存圖片' }
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
                    saveAsImage: { name: '主要NG原因佔比', title: '儲存圖片' }
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
                    data: data,
                    color: ['#EF4444', '#F97316', '#F59E0B', '#10B981', '#3B82F6', '#8B5CF6', '#EC4899']
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
                    saveAsImage: { name: '影響因素分析', title: '儲存圖片' }
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
                    data: values,
                    label: { show: true, position: 'right', color: chartTextColor },
                    itemStyle: {
                        color: new echarts.graphic.LinearGradient(1, 0, 0, 0, [
                            { offset: 0, color: '#8B5CF6' },
                            { offset: 1, color: '#5B21B6' }
                        ]),
                        borderRadius: [0, 4, 4, 0]
                    }
                }
            ]
        };
        factorChart.setOption(option);
    }

    function renderProductBarChart(names, values) {
        if(productChart) productChart.dispose();
        productChart = echarts.init(document.getElementById('product-chart'));

        const option = {
            toolbox: {
                feature: {
                    saveAsImage: { name: '各產品NG次數排名', title: '儲存圖片' }
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
            series: [
                {
                    name: '不良次數/包數',
                    type: 'bar',
                    data: values,
                    label: {
                        show: true,
                        position: 'right',
                        color: chartTextColor
                    },
                    itemStyle: {
                        color: new echarts.graphic.LinearGradient(1, 0, 0, 0, [
                            { offset: 0, color: '#10B981' },
                            { offset: 1, color: '#047857' }
                        ]),
                        borderRadius: [0, 4, 4, 0]
                    }
                }
            ]
        };
        productChart.setOption(option);
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
});
