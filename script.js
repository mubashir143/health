document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excel-upload');
    const dashboardResults = document.getElementById('dashboard-results');
    const emptyState = document.getElementById('empty-state');
    const loadingState = document.getElementById('loading-state');

    fileInput.addEventListener('change', handleFileUpload);

    // Tab Switching Logic
    const tabLinks = document.querySelectorAll('.nav-links li[data-tab]');
    const tabContents = document.querySelectorAll('.content-body > .tab-content');

    tabLinks.forEach(link => {
        link.addEventListener('click', () => {
            const targetTab = link.getAttribute('data-tab');

            // Update active link
            tabLinks.forEach(l => l.classList.remove('active'));
            link.classList.add('active');

            // Show target content
            tabContents.forEach(content => {
                if (content.id === `tab-${targetTab}`) {
                    content.style.display = 'block';
                } else {
                    content.style.display = 'none';
                }
            });
        });
    });

    function resetStates() {
        emptyState.style.display = 'block';
        dashboardResults.style.display = 'none';
        loadingState.style.display = 'none';
    }

    function handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        // Show loading state
        emptyState.style.display = 'none';
        dashboardResults.style.display = 'none';
        loadingState.style.display = 'block';

        const reader = new FileReader();
        reader.onload = function (event) {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                // Process only the first sheet for simplicity
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                // Convert sheet to JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "-" });

                if (jsonData.length === 0) {
                    alert('The uploaded file appears to be empty.');
                    resetStates();
                    return;
                }

                const headers = Object.keys(jsonData[0]);
                if (headers.length < 2) {
                    alert('The uploaded file must have at least two columns (Role and House Count).');
                    resetStates();
                    return;
                }

                processHealthData(jsonData);
                populateRawDataTable(jsonData);
            } catch (error) {
                console.error('Detailed Excel Error:', error);
                alert('Analysis Error: ' + error.message);
                resetStates();
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function levenshteinDistance(a, b) {
        const matrix = [];
        for (let i = 0; i <= b.length; i++) {
            matrix[i] = [i];
        }
        for (let j = 0; j <= a.length; j++) {
            matrix[0][j] = j;
        }
        for (let i = 1; i <= b.length; i++) {
            for (let j = 1; j <= a.length; j++) {
                if (b.charAt(i - 1) === a.charAt(j - 1)) {
                    matrix[i][j] = matrix[i - 1][j - 1];
                } else {
                    matrix[i][j] = Math.min(
                        matrix[i - 1][j - 1] + 1,
                        matrix[i][j - 1] + 1,
                        matrix[i - 1][j] + 1
                    );
                }
            }
        }
        return matrix[b.length][a.length];
    }

    function similarityRatio(str1, str2) {
        const maxLen = Math.max(str1.length, str2.length);
        if (maxLen === 0) return 1;
        const distance = levenshteinDistance(str1.toLowerCase(), str2.toLowerCase());
        return (maxLen - distance) / maxLen;
    }

    function standardizeJobTitles(data, roleCol, standardTitle = 'Community Health Inspector', threshold = 0.8) {
        data.forEach(row => {
            const title = String(row[roleCol] || '').trim();
            if (title && similarityRatio(title, standardTitle) >= threshold) {
                row[roleCol] = standardTitle;
            }
        });
    }

    const searchInput = document.querySelector('.search-container input');

    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        const rows = document.querySelectorAll('#raw-data-body tr');

        rows.forEach(row => {
            const text = row.textContent.toLowerCase();
            row.style.display = text.includes(query) ? '' : 'none';
        });
    });

    function populateRawDataTable(data) {
        const tableHead = document.querySelector('#raw-data-table thead tr');
        const tableBody = document.getElementById('raw-data-body');
        const countBadge = document.getElementById('data-count-badge');

        if (!data || data.length === 0) return;

        // Populate Headers
        const headers = Object.keys(data[0]);
        tableHead.innerHTML = '<th>#</th>' + headers.map(h => `<th>${h}</th>`).join('');

        // Populate Body
        tableBody.innerHTML = data.map((row, index) => `
            <tr>
                <td>${index + 1}</td>
                ${headers.map(h => `<td>${row[h]}</td>`).join('')}
            </tr>
        `).join('');

        countBadge.textContent = `${data.length} Records`;
    }

    function processHealthData(data) {
        // Find which column is the role column
        const headers = Object.keys(data[0]);
        const lastCol = headers[headers.length - 1];

        // Find role column: Look for "role", "designation", "category", or common health officer terms
        let roleCol = headers.find(h => {
            const lowerVal = h.toLowerCase();
            return lowerVal.includes('role') || lowerVal.includes('designation') || lowerVal.includes('category') || lowerVal.includes('position');
        }) || headers.find(h => {
            const sampleValues = data.slice(0, 5).map(row => String(row[h]).toLowerCase());
            return sampleValues.some(v => v.includes('health') || v.includes('worker') || v.includes('officer'));
        }) || (headers.length > 1 ? headers[headers.length - 2] : headers[0]);

        console.log(`Analyzing: Houses in [${lastCol}], Roles in [${roleCol}]`);

        // Standardize job titles
        standardizeJobTitles(data, roleCol);

        // Perform calculations
        const analyze = (subset) => {
            const totalUsers = subset.length;
            let activeUsers = 0;
            let totalHouses = 0;
            let dist = { '0': 0, '1-5': 0, '6-10': 0, '11+': 0 };

            subset.forEach(row => {
                let houseCount = row[lastCol];
                // If the value is "-" or empty, treat as 0
                if (houseCount === "-" || houseCount === "" || houseCount === undefined || houseCount === null) {
                    houseCount = 0;
                } else {
                    houseCount = parseInt(houseCount) || 0;
                }

                totalHouses += houseCount;

                if (houseCount !== 0) activeUsers++;

                if (houseCount === 0) dist['0']++;
                else if (houseCount >= 1 && houseCount <= 5) dist['1-5']++;
                else if (houseCount >= 6 && houseCount <= 10) dist['6-10']++;
                else if (houseCount >= 11) dist['11+']++;
            });

            return { totalUsers, activeUsers, totalHouses, dist };
        };

        // Filter datasets
        const LHW_KEYWORDS = ['lady health worker', 'lhw'];
        const CHO_KEYWORDS = ['community health officer', 'cho', 'chi', 'community health inspector'];

        const lhwData = data.filter(row => {
            const val = String(row[roleCol] || '').toLowerCase();
            return LHW_KEYWORDS.some(k => val.includes(k));
        });

        const choData = data.filter(row => {
            const val = String(row[roleCol] || '').toLowerCase();
            return CHO_KEYWORDS.some(k => val.includes(k));
        });

        const overallResults = analyze(data);
        const lhwResults = analyze(lhwData);
        const choResults = analyze(choData);

        // Calculate Role Counts for Breakdown
        const roleCounts = {};
        data.forEach(row => {
            const r = row[roleCol] || 'Other';
            roleCounts[r] = (roleCounts[r] || 0) + 1;
        });

        updateDashboard(overallResults, lhwResults, choResults, roleCounts, data, lastCol, roleCol);
    }

    function createMetricCard(label, value, icon, className = '') {
        return `
            <div class="col-md-4">
                <div class="metric-card ${className}">
                    <div class="metric-label">${label}</div>
                    <div class="metric-value">${value.toLocaleString()}</div>
                    <div class="metric-icon">
                        <i class="fas ${icon}"></i>
                    </div>
                </div>
            </div>
        `;
    }

    function createDistributionHTML(dist) {
        return `
            <div class="dist-item">
                <div class="dist-label">0 Houses</div>
                <div class="dist-value">${dist['0']}</div>
            </div>
            <div class="dist-item">
                <div class="dist-label">1-5 Houses</div>
                <div class="dist-value">${dist['1-5']}</div>
            </div>
            <div class="dist-item">
                <div class="dist-label">6-10 Houses</div>
                <div class="dist-value">${dist['6-10']}</div>
            </div>
            <div class="dist-item">
                <div class="dist-label">11+ Houses</div>
                <div class="dist-value">${dist['11+']}</div>
            </div>
        `;
    }

    function updateDashboard(overall, lhw, cho, roleCounts, data, lastCol, roleCol) {
        // Overall Metrics
        const overallMetricsEl = document.getElementById('overall-metrics');
        overallMetricsEl.innerHTML =
            createMetricCard('Total Registered Users', overall.totalUsers, 'fa-users') +
            createMetricCard('Active Service Providers', overall.activeUsers, 'fa-user-check') +
            createMetricCard('Total Houses Covered', overall.totalHouses, 'fa-home');

        document.getElementById('overall-distribution').innerHTML = createDistributionHTML(overall.dist);

        // Role Breakdown
        const roleBreakdownEl = document.getElementById('role-breakdown');
        roleBreakdownEl.innerHTML = Object.entries(roleCounts).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([role, count]) => `
            <div class="role-item">
                <span class="role-name">${role}</span>
                <span class="role-count bg-primary text-white">${count}</span>
            </div>
        `).join('');

        // LHW Metrics
        const lhwMetricsEl = document.getElementById('lhw-metrics');
        lhwMetricsEl.innerHTML =
            createMetricCard('LHW Total Users', lhw.totalUsers, 'fa-female', 'lhw') +
            createMetricCard('LHW Active Users', lhw.activeUsers, 'fa-clipboard-check', 'lhw') +
            createMetricCard('LHW Houses Covered', lhw.totalHouses, 'fa-house-medical', 'lhw');

        document.getElementById('lhw-distribution').innerHTML = createDistributionHTML(lhw.dist);

        // CHO Metrics
        const choMetricsEl = document.getElementById('cho-metrics');
        choMetricsEl.innerHTML =
            createMetricCard('CHO Total Users', cho.totalUsers, 'fa-user-doctor', 'cho') +
            createMetricCard('CHO Active Users', cho.activeUsers, 'fa-stethoscope', 'cho') +
            createMetricCard('CHO Houses Covered', cho.totalHouses, 'fa-hospital', 'cho');

        document.getElementById('cho-distribution').innerHTML = createDistributionHTML(cho.dist);

        // Populate Summary List Tab
        populateSummaryList(overall, lhw, cho);

        // Populate UC Analysis Tab
        populateUCAnalysis(data, lastCol, roleCol);

        // Display dashboard
        loadingState.style.display = 'none';
        dashboardResults.style.display = 'block';
    }

    function populateUCAnalysis(data, houseCol, roleCol) {
        const body = document.getElementById('uc-summary-body');
        const headers = Object.keys(data[0]);

        // Find UC Column
        let ucCol = headers.find(h => {
            const low = h.toLowerCase();
            return low.includes('uc') || low.includes('union council') || low.includes('area') || low.includes('location');
        });

        if (!ucCol) {
            // Fallback: look for common UC names like "Chak" in data
            for (let h of headers) {
                const sampleValues = data.slice(0, 10).map(row => String(row[h]).toLowerCase());
                if (sampleValues.some(v => v.includes('chak') || v.includes('uc') || /^\d+$/.test(v))) {
                    ucCol = h;
                    break;
                }
            }
        }

        if (!ucCol) ucCol = headers[0]; // Final fallback

        // Grouping logic
        const groups = {};
        const LHW_KEYWORDS = ['lady health worker', 'lhw'];
        const CHO_KEYWORDS = ['community health officer', 'cho', 'chi', 'community health inspector'];

        data.forEach(row => {
            const uc = String(row[ucCol] || 'Unknown UC');
            if (!groups[uc]) {
                const emptyGroup = () => ({ users: 0, active: 0, houses: 0, dist: { '0': 0, '1-5': 0, '6-10': 0, '11+': 0 } });
                groups[uc] = {
                    total: emptyGroup(),
                    lhw: emptyGroup(),
                    cho: emptyGroup()
                };
            }

            const role = String(row[roleCol] || '').toLowerCase();
            const isLHW = LHW_KEYWORDS.some(k => role.includes(k));
            const isCHO = CHO_KEYWORDS.some(k => role.includes(k));

            let houseCount = row[houseCol];
            if (houseCount === "-" || houseCount === "" || houseCount === undefined || houseCount === null) {
                houseCount = 0;
            } else {
                houseCount = parseInt(houseCount) || 0;
            }

            const updateSubgroup = (g) => {
                g.users++;
                g.houses += houseCount;
                if (houseCount !== 0) g.active++;

                if (houseCount === 0) g.dist['0']++;
                else if (houseCount >= 1 && houseCount <= 5) g.dist['1-5']++;
                else if (houseCount >= 6 && houseCount <= 10) g.dist['6-10']++;
                else if (houseCount >= 11) g.dist['11+']++;
            };

            updateSubgroup(groups[uc].total);
            if (isLHW) updateSubgroup(groups[uc].lhw);
            if (isCHO) updateSubgroup(groups[uc].cho);
        });

        renderUCTable(groups);
    }

    function renderUCTable(groups) {
        const body = document.getElementById('uc-summary-body');
        body.innerHTML = Object.entries(groups).map(([uc, g]) => `
            <tr>
                <td class="fw-bold sticky-column">${uc}</td>
                <!-- Overall -->
                <td class="table-primary-light">${g.total.users}</td>
                <td class="table-primary-light">${g.total.active}</td>
                <td class="table-primary-light">${g.total.houses.toLocaleString()}</td>
                <td class="table-primary-light">${g.total.dist['0']}</td>
                <td class="table-primary-light">${g.total.dist['1-5']}</td>
                <td class="table-primary-light">${g.total.dist['6-10']}</td>
                <td class="table-primary-light">${g.total.dist['11+']}</td>
                <!-- LHW -->
                <td class="table-info-light">${g.lhw.users}</td>
                <td class="table-info-light">${g.lhw.active}</td>
                <td class="table-info-light">${g.lhw.houses.toLocaleString()}</td>
                <td class="table-info-light">${g.lhw.dist['0']}</td>
                <td class="table-info-light">${g.lhw.dist['1-5']}</td>
                <td class="table-info-light">${g.lhw.dist['6-10']}</td>
                <td class="table-info-light">${g.lhw.dist['11+']}</td>
                <!-- CHO -->
                <td class="table-success-light">${g.cho.users}</td>
                <td class="table-success-light">${g.cho.active}</td>
                <td class="table-success-light">${g.cho.houses.toLocaleString()}</td>
                <td class="table-success-light">${g.cho.dist['0']}</td>
                <td class="table-success-light">${g.cho.dist['1-5']}</td>
                <td class="table-success-light">${g.cho.dist['6-10']}</td>
                <td class="table-success-light">${g.cho.dist['11+']}</td>
            </tr>
        `).join('');
    }

    // UC Filter Listener
    document.getElementById('uc-filter')?.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        const rows = document.querySelectorAll('#uc-summary-body tr');
        rows.forEach(row => {
            const ucName = row.cells[0].textContent.toLowerCase();
            row.style.display = ucName.includes(query) ? '' : 'none';
        });
    });

    function populateSummaryList(overall, lhw, cho) {
        const body = document.getElementById('summary-list-body');
        const rows = [
            { name: 'Overall Users', data: overall, class: 'fw-bold' },
            { name: 'Lady Health Workers (LHW)', data: lhw, class: '' },
            { name: 'Community Health Inspector (CHI)', data: cho, class: '' }
        ];

        body.innerHTML = rows.map(row => `
            <tr class="${row.class}">
                <td>${row.name}</td>
                <td>${row.data.totalUsers.toLocaleString()}</td>
                <td>${row.data.activeUsers.toLocaleString()}</td>
                <td>${row.data.totalHouses.toLocaleString()}</td>
                <td>${row.data.dist['0'].toLocaleString()}</td>
                <td>${row.data.dist['1-5'].toLocaleString()}</td>
                <td>${row.data.dist['6-10'].toLocaleString()}</td>
                <td>${row.data.dist['11+'].toLocaleString()}</td>
            </tr>
        `).join('');
    }

    // --- GET CHI NUMBERS LOGIC ---
    let chiFinalData = [];

    const chiFile1 = document.getElementById('chi-file1');
    const chiFile2 = document.getElementById('chi-file2');
    const chiWrapper1 = document.getElementById('chi-wrapper1');
    const chiWrapper2 = document.getElementById('chi-wrapper2');
    const chiName1 = document.getElementById('chi-name1');
    const chiName2 = document.getElementById('chi-name2');
    const chiProcessBtn = document.getElementById('chi-process-btn');
    const chiExportBtn = document.getElementById('chi-export-btn');
    const chiStatus = document.getElementById('chi-status');

    function setupChiFileInput(input, wrapper, nameLabel) {
        wrapper.addEventListener('click', () => input.click());

        input.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                nameLabel.innerHTML = `<strong>${e.target.files[0].name}</strong>`;
                wrapper.classList.add('border-primary');
            }
        });

        wrapper.addEventListener('dragover', (e) => {
            e.preventDefault();
            wrapper.style.backgroundColor = 'rgba(13, 110, 253, 0.05)';
        });

        wrapper.addEventListener('dragleave', () => {
            wrapper.style.backgroundColor = '';
        });

        wrapper.addEventListener('drop', (e) => {
            e.preventDefault();
            wrapper.style.backgroundColor = '';
            if (e.dataTransfer.files.length > 0) {
                input.files = e.dataTransfer.files;
                input.dispatchEvent(new Event('change'));
            }
        });
    }

    setupChiFileInput(chiFile1, chiWrapper1, chiName1);
    setupChiFileInput(chiFile2, chiWrapper2, chiName2);

    function readExcelAsJSON(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = e.target.result;
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false });
                    resolve(jsonData);
                } catch (err) { reject(err); }
            };
            reader.onerror = (err) => reject(err);
            reader.readAsArrayBuffer(file);
        });
    }

    function findKey(obj, possibleNames) {
        if (!obj || typeof obj !== 'object') return null;
        const keys = Object.keys(obj);

        // 1. Exact or Case-insensitive
        for (let pName of possibleNames) {
            const searchName = pName.toLowerCase();
            for (let key of keys) {
                if (key.toLowerCase() === searchName) return key;
            }
        }

        // 2. Normalized
        for (let pName of possibleNames) {
            const normName = pName.toLowerCase().replace(/[^a-z0-9]/g, '');
            for (let key of keys) {
                const normKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
                if (normKey === normName) return key;
            }
        }

        // 3. Includes
        for (let pName of possibleNames) {
            const searchName = pName.toLowerCase();
            if (searchName.length < 3) continue;
            for (let key of keys) {
                if (key.toLowerCase().includes(searchName)) return key;
            }
        }
        return null;
    }

    chiProcessBtn?.addEventListener('click', async () => {
        const f1 = chiFile1.files[0];
        const f2 = chiFile2.files[0];

        if (!f1 || !f2) {
            chiStatus.innerHTML = '<span class="text-danger">Please upload both files.</span>';
            return;
        }

        chiProcessBtn.disabled = true;
        chiProcessBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Processing...';
        chiStatus.textContent = 'Reading files...';

        try {
            const nonReportingData = await readExcelAsJSON(f1);
            const userProfileData = await readExcelAsJSON(f2);

            if (nonReportingData.length === 0 || userProfileData.length === 0) {
                throw new Error("One or both files appear to be empty.");
            }

            // identify headers once
            const r1 = nonReportingData[0];
            const desigKey = findKey(r1, ['designation', 'role', 'title', 'post', 'category']);
            const cnicKey1 = findKey(r1, ['cnicofcadre', 'cnic', 'idm', 'cadrecnic', 'id', 'cadre']);

            const r2 = userProfileData[0];
            const cnicKey2 = findKey(r2, ['username', 'cnic', 'nationalid', 'idnumber', 'identity', 'id']);
            const phoneKey = findKey(r2, ['phone', 'contact', 'mobile', 'cell', 'number', 'tel', 'whatsapp']);

            if (!desigKey || !cnicKey1 || !cnicKey2) {
                throw new Error("Could not find required columns (Designation/CNIC) in the uploaded files. Please check headers.");
            }

            const normalizeCnic = c => String(c || '').split('.')[0].replace(/[^0-9]/g, '');

            nonReportingData.forEach(row => {
                const desigVal = String(row[desigKey] || '').toLowerCase();

                // Matches "community", "comuntiy", "comunity", "cummunity" + "inspect" or "chi"
                const hasComm = /comm?un/i.test(desigVal) || /cummun/i.test(desigVal);
                const hasInsp = /inspec/i.test(desigVal) || /insp/i.test(desigVal);
                const isInspector = (hasComm && hasInsp) || desigVal.includes('chi') || (hasComm && /health/i.test(desigVal));

                if (isInspector) {
                    const targetCnic = normalizeCnic(row[cnicKey1]);
                    if (targetCnic.length < 5) return;

                    const matchProfiles = userProfileData.filter(pRow => normalizeCnic(pRow[cnicKey2]) === targetCnic);

                    matchProfiles.forEach(match => {
                        const tehKey1 = findKey(row, ['tehsil', 'district', 'area']);
                        const tehKey2 = findKey(match, ['tehsil', 'district', 'area']);
                        const nmKey1 = findKey(row, ['name', 'fullname', 'user']);
                        const nmKey2 = findKey(match, ['name', 'fullname', 'user']);

                        chiFinalData.push({
                            "Tehsil": (row[tehKey1] || match[tehKey2] || 'N/A'),
                            "Name": (row[nmKey1] || match[nmKey2] || 'N/A'),
                            "CNIC": row[cnicKey1],
                            "Phone Number": (match[phoneKey] || 'N/A')
                        });
                    });
                }
            });

            // Update UI
            const resultsBody = document.getElementById('chi-results-body');
            const countBadge = document.getElementById('chi-count-badge');

            if (chiFinalData.length > 0) {
                resultsBody.innerHTML = chiFinalData.map(d => `
                    <tr>
                        <td class="px-4">${d.Tehsil}</td>
                        <td>${d.Name}</td>
                        <td><code class="text-dark">${d["CNIC"]}</code></td>
                        <td><span class="badge bg-light text-dark border"><i class="fas fa-phone me-1 text-success"></i> ${d["Phone Number"]}</span></td>
                    </tr>
                `).join('');
                countBadge.textContent = `${chiFinalData.length} Matches`;
                chiStatus.innerHTML = `<span class="text-success">Found ${chiFinalData.length} matches!</span>`;
                chiExportBtn.disabled = false;

                // Switch to results tab automatically (using Bootstrap Tab API)
                const resultsTabTrigger = document.getElementById('chi-nav-results');
                if (resultsTabTrigger) {
                    const tab = new bootstrap.Tab(resultsTabTrigger);
                    tab.show();
                }
            } else {
                resultsBody.innerHTML = '<tr><td colspan="4" class="text-center py-5">No matching records found.</td></tr>';
                countBadge.textContent = '0 Matches';
                chiStatus.innerHTML = '<span class="text-warning">No matches found between the two files.</span>';
                chiExportBtn.disabled = true;
            }

        } catch (err) {
            console.error(err);
            chiStatus.innerHTML = `<span class="text-danger">Error: ${err.message}</span>`;
        } finally {
            chiProcessBtn.disabled = false;
            chiProcessBtn.innerHTML = '<i class="fas fa-sync-alt me-2"></i> Process and Match Records';
        }
    });

    chiExportBtn?.addEventListener('click', () => {
        if (chiFinalData.length === 0) return;
        const newSheet = XLSX.utils.json_to_sheet(chiFinalData);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Matched CHI");
        XLSX.writeFile(newWorkbook, "Matched_CHI_Contacts.xlsx");
    });
});
