document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excel-upload');
    const dashboardResults = document.getElementById('dashboard-results');
    const emptyState = document.getElementById('empty-state');
    const loadingState = document.getElementById('loading-state');

    // Global Filter and Calculations State
    let currentRawData = null;
    let activeFilters = {
        calcCol: '',
        district: 'ALL',
        tehsil: 'ALL',
        uc: 'ALL',
        selectedDesignations: []
    };

    // Filter UI Elements
    const globalFilterCard = document.getElementById('global-filter-card');
    const filterCalcCol = document.getElementById('filter-calc-col');
    const filterDistrict = document.getElementById('filter-district');
    const filterTehsil = document.getElementById('filter-tehsil');
    const filterUc = document.getElementById('filter-uc');
    const designationCheckboxContainer = document.getElementById('designation-checkbox-container');
    const btnDesigSelectAll = document.getElementById('btn-desig-select-all');
    const btnDesigClearAll = document.getElementById('btn-desig-clear-all');
    const btnResetFilters = document.getElementById('btn-reset-filters');

    fileInput.addEventListener('change', handleFileUpload);

    // Tab Switching Logic
    const tabLinks = document.querySelectorAll('.nav-links li[data-tab]');
    const tabContents = document.querySelectorAll('.content-body > .tab-content');

    // Mobile Menu Toggle
    const mobileMenuToggle = document.querySelector('.mobile-menu-toggle');
    const sidebar = document.querySelector('.sidebar');
    const sidebarOverlay = document.getElementById('sidebar-overlay');

    if (mobileMenuToggle) {
        mobileMenuToggle.addEventListener('click', () => {
            sidebar.classList.toggle('show');
            sidebarOverlay.classList.toggle('show');
        });

        sidebarOverlay.addEventListener('click', () => {
            sidebar.classList.remove('show');
            sidebarOverlay.classList.remove('show');
        });

        // Close sidebar on tab click (mobile)
        tabLinks.forEach(link => {
            link.addEventListener('click', () => {
                if (window.innerWidth <= 991) {
                    sidebar.classList.remove('show');
                    sidebarOverlay.classList.remove('show');
                }
            });
        });
    }

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

            // Show/hide filter card if data is uploaded
            const filterableTabs = ['dashboard', 'summary', 'district-summary', 'uc-wise', 'datalist'];
            if (globalFilterCard && currentRawData) {
                if (filterableTabs.includes(targetTab)) {
                    globalFilterCard.style.display = 'block';
                } else {
                    globalFilterCard.style.display = 'none';
                }
            }
        });
    });

    // Reset settings
    function resetStates() {
        emptyState.style.display = 'block';
        dashboardResults.style.display = 'none';
        loadingState.style.display = 'none';
        if (globalFilterCard) globalFilterCard.style.display = 'none';
        currentRawData = null;
    }

    // Dynamic Filter Event Handlers
    filterCalcCol?.addEventListener('change', applyFiltersAndCalculate);
    filterDistrict?.addEventListener('change', () => {
        rebuildTehsilDropdown();
        rebuildUcDropdown();
        applyFiltersAndCalculate();
    });
    filterTehsil?.addEventListener('change', () => {
        rebuildUcDropdown();
        applyFiltersAndCalculate();
    });
    filterUc?.addEventListener('change', applyFiltersAndCalculate);

    btnDesigSelectAll?.addEventListener('click', () => {
        const checkboxes = designationCheckboxContainer?.querySelectorAll('input[type="checkbox"]');
        checkboxes?.forEach(cb => cb.checked = true);
        applyFiltersAndCalculate();
    });

    btnDesigClearAll?.addEventListener('click', () => {
        const checkboxes = designationCheckboxContainer?.querySelectorAll('input[type="checkbox"]');
        checkboxes?.forEach(cb => cb.checked = false);
        applyFiltersAndCalculate();
    });

    btnResetFilters?.addEventListener('click', () => {
        if (!currentRawData) return;
        initializeFilters(currentRawData);
        applyFiltersAndCalculate();
    });

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

                currentRawData = jsonData;
                initializeFilters(jsonData);
                if (globalFilterCard) globalFilterCard.style.display = 'block';
                applyFiltersAndCalculate();
            } catch (error) {
                console.error('Detailed Excel Error:', error);
                alert('Analysis Error: ' + error.message);
                resetStates();
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // Dynamic Filter & Location Setup Helpers
    function initializeFilters(data) {
        if (!data || data.length === 0) return;
        const headers = Object.keys(data[0]);

        // Find District, Tehsil, and UC Columns
        let districtCol = getLocColumn('districtCol', headers);
        let tehsilCol = getLocColumn('tehsilCol', headers);
        let ucCol = getLocColumn('ucCol', headers);

        // Find Designation Column
        let roleCol = headers.find(h => {
            const lowerVal = h.toLowerCase();
            return lowerVal.includes('role') || lowerVal.includes('designation') || lowerVal.includes('category') || lowerVal.includes('position');
        }) || headers.find(h => {
            const sampleValues = data.slice(0, 5).map(row => String(row[h]).toLowerCase());
            return sampleValues.some(v => v.includes('health') || v.includes('worker') || v.includes('officer'));
        }) || (headers.length > 1 ? headers[headers.length - 2] : headers[0]);

        // 1. Populate Calculation Column (numeric columns from index 6 onwards)
        filterCalcCol.innerHTML = '';
        const defaultIndex = 6;
        headers.forEach((h, idx) => {
            if (idx >= 6) {
                const opt = document.createElement('option');
                opt.value = h;
                opt.textContent = h;
                if (idx === defaultIndex) {
                    opt.selected = true;
                    activeFilters.calcCol = h;
                }
                filterCalcCol.appendChild(opt);
            }
        });
        if (filterCalcCol.options.length === 0) {
            const lastHeader = headers[headers.length - 1];
            const opt = document.createElement('option');
            opt.value = lastHeader;
            opt.textContent = lastHeader;
            opt.selected = true;
            activeFilters.calcCol = lastHeader;
            filterCalcCol.appendChild(opt);
        }

        // Reset filter values
        activeFilters.district = 'ALL';
        activeFilters.tehsil = 'ALL';
        activeFilters.uc = 'ALL';

        // 2. Populate Designation Checklist
        const uniqueDesignations = new Set();
        data.forEach(row => {
            const r = String(row[roleCol] || 'Other').trim();
            if (r) uniqueDesignations.add(r);
        });

        designationCheckboxContainer.innerHTML = '';
        activeFilters.selectedDesignations = [...uniqueDesignations].sort();
        activeFilters.selectedDesignations.forEach((desig, idx) => {
            const id = `desig-cb-${idx}`;
            const div = document.createElement('div');
            div.className = 'form-check form-check-inline';
            div.innerHTML = `
                <input class="form-check-input designation-filter-cb" type="checkbox" id="${id}" value="${desig}" checked>
                <label class="form-check-label small text-dark" for="${id}">${desig}</label>
            `;
            div.querySelector('input').addEventListener('change', applyFiltersAndCalculate);
            designationCheckboxContainer.appendChild(div);
        });

        // Populate location options
        updateLocationDropdowns(data, districtCol, tehsilCol, ucCol);
    }

    function updateLocationDropdowns(data, districtCol, tehsilCol, ucCol) {
        const districts = new Set();
        const tehsilsMap = new Map();
        const ucsMap = new Map();

        data.forEach(row => {
            const dist = String(row[districtCol] || 'N/A').trim();
            const teh = String(row[tehsilCol] || 'N/A').trim();
            const uc = String(row[ucCol] || 'N/A').trim();

            districts.add(dist);
            
            if (!tehsilsMap.has(dist)) tehsilsMap.set(dist, new Set());
            tehsilsMap.get(dist).add(teh);

            if (!ucsMap.has(teh)) ucsMap.set(teh, new Set());
            ucsMap.get(teh).add(uc);
        });

        filterDistrict.innerHTML = '<option value="ALL">All Districts</option>';
        [...districts].sort().forEach(d => {
            const opt = document.createElement('option');
            opt.value = d;
            opt.textContent = d;
            filterDistrict.appendChild(opt);
        });

        window.locationData = {
            districtCol,
            tehsilCol,
            ucCol,
            tehsilsMap,
            ucsMap,
            districts: [...districts].sort()
        };

        rebuildTehsilDropdown();
        rebuildUcDropdown();
    }

    function rebuildTehsilDropdown() {
        const selectedDist = filterDistrict.value;
        filterTehsil.innerHTML = '<option value="ALL">All Tehsils</option>';

        if (selectedDist === 'ALL') {
            const allTehsils = new Set();
            window.locationData.tehsilsMap.forEach(set => set.forEach(t => allTehsils.add(t)));
            [...allTehsils].sort().forEach(t => {
                const opt = document.createElement('option');
                opt.value = t;
                opt.textContent = t;
                filterTehsil.appendChild(opt);
            });
        } else {
            const tehsils = window.locationData.tehsilsMap.get(selectedDist) || new Set();
            [...tehsils].sort().forEach(t => {
                const opt = document.createElement('option');
                opt.value = t;
                opt.textContent = t;
                filterTehsil.appendChild(opt);
            });
        }
    }

    function rebuildUcDropdown() {
        const selectedTeh = filterTehsil.value;
        const selectedDist = filterDistrict.value;
        filterUc.innerHTML = '<option value="ALL">All UCs</option>';

        if (selectedTeh === 'ALL') {
            const activeTehsils = new Set();
            if (selectedDist === 'ALL') {
                window.locationData.ucsMap.forEach((set, teh) => activeTehsils.add(teh));
            } else {
                const tehsils = window.locationData.tehsilsMap.get(selectedDist) || new Set();
                tehsils.forEach(t => activeTehsils.add(t));
            }

            const allUcs = new Set();
            activeTehsils.forEach(teh => {
                const ucs = window.locationData.ucsMap.get(teh) || new Set();
                ucs.forEach(u => allUcs.add(u));
            });

            [...allUcs].sort().forEach(u => {
                const opt = document.createElement('option');
                opt.value = u;
                opt.textContent = u;
                filterUc.appendChild(opt);
            });
        } else {
            const ucs = window.locationData.ucsMap.get(selectedTeh) || new Set();
            [...ucs].sort().forEach(u => {
                const opt = document.createElement('option');
                opt.value = u;
                opt.textContent = u;
                filterUc.appendChild(opt);
            });
        }
    }

    function applyFiltersAndCalculate() {
        if (!currentRawData || currentRawData.length === 0) return;

        activeFilters.calcCol = filterCalcCol.value;
        activeFilters.district = filterDistrict.value;
        activeFilters.tehsil = filterTehsil.value;
        activeFilters.uc = filterUc.value;

        // Gather checked designations
        const checkedBoxes = designationCheckboxContainer?.querySelectorAll('input[type="checkbox"]:checked');
        const selected = [];
        checkedBoxes?.forEach(cb => selected.push(cb.value));
        activeFilters.selectedDesignations = selected;

        const headers = Object.keys(currentRawData[0]);
        const loc = window.locationData || {
            districtCol: getLocColumn('districtCol', headers),
            tehsilCol: getLocColumn('tehsilCol', headers),
            ucCol: getLocColumn('ucCol', headers)
        };
        let filteredData = currentRawData;

        // Filter by Location
        if (activeFilters.district !== 'ALL') {
            filteredData = filteredData.filter(row => String(row[loc.districtCol] || '').trim() === activeFilters.district);
        }
        if (activeFilters.tehsil !== 'ALL') {
            filteredData = filteredData.filter(row => String(row[loc.tehsilCol] || '').trim() === activeFilters.tehsil);
        }
        if (activeFilters.uc !== 'ALL') {
            filteredData = filteredData.filter(row => String(row[loc.ucCol] || '').trim() === activeFilters.uc);
        }

        // Filter by Designation
        let roleCol = headers.find(h => {
            const lowerVal = h.toLowerCase();
            return lowerVal.includes('role') || lowerVal.includes('designation') || lowerVal.includes('category') || lowerVal.includes('position');
        }) || headers.find(h => {
            const sampleValues = currentRawData.slice(0, 5).map(row => String(row[h]).toLowerCase());
            return sampleValues.some(v => v.includes('health') || v.includes('worker') || v.includes('officer'));
        }) || (headers.length > 1 ? headers[headers.length - 2] : headers[0]);

        filteredData = filteredData.filter(row => {
            const r = String(row[roleCol] || 'Other').trim();
            return activeFilters.selectedDesignations.includes(r);
        });

        populateRawDataTable(filteredData);
        processHealthData(filteredData);
    }

    function findColumn(headers, keywords, ignoreKeywords, fallbackIndex) {
        // 1. Check for exact match (case-insensitive)
        for (const kw of keywords) {
            const found = headers.find(h => h.toLowerCase().trim() === kw);
            if (found) return found;
        }

        // 2. Check for matches containing the keywords but ignoring ID/code columns
        for (const kw of keywords) {
            const found = headers.find(h => {
                const low = h.toLowerCase();
                return low.includes(kw) && !ignoreKeywords.some(ignore => low.includes(ignore));
            });
            if (found) return found;
        }

        // 3. Fallback to any match containing the keywords
        for (const kw of keywords) {
            const found = headers.find(h => h.toLowerCase().includes(kw));
            if (found) return found;
        }

        // 4. Default fallback
        return headers[fallbackIndex] || headers[0];
    }

    function getLocColumn(colType, headers) {
        if (window.locationData && window.locationData[colType]) {
            return window.locationData[colType];
        }
        const ignore = ['code', 'id', 'no', 'number', 'num', 'key', 'fk', 'pk'];
        if (colType === 'districtCol') {
            return findColumn(headers, ['district', 'dist'], ignore, 0);
        } else if (colType === 'tehsilCol') {
            return findColumn(headers, ['tehsil', 'teh'], ignore, 1);
        } else if (colType === 'ucCol') {
            return findColumn(headers, ['uc', 'union council', 'union_council', 'area', 'location'], ignore, 2);
        }
        return headers[0];
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
        if (!data || data.length === 0) return;
        const headers = Object.keys(data[0]);
        // Use user selected calculation column or target Column G (7th column, index 6) for House Counts
        const houseCol = activeFilters.calcCol || headers[6] || headers[headers.length - 1];

        // Find role column: Look for "role", "designation", "category", or common health officer terms
        let roleCol = headers.find(h => {
            const lowerVal = h.toLowerCase();
            return lowerVal.includes('role') || lowerVal.includes('designation') || lowerVal.includes('category') || lowerVal.includes('position');
        }) || headers.find(h => {
            const sampleValues = data.slice(0, 5).map(row => String(row[h]).toLowerCase());
            return sampleValues.some(v => v.includes('health') || v.includes('worker') || v.includes('officer'));
        }) || (headers.length > 1 ? headers[headers.length - 2] : headers[0]);

        console.log(`Analyzing: Houses/Premises in [${houseCol}], Roles in [${roleCol}]`);

        // Standardize job titles
        standardizeJobTitles(data, roleCol);

        // Perform calculations
        const analyze = (subset) => {
            const totalUsers = subset.length;
            let activeUsers = 0;
            let totalHouses = 0;
            let dist = { '0': 0, '1-5': 0, '6-10': 0, '11+': 0 };

            subset.forEach(row => {
                let houseCount = row[houseCol];
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

        // Filter datasets based on standard keywords
        const LHW_KEYWORDS = ['lady health worker', 'lhw'];
        const CHO_KEYWORDS = ['community health officer', 'cho', 'chi', 'community health inspector'];

        const lhwData = data.filter(row => {
            const val = String(row[roleCol] || '').toLowerCase().replace(/\./g, '').trim();
            return val.includes('lady health worker') || val.includes('lhw');
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

        updateDashboard(overallResults, lhwResults, choResults, roleCounts, data, houseCol, roleCol);
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

    function updateDashboard(overall, lhw, cho, roleCounts, data, houseCol, roleCol) {
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

        // Populate Tehsil Summary Tab
        populateTehsilSummary(data, houseCol, roleCol);

        // Populate District Summary Tab
        populateDistrictSummary(data, houseCol, roleCol);

        // Populate UC Analysis Tab
        populateUCAnalysis(data, houseCol, roleCol);

        // Display dashboard
        loadingState.style.display = 'none';
        dashboardResults.style.display = 'block';
    }

    function populateUCAnalysis(data, houseCol, roleCol) {
        const body = document.getElementById('uc-summary-body');
        if (!currentRawData || currentRawData.length === 0) {
            if (body) body.innerHTML = '<tr><td colspan="22" class="text-center py-5 text-muted">No records found.</td></tr>';
            return;
        }
        const headers = Object.keys(currentRawData[0]);

        // Find UC Column
        const ucCol = getLocColumn('ucCol', headers);
        const tehsilCol = getLocColumn('tehsilCol', headers);

        // 1. Get all unique UCs from currentRawData (case-insensitive key normalization)
        const allUcs = [];
        const seenUcs = new Set();
        currentRawData.forEach(row => {
            const uc = String(row[ucCol] || 'Unknown UC').trim();
            const normUc = uc.toLowerCase();
            if (!seenUcs.has(normUc)) {
                seenUcs.add(normUc);
                allUcs.push(uc);
            }
        });

        // 2. Filter unique UCs based on active location filters
        let filteredUcs = allUcs;
        if (activeFilters.uc !== 'ALL') {
            filteredUcs = filteredUcs.filter(u => u.toLowerCase() === activeFilters.uc.toLowerCase());
        } else if (activeFilters.tehsil !== 'ALL') {
            const ucsInTehsil = new Set();
            currentRawData.forEach(row => {
                if (String(row[tehsilCol] || '').trim().toLowerCase() === activeFilters.tehsil.toLowerCase()) {
                    ucsInTehsil.add(String(row[ucCol] || '').trim().toLowerCase());
                }
            });
            filteredUcs = filteredUcs.filter(u => ucsInTehsil.has(u.toLowerCase()));
        } else if (activeFilters.district !== 'ALL') {
            const districtCol = getLocColumn('districtCol', headers);
            const ucsInDistrict = new Set();
            currentRawData.forEach(row => {
                if (String(row[districtCol] || '').trim().toLowerCase() === activeFilters.district.toLowerCase()) {
                    ucsInDistrict.add(String(row[ucCol] || '').trim().toLowerCase());
                }
            });
            filteredUcs = filteredUcs.filter(u => ucsInDistrict.has(u.toLowerCase()));
        }

        // 3. Initialize groups map for filtered UCs
        const groups = {};
        const keyMap = {};
        filteredUcs.forEach(uc => {
            const normUc = uc.toLowerCase();
            keyMap[normUc] = uc;

            const emptyGroup = () => ({ users: 0, active: 0, houses: 0, dist: { '0': 0, '1-5': 0, '6-10': 0, '11+': 0 } });
            groups[uc] = {
                total: emptyGroup(),
                lhw: emptyGroup(),
                cho: emptyGroup()
            };
        });

        // 4. Fill the groups using data
        const CHO_KEYWORDS = ['community health officer', 'cho', 'chi', 'community health inspector'];

        data.forEach(row => {
            const uc = String(row[ucCol] || 'Unknown UC').trim();
            const normUc = uc.toLowerCase();
            const displayKey = keyMap[normUc];

            if (displayKey && groups[displayKey]) {
                const role = String(row[roleCol] || '').toLowerCase().replace(/\./g, '').trim();
                const isLHW = role.includes('lady health worker') || role.includes('lhw');
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

                updateSubgroup(groups[displayKey].total);
                if (isLHW) updateSubgroup(groups[displayKey].lhw);
                if (isCHO) updateSubgroup(groups[displayKey].cho);
            }
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

    let tehsilFinalData = [];

    function populateTehsilSummary(data, houseCol, roleCol) {
        if (!currentRawData || currentRawData.length === 0) {
            tehsilFinalData = [];
            renderTehsilTable([]);
            const actionsWrapper = document.getElementById('tehsil-actions-wrapper');
            if (actionsWrapper) actionsWrapper.style.display = 'none';
            return;
        }
        const headers = Object.keys(currentRawData[0]);

        // Find District and Tehsil Columns
        const districtCol = getLocColumn('districtCol', headers);
        const tehsilCol = getLocColumn('tehsilCol', headers);

        // 1. Get all unique district-tehsil pairs from currentRawData (case-insensitive key normalization)
        const allTehsils = [];
        const seenKeys = new Set();
        currentRawData.forEach(row => {
            const district = String(row[districtCol] || 'N/A').trim();
            const tehsil = String(row[tehsilCol] || 'N/A').trim();
            const normKey = `${district.toLowerCase()}||${tehsil.toLowerCase()}`;
            if (!seenKeys.has(normKey)) {
                seenKeys.add(normKey);
                allTehsils.push({ district, tehsil });
            }
        });

        // 2. Filter the unique tehsils list based on location filters (District, Tehsil, UC)
        let filteredTehsils = allTehsils;
        if (activeFilters.district !== 'ALL') {
            filteredTehsils = filteredTehsils.filter(item => item.district.toLowerCase() === activeFilters.district.toLowerCase());
        }
        if (activeFilters.tehsil !== 'ALL') {
            filteredTehsils = filteredTehsils.filter(item => item.tehsil.toLowerCase() === activeFilters.tehsil.toLowerCase());
        }
        if (activeFilters.uc !== 'ALL') {
            const ucCol = getLocColumn('ucCol', headers);
            const tehsilsWithUc = new Set();
            currentRawData.forEach(row => {
                if (String(row[ucCol] || '').trim().toLowerCase() === activeFilters.uc.toLowerCase()) {
                    tehsilsWithUc.add(String(row[tehsilCol] || '').trim().toLowerCase());
                }
            });
            filteredTehsils = filteredTehsils.filter(item => tehsilsWithUc.has(item.tehsil.toLowerCase()));
        }

        // 3. Initialize groups map for the filtered tehsils
        const groups = {};
        const keyMap = {};
        filteredTehsils.forEach(item => {
            const displayKey = `${item.district}||${item.tehsil}`;
            const normKey = displayKey.toLowerCase();
            keyMap[normKey] = displayKey;

            const emptyGroup = () => ({
                users: 0,
                active: 0,
                zeroHouse: 0,
                houses: 0,
                dist: { '0': 0, '1-5': 0, '6-10': 0, '11+': 0 }
            });
            groups[displayKey] = {
                district: item.district,
                tehsil: item.tehsil,
                total: emptyGroup(),
                lhw: emptyGroup(),
                cho: emptyGroup()
            };
        });

        // 4. Fill the groups using data
        const CHO_KEYWORDS = ['community health officer', 'cho', 'chi', 'community health inspector'];

        data.forEach(row => {
            const district = String(row[districtCol] || 'N/A').trim();
            const tehsil = String(row[tehsilCol] || 'N/A').trim();
            const normKey = `${district.toLowerCase()}||${tehsil.toLowerCase()}`;
            const displayKey = keyMap[normKey];

            if (displayKey && groups[displayKey]) {
                const role = String(row[roleCol] || '').toLowerCase().replace(/\./g, '').trim();
                const isLHW = role.includes('lady health worker') || role.includes('lhw');
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
                    else g.zeroHouse++;

                    if (houseCount === 0) g.dist['0']++;
                    else if (houseCount >= 1 && houseCount <= 5) g.dist['1-5']++;
                    else if (houseCount >= 6 && houseCount <= 10) g.dist['6-10']++;
                    else if (houseCount >= 11) g.dist['11+']++;
                };

                updateSubgroup(groups[displayKey].total);
                if (isLHW) updateSubgroup(groups[displayKey].lhw);
                if (isCHO) updateSubgroup(groups[displayKey].cho);
            }
        });

        tehsilFinalData = Object.values(groups);
        
        renderTehsilTable(tehsilFinalData);

        // Show actions wrapper if we have data
        const actionsWrapper = document.getElementById('tehsil-actions-wrapper');
        if (actionsWrapper) {
            actionsWrapper.style.display = tehsilFinalData.length > 0 ? 'block' : 'none';
        }
    }

    function renderTehsilTable(tehsilList) {
        const body = document.getElementById('summary-list-body');
        if (!tehsilList || tehsilList.length === 0) {
            body.innerHTML = '<tr><td colspan="26" class="text-center py-5 text-muted">No records found.</td></tr>';
            return;
        }

        body.innerHTML = tehsilList.map(g => {
            const formatPct = (sub) => sub.users > 0 ? ((sub.zeroHouse / sub.users) * 100).toFixed(2) + '%' : '0.00%';

            return `
                <tr>
                    <td class="fw-bold sticky-column" style="left: 0; min-width: 130px; max-width: 130px; width: 130px; z-index: 5;">${g.district}</td>
                    <td class="fw-bold sticky-column" style="left: 130px; min-width: 130px; max-width: 130px; width: 130px; z-index: 5; border-right: 2px solid #e3e6f0 !important;">${g.tehsil}</td>
                    
                    <!-- Overall -->
                    <td class="table-primary-light">${g.total.users}</td>
                    <td class="table-primary-light">${g.total.active}</td>
                    <td class="table-primary-light">${g.total.zeroHouse}</td>
                    <td class="table-primary-light text-danger fw-bold">${formatPct(g.total)}</td>
                    <td class="table-primary-light">${g.total.houses.toLocaleString()}</td>
                    <td class="table-primary-light">${g.total.dist['1-5']}</td>
                    <td class="table-primary-light">${g.total.dist['6-10']}</td>
                    <td class="table-primary-light">${g.total.dist['11+']}</td>
                    
                    <!-- LHW -->
                    <td class="table-info-light">${g.lhw.users}</td>
                    <td class="table-info-light">${g.lhw.active}</td>
                    <td class="table-info-light">${g.lhw.zeroHouse}</td>
                    <td class="table-info-light text-danger fw-bold">${formatPct(g.lhw)}</td>
                    <td class="table-info-light">${g.lhw.houses.toLocaleString()}</td>
                    <td class="table-info-light">${g.lhw.dist['1-5']}</td>
                    <td class="table-info-light">${g.lhw.dist['6-10']}</td>
                    <td class="table-info-light">${g.lhw.dist['11+']}</td>
                    
                    <!-- CHI -->
                    <td class="table-success-light">${g.cho.users}</td>
                    <td class="table-success-light">${g.cho.active}</td>
                    <td class="table-success-light">${g.cho.zeroHouse}</td>
                    <td class="table-success-light text-danger fw-bold">${formatPct(g.cho)}</td>
                    <td class="table-success-light">${g.cho.houses.toLocaleString()}</td>
                    <td class="table-success-light">${g.cho.dist['1-5']}</td>
                    <td class="table-success-light">${g.cho.dist['6-10']}</td>
                    <td class="table-success-light">${g.cho.dist['11+']}</td>
                </tr>
            `;
        }).join('');
    }

    // Tehsil Filter Listener
    document.getElementById('tehsil-filter')?.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        const filtered = tehsilFinalData.filter(g => g.tehsil.toLowerCase().includes(query) || g.district.toLowerCase().includes(query));
        renderTehsilTable(filtered);
    });

    // Copy for Excel - Tehsil Summary
    document.getElementById('tehsil-copy-btn')?.addEventListener('click', () => {
        if (tehsilFinalData.length === 0) return;

        const headers = [
            'District', 'Tehsil',
            'Total Users (All)', 'Active (All)', '0 Houses (All)', 'Non Active % (All)', 'Total Houses (All)', '1-5 (All)', '6-10 (All)', '11+ (All)',
            'Total Users (LHW)', 'Active (LHW)', '0 Houses (LHW)', 'Non Active % (LHW)', 'Total Houses (LHW)', '1-5 (LHW)', '6-10 (LHW)', '11+ (LHW)',
            'Total Users (CHI)', 'Active (CHI)', '0 Houses (CHI)', 'Non Active % (CHI)', 'Total Houses (CHI)', '1-5 (CHI)', '6-10 (CHI)', '11+ (CHI)'
        ];

        const lines = [headers.join('\t')];

        tehsilFinalData.forEach(g => {
            const formatPct = (sub) => sub.users > 0 ? ((sub.zeroHouse / sub.users) * 100).toFixed(2) + '%' : '0.00%';
            
            const line = [
                g.district, g.tehsil,
                g.total.users, g.total.active, g.total.zeroHouse, formatPct(g.total), g.total.houses, g.total.dist['1-5'], g.total.dist['6-10'], g.total.dist['11+'],
                g.lhw.users, g.lhw.active, g.lhw.zeroHouse, formatPct(g.lhw), g.lhw.houses, g.lhw.dist['1-5'], g.lhw.dist['6-10'], g.lhw.dist['11+'],
                g.cho.users, g.cho.active, g.cho.zeroHouse, formatPct(g.cho), g.cho.houses, g.cho.dist['1-5'], g.cho.dist['6-10'], g.cho.dist['11+']
            ];
            lines.push(line.join('\t'));
        });

        const tsvText = lines.join('\n');
        const feedback = document.getElementById('tehsil-copy-feedback');

        const showSuccess = () => {
            if (feedback) {
                feedback.innerHTML = `<i class="fas fa-check-circle me-1"></i> Copied! Ready to paste into Excel.`;
                feedback.classList.add('show');
                setTimeout(() => feedback.classList.remove('show'), 4000);
            }
        };

        navigator.clipboard.writeText(tsvText).then(showSuccess).catch(() => {
            const ta = document.createElement('textarea');
            ta.value = tsvText;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
            showSuccess();
        });
    });

    // Export to Excel - Tehsil Summary
    document.getElementById('tehsil-export-btn')?.addEventListener('click', () => {
        if (tehsilFinalData.length === 0) return;

        const sheetData = tehsilFinalData.map(g => {
            const formatPct = (sub) => sub.users > 0 ? ((sub.zeroHouse / sub.users) * 100).toFixed(2) + '%' : '0.00%';
            return {
                "District": g.district,
                "Tehsil": g.tehsil,
                "Total Users (All)": g.total.users,
                "Active (All)": g.total.active,
                "0 Houses (All)": g.total.zeroHouse,
                "Non Active % (All)": formatPct(g.total),
                "Total Houses (All)": g.total.houses,
                "1-5 (All)": g.total.dist['1-5'],
                "6-10 (All)": g.total.dist['6-10'],
                "11+ (All)": g.total.dist['11+'],
                "Total Users (LHW)": g.lhw.users,
                "Active (LHW)": g.lhw.active,
                "0 Houses (LHW)": g.lhw.zeroHouse,
                "Non Active % (LHW)": formatPct(g.lhw),
                "Total Houses (LHW)": g.lhw.houses,
                "1-5 (LHW)": g.lhw.dist['1-5'],
                "6-10 (LHW)": g.lhw.dist['6-10'],
                "11+ (LHW)": g.lhw.dist['11+'],
                "Total Users (CHI)": g.cho.users,
                "Active (CHI)": g.cho.active,
                "0 Houses (CHI)": g.cho.zeroHouse,
                "Non Active % (CHI)": formatPct(g.cho),
                "Total Houses (CHI)": g.cho.houses,
                "1-5 (CHI)": g.cho.dist['1-5'],
                "6-10 (CHI)": g.cho.dist['6-10'],
                "11+ (CHI)": g.cho.dist['11+']
            };
        });

        const ws = XLSX.utils.json_to_sheet(sheetData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Tehsil Summary");
        XLSX.writeFile(wb, "Tehsil_Summary_Report.xlsx");
    });

    // --- DISTRICT SUMMARY LOGIC ---
    let districtFinalData = [];

    function populateDistrictSummary(data, houseCol, roleCol) {
        if (!currentRawData || currentRawData.length === 0) {
            districtFinalData = [];
            renderDistrictTable([]);
            const actionsWrapper = document.getElementById('district-actions-wrapper');
            if (actionsWrapper) actionsWrapper.style.display = 'none';
            return;
        }
        const headers = Object.keys(currentRawData[0]);

        // Find District Column
        const districtCol = getLocColumn('districtCol', headers);

        // 1. Get all unique districts from currentRawData (case-insensitive key normalization)
        const districts = [];
        const seenDistricts = new Set();
        currentRawData.forEach(row => {
            const district = String(row[districtCol] || 'N/A').trim();
            const normDist = district.toLowerCase();
            if (!seenDistricts.has(normDist)) {
                seenDistricts.add(normDist);
                districts.push(district);
            }
        });

        // 2. Filter unique districts based on active district filter
        let filteredDistricts = districts;
        if (activeFilters.district !== 'ALL') {
            filteredDistricts = filteredDistricts.filter(d => d.toLowerCase() === activeFilters.district.toLowerCase());
        }

        // 3. Initialize groups map for filtered districts
        const groups = {};
        const keyMap = {};
        filteredDistricts.forEach(dist => {
            const normDist = dist.toLowerCase();
            keyMap[normDist] = dist;

            const emptyGroup = () => ({
                users: 0,
                active: 0,
                zeroHouse: 0,
                houses: 0,
                dist: { '0': 0, '1-5': 0, '6-10': 0, '11+': 0 }
            });
            groups[dist] = {
                district: dist,
                total: emptyGroup(),
                lhw: emptyGroup(),
                cho: emptyGroup()
            };
        });

        // 4. Fill the groups using data
        const CHO_KEYWORDS = ['community health officer', 'cho', 'chi', 'community health inspector'];

        data.forEach(row => {
            const district = String(row[districtCol] || 'N/A').trim();
            const normDist = district.toLowerCase();
            const displayKey = keyMap[normDist];

            if (displayKey && groups[displayKey]) {
                const role = String(row[roleCol] || '').toLowerCase().replace(/\./g, '').trim();
                const isLHW = role.includes('lady health worker') || role.includes('lhw');
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
                    else g.zeroHouse++;

                    if (houseCount === 0) g.dist['0']++;
                    else if (houseCount >= 1 && houseCount <= 5) g.dist['1-5']++;
                    else if (houseCount >= 6 && houseCount <= 10) g.dist['6-10']++;
                    else if (houseCount >= 11) g.dist['11+']++;
                };

                updateSubgroup(groups[displayKey].total);
                if (isLHW) updateSubgroup(groups[displayKey].lhw);
                if (isCHO) updateSubgroup(groups[displayKey].cho);
            }
        });

        districtFinalData = Object.values(groups);
        
        renderDistrictTable(districtFinalData);

        // Show actions wrapper if we have data
        const actionsWrapper = document.getElementById('district-actions-wrapper');
        if (actionsWrapper) {
            actionsWrapper.style.display = districtFinalData.length > 0 ? 'block' : 'none';
        }
    }

    function renderDistrictTable(districtList) {
        const body = document.getElementById('district-summary-list-body');
        if (!districtList || districtList.length === 0) {
            body.innerHTML = '<tr><td colspan="25" class="text-center py-5 text-muted">No records found.</td></tr>';
            return;
        }

        body.innerHTML = districtList.map(g => {
            const formatPct = (sub) => sub.users > 0 ? ((sub.zeroHouse / sub.users) * 100).toFixed(2) + '%' : '0.00%';

            return `
                <tr>
                    <td class="fw-bold sticky-column" style="left: 0; min-width: 150px; max-width: 150px; width: 150px; z-index: 5; border-right: 2px solid #e3e6f0 !important;">${g.district}</td>
                    
                    <!-- Overall -->
                    <td class="table-primary-light">${g.total.users}</td>
                    <td class="table-primary-light">${g.total.active}</td>
                    <td class="table-primary-light">${g.total.zeroHouse}</td>
                    <td class="table-primary-light text-danger fw-bold">${formatPct(g.total)}</td>
                    <td class="table-primary-light">${g.total.houses.toLocaleString()}</td>
                    <td class="table-primary-light">${g.total.dist['1-5']}</td>
                    <td class="table-primary-light">${g.total.dist['6-10']}</td>
                    <td class="table-primary-light">${g.total.dist['11+']}</td>
                    
                    <!-- LHW -->
                    <td class="table-info-light">${g.lhw.users}</td>
                    <td class="table-info-light">${g.lhw.active}</td>
                    <td class="table-info-light">${g.lhw.zeroHouse}</td>
                    <td class="table-info-light text-danger fw-bold">${formatPct(g.lhw)}</td>
                    <td class="table-info-light">${g.lhw.houses.toLocaleString()}</td>
                    <td class="table-info-light">${g.lhw.dist['1-5']}</td>
                    <td class="table-info-light">${g.lhw.dist['6-10']}</td>
                    <td class="table-info-light">${g.lhw.dist['11+']}</td>
                    
                    <!-- CHI -->
                    <td class="table-success-light">${g.cho.users}</td>
                    <td class="table-success-light">${g.cho.active}</td>
                    <td class="table-success-light">${g.cho.zeroHouse}</td>
                    <td class="table-success-light text-danger fw-bold">${formatPct(g.cho)}</td>
                    <td class="table-success-light">${g.cho.houses.toLocaleString()}</td>
                    <td class="table-success-light">${g.cho.dist['1-5']}</td>
                    <td class="table-success-light">${g.cho.dist['6-10']}</td>
                    <td class="table-success-light">${g.cho.dist['11+']}</td>
                </tr>
            `;
        }).join('');
    }

    // District Filter Listener
    document.getElementById('district-filter')?.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        const filtered = districtFinalData.filter(g => g.district.toLowerCase().includes(query));
        renderDistrictTable(filtered);
    });

    // Copy for Excel - District Summary
    document.getElementById('district-copy-btn')?.addEventListener('click', () => {
        if (districtFinalData.length === 0) return;

        const headers = [
            'District',
            'Total Users (All)', 'Active (All)', '0 Houses (All)', 'Non Active % (All)', 'Total Houses (All)', '1-5 (All)', '6-10 (All)', '11+ (All)',
            'Total Users (LHW)', 'Active (LHW)', '0 Houses (LHW)', 'Non Active % (LHW)', 'Total Houses (LHW)', '1-5 (LHW)', '6-10 (LHW)', '11+ (LHW)',
            'Total Users (CHI)', 'Active (CHI)', '0 Houses (CHI)', 'Non Active % (CHI)', 'Total Houses (CHI)', '1-5 (CHI)', '6-10 (CHI)', '11+ (CHI)'
        ];

        const lines = [headers.join('\t')];

        districtFinalData.forEach(g => {
            const formatPct = (sub) => sub.users > 0 ? ((sub.zeroHouse / sub.users) * 100).toFixed(2) + '%' : '0.00%';
            
            const line = [
                g.district,
                g.total.users, g.total.active, g.total.zeroHouse, formatPct(g.total), g.total.houses, g.total.dist['1-5'], g.total.dist['6-10'], g.total.dist['11+'],
                g.lhw.users, g.lhw.active, g.lhw.zeroHouse, formatPct(g.lhw), g.lhw.houses, g.lhw.dist['1-5'], g.lhw.dist['6-10'], g.lhw.dist['11+'],
                g.cho.users, g.cho.active, g.cho.zeroHouse, formatPct(g.cho), g.cho.houses, g.cho.dist['1-5'], g.cho.dist['6-10'], g.cho.dist['11+']
            ];
            lines.push(line.join('\t'));
        });

        const tsvText = lines.join('\n');
        const feedback = document.getElementById('district-copy-feedback');

        const showSuccess = () => {
            if (feedback) {
                feedback.innerHTML = `<i class="fas fa-check-circle me-1"></i> Copied! Ready to paste into Excel.`;
                feedback.classList.add('show');
                setTimeout(() => feedback.classList.remove('show'), 4000);
            }
        };

        navigator.clipboard.writeText(tsvText).then(showSuccess).catch(() => {
            const ta = document.createElement('textarea');
            ta.value = tsvText;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
            showSuccess();
        });
    });

    // Export to Excel - District Summary
    document.getElementById('district-export-btn')?.addEventListener('click', () => {
        if (districtFinalData.length === 0) return;

        const sheetData = districtFinalData.map(g => {
            const formatPct = (sub) => sub.users > 0 ? ((sub.zeroHouse / sub.users) * 100).toFixed(2) + '%' : '0.00%';
            return {
                "District": g.district,
                "Total Users (All)": g.total.users,
                "Active (All)": g.total.active,
                "0 Houses (All)": g.total.zeroHouse,
                "Non Active % (All)": formatPct(g.total),
                "Total Houses (All)": g.total.houses,
                "1-5 (All)": g.total.dist['1-5'],
                "6-10 (All)": g.total.dist['6-10'],
                "11+ (All)": g.total.dist['11+'],
                "Total Users (LHW)": g.lhw.users,
                "Active (LHW)": g.lhw.active,
                "0 Houses (LHW)": g.lhw.zeroHouse,
                "Non Active % (LHW)": formatPct(g.lhw),
                "Total Houses (LHW)": g.lhw.houses,
                "1-5 (LHW)": g.lhw.dist['1-5'],
                "6-10 (LHW)": g.lhw.dist['6-10'],
                "11+ (LHW)": g.lhw.dist['11+'],
                "Total Users (CHI)": g.cho.users,
                "Active (CHI)": g.cho.active,
                "0 Houses (CHI)": g.cho.zeroHouse,
                "Non Active % (CHI)": formatPct(g.cho),
                "Total Houses (CHI)": g.cho.houses,
                "1-5 (CHI)": g.cho.dist['1-5'],
                "6-10 (CHI)": g.cho.dist['6-10'],
                "11+ (CHI)": g.cho.dist['11+']
            };
        });

        const ws = XLSX.utils.json_to_sheet(sheetData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "District Summary");
        XLSX.writeFile(wb, "District_Summary_Report.xlsx");
    });

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

    // --- ATTENDANT ANALYSIS LOGIC ---
    let attFinalData = [];

    const attFile1 = document.getElementById('att-file1');
    const attFile2 = document.getElementById('att-file2');
    const attWrapper1 = document.getElementById('att-wrapper1');
    const attWrapper2 = document.getElementById('att-wrapper2');
    const attName1 = document.getElementById('att-name1');
    const attName2 = document.getElementById('att-name2');
    const attProcessBtn = document.getElementById('att-process-btn');
    const attExportBtn = document.getElementById('att-export-btn');
    const attStatus = document.getElementById('att-status');

    if (attWrapper1) setupChiFileInput(attFile1, attWrapper1, attName1);
    if (attWrapper2) setupChiFileInput(attFile2, attWrapper2, attName2);

    function readExcelAsArray(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = e.target.result;
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    // Using header: 1 to get raw arrays for index-based access
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                    resolve(jsonData);
                } catch (err) { reject(err); }
            };
            reader.onerror = (err) => reject(err);
            reader.readAsArrayBuffer(file);
        });
    }

    attProcessBtn?.addEventListener('click', async () => {
        const f1 = attFile1.files[0];
        const f2 = attFile2.files[0];

        if (!f1 || !f2) {
            attStatus.innerHTML = '<span class="text-danger">Please upload both files for Attendant Analysis.</span>';
            return;
        }

        attProcessBtn.disabled = true;
        attProcessBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Processing...';
        attStatus.textContent = 'Analyzing data...';
        attFinalData = [];

        try {
            const data1 = await readExcelAsArray(f1);
            const data2 = await readExcelAsArray(f2);

            // Logic: 
            // File 1: Col E (index 4) = CNIC, Col F (index 5) = Designation
            // File 2: Col E (index 4) = CNIC, Col G (index 6) = Total Premises
            
            // Map File 2 for quick lookup
            const file2Map = new Map();
            data2.forEach((row, idx) => {
                if (idx === 0) return; // Skip header
                const cnic = String(row[4] || '').trim().replace(/[^0-9]/g, '');
                if (cnic) file2Map.set(cnic, row[6]);
            });

            const results = [];
            const file1Cnics = new Set(); // Tracks CHIs for Step 1
            const file1FullMap = new Map(); // Tracks ALL users for metadata lookup

            // Build full map from File 1 first
            data1.forEach((row, idx) => {
                if (idx === 0) return;
                const cnic = String(row[4] || '').trim().replace(/[^0-9]/g, '');
                if (cnic) {
                    file1FullMap.set(cnic, {
                        tehsil: String(row[1] || '').trim(),
                        uc: String(row[2] || '').trim(),
                        name: String(row[3] || '').trim()
                    });
                }
            });

            // 1. Process File 1 (filtered) for Matches and Disables
            data1.forEach((row, idx) => {
                if (idx === 0) return; // Skip header

                const designation = String(row[5] || '').toLowerCase();
                if (designation.includes("community health inspector") || designation.includes("chi")) {
                    const cnicRaw = String(row[4] || '').trim();
                    const cnic = cnicRaw.replace(/[^0-9]/g, '');
                    
                    file1Cnics.add(cnic);

                    const meta = file1FullMap.get(cnic);

                    if (file2Map.has(cnic)) {
                        results.push({ 
                            tehsil: meta.tehsil, uc: meta.uc, nameOfCadre: meta.name, 
                            cnic: cnicRaw, 
                            premises: file2Map.get(cnic),
                            status: "Match"
                        });
                    } else {
                        results.push({ 
                            tehsil: meta.tehsil, uc: meta.uc, nameOfCadre: meta.name, 
                            cnic: cnicRaw, 
                            premises: "0", 
                            status: "Disable"
                        });
                    }
                }
            });

            // 2. Process File 2 for "New" records (not in File 1 CHI list)
            data2.forEach((row, idx) => {
                if (idx === 0) return; // Skip header
                const cnicRaw = String(row[4] || '').trim();
                const cnic = cnicRaw.replace(/[^0-9]/g, '');
                
                if (cnic && !file1Cnics.has(cnic)) {
                    // Try to get metadata from File 1 first
                    let meta = file1FullMap.get(cnic);
                    
                    // If not in File 1, pick info from File 2 (Current row)
                    if (!meta) {
                        meta = {
                            tehsil: String(row[1] || '').trim() || "N/A",
                            uc: String(row[2] || '').trim() || "N/A",
                            name: String(row[3] || '').trim() || "N/A"
                        };
                    }

                    results.push({
                        tehsil: meta.tehsil,
                        uc: meta.uc,
                        nameOfCadre: meta.name,
                        cnic: cnicRaw,
                        premises: row[6] || "0",
                        status: "New"
                    });
                }
            });

            attFinalData = results;

            // Render
            const resultsBody = document.getElementById('att-results-body');
            if (results.length > 0) {
                resultsBody.innerHTML = results.map(r => {
                    let statusBadge = '';
                    if (r.status === 'Match') statusBadge = '<span class="badge bg-success px-2">Match</span>';
                    else if (r.status === 'New') statusBadge = '<span class="badge bg-primary px-2">New</span>';
                    else if (r.status === 'Disable') statusBadge = '<span class="badge bg-danger px-2">Disable</span>';

                    return `
                    <tr>
                        <td class="px-4">${r.tehsil}</td>
                        <td>${r.uc}</td>
                        <td>${r.nameOfCadre}</td>
                        <td>${r.cnic}</td>
                        <td class="fw-bold">${r.premises}</td>
                        <td>${statusBadge}</td>
                    </tr>
                `;}).join('');
                attStatus.innerHTML = `<span class="text-success">Processed ${results.length} records successfully!</span>`;
                attExportBtn.disabled = false;
            } else {
                resultsBody.innerHTML = '<tr><td colspan="6" class="text-center py-5">No records found.</td></tr>';
                attExportBtn.disabled = true;
            }

        } catch (err) {
            console.error(err);
            attStatus.innerHTML = `<span class="text-danger">Error: ${err.message}</span>`;
        } finally {
            attProcessBtn.disabled = false;
            attProcessBtn.innerHTML = '<i class="fas fa-sync me-2"></i> Process Attendant Data';
        }
    });

    attExportBtn?.addEventListener('click', () => {
        if (attFinalData.length === 0) return;
        const ws = XLSX.utils.json_to_sheet(attFinalData.map(d => ({ 
            "Tehsil": d.tehsil,
            "UC": d.uc,
            "Name of Cadre": d.nameOfCadre,
            "CNIC": d.cnic, 
            "Total Premises": d.premises,
            "Status": d.status
        })));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Analysis Results");
        XLSX.writeFile(wb, "Attendant_Analysis.xlsx");
    });

    // --- MULTI-FILE SUMMARY LOGIC ---
    const multiFileInput = document.getElementById('multi-excel-upload');
    const multiFileStatus = document.getElementById('multi-file-status');
    const multiSummaryContainer = document.getElementById('multi-summary-container');
    const multiCopyBtnWrapper = document.getElementById('multi-copy-btn-wrapper');
    const multiCopyAllBtn = document.getElementById('multi-copy-all-btn');
    const multiCopyFeedback = document.getElementById('multi-copy-feedback');

    // Stores data for all processed files so the copy button can access it
    let allFileSummaries = [];

    multiFileInput?.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        multiFileStatus.innerHTML = `<span class="spinner-border spinner-border-sm me-2"></span>Processing ${files.length} file(s)...`;
        multiSummaryContainer.innerHTML = '';
        multiCopyBtnWrapper.style.display = 'none';
        multiCopyFeedback.textContent = '';
        allFileSummaries = [];

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            try {
                const jsonData = await readExcelAsJSON(file);
                if (jsonData && jsonData.length > 0) {
                    const summaryData = calculateSummaryForFile(jsonData);
                    renderSummaryCard(file.name, summaryData, i);
                    allFileSummaries.push({ fileName: file.name, results: summaryData });
                } else {
                    renderErrorCard(file.name, "File appears to be empty.");
                }
            } catch (err) {
                renderErrorCard(file.name, err.message);
            }
        }
        
        multiFileStatus.innerHTML = `<span class="text-success">Finished processing ${files.length} file(s).</span>`;
        multiFileInput.value = ''; // Reset input

        // Show copy button only if we have at least one successful result
        if (allFileSummaries.length > 0) {
            multiCopyBtnWrapper.style.display = 'block';
        }
    });

    // --- COPY ALL FOR EXCEL ---
    multiCopyAllBtn?.addEventListener('click', () => {
        if (allFileSummaries.length === 0) return;

        const HEADERS = [
            'Category',
            'Total Users',
            'Active Users',
            '0 Houses',
            'Non Active User %',
            'Total Houses',
            '1-5 Houses',
            '6-10 Houses',
            '11+ Houses'
        ];

        const buildRow = (rowName, data) => {
            const nonActivePct = data.totalUsers > 0
                ? ((data.dist['0'] / data.totalUsers) * 100).toFixed(2) + '%'
                : '0.00%';
            return [
                rowName,
                data.totalUsers,
                data.activeUsers,
                data.dist['0'],
                nonActivePct,
                data.totalHouses,
                data.dist['1-5'],
                data.dist['6-10'],
                data.dist['11+']
            ].join('\t');
        };

        const tsvLines = [];

        allFileSummaries.forEach(({ fileName, results }, idx) => {
            // Blank separator between blocks (skip for first)
            if (idx > 0) tsvLines.push('');

            // File title row
            tsvLines.push(fileName);

            // Column headers
            tsvLines.push(HEADERS.join('\t'));

            // Data rows
            tsvLines.push(buildRow('Overall Users', results.overall));
            tsvLines.push(buildRow('Lady Health Workers (LHW)', results.lhw));
            tsvLines.push(buildRow('Community Health Inspector (CHI)', results.cho));
        });

        const tsvText = tsvLines.join('\n');

        const showSuccess = () => {
            multiCopyFeedback.innerHTML = `<i class="fas fa-check-circle me-1"></i> Copied! Ready to paste (Ctrl+V) into Excel.`;
            multiCopyFeedback.classList.add('show');
            setTimeout(() => {
                multiCopyFeedback.classList.remove('show');
            }, 4000);
        };

        navigator.clipboard.writeText(tsvText).then(showSuccess).catch(() => {
            // Fallback for browsers that block clipboard API
            const ta = document.createElement('textarea');
            ta.value = tsvText;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
            showSuccess();
        });
    });

    function calculateSummaryForFile(data) {
        const headers = Object.keys(data[0]);
        const houseCol = headers[6] || headers[headers.length - 1];
        
        let roleCol = headers.find(h => {
            const lowerVal = h.toLowerCase();
            return lowerVal.includes('role') || lowerVal.includes('designation') || lowerVal.includes('category') || lowerVal.includes('position');
        }) || headers.find(h => {
            const sampleValues = data.slice(0, 5).map(row => String(row[h]).toLowerCase());
            return sampleValues.some(v => v.includes('health') || v.includes('worker') || v.includes('officer'));
        }) || (headers.length > 1 ? headers[headers.length - 2] : headers[0]);

        standardizeJobTitles(data, roleCol);

        const analyze = (subset) => {
            const totalUsers = subset.length;
            let activeUsers = 0;
            let totalHouses = 0;
            let dist = { '0': 0, '1-5': 0, '6-10': 0, '11+': 0 };

            subset.forEach(row => {
                let houseCount = row[houseCol];
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

        const LHW_KEYWORDS = ['lady health worker', 'lhw'];
        const CHO_KEYWORDS = ['community health officer', 'cho', 'chi', 'community health inspector'];

        const lhwData = data.filter(row => {
            const val = String(row[roleCol] || '').trim().toLowerCase();
            return LHW_KEYWORDS.includes(val);
        });

        const choData = data.filter(row => {
            const val = String(row[roleCol] || '').toLowerCase();
            return CHO_KEYWORDS.some(k => val.includes(k));
        });

        return {
            overall: analyze(data),
            lhw: analyze(lhwData),
            cho: analyze(choData)
        };
    }

    function renderSummaryCard(fileName, results, index) {
        const rows = [
            { name: 'Overall Users', data: results.overall, class: 'fw-bold' },
            { name: 'Lady Health Workers (LHW)', data: results.lhw, class: '' },
            { name: 'Community Health Inspector (CHI)', data: results.cho, class: '' }
        ];

        const tbodyHTML = rows.map(row => {
            const nonActivePct = row.data.totalUsers > 0
                ? ((row.data.dist['0'] / row.data.totalUsers) * 100).toFixed(2) + '%'
                : '0.00%';

            return `
                <tr class="${row.class}">
                    <td>${row.name}</td>
                    <td>${row.data.totalUsers.toLocaleString()}</td>
                    <td>${row.data.activeUsers.toLocaleString()}</td>
                    <td>${row.data.dist['0'].toLocaleString()}</td>
                    <td class="text-danger fw-bold">${nonActivePct}</td>
                    <td>${row.data.totalHouses.toLocaleString()}</td>
                    <td>${row.data.dist['1-5'].toLocaleString()}</td>
                    <td>${row.data.dist['6-10'].toLocaleString()}</td>
                    <td>${row.data.dist['11+'].toLocaleString()}</td>
                </tr>
            `;
        }).join('');

        const cardHTML = `
            <div class="dashboard-card mb-4">
                <div class="card-header d-flex justify-content-between align-items-center">
                    <h3 class="card-title text-primary"><i class="fas fa-file-excel me-2"></i> ${fileName}</h3>
                </div>
                <div class="card-body p-0">
                    <div class="table-responsive">
                        <table class="table table-hover mb-0">
                            <thead class="table-light">
                                <tr>
                                    <th>Category</th>
                                    <th>Total Users</th>
                                    <th>Active Users</th>
                                    <th>0 Houses</th>
                                    <th>non active user %</th>
                                    <th>Total Houses</th>
                                    <th>1-5 Houses</th>
                                    <th>6-10 Houses</th>
                                    <th>11+ Houses</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${tbodyHTML}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        `;
        
        multiSummaryContainer.insertAdjacentHTML('beforeend', cardHTML);
    }

    function renderErrorCard(fileName, errorMessage) {
        const cardHTML = `
            <div class="dashboard-card mb-4 border-danger">
                <div class="card-header d-flex justify-content-between align-items-center">
                    <h3 class="card-title text-danger"><i class="fas fa-exclamation-triangle me-2"></i> ${fileName}</h3>
                </div>
                <div class="card-body">
                    <p class="text-danger mb-0">${errorMessage}</p>
                </div>
            </div>
        `;
        multiSummaryContainer.insertAdjacentHTML('beforeend', cardHTML);
    }
});
