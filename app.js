// Version: 3.5.2 (Customer Chain Dropdown and Optional Pack Classification)
// Initialize Lucide Icons & Node Listeners
document.addEventListener("DOMContentLoaded", () => {
  if (window.lucide) {
    window.lucide.createIcons();
  }
  const versionEl = document.getElementById('js-version-display');
  if (versionEl) {
    versionEl.textContent = '3.5.2 (Latest)';
  }
  
  // Customer Chain Dropdown Toggle
  const chainTrigger = document.getElementById('chain-select-trigger');
  const chainDropdown = document.getElementById('chain-select-dropdown');
  
  if (chainTrigger && chainDropdown) {
    chainTrigger.addEventListener('click', (e) => {
      e.stopPropagation();
      if (chainTrigger.hasAttribute('disabled')) return;
      chainDropdown.classList.toggle('hidden');
    });
    
    chainDropdown.addEventListener('click', (e) => {
      e.stopPropagation();
    });
    
    document.addEventListener('click', () => {
      chainDropdown.classList.add('hidden');
    });
  }
  
  // Excel Cleaner Script Panel Controls
  initExcelCleanerScript();
  
  // Atlan OLA Query Panel Controls
  initAtlanQueryScript();
});

// App State
const state = {
  olaFile: null,
  allocationFile: null,
  classificationFile: null,
  olaData: [], // Raw OLA rows
  allocationData: [], // Raw Allocation rows
  classificationMap: new Map(), // basepack -> classification override
  
  // Customer Chain Filter State
  availableChains: [],
  selectedChains: [],
  
  // Filtered Datasets
  filteredOla: [],
  filteredAllocation: [],
  
  // Analysis Results
  detailedBreakdown: [], // Detailed breakdown rows
  summaryRaisePo: [], // Pivot AOA
  
  // Matching lookup (bp_chain -> Allocation PO)
  matchingMap: new Map(),
  
  // Tree Data Structure
  treeData: {
    root: [],       // unique basepacks with average OLA < 80%
    accounts: {},   // AccountName -> { root: [], demand: [], supply: { all: [], classifications: { 'A': [], 'B': [], 'C': [], '#N/A': [] } } }
    demand: [],     // basepacks in Demand side (Raise PO)
    supply: {
      all: [],      // all open PO basepacks
      classifications: {
        'A': [],
        'B': [],
        'C': [],
        '#N/A': []
      }
    }
  },
  
  // Interactive Tree View State
  activeAccount: null,      // Selected account chain name in the tree
  activeBranch: null,       // 'demand' or 'supply'
  activeSubcategory: null,  // classification (A/B/C/#N/A) for demand, or 'fillRate' / 'ageing' for supply
  activeSupplyClass: null   // classification (A/B/C/#N/A) for supply issues (fillRate / ageing)
};

// Required columns validation (Note: Pack classification is optional in OLA)
const REQUIRED_COLUMNS = {
  ola: [
    'account', 'basepack', 'basepack_desc', 'business_unit', 'small_c', 
    'sales_category', 'brand', 'master_segment', 'location', 'pincode', 
    'week', 'month_year', 'depot', 'pack_tag', 'MRP', 'Average OLA'
  ],
  allocation: [
    'Basepack - Depot', 'Customer Chain', 'PO Date', 'PO Number', 'Requested Date', 
    'Sales Category', 'Basepack', 'Basepack Desc', 'Order Value Lacs', 
    'Invoiced Value Lacs', 'PO Status', 'Depot', 'PO Ageing', 'Po Ageing Bucket', 'Fill Rate'
  ],
  classification: [
    'Basepack', 'Pack classification'
  ]
};

// -------------------------------------------------------------
// DOM ELEMENTS & EVENT LISTENERS
// -------------------------------------------------------------
const themeToggle = document.getElementById('theme-toggle');
const olaDropzone = document.getElementById('ola-dropzone');
const allocationDropzone = document.getElementById('allocation-dropzone');
const classificationDropzone = document.getElementById('classification-dropzone');
const olaInput = document.getElementById('ola-file-input');
const allocationInput = document.getElementById('allocation-file-input');
const classificationInput = document.getElementById('classification-file-input');
const btnProcess = document.getElementById('btn-process');
const loader = document.getElementById('processing-loader');
const loaderMsg = document.getElementById('loader-message');
const weeksFilter = document.getElementById('weeks-filter');
const poDaysFilter = document.getElementById('po-days-filter');
const btnDownloadDemand = document.getElementById('btn-download-demand');
const btnDownloadSupply = document.getElementById('btn-download-supply');

// Theme Toggle
themeToggle.addEventListener('click', () => {
  document.body.classList.toggle('light-theme');
  document.body.classList.toggle('dark-theme');
  const isLight = document.body.classList.contains('light-theme');
  themeToggle.innerHTML = isLight ? '<i data-lucide="moon"></i>' : '<i data-lucide="sun"></i>';
  if (window.lucide) window.lucide.createIcons();
});

// Drag & Drop Handlers
function setupDropzone(dropzone, input, fileKey) {
  dropzone.addEventListener('click', (e) => {
    if (e.target !== input) {
      input.click();
    }
  });
  
  input.addEventListener('click', (e) => {
    e.stopPropagation();
  });
  
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });
  
  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });
  
  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
      handleFileSelect(e.dataTransfer.files[0], fileKey);
    }
  });
  
  input.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFileSelect(e.target.files[0], fileKey);
    }
  });
}

setupDropzone(olaDropzone, olaInput, 'ola');
setupDropzone(allocationDropzone, allocationInput, 'allocation');
setupDropzone(classificationDropzone, classificationInput, 'classification');

// Fuzzy Column Header Matching helper
function fuzzyMatchKey(reqKey, actualKeys) {
  const clean = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const reqClean = clean(reqKey);
  
  // 1. Exact cleaned match
  let found = actualKeys.find(k => clean(k) === reqClean);
  if (found) return found;
  
  // 2. Substring match fallback
  found = actualKeys.find(k => clean(k).includes(reqClean) || reqClean.includes(clean(k)));
  if (found) return found;
  
  // 3. Normalise abbreviations and common terminology variants
  const norm = (s) => clean(s)
    .replace(/invoiced/g, 'invoice')
    .replace(/number|num/g, 'no')
    .replace(/description/g, 'desc')
    .replace(/class+ification|classifcation/g, 'classification');
    
  const reqNorm = norm(reqKey);
  found = actualKeys.find(k => {
    const kNorm = norm(k);
    return kNorm === reqNorm || kNorm.includes(reqNorm) || reqNorm.includes(kNorm);
  });
  if (found) return found;
  
  // 4. Token-based matching for value and rate columns
  const getWords = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(w => w.length > 2);
  found = actualKeys.find(k => {
    const kWords = getWords(k);
    const kClean = clean(k);
    
    // Match 'Order Value' variants (e.g. Order Val, Order Value Lacs, Order Value)
    if (reqClean.includes('order') && reqClean.includes('value')) {
      return kWords.includes('order') && (kWords.includes('value') || kWords.includes('val') || kClean.includes('val'));
    }
    // Match 'Invoiced Value' variants (e.g. Invoice Val, Invoiced Value, Invoice Value)
    if (reqClean.includes('invoice') && reqClean.includes('value')) {
      return (kWords.includes('invoice') || kWords.includes('invoiced') || kClean.includes('invoice')) && 
             (kWords.includes('value') || kWords.includes('val') || kClean.includes('val'));
    }
    return false;
  });
  
  return found || null;
}

// File Processing Manager
function handleFileSelect(file, expectedKey) {
  const name = file.name.toLowerCase();
  let detectedKey = expectedKey;
  
  // Filename Keyword Auto-detection
  if (name.includes('ola') && !name.includes('allocation') && !name.includes('classification') && !name.includes('mapping')) {
    detectedKey = 'ola';
  } else if (name.includes('allocation')) {
    detectedKey = 'allocation';
  } else if (name.includes('classification') || name.includes('mapping')) {
    detectedKey = 'classification';
  }
  
  let zone = allocationDropzone;
  if (detectedKey === 'ola') zone = olaDropzone;
  else if (detectedKey === 'classification') zone = classificationDropzone;
  
  const details = document.getElementById(`${detectedKey}-file-details`);
  const instruction = zone ? zone.querySelector('.dropzone-instruction') : null;
  const keywordNode = zone ? zone.querySelector('.dropzone-keyword') : null;
  
  showLoader(`Reading ${file.name}...`);
  
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { 
        type: 'array', 
        cellDates: true,
        dense: true,
        cellStyles: false,
        cellFormula: false,
        cellHTML: false,
        cellText: false
      });
      
      let sheetName = '';
      let sheetData = [];
      
      // Look for the first sheet that actually has data
      for (const name of workbook.SheetNames) {
        const sheet = workbook.Sheets[name];
        
        // Fallback: If !ref is missing but !data is present, reconstruct it
        if (!sheet['!ref'] && sheet['!data']) {
          const maxRow = sheet['!data'].length - 1;
          let maxCol = 0;
          for (let r = 0; r <= maxRow; r++) {
            if (sheet['!data'][r] && sheet['!data'][r].length > maxCol) {
              maxCol = sheet['!data'][r].length;
            }
          }
          if (maxRow >= 0 && maxCol > 0) {
            sheet['!ref'] = XLSX.utils.encode_range(
              { r: 0, c: 0 },
              { r: maxRow, c: maxCol - 1 }
            );
          }
        }
        
        const parsed = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        if (parsed.length > 0) {
          sheetName = name;
          sheetData = parsed;
          break;
        }
      }
      
      if (sheetData.length === 0) {
        const firstSheetName = workbook.SheetNames && workbook.SheetNames.length > 0 ? workbook.SheetNames[0] : 'Unknown';
        throw new Error(`The workbook has no sheets containing data rows. First sheet "${firstSheetName}" was empty.`);
      }
      
      // Clean and normalize columns using fuzzy match
      const firstRow = sheetData[0];
      const actualKeys = Object.keys(firstRow);
      
      // Check column validation
      const required = REQUIRED_COLUMNS[detectedKey];
      const missing = required.filter(col => {
        if (col === 'Fill Rate') return false; // Optional column with fallback
        return fuzzyMatchKey(col, actualKeys) === null;
      });
      
      if (missing.length > 0) {
        const confirmMsg = `Warning: Missing expected columns in ${detectedKey.toUpperCase()} file:\n` + 
                           missing.join(', ') + 
                           `\n\nWe will try to process, but this might lead to errors. Continue?`;
        if (!confirm(confirmMsg)) {
          hideLoader();
          return;
        }
      }
      
      // Store cleaned data where keys are exactly the expected column names
      const cleanedData = sheetData.map(row => {
        const cleanedRow = {};
        required.forEach(reqKey => {
          const actualKey = fuzzyMatchKey(reqKey, actualKeys);
          cleanedRow[reqKey] = actualKey ? row[actualKey] : '';
        });
        
        // Retain unmapped keys just in case, but keep clean mapping
        Object.keys(row).forEach(k => {
          const cleanK = k.trim();
          if (!cleanedRow[cleanK]) {
            cleanedRow[cleanK] = row[k];
          }
        });
        return cleanedRow;
      });
      
      if (detectedKey === 'ola') {
        state.olaFile = file;
        state.olaData = cleanedData;
        
        // Extract unique weeks
        const weeks = [...new Set(cleanedData.map(r => String(r.week || '')))].filter(w => w.trim() !== '');
        
        if (zone) zone.classList.add('success');
        if (details) details.classList.remove('hidden');
        if (instruction) instruction.classList.add('hidden');
        if (keywordNode) keywordNode.classList.add('hidden');
        
        const fileNameEl = zone ? zone.querySelector('.file-name') : null;
        if (fileNameEl) fileNameEl.textContent = file.name;
        
        const rowCountEl = document.getElementById('ola-row-count');
        if (rowCountEl) rowCountEl.textContent = `${cleanedData.length.toLocaleString()} rows`;
        
        const weeksCountEl = document.getElementById('ola-weeks-count');
        if (weeksCountEl) weeksCountEl.textContent = `${weeks.length} weeks`;
      } else if (detectedKey === 'classification') {
        state.classificationFile = file;
        
        // Build classification lookup map
        state.classificationMap.clear();
        cleanedData.forEach(row => {
          const bp = String(row['Basepack'] || '').trim();
          const clsRaw = String(row['Pack classification'] || '').trim();
          if (bp && clsRaw) {
            const cls = (() => {
              const lower = clsRaw.toLowerCase();
              if (lower === 'a' || lower.includes('core')) return 'A';
              if (lower === 'b' || lower.includes('promo')) return 'B';
              if (lower === 'c' || lower.includes('new launch') || lower.includes('new')) return 'C';
              return clsRaw.toUpperCase();
            })();
            state.classificationMap.set(bp, cls);
          }
        });
        
        if (zone) zone.classList.add('success');
        if (details) details.classList.remove('hidden');
        if (instruction) instruction.classList.add('hidden');
        if (keywordNode) keywordNode.classList.add('hidden');
        
        const fileNameEl = zone ? zone.querySelector('.file-name') : null;
        if (fileNameEl) fileNameEl.textContent = file.name;
        
        const rowCountEl = document.getElementById('classification-row-count');
        if (rowCountEl) rowCountEl.textContent = `${state.classificationMap.size.toLocaleString()} basepacks`;
      } else {
        state.allocationFile = file;
        state.allocationData = cleanedData;
        
        // Find date range
        let minD = null, maxD = null;
        cleanedData.forEach(r => {
          const d = parseExcelDate(r['PO Date']);
          if (d) {
            if (!minD || d < minD) minD = d;
            if (!maxD || d > maxD) maxD = d;
          }
        });
        
        const dateRangeStr = minD && maxD ? 
          `${formatDate(minD)} to ${formatDate(maxD)}` : 
          'No valid PO dates';
          
        if (zone) zone.classList.add('success');
        if (details) details.classList.remove('hidden');
        if (instruction) instruction.classList.add('hidden');
        if (keywordNode) keywordNode.classList.add('hidden');
        
        const fileNameEl = zone ? zone.querySelector('.file-name') : null;
        if (fileNameEl) fileNameEl.textContent = file.name;
        
        const rowCountEl = document.getElementById('allocation-row-count');
        if (rowCountEl) rowCountEl.textContent = `${cleanedData.length.toLocaleString()} rows`;
        
        const dateRangeEl = document.getElementById('allocation-date-range');
        if (dateRangeEl) dateRangeEl.textContent = dateRangeStr;
        
        // Populate customer chains dropdown
        populateCustomerChainsDropdown();
      }
      
      // Enable process button if both files are loaded
      if (state.olaData.length > 0 && state.allocationData.length > 0) {
        if (btnProcess) btnProcess.removeAttribute('disabled');
      }
      
      hideLoader();
    } catch (err) {
      hideLoader();
      alert(`Error reading Excel file: ${err.message}`);
      console.error(err);
    }
  };
  
  reader.readAsArrayBuffer(file);
}

// Helper to show/hide loaders
function showLoader(message) {
  const msgEl = document.getElementById('loader-message');
  if (msgEl) {
    msgEl.textContent = message;
  }
  const loaderEl = document.getElementById('processing-loader');
  if (loaderEl) {
    loaderEl.classList.remove('hidden');
  }
}

function hideLoader() {
  const loaderEl = document.getElementById('processing-loader');
  if (loaderEl) {
    loaderEl.classList.add('hidden');
  }
}

// -------------------------------------------------------------
// FILTERING & DECISION-TREE MAPPING
// -------------------------------------------------------------
if (btnProcess) {
  btnProcess.addEventListener('click', runAnalysis);
}
if (btnDownloadDemand) {
  btnDownloadDemand.addEventListener('click', downloadExcelWorkbook);
}
if (btnDownloadSupply) {
  btnDownloadSupply.addEventListener('click', downloadSupplyExcelWorkbook);
}

function runAnalysis() {
  if (state.olaData.length === 0 || state.allocationData.length === 0) return;
  
  showLoader("Filtering datasets and executing decision-tree mapping...");
  
  setTimeout(() => {
    try {
      // Apply Classification Mapping overrides if uploaded
      if (state.classificationMap.size > 0) {
        state.olaData.forEach(row => {
          const bp = String(row.basepack || '').trim();
          if (state.classificationMap.has(bp)) {
            row['Pack classification'] = state.classificationMap.get(bp);
          }
        });
      }

      // 1. FILTER OLA TO LATEST N WEEKS
      const filterWeeksEl = document.getElementById('weeks-filter') || weeksFilter;
      const weeksToKeepCount = filterWeeksEl ? parseInt(filterWeeksEl.value, 10) : 2;
      filterOlaData(weeksToKeepCount);
      
      // 2. FILTER ALLOCATION POs BY DATE
      const filterPoDaysEl = document.getElementById('po-days-filter') || poDaysFilter;
      const poFilterMode = filterPoDaysEl ? filterPoDaysEl.value : '20';
      filterAllocationData(poFilterMode);
      
      // 3. FILTER BY SELECTED CUSTOMER CHAINS
      const selectedChains = state.selectedChains || [];
      state.filteredOla = state.filteredOla.filter(row => {
        const canonical = getCanonicalChainName(row.account);
        return selectedChains.includes(canonical);
      });
      state.filteredAllocation = state.filteredAllocation.filter(row => {
        const canonical = getCanonicalChainName(row['Customer Chain']);
        return selectedChains.includes(canonical);
      });
      
      // 4. EXECUTE MAPPING
      executeDecisionTreeAnalysis();
      
    } catch (err) {
      hideLoader();
      alert(`Error during analysis: ${err.message}`);
      console.error(err);
    }
  }, 100);
}

function filterOlaData(nWeeks) {
  // Matches "week 20-2026", "week 20 2026", or just "week 20", "week 21"
  const weekRegex = /(?:week\s*)?(\d+)(?:[-\s](\d{4}))?/i;
  
  const parsedWeeks = state.olaData.map((row) => {
    const wVal = String(row.week || '').trim();
    const match = wVal.match(weekRegex);
    
    if (match) {
      const weekNum = parseInt(match[1], 10);
      let year = match[2] ? parseInt(match[2], 10) : null;
      
      // Fallback for missing year in week column
      if (!year) {
        // Try parsing year from month_year. E.g. "May-26" or "May-2026"
        const myVal = String(row.month_year || '').trim();
        const yearMatch = myVal.match(/(?:\d{4}|\d{2})$/);
        if (yearMatch) {
          let y = parseInt(yearMatch[0], 10);
          year = y < 100 ? 2000 + y : y;
        } else {
          year = 2026; // Default fallback to current year
        }
      }
      
      return {
        row,
        weekStr: wVal,
        weekNum: weekNum,
        year: year,
        key: `${year}-${String(weekNum).padStart(2, '0')}` // YYYY-WW sort key
      };
    }
    return { row, weekStr: wVal, weekNum: 0, year: 0, key: '' };
  }).filter(item => item.key !== '');
  
  // Find unique weeks and sort descending (year desc, week desc)
  const uniqueWeeks = [...new Set(parsedWeeks.map(item => item.weekStr))];
  const sortedUniqueWeeks = uniqueWeeks.map(wStr => {
    const item = parsedWeeks.find(i => i.weekStr === wStr);
    return { weekStr: wStr, key: item.key, year: item.year, weekNum: item.weekNum };
  }).sort((a, b) => b.key.localeCompare(a.key));
  
  // Keep only latest nWeeks
  const latestWeeksToKeep = sortedUniqueWeeks.slice(0, nWeeks).map(w => w.weekStr);
  
  // Filter OLA
  state.filteredOla = state.olaData.filter(row => {
    const wVal = String(row.week || '').trim();
    return latestWeeksToKeep.includes(wVal);
  });
  
  console.log(`Filtered OLA from ${state.olaData.length} to ${state.filteredOla.length} rows (Latest ${nWeeks} weeks: ${latestWeeksToKeep.join(', ')})`);
}

// Filter Allocation PO Dates to last X Days of maximum PO date (default 20 days)
function filterAllocationData(daysMode) {
  if (daysMode === 'all') {
    state.filteredAllocation = [...state.allocationData];
    return;
  }
  
  const daysLimit = parseInt(daysMode, 10);
  
  // Parse all PO Dates and find max date
  let maxDate = null;
  const rowsWithParsedDates = state.allocationData.map(row => {
    const d = parseExcelDate(row['PO Date']);
    return { row, date: d };
  }).filter(item => item.date !== null);
  
  if (rowsWithParsedDates.length === 0) {
    state.filteredAllocation = [...state.allocationData];
    return;
  }
  
  // Find max date
  rowsWithParsedDates.forEach(item => {
    if (!maxDate || item.date > maxDate) maxDate = item.date;
  });
  
  // Compute minimum date cut-off (inclusive)
  const cutOffDate = new Date(maxDate.getTime());
  cutOffDate.setDate(maxDate.getDate() - daysLimit);
  cutOffDate.setHours(0,0,0,0);
  
  state.filteredAllocation = rowsWithParsedDates
    .filter(item => item.date >= cutOffDate)
    .map(item => item.row);
    
  console.log(`Filtered Allocation: Max PO Date is ${formatDate(maxDate)}. Cutoff is ${formatDate(cutOffDate)}. Rows kept: ${state.filteredAllocation.length} / ${state.allocationData.length}`);
}

// Conversion: Week string to Date Range
// weekStr e.g. "week 20-2026"
function getWeekDateRange(weekStr) {
  const match = weekStr.match(/(?:week\s*)?(\d+)[-\s](\d{4})/i);
  if (!match) return null;
  
  const weekNum = parseInt(match[1], 10);
  const year = parseInt(match[2], 10);
  
  // ISO Week 1 is the week with the first Thursday
  const jan4 = new Date(year, 0, 4);
  const dayOfWeek = jan4.getDay() === 0 ? 7 : jan4.getDay(); // 1 = Mon, 7 = Sun
  
  // Monday of Week 1
  const mondayW1 = new Date(jan4);
  mondayW1.setDate(jan4.getDate() - dayOfWeek + 1);
  
  // Monday of target week
  const monday = new Date(mondayW1);
  monday.setDate(mondayW1.getDate() + (weekNum - 1) * 7);
  
  // Sunday of target week
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  
  // Clear times
  monday.setHours(0,0,0,0);
  sunday.setHours(23,59,59,999);
  
  return { start: monday, end: sunday };
}

// Check month matching helper
// monthYearStr is OLA's month_year e.g. "May-2026" or "May-26" or date object
function matchesMonthYear(poDate, monthYearStr) {
  if (!poDate || !monthYearStr) return false;
  
  const d = parseExcelDate(poDate);
  if (!d) return false;
  
  // Parse monthYearStr. E.g. "May-2026"
  const parts = String(monthYearStr).split('-');
  if (parts.length < 2) return false;
  
  const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const targetMonth = parts[0].trim().toLowerCase().substring(0, 3);
  let targetYearStr = parts[1].trim();
  
  // Handle 2-digit vs 4-digit year
  let targetYear = parseInt(targetYearStr, 10);
  if (targetYearStr.length === 2) {
    targetYear += 2000;
  }
  
  const poMonth = d.getMonth(); // 0-11
  const poYear = d.getFullYear();
  
  const monthIdx = months.indexOf(targetMonth);
  if (monthIdx === -1) return false;
  
  return poMonth === monthIdx && poYear === targetYear;
}

// Core Matching Algorithm
// Core Matching Algorithm
function executeDecisionTreeAnalysis() {
  state.detailedBreakdown = [];
  state.summaryRaisePo = [];
  state.matchingMap.clear();
  
  console.log("=== RUNNING ANALYSIS ===");
  if (state.filteredAllocation.length > 0) {
    const keys = Object.keys(state.filteredAllocation[0]);
    console.log("Allocation clean keys:", keys);
    console.log("Sample allocation row:", state.filteredAllocation[0]);
  }
  
  const hasClassification = state.classificationMap.size > 0;
  
  // 1. Pre-process Allocation POs and group by (Basepack, Customer Chain)
  const allocationByBpChain = new Map();
  state.filteredAllocation.forEach(po => {
    const bp = String(po.Basepack || '').trim();
    const chain = getCanonicalChainName(po['Customer Chain']);
    if (!bp || !chain) return;
    
    const key = `${bp}_${chain}`;
    if (!allocationByBpChain.has(key)) {
      allocationByBpChain.set(key, []);
    }
    allocationByBpChain.get(key).push(po);
  });
  
  // 2. Identify all (basepack, chain) groups in OLA globally to match POs
  const bpChainGroups = new Map();
  state.filteredOla.forEach(row => {
    const bp = String(row.basepack || '').trim();
    const chain = getCanonicalChainName(row.account);
    if (!bp || !chain) return;
    
    const key = `${bp}_${chain}`;
    if (!bpChainGroups.has(key)) {
      bpChainGroups.set(key, []);
    }
    bpChainGroups.get(key).push(row);
  });
  
  const bpChainsNoPo = new Set();
  const matchedPosByBpChain = new Map();
  
  bpChainGroups.forEach((rows, key) => {
    const candidatePOs = allocationByBpChain.get(key) || [];
    if (candidatePOs.length === 0) {
      bpChainsNoPo.add(key);
    } else {
      let selectedPO = candidatePOs[0];
      let maxDate = parseExcelDate(selectedPO['PO Date']);
      
      for (let i = 1; i < candidatePOs.length; i++) {
        const d = parseExcelDate(candidatePOs[i]['PO Date']);
        if (d && (!maxDate || d > maxDate)) {
          maxDate = d;
          selectedPO = candidatePOs[i];
        }
      }
      matchedPosByBpChain.set(key, selectedPO);
      state.matchingMap.set(key, selectedPO); // Store for explorer reference
    }
  });
  
  // 3. Group OLA by (basepack, chain, location) to compute average OLA (For Excel)
  const bpChainLocGroups = new Map(); // key: "bp_chain_loc" -> { basepack, chain, location, rows: [] }
  state.filteredOla.forEach(row => {
    const bp = String(row.basepack || '').trim();
    const chain = getCanonicalChainName(row.account);
    const loc = String(row.location || '').trim();
    if (!bp || !chain || !loc) return;
    
    const key = `${bp}_${chain}_${loc}`;
    if (!bpChainLocGroups.has(key)) {
      bpChainLocGroups.set(key, {
        basepack: bp,
        chain: chain,
        location: loc,
        rows: []
      });
    }
    bpChainLocGroups.get(key).rows.push(row);
  });
  
  // 4. Filter and build detailedBreakdown (For Excel)
  bpChainLocGroups.forEach(grp => {
    const bp = grp.basepack;
    const chain = grp.chain;
    const loc = grp.location;
    const bpChainKey = `${bp}_${chain}`;
    const avgOla = grp.rows.reduce((sum, r) => sum + (Number(r['Average OLA']) || 0), 0) / grp.rows.length;
    
    if (avgOla < 80) {
      let shouldRaise = false;
      
      if (bpChainsNoPo.has(bpChainKey)) {
        shouldRaise = true;
      } else {
        const po = matchedPosByBpChain.get(bpChainKey);
        const status = String(po['PO Status'] || '').trim().toLowerCase();
        if (status === 'closed po' || status === 'closed') {
          shouldRaise = true;
        }
      }
      
      if (shouldRaise) {
        const firstRow = grp.rows[0] || {};
        state.detailedBreakdown.push({
          'Basepack': bp,
          'basepack_desc': firstRow.basepack_desc || (matchedPosByBpChain.has(bpChainKey) ? matchedPosByBpChain.get(bpChainKey)['Basepack Desc'] : ''),
          'Customer Chain': chain,
          'small_c': firstRow.small_c || '',
          'brand': firstRow.brand || '',
          'location': loc,
          'Pack classification': hasClassification && firstRow['Pack classification'] ? String(firstRow['Pack classification']).trim() : '',
          'pack_tag': firstRow.pack_tag || '',
          'average ola': Number(avgOla.toFixed(2))
        });
      }
    }
  });
  
  // 5. Build Tree Hierarchy Data (For UI Tree Explorer)
  state.treeData = {
    root: [],
    accounts: {},
    demand: [],
    supply: {
      all: [],
      classifications: {
        'A': [],
        'B': [],
        'C': [],
        '#N/A': []
      }
    }
  };
 
  bpChainGroups.forEach((rows, key) => {
    const parts = key.split('_');
    const bp = parts[0];
    const chain = parts[1];
    
    const avgOla = rows.reduce((sum, r) => sum + (Number(r['Average OLA']) || 0), 0) / rows.length;
    if (avgOla < 80) {
      const firstRow = rows[0] || {};
      const bpObj = {
        basepack: bp,
        desc: firstRow.basepack_desc || (matchedPosByBpChain.has(key) ? matchedPosByBpChain.get(key)['Basepack Desc'] : ''),
        customerChain: chain,
        brand: firstRow.brand || '',
        small_c: firstRow.small_c || '',
        classification: (() => {
          if (!hasClassification) return '#N/A';
          const rawVal = firstRow['Pack classification'];
          const raw = rawVal !== undefined && rawVal !== null ? String(rawVal).trim() : '';
          if (!raw) return '#N/A';
          const lower = raw.toLowerCase();
          if (lower === 'a' || lower.includes('core')) return 'A';
          if (lower === 'b' || lower.includes('promo')) return 'B';
          if (lower === 'c' || lower.includes('new launch') || lower.includes('new')) return 'C';
          return raw.toUpperCase();
        })(),
        avgOla: avgOla,
        rows: rows,
        pos: allocationByBpChain.get(key) || []
      };
      
      state.treeData.root.push(bpObj);
      
      // Initialize account structure if not exists
      if (!state.treeData.accounts[chain]) {
        state.treeData.accounts[chain] = {
          root: [],
          demand: [],
          supply: {
            all: [],
            classifications: {
              'A': [],
              'B': [],
              'C': [],
              '#N/A': []
            }
          }
        };
      }
      
      const accData = state.treeData.accounts[chain];
      accData.root.push(bpObj);
      
      // Determine Demand or Supply side
      let isClosed = false;
      let mostRecentPo = null;
      if (bpObj.pos.length > 0) {
        mostRecentPo = bpObj.pos[0];
        let maxDate = parseExcelDate(mostRecentPo['PO Date']);
        for (let i = 1; i < bpObj.pos.length; i++) {
          const d = parseExcelDate(bpObj.pos[i]['PO Date']);
          if (d && (!maxDate || d > maxDate)) {
            maxDate = d;
            mostRecentPo = bpObj.pos[i];
          }
        }
        const status = String(mostRecentPo['PO Status'] || '').trim().toLowerCase();
        if (status === 'closed po' || status === 'closed') {
          isClosed = true;
        }
      }
      
      const isDemand = (bpObj.pos.length === 0 || isClosed);
      if (isDemand) {
        bpObj.mostRecentPo = mostRecentPo;
        state.treeData.demand.push(bpObj);
        accData.demand.push(bpObj);
      } else {
        // Supply Side: open PO exists
        const openPos = bpObj.pos.filter(po => {
          const status = String(po['PO Status'] || '').trim().toLowerCase();
          return status === 'open po' || status === 'open';
        });
        
        bpObj.openPos = openPos;
        state.treeData.supply.all.push(bpObj);
        accData.supply.all.push(bpObj);
        
        // Calculate average PO Ageing
        let ageingSum = 0;
        let ageingCount = 0;
        openPos.forEach(po => {
          const ageing = Number(po['PO Ageing']) || 0;
          ageingSum += ageing;
          ageingCount++;
        });
        const avgAgeing = ageingCount > 0 ? (ageingSum / ageingCount) : null;
        bpObj.avgAgeing = avgAgeing;
        
        // Threshold check:
        // A > 3 days, B > 7 days, C > 11 days, unclassified > 7 days
        const cls = bpObj.classification;
        let isProblem = false;
        if (avgAgeing !== null) {
          if (hasClassification) {
            if (cls === 'A') {
              isProblem = avgAgeing > 3;
            } else if (cls === 'B') {
              isProblem = avgAgeing > 7;
            } else if (cls === 'C') {
              isProblem = avgAgeing > 11;
            } else {
              isProblem = avgAgeing > 7;
            }
          } else {
            // Default threshold is 7 days if no classification uploaded
            isProblem = avgAgeing > 7;
          }
        }
        bpObj.isProblem = isProblem;
        
        if (isProblem) {
          if (state.treeData.supply.classifications[cls]) {
            state.treeData.supply.classifications[cls].push(bpObj);
          } else {
            state.treeData.supply.classifications['#N/A'].push(bpObj);
          }
          
          if (accData.supply.classifications[cls]) {
            accData.supply.classifications[cls].push(bpObj);
          } else {
            accData.supply.classifications['#N/A'].push(bpObj);
          }
        }
      }
    }
  });
  
  // 6. Construct Summary Pivot (only if classification uploaded)
  generateSummaryRaisePoPivot();
  
  // 7. Render Interactive Tree
  renderInteractiveTree();
}

// -------------------------------------------------------------
// INTERACTIVE DECISION TREE RENDER FUNCTIONS
// -------------------------------------------------------------
function getSupplyProblemCount(account) {
  const accData = state.treeData.accounts[account];
  if (!accData) return 0;
  
  const hasClassification = state.classificationMap.size > 0;
  if (hasClassification) {
    return (
      accData.supply.classifications['A'].length + 
      accData.supply.classifications['B'].length + 
      accData.supply.classifications['C'].length + 
      accData.supply.classifications['#N/A'].length
    );
  } else {
    return accData.supply.all.filter(bp => bp.isProblem).length;
  }
}

function getAccountStyle(accountName) {
  const name = String(accountName).trim().toLowerCase();
  if (name === 'purplle') {
    return {
      color: '#d946ef',
      bgClass: 'rgba(217, 70, 239, 0.15)',
      icon: 'shopping-bag'
    };
  } else if (name === 'nykaa') {
    return {
      color: '#ec4899',
      bgClass: 'rgba(236, 72, 153, 0.15)',
      icon: 'shopping-bag'
    };
  } else if (name === 'myntra') {
    return {
      color: '#f43f5e',
      bgClass: 'rgba(244, 63, 94, 0.15)',
      icon: 'shopping-bag'
    };
  } else {
    return {
      color: '#06b6d4',
      bgClass: 'rgba(6, 182, 212, 0.15)',
      icon: 'box'
    };
  }
}

function selectTreeAccount(account) {
  state.activeAccount = account;
  state.activeBranch = null;
  state.activeSubcategory = null;
  state.activeSupplyClass = null;
  
  // Highlight active account card
  const cards = document.querySelectorAll('#grid-accounts .tree-node-card');
  cards.forEach(card => {
    if (card.dataset.account === account) {
      card.classList.add('active');
    } else {
      card.classList.remove('active');
    }
  });
  
  // Update counts on branches
  const demandCountEl = document.getElementById('val-demand-count');
  const supplyCountEl = document.getElementById('val-supply-count');
  
  const accData = state.treeData.accounts[account];
  const demandCount = accData ? accData.demand.length : 0;
  const supplyCount = getSupplyProblemCount(account);
  
  if (demandCountEl) demandCountEl.textContent = demandCount.toLocaleString();
  if (supplyCountEl) supplyCountEl.textContent = supplyCount.toLocaleString();
  
  // Update branches row title
  const branchesTitleEl = document.getElementById('lbl-branches-title');
  if (branchesTitleEl) {
    branchesTitleEl.textContent = `${account} Analysis Branches`;
  }
  
  // Show Branches Row and Connector
  const connBranches = document.getElementById('conn-to-branches');
  const rowBranches = document.getElementById('tree-row-branches');
  if (connBranches) connBranches.classList.remove('hidden');
  if (rowBranches) rowBranches.classList.remove('hidden');
  
  // Reset branch card active highlight states
  const cardDemand = document.getElementById('node-branch-demand');
  const cardSupply = document.getElementById('node-branch-supply');
  if (cardDemand) cardDemand.classList.remove('active');
  if (cardSupply) cardSupply.classList.remove('active');
  
  // Hide all lower rows
  const rowSub = document.getElementById('tree-row-subbranches');
  const connSub = document.getElementById('conn-to-subbranches');
  const rowCls = document.getElementById('tree-row-classifications');
  const connCls = document.getElementById('conn-to-classifications');
  const rowDetails = document.getElementById('tree-row-details');
  const connDetails = document.getElementById('conn-to-details');
  
  if (rowSub) rowSub.classList.add('hidden');
  if (connSub) connSub.classList.add('hidden');
  if (rowCls) rowCls.classList.add('hidden');
  if (connCls) connCls.classList.add('hidden');
  if (rowDetails) rowDetails.classList.add('hidden');
  if (connDetails) connDetails.classList.add('hidden');
  
  setTimeout(() => {
    if (rowBranches) rowBranches.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }, 50);
}

function renderInteractiveTree() {
  const rootCountEl = document.getElementById('val-root-count');
  if (rootCountEl) rootCountEl.textContent = state.treeData.root.length.toLocaleString();
  
  // Render dynamic account cards
  const gridAccounts = document.getElementById('grid-accounts');
  if (gridAccounts) {
    gridAccounts.innerHTML = '';
    
    // Sort accounts so Purplle, Nykaa, Myntra are first
    const sortedSelectedChains = [...(state.selectedChains || [])].sort((a, b) => {
      const priority = { 'purplle': 1, 'nykaa': 2, 'myntra': 3 };
      const pA = priority[a.toLowerCase()] || 99;
      const pB = priority[b.toLowerCase()] || 99;
      if (pA !== pB) return pA - pB;
      return a.localeCompare(b);
    });
    
    sortedSelectedChains.forEach(account => {
      const accData = state.treeData.accounts[account];
      const count = accData ? accData.root.length : 0;
      const style = getAccountStyle(account);
      
      const card = document.createElement('div');
      card.className = 'tree-node-card';
      card.dataset.account = account;
      card.style.borderWidth = '1.5px';
      
      card.innerHTML = `
        <div class="node-card-icon" style="color: ${style.color}; background: ${style.bgClass};">
          <i data-lucide="${style.icon}"></i>
        </div>
        <div class="node-card-title">${account}</div>
        <div class="node-card-desc">OLA Basepacks</div>
        <div class="node-card-badge" style="background: ${style.color};">${count}</div>
      `;
      
      card.addEventListener('click', () => {
        selectTreeAccount(account);
      });
      gridAccounts.appendChild(card);
    });
  }
  
  // Unhide the dashboard card
  const dashboardEl = document.getElementById('analysis-dashboard');
  if (dashboardEl) {
    dashboardEl.classList.remove('hidden');
  }
  
  // Clear active tree states and hide lower rows
  state.activeAccount = null;
  state.activeBranch = null;
  state.activeSubcategory = null;
  state.activeSupplyClass = null;
  
  const connBranches = document.getElementById('conn-to-branches');
  const rowBranches = document.getElementById('tree-row-branches');
  const rowSub = document.getElementById('tree-row-subbranches');
  const connSub = document.getElementById('conn-to-subbranches');
  const rowCls = document.getElementById('tree-row-classifications');
  const connCls = document.getElementById('conn-to-classifications');
  const rowDetails = document.getElementById('tree-row-details');
  const connDetails = document.getElementById('conn-to-details');
  
  if (connBranches) connBranches.classList.add('hidden');
  if (rowBranches) rowBranches.classList.add('hidden');
  if (rowSub) rowSub.classList.add('hidden');
  if (connSub) connSub.classList.add('hidden');
  if (rowCls) rowCls.classList.add('hidden');
  if (connCls) connCls.classList.add('hidden');
  if (rowDetails) rowDetails.classList.add('hidden');
  if (connDetails) connDetails.classList.add('hidden');
  
  const cardDemand = document.getElementById('node-branch-demand');
  const cardSupply = document.getElementById('node-branch-supply');
  if (cardDemand) cardDemand.classList.remove('active');
  if (cardSupply) cardSupply.classList.remove('active');
  
  // Re-attach level 2 click events once
  if (cardDemand && !cardDemand.dataset.listenerAttached) {
    cardDemand.addEventListener('click', () => {
      selectTreeBranch('demand');
    });
    cardDemand.dataset.listenerAttached = 'true';
  }
  
  if (cardSupply && !cardSupply.dataset.listenerAttached) {
    cardSupply.addEventListener('click', () => {
      selectTreeBranch('supply');
    });
    cardSupply.dataset.listenerAttached = 'true';
  }
  
  if (window.lucide) window.lucide.createIcons();
  hideLoader();
}

function selectTreeBranch(branch) {
  state.activeBranch = branch;
  state.activeSubcategory = null;
  state.activeSupplyClass = null;
  
  const cardDemand = document.getElementById('node-branch-demand');
  const cardSupply = document.getElementById('node-branch-supply');
  
  const hasClassification = state.classificationMap.size > 0;
  const accData = state.treeData.accounts[state.activeAccount];
  if (!accData) return;
  
  if (branch === 'demand') {
    if (cardDemand) cardDemand.classList.add('active');
    if (cardSupply) cardSupply.classList.remove('active');
    
    if (hasClassification) {
      renderDemandClassifications();
    } else {
      const rowSub = document.getElementById('tree-row-subbranches');
      const connSub = document.getElementById('conn-to-subbranches');
      if (rowSub) rowSub.classList.add('hidden');
      if (connSub) connSub.classList.add('hidden');
      
      const rowCls = document.getElementById('tree-row-classifications');
      const connCls = document.getElementById('conn-to-classifications');
      if (rowCls) rowCls.classList.add('hidden');
      if (connCls) connCls.classList.add('hidden');
      
      renderDetailsTable('demand', accData.demand);
    }
  } else if (branch === 'supply') {
    if (cardSupply) cardSupply.classList.add('active');
    if (cardDemand) cardDemand.classList.remove('active');
    
    if (hasClassification) {
      renderSupplyClassificationsDirect();
    } else {
      const rowSub = document.getElementById('tree-row-subbranches');
      const connSub = document.getElementById('conn-to-subbranches');
      if (rowSub) rowSub.classList.add('hidden');
      if (connSub) connSub.classList.add('hidden');
      
      const rowCls = document.getElementById('tree-row-classifications');
      const connCls = document.getElementById('conn-to-classifications');
      if (rowCls) rowCls.classList.add('hidden');
      if (connCls) connCls.classList.add('hidden');
      
      renderDetailsTable('supplyAgeing', accData.supply.all.filter(bp => bp.isProblem));
    }
  }
}

function renderDemandClassifications() {
  const accData = state.treeData.accounts[state.activeAccount];
  if (!accData) return;
  
  const groups = { 'A': 0, 'B': 0, 'C': 0, '#N/A': 0 };
  accData.demand.forEach(bp => {
    const cls = bp.classification;
    groups[cls] = (groups[cls] || 0) + 1;
  });
  
  const grid = document.getElementById('grid-subbranches');
  if (!grid) return;
  grid.innerHTML = '';
  
  const lblTitle = document.getElementById('lbl-subbranch-title');
  if (lblTitle) {
    lblTitle.textContent = "Demand Classifications";
  }
  
  ['A', 'B', 'C', '#N/A'].forEach(cls => {
    const count = groups[cls] || 0;
    const card = document.createElement('div');
    card.className = 'tree-node-card';
    card.dataset.classification = cls;
    
    let colorClass = 'color-blue';
    if (cls === 'A') colorClass = 'color-yellow';
    else if (cls === 'C') colorClass = 'color-secondary';
    
    card.innerHTML = `
      <div class="node-card-icon ${colorClass}"><i data-lucide="layers"></i></div>
      <div class="node-card-title">Class ${cls}</div>
      <div class="node-card-desc">Pack Classification</div>
      <div class="node-card-badge bg-blue">${count}</div>
    `;
    
    card.addEventListener('click', () => {
      selectDemandClassification(cls);
    });
    grid.appendChild(card);
  });
  
  if (window.lucide) window.lucide.createIcons();
  
  const rowSub = document.getElementById('tree-row-subbranches');
  const connSub = document.getElementById('conn-to-subbranches');
  if (rowSub) rowSub.classList.remove('hidden');
  if (connSub) connSub.classList.remove('hidden');
  
  const rowCls = document.getElementById('tree-row-classifications');
  const connCls = document.getElementById('conn-to-classifications');
  const rowDetails = document.getElementById('tree-row-details');
  const connDetails = document.getElementById('conn-to-details');
  if (rowCls) rowCls.classList.add('hidden');
  if (connCls) connCls.classList.add('hidden');
  if (rowDetails) rowDetails.classList.add('hidden');
  if (connDetails) connDetails.classList.add('hidden');
  
  setTimeout(() => {
    if (rowSub) rowSub.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }, 50);
}

function selectDemandClassification(cls) {
  state.activeSubcategory = cls;
  
  const cards = document.querySelectorAll('#grid-subbranches .tree-node-card');
  cards.forEach(card => {
    if (card.dataset.classification === cls) {
      card.classList.add('active');
    } else {
      card.classList.remove('active');
    }
  });
  
  const accData = state.treeData.accounts[state.activeAccount];
  if (!accData) return;
  const basepacks = accData.demand.filter(bp => bp.classification === cls);
  renderDetailsTable('demand', basepacks);
}

function renderSupplyClassificationsDirect() {
  const accData = state.treeData.accounts[state.activeAccount];
  if (!accData) return;
  
  const grid = document.getElementById('grid-subbranches');
  if (!grid) return;
  grid.innerHTML = '';
  
  const lblTitle = document.getElementById('lbl-subbranch-title');
  if (lblTitle) {
    lblTitle.textContent = "Supply Classifications (Problem Packs)";
  }
  
  ['A', 'B', 'C', '#N/A'].forEach(cls => {
    const count = accData.supply.classifications[cls].length;
    const card = document.createElement('div');
    card.className = 'tree-node-card';
    card.dataset.classification = cls;
    
    let colorClass = 'color-blue';
    if (cls === 'A') colorClass = 'color-yellow';
    else if (cls === 'C') colorClass = 'color-secondary';
    
    let thresholdDesc = 'PO Ageing > 7 Days';
    if (cls === 'A') thresholdDesc = 'PO Ageing > 3 Days';
    else if (cls === 'B') thresholdDesc = 'PO Ageing > 7 Days';
    else if (cls === 'C') thresholdDesc = 'PO Ageing > 11 Days';
    
    card.innerHTML = `
      <div class="node-card-icon ${colorClass}"><i data-lucide="layers"></i></div>
      <div class="node-card-title">Class ${cls}</div>
      <div class="node-card-desc">${thresholdDesc}</div>
      <div class="node-card-badge bg-orange">${count}</div>
    `;
    
    card.addEventListener('click', () => {
      selectSupplyClassificationDirect(cls);
    });
    grid.appendChild(card);
  });
  
  if (window.lucide) window.lucide.createIcons();
  
  const rowSub = document.getElementById('tree-row-subbranches');
  const connSub = document.getElementById('conn-to-subbranches');
  if (rowSub) rowSub.classList.remove('hidden');
  if (connSub) connSub.classList.remove('hidden');
  
  const rowCls = document.getElementById('tree-row-classifications');
  const connCls = document.getElementById('conn-to-classifications');
  const rowDetails = document.getElementById('tree-row-details');
  const connDetails = document.getElementById('conn-to-details');
  if (rowCls) rowCls.classList.add('hidden');
  if (connCls) connCls.classList.add('hidden');
  if (rowDetails) rowDetails.classList.add('hidden');
  if (connDetails) connDetails.classList.add('hidden');
  
  setTimeout(() => {
    if (rowSub) rowSub.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }, 50);
}

function selectSupplyClassificationDirect(cls) {
  state.activeSubcategory = cls;
  
  const cards = document.querySelectorAll('#grid-subbranches .tree-node-card');
  cards.forEach(card => {
    if (card.dataset.classification === cls) {
      card.classList.add('active');
    } else {
      card.classList.remove('active');
    }
  });
  
  const accData = state.treeData.accounts[state.activeAccount];
  if (!accData) return;
  const basepacks = accData.supply.classifications[cls];
  renderDetailsTable('supplyAgeing', basepacks);
}

function renderDetailsTable(type, basepacks) {
  const head = document.getElementById('details-table-head');
  const body = document.getElementById('details-table-body');
  if (!head || !body) return;
  
  head.innerHTML = '';
  body.innerHTML = '';
  
  const hasCls = state.classificationMap.size > 0;
  
  if (basepacks.length === 0) {
    body.innerHTML = '<tr><td colspan="100%" class="text-center text-muted" style="padding: 20px 0;">No basepacks found in this category.</td></tr>';
  } else {
    if (type === 'demand') {
      head.innerHTML = `
        <th>Basepack</th>
        <th>Description</th>
        <th>Customer Chain</th>
        <th>Brand</th>
        ${hasCls ? '<th>Classification</th>' : ''}
        <th class="text-right">Average OLA</th>
        <th>PO Status</th>
      `;
      basepacks.forEach(bp => {
        const tr = document.createElement('tr');
        const statusText = bp.mostRecentPo ? String(bp.mostRecentPo['PO Status']).trim() : 'No PO Found';
        tr.innerHTML = `
          <td class="font-semibold">${bp.basepack}</td>
          <td>${bp.desc}</td>
          <td>${bp.customerChain || '-'}</td>
          <td>${bp.brand}</td>
          ${hasCls ? `<td><span class="badge">${bp.classification}</span></td>` : ''}
          <td class="text-right font-semibold text-danger">${bp.avgOla.toFixed(1)}%</td>
          <td><span class="badge-issue red">${statusText}</span></td>
        `;
        body.appendChild(tr);
      });
    } else if (type === 'fillRate') {
      head.innerHTML = `
        <th>Basepack</th>
        <th>Description</th>
        <th>Customer Chain</th>
        <th>Brand</th>
        ${hasCls ? '<th>Classification</th>' : ''}
        <th class="text-right">Average OLA</th>
        <th class="text-right">Avg Fill Rate</th>
        <th class="text-right">Open PO Count</th>
      `;
      basepacks.forEach(bp => {
        const tr = document.createElement('tr');
        const fillRatePct = bp.avgFillRate !== null ? `${(bp.avgFillRate * 100).toFixed(1)}%` : '-';
        tr.innerHTML = `
          <td class="font-semibold">${bp.basepack}</td>
          <td>${bp.desc}</td>
          <td>${bp.customerChain || '-'}</td>
          <td>${bp.brand}</td>
          ${hasCls ? `<td><span class="badge">${bp.classification}</span></td>` : ''}
          <td class="text-right font-semibold text-warning">${bp.avgOla.toFixed(1)}%</td>
          <td class="text-right font-semibold text-danger">${fillRatePct}</td>
          <td class="text-right">${bp.openPos.length}</td>
        `;
        body.appendChild(tr);
      });
    } else if (type === 'ageing' || type === 'supplyAgeing') {
      head.innerHTML = `
        <th>Basepack</th>
        <th>Description</th>
        <th>Customer Chain</th>
        <th>Brand</th>
        ${hasCls ? '<th>Classification</th>' : ''}
        <th class="text-right">Average OLA</th>
        <th class="text-right">Avg PO Ageing</th>
        <th class="text-right">Open PO Count</th>
      `;
      basepacks.forEach(bp => {
        const tr = document.createElement('tr');
        const ageingText = bp.avgAgeing !== null ? `${bp.avgAgeing.toFixed(1)} days` : '-';
        tr.innerHTML = `
          <td class="font-semibold">${bp.basepack}</td>
          <td>${bp.desc}</td>
          <td>${bp.customerChain || '-'}</td>
          <td>${bp.brand}</td>
          ${hasCls ? `<td><span class="badge">${bp.classification}</span></td>` : ''}
          <td class="text-right font-semibold text-warning">${bp.avgOla.toFixed(1)}%</td>
          <td class="text-right font-semibold text-danger">${ageingText}</td>
          <td class="text-right">${bp.openPos.length}</td>
        `;
        body.appendChild(tr);
      });
    }
  }
  
  const rowDetails = document.getElementById('tree-row-details');
  const connDetails = document.getElementById('conn-to-details');
  if (rowDetails) rowDetails.classList.remove('hidden');
  if (connDetails) connDetails.classList.remove('hidden');
  
  setTimeout(() => {
    if (rowDetails) rowDetails.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }, 50);
}

// Generate Summary Raise PO Pivot table structure
function generateSummaryRaisePoPivot() {
  state.summaryRaisePo = [];
  
  if (state.detailedBreakdown.length === 0) return;
  
  const hasClassification = state.classificationMap.size > 0;
  if (!hasClassification) return;
  
  // Extract unique classifications
  const rowGroups = new Map(); // key: "cls" -> { classification, basepacks: Set, locations: { loc: { sum, count } } }
  const uniqueLocations = new Set();
  
  state.detailedBreakdown.forEach(row => {
    const cls = String(row['Pack classification'] || '').trim();
    if (!cls) return; // IGNORE Blank pack classification
    
    const loc = String(row.location || 'Blank').trim();
    const ola = Number(row['average ola']) || 0;
    const bp = String(row.Basepack || '').trim();
    
    uniqueLocations.add(loc);
    
    if (!rowGroups.has(cls)) {
      rowGroups.set(cls, {
        classification: cls,
        basepacks: new Set(),
        locations: {}
      });
    }
    
    const group = rowGroups.get(cls);
    if (bp) {
      group.basepacks.add(bp);
    }
    
    if (!group.locations[loc]) {
      group.locations[loc] = { sum: 0, count: 0 };
    }
    group.locations[loc].sum += ola;
    group.locations[loc].count += 1;
  });
  
  const sortedLocations = [...uniqueLocations].sort();
  
  // Construct AOA (Array of Arrays) for Excel output
  // Header Row: Pack classification, Unique Basepacks, ...locations, Average
  const header = ['Pack classification', 'Unique Basepacks', ...sortedLocations, 'Average'];
  state.summaryRaisePo.push(header);
  
  // Values Rows
  const sortedClasses = [...rowGroups.keys()].sort((a, b) => a.localeCompare(b));
  
  sortedClasses.forEach(cls => {
    const grp = rowGroups.get(cls);
    const row = [grp.classification, grp.basepacks.size];
    let sumLocationAverages = 0;
    let countLocationsWithData = 0;
    
    sortedLocations.forEach(loc => {
      const locData = grp.locations[loc];
      if (locData && locData.count > 0) {
        const avg = locData.sum / locData.count;
        const avgVal = Number(avg.toFixed(2));
        row.push(avgVal);
        sumLocationAverages += avgVal;
        countLocationsWithData++;
      } else {
        row.push(''); // Leave blank as requested
      }
    });
    
    // Average across all locations with data
    if (countLocationsWithData > 0) {
      const overallAvg = sumLocationAverages / countLocationsWithData;
      row.push(Number(overallAvg.toFixed(2)));
    } else {
      row.push('');
    }
    
    state.summaryRaisePo.push(row);
  });
}

// -------------------------------------------------------------
// EXECUTIVE DASHBOARD STATS & CHARTS
// -------------------------------------------------------------
function renderKPIs() {
  const sheetCountDetailed = document.getElementById('sheet-count-detailed');
  if (sheetCountDetailed) sheetCountDetailed.textContent = state.detailedBreakdown.length;
}

function recreateCharts() {
  // Charts have been removed from the Executive Dashboard.
}

// -------------------------------------------------------------
// INTERACTIVE DRILL-DOWN TREE TABLE
// -------------------------------------------------------------
// ANALYTICAL PROCESSING AND REPORT PREPARATION
// -------------------------------------------------------------

// -------------------------------------------------------------
// EXCEL EXPORTER
// -------------------------------------------------------------
async function downloadExcelWorkbook() {
  if (state.filteredOla.length === 0) return;
  
  showLoader("Generating Stylized Excel Workbook...");
  
  setTimeout(async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      
      // Common Styling Definitions
      const headerFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '1F4E78' } // Premium Dark Navy Blue
      };
      
      const headerFont = {
        name: 'Segoe UI',
        size: 11,
        bold: true,
        color: { argb: 'FFFFFF' }
      };
      
      const dataFont = {
        name: 'Segoe UI',
        size: 10,
        color: { argb: '000000' }
      };
      
      const thinBorder = {
        top: { style: 'thin', color: { argb: 'D3D3D3' } },
        left: { style: 'thin', color: { argb: 'D3D3D3' } },
        bottom: { style: 'thin', color: { argb: 'D3D3D3' } },
        right: { style: 'thin', color: { argb: 'D3D3D3' } }
      };
      
      const zebraFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F6FA' } // Very Light Blue Tint
      };

      // -------------------------------------------------------------
      // SHEET 1: Detailed breakdown
      // -------------------------------------------------------------
      const ws1 = workbook.addWorksheet("Detailed breakdown");
      ws1.views = [{ showGridLines: true }];
      
      const headers1 = [
        'Account', 'Basepack', 'basepack_desc', 'Customer Chain', 'small_c', 'brand', 'location', 
        'Pack classification', 'pack_tag', 'average ola'
      ];
      
      ws1.addRow(headers1);
      state.detailedBreakdown.forEach(row => {
        ws1.addRow([
          row['Customer Chain'], row['Basepack'], row['basepack_desc'], row['Customer Chain'], row['small_c'], row['brand'], row['location'],
          row['Pack classification'], row['pack_tag'], row['average ola']
        ]);
      });
      
      styleWorksheet(ws1, headers1.length, false);

      // -------------------------------------------------------------
      // SHEET 2: Summary (Only if classification mapping uploaded)
      // -------------------------------------------------------------
      const hasCls = state.classificationMap.size > 0;
      if (hasCls && state.summaryRaisePo.length > 0) {
        const ws2 = workbook.addWorksheet("Summary");
        ws2.views = [{ showGridLines: true }];
        
        state.summaryRaisePo.forEach(row => {
          ws2.addRow(row);
        });
        
        const summaryHeaders = state.summaryRaisePo[0];
        styleWorksheet(ws2, summaryHeaders.length, true);
      }

      // Helper function to style worksheet
      function styleWorksheet(ws, colCount, isSummary = false) {
        // Style Header Row
        const headerRow = ws.getRow(1);
        headerRow.height = 26;
        for (let c = 1; c <= colCount; c++) {
          const cell = headerRow.getCell(c);
          cell.fill = headerFill;
          cell.font = headerFont;
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.border = thinBorder;
        }
        
        // Style Data Rows
        const rowCount = ws.rowCount;
        for (let r = 2; r <= rowCount; r++) {
          const row = ws.getRow(r);
          row.height = 20;
          const isEven = (r % 2 === 0);
          
          for (let c = 1; c <= colCount; c++) {
            const cell = row.getCell(c);
            cell.font = dataFont;
            cell.border = thinBorder;
            if (isEven) {
              cell.fill = zebraFill;
            }
            
            const colHeader = String(ws.getRow(1).getCell(c).value || '');
            const val = cell.value;
            const colHeaderClean = colHeader.trim().toLowerCase();
            
            // Handle IDs: Basepack and PO Number should be strings and not formatted with commas
            if (colHeaderClean === 'basepack' || colHeaderClean === 'po number') {
              cell.numFmt = '@'; // Force text format
              cell.alignment = { vertical: 'middle', horizontal: 'left' };
              if (cell.value !== null && cell.value !== undefined) {
                cell.value = String(cell.value);
              }
            } else if (isSummary) {
              if (c === 1) {
                // Pack classification
                cell.alignment = { vertical: 'middle', horizontal: 'left' };
              } else if (c === 2) {
                // Unique Basepacks
                cell.alignment = { vertical: 'middle', horizontal: 'right' };
                cell.numFmt = '#,##0';
              } else {
                // locations (averages)
                cell.alignment = { vertical: 'middle', horizontal: 'right' };
                if (typeof val === 'number') {
                  cell.numFmt = '#,##0.0';
                }
              }
            } else {
              // Standard Sheets
              if (typeof val === 'number') {
                cell.alignment = { vertical: 'middle', horizontal: 'right' };
                if (colHeaderClean.includes('ola') || colHeaderClean.includes('value')) {
                  cell.numFmt = '#,##0.0';
                } else if (colHeaderClean.includes('rate') || colHeaderClean.includes('percentage')) {
                  cell.numFmt = '0.00';
                } else {
                  cell.numFmt = '#,##0';
                }
              } else if (val instanceof Date || colHeaderClean.includes('date')) {
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
              } else if (colHeaderClean.includes('number') || colHeaderClean.includes('code') || colHeaderClean.includes('week') || colHeaderClean.includes('pincode')) {
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
              } else {
                cell.alignment = { vertical: 'middle', horizontal: 'left' };
              }
            }
          }
        }
        
        // Auto-fit Columns
        for (let c = 1; c <= colCount; c++) {
          let maxLength = 10;
          ws.getColumn(c).eachCell({ includeEmpty: true }, (cell) => {
            const cellVal = cell.value;
            if (cellVal !== undefined && cellVal !== null) {
              let strLen = 0;
              if (cellVal instanceof Date) {
                strLen = 10;
              } else if (typeof cellVal === 'number' && cell.numFmt === '0.0"%"') {
                strLen = cellVal.toFixed(1).length + 1;
              } else if (typeof cellVal === 'number') {
                strLen = cellVal.toFixed(1).length;
              } else {
                strLen = String(cellVal).length;
              }
              if (strLen > maxLength) {
                maxLength = strLen;
              }
            }
          });
          ws.getColumn(c).width = Math.min(maxLength + 4, 30);
        }
      }

      // Generate buffer and trigger download
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const filename = `DEMAND_RAISE_PO.xlsx`;
      
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = filename;
      link.click();
      
      hideLoader();
    } catch (err) {
      hideLoader();
      alert(`Error exporting styled Excel workbook: ${err.message}`);
      console.error(err);
    }
  }, 100);
}

async function downloadSupplyExcelWorkbook() {
  const problemPacks = [];
  const hasClassification = state.classificationMap.size > 0;
  if (hasClassification) {
    ['A', 'B', 'C', '#N/A'].forEach(cls => {
      const list = state.treeData.supply.classifications[cls] || [];
      problemPacks.push(...list);
    });
  } else {
    problemPacks.push(...state.treeData.supply.all.filter(bp => bp.isProblem));
  }
  
  if (problemPacks.length === 0) {
    alert("No supply side issue basepacks found to download.");
    return;
  }
  
  // Sort problem packs by basepack code numerically/alphabetically
  problemPacks.sort((a, b) => String(a.basepack).localeCompare(String(b.basepack), undefined, {numeric: true}));
  
  showLoader("Generating Stylized Supply Excel Workbook...");
  
  setTimeout(async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      
      // Common Styling Definitions
      const headerFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '1F4E78' } // Premium Dark Navy Blue
      };
      
      const headerFont = {
        name: 'Segoe UI',
        size: 11,
        bold: true,
        color: { argb: 'FFFFFF' }
      };
      
      const dataFont = {
        name: 'Segoe UI',
        size: 10,
        color: { argb: '000000' }
      };
      
      const thinBorder = {
        top: { style: 'thin', color: { argb: 'D3D3D3' } },
        left: { style: 'thin', color: { argb: 'D3D3D3' } },
        bottom: { style: 'thin', color: { argb: 'D3D3D3' } },
        right: { style: 'thin', color: { argb: 'D3D3D3' } }
      };
      
      const zebraFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F6FA' } // Very Light Blue Tint
      };
 
      const ws = workbook.addWorksheet("Supply Ageing Issues");
      ws.views = [{ showGridLines: true }];
      
      const headers = [
        'Account', 'basepack', 'basepack_desc', 'Customer Chain', 'small_c', 'average ola 2 weeks', 'po date', 'Po ageing', 'po number', 'po ageing bucket'
      ];
      
      ws.addRow(headers);
      
      problemPacks.forEach(bpObj => {
        const openPos = bpObj.openPos || [];
        const cls = bpObj.classification;
        
        let threshold = 7;
        if (hasClassification) {
          if (cls === 'A') threshold = 3;
          else if (cls === 'B') threshold = 7;
          else if (cls === 'C') threshold = 11;
        }
        
        openPos.forEach(po => {
          const poAgeing = po['PO Ageing'] !== undefined && po['PO Ageing'] !== null && po['PO Ageing'] !== '' ? Number(po['PO Ageing']) : 0;
          if (poAgeing > threshold) {
            const poDateObj = parseExcelDate(po['PO Date']);
            const chainName = bpObj.customerChain || String(po['Customer Chain'] || '');
            ws.addRow([
              chainName,
              String(bpObj.basepack),
              bpObj.desc,
              chainName,
              bpObj.small_c || '',
              bpObj.avgOla,
              poDateObj ? formatDate(poDateObj) : '',
              poAgeing,
              String(po['PO Number'] || ''),
              po['Po Ageing Bucket'] || ''
            ]);
          }
        });
      });
      
      // Style Header Row
      const headerRow = ws.getRow(1);
      headerRow.height = 26;
      for (let c = 1; c <= headers.length; c++) {
        const cell = headerRow.getCell(c);
        cell.fill = headerFill;
        cell.font = headerFont;
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = thinBorder;
      }
      
      // Style Data Rows
      const rowCount = ws.rowCount;
      for (let r = 2; r <= rowCount; r++) {
        const row = ws.getRow(r);
        row.height = 20;
        const isEven = (r % 2 === 0);
        
        for (let c = 1; c <= headers.length; c++) {
          const cell = row.getCell(c);
          cell.font = dataFont;
          cell.border = thinBorder;
          if (isEven) {
            cell.fill = zebraFill;
          }
          
          const colHeader = headers[c - 1];
          const val = cell.value;
          const colHeaderClean = colHeader.trim().toLowerCase();
          
          if (colHeaderClean === 'basepack' || colHeaderClean === 'po number') {
            cell.numFmt = '@'; // Force text format
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
            if (cell.value !== null && cell.value !== undefined) {
              cell.value = String(cell.value);
            }
          } else if (colHeaderClean.includes('ola')) {
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
            if (typeof val === 'number') {
              cell.numFmt = '#,##0.0';
            }
          } else if (colHeaderClean.includes('ageing') && !colHeaderClean.includes('bucket')) {
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
            if (typeof val === 'number') {
              cell.numFmt = '#,##0';
            }
          } else if (colHeaderClean.includes('bucket')) {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
          } else if (colHeaderClean.includes('date')) {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
          } else {
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
          }
        }
      }
      
      // Auto-fit Columns
      for (let c = 1; c <= headers.length; c++) {
        let maxLength = 10;
        ws.getColumn(c).eachCell({ includeEmpty: true }, (cell) => {
          const cellVal = cell.value;
          if (cellVal !== undefined && cellVal !== null) {
            let strLen = 0;
            if (typeof cellVal === 'number') {
              strLen = cellVal.toFixed(1).length;
            } else {
              strLen = String(cellVal).length;
            }
            if (strLen > maxLength) {
              maxLength = strLen;
            }
          }
        });
        ws.getColumn(c).width = Math.min(maxLength + 4, 30);
      }

      // Generate buffer and trigger download
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const filename = `SUPPLY_AGEING.xlsx`;
      
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = filename;
      link.click();
      
      hideLoader();
    } catch (err) {
      hideLoader();
      alert(`Error exporting styled Supply Excel workbook: ${err.message}`);
      console.error(err);
    }
  }, 100);
}

// -------------------------------------------------------------
// UTILITIES (DATES, DATA TYPES)
// -------------------------------------------------------------
function parseExcelDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    // SheetJS date serial number offset
    // 25569 = Jan 1 1970. 86400 * 1000 ms in a day
    const utcDate = new Date((val - 25569) * 86400 * 1000);
    // Correct for local time offset
    const localDate = new Date(utcDate.getTime() + utcDate.getTimezoneOffset() * 60 * 1000);
    return localDate;
  }
  
  // Try custom string date formats commonly found in excel: YYYY-MM-DD, DD/MM/YYYY, etc
  const strVal = String(val).trim();
  
  // Matches dd-mmm-yy or dd-mmm-yyyy (e.g. 26-May-26 or 26-May-2026)
  const mmmMatch = strVal.match(/^(\d{1,2})[\/\-]([a-zA-Z]{3,9})[\/\-](\d{2,4})/);
  if (mmmMatch) {
    const day = parseInt(mmmMatch[1], 10);
    const mStr = mmmMatch[2].toLowerCase().substring(0, 3);
    const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
    const month = months.indexOf(mStr);
    let year = parseInt(mmmMatch[3], 10);
    if (mmmMatch[3].length === 2) {
      year += 2000;
    }
    if (month >= 0) {
      return new Date(year, month, day);
    }
  }
  
  // Matches dd-mm-yyyy or dd/mm/yyyy
  const dmyMatch = strVal.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10) - 1; // 0-indexed
    const year = parseInt(dmyMatch[3], 10);
    return new Date(year, month, day);
  }
  
  // Matches yyyy-mm-dd
  const ymdMatch = strVal.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (ymdMatch) {
    const year = parseInt(ymdMatch[1], 10);
    const month = parseInt(ymdMatch[2], 10) - 1;
    const day = parseInt(ymdMatch[3], 10);
    return new Date(year, month, day);
  }
  
  const d = new Date(strVal);
  if (!isNaN(d.getTime())) return d;
  
  return null;
}

function formatDate(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) return '';
  const day = String(date.getDate()).padStart(2, '0');
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const month = months[date.getMonth()];
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}

function parseNumber(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  const cleanStr = String(val).replace(/[$,]/g, '').trim();
  const num = parseFloat(cleanStr);
  return isNaN(num) ? 0 : num;
}

function getFillRate(po) {
  // Check mapped key first (robustly set by REQUIRED_COLUMNS.allocation and fuzzyMatchKey)
  if (po['Fill Rate'] !== undefined && po['Fill Rate'] !== null && po['Fill Rate'] !== '') {
    return parsePercentOrDecimal(po['Fill Rate']);
  }
  
  const keys = Object.keys(po);
  const fillRateKey = keys.find(k => {
    const cleanK = k.toLowerCase().replace(/[^a-z0-9%]/g, '');
    return cleanK === 'fillrate' || cleanK === 'fillrate%' || cleanK === 'fr%' || cleanK === 'fr';
  });
  
  if (fillRateKey && po[fillRateKey] !== undefined && po[fillRateKey] !== null && po[fillRateKey] !== '') {
    return parsePercentOrDecimal(po[fillRateKey]);
  }
  
  const orderVal = parseNumber(po['Order Value Lacs']);
  const invoicedVal = parseNumber(po['Invoiced Value Lacs']);
  if (orderVal > 0) {
    return invoicedVal / orderVal;
  }
  
  return 1.0;
}

function parsePercentOrDecimal(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') {
    return val > 2 ? val / 100 : val;
  }
  const str = String(val).replace(/[$,]/g, '').trim();
  if (str.endsWith('%')) {
    const num = parseFloat(str.replace('%', ''));
    return isNaN(num) ? 0 : num / 100;
  }
  const num = parseFloat(str);
  if (isNaN(num)) return 0;
  return num > 2 ? num / 100 : num;
}

// -------------------------------------------------------------
// CUSTOM CUSTOMER CHAIN MULTI-SELECT FUNCTIONS
// -------------------------------------------------------------
function getCanonicalChainName(name) {
  const clean = String(name || '').trim().toLowerCase();
  if (clean === 'purple' || clean === 'purplle') {
    return 'Purplle';
  }
  if (!clean) return '';
  return clean.charAt(0).toUpperCase() + clean.slice(1);
}

function populateCustomerChainsDropdown() {
  const chainsSet = new Set();
  state.allocationData.forEach(row => {
    const chain = String(row['Customer Chain'] || '').trim();
    if (chain) {
      chainsSet.add(getCanonicalChainName(chain));
    }
  });
  
  const sortedChains = [...chainsSet].sort((a, b) => a.localeCompare(b));
  
  const listContainer = document.getElementById('chain-options-list');
  const selectTrigger = document.getElementById('chain-select-trigger');
  
  if (!listContainer || !selectTrigger) return;
  
  listContainer.innerHTML = '';
  
  const selectTriggerText = selectTrigger.querySelector('span');
  
  if (sortedChains.length === 0) {
    if (selectTriggerText) selectTriggerText.textContent = 'No Customer Chains Found';
    selectTrigger.setAttribute('disabled', 'true');
    return;
  }
  
  // Enable dropdown trigger
  selectTrigger.removeAttribute('disabled');
  
  // Update state
  state.availableChains = sortedChains;
  state.selectedChains = [...sortedChains]; // default select all
  
  updateChainTriggerText();
  
  // Add option checkboxes
  sortedChains.forEach(chain => {
    const optionDiv = document.createElement('div');
    optionDiv.className = 'custom-select-option';
    optionDiv.dataset.value = chain;
    
    const chk = document.createElement('input');
    chk.type = 'checkbox';
    chk.id = `chk-chain-${chain.replace(/\s+/g, '-')}`;
    chk.checked = true;
    
    const lbl = document.createElement('label');
    lbl.htmlFor = chk.id;
    lbl.textContent = chain;
    
    optionDiv.appendChild(chk);
    optionDiv.appendChild(lbl);
    
    // Checkbox change listener
    chk.addEventListener('change', (e) => {
      e.stopPropagation();
      if (chk.checked) {
        if (!state.selectedChains.includes(chain)) {
          state.selectedChains.push(chain);
        }
      } else {
        state.selectedChains = state.selectedChains.filter(c => c !== chain);
      }
      
      // Update Select All state
      const chkAll = document.getElementById('chk-chain-all');
      if (chkAll) {
        chkAll.checked = (state.selectedChains.length === state.availableChains.length);
      }
      
      updateChainTriggerText();
    });
    
    // Div click checks/unchecks
    optionDiv.addEventListener('click', (e) => {
      if (e.target !== chk && e.target !== lbl) {
        chk.checked = !chk.checked;
        chk.dispatchEvent(new Event('change'));
      }
    });
    
    listContainer.appendChild(optionDiv);
  });
  
  // Setup Select All checkbox behavior
  const chkAll = document.getElementById('chk-chain-all');
  if (chkAll) {
    chkAll.checked = true;
    
    // Re-attach select all event listener
    const newChkAll = chkAll.cloneNode(true);
    chkAll.parentNode.replaceChild(newChkAll, chkAll);
    
    newChkAll.addEventListener('change', (e) => {
      e.stopPropagation();
      const checked = newChkAll.checked;
      if (checked) {
        state.selectedChains = [...state.availableChains];
      } else {
        state.selectedChains = [];
      }
      
      // Check/uncheck all option checkboxes
      state.availableChains.forEach(chain => {
        const input = document.getElementById(`chk-chain-${chain.replace(/\s+/g, '-')}`);
        if (input) input.checked = checked;
      });
      
      updateChainTriggerText();
    });
    
    const selectAllDiv = document.querySelector('.select-all-option');
    if (selectAllDiv) {
      const newSelectAllDiv = selectAllDiv.cloneNode(true);
      selectAllDiv.parentNode.replaceChild(newSelectAllDiv, selectAllDiv);
      
      const chkAllReplaced = document.getElementById('chk-chain-all');
      newSelectAllDiv.addEventListener('click', (e) => {
        if (e.target !== chkAllReplaced && e.target !== newSelectAllDiv.querySelector('label')) {
          chkAllReplaced.checked = !chkAllReplaced.checked;
          chkAllReplaced.dispatchEvent(new Event('change'));
        }
      });
    }
  }
}

function updateChainTriggerText() {
  const selectTrigger = document.getElementById('chain-select-trigger');
  if (!selectTrigger) return;
  const selectTriggerText = selectTrigger.querySelector('span');
  if (!selectTriggerText) return;
  
  const count = state.selectedChains.length;
  const total = state.availableChains.length;
  
  if (count === total) {
    selectTriggerText.textContent = `Customer Chain: All (${total})`;
  } else if (count === 0) {
    selectTriggerText.textContent = `Customer Chain: None Selected`;
  } else if (count <= 2) {
    selectTriggerText.textContent = `Customer Chain: ${state.selectedChains.join(', ')}`;
  } else {
    selectTriggerText.textContent = `Customer Chain: ${count} of ${total} Selected`;
  }
  
  // Disable process button if no chains are selected
  if (count === 0) {
    btnProcess.setAttribute('disabled', 'true');
  } else if (state.olaData.length > 0 && state.allocationData.length > 0) {
    btnProcess.removeAttribute('disabled');
  }
}

// -------------------------------------------------------------
// EXCEL CLEANER SCRIPT COPILOT FUNCTIONS
// -------------------------------------------------------------
const DEFAULT_CLEANER_SCRIPT = `function main(workbook: ExcelScript.Workbook) {
	// Deletes all unnecessary sheets to shrink workbook file size
	let basepack_Depot_ATP_View = workbook.getWorksheet("Basepack-Depot-ATP View");
	if (basepack_Depot_ATP_View) basepack_Depot_ATP_View.delete();
	
	let pivot_Use = workbook.getWorksheet("Pivot Use");
	if (pivot_Use) pivot_Use.delete();
	
	let sheet2 = workbook.getWorksheet("Sheet2");
	if (sheet2) sheet2.delete();
	
	let new_base_use_advance = workbook.getWorksheet("New base use advance");
	if (new_base_use_advance) new_base_use_advance.delete();
	
	let new_Base_Use = workbook.getWorksheet("New Base Use");
	if (new_Base_Use) new_Base_Use.delete();
	
	let new_Base_Pivot = workbook.getWorksheet("New Base Pivot");
	if (new_Base_Pivot) new_Base_Pivot.delete();
	
	let quota_Loss_Summary = workbook.getWorksheet("Quota Loss Summary");
	if (quota_Loss_Summary) quota_Loss_Summary.delete();
	
	let sheet3 = workbook.getWorksheet("Sheet3");
	if (sheet3) sheet3.delete();
	
	let quota_Loss_Working = workbook.getWorksheet("Quota Loss Working");
	if (quota_Loss_Working) quota_Loss_Working.delete();
	
	let quota_Loss_Pivot = workbook.getWorksheet("Quota Loss Pivot");
	if (quota_Loss_Pivot) quota_Loss_Pivot.delete();
	
	let sheet1 = workbook.getWorksheet("Sheet1");
	if (sheet1) sheet1.delete();
	
	let summary_Pivot = workbook.getWorksheet("Summary Pivot");
	if (summary_Pivot) summary_Pivot.delete();
	
	let use_SI = workbook.getWorksheet("Use SI");
	if (use_SI) use_SI.delete();
	
	let summary = workbook.getWorksheet("Summary");
	if (summary) summary.delete();
	
	let aTP = workbook.getWorksheet("ATP");
	if (aTP) aTP.delete();
	
	let winter_Tagging_Summary = workbook.getWorksheet("Winter Tagging Summary");
	if (winter_Tagging_Summary) winter_Tagging_Summary.delete();
	
	// Filters base data sheet to target channels (Myntra, Nykaa, Purplle) and exports them
	let base_File = workbook.getWorksheet("Base File");
	if (base_File) {
		// Toggle auto filter on base_File
		base_File.getAutoFilter().apply(base_File.getRange("B1"));
		// Apply values filter on base_File (keeps only target chains)
		base_File.getAutoFilter().apply(base_File.getAutoFilter().getRange(), 2, { 
			filterOn: ExcelScript.FilterOn.values, 
			values: ["MYNTRA", "NYKAA", "PURPLE"] 
		});
		
		// Add a new clean worksheet
		let sheet1_1 = workbook.addWorksheet();
		
		// Copies only filtered rows (automatically handles any row length dynamically)
		const sourceRange = base_File.getUsedRange();
		if (sourceRange) {
			sheet1_1.getRange("A1").copyFrom(sourceRange, ExcelScript.RangeCopyType.all, false, false);
		}
		
		// Delete the original bloated worksheet
		base_File.delete();
		
		// Rename the clean sheet to "Base File"
		sheet1_1.setName("Base File");
	}
}`;

function initExcelCleanerScript() {
  const btnShow = document.getElementById('btn-show-script');
  const btnClose = document.getElementById('btn-close-script');
  const btnCopy = document.getElementById('btn-copy-script');
  const btnEdit = document.getElementById('btn-edit-script');
  const btnCancel = document.getElementById('btn-cancel-edit');
  const btnSave = document.getElementById('btn-save-script');
  
  const panel = document.getElementById('script-panel');
  const viewMode = document.getElementById('script-view-mode');
  const editMode = document.getElementById('script-edit-mode');
  
  const displayPre = document.getElementById('script-code-display');
  const textarea = document.getElementById('script-textarea');
  
  if (!panel || !displayPre || !textarea) return;
  
  // Load script from localStorage or fallback to default
  let script = localStorage.getItem('ola_allocation_cleaner_script');
  
  // Force update to the new version if it was the old default script or doesn't exist
  if (!script || script.includes("requiredCols") || script.includes("totalCols")) {
    script = DEFAULT_CLEANER_SCRIPT;
    localStorage.setItem('ola_allocation_cleaner_script', DEFAULT_CLEANER_SCRIPT);
  }
  
  displayPre.textContent = script;
  textarea.value = script;
  
  // Show/Hide Panel
  if (btnShow) {
    btnShow.addEventListener('click', () => {
      panel.classList.toggle('hidden');
      if (!panel.classList.contains('hidden')) {
        panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    });
  }
  
  if (btnClose) {
    btnClose.addEventListener('click', () => {
      panel.classList.add('hidden');
    });
  }
  
  // Copy to Clipboard
  if (btnCopy) {
    btnCopy.addEventListener('click', () => {
      navigator.clipboard.writeText(script).then(() => {
        const span = btnCopy.querySelector('span');
        const icon = btnCopy.querySelector('i');
        
        if (span) span.textContent = 'Copied!';
        if (btnCopy) btnCopy.style.background = 'var(--color-success)';
        
        setTimeout(() => {
          if (span) span.textContent = 'Copy Cleaner';
          if (btnCopy) btnCopy.style.background = '';
        }, 2000);
      }).catch(err => {
        alert('Failed to copy text: ' + err);
      });
    });
  }
  
  // Toggle Edit Mode
  if (btnEdit) {
    btnEdit.addEventListener('click', () => {
      viewMode.classList.add('hidden');
      editMode.classList.remove('hidden');
    });
  }
  
  if (btnCancel) {
    btnCancel.addEventListener('click', () => {
      textarea.value = script;
      editMode.classList.add('hidden');
      viewMode.classList.remove('hidden');
    });
  }
  
  // Save Script
  if (btnSave) {
    btnSave.addEventListener('click', () => {
      const newScript = textarea.value;
      script = newScript;
      localStorage.setItem('ola_allocation_cleaner_script', newScript);
      displayPre.textContent = newScript;
      
      editMode.classList.add('hidden');
      viewMode.classList.remove('hidden');
    });
  }
}

// -------------------------------------------------------------
// ATLAN OLA QUERY COPILOT FUNCTIONS
// -------------------------------------------------------------
const DEFAULT_ATLAN_QUERY = `SELECT
    "account",
    "basepack",
    "basepack_desc",
    "business_unit",
    "small_c",
    "sales_category",
    "brand",
    "master_segment",
    "location",
    "pincode",
    "week",
    "month_year",
    "depot",
    "pack_tag",
    AVG("mrp") AS "MRP",
    AVG("ola_percentage_final") AS "Average OLA"
FROM "ecom_dailyola_allaccts_latest2months"
WHERE 
    "account" IN ('PURPLLE', 'MYNTRA', 'NYKAA')
    
    -- ✅ Only for PURPLLE: consideration and pack_tag filters
    AND (
        ("account" = 'PURPLLE' AND "consideration" = 'TRUE' AND "pack_tag" IN ('TOP PACK','BHP','LAUNCH'))
        OR ("account" IN ('MYNTRA', 'NYKAA'))
    )
    
    -- ✅ Current month + last month filter (applies to all accounts)
    AND date_parse(replace("month_year", ' ', ''), '%b-%Y') >= date_trunc('month', current_date - INTERVAL '1' MONTH)
    AND date_parse(replace("month_year", ' ', ''), '%b-%Y') < date_trunc('month', current_date + INTERVAL '1' MONTH)

GROUP BY  
    "account",
    "basepack",
    "basepack_desc",
    "business_unit",
    "small_c",
    "sales_category",
    "brand",
    "master_segment",
    "location",
    "pincode",
    "week",
    "month_year",
    "depot",
    "pack_tag"

HAVING 
    AVG("ola_percentage_final") < 80

ORDER BY 
    "month_year" DESC,
    "depot" DESC,
    "pincode" DESC,
    "location" DESC,
    "master_segment" DESC,
    "brand" DESC,
    "sales_category" DESC,
    "small_c" DESC,
    "business_unit" DESC,
    "basepack_desc" DESC,
    "basepack" DESC;`;

function initAtlanQueryScript() {
  const btnShow = document.getElementById('btn-show-atlan');
  const btnClose = document.getElementById('btn-close-atlan');
  const btnCopy = document.getElementById('btn-copy-atlan');
  const btnEdit = document.getElementById('btn-edit-atlan');
  const btnCancel = document.getElementById('btn-cancel-edit-atlan');
  const btnSave = document.getElementById('btn-save-atlan');
  
  const panel = document.getElementById('atlan-panel');
  const viewMode = document.getElementById('atlan-view-mode');
  const editMode = document.getElementById('atlan-edit-mode');
  
  const displayPre = document.getElementById('atlan-code-display');
  const textarea = document.getElementById('atlan-textarea');
  
  if (!panel || !displayPre || !textarea) return;
  
  // Load query from localStorage or fallback to default
  let query = localStorage.getItem('ola_atlan_query');
  if (!query || !query.includes("ecom_dailyola_allaccts_latest2months")) {
    query = DEFAULT_ATLAN_QUERY;
    localStorage.setItem('ola_atlan_query', DEFAULT_ATLAN_QUERY);
  }
  
  displayPre.textContent = query;
  textarea.value = query;
  
  // Show/Hide Panel
  if (btnShow) {
    btnShow.addEventListener('click', () => {
      panel.classList.toggle('hidden');
      if (!panel.classList.contains('hidden')) {
        panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    });
  }
  
  if (btnClose) {
    btnClose.addEventListener('click', () => {
      panel.classList.add('hidden');
    });
  }
  
  // Copy to Clipboard
  if (btnCopy) {
    btnCopy.addEventListener('click', () => {
      navigator.clipboard.writeText(query).then(() => {
        const span = btnCopy.querySelector('span');
        
        if (span) span.textContent = 'Copied!';
        if (btnCopy) btnCopy.style.background = 'var(--color-success)';
        
        setTimeout(() => {
          if (span) span.textContent = 'Copy Query';
          if (btnCopy) btnCopy.style.background = '';
        }, 2000);
      }).catch(err => {
        alert('Failed to copy text: ' + err);
      });
    });
  }
  
  // Toggle Edit Mode
  if (btnEdit) {
    btnEdit.addEventListener('click', () => {
      viewMode.classList.add('hidden');
      editMode.classList.remove('hidden');
    });
  }
  
  if (btnCancel) {
    btnCancel.addEventListener('click', () => {
      textarea.value = query;
      editMode.classList.add('hidden');
      viewMode.classList.remove('hidden');
    });
  }
  
  // Save Query
  if (btnSave) {
    btnSave.addEventListener('click', () => {
      const newQuery = textarea.value;
      query = newQuery;
      localStorage.setItem('ola_atlan_query', newQuery);
      displayPre.textContent = newQuery;
      
      editMode.classList.add('hidden');
      viewMode.classList.remove('hidden');
    });
  }
}



