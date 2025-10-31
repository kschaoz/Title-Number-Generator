// --- Constants ---
// We will load the SheetJS library dynamically
const SHEET_JS_URL = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";

const MAX_POINTS = 100;
const MIN_POINTS = 2;
const CONFIDENCE_THRESHOLD = 90;

// --- Global State ---
let allUploadedData = []; // This will hold all data from the Excel file
let normalizedTypeMap = new Map(); // Stores normalized_type -> Original_Type

/**
 * Helper function to dynamically load a script
 */
function loadScript(url, callback) {
  const script = document.createElement('script');
  script.src = url;
  script.onload = callback; // Run the callback function once the script is loaded
  script.onerror = () => {
    // Show a user-friendly error if the script fails to load
    const errorContainer = document.getElementById("error-container");
    const errorText = document.getElementById("error-text");
    if(errorContainer && errorText) {
        errorText.innerHTML = "<b>Critical Error:</b> Failed to load Excel processing library. Please check your internet connection (or disable ad-blockers) and refresh.";
        errorContainer.classList.remove("hidden");
    }
  };
  document.body.appendChild(script);
}

/**
 * This function is the callback that runs *after* SheetJS is loaded
 */
function onSheetJsLoaded() {
  if (!window.XLSX) {
    console.error("SheetJS library object (window.XLSX) not found.");
    return;
  }
  // Now that the library is loaded, initialize the main application logic
  initializeApp();
}

/**
 * Main application logic
 */
function initializeApp() {
  
  // --- Get DOM Elements ---
  const form = document.getElementById("calculator-form");
  const dataPointContainer = document.getElementById("data-point-container");
  const addPointBtn = document.getElementById("add-point-btn");
  const targetHouse_el = document.getElementById("target-house");
  
  const resultContainer = document.getElementById("result-container");
  const resultText = document.getElementById("result-text");
  const formulaText = document.getElementById("formula-text");
  const resultTitleType = document.getElementById("result-title-type");
  
  const analysisLevel = document.getElementById("analysis-level");
  const analysisText = document.getElementById("analysis-text");
  
  const errorContainer = document.getElementById("error-container");
  const errorText = document.getElementById("error-text");
  
  // Excel Upload Elements
  const excelUpload = document.getElementById("excel-upload");
  const uploadProcessingMsg = document.getElementById("upload-processing-msg");
  const dropZone = document.querySelector(".upload-drop-zone");

  // In-Page Type Selector Elements
  const typeSelectorContainer = document.getElementById("type-selector-container");
  const typeSelectorCheckboxes = document.getElementById("type-selector-checkboxes");
  
  // --- 1. Excel Upload Logic ---
  
  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });
  dropZone.addEventListener("dragleave", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
  });
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    const files = e.dataTransfer.files;
    if (files.length > 0) handleFile(files[0]);
  });
  excelUpload.addEventListener("change", (e) => {
    const files = e.target.files;
    if (files.length > 0) handleFile(files[0]);
  });
  
  function handleFile(file) {
    const validTypes = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel"
    ];
    if (!validTypes.includes(file.type) && !file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      showError("Invalid file type. Please upload an .xlsx or .xls file.");
      return;
    }
    
    uploadProcessingMsg.classList.remove("hidden");
    errorContainer.classList.add("hidden");
    typeSelectorContainer.classList.add("hidden");
    clearDataPoints();
    
    const reader = new FileReader();
    
    reader.onload = (event) => {
      let workbook, worksheet, excelData;
      try {
        const data = event.target.result;
        workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        worksheet = workbook.Sheets[sheetName];
        excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      } catch (err) {
        console.error("Error reading Excel structure:", err);
        showError("Could not read the Excel file. It may be corrupt or an invalid format.");
        uploadProcessingMsg.classList.add("hidden");
        excelUpload.value = null;
        return;
      }

      try {
        processExcelData(excelData);
      } catch (err) {
        console.error("Error processing Excel data:", err);
        showError(`File read, but data processing failed: ${err.message}. Please check your template.`);
      } finally {
        uploadProcessingMsg.classList.add("hidden"); 
        excelUpload.value = null; 
      }
    };
    
    reader.onerror = () => {
      showError("Error reading file.");
      uploadProcessingMsg.classList.add("hidden");
    };
    
    reader.readAsArrayBuffer(file);
  }
  
  /**
   * Parses the Excel data and populates the global data array
   */
  function processExcelData(data) {
    if (data.length < 26) {
      showError("Invalid Excel template. The file must have at least 26 rows.");
      return;
    }
    
    const houseLine = data[10] || []; // Row 11
    const titleLine = data[25] || []; // Row 26
    
    allUploadedData = []; 
    normalizedTypeMap.clear(); // Clear the old map
    
    for (let i = 3; i < houseLine.length; i++) {
      try {
        const houseCell = String(houseLine[i] || "");
        const titleCell = String(titleLine[i] || "");
        
        if (houseCell && titleCell) {
          const houseNum = parseHouseNumber(houseCell);
          const indication = parseTitleIndication(titleCell); 
          const titleNum = parseTitleNumber(titleCell);
          
          if (houseNum && titleNum) {
            allUploadedData.push({ house: houseNum, title: titleNum, type: indication });
            
            // --- NEW Normalized Map Logic ---
            if (indication) { 
              const normalized = normalizeType(indication);
              if (!normalizedTypeMap.has(normalized)) {
                normalizedTypeMap.set(normalized, indication); // Store original case
              }
            }
          }
        }
      } catch (colErr) {
        console.warn(`Skipped a column (index ${i}) due to a processing error:`, colErr.message);
      }
    }
    
    // Always populate the data points
    populateDataPoints(allUploadedData);
    
    // Only show the selector if there are types
    if (normalizedTypeMap.size > 0) {
      populateTypeSelector();
    }
  }
  
  /**
   * *** UPDATED ***
   * Populates the checkbox filter UI
   */
  function populateTypeSelector() {
    typeSelectorCheckboxes.innerHTML = ''; // Clear old checkboxes
    
    // 1. Create the "All (Manual Remove)" button
    const allDiv = document.createElement('div');
    allDiv.className = 'checkbox-filter-group all-filter';
    allDiv.innerHTML = `
      <input type="checkbox" id="filter-all" value="ALL" checked>
      <label for="filter-all">All (Manual Remove)</label>
    `;
    typeSelectorCheckboxes.appendChild(allDiv);
    
    // 2. Create buttons for each type
    for (const [normalized, original] of normalizedTypeMap.entries()) {
      const div = document.createElement('div');
      div.className = 'checkbox-filter-group';
      // Use original for label, normalized for id/value
      div.innerHTML = `
        <input type="checkbox" id="filter-${normalized}" value="${normalized}" checked>
        <label for="filter-${normalized}">${original}</label>
      `;
      typeSelectorCheckboxes.appendChild(div);
    }
    
    // Add event listeners to all checkboxes
    typeSelectorCheckboxes.querySelectorAll('input[type="checkbox"]').forEach(cb => {
      cb.addEventListener('change', onFilterChange);
    });
    
    typeSelectorContainer.classList.remove('hidden');
  }
  
  /**
   * *** NEW ***
   * Handles click on any filter checkbox
   */
  function onFilterChange(e) {
    const allCheckbox = document.getElementById('filter-all');
    const typeCheckboxes = typeSelectorCheckboxes.querySelectorAll('input:not(#filter-all)');

    if (e.target.id === 'filter-all') {
      // If "All" is clicked, set all others to match it
      typeCheckboxes.forEach(cb => cb.checked = allCheckbox.checked);
    } else {
      // If a specific type is clicked, check if all are checked
      const allChecked = Array.from(typeCheckboxes).every(cb => cb.checked);
      allCheckbox.checked = allChecked;
    }
    
    filterDataPoints();
  }
  
  /**
   * *** NEW ***
   * Hides or shows data points based on the active filters
   */
  function filterDataPoints() {
    // 1. Get all active *normalized* filter types
    const activeFilters = new Set();
    const typeCheckboxes = typeSelectorCheckboxes.querySelectorAll('input:not(#filter-all):checked');
    typeCheckboxes.forEach(cb => activeFilters.add(cb.value));

    // 2. Filter the visible rows
    const allRows = dataPointContainer.querySelectorAll('.input-group');
    allRows.forEach(row => {
      const rowTypeNormalized = row.dataset.titleType;
      
      // Show if "All" is checked OR if the row's type is in the active set
      if (document.getElementById('filter-all').checked || activeFilters.has(rowTypeNormalized)) {
        row.classList.remove('hidden');
      } else {
        row.classList.add('hidden');
      }
    });
  }

  /**
   * Takes an array of data objects and populates the UI
   */
  function populateDataPoints(dataToPopulate) {
    clearDataPoints(); // Clear any manual entries
    
    let pointsFound = 0;
    dataToPopulate.forEach(item => {
      createDataPointWithValue(item.house, item.title, item.type);
      pointsFound++;
    });
    
    if (pointsFound < 2) {
      showError(`Found ${pointsFound} valid data pairs. Not enough data to calculate a pattern.`);
      createBlankDataPoint();
      createBlankDataPoint();
    }
  }

  
  /** Smart Parsers (Updated) */
  function normalizeType(text) {
    return text.toLowerCase().replace(/\s+/g, '');
  }
  
  function parseHouseNumber(cellText) {
    const numMatch = cellText.match(/(\d+)/);
    if (!numMatch) return null;
    let houseNum = parseInt(numMatch[1], 10);
    const numEndIndex = numMatch.index + numMatch[1].length;
    const restOfString = cellText.substring(numEndIndex).trim();
    if (restOfString.toUpperCase().startsWith('A')) {
      houseNum += 1;
    }
    return houseNum.toString();
  }
  
  function parseTitleNumber(cellText) {
    const numMatches = cellText.match(/(\d{5,})/g);
    if (!numMatches) return null;
    return numMatches[numMatches.length - 1];
  }
  
  function parseTitleIndication(cellText) {
    const numMatches = cellText.match(/(\d{5,})/g);
    if (!numMatches) return "";
    const lastNum = numMatches[numMatches.length - 1];
    const lastNumIndex = cellText.lastIndexOf(lastNum);
    return cellText.substring(0, lastNumIndex).trim();
  }
  

  
  // --- 2. Dynamic Add/Remove Logic ---

  function clearDataPoints() {
    dataPointContainer.innerHTML = '';
    typeSelectorContainer.classList.add('hidden');
    typeSelectorCheckboxes.innerHTML = ''; // Clear checkboxes too
  }

  function createBlankDataPoint() {
    createDataPointWithValue('', '', '');
  }
  
  function createDataPointWithValue(house, title, indication = '') {
    if (dataPointContainer.children.length >= MAX_POINTS) return;
    
    const newGroup = document.createElement('div');
    newGroup.className = 'input-group';
    // *** Store the NORMALIZED type for filtering ***
    newGroup.dataset.titleType = normalizeType(indication); 
    
    newGroup.innerHTML = `
      <div>
        <label>House No.</label>
        <input type="number" class="house-input" required>
      </div>
      <div>
        <label>Title No.</label>
        <input type="number" class="title-input" required>
      </div>
      <div class="title-type-wrapper">
        <label>Title Type</label>
        <input type="text" class="title-type-input" readonly>
      </div>
      <button type="button" class="remove-btn" title="Remove data point">&times;</button>
    `;
    
    newGroup.querySelector('.house-input').value = house;
    newGroup.querySelector('.title-input').value = title;
    
    const titleTypeInput = newGroup.querySelector('.title-type-input');
    const titleTypeWrapper = newGroup.querySelector('.title-type-wrapper');
    
    titleTypeInput.value = indication; // Show the *original* text
    
    if (!indication) {
      titleTypeWrapper.classList.add("hidden");
    }
    
    dataPointContainer.appendChild(newGroup);
    updatePointUI();
  }

  function updatePointUI() {
    const dataGroups = dataPointContainer.querySelectorAll('.input-group');
    const count = dataGroups.length;

    dataGroups.forEach((group, index) => {
      const houseLabel = group.querySelector('.house-input').previousElementSibling;
      if (houseLabel) houseLabel.textContent = `House No. ${index + 1}`;
      
      const titleLabel = group.querySelector('.title-input').previousElementSibling;
      if (titleLabel) titleLabel.textContent = `Title No. ${index + 1}`;
      
      const typeLabel = group.querySelector('.title-type-input').previousElementSibling;
      if (typeLabel) typeLabel.textContent = `Title Type ${index + 1}`;
    });
    
    addPointBtn.disabled = (count >= MAX_POINTS);

    const removeButtons = dataPointContainer.querySelectorAll('.remove-btn');
    removeButtons.forEach(btn => {
      btn.disabled = (count <= MIN_POINTS);
    });
  }
  
  addPointBtn.addEventListener("click", createBlankDataPoint);
  
  dataPointContainer.addEventListener("click", (e) => {
    if (e.target && e.target.classList.contains('remove-btn')) {
      e.target.closest('.input-group').remove();
      updatePointUI();
    }
  });

  // --- 3. Form Submission & Calculation Logic (UPDATED) ---

  form.addEventListener("submit", (e) => {
    e.preventDefault(); 
    
    resultContainer.classList.add("hidden");
    errorContainer.classList.add("hidden");
    resultContainer.classList.remove("success", "warning", "info");

    // A. Get Target and Parity
    const x_target = parseInt(targetHouse_el.value);
    if (isNaN(x_target)) {
        showError("Please enter a Target House Number.");
        return;
    }
    const targetParity = x_target % 2; 
    const parityType = targetParity === 0 ? "even" : "odd";
    
    // B. Get and AUTO-FILTER Data Points
    const filtered_x = []; 
    const filtered_y = []; 
    
    // *** UPDATED: Get the active filter types ***
    const activeFilterLabels = [];
    const allCheckbox = document.getElementById('filter-all');
    let activeFilterDisplay = "N/A (Manual Mode)";
    
    if (allCheckbox && !allCheckbox.checked) {
      const activeCheckboxes = typeSelectorCheckboxes.querySelectorAll('input:not(#filter-all):checked');
      activeCheckboxes.forEach(cb => {
        // Use the original text from the map
        activeFilterLabels.push(normalizedTypeMap.get(cb.value)); 
      });
      if (activeFilterLabels.length > 0) {
        activeFilterDisplay = activeFilterLabels.join(', ');
      }
    }
    
    const dataGroups = dataPointContainer.querySelectorAll('.input-group');
    
    dataGroups.forEach(group => {
      const isVisible = !group.classList.contains('hidden');
      
      if (isVisible) {
        const x_input = group.querySelector('.house-input').value;
        const y_input = group.querySelector('.title-input').value;
        
        if (x_input && y_input) {
          const x_val = parseInt(x_input);
          if (x_val % 2 === targetParity) {
            filtered_x.push(x_val);
            filtered_y.push(parseInt(y_input));
          }
        }
      }
    });

    // C. Validation on Filtered Data
    if (filtered_x.length < MIN_POINTS) {
      showError(`Not enough matching data. Please select a filter and ensure at least ${MIN_POINTS} <b>${parityType}</b> house numbers are visible.`);
      return;
    }

    // D. Transform 'x' values
    let transform_fn, transform_name;
    if (targetParity === 0) { // Even
      transform_fn = (x) => x / 2;
      transform_name = "(HouseNo / 2)";
    } else { // Odd
      transform_fn = (x) => (x + 1) / 2;
      transform_name = "(HouseNo + 1) / 2";
    }
    
    const x_prime_values = filtered_x.map(transform_fn);
    
    // E. Main Regression (Attempt 1)
    let main_regression = calculateRegression(x_prime_values, filtered_y);
    if (!main_regression) {
        showError("Cannot calculate a pattern: all entered House Numbers are identical.");
        return;
    }
    
    let r2_percent_initial = main_regression.r2 * 100;
    
    let final_regression = main_regression;
    let final_r2_percent = r2_percent_initial;
    let analysis_level = 'success';
    let analysis_message = `The data points form a strong linear pattern. The result is likely correct.`;

    // F. Outlier Auto-Correction Logic
    if (r2_percent_initial < CONFIDENCE_THRESHOLD && filtered_x.length >= 3) {
      const outlier_info = findOutlierByResidual(
        x_prime_values, 
        filtered_y, 
        filtered_x, 
        main_regression.m, 
        main_regression.c
      );
      
      const corrected_x_prime = x_prime_values.filter((_, idx) => idx !== outlier_info.outlier_index);
      const corrected_y = filtered_y.filter((_, idx) => idx !== outlier_info.outlier_index);
      
      let corrected_regression = calculateRegression(corrected_x_prime, corrected_y);
      
      if (corrected_regression && (corrected_regression.r2 * 100) >= CONFIDENCE_THRESHOLD) {
        final_regression = corrected_regression;
        final_r2_percent = corrected_regression.r2 * 100;
        analysis_level = 'info'; 
        analysis_message = `<b>Note:</b> We automatically ignored <b>House No. ${outlier_info.outlier_house}</b> as it was an outlier. This improved the confidence from <b>${r2_percent_initial.toFixed(1)}%</b> to <b>${final_r2_percent.toFixed(1)}%</b>.`;
      
      } else {
        analysis_level = 'warning';
        analysis_message = `The data points are inconsistent and do not form a clear pattern. The result is likely INCORRECT. Please double-check your data entries.`;
      }
    } else if (r2_percent_initial < CONFIDENCE_THRESHOLD) {
      analysis_level = 'warning';
      analysis_message = `The data points form a line, but more data is needed to confirm the pattern.`;
    }

    // G. Calculate Final Result
    const n_target = transform_fn(x_target);
    const y_target = (final_regression.m * n_target) + final_regression.c;
    const finalResult = Math.round(y_target);
    
    // H. Display Result
    showResult(
      x_target, 
      finalResult, 
      final_regression.m, 
      final_regression.c, 
      transform_name, 
      analysis_level, 
      final_r2_percent,
      analysis_message,
      activeFilterDisplay // Pass the filter display string
    );
  });
  
  
  // --- 4. Calculation Functions (No changes here) ---
  
  function calculateRegression(x_arr, y_arr) {
    const n = x_arr.length;
    if (n < 2) return null;

    const x_mean = x_arr.reduce((a, b) => a + b) / n;
    const y_mean = y_arr.reduce((a, b) => a + b) / n;

    let m_numerator = 0;
    let m_denominator = 0;

    for (let i = 0; i < n; i++) {
      m_numerator += (x_arr[i] - x_mean) * (y_arr[i] - y_mean);
      m_denominator += (x_arr[i] - x_mean) ** 2;
    }
    
    if (m_denominator === 0) {
      return null; 
    }

    const m = m_numerator / m_denominator;
    const c = y_mean - (m * x_mean);

    let ss_total = 0;
    let ss_residual = 0;
    
    for (let i = 0; i < y_arr.length; i++) {
      const y_predicted = (m * x_arr[i]) + c;
      ss_total += (y_arr[i] - y_mean) ** 2;
      ss_residual += (y_arr[i] - y_predicted) ** 2;
    }
    
    let r2 = 0;
    if (ss_total === 0) {
      r2 = ss_residual === 0 ? 1 : 0;
    } else {
      r2 = 1 - (ss_residual / ss_total);
    }
    
    return { m, c, r2 };
  }
  
  function findOutlierByResidual(x_prime_values, y_values, x_values, m, c) {
    let max_residual_sq = -1;
    let outlier_index = -1;

    for (let i = 0; i < x_values.length; i++) {
      const n_i = x_prime_values[i];
      const y_i = y_values[i];
      
      const predicted_y = (m * n_i) + c;
      const residual_sq = (y_i - predicted_y) ** 2;
      
      if (residual_sq > max_residual_sq) {
        max_residual_sq = residual_sq;
        outlier_index = i;
      }
    }
    
    return {
      outlier_house: x_values[outlier_index],
      outlier_index: outlier_index
    };
  }

  
  // --- 5. Helper Functions (UPDATED) ---
  
  function showError(message) {
    errorText.innerHTML = message; 
    errorContainer.classList.remove("hidden");
  }
  
  function showResult(targetHouse, result, m, c, transform_name, level, r2_percent, message, activeFilterType) {
    const h2_el = resultContainer.querySelector('h2');
    h2_el.innerHTML = `House Number: <span class="target-house-display">${targetHouse}</span><br>Calculated Title Number:`;
    
    resultText.textContent = result;
    
    // Show the Title Type(s) in the result
    resultTitleType.textContent = activeFilterType;
    
    const c_string = c >= 0 ? `+ ${c.toFixed(2)}` : `- ${Math.abs(c).toFixed(2)}`;
    
    formulaText.innerHTML = `
      Let <b>n = ${transform_name}</b><br>
      Formula: <b>Title = (${m.toFixed(4)} * n) ${c_string}</b>
    `;
    
    resultContainer.classList.add(level);
    
    if(level === 'success' || (level === 'info' && r2_percent >= CONFIDENCE_THRESHOLD)) {
      analysisLevel.textContent = `High Confidence (${r2_percent.toFixed(1)}%)`;
    } else if (level === 'warning') {
      analysisLevel.textContent = `Low Confidence (${r2_percent.toFixed(1)}%)`;
    } else { 
      analysisLevel.textContent = `Analysis Note`;
    }
    
    analysisText.innerHTML = message;
    
    resultContainer.classList.remove("hidden");
  }
  
  // --- Initialize UI ---
  // Clear the container first, *then* add the 2 blank points
  clearDataPoints();
  createBlankDataPoint();
  createBlankDataPoint();
  
} // End of initializeApp()


// --- Main Entry Point ---
// Wait for the DOM to be ready, then start loading the SheetJS script
document.addEventListener("DOMContentLoaded", () => {
  loadScript(SHEET_JS_URL, onSheetJsLoaded);
});