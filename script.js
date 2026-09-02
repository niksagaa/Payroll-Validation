let isMappingActive = false;
let globalDr2Cols = [];
let allSheetNames = [];
let currentSheetIndex = 0;
let accumulatedMappings = {};
let accumulatedIdKeys = {};

function handleFileSelect(input, containerId, textId) {
    if (input.files && input.files[0]) {
        document.getElementById(textId).innerText = "📄 " + input.files[0].name;
        document.getElementById(containerId).style.display = 'block';

        if (input.id === 'payroll_inputs') {
            fetchSheetNames(input.files[0]);
        }
    }
}

function checkExistingInputs() {
    const inputElem = document.getElementById('payroll_inputs');
    if (inputElem.files && inputElem.files[0] && document.getElementById('sheet_name').options.length <= 1) {
        fetchSheetNames(inputElem.files[0]);
    }
}

function toggleSettings() {
    const content = document.getElementById('settings_content');
    content.style.display = content.style.display === 'block' ? 'none' : 'block';
}

async function fetchSheetNames(file) {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('password', document.getElementById('password').value || '');

    const sheetSelect = document.getElementById('sheet_name');
    sheetSelect.innerHTML = '<option value="">Loading sheets...</option>';

    try {
        const res = await fetch('/get-sheets', { method: 'POST', body: formData });
        if (!res.ok) throw new Error("Failed to read sheets.");
        const data = await res.json();

        sheetSelect.innerHTML = '';

        if (data.sheet_names && data.sheet_names.length > 0) {
            allSheetNames = data.sheet_names;
            currentSheetIndex = 0;

            allSheetNames.forEach((sheet, idx) => {
                const opt = document.createElement('option');
                opt.value = sheet;
                opt.textContent = sheet;
                if (idx === 0) opt.selected = true;
                sheetSelect.appendChild(opt);
            });
        } else {
            allSheetNames = ["Sheet1"];
            sheetSelect.innerHTML = '<option value="Sheet1">Sheet1</option>';
        }
    } catch (err) {
        console.error("Error fetching sheets:", err);
        allSheetNames = ["Sheet1"];
        sheetSelect.innerHTML = '<option value="Sheet1">Sheet1</option>';
    }
}

async function onSheetChange() {
    const selectedSheet = document.getElementById('sheet_name').value;
    const foundIdx = allSheetNames.indexOf(selectedSheet);
    if (foundIdx !== -1) {
        currentSheetIndex = foundIdx;
    }
    if (isMappingActive) {
        await triggerAnalyzeMappingPreview(true);
    }
}

async function startMappingFlow() {
    const inputs = document.getElementById('payroll_inputs').files[0];
    const master = document.getElementById('payroll_masterfile').files[0];
    if (!inputs || !master) {
        alert("Please upload both Payroll Inputs and Masterfile files.");
        return;
    }

    isMappingActive = true;
    document.getElementById('mappingSectionWrapper').style.display = 'block';
    document.getElementById('initialBtnGroup').style.display = 'none';
    document.getElementById('multiTabBtnGroup').style.display = 'flex';

    currentSheetIndex = 0;
    accumulatedMappings = {};
    accumulatedIdKeys = {};

    await triggerAnalyzeMappingPreview(false);
}

async function triggerAnalyzeMappingPreview(isTabSwitch) {
    const inputs = document.getElementById('payroll_inputs').files[0];
    const master = document.getElementById('payroll_masterfile').files[0];
    const password = document.getElementById('password').value;
    const currentSheet = allSheetNames[currentSheetIndex] || document.getElementById('sheet_name').value;

    document.getElementById('sheet_name').value = currentSheet;
    document.getElementById('tabIndicator').innerText = `Tab ${currentSheetIndex + 1} of ${allSheetNames.length}: ${currentSheet}`;

    const formData = new FormData();
    formData.append('payroll_inputs', inputs);
    formData.append('payroll_masterfile', master);
    formData.append('password', password);
    formData.append('sheet_name', currentSheet);

    const empColSelect = document.getElementById('employee_id_col');
    formData.append('employee_id_col', empColSelect.value);

    try {
        const res = await fetch('/analyze-mapping', { method: 'POST', body: formData });
        if (!res.ok) {
            const errData = await res.json();
            throw new Error(errData.detail || "Analysis failed.");
        }
        const data = await res.json();
        globalDr2Cols = data.dr2_columns;

        empColSelect.innerHTML = '<option value="">-- Auto-detect --</option>';
        data.main_columns.forEach(col => {
            const opt = document.createElement('option');
            opt.value = col;
            opt.textContent = col;
            if (col === data.detected_key_main) opt.selected = true;
            empColSelect.appendChild(opt);
        });

        renderMappingUI(data.data_cols, data.initial_mapping, data.dr2_columns);
    } catch (err) {
        alert("Error: " + err.message);
    }
}

function renderMappingUI(dataCols, initialMapping, dr2Cols) {
    const container = document.getElementById('mappingContainer');
    container.innerHTML = '';

    dataCols.forEach(mCol => {
        const row = document.createElement('div');
        row.className = 'mapping-row';

        const colNameDiv = document.createElement('div');
        colNameDiv.className = 'mapping-col-name';
        colNameDiv.textContent = mCol;

        const arrowDiv = document.createElement('div');
        arrowDiv.className = 'mapping-arrow';
        arrowDiv.textContent = '➔';

        const selectElem = document.createElement('select');
        selectElem.className = 'custom-select mapping-select';
        selectElem.dataset.mainCol = mCol;

        const skipOpt = document.createElement('option');
        skipOpt.value = '-- Skip / Do Not Compare --';
        skipOpt.textContent = '-- Skip / Do Not Compare --';
        selectElem.appendChild(skipOpt);

        dr2Cols.forEach(dCol => {
            const opt = document.createElement('option');
            opt.value = dCol;
            opt.textContent = dCol;
            if (initialMapping[mCol] === dCol) {
                opt.selected = true;
            }
            selectElem.appendChild(opt);
        });

        row.appendChild(colNameDiv);
        row.appendChild(arrowDiv);
        row.appendChild(selectElem);
        container.appendChild(row);
    });
}

function saveCurrentTabMapping() {
    const mapping = {};
    const selects = document.querySelectorAll('.mapping-select');
    selects.forEach(sel => {
        mapping[sel.dataset.mainCol] = sel.value;
    });

    const currentSheet = allSheetNames[currentSheetIndex];
    accumulatedMappings[currentSheet] = mapping;
    accumulatedIdKeys[currentSheet] = document.getElementById('employee_id_col').value;
}

async function handleNextTab() {
    saveCurrentTabMapping();
    currentSheetIndex++;

    if (currentSheetIndex < allSheetNames.length) {
        await triggerAnalyzeMappingPreview(true);
    } else {
        await submitFinalValidation();
    }
}

async function skipCurrentTab() {
    const currentSheet = allSheetNames[currentSheetIndex];
    accumulatedMappings[currentSheet] = { "__SKIPPED__": true };
    currentSheetIndex++;

    if (currentSheetIndex < allSheetNames.length) {
        await triggerAnalyzeMappingPreview(true);
    } else {
        await submitFinalValidation();
    }
}

async function submitFinalValidation() {
    const inputs = document.getElementById('payroll_inputs').files[0];
    const master = document.getElementById('payroll_masterfile').files[0];
    const password = document.getElementById('password').value;

    const formData = new FormData();
    formData.append('payroll_inputs', inputs);
    formData.append('payroll_masterfile', master);
    formData.append('password', password);
    formData.append('all_mappings_json', JSON.stringify(accumulatedMappings));
    formData.append('all_id_keys_json', JSON.stringify(accumulatedIdKeys));

    document.getElementById('progressWrapper').style.display = 'block';
    document.getElementById('progressStatus').innerText = 'Generating validation report...';
    document.getElementById('progressBarFill').style.width = '70%';

    try {
        const res = await fetch('/validate-mapped', { method: 'POST', body: formData });
        if (!res.ok) {
            const errData = await res.json();
            throw new Error(errData.detail || "Validation failed.");
        }

        document.getElementById('progressBarFill').style.width = '100%';
        document.getElementById('progressStatus').innerText = 'Download complete!';

        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = "Payroll_Validated_Report.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();

        setTimeout(() => { resetForm(); }, 2000);
    } catch (err) {
        alert("Error during validation: " + err.message);
        document.getElementById('progressWrapper').style.display = 'none';
    }
}

function clearFile(event, inputId, badgeId) {
    event.stopPropagation();
    document.getElementById(inputId).value = '';
    document.getElementById(badgeId).style.display = 'none';
    if (inputId === 'payroll_inputs') {
        document.getElementById('sheet_name').innerHTML = '<option value="">-- Please upload Payroll Inputs File first --</option>';
        allSheetNames = [];
    }
}

function resetForm() {
    document.getElementById('payrollForm').reset();
    document.getElementById('inputs_badge_container').style.display = 'none';
    document.getElementById('master_badge_container').style.display = 'none';
    document.getElementById('mappingSectionWrapper').style.display = 'none';
    document.getElementById('initialBtnGroup').style.display = 'flex';
    document.getElementById('multiTabBtnGroup').style.display = 'none';
    document.getElementById('progressWrapper').style.display = 'none';
    document.getElementById('sheet_name').innerHTML = '<option value="">-- Please upload Payroll Inputs File first --</option>';
    isMappingActive = false;
    currentSheetIndex = 0;
    accumulatedMappings = {};
    accumulatedIdKeys = {};
}