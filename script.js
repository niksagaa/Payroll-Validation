let isMappingActive = false;
let globalDr2Cols = [];
let allSheetNames = [];
let currentSheetIndex = 0;
let accumulatedMappings = {};
let accumulatedIdKeys = {};

const passwordInput = document.getElementById('password') || document.getElementById('passwordInput');

function goToStep(stepNumber) {
    document.querySelectorAll('.step-panel').forEach(panel => panel.classList.remove('active-panel'));
    const targetPanel = document.getElementById(`panelStep${stepNumber}`);
    if (targetPanel) targetPanel.classList.add('active-panel');

    for (let i = 1; i <= 2; i++) {
        const ind = document.getElementById(`indicatorStep${i}`);
        if (ind) {
            ind.classList.remove('active', 'completed');
            if (i < stepNumber) ind.classList.add('completed');
            else if (i === stepNumber) ind.classList.add('active');
        }
    }
}

function handleFileSelect(input, containerId, textId) {
    if (input.files && input.files[0]) {
        const textElem = document.getElementById(textId);
        const containerElem = document.getElementById(containerId);
        if (textElem) textElem.innerText = "📄 " + input.files[0].name;
        if (containerElem) containerElem.style.display = 'block';
    }
}

async function executeSmartAutoValidation() {
    const inputs = document.getElementById('payroll_inputs').files[0];
    const master = document.getElementById('payroll_masterfile').files[0];
    if (!inputs || !master) {
        alert("Please upload both Payroll Inputs and Masterfile first.");
        return;
    }
    await submitFinalValidation(true);
}

async function startMappingFlowAndAdvance() {
    const inputs = document.getElementById('payroll_inputs').files[0];
    const master = document.getElementById('payroll_masterfile').files[0];
    const password = passwordInput ? passwordInput.value : '';

    if (!inputs || !master) {
        alert("Please upload both Payroll Inputs and Masterfile in Step 1 first.");
        return;
    }

    const formDataCheck = new FormData();
    formDataCheck.append('file', inputs);
    formDataCheck.append('password', password);

    try {
        const checkRes = await fetch('/get-sheets', { method: 'POST', body: formDataCheck });
        if (!checkRes.ok) {
            const errData = await checkRes.json().catch(() => ({}));
            throw new Error(errData.detail || "Mali ang password o hindi mabasa ang file.");
        }
        const checkData = await checkRes.json();
        allSheetNames = (checkData.sheet_names && checkData.sheet_names.length > 0) ? checkData.sheet_names : ["Sheet1"];
    } catch (err) {
        alert("Password Error: " + err.message);
        return;
    }

    isMappingActive = true;
    goToStep(2);
    currentSheetIndex = 0;
    accumulatedMappings = {};
    accumulatedIdKeys = {};

    await triggerAnalyzeMappingPreview();
}

function renderTabPills() {
    const container = document.getElementById('tabPillsContainer');
    if (!container) return;
    container.innerHTML = '';

    allSheetNames.forEach((sheet, idx) => {
        const pill = document.createElement('div');
        pill.className = `tab-pill ${idx === currentSheetIndex ? 'active-tab-pill' : ''}`;
        pill.textContent = `${idx + 1}. ${sheet}`;
        pill.onclick = async () => {
            saveCurrentTabMapping();
            currentSheetIndex = idx;
            await triggerAnalyzeMappingPreview();
        };
        container.appendChild(pill);
    });
}

async function triggerAnalyzeMappingPreview() {
    const inputs = document.getElementById('payroll_inputs').files[0];
    const master = document.getElementById('payroll_masterfile').files[0];
    const password = passwordInput ? passwordInput.value : '';
    const currentSheet = allSheetNames[currentSheetIndex] || 'Sheet1';
    
    renderTabPills();

    const formData = new FormData();
    formData.append('payroll_inputs', inputs);
    formData.append('payroll_masterfile', master);
    formData.append('password', password);
    formData.append('sheet_name', currentSheet);

    const empColSelect = document.getElementById('employee_id_col');
    if (empColSelect) {
        if (accumulatedIdKeys[currentSheet]) {
            empColSelect.value = accumulatedIdKeys[currentSheet];
        }
        formData.append('employee_id_col', empColSelect.value);
    }

    try {
        const res = await fetch('/analyze-mapping', { method: 'POST', body: formData });
        if (!res.ok) {
            const errData = await res.json();
            throw new Error(errData.detail || "Analysis failed.");
        }
        const data = await res.json();
        globalDr2Cols = data.dr2_columns;

        if (empColSelect && !accumulatedIdKeys[currentSheet]) {
            empColSelect.innerHTML = '<option value="">-- Automatically detected --</option>';
            data.main_columns.forEach(col => {
                const opt = document.createElement('option');
                opt.value = col;
                opt.textContent = col;
                if (col === data.detected_key_main) opt.selected = true;
                empColSelect.appendChild(opt);
            });
        }

        const savedMapping = accumulatedMappings[currentSheet] || data.initial_mapping;
        renderMappingUI(data.data_cols, savedMapping, data.dr2_columns);
    } catch (err) {
        alert("Error: " + err.message);
    }
}

function renderMappingUI(dataCols, initialMapping, dr2Cols) {
    const container = document.getElementById('mappingContainer');
    if (!container) return;
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
            if (initialMapping[mCol] === dCol) opt.selected = true;
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
    document.querySelectorAll('.mapping-select').forEach(sel => {
        mapping[sel.dataset.mainCol] = sel.value;
    });
    const currentSheet = allSheetNames[currentSheetIndex];
    accumulatedMappings[currentSheet] = mapping;
    
    const empColSelect = document.getElementById('employee_id_col');
    accumulatedIdKeys[currentSheet] = empColSelect ? empColSelect.value : '';
}

async function handleNextTab() {
    saveCurrentTabMapping();
    currentSheetIndex++;
    if (currentSheetIndex < allSheetNames.length) {
        await triggerAnalyzeMappingPreview();
    } else {
        await submitFinalValidation(false);
    }
}

async function skipCurrentTab() {
    const currentSheet = allSheetNames[currentSheetIndex];
    accumulatedMappings[currentSheet] = { "__SKIPPED__": true };
    currentSheetIndex++;
    if (currentSheetIndex < allSheetNames.length) {
        await triggerAnalyzeMappingPreview();
    } else {
        await submitFinalValidation(false);
    }
}

async function submitFinalValidation(isAutoRun = false) {
    const inputs = document.getElementById('payroll_inputs').files[0];
    const master = document.getElementById('payroll_masterfile').files[0];
    const password = passwordInput ? passwordInput.value : '';

    const formData = new FormData();
    formData.append('payroll_inputs', inputs);
    formData.append('payroll_masterfile', master);
    formData.append('password', password);
    formData.append('all_mappings_json', JSON.stringify(accumulatedMappings));
    formData.append('all_id_keys_json', JSON.stringify(accumulatedIdKeys));

    const progressWrapper = document.getElementById('progressWrapper');
    const progressStatus = document.getElementById('progressStatus');
    const progressBarFill = document.getElementById('progressBarFill');

    if (isAutoRun) goToStep(2);
    if (progressWrapper) progressWrapper.style.display = 'block';
    if (progressStatus) progressStatus.innerText = 'Generating validated report...';
    if (progressBarFill) progressBarFill.style.width = '70%';

    try {
        const res = await fetch('/validate-mapped', { method: 'POST', body: formData });
        if (!res.ok) {
            const errData = await res.json();
            throw new Error(errData.detail || "Validation failed.");
        }

        if (progressBarFill) progressBarFill.style.width = '100%';
        if (progressStatus) progressStatus.innerText = 'Report successfully downloaded!';

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
        alert("Validation Error: " + err.message);
        if (progressWrapper) progressWrapper.style.display = 'none';
        goToStep(1);
    }
}

function clearFile(event, inputId, badgeId) {
    event.stopPropagation();
    const inputElem = document.getElementById(inputId);
    const badgeElem = document.getElementById(badgeId);
    if (inputElem) inputElem.value = '';
    if (badgeElem) badgeElem.style.display = 'none';
    if (inputId === 'payroll_inputs') allSheetNames = [];
}

function resetForm() {
    const formElem = document.getElementById('payrollForm');
    if (formElem) formElem.reset();
    
    const inputBadge = document.getElementById('inputs_badge_container');
    const masterBadge = document.getElementById('master_badge_container');
    const progressWrap = document.getElementById('progressWrapper');

    if (inputBadge) inputBadge.style.display = 'none';
    if (masterBadge) masterBadge.style.display = 'none';
    if (progressWrap) progressWrap.style.display = 'none';
    
    goToStep(1);
    isMappingActive = false;
    currentSheetIndex = 0;
    allSheetNames = [];
    accumulatedMappings = {};
    accumulatedIdKeys = {};
}