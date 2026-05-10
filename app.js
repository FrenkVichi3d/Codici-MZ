let currentSchema = [];
let selectedSheet = null;
let selectedLevel1 = null;
let selectedLevel2 = null;

// DOM Elements
const sheetSelect = document.getElementById('sheetSelect');
const level1Select = document.getElementById('level1Select');
const level2Select = document.getElementById('level2Select');
const progressiveInput = document.getElementById('progressiveInput');
const descriptionInput = document.getElementById('descriptionInput');
const generateBtn = document.getElementById('generateBtn');

const resultCard = document.getElementById('resultCard');
const finalCodeEl = document.getElementById('finalCode');
const codeBreakdownEl = document.getElementById('codeBreakdown');
const copyBtn = document.getElementById('copyBtn');
const saveBtn = document.getElementById('saveBtn');
const syncBtn = document.getElementById('syncBtn');

const historyList = document.getElementById('historyList');
const clearHistoryBtn = document.getElementById('clearHistoryBtn');

const appMain = document.getElementById('appMain');
const loadingOverlay = document.getElementById('loadingOverlay');

// Office Initialization
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // App is loaded inside Excel
        initApp();
    } else {
        // App opened in regular browser
        loadingOverlay.innerHTML = '<i class="fa-solid fa-triangle-exclamation fa-3x" style="color: var(--danger);"></i><p style="margin-top: 15px; font-weight: 600;">Errore di Contesto</p><p style="font-size: 0.8rem; color: var(--text-muted); padding: 0 20px;">Questa app deve essere aperta all\'interno di Excel come Componente Aggiuntivo Web (Web Add-in).</p>';
    }
});

async function initApp() {
    loadHistory();
    await fetchSchemaFromExcel();
    
    loadingOverlay.style.display = 'none';
    appMain.style.display = 'block';
}

function extractProgressive(code_str) {
    if (!code_str) return 0;
    const match = String(code_str).match(/\.(\d{4})\b/);
    if (match) {
        return parseInt(match[1], 10);
    }
    return 0;
}

async function fetchSchemaFromExcel() {
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();
            
            currentSchema = [];
            
            for (let i = 0; i < sheets.items.length; i++) {
                const sheet = sheets.items[i];
                const sheetName = sheet.name;
                const sheetObj = { sheet: sheetName, categories: [] };
                
                const usedRange = sheet.getUsedRangeOrNullObject();
                await context.sync();
                
                if (usedRange.isNullObject) {
                    continue; // Skip empty sheets
                }
                
                usedRange.load("values");
                await context.sync();
                
                const values = usedRange.values;
                const is_4_39 = sheetName.includes('4.39');
                
                let current_level1 = null;
                let current_level2 = null;
                
                if (is_4_39) {
                    current_level1 = {name: 'ELETTRICI', code: '39', prefix: '4', subcategories: [], max_progressive: 0};
                    sheetObj.categories.push(current_level1);
                }
                
                for (let rowIdx = 0; rowIdx < values.length; rowIdx++) {
                    const row = values[rowIdx];
                    const v0 = row[0] ? String(row[0]).trim() : "";
                    const v1 = row[1] ? String(row[1]).trim() : "";
                    const v2 = row[2] ? String(row[2]).trim() : "";
                    const v3 = row.length > 3 && row[3] ? String(row[3]).trim() : "";
                    const v4 = row.length > 4 && row[4] ? String(row[4]).trim() : "";
                    const v5 = row.length > 5 && row[5] ? String(row[5]).trim() : "";
                    
                    if (is_4_39) {
                        if (v1 && v2 && v2 !== 'nan' && v2 !== 'N. PZ') {
                            let c2 = v2.replace('.0', '');
                            c2 = isNaN(c2) ? c2 : c2.padStart(2, '0');
                            current_level2 = {name: v1, code: c2, max_progressive: 0};
                            current_level1.subcategories.push(current_level2);
                        }
                        if (v3) {
                            const prog = extractProgressive(v3);
                            if (current_level2 && prog > current_level2.max_progressive) {
                                current_level2.max_progressive = prog;
                            } else if (current_level1 && prog > current_level1.max_progressive) {
                                current_level1.max_progressive = prog;
                            }
                        }
                    } else {
                        if (v1 && v2 && v2 !== 'nan' && v2 !== 'N. PZ') {
                            let c1 = v2.replace('.0', '');
                            c1 = isNaN(c1) ? c1 : c1.padStart(2, '0');
                            let p = v0 ? v0.replace('.0', '').replace('.', '') : sheetName.split(' ')[0];
                            current_level1 = {name: v1, code: c1, prefix: p, subcategories: [], max_progressive: 0};
                            sheetObj.categories.push(current_level1);
                            current_level2 = null;
                        }
                        if (v3 && v4 && v4 !== 'nan' && current_level1) {
                            let c2 = v4.replace('.0', '');
                            c2 = isNaN(c2) ? c2 : c2.padStart(2, '0');
                            current_level2 = {name: v3, code: c2, max_progressive: 0};
                            current_level1.subcategories.push(current_level2);
                        }
                        if (v5) {
                            const prog = extractProgressive(v5);
                            if (current_level2 && prog > current_level2.max_progressive) {
                                current_level2.max_progressive = prog;
                            } else if (current_level1 && prog > current_level1.max_progressive) {
                                current_level1.max_progressive = prog;
                            }
                        }
                    }
                }
                
                if (sheetObj.categories.length > 0) {
                    currentSchema.push(sheetObj);
                }
            }
            populateSheets();
        });
    } catch (error) {
        console.error("Errore durante la lettura da Excel:", error);
        // Error is hidden from UI usually, just logged
    }
}

function populateSheets() {
    sheetSelect.innerHTML = '<option value="">Seleziona...</option>';
    currentSchema.forEach((sheet, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = sheet.sheet;
        sheetSelect.appendChild(option);
    });
}

sheetSelect.addEventListener('change', (e) => {
    const index = e.target.value;
    if (index === '') {
        selectedSheet = null;
        level1Select.innerHTML = '<option value="">Seleziona Categoria Principale prima...</option>';
        level1Select.disabled = true;
        level2Select.innerHTML = '<option value="">Seleziona Livello 1 prima...</option>';
        level2Select.disabled = true;
        progressiveInput.value = 1;
        return;
    }
    
    selectedSheet = currentSchema[index];
    populateLevel1();
});

level1Select.addEventListener('change', (e) => {
    const index = e.target.value;
    if (index === '') {
        selectedLevel1 = null;
        level2Select.innerHTML = '<option value="">Seleziona Livello 1 prima...</option>';
        level2Select.disabled = true;
        progressiveInput.value = 1;
        return;
    }
    
    selectedLevel1 = selectedSheet.categories[index];
    populateLevel2();
});

level2Select.addEventListener('change', (e) => {
    const index = e.target.value;
    if (index === '') {
        selectedLevel2 = null;
        progressiveInput.value = (selectedLevel1.max_progressive || 0) + 1;
        return;
    }
    selectedLevel2 = selectedLevel1.subcategories[index];
    if (selectedLevel2) {
        progressiveInput.value = (selectedLevel2.max_progressive || 0) + 1;
    }
});

function populateLevel1() {
    level1Select.innerHTML = '<option value="">Seleziona...</option>';
    if (selectedSheet.categories.length === 0) {
        level1Select.innerHTML = '<option value="">Nessun Livello 1</option>';
        level1Select.disabled = true;
        level2Select.innerHTML = '<option value="">-</option>';
        level2Select.disabled = true;
        selectedLevel1 = null;
        selectedLevel2 = null;
        progressiveInput.value = 1;
        return;
    }
    
    selectedSheet.categories.forEach((cat, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = `${cat.code} - ${cat.name}`;
        level1Select.appendChild(option);
    });
    level1Select.disabled = false;
    level2Select.innerHTML = '<option value="">Seleziona Livello 1 prima...</option>';
    level2Select.disabled = true;
    selectedLevel1 = null;
    selectedLevel2 = null;
    progressiveInput.value = 1;
}

function populateLevel2() {
    level2Select.innerHTML = '<option value="">Seleziona (Opzionale)...</option>';
    if (!selectedLevel1.subcategories || selectedLevel1.subcategories.length === 0) {
        level2Select.innerHTML = '<option value="">Nessun Livello 2</option>';
        level2Select.disabled = true;
        selectedLevel2 = null;
        progressiveInput.value = (selectedLevel1.max_progressive || 0) + 1;
        return;
    }
    
    selectedLevel1.subcategories.forEach((sub, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = `${sub.code} - ${sub.name}`;
        level2Select.appendChild(option);
    });
    level2Select.disabled = false;
    selectedLevel2 = null;
    progressiveInput.value = (selectedLevel1.max_progressive || 0) + 1;
}

syncBtn.addEventListener('click', async () => {
    const icon = syncBtn.querySelector('i');
    icon.classList.add('fa-spin');
    
    const currSheetIdx = sheetSelect.value;
    const currL1Idx = level1Select.value;
    const currL2Idx = level2Select.value;
    
    await fetchSchemaFromExcel();
    
    if (currSheetIdx !== "") {
        sheetSelect.value = currSheetIdx;
        selectedSheet = currentSchema[currSheetIdx];
        populateLevel1();
        if (currL1Idx !== "") {
            level1Select.value = currL1Idx;
            selectedLevel1 = selectedSheet.categories[currL1Idx];
            populateLevel2();
            if (currL2Idx !== "") {
                level2Select.value = currL2Idx;
                selectedLevel2 = selectedLevel1.subcategories[currL2Idx];
            }
        }
    }
    
    if (selectedLevel2) {
        progressiveInput.value = (selectedLevel2.max_progressive || 0) + 1;
    } else if (selectedLevel1) {
        progressiveInput.value = (selectedLevel1.max_progressive || 0) + 1;
    }
    
    icon.classList.remove('fa-spin');
    icon.className = 'fa-solid fa-check text-success';
    setTimeout(() => {
        icon.className = 'fa-solid fa-rotate';
    }, 2000);
});

generateBtn.addEventListener('click', () => {
    if (!selectedSheet || !selectedLevel1) {
        alert('Seleziona almeno la Categoria Principale e il Livello 1.');
        return;
    }
    
    let prefix = selectedLevel1.prefix;
    if (!prefix) {
        prefix = selectedSheet.sheet.split(' ')[0].replace('.', '');
    }
    
    const lvl1Code = selectedLevel1.code;
    const lvl2Code = selectedLevel2 ? selectedLevel2.code : '00';
    let progVal = parseInt(progressiveInput.value) || 1;
    const progCode = progVal.toString().padStart(4, '0');
    
    const description = descriptionInput.value.trim().toUpperCase();
    if (!description) {
        alert('La descrizione è obbligatoria per salvare sul file Excel.');
        return;
    }
    
    let finalCode = `${prefix}.${lvl1Code}.${lvl2Code}.${progCode} - ${description}`;
    finalCodeEl.textContent = finalCode;
    
    codeBreakdownEl.innerHTML = `
        <div class="breakdown-item"><span class="breakdown-label">Prefisso</span><span class="breakdown-value">${prefix}</span></div>
        <div class="breakdown-item"><span class="breakdown-label">Livello 1</span><span class="breakdown-value">${lvl1Code}</span></div>
        <div class="breakdown-item"><span class="breakdown-label">Livello 2</span><span class="breakdown-value">${lvl2Code}</span></div>
        <div class="breakdown-item"><span class="breakdown-label">Progressivo</span><span class="breakdown-value">${progCode}</span></div>
    `;
    
    resultCard.style.display = 'block';
    resultCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
});

copyBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(finalCodeEl.textContent).then(() => {
        const icon = copyBtn.querySelector('i');
        icon.className = 'fa-solid fa-check text-success';
        setTimeout(() => icon.className = 'fa-regular fa-copy', 2000);
    });
});

saveBtn.addEventListener('click', async () => {
    const codeString = finalCodeEl.textContent;
    
    saveBtn.disabled = true;
    saveBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Scrittura...';
    
    try {
        await insertCodeIntoExcel(selectedSheet.sheet, selectedLevel1.code, selectedLevel2 ? selectedLevel2.code : '00', codeString);
        
        // Successo
        if (selectedLevel2) selectedLevel2.max_progressive = parseInt(progressiveInput.value);
        else if (selectedLevel1) selectedLevel1.max_progressive = parseInt(progressiveInput.value);
        
        const codeOnly = codeString.split(' - ')[0];
        const desc = codeString.substring(codeOnly.length + 3);
        let history = JSON.parse(localStorage.getItem('mz_code_history') || '[]');
        history.unshift({ code: codeOnly, description: desc || 'Nessuna descrizione', timestamp: new Date().toISOString() });
        if (history.length > 50) history = history.slice(0, 50);
        localStorage.setItem('mz_code_history', JSON.stringify(history));
        
        progressiveInput.value = parseInt(progressiveInput.value) + 1;
        descriptionInput.value = '';
        loadHistory();
        
        saveBtn.innerHTML = '<i class="fa-solid fa-check"></i> Inserito con successo!';
        saveBtn.classList.add('btn-success');
        
        setTimeout(() => {
            saveBtn.disabled = false;
            saveBtn.innerHTML = '<i class="fa-solid fa-file-excel"></i> Inserisci riga in Excel';
            saveBtn.classList.remove('btn-success');
        }, 3000);
        
    } catch (err) {
        console.error(err);
        alert('Errore durante l\'inserimento: ' + err.message);
        saveBtn.disabled = false;
        saveBtn.innerHTML = '<i class="fa-solid fa-file-excel"></i> Riprova Inserimento';
    }
});

async function insertCodeIntoExcel(sheetName, level1Code, level2Code, newCodeString) {
    return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const usedRange = sheet.getUsedRange();
        usedRange.load(["values", "rowIndex"]);
        await context.sync();
        
        const values = usedRange.values;
        const startRowIdx = usedRange.rowIndex;
        
        let target_l1_found = false;
        let target_l2_found = false;
        let last_item_row_idx = -1;
        
        const is_4_39 = sheetName.includes('4.39');
        
        for (let i = 0; i < values.length; i++) {
            const row = values[i];
            const v1 = row[1] ? String(row[1]).trim() : "";
            const v2 = row[2] ? String(row[2]).trim() : "";
            const v3 = row.length > 3 && row[3] ? String(row[3]).trim() : "";
            const v4 = row.length > 4 && row[4] ? String(row[4]).trim() : "";
            const v5 = row.length > 5 && row[5] ? String(row[5]).trim() : "";
            
            if (is_4_39) {
                target_l1_found = true;
                if (v1 && v2 && v2 !== 'N. PZ') {
                    let c2 = v2.replace('.0', '');
                    c2 = isNaN(c2) ? c2 : c2.padStart(2, '0');
                    if (c2 === level2Code) {
                        target_l2_found = true;
                        last_item_row_idx = i;
                    } else if (target_l2_found) break;
                }
                if (target_l2_found && v3) last_item_row_idx = i;
            } else {
                if (v1 && v2 && v2 !== 'N. PZ') {
                    let c1 = v2.replace('.0', '');
                    c1 = isNaN(c1) ? c1 : c1.padStart(2, '0');
                    if (c1 === level1Code) {
                        target_l1_found = true;
                        if (level2Code === '00') last_item_row_idx = i;
                    } else if (target_l1_found && target_l2_found) break;
                    else if (target_l1_found && !target_l2_found && level2Code === '00') break;
                }
                if (target_l1_found && v3 && v4) {
                    let c2 = v4.replace('.0', '');
                    c2 = isNaN(c2) ? c2 : c2.padStart(2, '0');
                    if (c2 === level2Code) {
                        target_l2_found = true;
                        last_item_row_idx = i;
                    } else if (target_l2_found) break;
                }
                if (target_l2_found && v5) last_item_row_idx = i;
                if (target_l1_found && !target_l2_found && level2Code === '00' && v5) last_item_row_idx = i;
            }
        }
        
        if (!target_l1_found) throw new Error("Livello 1 non trovato nel foglio.");
        if (level2Code !== '00' && !target_l2_found) throw new Error("Livello 2 non trovato nel foglio.");
        
        let insertAbsoluteRow = -1;
        if (last_item_row_idx !== -1) {
            insertAbsoluteRow = startRowIdx + last_item_row_idx + 1; // Row *after* the last found item
        } else {
            insertAbsoluteRow = startRowIdx + values.length; // End of sheet
        }
        
        // Insert a new row by shifting down
        const rangeToInsert = sheet.getRangeByIndexes(insertAbsoluteRow, 0, 1, 10).getEntireRow();
        rangeToInsert.insert(Excel.InsertShiftDirection.down);
        
        // Write the code in the specific column
        let targetCell;
        if (is_4_39) {
            targetCell = sheet.getCell(insertAbsoluteRow, 3); // Column D is index 3
        } else {
            targetCell = sheet.getCell(insertAbsoluteRow, 5); // Column F is index 5
        }
        
        targetCell.values = [[newCodeString]];
        targetCell.format.font.name = "Calibri";
        targetCell.format.font.size = 12;
        
        await context.sync();
    });
}

function loadHistory() {
    const history = JSON.parse(localStorage.getItem('mz_code_history') || '[]');
    if (history.length === 0) {
        historyList.innerHTML = '<li class="empty-state">Nessun codice generato.</li>';
        return;
    }
    
    historyList.innerHTML = '';
    history.forEach(item => {
        const li = document.createElement('li');
        li.className = 'history-item';
        const date = new Date(item.timestamp).toLocaleDateString('it-IT', {
            hour: '2-digit', minute:'2-digit'
        });
        
        li.innerHTML = `
            <div>
                <div class="history-code">${item.code}</div>
                <div class="history-desc">${item.description} • ${date}</div>
            </div>
            <button class="history-copy" title="Copia" data-code="${item.code} - ${item.description !== 'Nessuna descrizione' ? item.description : ''}">
                <i class="fa-regular fa-copy"></i>
            </button>
        `;
        historyList.appendChild(li);
    });
    
    document.querySelectorAll('.history-copy').forEach(btn => {
        btn.addEventListener('click', function() {
            let text = this.getAttribute('data-code');
            if (text.endsWith(' - ')) text = text.replace(' - ', '');
            navigator.clipboard.writeText(text).then(() => {
                const icon = this.querySelector('i');
                icon.className = 'fa-solid fa-check text-success';
                setTimeout(() => icon.className = 'fa-regular fa-copy', 2000);
            });
        });
    });
}

clearHistoryBtn.addEventListener('click', () => {
    if (confirm('Sei sicuro di voler svuotare lo storico?')) {
        localStorage.removeItem('mz_code_history');
        loadHistory();
    }
});
