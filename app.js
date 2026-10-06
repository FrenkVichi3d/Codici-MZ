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
const syncBtn = document.getElementById('syncBtn');
const historyList = document.getElementById('historyList');
const clearHistoryBtn = document.getElementById('clearHistoryBtn');
const appMain = document.getElementById('appMain');
const loadingOverlay = document.getElementById('loadingOverlay');

// Global Search DOM Elements
const globalSearchInput = document.getElementById('globalSearchInput');
const clearSearchBtn = document.getElementById('clearSearchBtn');
const searchResultsContainer = document.getElementById('searchResultsContainer');
const searchResultsList = document.getElementById('searchResultsList');
const searchCount = document.getElementById('searchCount');

// Presets & Smart Suggestion DOM Elements
const presetsList = document.getElementById('presetsList');
const addPresetBtn = document.getElementById('addPresetBtn');
const presetAddBox = document.getElementById('presetAddBox');
const presetNameInput = document.getElementById('presetNameInput');
const presetConfirmBtn = document.getElementById('presetConfirmBtn');
const presetCancelBtn = document.getElementById('presetCancelBtn');
const presetErrorMsg = document.getElementById('presetErrorMsg');
const smartSuggestionCard = document.getElementById('smartSuggestionCard');
const smartSuggestionPath = document.getElementById('smartSuggestionPath');
const smartSuggestionReason = document.getElementById('smartSuggestionReason');
const applySuggestionBtn = document.getElementById('applySuggestionBtn');
const dismissSuggestionBtn = document.getElementById('dismissSuggestionBtn');

// Office Initialization
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initApp();
    } else {
        if (loadingOverlay) loadingOverlay.style.display = 'none';
        if (appMain) appMain.style.display = 'block';
    }
});

async function initApp() {
    try {
        loadHistory();
        await fetchSchemaFromExcel();
        loadPresets();
        initSmartSuggestion();
    } catch (e) {
        // MOSTRA L'ERRORE DIRETTAMENTE NELL'INTERFACCIA PER DEBUG
        if (loadingOverlay) {
            loadingOverlay.innerHTML = `<i class="fa-solid fa-triangle-exclamation fa-3x" style="color:var(--danger);"></i>
                                       <p style="margin-top:15px; color:var(--danger); font-weight:bold;">Errore di Caricamento</p>
                                       <p style="font-size:0.75rem; color:var(--text-main); padding:0 10px;">${e.message}</p>
                                       <button onclick="location.reload()" style="margin-top:10px; padding:5px 10px; border-radius:5px; border:1px solid var(--border); cursor:pointer;">Riprova</button>`;
            loadingOverlay.style.display = 'block';
        }
        console.error("Init Error:", e);
    } finally {
        if (loadingOverlay && !loadingOverlay.innerHTML.includes('Errore')) {
            loadingOverlay.style.display = 'none';
        }
        if (appMain) appMain.style.display = 'block';
    }
}

function extractProgressive(code_str) {
    if (!code_str) return 0;
    const match = String(code_str).match(/\.(\d{4})\b/);
    return match ? parseInt(match[1], 10) : 0;
}

async function fetchSchemaFromExcel() {
    return Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        const sheetItems = sheets.items;
        const rangeObjects = [];
        
        // Prima fase: identifichiamo i range usati
        for (let i = 0; i < sheetItems.length; i++) {
            const s = sheetItems[i];
            rangeObjects.push({
                name: s.name,
                range: s.getUsedRangeOrNullObject()
            });
        }

        await context.sync();

        const activeRanges = [];
        rangeObjects.forEach(obj => {
            if (!obj.range.isNullObject) {
                // Carichiamo rowIndex e values in modo esplicito come stringa
                obj.range.load("values, rowIndex");
                activeRanges.push(obj);
            }
        });

        await context.sync();

        currentSchema = [];
        window.allExistingItems = [];

        activeRanges.forEach(obj => {
            const sheetName = obj.name;
            const values = obj.range.values;
            const startRow = obj.range.rowIndex;
            const sheetObj = { sheet: sheetName, categories: [] };
            const is_4_39 = sheetName.includes('4.39');

            let current_level1 = null;
            let current_level2 = null;

            if (is_4_39) {
                current_level1 = { name: 'ELETTRICI', code: '39', prefix: '4', subcategories: [], max_progressive: 0 };
                sheetObj.categories.push(current_level1);
            }

            for (let rowIdx = 0; rowIdx < values.length; rowIdx++) {
                const row = values[rowIdx];
                if (!row || row.length < 3) continue;

                const v0 = row[0] ? String(row[0]).trim() : "";
                const v1 = row[1] ? String(row[1]).trim() : "";
                const v2 = row[2] ? String(row[2]).trim() : "";
                const v3 = row.length > 3 ? String(row[3]).trim() : "";
                const v4 = row.length > 4 ? String(row[4]).trim() : "";
                const v5 = row.length > 5 ? String(row[5]).trim() : "";

                // POPOLAMENTO CATALOGO RICERCA
                // Riconosciamo un codice se inizia con un numero e contiene punti
                let possibleCode = is_4_39 ? v3 : v5;
                if (possibleCode && possibleCode.length > 5 && /^\d+\./.test(possibleCode)) {
                    window.allExistingItems.push({
                        sheet: sheetName,
                        text: possibleCode,
                        rowIndex: startRow + rowIdx,
                        colIndex: is_4_39 ? 3 : 5
                    });
                }

                if (is_4_39) {
                    if (v1 && v2 && v2 !== 'nan' && v2 !== 'N. PZ') {
                        let c2 = v2.replace('.0', '').padStart(2, '0');
                        current_level2 = { name: v1, code: c2, max_progressive: 0 };
                        current_level1.subcategories.push(current_level2);
                    }
                    if (v3) {
                        const prog = extractProgressive(v3);
                        if (current_level2 && prog > current_level2.max_progressive) current_level2.max_progressive = prog;
                        else if (current_level1 && prog > current_level1.max_progressive) current_level1.max_progressive = prog;
                    }
                } else {
                    if (v1 && v2 && v2 !== 'nan' && v2 !== 'N. PZ') {
                        let c1 = v2.replace('.0', '').padStart(2, '0');
                        let p = v0 ? v0.replace('.0', '').replace('.', '') : sheetName.split(' ')[0];
                        current_level1 = { name: v1, code: c1, prefix: p, subcategories: [], max_progressive: 0 };
                        sheetObj.categories.push(current_level1);
                        current_level2 = null;
                    }
                    if (v3 && v4 && v4 !== 'nan' && current_level1) {
                        let c2 = v4.replace('.0', '').padStart(2, '0');
                        current_level2 = { name: v3, code: c2, max_progressive: 0 };
                        current_level1.subcategories.push(current_level2);
                    }
                    if (v5) {
                        const prog = extractProgressive(v5);
                        if (current_level2 && prog > current_level2.max_progressive) current_level2.max_progressive = prog;
                        else if (current_level1 && prog > current_level1.max_progressive) current_level1.max_progressive = prog;
                    }
                }
            }
            if (sheetObj.categories.length > 0) currentSchema.push(sheetObj);
        });
        populateSheets();
    });
}

function populateSheets() {
    if (!sheetSelect) return;
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
        level1Select.innerHTML = '<option value="">Seleziona Categoria...</option>';
        level1Select.disabled = true;
        level2Select.innerHTML = '<option value="">Seleziona Livello 1...</option>';
        level2Select.disabled = true;
        progressiveInput.value = 1;
        highlightActivePreset();
        return;
    }
    selectedSheet = currentSchema[index];
    populateLevel1();
    highlightActivePreset();
});

level1Select.addEventListener('change', (e) => {
    const index = e.target.value;
    if (index === '') {
        selectedLevel1 = null;
        level2Select.innerHTML = '<option value="">Seleziona Livello 1...</option>';
        level2Select.disabled = true;
        progressiveInput.value = 1;
        highlightActivePreset();
        return;
    }
    selectedLevel1 = selectedSheet.categories[index];
    populateLevel2();
    highlightActivePreset();
});

level2Select.addEventListener('change', (e) => {
    const index = e.target.value;
    if (index === '') {
        selectedLevel2 = null;
        progressiveInput.value = (selectedLevel1.max_progressive || 0) + 1;
        highlightActivePreset();
        return;
    }
    selectedLevel2 = selectedLevel1.subcategories[index];
    if (selectedLevel2) progressiveInput.value = (selectedLevel2.max_progressive || 0) + 1;
    highlightActivePreset();
});

function populateLevel1() {
    level1Select.innerHTML = '<option value="">Seleziona...</option>';
    if (!selectedSheet || selectedSheet.categories.length === 0) {
        level1Select.disabled = true;
        return;
    }
    selectedSheet.categories.forEach((cat, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = `${cat.code} - ${cat.name}`;
        level1Select.appendChild(option);
    });
    level1Select.disabled = false;
    level2Select.disabled = true;
    progressiveInput.value = 1;
}

function populateLevel2() {
    level2Select.innerHTML = '<option value="">Seleziona (Opzionale)...</option>';
    if (!selectedLevel1.subcategories || selectedLevel1.subcategories.length === 0) {
        level2Select.disabled = true;
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
    progressiveInput.value = (selectedLevel1.max_progressive || 0) + 1;
}

syncBtn.addEventListener('click', async () => {
    const icon = syncBtn.querySelector('i');
    icon.classList.add('fa-spin');
    await fetchSchemaFromExcel();
    loadPresets();
    icon.classList.remove('fa-spin');
    icon.className = 'fa-solid fa-check text-success';
    setTimeout(() => { icon.className = 'fa-solid fa-rotate'; }, 2000);
});

generateBtn.addEventListener('click', async () => {
    if (!selectedSheet || !selectedLevel1) {
        alert('Seleziona i parametri.');
        return;
    }
    const description = descriptionInput.value.trim().toUpperCase();
    if (!description) {
        alert('Inserisci la descrizione.');
        return;
    }

    let prefix = selectedLevel1.prefix || selectedSheet.sheet.split(' ')[0].replace('.', '');
    const lvl1Code = selectedLevel1.code;
    const lvl2Code = selectedLevel2 ? selectedLevel2.code : '00';
    let progVal = parseInt(progressiveInput.value) || 1;
    const progCode = progVal.toString().padStart(4, '0');
    let finalCode = `${prefix}.${lvl1Code}.${lvl2Code}.${progCode} - ${description}`;

    // COPIA AUTOMATICA NEGLI APPUNTI (SINCRONA, PRIMA DELL'INSERIMENTO)
    // Questo garantisce che funzioni perché avviene subito dopo il click dell'utente
    try {
        const textArea = document.createElement("textarea");
        textArea.value = finalCode;
        textArea.style.position = "fixed";  // Nasconde la textarea
        textArea.style.left = "-9999px";
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
    } catch (e) {
        console.log("Copia automatica fallita:", e);
    }

    const originalContent = generateBtn.innerHTML;
    generateBtn.disabled = true;
    generateBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Inserimento...';

    try {
        await insertCodeIntoExcel(selectedSheet.sheet, selectedLevel1.code, lvl2Code, finalCode);

        if (selectedLevel2) selectedLevel2.max_progressive = progVal;
        else selectedLevel1.max_progressive = progVal;

        finalCodeEl.textContent = finalCode;
        codeBreakdownEl.innerHTML = `
            <div class="breakdown-item"><span class="breakdown-label">Prefisso</span><span class="breakdown-value">${prefix}</span></div>
            <div class="breakdown-item"><span class="breakdown-label">Livello 1</span><span class="breakdown-value">${lvl1Code}</span></div>
            <div class="breakdown-item"><span class="breakdown-label">Livello 2</span><span class="breakdown-value">${lvl2Code}</span></div>
            <div class="breakdown-item"><span class="breakdown-label">Progressivo</span><span class="breakdown-value">${progCode}</span></div>
        `;
        resultCard.style.display = 'block';
        resultCard.classList.add('success-border');

        const codeOnly = finalCode.split(' - ')[0];
        let history = JSON.parse(localStorage.getItem('mz_code_history') || '[]');
        history.unshift({ code: codeOnly, description: description, timestamp: new Date().toISOString() });
        localStorage.setItem('mz_code_history', JSON.stringify(history.slice(0, 50)));
        loadHistory();

        progressiveInput.value = progVal + 1;
        descriptionInput.value = '';
        hideSmartSuggestion();
        generateBtn.innerHTML = '<i class="fa-solid fa-check"></i> Copiato & Inserito!';
        generateBtn.classList.add('btn-success');

        setTimeout(() => {
            generateBtn.disabled = false;
            generateBtn.innerHTML = originalContent;
            generateBtn.classList.remove('btn-success');
        }, 3000);
    } catch (err) {
        alert('Errore: ' + err.message);
        generateBtn.disabled = false;
        generateBtn.innerHTML = originalContent;
    }
});

async function insertCodeIntoExcel(sheetName, level1Code, level2Code, newCodeString) {
    return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        sheet.activate();
        const usedRange = sheet.getUsedRange();
        usedRange.load(["values", "rowIndex"]);
        await context.sync();

        const values = usedRange.values;
        const startRowIdx = usedRange.rowIndex;
        const is_4_39 = sheetName.includes('4.39');
        let target_l1_found = false, target_l2_found = false, last_item_row_idx = -1;

        for (let i = 0; i < values.length; i++) {
            const row = values[i];
            const v1 = row[1] ? String(row[1]).trim() : "", v2 = row[2] ? String(row[2]).trim() : "";
            const v3 = row.length > 3 && row[3] ? String(row[3]).trim() : "";
            const v4 = row.length > 4 && row[4] ? String(row[4]).trim() : "", v5 = row.length > 5 && row[5] ? String(row[5]).trim() : "";

            if (is_4_39) {
                target_l1_found = true;
                if (v1 && v2 && v2 !== 'N. PZ') {
                    let c2 = v2.replace('.0', '').padStart(2, '0');
                    if (c2 === level2Code) { target_l2_found = true; last_item_row_idx = i; }
                    else if (target_l2_found) break;
                }
                if (target_l2_found && v3) last_item_row_idx = i;
            } else {
                if (v1 && v2 && v2 !== 'N. PZ') {
                    let c1 = v2.replace('.0', '').padStart(2, '0');
                    if (c1 === level1Code) { target_l1_found = true; if (level2Code === '00') last_item_row_idx = i; }
                    else if (target_l1_found) break;
                }
                if (target_l1_found && v3 && v4) {
                    let c2 = v4.replace('.0', '').padStart(2, '0');
                    if (c2 === level2Code) { target_l2_found = true; last_item_row_idx = i; }
                    else if (target_l2_found) break;
                }
                if (target_l2_found && v5) last_item_row_idx = i;
                if (target_l1_found && !target_l2_found && level2Code === '00' && v5) last_item_row_idx = i;
            }
        }
        if (!target_l1_found) throw new Error("Categoria non trovata.");
        let insertAbsoluteRow = (last_item_row_idx !== -1) ? startRowIdx + last_item_row_idx + 1 : startRowIdx + values.length;
        const rangeToInsert = sheet.getRangeByIndexes(insertAbsoluteRow, 0, 1, 10).getEntireRow();
        rangeToInsert.insert(Excel.InsertShiftDirection.down);
        let targetCell = sheet.getCell(insertAbsoluteRow, is_4_39 ? 3 : 5);
        targetCell.values = [[newCodeString]];
        targetCell.format.font.name = "Calibri";
        targetCell.format.font.size = 12;
        targetCell.select();
        await context.sync();
    });
}

copyBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(finalCodeEl.textContent).then(() => {
        const icon = copyBtn.querySelector('i');
        icon.className = 'fa-solid fa-check text-success';
        setTimeout(() => icon.className = 'fa-regular fa-copy', 2000);
    });
});

function loadHistory() {
    const history = JSON.parse(localStorage.getItem('mz_code_history') || '[]');
    if (!historyList) return;
    if (history.length === 0) {
        historyList.innerHTML = '<li class="empty-state">Nessun codice.</li>';
        return;
    }
    historyList.innerHTML = '';
    history.forEach(item => {
        const li = document.createElement('li');
        li.className = 'history-item';
        const date = new Date(item.timestamp).toLocaleDateString('it-IT', { hour: '2-digit', minute: '2-digit' });
        li.innerHTML = `<div><div class="history-code">${item.code}</div><div class="history-desc">${item.description} • ${date}</div></div>
            <button class="history-copy" title="Copia" data-code="${item.code} - ${item.description}"><i class="fa-regular fa-copy"></i></button>`;
        historyList.appendChild(li);
    });
    document.querySelectorAll('.history-copy').forEach(btn => {
        btn.addEventListener('click', function () {
            navigator.clipboard.writeText(this.getAttribute('data-code')).then(() => {
                const icon = this.querySelector('i');
                icon.className = 'fa-solid fa-check text-success';
                setTimeout(() => icon.className = 'fa-regular fa-copy', 2000);
            });
        });
    });
}

clearHistoryBtn.addEventListener('click', () => {
    if (confirm('Svuotare?')) { localStorage.removeItem('mz_code_history'); loadHistory(); }
});

// --- LOGICA MOTORE DI RICERCA GLOBALE ---

if (globalSearchInput) {
    globalSearchInput.addEventListener('input', (e) => {
        const term = e.target.value.trim().toLowerCase();
        
        // Gestione pulsante cancella
        if (term.length > 0) {
            clearSearchBtn.style.display = 'block';
        } else {
            clearSearchBtn.style.display = 'none';
            searchResultsContainer.style.display = 'none';
            return;
        }
        
        // Ricerca solo dopo 2 caratteri per performance
        if (term.length < 2) return;

        // Supporto ricerca multi-parola (es: "LAM SP4")
        const terms = term.split(' ').filter(t => t);
        
        const results = (window.allExistingItems || []).filter(item => {
            const textLower = item.text.toLowerCase();
            return terms.every(t => textLower.includes(t));
        });
        
        renderSearchResults(results.slice(0, 50)); // Limitiamo a 50 per fluidità
    });
}

if (clearSearchBtn) {
    clearSearchBtn.addEventListener('click', () => {
        globalSearchInput.value = '';
        globalSearchInput.dispatchEvent(new Event('input'));
    });
}

function renderSearchResults(results) {
    if (!searchResultsList || !searchResultsContainer || !searchCount) return;
    
    searchResultsList.innerHTML = '';
    
    if (results.length === 0) {
        searchResultsList.innerHTML = '<li class="empty-state">Nessun componente trovato.</li>';
        searchCount.textContent = '0 risultati';
    } else {
        searchCount.textContent = `${results.length} risultati trovati`;
        
        results.forEach(item => {
            const li = document.createElement('li');
            li.className = 'history-item';
            li.style.cursor = 'pointer'; // Indica che è cliccabile
            
            // Separiamo codice e descrizione per lo stile
            const parts = item.text.split(' - ');
            const code = parts[0];
            const desc = parts.slice(1).join(' - ');
            
            li.innerHTML = `
                <div class="result-info" style="flex-grow: 1; padding-right: 10px;">
                    <div class="history-code" style="font-size:0.85rem; color: var(--primary); font-weight: 700;">${code}</div>
                    <div class="history-desc" style="font-size:0.8rem; line-height: 1.2; margin-top: 2px;">${desc}</div>
                    <div style="font-size:0.65rem; color: var(--text-muted); margin-top: 4px; display: flex; align-items: center; gap: 4px;">
                        <i class="fa-solid fa-location-dot"></i> Vai a riga ${item.rowIndex + 1} • <strong>${item.sheet}</strong>
                    </div>
                </div>
                <button class="history-copy search-copy-btn" title="Copia" data-full="${item.text}">
                    <i class="fa-regular fa-copy"></i>
                </button>
            `;
            
            // Evento per andare alla cella cliccando sul corpo del risultato
            li.querySelector('.result-info').addEventListener('click', () => {
                goToExcelCell(item.sheet, item.rowIndex, item.colIndex);
            });

            searchResultsList.appendChild(li);
        });
        
        // Attiviamo i pulsanti di copia nei risultati
        searchResultsList.querySelectorAll('.search-copy-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                const textToCopy = this.getAttribute('data-full');
                copyToClipboard(textToCopy, this);
            });
        });
    }
    searchResultsContainer.style.display = 'block';
}

// Funzione di utilità per la copia (riutilizzabile)
function copyToClipboard(text, btnElement) {
    try {
        const textArea = document.createElement("textarea");
        textArea.value = text;
        textArea.style.position = "fixed";
        textArea.style.left = "-9999px";
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
        
        const icon = btnElement.querySelector('i');
        if (icon) {
            icon.className = 'fa-solid fa-check text-success';
            setTimeout(() => icon.className = 'fa-regular fa-copy', 2000);
        }
    } catch(e) {
        console.error("Errore copia:", e);
    }
}

// Nuova funzione per navigare verso una cella specifica
async function goToExcelCell(sheetName, rowIndex, colIndex) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const cell = sheet.getCell(rowIndex, colIndex);
            sheet.activate();
            cell.select();
            await context.sync();
        });
    } catch (error) {
        console.error("Errore navigazione cella:", error);
    }
}

// ==========================================
// SELEZIONE PROGRAMMATICA CATEGORIE
// ==========================================
function applyCategorySelectionByCodes(sheetName, lvl1Code, lvl2Code) {
    if (!currentSchema || currentSchema.length === 0 || !sheetName) return false;
    
    // Trova l'indice del foglio
    const sNameNorm = sheetName.trim().toUpperCase();
    const sheetIdx = currentSchema.findIndex(s => s.sheet.trim().toUpperCase() === sNameNorm);
    if (sheetIdx === -1) return false;
    
    sheetSelect.value = String(sheetIdx);
    selectedSheet = currentSchema[sheetIdx];
    populateLevel1();
    
    if (!lvl1Code) {
        highlightActivePreset();
        return true;
    }
    
    // Trova Livello 1
    const targetL1 = String(lvl1Code).padStart(2, '0');
    const lvl1Idx = selectedSheet.categories.findIndex(c => String(c.code).padStart(2, '0') === targetL1);
    if (lvl1Idx === -1) {
        highlightActivePreset();
        return true;
    }
    
    level1Select.value = String(lvl1Idx);
    selectedLevel1 = selectedSheet.categories[lvl1Idx];
    populateLevel2();
    
    // Trova Livello 2 (se presente)
    if (lvl2Code && lvl2Code !== '00' && selectedLevel1.subcategories && selectedLevel1.subcategories.length > 0) {
        const targetL2 = String(lvl2Code).padStart(2, '0');
        const lvl2Idx = selectedLevel1.subcategories.findIndex(sub => String(sub.code).padStart(2, '0') === targetL2);
        if (lvl2Idx !== -1) {
            level2Select.value = String(lvl2Idx);
            selectedLevel2 = selectedLevel1.subcategories[lvl2Idx];
            progressiveInput.value = (selectedLevel2.max_progressive || 0) + 1;
        } else {
            selectedLevel2 = null;
            progressiveInput.value = (selectedLevel1.max_progressive || 0) + 1;
        }
    } else {
        selectedLevel2 = null;
        progressiveInput.value = (selectedLevel1.max_progressive || 0) + 1;
    }
    
    highlightActivePreset();
    return true;
}

// ==========================================
// GESTIONE SCORCIATOIE RAPIDE (PRESET)
// ==========================================
let currentPresets = [];

function loadPresets() {
    try {
        const saved = localStorage.getItem('mz_code_presets');
        if (saved) {
            currentPresets = JSON.parse(saved);
        } else {
            currentPresets = generateDefaultPresets();
            localStorage.setItem('mz_code_presets', JSON.stringify(currentPresets));
        }
    } catch (e) {
        currentPresets = generateDefaultPresets();
    }
    renderPresets();
}

function generateDefaultPresets() {
    const defaults = [];
    if (!currentSchema) return defaults;

    currentSchema.forEach(sheet => {
        const sName = sheet.sheet.toUpperCase();
        if (sName.includes('5') || sName.includes('SEMILAVORATO')) {
            const cat = sheet.categories.find(c => String(c.code).padStart(2, '0') === '01');
            if (cat) {
                const sub = (cat.subcategories || []).find(s => String(s.code).padStart(2, '0') === '01');
                defaults.push({
                    id: 'p_semi_carp',
                    label: 'Carpenteria Acciaio',
                    sheetName: sheet.sheet,
                    lvl1Code: '01',
                    lvl2Code: sub ? '01' : '00'
                });
            }
        } else if (sName.includes('4.39') || sName.includes('ELETTRICI')) {
            const cat = sheet.categories.find(c => String(c.code).padStart(2, '0') === '39');
            if (cat) {
                defaults.push({
                    id: 'p_elec',
                    label: 'Elettrici MZ',
                    sheetName: sheet.sheet,
                    lvl1Code: '39',
                    lvl2Code: '00'
                });
            }
        } else if (sName.includes('4') && sName.includes('COMMERCIALE')) {
            if (sheet.categories.length > 0) {
                const cat = sheet.categories[0];
                defaults.push({
                    id: 'p_comm',
                    label: 'Commerciale',
                    sheetName: sheet.sheet,
                    lvl1Code: cat.code,
                    lvl2Code: '00'
                });
            }
        }
    });
    return defaults.slice(0, 4);
}

function renderPresets() {
    if (!presetsList) return;
    presetsList.innerHTML = '';
    
    if (currentPresets.length === 0) {
        presetsList.innerHTML = '<span class="presets-empty">Nessuna scorciatoia. Clicca su "+ Salva Preferito" per aggiungerne una.</span>';
        return;
    }
    
    currentPresets.forEach(preset => {
        const chip = document.createElement('div');
        chip.className = 'preset-chip';
        chip.setAttribute('data-id', preset.id);
        chip.innerHTML = `
            <span>${preset.label}</span>
            <button type="button" class="preset-delete" title="Rimuovi scorciatoia" data-id="${preset.id}">&times;</button>
        `;
        
        chip.addEventListener('click', (e) => {
            if (e.target.closest('.preset-delete')) return;
            applyCategorySelectionByCodes(preset.sheetName, preset.lvl1Code, preset.lvl2Code);
        });
        
        const delBtn = chip.querySelector('.preset-delete');
        delBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            removePreset(preset.id);
        });
        
        presetsList.appendChild(chip);
    });
    
    highlightActivePreset();
}

function removePreset(id) {
    currentPresets = currentPresets.filter(p => p.id !== id);
    localStorage.setItem('mz_code_presets', JSON.stringify(currentPresets));
    renderPresets();
}

function openPresetAddBox() {
    if (!presetAddBox || !presetNameInput) return;
    
    // Validazione: deve esserci almeno foglio e livello 1
    if (!selectedSheet || !selectedLevel1) {
        showPresetError('Seleziona prima una Categoria Principale e un Livello 1!');
        return;
    }
    
    // Genera nome suggerito chiaro e sintetico
    let defaultLabel = '';
    if (selectedLevel2) {
        defaultLabel = `${selectedLevel1.name.split(' ')[0]} / ${selectedLevel2.name}`;
    } else {
        defaultLabel = `${selectedSheet.sheet.split(' ')[0]} ${selectedLevel1.name}`;
    }
    if (defaultLabel.length > 25) defaultLabel = defaultLabel.substring(0, 24);
    
    presetNameInput.value = defaultLabel;
    if (presetErrorMsg) presetErrorMsg.style.display = 'none';
    presetAddBox.style.display = 'flex';
    presetNameInput.focus();
    presetNameInput.select();
}

function closePresetAddBox() {
    if (presetAddBox) presetAddBox.style.display = 'none';
    if (presetErrorMsg) presetErrorMsg.style.display = 'none';
    if (presetNameInput) presetNameInput.value = '';
}

function showPresetError(msg) {
    if (!presetErrorMsg) return;
    presetErrorMsg.textContent = msg;
    presetErrorMsg.style.display = 'block';
    setTimeout(() => {
        if (presetErrorMsg) presetErrorMsg.style.display = 'none';
    }, 3500);
}

function confirmPresetAdd() {
    if (!selectedSheet || !selectedLevel1) {
        showPresetError('Seleziona prima Categoria e Livello 1!');
        return;
    }
    
    const label = presetNameInput ? presetNameInput.value.trim() : '';
    if (!label) {
        showPresetError('Inserisci un nome per la scorciatoia!');
        if (presetNameInput) presetNameInput.focus();
        return;
    }
    
    const sheetName = selectedSheet.sheet;
    const lvl1Code = selectedLevel1.code;
    const lvl2Code = selectedLevel2 ? selectedLevel2.code : '00';
    
    const newPreset = {
        id: 'p_' + Date.now(),
        label: label,
        sheetName: sheetName,
        lvl1Code: lvl1Code,
        lvl2Code: lvl2Code
    };
    
    currentPresets.push(newPreset);
    try {
        localStorage.setItem('mz_code_presets', JSON.stringify(currentPresets));
    } catch (e) {
        console.error("Errore salvataggio preferiti:", e);
    }
    
    closePresetAddBox();
    renderPresets();
}

if (addPresetBtn) {
    addPresetBtn.addEventListener('click', () => {
        if (presetAddBox && presetAddBox.style.display === 'flex') {
            closePresetAddBox();
        } else {
            openPresetAddBox();
        }
    });
}

if (presetConfirmBtn) {
    presetConfirmBtn.addEventListener('click', () => {
        confirmPresetAdd();
    });
}

if (presetCancelBtn) {
    presetCancelBtn.addEventListener('click', () => {
        closePresetAddBox();
    });
}

if (presetNameInput) {
    presetNameInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            confirmPresetAdd();
        } else if (e.key === 'Escape') {
            closePresetAddBox();
        }
    });
}

function highlightActivePreset() {
    if (!presetsList || !selectedSheet || !selectedLevel1) {
        if (presetsList) {
            presetsList.querySelectorAll('.preset-chip').forEach(c => c.classList.remove('active'));
        }
        return;
    }
    const sName = selectedSheet.sheet.trim().toUpperCase();
    const l1 = String(selectedLevel1.code).padStart(2, '0');
    const l2 = selectedLevel2 ? String(selectedLevel2.code).padStart(2, '0') : '00';
    
    presetsList.querySelectorAll('.preset-chip').forEach(chip => {
        const id = chip.getAttribute('data-id');
        const p = currentPresets.find(item => item.id === id);
        if (p && p.sheetName.trim().toUpperCase() === sName && 
            String(p.lvl1Code).padStart(2, '0') === l1 && 
            (String(p.lvl2Code).padStart(2, '0') === l2 || (p.lvl2Code === '00' && !selectedLevel2))) {
            chip.classList.add('active');
        } else {
            chip.classList.remove('active');
        }
    });
}

// ==========================================
// SUGGERITORE INTELLIGENTE DALLA DESCRIZIONE
// ==========================================
let currentSuggestedCategory = null;
let suggestionDebounceTimer = null;

function initSmartSuggestion() {
    if (!descriptionInput) return;
    
    descriptionInput.addEventListener('input', (e) => {
        clearTimeout(suggestionDebounceTimer);
        const text = e.target.value.trim();
        if (text.length < 3) {
            hideSmartSuggestion();
            return;
        }
        suggestionDebounceTimer = setTimeout(() => {
            analyzeDescriptionForSuggestion(text);
        }, 180);
    });
    
    if (applySuggestionBtn) {
        applySuggestionBtn.addEventListener('click', () => {
            if (!currentSuggestedCategory) return;
            const ok = applyCategorySelectionByCodes(
                currentSuggestedCategory.sheetName,
                currentSuggestedCategory.lvl1Code,
                currentSuggestedCategory.lvl2Code
            );
            if (ok) {
                hideSmartSuggestion();
            }
        });
    }
    
    if (dismissSuggestionBtn) {
        dismissSuggestionBtn.addEventListener('click', () => {
            hideSmartSuggestion();
        });
    }
}

function hideSmartSuggestion() {
    currentSuggestedCategory = null;
    if (smartSuggestionCard) {
        smartSuggestionCard.style.display = 'none';
    }
}

function analyzeDescriptionForSuggestion(text) {
    if (!smartSuggestionCard || !window.allExistingItems || window.allExistingItems.length === 0) {
        hideSmartSuggestion();
        return;
    }
    
    const rawTokens = text.toUpperCase().split(/[\s,./\-_+]+/).filter(t => t.length >= 2);
    const stopWords = new Set(["PER", "CON", "DEL", "DEI", "DELLA", "DELLE", "NEL", "SUL", "ALLA", "ALLE", "NON"]);
    const tokens = rawTokens.filter(t => !stopWords.has(t));
    if (tokens.length === 0) {
        hideSmartSuggestion();
        return;
    }
    
    const candidateMap = new Map();
    
    // 1. Confronto con i componenti storici registrati
    for (let i = 0; i < window.allExistingItems.length; i++) {
        const item = window.allExistingItems[i];
        const itemText = item.text.toUpperCase();
        
        const parts = itemText.split(' - ');
        if (parts.length < 2) continue;
        const codePart = parts[0].trim();
        const descPart = parts.slice(1).join(' - ');
        
        const codeSegments = codePart.split('.');
        if (codeSegments.length < 4) continue;
        
        const lvl1Code = codeSegments[1];
        const lvl2Code = codeSegments[2];
        const sheetName = item.sheet;
        
        let matchScore = 0;
        for (const token of tokens) {
            if (descPart.includes(token)) {
                matchScore += 1;
                const reg = new RegExp('\\b' + token + '\\b');
                if (reg.test(descPart)) matchScore += 1.5;
            }
        }
        
        if (matchScore > 0) {
            const key = `${sheetName}:::${lvl1Code}:::${lvl2Code}`;
            if (!candidateMap.has(key)) {
                candidateMap.set(key, {
                    sheetName,
                    lvl1Code,
                    lvl2Code,
                    score: 0,
                    sampleDesc: descPart,
                    matchCount: 0
                });
            }
            const cand = candidateMap.get(key);
            cand.score += matchScore;
            cand.matchCount += 1;
        }
    }
    
    // 2. Bonus diretto da nomi categorie e sottocategorie
    if (currentSchema) {
        currentSchema.forEach(sheet => {
            const sName = sheet.sheet;
            sheet.categories.forEach(cat => {
                const catName = cat.name.toUpperCase();
                tokens.forEach(token => {
                    if (catName.includes(token)) {
                        const key = `${sName}:::${cat.code}:::00`;
                        if (!candidateMap.has(key)) {
                            candidateMap.set(key, { sheetName: sName, lvl1Code: cat.code, lvl2Code: '00', score: 0, sampleDesc: catName, matchCount: 0 });
                        }
                        candidateMap.get(key).score += 4;
                    }
                });
                
                (cat.subcategories || []).forEach(sub => {
                    const subName = sub.name.toUpperCase();
                    tokens.forEach(token => {
                        if (subName.includes(token)) {
                            const key = `${sName}:::${cat.code}:::${sub.code}`;
                            if (!candidateMap.has(key)) {
                                candidateMap.set(key, { sheetName: sName, lvl1Code: cat.code, lvl2Code: sub.code, score: 0, sampleDesc: subName, matchCount: 0 });
                            }
                            candidateMap.get(key).score += 5;
                        }
                    });
                });
            });
        });
    }
    
    if (candidateMap.size === 0) {
        hideSmartSuggestion();
        return;
    }
    
    let bestCandidate = null;
    let maxScore = 0;
    for (const cand of candidateMap.values()) {
        if (cand.score > maxScore) {
            maxScore = cand.score;
            bestCandidate = cand;
        }
    }
    
    if (!bestCandidate || maxScore < 2.5) {
        hideSmartSuggestion();
        return;
    }
    
    // Se la categoria corrente è già identica, non disturbare l'utente
    if (selectedSheet && selectedLevel1) {
        const curSheet = selectedSheet.sheet.trim().toUpperCase();
        const curL1 = String(selectedLevel1.code).padStart(2, '0');
        const curL2 = selectedLevel2 ? String(selectedLevel2.code).padStart(2, '0') : '00';
        if (curSheet === bestCandidate.sheetName.trim().toUpperCase() && 
            curL1 === String(bestCandidate.lvl1Code).padStart(2, '0') && 
            curL2 === String(bestCandidate.lvl2Code).padStart(2, '0')) {
            hideSmartSuggestion();
            return;
        }
    }
    
    // Ricava i nomi leggibili per la barra di suggerimento
    const sheetObj = currentSchema.find(s => s.sheet.trim().toUpperCase() === bestCandidate.sheetName.trim().toUpperCase());
    if (!sheetObj) {
        hideSmartSuggestion();
        return;
    }
    const catObj = sheetObj.categories.find(c => String(c.code).padStart(2, '0') === String(bestCandidate.lvl1Code).padStart(2, '0'));
    if (!catObj) {
        hideSmartSuggestion();
        return;
    }
    let subObj = null;
    if (bestCandidate.lvl2Code && bestCandidate.lvl2Code !== '00' && catObj.subcategories) {
        subObj = catObj.subcategories.find(s => String(s.code).padStart(2, '0') === String(bestCandidate.lvl2Code).padStart(2, '0'));
    }
    
    currentSuggestedCategory = bestCandidate;
    
    let path = `<strong>${sheetObj.sheet}</strong> › ${catObj.code} ${catObj.name}`;
    if (subObj) {
        path += ` › ${subObj.code} ${subObj.name}`;
    }
    smartSuggestionPath.innerHTML = path;
    
    let sample = bestCandidate.sampleDesc;
    if (sample && sample.length > 32) sample = sample.substring(0, 30) + '...';
    smartSuggestionReason.innerHTML = sample ? `Simile a: <em>"${sample}"</em>` : `Rilevato da parole chiave`;
    
    smartSuggestionCard.style.display = 'block';
}
