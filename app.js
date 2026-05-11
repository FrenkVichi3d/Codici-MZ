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
    } catch (e) {
        console.error("Init Error:", e);
    } finally {
        if (loadingOverlay) loadingOverlay.style.display = 'none';
        if (appMain) appMain.style.display = 'block';
    }
}

function extractProgressive(code_str) {
    if (!code_str) return 0;
    const match = String(code_str).match(/\.(\d{4})\b/);
    return match ? parseInt(match[1], 10) : 0;
}

async function fetchSchemaFromExcel() {
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const sheetItems = sheets.items;
            const rangeObjects = sheetItems.map(sheet => ({
                name: sheet.name,
                range: sheet.getUsedRangeOrNullObject()
            }));

            await context.sync();

            const activeRanges = [];
            rangeObjects.forEach(obj => {
                if (!obj.range.isNullObject) {
                    obj.range.load("values");
                    activeRanges.push(obj);
                }
            });

            await context.sync();

            currentSchema = [];
            window.allExistingItems = []; // Catalogo globale per la ricerca veloce

            activeRanges.forEach(obj => {
                const sheetName = obj.name;
                const values = obj.range.values;
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
                    const v3 = row.length > 3 && row[3] ? String(row[3]).trim() : "";
                    const v4 = row.length > 4 && row[4] ? String(row[4]).trim() : "";
                    const v5 = row.length > 5 && row[5] ? String(row[5]).trim() : "";

                    // Aggiunta al catalogo globale con posizione riga
                    let codeString = "";
                    if (is_4_39 && v3 && v3.includes('-')) codeString = v3;
                    else if (!is_4_39 && v5 && v5.includes('-')) codeString = v5;
                    
                    if (codeString) {
                        window.allExistingItems.push({
                            sheet: sheetName,
                            text: codeString,
                            rowIndex: obj.range.rowIndex + rowIdx,
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
    } catch (error) {
        console.error("Errore schema:", error);
    }
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
        return;
    }
    selectedSheet = currentSchema[index];
    populateLevel1();
});

level1Select.addEventListener('change', (e) => {
    const index = e.target.value;
    if (index === '') {
        selectedLevel1 = null;
        level2Select.innerHTML = '<option value="">Seleziona Livello 1...</option>';
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
    if (selectedLevel2) progressiveInput.value = (selectedLevel2.max_progressive || 0) + 1;
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
