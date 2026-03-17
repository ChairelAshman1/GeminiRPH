// DOM Elements Helpers
const $ = (id) => document.getElementById(id);

// State
let dskpBuffer = [];
let rptBuffer = [];
let DATA_STORE = {}; // Structure: { "Tingkatan 1": { "Sejarah": [ {week, theme...} ] } }
let API_KEY = ""; // Holds the currently active key
let API_KEYS_LIST = []; // Holds all available keys
let CURRENT_MODE = "WEB"; // WEB, API
const DB_NAME = 'RPH_DB';
const STORE_NAME = 'data_store';
const DB_VERSION = 3; 

// Auto-Gen Timer
let autoGenTimer = null;
let currentAbortController = null; // Control cancellation

// --- INIT ---
document.addEventListener('DOMContentLoaded', async () => {
    
    // Load Settings & Keys
    initKeySystem();

    // Load API Usage Count
    updateAPIUsageDisplay();

    // Determine Initial Mode
    initSystemMode();

    // Restore Data Store from DB
    try {
        const savedData = await loadDataFromDB();
        if (savedData) {
            DATA_STORE = savedData;
            renderDataResults();
        }
    } catch(e) {
        console.error("Failed to load data from DB", e);
    }

    // --- ATTACH LISTENERS ---
    attachListeners();
    
    // Listen for Copy Events from Content Script
    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
        if (request.action === "COPIED_TEXT_DETECTED") {
            handleCopiedText(request.text);
        }
        else if (request.action === "COPY_TRIGGERED") {
            // Content script detected a copy, but couldn't get text (e.g. Google Sheets)
            // We read from clipboard here since we have permission
            setTimeout(async () => {
                try {
                    const text = await navigator.clipboard.readText();
                    if (text) handleCopiedText(text);
                } catch (e) {
                    console.error("Clipboard read failed", e);
                }
            }, 100);
        } else if (request.action === "MACRO_COMPLETED") {
            showToast("BERJAYA: Sistem selesai menjalankan tugas!", "success");
            resetCopyButtons();
        }
    });
});

function handleCopiedText(text) {
    if (!text || text.length < 3) return;
    
    // Regex for Standard Pembelajaran (e.g., "1.1.1 Text") -> 3 levels
    const spRegex = /^\d+\.\d+\.\d+\s+/;
    // Regex for Standard Kandungan (e.g., "1.1 Text") -> 2 levels
    const skRegex = /^\d+\.\d+\s+/;
    
    let isPasted = false;

    if (spRegex.test(text)) {
        // It is SP
        if (!$('standardInputs').classList.contains('hide')) {
            toggleInputType('inputSP', 'text');
            const el = $('inputSP');
            el.value = text;
            // Flash effect
            el.style.borderColor = '#4ade80';
            setTimeout(() => el.style.borderColor = '#334155', 500);
            showToast("SP dikesan & ditampal!", "success");
            isPasted = true;
        }
    } else if (skRegex.test(text)) {
        // It is SK (matches 1.1 but not 1.1.1 check above)
         if (!$('standardInputs').classList.contains('hide')) {
            toggleInputType('inputSK', 'text');
             const el = $('inputSK');
            el.value = text;
            // Flash effect
            el.style.borderColor = '#4ade80';
            setTimeout(() => el.style.borderColor = '#334155', 500);
            showToast("SK dikesan & ditampal!", "success");
            isPasted = true;
        }
    }

    // Auto Click Logic for Web Mode
    if (isPasted && CURRENT_MODE === 'WEB') {
        if (autoGenTimer) clearTimeout(autoGenTimer);
        
        // Update toast
        const mb = $('messageBar'); 
        mb.textContent = "Auto-Jana (Web) dalam 1s..."; 
        mb.className = "show info";
        
        autoGenTimer = setTimeout(() => {
            const btn = $('apiGenerateBtn');
            if (btn) btn.click();
        }, 1000);
    }
}

function initKeySystem() {
    // 1. Load active key (Legacy support)
    API_KEY = localStorage.getItem("GEMINI_API_KEY") || "";
    
    // 2. Load key list
    try {
        const storedList = localStorage.getItem("GEMINI_API_KEYS_LIST");
        API_KEYS_LIST = storedList ? JSON.parse(storedList) : [];
    } catch(e) {
        API_KEYS_LIST = [];
    }

    // 3. Migration: If API_KEY exists but not in list, add it
    if (API_KEY && API_KEY.length > 5) {
        if (!API_KEYS_LIST.includes(API_KEY)) {
            API_KEYS_LIST.push(API_KEY);
            saveKeyList();
        }
    } else if (API_KEYS_LIST.length > 0 && !API_KEY) {
        // If no active key but list exists, pick first one
        selectKey(API_KEYS_LIST[0], false); // Don't reset usage on init
    }

    renderKeyList();
}

function saveKeyList() {
    localStorage.setItem("GEMINI_API_KEYS_LIST", JSON.stringify(API_KEYS_LIST));
}

function selectKey(key, resetUsage = true) {
    API_KEY = key;
    localStorage.setItem("GEMINI_API_KEY", key);
    
    // Auto switch to API mode if valid key selected
    if(key.length > 5) {
        CURRENT_MODE = 'API';
        localStorage.setItem("PREFERRED_MODE", 'API');
    }

    if (resetUsage) {
        resetApiUsage();
    }
    
    updateStatusUI();
    renderKeyList();
}

function renderKeyList() {
    const listContainer = $('keyList');
    listContainer.innerHTML = "";

    if (API_KEYS_LIST.length === 0) {
        listContainer.innerHTML = '<div style="font-size:0.7rem; color:#64748b; text-align:center; padding:0.5rem;">Tiada Key Disimpan.</div>';
        return;
    }
    
    // Color Palette mapping logic is implicitly handled via CSS classes .color-0 to .color-4
    // We cycle through these based on index.

    API_KEYS_LIST.forEach((key, index) => {
        const isActive = (key === API_KEY);
        const colorClass = `color-${index % 5}`;
        
        const item = document.createElement('div');
        item.className = `key-item ${isActive ? 'active' : ''}`;
        
        // Mask: Show only last 4 chars
        const visiblePart = key.slice(-4);
        const mask = document.createElement('div');
        mask.className = `key-mask ${colorClass}`; // Apply distinct color to text
        mask.textContent = `Key ${index + 1}: •••••${visiblePart}`;
        
        // Checkbox Indicator
        const check = document.createElement('div');
        check.className = 'key-check';

        // Delete Button
        const delBtn = document.createElement('button');
        delBtn.className = 'btn-icon-small';
        delBtn.textContent = '✕';
        delBtn.title = "Padam Key";
        delBtn.onclick = (e) => {
            e.stopPropagation();
            if(confirm("Padam API Key ini?")) {
                API_KEYS_LIST.splice(index, 1);
                saveKeyList();
                // If deleted active key, fallback
                if (key === API_KEY) {
                    API_KEY = "";
                    localStorage.removeItem("GEMINI_API_KEY");
                    if(API_KEYS_LIST.length > 0) selectKey(API_KEYS_LIST[0]);
                    else {
                        CURRENT_MODE = "WEB";
                        updateStatusUI();
                    }
                }
                renderKeyList();
            }
        };

        item.onclick = () => {
            if (!isActive) {
                selectKey(key, true);
                showToast(`Key ${index + 1} diaktifkan. Penggunaan direset.`, "success");
            }
        };

        item.append(check, mask, delBtn);
        listContainer.appendChild(item);
    });
}

function initSystemMode() {
    // Initial Mode Logic
    const savedMode = localStorage.getItem("PREFERRED_MODE");
    
    if (savedMode === 'API' && API_KEY && API_KEY.length > 5) {
        CURRENT_MODE = 'API';
    } else if (savedMode === 'WEB') {
        CURRENT_MODE = 'WEB';
    } else {
        // Default behavior: Prefer API if key exists, else Web
        CURRENT_MODE = (API_KEY && API_KEY.length > 5) ? 'API' : 'WEB';
    }
    
    updateStatusUI(); 
}

function updateStatusUI() {
    const badge = $('statusBadge');
    const text = $('statusText');
    const dot = badge.querySelector('.status-dot');

    // Reset classes and inline styles
    badge.className = 'status-badge';
    text.style.color = ""; // Reset custom color

    if (CURRENT_MODE === 'API') {
        badge.classList.add('mode-api', 'active');
        text.textContent = "API";
        dot.style.backgroundColor = "#4ade80"; // Green
        
        // --- CUSTOM COLOR LOGIC FOR API MODE ---
        // Find index of current API_KEY to match the color from the list
        const keyIndex = API_KEYS_LIST.indexOf(API_KEY);
        if (keyIndex !== -1) {
            // Map index to color code manually to sync with CSS
            const colors = ["#22d3ee", "#f472b6", "#fbbf24", "#a78bfa", "#f87171"];
            const activeColor = colors[keyIndex % 5];
            text.style.color = activeColor;
            text.style.fontWeight = "800";
            text.style.textShadow = "0 0 5px rgba(0,0,0,0.5)";
        }
    } else {
        badge.classList.add('mode-web', 'active');
        text.textContent = "Web";
        dot.style.backgroundColor = "#4ade80";
        text.style.color = "#93c5fd"; // Default blueish for Web
    }
    
    // Toggle Web Scrape Button based on mode
    const webBtn = $('btnWebScrape');
    if (CURRENT_MODE === 'WEB') {
        webBtn.classList.remove('hide');
    } else {
        webBtn.classList.add('hide');
    }
}

function attachListeners() {
    
    // API Key Toggle
    $('btnApiKeyToggle').addEventListener('click', () => {
        const sec = $('apiKeySection');
        sec.classList.toggle('show');
        if(sec.classList.contains('show')) $('apiKeyInput').focus();
    });

    // Add Key Button
    $('btnAddKey').addEventListener('click', () => {
        const val = $('apiKeyInput').value.trim();
        if (!val || val.length < 10) {
            showToast("Key tidak sah atau terlalu pendek.", "error");
            return;
        }

        if (API_KEYS_LIST.includes(val)) {
            showToast("Key ini sudah wujud.", "error");
            return;
        }

        // Add key
        API_KEYS_LIST.push(val);
        saveKeyList();
        $('apiKeyInput').value = ""; // Clear input
        
        // Automatically select the new key
        selectKey(val, true); 
        
        showToast("Key Baru ditambah & diaktifkan!", "success");
    });

    // Help Button
    $('btnShowHelp').addEventListener('click', () => $('helpModal').classList.add('show'));
    $('btnCloseHelp').addEventListener('click', () => $('helpModal').classList.remove('show'));
    
    // Mode Selection Logic
    $('statusBadge').addEventListener('click', () => {
        const modal = $('modeModal');
        const btnAPI = $('btnSelectAPI');
        
        // Check API availability
        if(!API_KEY || API_KEY.length < 5) {
            btnAPI.classList.add('btn-disabled');
            btnAPI.title = "Sila masukkan API Key dahulu.";
        } else {
            btnAPI.classList.remove('btn-disabled');
            btnAPI.title = "";
        }
        
        modal.classList.add('show');
    });
    
    $('btnCloseMode').addEventListener('click', () => $('modeModal').classList.remove('show'));
    
    $('btnSelectAPI').addEventListener('click', () => {
        if(!API_KEY || API_KEY.length < 5) {
            showToast("Sila masukkan API Key dalam tetapan dahulu.", "error");
            return;
        }
        CURRENT_MODE = 'API';
        localStorage.setItem("PREFERRED_MODE", 'API');
        updateStatusUI();
        $('modeModal').classList.remove('show');
        showToast("Mode ditukar kepada API.", "success");
    });
    
    $('btnSelectWeb').addEventListener('click', () => {
        CURRENT_MODE = 'WEB';
        localStorage.setItem("PREFERRED_MODE", 'WEB');
        updateStatusUI();
        $('modeModal').classList.remove('show');
        showToast("Mode ditukar kepada Web.", "info");
    });

    // Cancel Loading Button
    $('btnCancelLoading').addEventListener('click', () => {
        if (currentAbortController) {
            currentAbortController.abort();
            currentAbortController = null;
        }
        showLoading(false);
        showToast("Proses dibatalkan.", "info");
    });

    // MODE SWITCHING LISTENER
    $('rphMode').addEventListener('change', (e) => {
        const mode = e.target.value;
        
        // Hide all Step 2 containers first
        $('standardInputs').classList.add('hide');
        $('examInputs').classList.add('hide');
        $('specialInputs').classList.add('hide');

        // Reset Inputs to default (Text) when mode changes
        toggleInputType('inputTema', 'text');
        toggleInputType('inputTajuk', 'text');
        toggleInputType('inputSK', 'text');
        toggleInputType('inputSP', 'text');

        // Visibility logic based on mode
        if (mode === 'special') {
            $('specialInputs').classList.remove('hide');
        } else if (mode === 'exam') {
            $('examInputs').classList.remove('hide');
        } else {
            $('standardInputs').classList.remove('hide');
        }
    });

    // FIELD EDIT BUTTONS LISTENERS
    $('btnEditTema').addEventListener('click', () => toggleInputType('inputTema', 'text'));
    $('btnEditTajuk').addEventListener('click', () => toggleInputType('inputTajuk', 'text'));
    $('btnEditSK').addEventListener('click', () => toggleInputType('inputSK', 'text'));
    $('btnEditSP').addEventListener('click', () => toggleInputType('inputSP', 'text'));

    // Clear Outputs
    $('btnClearOutputs').addEventListener('click', () => {
        const emptyState = '<div style="text-align:center; color:#94a3b8; font-size:0.75rem; padding: 1rem;">Data akan muncul di sini...</div>';
        $('listObjektif').innerHTML = emptyState;
        $('listKriteria').innerHTML = emptyState;
        $('listAktiviti').innerHTML = emptyState;
        showToast("Semua medan dikosongkan.");
    });

    // Clear Step 2 Inputs
    $('btnClearStep2').addEventListener('click', () => {
        // Reset Standard Inputs
        toggleInputType('inputTema', 'text', "");
        toggleInputType('inputTajuk', 'text', "");
        toggleInputType('inputSK', 'text', "");
        toggleInputType('inputSP', 'text', "");
        $('rptRawData').value = "";
        
        // Reset Exam/Special
        $('inputExamSubject').value = "";
        $('inputExamType').value = "";
        $('specialSubject').value = "";
        $('specialForm').value = "";
        $('specialActivity').value = "";
        
        showToast("Langkah 2 dikosongkan.");
    });

    // Paste All
    document.querySelectorAll('.btn-paste-all').forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.dataset.target;
            const container = $(targetId);
            const texts = [];
            container.querySelectorAll('textarea').forEach(inp => {
                const val = inp.value.trim();
                if(val && val.length > 2) texts.push(val);
            });
            
            if(texts.length > 0) inject(texts.join('\n'));
            else showToast("Tiada data untuk disalin", "error");
        });
    });

    // Tabs
    document.querySelectorAll('.tab-btn').forEach(b => {
        b.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            b.classList.add('active');
            $(`tab-${b.dataset.tab}`).classList.add('active');
        });
    });
    
    // Toggle Upload Section
    $('btnToggleUpload').addEventListener('click', () => {
        const sec = $('dataUploadSection');
        if (sec.classList.contains('expanded')) {
            sec.classList.remove('expanded');
        } else {
            sec.classList.add('expanded');
        }
    });

    // File Input Handlers
    $('dskpFileInput').addEventListener('change', (e) => handleFileUpload(e.target.files, 'DSKP'));
    $('rptFileInput').addEventListener('change', (e) => handleFileUpload(e.target.files, 'RPT'));

    async function handleFileUpload(files, type) {
        if (files.length === 0) return;
        showLoading(true, `Membaca ${type}...`);
        
        for (const file of Array.from(files)) {
            try {
                let dataContent = "";
                if (file.type === "application/pdf" || file.name.toLowerCase().endsWith('.pdf')) {
                    dataContent = await readFileAsBase64(file);
                } else {
                    dataContent = await readFileAsText(file);
                }
                
                const fileObj = { name: file.name, data: dataContent, type: file.type };
                if (type === 'DSKP') dskpBuffer.push(fileObj);
                else rptBuffer.push(fileObj);
                
            } catch(e) { console.error(e); }
        }
        showLoading(false);
        showToast(`${type} ditambah.`, "success");
    }

    $('btnClearDataStore').addEventListener('click', async () => {
        if(confirm("Padam semua data analisis?")) {
            DATA_STORE = {};
            dskpBuffer = [];
            rptBuffer = [];
            await saveDataToDB(DATA_STORE);
            renderDataResults();
            showToast("Data dipadam.", "info");
        }
    });

    // --- SMART LEARNING LOGIC ---
    $('btnLearnData').addEventListener('click', async () => {
        if (dskpBuffer.length === 0 && rptBuffer.length === 0) {
            return showToast("Sila muat naik DSKP atau RPT.", "error");
        }
        await processSmartLearning();
    });

    // Logic 1: Search (Uses Local DATA_STORE)
    $('btnAnalyzeRPT').addEventListener('click', () => {
        const week = $('inputRPTWeek').value.trim();
        const topic = $('inputRPTTopic').value.trim();
        
        if (!week && !topic) return showToast("Isi Minggu atau Kata Kunci", "error");
        performDeepSearch(week, topic);
    });
    
    // Button: Web Scrape (Only for Web Mode) - IMPROVED SMART FIND
    $('btnWebScrape').addEventListener('click', async () => {
        showLoading(true, "Mencari data di Gemini...");
        
        // 1. Find any Gemini tab (Active or Background)
        chrome.tabs.query({ url: "*://gemini.google.com/*" }, tabs => {
            let targetTab = null;
            
            // Prefer active Gemini tab
            const active = tabs.find(t => t.active);
            if(active) targetTab = active;
            else if(tabs.length > 0) targetTab = tabs[0]; // Or fallback to first found
            
            if(!targetTab) {
                showLoading(false);
                showToast("Sila buka Gemini (gemini.google.com)", "error");
                return;
            }

            // 2. Extract data from that tab
            chrome.tabs.sendMessage(targetTab.id, { action: "GET_GEMINI_RESPONSE" }, response => {
                showLoading(false);
                
                if (chrome.runtime.lastError) {
                    console.warn(chrome.runtime.lastError);
                    showToast("Sila refresh tab Gemini.", "error");
                    return;
                }

                if (response && response.text) {
                    renderOutputs(response.text);
                    showToast("Data diambil dari Gemini!", "success");
                } else {
                    showToast("Tiada data RPH dijumpai di Gemini.", "error");
                }
            });
        });
    });

    // Logic 2: Generate RPH (Uses AI)
    $('apiGenerateBtn').addEventListener('click', async () => {
        const mode = $('rphMode').value;
        let inputs = getInputs(mode);
        if (!inputs) return;
        const prompt = buildRPHPrompt(mode, inputs);
        smartGenerate(prompt, 'generate', inputs); // Pass inputs for web fallback
    });

    // --- SMART EXPORT (PASTE ALL) ---
    $('btnExportHelp').addEventListener('click', () => $('exportHelpModal').classList.add('show'));
    $('btnCloseExportHelp').addEventListener('click', () => $('exportHelpModal').classList.remove('show'));

    // Google Sheet / Apps Script Integration
    function getSheetAppScriptUrl() {
        return $('sheetAppScriptUrl')?.value.trim();
    }

    function getSheetId() {
        return $('sheetIdInput')?.value.trim();
    }

    function getSheetName() {
        return $('sheetNameInput')?.value.trim() || 'Sheet1';
    }

    function getSheetRange() {
        return $('sheetRangeInput')?.value.trim() || 'A2';
    }

    function setSheetStatus(message, isSuccess = true) {
        const status = $('sheetStatus');
        if (!status) return;
        status.textContent = `Status: ${message}`;
        status.style.color = isSuccess ? '#86efac' : '#fca5a5';
    }

    function saveSheetSettings() {
        localStorage.setItem('SHEET_WEBAPP_URL', getSheetAppScriptUrl() || '');
        localStorage.setItem('SHEET_ID', getSheetId() || '');
        localStorage.setItem('SHEET_NAME', getSheetName());
        localStorage.setItem('SHEET_RANGE', getSheetRange());
    }

    function loadSheetSettings() {
        const url = localStorage.getItem('SHEET_WEBAPP_URL') || '';
        const sheetId = localStorage.getItem('SHEET_ID') || '';
        const name = localStorage.getItem('SHEET_NAME') || 'Sheet1';
        const range = localStorage.getItem('SHEET_RANGE') || 'A2';
        if ($('sheetAppScriptUrl')) $('sheetAppScriptUrl').value = url;
        if ($('sheetIdInput')) $('sheetIdInput').value = sheetId;
        if ($('sheetNameInput')) $('sheetNameInput').value = name;
        if ($('sheetRangeInput')) $('sheetRangeInput').value = range;
        if (url) setSheetStatus('URL disimpan. Sila sambung.', true);
    }

    async function connectSheet() {
        const url = getSheetAppScriptUrl();
        if (!url) {
            setSheetStatus('Sila masukkan Apps Script URL.', false);
            return;
        }

        saveSheetSettings();
        setSheetStatus('Menyambung...', true);

        try {
            const f = await fetch(`${url}?action=ping`, { method: 'GET', cache: 'no-cache' });
            const j = await f.json();
            if (j && j.status === 'ok') {
                setSheetStatus('Sambungan berjaya: ' + (j.message || 'Bersedia'), true);
            } else {
                setSheetStatus('Sambungan gagal: ' + (j.message || 'tidak tersedia'), false);
            }
        } catch (error) {
            setSheetStatus('Sambungan gagal: ' + error.message, false);
        }
    }

    async function pushDataToSheet() {
        const url = getSheetAppScriptUrl();
        if (!url) {
            setSheetStatus('Sila masukkan Apps Script URL dahulu.', false);
            return;
        }

        const sheetName = getSheetName();
        const range = getSheetRange();

        const objektif = $('listObjektif')?.innerText.trim();
        const kriteria = $('listKriteria')?.innerText.trim();
        const aktiviti = $('listAktiviti')?.innerText.trim();
        if (!objektif && !kriteria && !aktiviti) {
            setSheetStatus('Tiada data hasil untuk dihantar.', false);
            return;
        }

        const payload = {
            action: 'write',
            sheetId: getSheetId(),
            sheetName,
            range,
            data: {
                objektif,
                kriteria,
                aktiviti,
                timestamp: new Date().toISOString()
            }
        };

        setSheetStatus('Menghantar data...', true);

        try {
            const resp = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const result = await resp.json();
            if (result && result.status === 'ok') {
                setSheetStatus('Data berjaya dihantar.', true);
            } else {
                setSheetStatus('Gagal menghantar: ' + (result.message || 'unknown'), false);
            }
        } catch (e) {
            setSheetStatus('Gagal menghantar: ' + e.message, false);
        }
    }

    loadSheetSettings();

    $('btnConnectSheet').addEventListener('click', connectSheet);
    $('btnPushToSheet').addEventListener('click', pushDataToSheet);

    // Copy Button (Full)
    $('btnSmartCopy').addEventListener('click', () => {
        // ALLOW ALL MODES to run the macro.
        
        // New Macro Sequence
        const commands = constructFullMacroSequence();
        if(!commands || commands.length === 0) return showToast("Sila jana RPH dahulu.", "error");
        
        // UI State Update
        $('btnSmartCopy').style.display = 'none';
        $('btnObjektifCopy').style.display = 'none';
        $('btnStopCopy').style.display = 'block';

        inject(commands);
    });

    // Copy Button (From Objektif)
    $('btnObjektifCopy').addEventListener('click', () => {
        const commands = constructObjektifMacroSequence();
        // Check if there is generated data
        const hasData = $('listObjektif').innerText.length > 20 || $('listAktiviti').innerText.length > 20;
        if(!hasData) return showToast("Sila jana RPH dahulu.", "error");

        // UI State Update
        $('btnSmartCopy').style.display = 'none';
        $('btnObjektifCopy').style.display = 'none';
        $('btnStopCopy').style.display = 'block';

        inject(commands);
    });

    // Stop Button
    $('btnStopCopy').addEventListener('click', () => {
        chrome.tabs.query({active:true, currentWindow:true}, tabs => {
            if(!tabs[0]) return;
            chrome.tabs.sendMessage(tabs[0].id, { action: "STOP_MACRO" }, () => {
                showToast("Proses dihentikan.", "info");
                resetCopyButtons();
            });
        });
    });
}

function resetCopyButtons() {
    const btnSmart = $('btnSmartCopy');
    const btnObj = $('btnObjektifCopy');
    const btnStop = $('btnStopCopy');
    if(btnSmart) btnSmart.style.display = 'flex';
    if(btnObj) btnObj.style.display = 'flex';
    if(btnStop) btnStop.style.display = 'none';
}

// --- NETWORK HELPER: RETRY LOGIC & USAGE COUNT ---
function getTodayDate() {
    return new Date().toISOString().split('T')[0];
}

function updateAPIUsageDisplay() {
    const today = getTodayDate();
    let usageData = JSON.parse(localStorage.getItem('API_USAGE') || '{}');
    
    if (usageData.date !== today) {
        usageData = { date: today, count: 0 };
        localStorage.setItem('API_USAGE', JSON.stringify(usageData));
    }
    
    const display = $('apiUsageBadge');
    if(display) display.textContent = `Use: ${usageData.count}/20`;
}

function resetApiUsage() {
    const today = getTodayDate();
    const usageData = { date: today, count: 0 };
    localStorage.setItem('API_USAGE', JSON.stringify(usageData));
    updateAPIUsageDisplay();
}

function incrementAPICount() {
    const today = getTodayDate();
    let usageData = JSON.parse(localStorage.getItem('API_USAGE') || '{}');
    if (usageData.date !== today) usageData = { date: today, count: 0 };
    usageData.count++;
    localStorage.setItem('API_USAGE', JSON.stringify(usageData));
    updateAPIUsageDisplay();
}

async function fetchWithRetry(url, options, retries = 3, backoff = 1500) {
    incrementAPICount();
    
    // Attach AbortController signal
    if (currentAbortController) {
        options.signal = currentAbortController.signal;
    }

    for (let i = 0; i <= retries; i++) {
        try {
            const response = await fetch(url, options);
            if (response.ok) return response;
            if (response.status === 503 || (response.status >= 500 && response.status < 600)) {
                if (i < retries) {
                    const delay = backoff * Math.pow(2, i); 
                    await new Promise(resolve => setTimeout(resolve, delay));
                    continue;
                }
            }
            return response; 
        } catch (error) {
             if (error.name === 'AbortError') throw error; // Don't retry if aborted
             
            if (i < retries) {
                const delay = backoff * Math.pow(2, i);
                await new Promise(resolve => setTimeout(resolve, delay));
                continue;
            }
            throw error;
        }
    }
}

// --- SMART LEARNING ENGINE (ANALISIS DATA) ---

async function processSmartLearning() {
    if (CURRENT_MODE !== 'API') return showToast("Sila gunakan Mode API untuk analisis.", "error");

    const allFiles = [...dskpBuffer, ...rptBuffer];
    const totalFiles = allFiles.length;
    
    currentAbortController = new AbortController();
    showLoading(true, "Menganalisis & Menyusun Data...", 0);
    
    let processedCount = 0;
    
    for (const file of allFiles) {
        // Check cancel
        if (currentAbortController.signal.aborted) break;

        processedCount++;
        const percent = Math.round((processedCount / totalFiles) * 100);
        updateProgress(percent, `Mengimbas fail ${processedCount}/${totalFiles}: ${file.name}`);
        
        try {
            await analyzeSingleFile(file);
        } catch (e) {
            if (e.name === 'AbortError') {
                console.log("Analysis aborted");
                break;
            }
            console.error(`Gagal analisis fail ${file.name}`, e);
            // Don't stop process, just log
        }
    }
    
    await saveDataToDB(DATA_STORE);
    renderDataResults();
    
    // Hide Upload Section
    $('dataUploadSection').classList.remove('expanded');
    dskpBuffer = []; 
    rptBuffer = [];
    currentAbortController = null;
    showLoading(false);
    showToast("Analisis Selesai & Data Dikemaskini!", "success");
}

async function analyzeSingleFile(file) {
    const instruction = `
    TASK: DOUBLE ANALYZE THIS MALAYSIAN EDUCATION DOCUMENT (DSKP/RPT).
    
    STEP 1: EXTRACTION
    - Identify FORM/YEAR (e.g., "Tingkatan 1", "Tahun 4").
    - Identify SUBJECT.
    - Extract Week, Theme, Topic, SK, SP.

    STEP 2: INTELLIGENT STANDARDIZATION & ACRONYM RESOLUTION
    - YOU MUST EXPAND ACRONYMS for Subject Names.
      - "RBT" -> "Reka Bentuk dan Teknologi"
      - "BM" -> "Bahasa Melayu"
      - "BI" -> "Bahasa Inggeris"
      - "PJPK" -> "Pendidikan Jasmani dan Kesihatan"
      - "ASK" -> "Asas Sains Komputer"
      - "PSV" -> "Pendidikan Seni Visual"
      - "PM" -> "Pendidikan Moral"
      - "PI" -> "Pendidikan Islam"
      - "SEJ" -> "Sejarah"
      - "GEO" -> "Geografi"
    - If subject is missing, infer it from content (e.g. keywords "Algoritma" = ASK).
    
    - STANDARDIZE FORM/YEAR:
      - "Form 1", "Tingkatan Satu" -> "Tingkatan 1"
      - "Year 4", "Darjah 4" -> "Tahun 4"

    OUTPUT FORMAT (STRICT MINIFIED JSON ARRAY):
    [
      {
        "form": "Tingkatan 1",
        "subject": "Reka Bentuk dan Teknologi",
        "minggu": 1, 
        "tema": "...", 
        "tajuk": "...", 
        "sk": "...", 
        "sp": "..." 
      }
    ]
    NO MARKDOWN. NO TEXT. ONLY JSON.
    `;

    const parts = [
        { text: instruction },
        { text: `FILENAME: ${file.name}` }
    ];

    if(file.type.includes('pdf')) {
        parts.push({ inlineData: { mimeType: "application/pdf", data: file.data } });
    } else {
        parts.push({ text: `CONTENT: ${file.data.substring(0, 30000)}` });
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;
    
    // Config for JSON mode
    const payload = {
        contents: [{ parts: parts }],
        generationConfig: { responseMimeType: "application/json" }
    };
    
    const response = await fetchWithRetry(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
    });

    if (!response.ok) throw new Error("API Error " + response.status);
    
    const data = await response.json();
    
    if (data.candidates?.[0]?.finishReason === 'SAFETY') {
        throw new Error("Content blocked by safety filters.");
    }
    
    const jsonStr = data.candidates?.[0]?.content?.parts?.[0]?.text || "";
    if(!jsonStr) throw new Error("Empty response");

    try {
        const cleanJson = jsonStr.replace(/```json/g, '').replace(/```/g, '').trim();
        const parsedArray = JSON.parse(cleanJson);
        
        if (Array.isArray(parsedArray)) {
            mergeDataIntoStore(parsedArray);
        }
    } catch (parseErr) {
        console.warn("JSON Parse Error", parseErr);
        // Try repair or ignore
        throw new Error("Invalid JSON format. File might be too complex.");
    }
}

// Standardization Helpers
function standardizeKey(key) {
    if (!key) return "Unknown";
    let norm = key.trim();
    // Title Case
    norm = norm.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
    
    // Standardize Form names
    if (/tingkatan|form|t/i.test(norm)) {
        if (/satu|one|1/i.test(norm)) return "Tingkatan 1";
        if (/dua|two|2/i.test(norm)) return "Tingkatan 2";
        if (/tiga|three|3/i.test(norm)) return "Tingkatan 3";
        if (/empat|four|4/i.test(norm)) return "Tingkatan 4";
        if (/lima|five|5/i.test(norm)) return "Tingkatan 5";
    }
    // Standardize Year names
    if (/tahun|year|darjah/i.test(norm)) {
        if (/satu|one|1/i.test(norm)) return "Tahun 1";
        if (/dua|two|2/i.test(norm)) return "Tahun 2";
        if (/tiga|three|3/i.test(norm)) return "Tahun 3";
        if (/empat|four|4/i.test(norm)) return "Tahun 4";
        if (/lima|five|5/i.test(norm)) return "Tahun 5";
        if (/enam|six|6/i.test(norm)) return "Tahun 6";
    }
    
    return norm;
}

// Smart Subject Name Standardization (Acronyms Resolution)
function standardizeSubjectName(name) {
    if (!name) return "Umum";
    let n = name.toUpperCase().replace(/[.,&]/g, ' ').trim();
    
    // Dictionary check
    if(n.includes("RBT") || n.includes("REKA BENTUK")) return "Reka Bentuk dan Teknologi";
    if(n.includes("BM") || n.includes("BAHASA MELAYU")) return "Bahasa Melayu";
    if(n.includes("BI") || n.includes("ENGLISH") || n.includes("INGGERIS")) return "Bahasa Inggeris";
    if(n.includes("ASK") || n.includes("SAINS KOMPUTER")) return "Asas Sains Komputer";
    if(n.includes("PJPK") || n.includes("JASMANI") || n.includes("PJK")) return "Pendidikan Jasmani dan Kesihatan";
    if(n.includes("PSV") || n.includes("SENI VISUAL")) return "Pendidikan Seni Visual";
    if(n.includes("PM") || n.includes("MORAL")) return "Pendidikan Moral";
    if(n.includes("PI") || n.includes("ISLAM") || n.includes("PAI")) return "Pendidikan Islam";
    if(n.includes("SEJ") || n.includes("SEJARAH")) return "Sejarah";
    if(n.includes("GEO") || n.includes("GEOGRAFI")) return "Geografi";
    if(n.includes("MAT") || n.includes("MATEMATIK")) return "Matematik";
    if(n.includes("SN") || n.includes("SAINS")) return "Sains";

    // Fallback: Title Case
    return name.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
}

function mergeDataIntoStore(items) {
    items.forEach(item => {
        const form = standardizeKey(item.form);
        const subj = standardizeSubjectName(item.subject);
        
        if (!DATA_STORE[form]) DATA_STORE[form] = {};
        
        // Deep Merging Strategy
        const existingList = DATA_STORE[form][subj] || [];
        const map = new Map();
        
        // Load existing
        existingList.forEach(ex => {
            const w = parseInt(ex.minggu);
            if(!isNaN(w)) map.set(w, ex);
        });
        
        // Merge new (Overwrite logic: if new data is more detailed, use it)
        const w = parseInt(item.minggu);
        if(!isNaN(w)) {
            const existing = map.get(w);
            let merged = item;

            if (existing) {
                // If existing has data but new one is empty, keep existing
                if (!item.tema && existing.tema) merged.tema = existing.tema;
                if (!item.tajuk && existing.tajuk) merged.tajuk = existing.tajuk;
                if (!item.sk && existing.sk) merged.sk = existing.sk;
                if (!item.sp && existing.sp) merged.sp = existing.sp;
            }
            
            map.set(w, {
                minggu: w,
                tema: merged.tema || "",
                tajuk: merged.tajuk || "",
                sk: merged.sk || "",
                sp: merged.sp || ""
            });
        }
        
        // Convert back to sorted array
        DATA_STORE[form][subj] = Array.from(map.values()).sort((a,b) => a.minggu - b.minggu);
    });
}

// --- UI HELPERS ---
function showLoading(show, text, percent=0) { 
    const ol = $('loadingOverlay');
    if (show) {
        ol.classList.add('show');
        $('loadingText').textContent = text;
        updateProgress(percent, text);
    } else {
        ol.classList.remove('show');
    }
}

function showToast(message, type = 'info') {
    const mb = $('messageBar');
    if (!mb) return;
    
    // Reset animation
    mb.style.transition = 'none';
    mb.classList.remove('show', 'error', 'success');
    void mb.offsetWidth; // trigger reflow
    
    mb.style.transition = 'transform 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275), opacity 0.3s ease';
    mb.textContent = message;
    mb.classList.add('show');
    if (type === 'error') mb.classList.add('error');
    if (type === 'success') mb.classList.add('success');

    // Clear existing timeout if any
    if (mb.hideTimeout) clearTimeout(mb.hideTimeout);

    mb.hideTimeout = setTimeout(() => {
        mb.classList.remove('show');
    }, 3000);
}

function updateProgress(percent, text) {
    $('progressBar').style.width = `${percent}%`;
    if(text) $('loadingSubtext').textContent = text;
}

function renderDataResults() {
    const tabsContainer = $('formTabsContainer');
    const resultsArea = $('dataResultsArea');
    
    tabsContainer.innerHTML = "";
    resultsArea.innerHTML = "";
    
    const forms = Object.keys(DATA_STORE).sort();
    if (forms.length === 0) {
        resultsArea.innerHTML = '<div class="empty-files">Tiada data dianalisis. Sila muat naik DSKP & RPT.</div>';
        return;
    }
    
    // Create Form Tabs
    forms.forEach((form, idx) => {
        const btn = document.createElement('button');
        btn.className = `data-tab-btn ${idx===0 ? 'active' : ''}`;
        btn.textContent = form;
        btn.dataset.form = form;
        btn.onclick = () => {
            document.querySelectorAll('#formTabsContainer .data-tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            renderFormContent(form);
        };
        tabsContainer.appendChild(btn);
    });
    
    // Render first form initially
    renderFormContent(forms[0]);
}

function renderFormContent(formName) {
    const resultsArea = $('dataResultsArea');
    resultsArea.innerHTML = "";
    
    const subjects = DATA_STORE[formName] || {};
    const subjectNames = Object.keys(subjects).sort();
    
    if (subjectNames.length === 0) {
        resultsArea.innerHTML = '<div class="empty-files">Tiada subjek.</div>';
        return;
    }

    // 1. Subject Navigation Bar (Chips)
    const subjNav = document.createElement('div');
    subjNav.style.display = 'flex';
    subjNav.style.gap = '0.5rem';
    subjNav.style.overflowX = 'auto';
    subjNav.style.marginBottom = '0.5rem';
    subjNav.style.paddingBottom = '0.25rem';
    
    // Container for Cards
    const cardsContainer = document.createElement('div');

    const renderSubjectCards = (subj) => {
        cardsContainer.innerHTML = "";
        
        // Header with Global Edit Button
        const headerContainer = document.createElement('div');
        headerContainer.style.display = 'flex';
        headerContainer.style.justifyContent = 'space-between';
        headerContainer.style.alignItems = 'center';
        headerContainer.style.marginBottom = "0.5rem";

        // Subject Title / Input
        const titleDisplay = document.createElement('h3');
        titleDisplay.textContent = subj;
        titleDisplay.style.color = "#c084fc";
        titleDisplay.style.fontSize = "0.8rem";
        titleDisplay.style.margin = "0";

        const titleInput = document.createElement('input');
        titleInput.type = 'text';
        titleInput.value = subj;
        titleInput.style.display = 'none';
        titleInput.style.fontSize = "0.8rem";
        titleInput.style.padding = "0.2rem";
        titleInput.style.width = "70%";

        const editBtn = document.createElement('button');
        editBtn.className = "btn-secondary bg-blue";
        editBtn.style.fontSize = "0.6rem";
        editBtn.textContent = "Edit Data";
        
        let isEditing = false;
        
        editBtn.onclick = () => {
            isEditing = !isEditing;
            const inputs = cardsContainer.querySelectorAll('.data-edit-input');
            
            if(isEditing) {
                // Switch to Edit Mode
                editBtn.textContent = "Simpan";
                editBtn.classList.replace("bg-blue", "bg-green");
                
                // Show title input
                titleDisplay.style.display = 'none';
                titleInput.style.display = 'block';
                
                // Enable card inputs
                inputs.forEach(i => i.disabled = false);
            } else {
                // Save Mode
                const newSubjName = standardizeSubjectName(titleInput.value);
                const oldSubjName = subj;

                // 1. If Name Changed, Move Data
                if (newSubjName !== oldSubjName) {
                    if (DATA_STORE[formName][newSubjName]) {
                        // Merge if target exists (simple concat for now, sophisticated merge handled by analyzer)
                        DATA_STORE[formName][newSubjName] = DATA_STORE[formName][newSubjName].concat(subjects[oldSubjName]);
                    } else {
                        // Create new
                        DATA_STORE[formName][newSubjName] = subjects[oldSubjName];
                    }
                    delete DATA_STORE[formName][oldSubjName];
                    
                    // Trigger refresh
                    saveDataToDB(DATA_STORE);
                    renderFormContent(formName); // Re-render whole form
                    showToast(`Subjek ditukar ke ${newSubjName}`, "success");
                    return; 
                }

                editBtn.textContent = "Edit Data";
                editBtn.classList.replace("bg-green", "bg-blue");
                
                // Revert title UI
                titleDisplay.style.display = 'block';
                titleInput.style.display = 'none';
                
                // Disable inputs
                inputs.forEach(i => i.disabled = true);
                saveDataToDB(DATA_STORE);
                showToast("Data Disimpan", "success");
            }
        };

        headerContainer.append(titleDisplay, titleInput, editBtn);
        cardsContainer.appendChild(headerContainer);

        const items = subjects[subj] || [];
        if(items.length === 0) {
            cardsContainer.innerHTML += '<div class="empty-files">Tiada data minggu.</div>';
            return;
        }

        items.forEach((item) => {
            const card = document.createElement('div');
            card.className = "data-card";
            
            const createField = (label, val, key) => {
                const d = document.createElement('div');
                d.style.marginBottom = "0.25rem";
                d.innerHTML = `<span style="font-size:0.6rem; color:#64748b;">${label}:</span>`;
                const inp = document.createElement('input');
                inp.className = "data-edit-input";
                inp.value = val || "";
                inp.disabled = true;
                inp.onchange = (e) => item[key] = e.target.value;
                d.appendChild(inp);
                return d;
            };
            
            card.appendChild(createField("Minggu", item.minggu, "minggu"));
            card.appendChild(createField("Tema", item.tema, "tema"));
            card.appendChild(createField("Tajuk", item.tajuk, "tajuk"));
            card.appendChild(createField("SK", item.sk, "sk"));
            card.appendChild(createField("SP", item.sp, "sp"));
            
            cardsContainer.appendChild(card);
        });
    };

    subjectNames.forEach((subj, idx) => {
        const chip = document.createElement('button');
        chip.className = `btn-secondary ${idx===0 ? 'bg-purple' : ''}`;
        chip.style.fontSize = "0.65rem";
        chip.textContent = subj;
        chip.onclick = () => {
            // Reset chips
            Array.from(subjNav.children).forEach(c => c.className = 'btn-secondary');
            chip.className = 'btn-secondary bg-purple';
            renderSubjectCards(subj);
        };
        subjNav.appendChild(chip);
    });

    // Initial Render
    renderSubjectCards(subjectNames[0]);

    resultsArea.append(subjNav, cardsContainer);
}

// Helper to switch between Input and Select
function toggleInputType(id, type, options = []) {
    const input = $(id);
    const select = $(id + 'Select');
    
    if(!input || !select) return; // Error prevention

    select.onchange = null;
    select.innerHTML = ''; // Clear previous

    if (type === 'select' && options.length > 1) {
        // Multi-option: Show Select
        select.innerHTML = '<option value="">-- Sila Pilih / Gabung Semua --</option>';
        
        const optAll = document.createElement('option');
        optAll.value = options.join(', ');
        optAll.textContent = `[GABUNG SEMUA]`;
        optAll.style.fontWeight = 'bold';
        optAll.style.color = '#4ade80';
        select.appendChild(optAll);

        options.forEach(optVal => {
            const opt = document.createElement('option');
            opt.value = optVal;
            opt.textContent = optVal.substring(0, 60) + (optVal.length > 60 ? '...' : '');
            select.appendChild(opt);
        });
        // Custom Option
        const customOpt = document.createElement('option');
        customOpt.value = "CUSTOM_MODE";
        customOpt.textContent = "✏️ [Custom] Tulis Sendiri...";
        customOpt.style.color = '#fbbf24'; 
        select.appendChild(customOpt);

        input.classList.add('hide');
        select.classList.remove('hide');
        select.value = ""; // Default empty

        select.onchange = function() {
            if (this.value === "CUSTOM_MODE") {
                select.classList.add('hide');
                input.classList.remove('hide');
                input.value = ""; input.focus();
            }
        };
    } else {
        // Single option or Text Mode: Show Input
        input.classList.remove('hide');
        select.classList.add('hide');
        
        // Populate if data exists
        if (Array.isArray(options)) {
            if (options.length > 0) input.value = options.join(', ');
        } else if (typeof options === 'string') {
            input.value = options;
        } else {
            input.value = "";
        }
    }
}

function performDeepSearch(targetWeek, targetTopic) {
    // 1. Reset all fields first
    toggleInputType('inputTema', 'text', "");
    toggleInputType('inputTajuk', 'text', "");
    toggleInputType('inputSK', 'text', "");
    toggleInputType('inputSP', 'text', "");

    // Search within DATA_STORE
    let matches = [];
    const qW = parseInt(targetWeek);
    const qT = targetTopic.toLowerCase();
    
    for (const form in DATA_STORE) {
        for (const subj in DATA_STORE[form]) {
            const list = DATA_STORE[form][subj];
            list.forEach(item => {
                let match = false;
                if (!isNaN(qW) && parseInt(item.minggu) === qW) match = true;
                if (qT && (
                    (item.tema && item.tema.toLowerCase().includes(qT)) ||
                    (item.tajuk && item.tajuk.toLowerCase().includes(qT)) ||
                    (item.sk && item.sk.toLowerCase().includes(qT)) ||
                    (item.sp && item.sp.toLowerCase().includes(qT))
                )) match = true;
                
                if (match) matches.push(item);
            });
        }
    }

    if (matches.length > 0) {
        // Use Set to remove exact duplicates
        const uniqueThemes = Array.from(new Set(matches.map(m => m.tema).filter(Boolean)));
        const uniqueTopics = Array.from(new Set(matches.map(m => m.tajuk).filter(Boolean)));
        const uniqueSK = Array.from(new Set(matches.map(m => m.sk).filter(Boolean)));
        const uniqueSP = Array.from(new Set(matches.map(m => m.sp).filter(Boolean)));

        toggleInputType('inputTema', uniqueThemes.length > 1 ? 'select' : 'text', uniqueThemes);
        toggleInputType('inputTajuk', uniqueTopics.length > 1 ? 'select' : 'text', uniqueTopics);
        toggleInputType('inputSK', uniqueSK.length > 1 ? 'select' : 'text', uniqueSK);
        toggleInputType('inputSP', uniqueSP.length > 1 ? 'select' : 'text', uniqueSP);

        $('rptRawData').value = `Data dijumpai (${matches.length} padanan).`;
        showToast("Data Dijumpai & Dikemaskini!", "success");
    } else {
        showToast("Tiada padanan dalam Pengurusan Data.", "error");
    }
}

// --- FILE HELPERS ---
async function readFileAsBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

function readFileAsText(f) { 
    return new Promise((r) => { 
        const rr=new FileReader(); 
        rr.onload=()=>r(rr.result); 
        rr.readAsText(f); 
    }); 
}

// --- SMART EXECUTION ENGINE ---
async function smartGenerate(promptText, actionType, fallbackData = null) {
    if (CURRENT_MODE === 'API') {
        currentAbortController = new AbortController();
        showLoading(true, "Menjana (Cloud API)...");
        
        // Key Rotation Logic
        let attempts = 0;
        const totalKeys = API_KEYS_LIST.length > 0 ? API_KEYS_LIST.length : 1;
        let success = false;

        while (attempts < totalKeys) {
            try {
                const result = await callCloudAPI(promptText);
                handleSuccess(result, actionType);
                success = true;
                break; // Exit loop on success
            } catch (e) {
                if (e.name === 'AbortError') {
                    console.log("Generation aborted");
                    return;
                }

                // Check for Quota/Limit errors
                const isLimitError = e.message.includes("Quota Exceeded") || 
                                     e.message.includes("Server Busy") || 
                                     e.message.includes("429");

                if (isLimitError && API_KEYS_LIST.length > 1) {
                    attempts++;
                    if (attempts < totalKeys) {
                        // Logic to switch to next key
                        let currentIndex = API_KEYS_LIST.indexOf(API_KEY);
                        // If current key not found (legacy/deleted), start from 0
                        if (currentIndex === -1) currentIndex = 0;
                        
                        const nextIndex = (currentIndex + 1) % API_KEYS_LIST.length;
                        const nextKey = API_KEYS_LIST[nextIndex];
                        
                        showToast(`Had Key ${currentIndex + 1} penuh. Menukar ke Key ${nextIndex + 1}...`, "error");
                        
                        // Switch Key (updates global API_KEY and UI)
                        selectKey(nextKey, true);
                        
                        // Brief pause to allow UI update and prevent instant spamming
                        await new Promise(resolve => setTimeout(resolve, 1500));
                        continue; // Retry loop
                    }
                }
                
                // If error is not quota related, or we've tried all keys
                console.warn("API Call Failed:", e);
                break;
            }
        }

        showLoading(false);
        currentAbortController = null;

        if (success) {
            return;
        }

        // If we reach here, all attempts failed. Fallback to Web.
        showToast("Semua API Key Had Dicapai. Menukar ke Web Mode...", "error");
        CURRENT_MODE = 'WEB';
        localStorage.setItem("PREFERRED_MODE", 'WEB'); // Persist preference
        updateStatusUI();
        
        // Trigger Web Fallback
        webFallback('raw', { prompt: promptText });
        return; 
    }
    
    if (CURRENT_MODE === 'WEB') {
        showToast("Menghantar ke Gemini Web...", "info");
        // Auto paste to Gemini Web logic
        webFallback('raw', { prompt: promptText });
        showLoading(false);
    }
}

function handleSuccess(responseText, actionType) {
    renderOutputs(responseText);
    showToast("RPH Siap Dijana!", "success");
}

async function callCloudAPI(promptText) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;
    const payload = { contents: [{ parts: [{ text: promptText }] }] };
    
    // Pass controller signal if active
    const options = {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
    };
    
    const response = await fetchWithRetry(url, options);

    if (!response.ok) {
        if (response.status === 429) throw new Error("Quota Exceeded");
        if (response.status === 503) throw new Error("Server Busy");
        throw new Error(`API Error ${response.status}`);
    }
    const data = await response.json();
    return data.candidates?.[0]?.content?.parts?.[0]?.text || null;
}

// --- WEB FALLBACK (ROBUST) ---
async function webFallback(mode, inputs) {
    let prompt = inputs.prompt || buildRPHPrompt(mode, inputs);
    
    // Check if Gemini is already open
    chrome.tabs.query({ url: "*://gemini.google.com/*" }, tabs => {
        let targetTabId = null;

        if (tabs && tabs.length > 0) {
             // Found existing tab, use the first one
             targetTabId = tabs[0].id;
             chrome.tabs.update(targetTabId, { active: true });
             showToast("Membuka tab Gemini sedia ada...", "info");
             
             // Immediate send for existing tab
             sendMessageWithRetry(targetTabId, prompt);

        } else {
             // Create new tab
             chrome.tabs.create({ url: 'https://gemini.google.com/app' }, (newTab) => {
                 targetTabId = newTab.id;
                 // Poll wait for new tab
                 sendMessageWithRetry(targetTabId, prompt, true);
             });
             showToast("Membuka tab Gemini baru...", "info");
        }
    });
}

function sendMessageWithRetry(tabId, prompt, isNewTab = false) {
    let attempts = 0;
    // Increase attempts: 30 seconds for new tab, 10 seconds for existing
    const maxAttempts = isNewTab ? 30 : 10; 
    
    const interval = setInterval(() => {
        attempts++;
        if(attempts > maxAttempts) {
            clearInterval(interval);
            // Feedback to user
            showToast("Gagal menyambung ke Gemini Web. Sila pastikan tab aktif.", "error");
            return;
        }

        chrome.tabs.sendMessage(tabId, { action: "PASTE_TO_GEMINI_WEB", text: prompt }, response => {
            if (chrome.runtime.lastError) {
                // Tab might be loading or content script not injected yet
                return;
            }
            if (response && response.success) {
                clearInterval(interval); 
                showToast("Arahan berjaya dihantar!", "success");
            } else if (response && response.status === "generating_aborted") {
                clearInterval(interval);
                showToast("Gemini sedang menjana. Sila tunggu.", "error");
            }
        });
    }, 1000); 
}

// Prompt Builder
function buildRPHPrompt(mode, d) {
    const plainTextDisclaimer = `
PENTING:
- Sediakan output dalam format teks yang ringkas.
- JANGAN guna formula matematik. Guna simbol biasa sahaja.
- LIMIT: MAKSIMUM 6 AYAT/POINT BAGI SETIAP BAHAGIAN.
- PASTIKAN ADA NOMBOR (1., 2.) UNTUK MUDAHKAN PARSING.

CONSTRAINT (WAJIB IKUT):
1. OBJEKTIF: Terus tulis ayat objektif. JANGAN mula dengan "Murid dapat". JANGAN tulis pendahuluan.
2. KRITERIA: Terus tulis kriteria. JANGAN mula dengan "Murid dapat".
3. AKTIVITI: TEPAT 6 langkah (Wajib). Jangan kurang, jangan lebih. Tulis ayat aktiviti sahaja. JANGAN letak label "Engage", "Explore", "Evaluate", "Rumusan", "Penutup".
4. GAMBAR: JANGAN hasilkan sebarang gambar, imej, atau diagram. HANYA TEKS.
`;

    // INTELLIGENT AUTO-FILL INSTRUCTION
    // If standard fields are missing, user wants AI to generate them.
    let autoFillInstruction = "";
    if (mode === 'standard' && (!d.sk || d.sk.length < 3 || !d.sp || d.sp.length < 3)) {
        autoFillInstruction = `
        IMPORTANT: The user has NOT provided specific Standard Kandungan (SK) or Standard Pembelajaran (SP).
        BASED ON THE THEME/TOPIC: "${d.tema} - ${d.tajuk}", YOU MUST:
        1. SUGGEST a relevant SK and SP from the standard Malaysian Curriculum.
        2. Use these suggested SK/SP to generate the Objectives and Activities.
        `;
    }

    if (mode === 'standard') {
        return `
PERANAN: Pakar RPH Pendidikan Malaysia (Format TS25).
INPUT: Tema:${d.tema}, Tajuk:${d.tajuk}, SK:${d.sk}, SP:${d.sp}, Nota:${d.note}

${autoFillInstruction}
${plainTextDisclaimer}

ARAHAN RPH KHUSUS:
1. [OBJEKTIF]: Format ABCD (Audience, Behavior, Condition, Degree). Mula dengan kata kerja.
2. [KRITERIA]: Kriteria kejayaan spesifik.
3. [AKTIVITI]: TEPAT 6 langkah (Wajib).

FORMAT OUTPUT WAJIB:
[OBJEKTIF]
1. ...
[KRITERIA]
1. ...
[AKTIVITI]
1. ...
2. ...
3. ...
4. ...
5. ...
6. ...
`;
    } 
    else if (mode === 'exam') {
        return `
Sila berikan maklumat peperiksaan di bawah, kemudian fokuskan penjanaan Rancangan Pengajaran Harian (RPH) kepada bahagian **Objektif Pembelajaran**, **Kriteria Kejayaan (KK)**, dan **Aktiviti Pengajaran & Pembelajaran (PdP)** yang terperinci.

1.  Mata Pelajaran: ${d.tema}
2.  Jenis Peperiksaan: ${d.examType || "Peperiksaan"}
-------------------------------------------------------------------------

**RANCANGAN PENGAJARAN HARIAN (RPH) UNIVERSAL PENGAWASAN PEPERIKSAAN**
-------------------------------------------------------------------------

**A. BUTIRAN UMUM**

| Kategori | Input |
| :--- | :--- |
| **Mata Pelajaran** | **${d.tema}** |
| **Jenis Peperiksaan** | **${d.examType}** |
|| **Guru Bertugas** | **[Guru Pengawas]** |
| **Tema / Tajuk** | Pengawasan Peperiksaan **${d.examType}**: Kertas **${d.tema}** |

**B. OBJEKTIF DAN KRITERIA KEJAYAAN (FOKUS UTAMA)**

* **Objektif Pembelajaran:**
    [OBJEKTIF]
    1.  Murid dapat menduduki peperiksaan mengikut prosedur dan peraturan bilik peperiksaan yang ditetapkan (Aspek Pengurusan & Disiplin).
    2.  Murid dapat menjawab soalan dengan penuh kejujuran, fokus, dan memanfaatkan masa yang diperuntukkan secara optimum (Aspek Integriti & Ketekunan).
    [END OBJEKTIF]

* **Kriteria Kejayaan (KK):**
    [KRITERIA]
    1.  Murid berjaya melengkapkan maklumat peribadi dan butiran kertas soalan pada skrip jawapan dengan betul.
    2.  Murid menjawab soalan secara bersendirian (tiada salah laku dikesan) dan kekal berada di tempat duduk sehingga tamat masa.
    3.  Murid menyerahkan kertas jawapan dengan teratur kepada guru pengawas apabila diarah berhenti menulis.
    [END KRITERIA]

**C. AKTIVITI PENGAJARAN DAN PEMBELAJARAN (PdP) (FOKUS UTAMA)**
    [AKTIVITI]
    1. **Penyediaan Persekitaran:** Guru mengarahkan semua murid meletakkan semua barang peribadi kecuali alat tulis yang dibenarkan di lokasi yang ditetapkan.
    2. **Semakan Kehadiran:** Guru memastikan setiap murid berada di tempat duduk yang betul dan merekodkan kehadiran.
    3. **Edaran Kertas & Arahan:** Guru mengedarkan kertas soalan dan memberi arahan terakhir (semak muka surat, masa tamat).
    4. **Proses Menjawab:** Murid mula menjawab soalan. Guru mengawasi dengan berjalan di sekitar bilik secara berkala.
    5. **Amaran Masa:** Guru memberi amaran masa secara jelas pada 30 minit terakhir dan 5 minit terakhir.
    6. **Penutup Sesi:** Guru mengarahkan berhenti menulis, mengutip skrip jawapan, dan mengira bilangan skrip sebelum membenarkan murid keluar.
    [END AKTIVITI]
`;
    }
    else if (mode === 'special') {
        return `
=========================================================================
**PROMPT RPH AKTIVITI SELEPAS UASA (KREATIF & MENARIK)**
=========================================================================
Gunakan maklumat input di bawah untuk menjana Rancangan Pengajaran Harian (RPH) bagi aktiviti selepas Ujian Akhir Sesi Akademik (UASA). Aktiviti yang dihasilkan mestilah **kreatif, menyeronokkan, dan mempunyai nilai pendidikan** yang santai dan berkaitan dengan subjek yang disenaraikan.

**Maklumat Input Wajib:**
1. Subjek: ${d.subject}
2. Tingkatan: ${d.tingkatan}

**Maklumat Input Pilihan:**
3. Jenis Aktiviti: ${d.activity || "Aktiviti Bebas Kreatif"}

**Struktur Aktiviti Output (Berfokus):**
Output mesti mengandungi bahagian-bahagian berikut dalam format yang terperinci:

1.  **Nama Projek/Aktiviti Kreatif:** (Sertakan nama yang menarik dan relevan)
2.  **Tema Utama:** (Penerangan ringkas tentang objektif santai berkaitan subjek)
3.  **Objektif Aktiviti:** (Apa yang pelajar akan capai secara santai dan berkumpulan)
4.  **Langkah Pelaksanaan (RPH Ringkas):** (Prosedur langkah demi langkah yang mudah diikuti untuk guru dan murid)

**FORMAT OUTPUT WAJIB UNTUK PARSING:**
[OBJEKTIF]
1. ...
[KRITERIA]
1. ...
[AKTIVITI]
1. ...
2. ...
3. ...
4. ...
5. ...
6. ...
`;
    }
}

function renderOutputs(text) {
    const parse = (t1, t2) => {
        const m = text.match(new RegExp(`\\[${t1}\\]([\\s\\S]*?)(\\[${t2}\\]|$)`, 'i'));
        return m ? m[1].trim() : "";
    };

    // Special parsing for "Aktiviti Khas" title extraction if needed
    if($('rphMode').value === 'special') {
        const nameMatch = text.match(/Nama Projek.*?:\s*(.*)/i);
        const themeMatch = text.match(/Tema Utama.*?:\s*(.*)/i);
        if(nameMatch) $('inputTajuk').value = nameMatch[1].trim();
        if(themeMatch) $('inputTema').value = themeMatch[1].trim();
    }
    
    // Exam parsing if needed (mostly static but good to fill)
    if($('rphMode').value === 'exam') {
        $('inputTajuk').value = `Pengawasan Peperiksaan ${$('inputExamType').value}`;
    }

    let objText = parse("OBJEKTIF", "KRITERIA") || parse("OBJEKTIF", "END") || parse("OBJEKTIF", "AKTIVITI"); // fallback
    let kritText = parse("KRITERIA", "AKTIVITI") || parse("KRITERIA", "END");
    let aktText = parse("AKTIVITI", "END");

    // Helper regex for cleaning explanations (Shared)
    // Removes: (Audience), (Condition), (SP 1.2.3), (2.1.4) etc if formatted as code/label
    const cleanRegex = /\s*\((?:SP\s*[\d\.]+|Audience|Behavior|Behaviour|Condition|Degree|Audiens|Tingkah\s*laku|Syarat|Tahap|[ABCD])\)/gi;

    // --- CLEAN & PROCESS OBJEKTIF ---
    if (objText) {
         objText = objText.replace(cleanRegex, "");
         objText = objText.replace(/[ \t]{2,}/g, " ");

         // Intelligent Split: Join lines that don't start with a number
         objText = objText.replace(/\n(?!\s*\d+[\.\)]\s)/g, " ");
         // Ensure list items start on new line
         objText = objText.replace(/(\s+)(\d+[\.\)]\s)/g, "\n$2");

         let lines = objText.split('\n').map(l => l.trim()).filter(l => l.length > 2);
         if (lines.length > 6) lines = lines.slice(0, 6);
         objText = lines.join('\n');
    }

    // --- PROCESS KRITERIA (Limit 6) ---
    if (kritText) {
         kritText = kritText.replace(cleanRegex, "");

         kritText = kritText.replace(/\n(?!\s*\d+[\.\)]\s)/g, " ");
         kritText = kritText.replace(/(\s+)(\d+[\.\)]\s)/g, "\n$2");

         let lines = kritText.split('\n').map(l => l.trim()).filter(l => l.length > 2);
         if (lines.length > 6) lines = lines.slice(0, 6);
         kritText = lines.join('\n');
    }

    // --- CLEAN & PROCESS AKTIVITI ---
    if (aktText) {
        // Remove common labels
        aktText = aktText.replace(/[\(\[\{]?(Fasa|Engage|Explore|Explain|Elaborate|Evaluate|Rumusan|Penutup|Pentaksiran|Langkah)[\)\]\}]?\s*[:\-]?\s*/gi, "");
        
        // Remove explanations (SP codes, Degree, etc)
        aktText = aktText.replace(cleanRegex, "");

        // Pre-process: Strip markdown bold/italic used on numbers to ensure regex matching
        aktText = aktText.replace(/\*\*/g, "").replace(/\*/g, "");

        // Intelligent split
        // 1. Merge lines that are NOT list starts (Fixes "sentence broken by newline" issue)
        // Checks for newline NOT followed by "1. ", "2. " etc.
        aktText = aktText.replace(/\n(?!\s*\d+[\.\)]\s)/g, " ");
        
        // 2. Ensure list items start on new line (Fixes "1. A 2. B" issue)
        // Requires space after dot/paren to avoid breaking versions like 2.1.4
        aktText = aktText.replace(/(\s+)(\d+[\.\)]\s)/g, "\n$2");

        let lines = aktText.split('\n').map(l => l.trim()).filter(l => l.length > 5);
        if (lines.length > 6) {
            lines = lines.slice(0, 6);
        }
        aktText = lines.join('\n');
    }

    generateList($('listObjektif'), objText);
    generateList($('listKriteria'), kritText);
    generateList($('listAktiviti'), aktText);
}

function generateList(container, text) {
    container.innerHTML = "";
    if (!text) return;
    const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 2);
    lines.forEach((line, index) => {
        // Clean leading numbers/bullets
        const cleanContent = line.replace(/^[\d]+[\.\)\-]\s*/, '').replace(/\*\*/g, '').trim();
        const row = createListRow(cleanContent, index + 1, container);
        container.appendChild(row);
    });
}

// Helper to create a single row element with split buttons
function createListRow(content, number, container) {
    const row = document.createElement('div');
    row.className = 'list-row-compact';

    const numLabel = document.createElement('div');
    numLabel.className = 'item-number';
    numLabel.textContent = number + ".";

    const textarea = document.createElement('textarea');
    textarea.value = content;
    textarea.rows = 1;
    textarea.addEventListener('input', function() {
        this.style.height = 'auto';
        this.style.height = (this.scrollHeight) + 'px';
    });

    // Actions Container
    const actionDiv = document.createElement('div');
    actionDiv.className = 'row-actions';

    // 1. Paste Button (Left side of actions)
    const pasteBtn = document.createElement('button');
    pasteBtn.className = 'btn-icon paste-mini';
    pasteBtn.innerHTML = `📋`;
    pasteBtn.title = "Salin";
    pasteBtn.onclick = () => inject(textarea.value.trim());

    // 2. Container for Split Buttons (Right side)
    const splitDiv = document.createElement('div');
    splitDiv.className = 'split-actions';

    // 3. Delete Button (Top)
    const delBtn = document.createElement('button');
    delBtn.className = 'btn-icon delete-mini half';
    delBtn.innerHTML = `✕`;
    delBtn.title = "Padam Baris";
    delBtn.onclick = () => {
        row.remove();
        updateListNumbers(container);
    };

    // 4. Add Button (Bottom)
    const addBtn = document.createElement('button');
    addBtn.className = 'btn-icon add-mini half';
    addBtn.innerHTML = `+`;
    addBtn.title = "Tambah Baris Bawah";
    addBtn.onclick = () => {
        // Create new empty row
        const newRow = createListRow("", 0, container); // Number will be fixed by updateListNumbers
        if (row.nextSibling) {
            container.insertBefore(newRow, row.nextSibling);
        } else {
            container.appendChild(newRow);
        }
        updateListNumbers(container);
    };

    splitDiv.append(delBtn, addBtn);
    actionDiv.append(pasteBtn, splitDiv);
    row.append(numLabel, textarea, actionDiv);
    
    return row;
}

// Helper to re-calculate numbers after add/delete
function updateListNumbers(container) {
    const rows = container.querySelectorAll('.list-row-compact');
    rows.forEach((row, index) => {
        const label = row.querySelector('.item-number');
        if (label) label.textContent = (index + 1) + ".";
    });
}

// Updated Inject for MACRO logic
async function inject(commands) {
    // Legacy support: if array of strings passed (from single copy buttons), convert to paste commands
    if (Array.isArray(commands) && typeof commands[0] === 'string') {
        const rows = commands;
        commands = [];
        for (const text of rows) {
            if(text) commands.push({ type: 'paste', text });
            commands.push({ type: 'move', key: 'ArrowDown', count: 1 });
        }
    } else if (typeof commands === 'string') {
         // Single string paste
         commands = [{ type: 'paste', text: commands }];
    }

    chrome.tabs.query({active:true, currentWindow:true}, tabs => {
        if(!tabs[0]) return showToast("Sila buka Spreadsheet.", "error");
        
        chrome.tabs.sendMessage(tabs[0].id, { 
            action: "FULL_AUTO_MACRO", 
            commands: commands 
        }, r => {
            if(chrome.runtime.lastError) {
                console.error(chrome.runtime.lastError);
                showToast("Sila Refresh page Spreadsheet.", "error");
            } else {
                if(r && r.status === "macro_started") {
                    showToast("Proses Salin Bermula...", "info");
                } else if(r && r.status === "macro_completed") {
                    showToast("Berjaya! (Macro)", "success");
                    // Reset UI
                    resetCopyButtons();
                }
            }
        });
    });
}

function getInputs(mode) {
    if (mode === 'standard') {
        const getVal = (id) => {
            const select = $(id + 'Select');
            if (select && !select.classList.contains('hide')) return select.value;
            return $(id).value;
        };
        // Allow Partial Inputs for Auto-Fill logic
        return { 
            tema: getVal('inputTema'), 
            tajuk: getVal('inputTajuk'), 
            sk: getVal('inputSK'), 
            sp: getVal('inputSP'), 
            note: $('rptRawData').value 
        };
    } else if (mode === 'exam') {
        return { tema: $('inputExamSubject').value, examType: $('inputExamType').value };
    } else {
        return { subject: $('specialSubject').value, tingkatan: $('specialForm').value, activity: $('specialActivity').value };
    }
}

// DB Helpers
function openDB(){ return new Promise((r,j)=>{ const q=indexedDB.open(DB_NAME,DB_VERSION); q.onupgradeneeded=e=>e.target.result.createObjectStore(STORE_NAME); q.onsuccess=e=>r(e.target.result); q.onerror=j; }); }
async function saveDataToDB(data){ const db=await openDB(); const tx=db.transaction(STORE_NAME,'readwrite'); tx.objectStore(STORE_NAME).put(data, 'MAIN_STORE'); }
async function loadDataFromDB(){ const db=await openDB(); return new Promise(r=>{ db.transaction(STORE_NAME).objectStore(STORE_NAME).get('MAIN_STORE').onsuccess=e=>r(e.target.result); }); }

// --- SPECIFIC MACRO SEQUENCE REWRITE ---
function constructFullMacroSequence() {
    // SMART INPUT RETRIEVAL
    let tema = $('inputTema').value || "";
    if (!tema && $('inputExamSubject').value) tema = $('inputExamSubject').value;
    if (!tema && $('specialSubject').value) tema = $('specialSubject').value;

    let tajuk = $('inputTajuk').value || "";
    if (!tajuk && $('inputExamType').value) tajuk = `Pengawasan Peperiksaan: ${$('inputExamType').value}`;
    if (!tajuk && $('specialActivity').value) tajuk = $('specialActivity').value;

    // Standard inputs can be multi-line
    const skText = $('inputSK').value || "";
    const skLines = skText.split('\n').map(s=>s.trim()).filter(s=>s);
    
    const spText = $('inputSP').value || "";
    const spLines = spText.split('\n').map(s=>s.trim()).filter(s=>s);

    const commands = [];
    const addMove = (key, count=1) => commands.push({ type: 'move', key, count });
    const addPaste = (text) => commands.push({ type: 'paste', text });

    // --- START POINT: A47 (KELAS/CLASS) ---

    // 1. Arrow Down 1x, Right 1x: Tema 1
    addMove('ArrowDown', 1);
    addMove('ArrowRight', 1);
    if(tema) addPaste(tema);

    // 2. Arrow Down 1x: Tema 2
    addMove('ArrowDown', 1);
    
    // 3. Arrow Down 1x: Tema 3
    addMove('ArrowDown', 1);

    // 4. Arrow Down 2x: Tajuk 1
    addMove('ArrowDown', 2);
    if(tajuk) addPaste(tajuk);

    // 5. Arrow Right 1x: Tajuk 2
    addMove('ArrowRight', 1);

    // 6. Arrow Right 1x: Tajuk 3
    addMove('ArrowRight', 1);

    // 7. Arrow Down 1x, Arrow Left 2x: SK 1 (Tajuk 1)
    addMove('ArrowDown', 1);
    addMove('ArrowLeft', 2);
    if(skLines[0]) addPaste(skLines[0]);

    // 8. Arrow Down 1x: SK 2 (Tajuk 1)
    addMove('ArrowDown', 1);
    if(skLines[1]) addPaste(skLines[1]);

    // 9. Right 1x, Up 1x: SK 1 (Tajuk 2)
    addMove('ArrowRight', 1);
    addMove('ArrowUp', 1);

    // 10. Arrow Down 1x: SK 2 (Tajuk 2)
    addMove('ArrowDown', 1);

    // 11. Right 1x, Up 1x: SK 1 (Tajuk 3)
    addMove('ArrowRight', 1);
    addMove('ArrowUp', 1);

    // 12. Arrow Down 1x: SK 2 (Tajuk 3)
    addMove('ArrowDown', 1);

    // 13. Down 1x, Left 2x: SP 1 (Tajuk 1) start
    addMove('ArrowDown', 1);
    addMove('ArrowLeft', 2);

    // Fill SP Column 1 (5 rows)
    for(let i=0; i<5; i++) {
        if(spLines[i]) addPaste(spLines[i]);
        if(i < 4) addMove('ArrowDown', 1);
    }
    // Cursor is at SP Row 5

    // 14. Right 1x, Up 4x: SP 1 (Tajuk 2) start
    addMove('ArrowRight', 1);
    addMove('ArrowUp', 4);

    // Fill SP Column 2 (5 rows)
    for(let i=0; i<5; i++) {
        // Assuming no specific data for col 2, just move
        if(i < 4) addMove('ArrowDown', 1);
    }

    // 15. Right 1x, Up 4x: SP 1 (Tajuk 3) start
    addMove('ArrowRight', 1);
    addMove('ArrowUp', 4);

    // Fill SP Column 3 (5 rows)
    for(let i=0; i<5; i++) {
        // Assuming no specific data for col 3, just move
        if(i < 4) addMove('ArrowDown', 1);
    }
    // Cursor at SP Row 5 (Tajuk 3)

    // 16. Down 2x: Objektif
    addMove('ArrowDown', 2);
    
    // Append the Objektif-end logic
    const tailCommands = constructObjektifMacroSequence();
    return commands.concat(tailCommands);
}

function constructObjektifMacroSequence() {
    const commands = [];
    const addMove = (key, count=1) => commands.push({ type: 'move', key, count });
    const addPaste = (text) => commands.push({ type: 'paste', text });

    const getList = (id) => {
        const items = [];
        $(id).querySelectorAll('textarea').forEach(t => items.push(t.value.trim()));
        return items;
    };
    const opList = getList('listObjektif');
    const kkList = getList('listKriteria');
    const aktList = getList('listAktiviti');

    // Objektif
    for (let i = 0; i < 6; i++) {
        if(opList[i]) addPaste(opList[i]);
        if(i < 5) addMove('ArrowDown', 1);
    }

    // Previous logic: End of Obj -> Down 2 -> Kriteria
    addMove('ArrowDown', 2);
    
    // Kriteria
    for (let i = 0; i < 6; i++) {
        if(kkList[i]) addPaste(kkList[i]);
        if(i < 5) addMove('ArrowDown', 1);
    }

    // Previous logic: End of Krit -> Down 1 -> Aktiviti
    addMove('ArrowDown', 1);

    // Aktiviti
    for (let i = 0; i < 6; i++) {
        if(aktList[i]) addPaste(aktList[i]);
        if(i < 5) addMove('ArrowDown', 1);
    }

    return commands;
}