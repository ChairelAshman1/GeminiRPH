// Track the last focused element
let lastFocusedElement = null;
document.addEventListener('focus', (e) => {
    lastFocusedElement = e.target;
}, true);

// Helper for delays
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// Macro Control Flags
let isMacroRunning = false;
let stopMacroRequest = false;

// LISTEN FOR USER COPY EVENTS (To Auto-fill Extension Fields)
document.addEventListener('copy', () => {
    try {
        const selection = window.getSelection();
        const text = selection ? selection.toString() : "";
        
        if (text) {
            chrome.runtime.sendMessage({ action: "COPIED_TEXT_DETECTED", text: text });
        } else {
            chrome.runtime.sendMessage({ action: "COPY_TRIGGERED" });
        }
    } catch(e) {
        console.log("Copy detection failed", e);
    }
});

// LISTENER FOR PASTE ACTIONS AND GEMINI SCRAPING
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    
    // ACTION: STOP MACRO
    if (request.action === "STOP_MACRO") {
        if (isMacroRunning) {
            stopMacroRequest = true;
            sendResponse({ status: "macro_stopping" });
        } else {
            sendResponse({ status: "not_running" });
        }
        return true;
    }

    // ACTION: FULL AUTO MACRO (Complex Navigation & Paste)
    if (request.action === "FULL_AUTO_MACRO") {
        (async () => {
            try {
                if (isMacroRunning) {
                    sendResponse({ status: "already_running" });
                    return;
                }

                isMacroRunning = true;
                stopMacroRequest = false;
                
                sendResponse({ status: "macro_started" });

                // --- FOCUS LOGIC ---
                // Ensure we are focused on the grid before starting
                let target = lastFocusedElement;
                if (!target || !document.body.contains(target)) {
                    target = document.querySelector('#waffle-grid-container');
                    if(!target) target = document.querySelector('.grid-container');
                    if(!target) target = document.querySelector('body');
                }
                
                if (target) {
                    target.focus();
                    target.dispatchEvent(new MouseEvent('mousedown', {bubbles: true, cancelable: true, view: window}));
                    target.dispatchEvent(new MouseEvent('mouseup', {bubbles: true, cancelable: true, view: window}));
                }

                const commands = request.commands;
                
                for (const cmd of commands) {
                    if (stopMacroRequest) {
                        console.log("Macro Stopped by User");
                        break;
                    }

                    if (cmd.type === 'move') {
                        for (let i = 0; i < (cmd.count || 1); i++) {
                            if (stopMacroRequest) break;
                            simulateKey(cmd.key);
                            // Increased delay for stability (prevents skipping)
                            await sleep(150); 
                        }
                    } else if (cmd.type === 'paste') {
                        if (cmd.text) {
                            // Use pure Clipboard simulation
                            // This DOES NOT move the cursor, keeping mapping intact.
                            simulatePaste(cmd.text);
                            await sleep(350); // Allow Sheets to parse the CSV data
                        }
                    }
                }
                
                isMacroRunning = false;
                stopMacroRequest = false;
                
                // NOTIFY POPUP COMPLETION
                chrome.runtime.sendMessage({ action: "MACRO_COMPLETED" });

            } catch (e) {
                console.error("Macro Failed", e);
                isMacroRunning = false;
            }
        })();
        return true; 
    }

    // ACTION: TRY PASTE (Single Item)
    if (request.action === "TRY_PASTE") {
        try {
            if (lastFocusedElement) lastFocusedElement.focus();
            else window.focus();

            simulatePaste(request.text);
            sendResponse({ status: "pasted" });
        } catch (e) {
            console.error("Paste failed:", e);
            sendResponse({ status: "error" });
        }
        return true; 
    }

    // ACTION: PASTE TO GEMINI WEB
    if (request.action === "PASTE_TO_GEMINI_WEB") {
        try {
            // Check if generating
             if (document.querySelector('.stop-button') || document.querySelector('[aria-label="Stop generating"]')) {
                 sendResponse({ success: false, status: "generating_aborted" });
                 return true;
             }

            const promptText = request.text;
            // Robust selector for Gemini input (including new rich-textarea)
            const editorDiv = document.querySelector('div[contenteditable="true"]') || 
                              document.querySelector('div[role="textbox"]') ||
                              document.querySelector('rich-textarea div p');
            
            if (editorDiv) {
                editorDiv.focus();
                // Select all and insert text
                document.execCommand('selectAll', false, null);
                document.execCommand('insertText', false, promptText);
                
                // Wait for UI to update state (button enable)
                setTimeout(() => {
                    const sendBtn = document.querySelector('button[aria-label="Send message"]') || 
                                    document.querySelector('button[aria-label="Hantar mesej"]') || 
                                    document.querySelector('.send-button') ||
                                    document.querySelector('button[data-testid="send-button"]');
                    
                    if (sendBtn) {
                        if(!sendBtn.disabled && !sendBtn.getAttribute('aria-disabled')) {
                            sendBtn.click();
                        } else {
                            // Force click anyway just in case
                            sendBtn.click();
                        }
                    }
                }, 800); // Increased delay

                sendResponse({ success: true });
            } else {
                sendResponse({ success: false, error: "Editor not found" });
            }
        } catch(e) {
            sendResponse({ success: false, error: e.toString() });
        }
        return true;
    }

    // ACTION: GET RESPONSE FROM GEMINI WEB
    if (request.action === "GET_GEMINI_RESPONSE") {
        try {
            const candidates = document.querySelectorAll('.model-response-text, [data-message-id]'); 
            let foundText = "";
            for (let i = candidates.length - 1; i >= 0; i--) {
                const text = candidates[i].innerText || "";
                if (text.includes("[OBJEKTIF]") && text.includes("[AKTIVITI]")) {
                    foundText = text;
                    break;
                }
            }
            sendResponse({ text: foundText });
        } catch (e) {
            sendResponse({ text: "" });
        }
        return true;
    }
});

// --- SIMULATION HELPERS ---

function simulatePaste(text) {
    const target = document.activeElement || document.querySelector('.grid-container') || document.body;
    
    // --- KEY FIX FOR UNMERGED CELLS ---
    // Google Sheets unmerges cells if you paste text with newlines directly.
    // However, if the text is formatted as a quoted CSV string, Sheets treats it as a single cell value.
    // We check for newlines, tabs, or quotes, and wrap the entire string in double quotes.
    // Inner quotes must be escaped as double double-quotes ("").
    
    let processText = text;
    if (/[\n\r\t"]/.test(text)) {
        processText = `"${text.replace(/"/g, '""')}"`;
    }

    // Use the modern DataTransfer API to create a clipboard event
    const dt = new DataTransfer();
    dt.setData('text/plain', processText);
    
    const pasteEvent = new ClipboardEvent('paste', {
        bubbles: true,
        cancelable: true,
        clipboardData: dt,
        view: window
    });
    
    target.dispatchEvent(pasteEvent);
}

function simulateKey(keyName) {
    const target = document.activeElement || document.body;
    
    const keyMap = {
        'ArrowDown': 40,
        'ArrowUp': 38,
        'ArrowLeft': 37,
        'ArrowRight': 39,
        'Enter': 13,
        'Tab': 9,
        'Escape': 27
    };
    
    const code = keyMap[keyName] || 0;
    
    const eventOptions = {
        key: keyName,
        code: keyName,
        keyCode: code,
        which: code,
        bubbles: true,
        cancelable: true,
        view: window
    };

    target.dispatchEvent(new KeyboardEvent('keydown', eventOptions));
    target.dispatchEvent(new KeyboardEvent('keyup', eventOptions));
}