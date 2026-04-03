/* global document, Office, Word, fetch, window, navigator, URL, Blob, FileReader */

let currentButtons = [];
let currentPersonas = [];
let promptHistory = [];
let replaceHistory = [];
let appLang = 'fr'; 

const i18n = {
    fr: {
        tabCorr: "✏️ Correcteur", tabAss: "💬 Assistant", tabHelp: "❓ Aide",
        setSettings: "⚙️ Paramètres IA & Personas", setBaseLang: "Langue de base de l'IA :",
        setBaseHint: "ℹ️ Force l'IA à utiliser cette langue par défaut.",
        setModel: "Nom du modèle :", setKey: "Clé API (Optionnel) :", setStream: "⚡ Activer le Streaming API",
        setStreamHint: "ℹ️ Répond en direct mot à mot.",
        setPersHint: "📝 <b>Personas (Profils)</b> : Créez des profils pour dicter à l'IA comment se comporter.",
        btnAddPers: "➕ Ajouter ce Persona", btnExp: "💾 Exporter Config", btnImp: "📂 Importer Config",
        lblPersActive: "🎭 Persona actif (Style de l'IA) :", setBtns: "🛠️ Paramètres des boutons rapides",
        btnAddBtn: "➕ Ajouter ce bouton", guideSel: "🖱️ Sélectionnez le texte à modifier dans votre document.",
        lblTrad: "Langue de traduction :", hintTrad: "💡 Par défaut (si vide) : <b>Anglais</b>",
        hintOpts: "Allez dans les options pour ajouter des boutons d'accès rapide.",
        btnCont: "➡️ Continuer la rédaction (Générer la suite)", lblInst: "Instruction personnalisée :",
        hintInst: "💡 Par défaut (si vide) : <b>Corrige l'orthographe et la grammaire de ce texte.</b>",
        btnClear: "🗑️ Effacer l'instruction", noteShort: "⚠️ Note : Une sélection trop courte peut entraîner des choix inadaptés.",
        optCtx: "🔍 Contexte élargi (Lit le paragraphe complet)", optTrk: "📝 Activer le 'Suivi des modifications'",
        optExp: "💡 Ajouter une bulle d'explication", optFmt: "🎨 Conserver les formats multiples (Expérimental)", optHint: "ℹ️ Cocher ces éléments accroît le temps de réponse (tokens).",
        btnRun: "✨ Lancer l'IA", histReq: "🕒 Historique des requêtes (Instructions)",
        histRep: "⏪ Historique des textes remplacés", histHint: "💡 Cliquez sur un ancien texte pour copier l'original.",
        histLimit: "ℹ️ Taille de l'historique maximum : 10 éléments.",
        chatGuide: "💡 Discutez librement avec l'IA. Si vous surlignez du texte dans Word, l'IA l'utilisera automatiquement.",
        chatMsg: "Sentinel AI prêt, Posez-moi une question.",
        
        hTitle1: "🧠 Comment fonctionne l'outil ?", hSub1: "✏️ Onglet Correcteur :", hTxt1: "Modifie directement Word. Surlignez un texte, choisissez une instruction et l'IA remplacera votre sélection.",
        hSub2: "💬 Onglet Assistant :", hTxt2: "Chat intelligent. Si vous surlignez du texte, l'IA l'utilisera comme contexte.",
        hLocalTitle: "🖥️ IA Locale (Gratuit & Privé)", hLocalDesc: "Idéal pour la confidentialité. L'IA tourne sur votre ordinateur.",
        hLmsNote: "Modèle: Laissez vide. Clé API: Laissez vide. (Activez l'option CORS dans le serveur !)",
        hOllamaNote: "Modèle: llama3, mistral... (Nécessite de configurer OLLAMA_ORIGINS)",
        hCloudTitle: "☁️ IA en Ligne (Cloud)", hCloudDesc: "Pour utiliser les modèles les plus performants. Nécessite une clé API.",
        hClaudeInfo: "<b>Claude & Autres (via OpenRouter) :</b><br><code>https://openrouter.ai/api/v1/chat/completions</code><br><span style='color:var(--text-muted); font-size:10px;'>Modèles : anthropic/claude-3-opus, google/gemini-pro</span>",
        toast: "✅ Texte d'origine copié !"
    },
    en: {
        tabCorr: "✏️ Editor", tabAss: "💬 Assistant", tabHelp: "❓ Help",
        setSettings: "⚙️ AI Settings & Personas", setBaseLang: "AI Base Language:",
        setBaseHint: "ℹ️ Forces the AI to use this default language.",
        setModel: "Model Name:", setKey: "API Key (Optional):", setStream: "⚡ Enable Streaming API",
        setStreamHint: "ℹ️ Streams response word by word.",
        setPersHint: "📝 <b>Personas (Profiles)</b>: Create profiles to dictate AI behavior.",
        btnAddPers: "➕ Add Persona", btnExp: "💾 Export Config", btnImp: "📂 Import Config",
        lblPersActive: "🎭 Active Persona (AI Style):", setBtns: "🛠️ Quick Buttons Settings",
        btnAddBtn: "➕ Add Button", guideSel: "🖱️ Select the text to modify in your document.",
        lblTrad: "Target Translation Language:", hintTrad: "💡 Default (if empty): <b>English</b>",
        hintOpts: "Go to settings to add custom quick access buttons.",
        btnCont: "➡️ Continue Writing (Generate Next)", lblInst: "Custom Instruction:",
        hintInst: "💡 Default (if empty): <b>Fix spelling and grammar.</b>",
        btnClear: "🗑️ Clear instruction", noteShort: "⚠️ Note: A very short selection may lead to poor context choices.",
        optCtx: "🔍 Extended Context (Reads paragraph)", optTrk: "📝 Enable 'Track Changes'",
        optExp: "💡 Add explanation bubble", optFmt: "🎨 Keep mixed formatting (Experimental)", optHint: "ℹ️ Checking these increases API requests and time (tokens).",
        btnRun: "✨ Run AI", histReq: "🕒 Prompt History",
        histRep: "⏪ Replaced Text History", histHint: "💡 Click on an old text to copy the original.",
        histLimit: "ℹ️ Maximum history size: 10 items.",
        chatGuide: "💡 Chat freely. If you highlight text in Word, the AI will use it as context.",
        chatMsg: "Sentinel AI ready, Ask me a question.",
        
        hTitle1: "🧠 How does it work?", hSub1: "✏️ Editor Tab:", hTxt1: "Modifies Word directly. Highlight text, choose a prompt, and the AI replaces it.",
        hSub2: "💬 Assistant Tab:", hTxt2: "Smart chat. Highlights in Word are used as context.",
        hLocalTitle: "🖥️ Local AI (Free & Private)", hLocalDesc: "Best for privacy. The AI runs on your computer.",
        hLmsNote: "Model: Leave empty. API Key: Leave empty. (Enable CORS in the server!)",
        hOllamaNote: "Model: llama3, mistral... (Requires setting up OLLAMA_ORIGINS)",
        hCloudTitle: "☁️ Cloud AI (Online)", hCloudDesc: "For the most powerful AIs. Requires a secret API key.",
        hClaudeInfo: "<b>Claude & Others (via OpenRouter):</b><br><code>https://openrouter.ai/api/v1/chat/completions</code><br><span style='color:var(--text-muted); font-size:10px;'>Models: anthropic/claude-3-opus, google/gemini-pro</span>",
        toast: "✅ Original text copied!"
    }
};

const baseButtonsData = {
    "btn_pro": { title: "👔 Pro", prompt: "Reformule ce texte en {langue_base} pour qu'il soit très professionnel, formel et adapté à un contexte d'entreprise." },
    "btn_concis": { title: "✂️ Concis", prompt: "Raccourcis ce texte en {langue_base} pour le rendre plus direct et concis." },
    "btn_fluide": { title: "✍️ Fluide", prompt: "Améliore la fluidité et le style de ce texte en {langue_base}." },
    "btn_trad": { title: "🌍 Traduire", prompt: "Traduis ce texte en {langue} de manière naturelle et idiomatique." }
};

const basePersonasData = {
    "pers_std": { name: "Standard (Neutre)", prompt: "Tu es un processeur de texte professionnel et neutre." },
    "pers_avo": { name: "Juridique", prompt: "Tu es un expert juridique. Utilise un vocabulaire formel, précis et légal." },
    "pers_web": { name: "Créatif / Web", prompt: "Tu es un rédacteur web créatif. Ton ton est engageant, moderne et dynamique." }
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    setupLanguage();
    setupTheme(); 
    chargerParametres();
    initData();
    setupTabs();

    const inputs = ["apiUrl", "modelName", "apiKey", "customPrompt", "targetLanguage", "baseLanguage", "activePersonaSelect"];
    inputs.forEach(id => document.getElementById(id).addEventListener("input", sauvegarderParametres));
    const checkboxes = ["chkContext", "chkTrack", "chkExplain", "chkFormat", "chkStreaming"];
    checkboxes.forEach(id => document.getElementById(id).addEventListener("change", sauvegarderParametres));

    document.getElementById("btnClearPrompt").onclick = () => { document.getElementById("customPrompt").value = ""; sauvegarderParametres(); };
    document.getElementById("btnAddButton").onclick = addNewButton;
    document.getElementById("btnAddPersona").onclick = addNewPersona;
    
    document.getElementById("btnExportConfig").onclick = exportConfig;
    document.getElementById("btnImportClick").onclick = () => document.getElementById("btnImportConfig").click();
    document.getElementById("btnImportConfig").addEventListener("change", importConfig);

    document.getElementById("btnCorriger").onclick = () => handleIAAction('replace');
    document.getElementById("btnContinueText").onclick = () => handleIAAction('continue');
    
    document.getElementById("btnSendChat").onclick = handleChat;
    document.getElementById("chatInput").addEventListener('keypress', function (e) {
        if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleChat(); }
    });
  }
});

function setupLanguage() {
    appLang = window.localStorage.getItem("ai_appLang") || "fr";
    updateUITexts();
    
    const langBtn = document.getElementById("langToggle");
    langBtn.innerText = appLang === "fr" ? "🇬🇧" : "🇫🇷";
    
    langBtn.onclick = () => {
        appLang = appLang === "fr" ? "en" : "fr";
        window.localStorage.setItem("ai_appLang", appLang);
        langBtn.innerText = appLang === "fr" ? "🇬🇧" : "🇫🇷";
        updateUITexts();
    };
}

function updateUITexts() {
    const dict = i18n[appLang];
    document.querySelectorAll("[data-i18n]").forEach(el => {
        const key = el.getAttribute("data-i18n");
        if (dict[key]) el.innerHTML = dict[key];
    });
    
    document.getElementById("chatInput").placeholder = appLang === "fr" ? "Tapez votre message... (Entrée pour envoyer)" : "Type your message... (Enter to send)";
    document.getElementById("customPrompt").placeholder = appLang === "fr" ? "Tapez votre instruction ici..." : "Type your instruction here...";
}

function setupTheme() {
    const isDark = window.localStorage.getItem("ai_darkMode") === "true";
    if (isDark) { document.body.classList.add("dark-mode"); }
    
    const themeBtn = document.getElementById("themeToggle");
    themeBtn.onclick = () => {
        document.body.classList.toggle("dark-mode");
        const darkOn = document.body.classList.contains("dark-mode");
        window.localStorage.setItem("ai_darkMode", darkOn);
    };
}

function setupTabs() {
    document.getElementById("btnTabCorrector").onclick = () => switchTab("tabCorrector", "btnTabCorrector");
    document.getElementById("btnTabAssistant").onclick = () => switchTab("tabAssistant", "btnTabAssistant");
    document.getElementById("btnTabHelp").onclick = () => switchTab("tabHelp", "btnTabHelp");
}

function switchTab(tabId, btnId) {
    document.querySelectorAll(".tab-content").forEach(el => el.classList.remove("active"));
    document.querySelectorAll(".tab-btn").forEach(el => el.classList.remove("active"));
    document.getElementById(tabId).classList.add("active");
    document.getElementById(btnId).classList.add("active");
}

function exportConfig() {
    const config = {
        apiUrl: window.localStorage.getItem("ai_apiUrl"),
        modelName: window.localStorage.getItem("ai_modelName"),
        customButtons: currentButtons,
        personas: currentPersonas,
        promptHistory: promptHistory,
        baseLanguage: window.localStorage.getItem("ai_baseLanguage"),
        chkStreaming: document.getElementById("chkStreaming").checked,
        chkFormat: document.getElementById("chkFormat").checked
    };
    const blob = new Blob([JSON.stringify(config, null, 2)], {type: "application/json"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = "smartscribe_config.json";
    a.click();
    showToast(appLang === "fr" ? "💾 Configuration exportée !" : "💾 Config exported!");
}

function importConfig(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const config = JSON.parse(e.target.result);
            if (config.apiUrl) window.localStorage.setItem("ai_apiUrl", config.apiUrl);
            if (config.modelName) window.localStorage.setItem("ai_modelName", config.modelName);
            if (config.customButtons) window.localStorage.setItem("ai_customButtons", JSON.stringify(config.customButtons));
            if (config.personas) window.localStorage.setItem("ai_personas", JSON.stringify(config.personas));
            if (config.promptHistory) window.localStorage.setItem("ai_promptHistory", JSON.stringify(config.promptHistory));
            if (config.baseLanguage) window.localStorage.setItem("ai_baseLanguage", config.baseLanguage);
            if (config.chkStreaming !== undefined) window.localStorage.setItem("ai_chkStreaming", config.chkStreaming);
            if (config.chkFormat !== undefined) window.localStorage.setItem("ai_chkFormat", config.chkFormat);
            
            chargerParametres(); initData();
            showToast(appLang === "fr" ? "📂 Configuration importée !" : "📂 Config imported!");
        } catch(err) { alert("Invalid config file."); }
    };
    reader.readAsText(file);
    event.target.value = ""; 
}

function initData() {
    try { currentButtons = JSON.parse(window.localStorage.getItem("ai_customButtons")) || []; } catch(e) { currentButtons = []; }
    if (currentButtons.length < 4 || !currentButtons[0].id) { currentButtons = Object.keys(baseButtonsData).map(k => ({ id: k, ...baseButtonsData[k] })); }
    
    try { currentPersonas = JSON.parse(window.localStorage.getItem("ai_personas")) || []; } catch(e) { currentPersonas = []; }
    if (currentPersonas.length === 0) { currentPersonas = Object.keys(basePersonasData).map(k => ({ id: k, ...basePersonasData[k] })); }

    try { promptHistory = JSON.parse(window.localStorage.getItem("ai_promptHistory")) || []; } catch(e) {}
    try { replaceHistory = JSON.parse(window.localStorage.getItem("ai_replaceHistory")) || []; } catch(e) {}
    renderUI();
}

function renderUI() { renderButtons(); renderPersonas(); renderHistories(); }

function renderButtons() {
    const quickContainer = document.getElementById("quickActionsContainer");
    quickContainer.innerHTML = "";
    currentButtons.forEach((btn) => {
        const b = document.createElement("button"); b.className = "btn-quick"; b.innerText = btn.title; b.title = btn.prompt;
        b.onclick = () => {
            let langTarget = document.getElementById("targetLanguage").value.trim() || (appLang === 'fr' ? "anglais" : "English");
            let baseLang = document.getElementById("baseLanguage").value.trim() || (appLang === 'fr' ? "Français" : "English");
            
            let finalPrompt = btn.prompt.replace(/\{langue\}/g, langTarget).replace(/\{langue_base\}/g, baseLang);
            document.getElementById("customPrompt").value = finalPrompt;
            sauvegarderParametres();
        };
        quickContainer.appendChild(b);
    });

    const editorList = document.getElementById("settingsButtonsList"); editorList.innerHTML = "";
    currentButtons.forEach((btn, index) => {
        const row = document.createElement("div"); row.className = "edit-row";
        const tInp = document.createElement("input"); tInp.value = btn.title; tInp.onchange = (e) => { currentButtons[index].title = e.target.value; saveDataAndRender(); };
        const pInp = document.createElement("textarea"); pInp.value = btn.prompt; pInp.onchange = (e) => { currentButtons[index].prompt = e.target.value; saveDataAndRender(); };
        row.appendChild(tInp); row.appendChild(pInp);
        if (btn.id && baseButtonsData[btn.id]) {
            const res = document.createElement("button"); res.innerHTML = "🔄"; res.className = "btn-reset";
            res.onclick = () => { currentButtons[index] = {id: btn.id, ...baseButtonsData[btn.id]}; saveDataAndRender(); };
            row.appendChild(res);
        } else {
            const del = document.createElement("button"); del.innerHTML = "❌"; del.className = "btn-delete";
            del.onclick = () => { currentButtons.splice(index, 1); saveDataAndRender(); };
            row.appendChild(del);
        }
        editorList.appendChild(row);
    });
}

function addNewButton() {
    const t = document.getElementById("newBtnTitle").value.trim(); const p = document.getElementById("newBtnPrompt").value.trim();
    if (t && p) { currentButtons.push({ id: "cb_"+Date.now(), title: t, prompt: p }); document.getElementById("newBtnTitle").value = ""; document.getElementById("newBtnPrompt").value = ""; saveDataAndRender(); }
}

function renderPersonas() {
    const select = document.getElementById("activePersonaSelect"); const savedActive = window.localStorage.getItem("ai_activePersonaSelect");
    select.innerHTML = "";
    currentPersonas.forEach(p => {
        const opt = document.createElement("option"); opt.value = p.id; opt.innerText = p.name;
        if(p.id === savedActive) opt.selected = true; select.appendChild(opt);
    });

    const editorList = document.getElementById("settingsPersonasList"); editorList.innerHTML = "";
    currentPersonas.forEach((p, index) => {
        const row = document.createElement("div"); row.className = "edit-row";
        const tInp = document.createElement("input"); tInp.value = p.name; tInp.onchange = (e) => { currentPersonas[index].name = e.target.value; saveDataAndRender(); };
        const pInp = document.createElement("textarea"); pInp.value = p.prompt; pInp.onchange = (e) => { currentPersonas[index].prompt = e.target.value; saveDataAndRender(); };
        row.appendChild(tInp); row.appendChild(pInp);
        if (p.id && basePersonasData[p.id]) {
            const res = document.createElement("button"); res.innerHTML = "🔄"; res.className = "btn-reset";
            res.onclick = () => { currentPersonas[index] = {id: p.id, ...basePersonasData[p.id]}; saveDataAndRender(); }; row.appendChild(res);
        } else {
            const del = document.createElement("button"); del.innerHTML = "❌"; del.className = "btn-delete";
            del.onclick = () => { currentPersonas.splice(index, 1); saveDataAndRender(); }; row.appendChild(del);
        }
        editorList.appendChild(row);
    });
}

function addNewPersona() {
    const t = document.getElementById("newPersonaTitle").value.trim(); const p = document.getElementById("newPersonaPrompt").value.trim();
    if (t && p) { currentPersonas.push({ id: "cp_"+Date.now(), name: t, prompt: p }); document.getElementById("newPersonaTitle").value = ""; document.getElementById("newPersonaPrompt").value = ""; saveDataAndRender(); }
}

function showToast(msgText) {
    const toast = document.getElementById("toastMsg");
    toast.innerText = msgText || i18n[appLang].toast;
    toast.classList.add("show");
    setTimeout(() => { toast.classList.remove("show"); }, 2500);
}

function renderHistories() {
    const pList = document.getElementById("promptHistoryList"); pList.innerHTML = "";
    [...promptHistory].reverse().forEach(h => {
        const div = document.createElement("div"); div.className = "history-item"; div.innerHTML = `<span class="history-text">${h}</span>`;
        div.onclick = () => { document.getElementById("customPrompt").value = h; sauvegarderParametres(); }; pList.appendChild(div);
    });

    const rList = document.getElementById("replaceHistoryList"); rList.innerHTML = "";
    [...replaceHistory].reverse().forEach(h => {
        const div = document.createElement("div"); div.className = "history-item"; 
        div.title = appLang === "fr" ? "Cliquez pour copier" : "Click to copy";
        const displayOld = h.shortOld || h.old; const displayNew = h.shortNew || h.new; const copyText = h.fullOld || h.old;
        div.innerHTML = `<span class="history-text" style="text-decoration:line-through; color:#a00;">${displayOld}</span><span class="history-text" style="color:#0a0;">${displayNew}</span>`;
        div.onclick = () => { navigator.clipboard.writeText(copyText).then(() => { showToast(); }).catch(err => console.error(err)); };
        rList.appendChild(div);
    });
}

function addPromptHistory(text) {
    if(!text || text === "") return;
    promptHistory = promptHistory.filter(h => h !== text); promptHistory.push(text);
    if(promptHistory.length > 10) promptHistory.shift(); saveDataAndRender();
}

function addReplaceHistory(oldTxt, newTxt) {
    replaceHistory.push({ fullOld: oldTxt, fullNew: newTxt, shortOld: oldTxt.length > 50 ? oldTxt.substring(0,50)+"..." : oldTxt, shortNew: newTxt.length > 50 ? newTxt.substring(0,50)+"..." : newTxt });
    if(replaceHistory.length > 10) replaceHistory.shift(); saveDataAndRender();
}

function saveDataAndRender() {
    window.localStorage.setItem("ai_customButtons", JSON.stringify(currentButtons));
    window.localStorage.setItem("ai_personas", JSON.stringify(currentPersonas));
    window.localStorage.setItem("ai_promptHistory", JSON.stringify(promptHistory));
    window.localStorage.setItem("ai_replaceHistory", JSON.stringify(replaceHistory));
    renderUI();
}

function sauvegarderParametres() {
    window.localStorage.setItem("ai_apiUrl", document.getElementById("apiUrl").value);
    window.localStorage.setItem("ai_modelName", document.getElementById("modelName").value);
    window.localStorage.setItem("ai_apiKey", document.getElementById("apiKey").value);
    window.localStorage.setItem("ai_customPrompt", document.getElementById("customPrompt").value);
    window.localStorage.setItem("ai_targetLanguage", document.getElementById("targetLanguage").value);
    window.localStorage.setItem("ai_baseLanguage", document.getElementById("baseLanguage").value);
    window.localStorage.setItem("ai_activePersonaSelect", document.getElementById("activePersonaSelect").value);
    window.localStorage.setItem("ai_chkContext", document.getElementById("chkContext").checked);
    window.localStorage.setItem("ai_chkTrack", document.getElementById("chkTrack").checked);
    window.localStorage.setItem("ai_chkExplain", document.getElementById("chkExplain").checked);
    window.localStorage.setItem("ai_chkFormat", document.getElementById("chkFormat").checked);
    window.localStorage.setItem("ai_chkStreaming", document.getElementById("chkStreaming").checked);
}

function chargerParametres() {
    document.getElementById("apiUrl").value = window.localStorage.getItem("ai_apiUrl") || "http://localhost:1234/v1/chat/completions";
    document.getElementById("modelName").value = window.localStorage.getItem("ai_modelName") || "mistral-7b";
    document.getElementById("apiKey").value = window.localStorage.getItem("ai_apiKey") || "";
    document.getElementById("customPrompt").value = window.localStorage.getItem("ai_customPrompt") || "";
    document.getElementById("targetLanguage").value = window.localStorage.getItem("ai_targetLanguage") || "";
    
    let savedBaseLang = window.localStorage.getItem("ai_baseLanguage");
    if(!savedBaseLang) { savedBaseLang = (appLang === 'fr' ? "Français" : "English"); }
    document.getElementById("baseLanguage").value = savedBaseLang;

    document.getElementById("activePersonaSelect").value = window.localStorage.getItem("ai_activePersonaSelect") || "";
    document.getElementById("chkContext").checked = window.localStorage.getItem("ai_chkContext") === "true";
    document.getElementById("chkTrack").checked = window.localStorage.getItem("ai_chkTrack") === "true";
    document.getElementById("chkExplain").checked = window.localStorage.getItem("ai_chkExplain") === "true";
    document.getElementById("chkFormat").checked = window.localStorage.getItem("ai_chkFormat") === "true";
    document.getElementById("chkStreaming").checked = window.localStorage.getItem("ai_chkStreaming") === "true";
}

async function callAI(systemMsg, userMsg, onChunkCallback = null) {
    const apiUrl = document.getElementById("apiUrl").value;
    const modelName = document.getElementById("modelName").value;
    const apiKey = document.getElementById("apiKey").value;
    const isStreaming = document.getElementById("chkStreaming").checked;
    
    const headers = { "Content-Type": "application/json" };
    if (apiKey) headers["Authorization"] = `Bearer ${apiKey}`;

    const payload = {
        model: modelName,
        messages: [{ role: "system", content: systemMsg }, { role: "user", content: userMsg }],
        temperature: 0.2,
        stream: isStreaming
    };

    const response = await fetch(apiUrl, { method: "POST", headers: headers, body: JSON.stringify(payload) });
    
    if (!response.ok) {
        const err = await response.json();
        throw new Error(err.error?.message || "Erreur de connexion serveur");
    }

    if (!isStreaming) {
        const data = await response.json();
        return data.choices[0].message.content.trim();
    } else {
        const reader = response.body.getReader();
        const decoder = new TextDecoder("utf-8");
        let fullText = "";

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            const chunk = decoder.decode(value, { stream: true });
            const lines = chunk.split("\n");
            
            for (const line of lines) {
                if (line.trim().startsWith("data: ") && line.trim() !== "data: [DONE]") {
                    try {
                        const data = JSON.parse(line.trim().substring(6));
                        if (data.choices && data.choices[0].delta && data.choices[0].delta.content) {
                            fullText += data.choices[0].delta.content;
                            if (onChunkCallback) onChunkCallback(fullText);
                        }
                    } catch(e) {}
                }
            }
        }
        return fullText.trim();
    }
}

async function handleChat() {
    const inputEl = document.getElementById("chatInput");
    const msg = inputEl.value.trim();
    if (!msg) return;
    
    appendChatMessage(msg, "user");
    inputEl.value = "";
    const box = document.getElementById("chatBox");

    const loadingDiv = document.createElement("div"); loadingDiv.className = "chat-msg ai"; loadingDiv.innerText = "✍️...";
    box.appendChild(loadingDiv); box.scrollTop = box.scrollHeight;

    let contextText = "";
    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        if(selection.text.trim() !== "") contextText = selection.text;
    });

    const activePersonaId = document.getElementById("activePersonaSelect").value;
    const activePersona = currentPersonas.find(p => p.id === activePersonaId);
    let personaPrompt = activePersona ? activePersona.prompt : "Tu es un assistant neutre.";

    let baseLang = document.getElementById("baseLanguage").value.trim() || "Français";

    let systemMsg = personaPrompt + ` Règle ABSOLUE : Tu dois IMPÉRATIVEMENT répondre dans la langue de l'utilisateur ou en ${baseLang}.`;
    let userMsg = msg;
    if(contextText) userMsg += `\n\n[CONTEXTE DU DOCUMENT SÉLECTIONNÉ] : "${contextText}"`;

    try {
        let responseText = await callAI(systemMsg, userMsg, (streamingText) => {
            loadingDiv.innerText = streamingText; 
            box.scrollTop = box.scrollHeight;
        });
        loadingDiv.innerText = responseText; 
    } catch(err) {
        box.removeChild(loadingDiv);
        appendChatMessage("❌ Erreur : " + err.message, "ai");
    }
}

async function handleIAAction(mode) {
  const btn = document.getElementById("btnCorriger");
  const spinner = document.getElementById("btnSpinner");
  const btnText = document.getElementById("btnText");
  const statusMsg = document.getElementById("statusMsg");
  const streamPreview = document.getElementById("streamPreview");
  
  const optContext = document.getElementById("chkContext").checked;
  const optTrack = document.getElementById("chkTrack").checked;
  const optExplain = document.getElementById("chkExplain").checked;
  const optFormat = document.getElementById("chkFormat").checked;
  const isStreaming = document.getElementById("chkStreaming").checked;

  let instructionText = document.getElementById("customPrompt").value.trim() || (appLang === "fr" ? "Corrige l'orthographe et la grammaire." : "Fix spelling and grammar.");
  if(mode === 'replace') addPromptHistory(instructionText);

  const activePersonaId = document.getElementById("activePersonaSelect").value;
  const activePersona = currentPersonas.find(p => p.id === activePersonaId);
  const personaPrompt = activePersona ? activePersona.prompt : "";

  let baseLang = document.getElementById("baseLanguage").value.trim() || "Français";

  btn.disabled = true; spinner.style.display = "block"; btnText.innerText = "..."; statusMsg.innerText = "";
  if (isStreaming) { streamPreview.style.display = "block"; streamPreview.innerText = "✍️..."; }

  await Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      selection.load("text");
      
      let htmlResult;
      if (optFormat && mode === 'replace') {
          htmlResult = selection.getHtml();
      }

      let contextResult;
      if (optContext) {
        const paragraphs = selection.paragraphs; 
        paragraphs.load("text"); 
        contextResult = paragraphs;
      }
      
      await context.sync();

      if (selection.text.trim() === "") throw new Error(appLang === "fr" ? "Veuillez d'abord surligner du texte !" : "Please select text first!");
      const originalText = selection.text;
      
      let contextText = "";
      if (optContext && contextResult) {
          contextText = contextResult.items.map(p => p.text).join(" ");
      }

      let originalHtml = "";
      if (optFormat && htmlResult) {
          originalHtml = htmlResult.value;
      }

      let systemPrompt = personaPrompt + ` Règle CRUCIALE : Tu dois IMPÉRATIVEMENT écrire ta réponse en ${baseLang}. NE TRADUIS PAS dans une autre langue, sauf si l'instruction contient explicitement le mot 'Traduis' ou 'Translate'. `;
      let userPrompt = "";

      if (mode === 'continue') {
          systemPrompt += "L'utilisateur veut que tu écrives la SUITE LOGIQUE du texte qu'il te donne. Adopte le même ton, le même style, et génère le prochain paragraphe de manière fluide. Renvoie UNIQUEMENT le nouveau texte généré, sans répéter l'original, sans guillemets.";
          userPrompt = `Texte actuel :\n"${originalText}"\n\nGénère la suite :`;
      } else {
          userPrompt = `Instruction : ${instructionText} (IMPORTANT : Réponds en ${baseLang} si le texte n'est pas à traduire)\n\n`;
          if (optContext && contextText) userPrompt += `Pour t'aider, voici le contexte global : "${contextText}"\n\n`;
          
          if (optFormat) {
              systemPrompt += " Règle ABSOLUE : Tu vas recevoir un code source HTML contenant le texte. Tu DOIS IMPÉRATIVEMENT renvoyer le résultat sous forme de code HTML strict en conservant exactement toutes les balises de style (couleurs, polices, tailles, balises MSO) d'origine. Ne casse surtout pas la mise en forme.";
              userPrompt += `Code HTML exact à corriger : \n\`\`\`html\n${originalHtml}\n\`\`\`\n\nRenvoye uniquement le code HTML final corrigé :`;
          } else if (optExplain) {
              systemPrompt += "Sépare ta réponse en deux blocs avec ces balises exactes : <TEXTE> pour la correction finale, et <EXPLICATION> pour expliquer.";
              userPrompt += `Texte exact à corriger : "${originalText}"\n\nRéponds avec <TEXTE>...</TEXTE> et <EXPLICATION>...</EXPLICATION>.`;
          } else {
              systemPrompt += "Règle ABSOLUE : Renvoie UNIQUEMENT le texte final corrigé. Aucun préfixe, aucun guillemet.";
              userPrompt += `Texte exact à corriger : "${originalText}"\n\nTexte final :`;
          }
      }

      let reponseLLM = await callAI(systemPrompt, userPrompt, (streamingText) => {
          streamPreview.innerText = streamingText; 
          streamPreview.scrollTop = streamPreview.scrollHeight;
      });
      
      let texteFinal = reponseLLM;
      let explicationText = "";

      if (mode === 'replace' && !optFormat && optExplain) {
        const tMatch = reponseLLM.match(/<TEXTE>([\s\S]*?)<\/TEXTE>/i);
        const eMatch = reponseLLM.match(/<EXPLICATION>([\s\S]*?)<\/EXPLICATION>/i);
        if (tMatch) texteFinal = tMatch[1].trim();
        if (eMatch) explicationText = eMatch[1].trim();
      }

      texteFinal = texteFinal.replace(/^(Correction|Texte corrigé|Voici le texte|Texte final|Texte)[\s]*:[\s]*/i, "").trim();
      if (!optFormat) { texteFinal = texteFinal.replace(/^["']|["']$/g, "").trim(); }

      if (optTrack) context.document.changeTrackingMode = "TrackAll"; 
      else context.document.changeTrackingMode = "Off"; 

      let insertedRange;
      if (mode === 'continue') {
          insertedRange = selection.insertText(" " + texteFinal, Word.InsertLocation.after);
      } else {
          if (optFormat) {
              // Nettoyage des balises markdown si l'IA en a rajouté
              let cleanHtml = texteFinal.replace(/```html\n?/gi, "").replace(/```\n?/g, "").trim();
              insertedRange = selection.insertHtml(cleanHtml, Word.InsertLocation.replace);
          } else {
              insertedRange = selection.insertText(texteFinal, Word.InsertLocation.replace);
          }
          addReplaceHistory(originalText, texteFinal);
      }
      
      if (mode === 'replace' && optExplain && explicationText && !optFormat) insertedRange.insertComment(`🤖 IA : ${explicationText}`);

      await context.sync();
      statusMsg.style.color = "green"; statusMsg.innerText = i18n[appLang].toast;

    } catch (error) {
      statusMsg.style.color = "#d83b01";
      statusMsg.innerText = error.message;
    } finally {
      btn.disabled = false; spinner.style.display = "none"; btnText.innerText = i18n[appLang].btnRun;
      streamPreview.style.display = "none"; streamPreview.innerText = "";
      setTimeout(() => statusMsg.innerText = "", 4000);
    }
  });
}