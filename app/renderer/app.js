(() => {
  const STORAGE_KEY = "notations.state.v1";
  const PAGE_DPI = 96;
  const EDITOR_MIN_HEIGHT_PX = 420;
  const EDITOR_SCROLL_BUFFER_PX = 120;
  const FONT_SIZE_MIN = 10;
  const FONT_SIZE_MAX = 100;
  const FONT_SIZE_STEP = 2;
  const FONT_SIZE_OPTIONS = Array.from(
    { length: Math.floor((FONT_SIZE_MAX - FONT_SIZE_MIN) / FONT_SIZE_STEP) + 1 },
    (_, index) => FONT_SIZE_MIN + index * FONT_SIZE_STEP
  );

  const PAPER_PRESETS = {
    letter: { id: "letter", label: "US Letter", widthIn: 8.5, heightIn: 11, printKeyword: "Letter" },
    a4: { id: "a4", label: "A4", widthIn: 8.2677, heightIn: 11.6929, printKeyword: "A4" }
  };

  const FONT_PRESETS = {
    monospace: '"Inconsolata", "SFMono-Regular", "Menlo", "Consolas", monospace',
    sans: '"Helvetica Neue", "Arial", sans-serif',
    serif: '"AGP", "Times New Roman", serif'
  };

  const DEFAULT_SETTINGS = {
    paperSize: "letter",
    fontFamily: "monospace",
    fontSize: 18,
    lineHeight: 1.5,
    margins: { top: 0.42, right: 1.12, bottom: 0.75, left: 0.42 }
  };

  const EMPTY_PLACEHOLDER = [
    "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.",
    "At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.",
    "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat."
  ].join(" ");

  const SAMPLE_TEXT = [
    "This is the first paragraph of the text. Its main purpose is to provide quick context for the contents of this notation.",
    "",
    "It also works as a visual element in the grid."
  ].join("\n");

  function nowISO() {
    return new Date().toISOString();
  }

  function clamp(n, min, max) {
    return Math.max(min, Math.min(max, n));
  }

  function normalizeFontSize(value, fallback = DEFAULT_SETTINGS.fontSize) {
    const source = Number(value);
    const base = Number.isFinite(source) ? source : Number(fallback);
    const clamped = clamp(base, FONT_SIZE_MIN, FONT_SIZE_MAX);
    const stepped = Math.round(clamped / FONT_SIZE_STEP) * FONT_SIZE_STEP;
    return clamp(stepped, FONT_SIZE_MIN, FONT_SIZE_MAX);
  }

  function normalizeMarginValue(value, fallback) {
    const n = Number(value);
    if (!Number.isFinite(n) || n < 0) return Number(fallback);
    return Number(n.toFixed(2));
  }

  function normalizeMargins(raw, fallback = DEFAULT_SETTINGS.margins) {
    const source = raw && typeof raw === "object" ? raw : {};
    return {
      top: normalizeMarginValue(source.top, fallback.top),
      right: normalizeMarginValue(source.right, fallback.right),
      bottom: normalizeMarginValue(source.bottom, fallback.bottom),
      left: normalizeMarginValue(source.left, fallback.left)
    };
  }

  function ensureSheetMargins(sheet, fallback = DEFAULT_SETTINGS.margins) {
    if (!sheet || sheet.kind !== "sheet") {
      return normalizeMargins(null, fallback);
    }
    sheet.margins = normalizeMargins(sheet.margins, fallback);
    return sheet.margins;
  }

  let titleMeasureEl = null;
  let titleMeasureCanvas = null;
  let titleMeasureCtx = null;

  function ensureTitleMeasurers() {
    if (!titleMeasureCanvas) {
      titleMeasureCanvas = document.createElement("canvas");
      titleMeasureCtx = titleMeasureCanvas.getContext("2d");
    }
    if (!titleMeasureEl) {
      titleMeasureEl = document.createElement("div");
      titleMeasureEl.setAttribute("aria-hidden", "true");
      Object.assign(titleMeasureEl.style, {
        position: "fixed",
        left: "-10000px",
        top: "-10000px",
        visibility: "hidden",
        pointerEvents: "none",
        zIndex: "-1",
        whiteSpace: "normal",
        wordBreak: "break-word"
      });
      document.body.appendChild(titleMeasureEl);
    }
  }

  function getUIFontFamily() {
    const v = getComputedStyle(document.documentElement).getPropertyValue("--ui-font-family").trim();
    return v || '"INDP", "Helvetica Neue", Arial, sans-serif';
  }

  function measureSingleLineWidth(text, sizePx, weight, family) {
    ensureTitleMeasurers();
    if (!titleMeasureCtx) return 0;
    titleMeasureCtx.font = `${weight} ${sizePx}px ${family}`;
    return titleMeasureCtx.measureText(text).width;
  }

  function fitSingleLineTitle(text, cfg) {
    const family = getUIFontFamily();
    let lo = cfg.min;
    let hi = cfg.max;
    let best = cfg.min;
    while (lo <= hi) {
      const mid = Math.floor((lo + hi) / 2);
      const width = measureSingleLineWidth(text, mid, cfg.weight, family);
      if (width <= cfg.maxWidth) {
        best = mid;
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
    return best;
  }

  function fitMultilineTitle(text, cfg) {
    ensureTitleMeasurers();
    const family = getUIFontFamily();
    let lo = cfg.min;
    let hi = cfg.max;
    let best = cfg.min;

    while (lo <= hi) {
      const mid = Math.floor((lo + hi) / 2);
      titleMeasureEl.style.width = `${cfg.maxWidth}px`;
      titleMeasureEl.style.fontSize = `${mid}px`;
      titleMeasureEl.style.fontWeight = String(cfg.weight);
      titleMeasureEl.style.lineHeight = String(cfg.lineHeight);
      titleMeasureEl.style.fontFamily = family;
      titleMeasureEl.textContent = text;

      if (titleMeasureEl.scrollHeight <= cfg.maxHeight) {
        best = mid;
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }

    return best;
  }

  function applyDynamicCardTitleSizing(card, titleText) {
    const titleEl = card.querySelector(".card-title");
    if (!titleEl) return;
    const clean = String(titleText || "").trim();
    if (!clean) return;

    if (card.classList.contains("card-date")) {
      const size = fitSingleLineTitle(clean, {
        min: 36,
        max: 70,
        maxWidth: 168,
        weight: 600
      });
      titleEl.style.fontSize = `${size}px`;
      titleEl.style.lineHeight = "1";
      return;
    }

    if (!card.classList.contains("card-title-only")) return;

    if (clean.length <= 4) {
      const size = fitSingleLineTitle(clean, {
        min: 52,
        max: 88,
        maxWidth: 168,
        weight: 600
      });
      titleEl.style.fontSize = `${size}px`;
      titleEl.style.lineHeight = "0.9";
      return;
    }

    const size = fitMultilineTitle(clean, {
      min: 20,
      max: 56,
      maxWidth: 168,
      maxHeight: 120,
      lineHeight: 1.1,
      weight: 600
    });
    titleEl.style.fontSize = `${size}px`;
    titleEl.style.lineHeight = "1.1";
  }

  function safeName(name, fallback = "Untitled") {
    const v = String(name || "").trim();
    return v ? v : fallback;
  }

  function normalizeTagName(raw) {
    return String(raw == null ? "" : raw)
      .trim()
      .replace(/\s+/g, " ");
  }

  function findTagIndex(tags, rawTag) {
    const key = normalizeTagName(rawTag).toLocaleLowerCase();
    if (!key) return -1;
    return tags.findIndex((tag) => String(tag).toLocaleLowerCase() === key);
  }

  function ensureSheetTags(sheet) {
    if (!sheet || sheet.kind !== "sheet") return [];
    if (!Array.isArray(sheet.tags)) {
      sheet.tags = [];
      return sheet.tags;
    }

    const seen = new Set();
    sheet.tags = sheet.tags
      .map((tag) => normalizeTagName(tag))
      .filter((tag) => {
        if (!tag) return false;
        const key = tag.toLocaleLowerCase();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    return sheet.tags;
  }

  function addTagToSheet(sheet, rawTag) {
    const tags = ensureSheetTags(sheet);
    const next = normalizeTagName(rawTag);
    if (!next) return false;
    if (findTagIndex(tags, next) !== -1) return false;
    tags.push(next);
    return true;
  }

  function removeTagFromSheet(sheet, rawTag) {
    const tags = ensureSheetTags(sheet);
    const index = findTagIndex(tags, rawTag);
    if (index === -1) return false;
    tags.splice(index, 1);
    return true;
  }

  function parseTagOperation(rawInput) {
    const trimmed = String(rawInput == null ? "" : rawInput).trim();
    if (!trimmed) return null;
    const marker = trimmed[0];
    const mode = marker === "-" ? "remove" : "add";
    const value = marker === "+" || marker === "-" ? trimmed.slice(1) : trimmed;
    const tag = normalizeTagName(value);
    if (!tag) return null;
    return { mode, tag };
  }

  function applyTagOperation(sheet, rawInput) {
    if (!sheet || sheet.kind !== "sheet") return false;
    const operation = parseTagOperation(rawInput);
    if (!operation) return false;
    const changed =
      operation.mode === "remove" ? removeTagFromSheet(sheet, operation.tag) : addTagToSheet(sheet, operation.tag);
    if (!changed) return false;
    sheet.updatedAt = nowISO();
    saveState();
    return true;
  }

  function createId(prefix) {
    return `${prefix}_${Math.random().toString(36).slice(2, 10)}`;
  }

  function createStack({ title, parentId = null, previewCount = null }) {
    const time = nowISO();
    return {
      id: createId("stack"),
      kind: "stack",
      title: safeName(title, "Untitled Stack"),
      parentId,
      previewCount: Number.isFinite(previewCount) ? previewCount : null,
      children: [],
      createdAt: time,
      updatedAt: time
    };
  }

  function createSheet({ title, body = "", parentId = null, subtitle = "" }) {
    const time = nowISO();
    return {
      id: createId("sheet"),
      kind: "sheet",
      title: safeName(title),
      subtitle: String(subtitle || ""),
      body,
      parentId,
      tags: [],
      margins: { ...DEFAULT_SETTINGS.margins },
      createdAt: time,
      updatedAt: time
    };
  }

  function seedState() {
    const root = createStack({ title: "Notations", parentId: null });
    root.id = "root";

    const stackA = createStack({ title: "This is a stack", parentId: root.id, previewCount: 5 });
    const food = createStack({ title: "Food Notes", parentId: root.id, previewCount: 33 });
    const poems = createStack({ title: "Poems", parentId: root.id, previewCount: 25 });
    const noteA = createSheet({ title: "This is the title", body: SAMPLE_TEXT, parentId: root.id });
    const noteB = createSheet({ title: "This is a notation with a longer title", body: "", parentId: root.id });
    const noteDate = createSheet({ title: "12-01-14", body: "", parentId: root.id });
    const noteEtc = createSheet({ title: "ETC.", body: "", parentId: root.id });

    const foodA = createSheet({
      title: "Huckleberry Pie Recipe",
      body: SAMPLE_TEXT,
      parentId: food.id
    });
    const foodB = createSheet({
      title: "A Moveable Feast - Chapter 1",
      body: Array(80).fill("A line of sample text to validate screen and print parity.").join("\n"),
      parentId: food.id,
      subtitle: "Chapter 1 - Jan. 4 1920"
    });
    foodB.tags = ["first person", "biographical", "final draft", "chapter"];

    const stackAItem = createSheet({ title: "Draft", body: "Stack sample note.", parentId: stackA.id });
    const poemsA = createSheet({ title: "Poems", body: "Line one\nLine two\nLine three", parentId: poems.id });

    root.children = [stackA.id, food.id, noteA.id, noteB.id, poems.id, noteDate.id, noteEtc.id];
    stackA.children = [stackAItem.id];
    food.children = [foodA.id, foodB.id];
    poems.children = [poemsA.id];

    const containers = {
      [root.id]: root,
      [stackA.id]: stackA,
      [food.id]: food,
      [poems.id]: poems,
      [noteA.id]: noteA,
      [noteB.id]: noteB,
      [noteDate.id]: noteDate,
      [noteEtc.id]: noteEtc,
      [foodA.id]: foodA,
      [foodB.id]: foodB,
      [stackAItem.id]: stackAItem,
      [poemsA.id]: poemsA
    };

    return {
      auth: { loggedIn: false, username: "" },
      settings: JSON.parse(JSON.stringify(DEFAULT_SETTINGS)),
      rootId: root.id,
      currentStackId: root.id,
      activeSheetId: null,
      containers,
      ui: { selectedCardId: null, settingsOpen: false, tagsHidden: false, zenMode: false, typewriterMode: false }
    };
  }

  function loadStateFromLocalStorage() {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        return seedState();
      }
      const parsed = JSON.parse(raw);
      return normalizeState(parsed);
    } catch (_) {
      return seedState();
    }
  }

  async function loadState() {
    try {
      if (window.notations && typeof window.notations.loadState === "function") {
        const response = await window.notations.loadState();
        if (response && response.ok && response.state) {
          const normalized = normalizeState(response.state);
          window.localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized));
          return normalized;
        }
      }
    } catch (_) {
      // fall through to local storage
    }
    return loadStateFromLocalStorage();
  }

  function normalizeState(raw) {
    const seeded = seedState();
    const merged = {
      ...seeded,
      ...raw,
      auth: { ...seeded.auth, ...(raw.auth || {}) },
      settings: {
        ...seeded.settings,
        ...(raw.settings || {}),
        margins: { ...seeded.settings.margins, ...((raw.settings && raw.settings.margins) || {}) }
      },
      ui: { ...seeded.ui, ...(raw.ui || {}) }
    };

    if (!merged.containers || !merged.containers[merged.rootId]) {
      return seeded;
    }

    if (!raw || !raw.settings) {
      merged.settings.fontFamily = DEFAULT_SETTINGS.fontFamily;
      merged.settings.fontSize = DEFAULT_SETTINGS.fontSize;
      merged.settings.lineHeight = DEFAULT_SETTINGS.lineHeight;
      merged.settings.margins = { ...DEFAULT_SETTINGS.margins };
    } else {
      if (!raw.settings.fontFamily || raw.settings.fontFamily === "sans") {
        merged.settings.fontFamily = DEFAULT_SETTINGS.fontFamily;
      }
      if (!Number.isFinite(raw.settings.fontSize) || raw.settings.fontSize === 16) {
        merged.settings.fontSize = DEFAULT_SETTINGS.fontSize;
      }
      const m = raw.settings.margins || {};
      if (
        Number(m.top) === 0.75 &&
        Number(m.right) === 0.75 &&
        Number(m.bottom) === 0.75 &&
        Number(m.left) === 0.75
      ) {
        merged.settings.margins = { ...DEFAULT_SETTINGS.margins };
      }
    }
    merged.settings.fontSize = normalizeFontSize(merged.settings.fontSize);

    Object.keys(merged.containers).forEach((id) => {
      const node = merged.containers[id];
      if (node && node.kind === "sheet") {
        ensureSheetTags(node);
        ensureSheetMargins(node, merged.settings.margins || DEFAULT_SETTINGS.margins);
      }
    });

    return merged;
  }

  function enforceDemoGridLayout(nextState) {
    const root = nextState.containers[nextState.rootId];
    if (!root || root.kind !== "stack") return;

    function byId(id) {
      return nextState.containers[id] || null;
    }

    function childIdsOfKind(kind) {
      return (root.children || []).filter((id) => {
        const node = byId(id);
        return node && node.kind === kind;
      });
    }

    function findChildByTitle(kind, titles) {
      const wanted = new Set(Array.isArray(titles) ? titles : [titles]);
      const hit = childIdsOfKind(kind).find((id) => wanted.has(byId(id).title));
      return hit ? byId(hit) : null;
    }

    function findAnyByTitle(kind, titles) {
      const wanted = new Set(Array.isArray(titles) ? titles : [titles]);
      const ids = Object.keys(nextState.containers);
      for (let i = 0; i < ids.length; i += 1) {
        const node = byId(ids[i]);
        if (!node || node.kind !== kind) continue;
        if (wanted.has(node.title)) return node;
      }
      return null;
    }

    function attachToRoot(node) {
      if (!node) return;
      node.parentId = root.id;
      if (!Array.isArray(root.children)) {
        root.children = [];
      }
      if (!root.children.includes(node.id)) {
        root.children.push(node.id);
      }
    }

    function ensureStack(primaryTitle, opts = {}) {
      const titles = [primaryTitle].concat(opts.fallbackTitles || []);
      let node = findChildByTitle("stack", titles) || findAnyByTitle("stack", titles);
      if (!node) {
        node = createStack({ title: primaryTitle, parentId: root.id, previewCount: opts.previewCount || null });
        nextState.containers[node.id] = node;
      }
      node.title = primaryTitle;
      node.previewCount = Number.isFinite(opts.previewCount) ? opts.previewCount : node.previewCount;
      attachToRoot(node);
      return node;
    }

    function ensureSheet(primaryTitle, opts = {}) {
      const titles = [primaryTitle].concat(opts.fallbackTitles || []);
      let node = findChildByTitle("sheet", titles) || findAnyByTitle("sheet", titles);
      if (!node) {
        node = createSheet({ title: primaryTitle, body: "", parentId: root.id });
        nextState.containers[node.id] = node;
      }
      node.title = primaryTitle;
      node.parentId = root.id;
      if (typeof opts.body === "string") {
        node.body = opts.body;
      }
      attachToRoot(node);
      return node;
    }

    const stackPrimary = ensureStack("This is a stack", { fallbackTitles: ["Untitled Stack"], previewCount: 5 });
    const stackFood = ensureStack("Food Notes", { previewCount: 33 });
    const noteTitle = ensureSheet("This is the title", { body: SAMPLE_TEXT });
    const noteLong = ensureSheet("This is a notation with a longer title", {
      fallbackTitles: ["This is a given title that is longer"],
      body: ""
    });
    const stackPoems = ensureStack("Poems", { previewCount: 25 });
    const noteDate = ensureSheet("12-01-14", { fallbackTitles: ["12-13-14"], body: "" });
    const noteEtc = ensureSheet("ETC.", { body: "" });

    const canonical = [stackPrimary.id, stackFood.id, noteTitle.id, noteLong.id, stackPoems.id, noteDate.id, noteEtc.id];
    root.children = canonical;

    nextState.currentStackId = root.id;
    nextState.ui = nextState.ui || {};
    nextState.ui.selectedCardId = null;
  }

  function saveState() {
    const serialized = JSON.stringify(state);
    window.localStorage.setItem(STORAGE_KEY, serialized);
    if (window.notations && typeof window.notations.saveState === "function") {
      window.notations.saveState({ serialized }).catch(() => {});
    }
  }

  let state = seedState();

  const el = {
    loginView: document.getElementById("loginView"),
    libraryView: document.getElementById("libraryView"),
    editorView: document.getElementById("editorView"),
    settingsPanel: document.getElementById("settingsPanel"),
    loginEnter: document.getElementById("loginEnter"),
    usernameInput: document.getElementById("usernameInput"),
    passwordInput: document.getElementById("passwordInput"),
    libraryBreadcrumb: document.getElementById("libraryBreadcrumb"),
    editorBreadcrumbTrail: document.getElementById("editorBreadcrumbTrail"),
    libraryGrid: document.getElementById("libraryGrid"),
    sheetTitle: document.getElementById("sheetTitle"),
    sheetSubtitle: document.getElementById("sheetSubtitle"),
    sheetBody: document.getElementById("sheetBody"),
    editorMain: document.querySelector("#editorView .editor-main"),
    quickFontFamily: document.getElementById("quickFontFamily"),
    quickFontSize: document.getElementById("quickFontSize"),
    printBtn: document.getElementById("printBtn"),
    pdfBtn: document.getElementById("pdfBtn"),
    txtBtn: document.getElementById("txtBtn"),
    editorPages: document.getElementById("editorPages"),
    pageFrames: document.getElementById("pageFrames"),
    tagsPanel: document.getElementById("tagsPanel"),
    editorMeta: document.getElementById("editorMeta"),
    tagChips: document.getElementById("tagChips"),
    paperSizeSelect: document.getElementById("paperSizeSelect"),
    fontFamilySelect: document.getElementById("fontFamilySelect"),
    fontSizeInput: document.getElementById("fontSizeInput"),
    marginTopInput: document.getElementById("marginTopInput"),
    marginRightInput: document.getElementById("marginRightInput"),
    marginBottomInput: document.getElementById("marginBottomInput"),
    marginLeftInput: document.getElementById("marginLeftInput"),
    resetDefaults: document.getElementById("resetDefaults"),
    printDocument: document.getElementById("printDocument"),
    choiceDialog: document.getElementById("choiceDialog"),
    choiceDialogScrim: document.getElementById("choiceDialogScrim"),
    choiceDialogTitle: document.getElementById("choiceDialogTitle"),
    choiceDialogMessage: document.getElementById("choiceDialogMessage"),
    choiceDialogActions: document.getElementById("choiceDialogActions"),
    commandPalette: document.getElementById("commandPalette"),
    commandPaletteScrim: document.getElementById("commandPaletteScrim"),
    commandPaletteInput: document.getElementById("commandPaletteInput"),
    commandPaletteSubtitle: document.getElementById("commandPaletteSubtitle"),
    commandPaletteList: document.getElementById("commandPaletteList"),
    commandPaletteEmpty: document.getElementById("commandPaletteEmpty"),
    toastRoot: document.getElementById("toastRoot")
  };

  function populateFontSizeOptions() {
    [el.quickFontSize, el.fontSizeInput].forEach((selectEl) => {
      if (!selectEl) return;
      selectEl.innerHTML = "";
      FONT_SIZE_OPTIONS.forEach((size) => {
        const option = document.createElement("option");
        option.value = String(size);
        option.textContent = String(size);
        selectEl.appendChild(option);
      });
    });
  }

  const dragState = {
    active: false,
    side: null
  };

  const homeTriggers = Array.from(document.querySelectorAll(".home-trigger"));
  const settingsTriggers = Array.from(document.querySelectorAll(".settings-trigger"));
  let zenTargetTextColumnWidthPx = null;
  let hasInitializedSheetLayout = false;
  let suppressHashChange = false;
  let activeDialogResolve = null;
  let activeDialogCancelValue = null;
  let previousActiveElement = null;
  let toastTimer = null;
  let typewriterMirrorEl = null;
  let typewriterMirrorBeforeNode = null;
  let typewriterMirrorMarker = null;
  let typewriterMirrorStyleKey = "";
  let typewriterMirrorWidth = -1;
  let typewriterSyncRaf = null;
  let typewriterPendingForce = false;
  let typewriterValueRevision = 0;
  let lastTypewriterVisualLine = null;
  let lastTypewriterMeasureRevision = -1;
  let lastTypewriterMeasuredCaretIndex = -1;
  const commandPaletteState = {
    open: false,
    mode: "root",
    selectedIndex: 0,
    previousActiveElement: null,
    items: []
  };
  const tagComposerState = {
    open: null,
    close: null
  };

  function closeChoiceDialog(result) {
    if (typeof activeDialogResolve !== "function") return;
    const resolve = activeDialogResolve;
    activeDialogResolve = null;
    activeDialogCancelValue = null;

    el.choiceDialog.classList.add("hidden");
    el.choiceDialog.setAttribute("aria-hidden", "true");
    el.choiceDialogActions.innerHTML = "";
    document.body.classList.remove("dialog-open");

    const restoreTarget = previousActiveElement;
    previousActiveElement = null;
    if (restoreTarget && typeof restoreTarget.focus === "function") {
      restoreTarget.focus();
    }

    resolve(result);
  }

  function showChoiceDialog({ title, message = "", actions = [], cancelValue = null }) {
    if (!el.choiceDialog || !el.choiceDialogActions) {
      return Promise.resolve(cancelValue);
    }

    return new Promise((resolve) => {
      if (typeof activeDialogResolve === "function") {
        const pendingResolve = activeDialogResolve;
        activeDialogResolve = null;
        pendingResolve(activeDialogCancelValue);
      }

      activeDialogResolve = resolve;
      activeDialogCancelValue = cancelValue;
      previousActiveElement =
        document.activeElement && typeof document.activeElement.focus === "function" ? document.activeElement : null;

      const dialogTitle = String(title || "");
      const dialogMessage = String(message || "");

      el.choiceDialogTitle.textContent = dialogTitle;
      el.choiceDialogMessage.textContent = dialogMessage;
      el.choiceDialogMessage.classList.toggle("hidden", !dialogMessage);
      el.choiceDialogActions.innerHTML = "";

      let focusTarget = null;
      actions.forEach((action, index) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = `dialog-btn dialog-btn-${action.variant || "secondary"}`;
        button.textContent = String(action.label || "");
        button.addEventListener("click", () => {
          closeChoiceDialog(action.id);
        });
        el.choiceDialogActions.appendChild(button);
        if (!focusTarget && (action.autoFocus || index === 0)) {
          focusTarget = button;
        }
      });

      el.choiceDialog.classList.remove("hidden");
      el.choiceDialog.setAttribute("aria-hidden", "false");
      document.body.classList.add("dialog-open");
      window.requestAnimationFrame(() => {
        if (focusTarget) {
          focusTarget.focus();
        }
      });
    });
  }

  function isEditableTarget(target) {
    if (!target || !(target instanceof Element)) return false;
    if (target.closest("#commandPalette")) return true;
    const tag = target.tagName ? target.tagName.toLowerCase() : "";
    if (tag === "textarea" || tag === "select") return true;
    if (tag === "input") {
      const type = String(target.type || "").toLowerCase();
      return type !== "button" && type !== "checkbox" && type !== "radio" && type !== "submit";
    }
    return target.isContentEditable || !!target.closest("[contenteditable='true']");
  }

  function shouldOpenCommandPalette(event) {
    if (commandPaletteState.open) return false;
    if (typeof activeDialogResolve === "function") return false;
    if (event.defaultPrevented) return false;
    if (event.repeat) return false;
    if (!event.metaKey || event.ctrlKey || event.altKey) return false;
    if (event.code !== "Slash") return false;
    return true;
  }

  function buildStackPathLabel(stackId) {
    const stack = state.containers[stackId];
    if (!stack || stack.kind !== "stack") return "";
    if (stack.id === state.rootId) {
      return `${safeName(stack.title)} (Root)`;
    }
    return getStackTrail(stack.id)
      .map((entry) => safeName(entry.title))
      .join(" / ");
  }

  function getAllStacksSorted() {
    const stacks = Object.values(state.containers).filter((node) => node && node.kind === "stack");
    return stacks.sort((a, b) => {
      if (a.id === state.rootId) return -1;
      if (b.id === state.rootId) return 1;
      const labelA = buildStackPathLabel(a.id).toLocaleLowerCase();
      const labelB = buildStackPathLabel(b.id).toLocaleLowerCase();
      return labelA.localeCompare(labelB);
    });
  }

  function showToast(message, options = {}) {
    if (!el.toastRoot) return;
    const duration = Number.isFinite(options.duration) ? options.duration : 1800;
    const node = document.createElement("div");
    node.className = "toast";
    node.textContent = String(message || "");
    el.toastRoot.innerHTML = "";
    el.toastRoot.appendChild(node);
    requestAnimationFrame(() => node.classList.add("show"));
    if (toastTimer) {
      clearTimeout(toastTimer);
    }
    toastTimer = setTimeout(() => {
      node.classList.remove("show");
      setTimeout(() => {
        if (el.toastRoot.contains(node)) {
          el.toastRoot.removeChild(node);
        }
      }, 140);
    }, duration);
  }

  function basename(filePath) {
    return String(filePath || "")
      .split(/[\\/]/)
      .filter(Boolean)
      .pop();
  }

  async function runPrintCommand() {
    const sheet = getActiveSheet();
    if (!sheet) {
      showToast("Open a notation first.");
      return;
    }
    preparePrintDocument();
    const result = await window.notations.printDocument(buildPrintRequestPayload());
    if (result && result.ok) {
      showToast("Print dialog opened.");
      return;
    }
    if (result && result.cancelled) return;
    const reason = result && result.reason ? String(result.reason) : "";
    showToast(reason ? `Print failed: ${reason}` : "Print failed.");
  }

  async function runPdfExportCommand() {
    const sheet = getActiveSheet();
    if (!sheet) return;
    preparePrintDocument();
    const result = await window.notations.exportPdf({ defaultFilename: safeFilename(sheet.title) });
    if (result && result.ok) {
      const label = basename(result.filePath) || "PDF";
      showToast(`Exported ${label}`);
      return;
    }
    if (result && result.cancelled) return;
    showToast("PDF export failed.");
  }

  async function runTextExportCommand() {
    const sheet = getActiveSheet();
    if (!sheet) return;
    const result = await window.notations.exportText({
      defaultFilename: safeFilename(sheet.title),
      body: sheet.body
    });
    if (result && result.ok) {
      const label = basename(result.filePath) || "TXT";
      showToast(`Exported ${label}`);
      return;
    }
    if (result && result.cancelled) return;
    showToast("Text export failed.");
  }

  function syncSettingsControlsFromState() {
    const margins = getActiveSheetMargins();
    state.settings.fontSize = normalizeFontSize(state.settings.fontSize);
    el.quickFontFamily.value = state.settings.fontFamily;
    el.quickFontSize.value = String(state.settings.fontSize);
    el.paperSizeSelect.value = state.settings.paperSize;
    el.fontFamilySelect.value = state.settings.fontFamily;
    el.fontSizeInput.value = String(state.settings.fontSize);
    syncMarginInputBounds();
    el.marginTopInput.value = String(margins.top);
    el.marginRightInput.value = String(margins.right);
    el.marginBottomInput.value = String(margins.bottom);
    el.marginLeftInput.value = String(margins.left);
  }

  function resetActiveMarginsCommand() {
    const activeSheet = getActiveSheet();
    if (!activeSheet) {
      showToast("Open a notation first.");
      return;
    }
    activeSheet.margins = { ...DEFAULT_SETTINGS.margins };
    activeSheet.updatedAt = nowISO();
    syncSettingsControlsFromState();
    saveState();
    applyTypographyAndGeometry();
    updateMeta();
    showToast("Margins reset.");
  }

  function resetTypographyDefaultsCommand() {
    state.settings.fontFamily = DEFAULT_SETTINGS.fontFamily;
    state.settings.fontSize = DEFAULT_SETTINGS.fontSize;
    state.settings.margins = { ...DEFAULT_SETTINGS.margins };

    const activeSheet = getActiveSheet();
    if (activeSheet) {
      activeSheet.margins = { ...DEFAULT_SETTINGS.margins };
      activeSheet.updatedAt = nowISO();
    }

    syncSettingsControlsFromState();
    saveState();
    applyTypographyAndGeometry();
    if (activeSheet) {
      updateMeta();
    }
    showToast("Defaults reset (font, size, margins).");
  }

  function setFontFamilyCommand(family, label) {
    if (!FONT_PRESETS[family]) return;
    state.settings.fontFamily = family;
    syncSettingsControlsFromState();
    saveState();
    applyTypographyAndGeometry();
    showToast(`Font set to ${label}.`);
  }

  function adjustFontSizeCommand(delta) {
    const nextSize = normalizeFontSize(Number(state.settings.fontSize || DEFAULT_SETTINGS.fontSize) + Number(delta || 0));
    if (nextSize === state.settings.fontSize) {
      showToast(`Font size is already ${nextSize}.`);
      return;
    }
    state.settings.fontSize = nextSize;
    syncSettingsControlsFromState();
    saveState();
    applyTypographyAndGeometry();
    showToast(`Font size: ${nextSize}`);
  }

  function measureCurrentTextColumnWidthPx() {
    const rect = el.sheetBody.getBoundingClientRect();
    if (!rect.width) return null;
    const sheetStyle = getComputedStyle(el.sheetBody);
    const paddingLeft = Number.parseFloat(sheetStyle.paddingLeft || "0") || 0;
    const paddingRight = Number.parseFloat(sheetStyle.paddingRight || "0") || 0;
    const contentWidth = rect.width - paddingLeft - paddingRight;
    return contentWidth > 0 ? contentWidth : null;
  }

  function toggleOptionsCommand() {
    state.ui.settingsOpen = !state.ui.settingsOpen;
    saveState();
    renderSettingsVisibility();
    if (state.activeSheetId) {
      requestAnimationFrame(() => applyTypographyAndGeometry());
    }
    showToast(state.ui.settingsOpen ? "Options shown." : "Options hidden.");
  }

  function setZenMode(enabled, options = {}) {
    const { showFeedback = true } = options;
    const activeSheet = getActiveSheet();
    if (!activeSheet && enabled) {
      showToast("Open a notation first.");
      return false;
    }
    if (enabled && !state.ui.zenMode && activeSheet) {
      const currentWidth = measureCurrentTextColumnWidthPx();
      if (currentWidth) {
        zenTargetTextColumnWidthPx = currentWidth;
      }
    }
    if (!enabled) {
      zenTargetTextColumnWidthPx = null;
    }
    state.ui.zenMode = !!enabled;
    if (state.ui.zenMode) {
      state.ui.settingsOpen = false;
    }
    saveState();
    renderSettingsVisibility();
    renderZenMode();
    if (activeSheet) {
      applyTypographyAndGeometry();
    }
    if (showFeedback) {
      showToast(state.ui.zenMode ? "Zen Mode on." : "Zen Mode off.");
    }
    return true;
  }

  function toggleZenModeCommand() {
    setZenMode(!state.ui.zenMode);
  }

  function resetTypewriterTracking() {
    lastTypewriterVisualLine = null;
    lastTypewriterMeasureRevision = -1;
    lastTypewriterMeasuredCaretIndex = -1;
  }

  function setTypewriterMode(enabled, options = {}) {
    const { showFeedback = true } = options;
    const activeSheet = getActiveSheet();
    if (!activeSheet && enabled) {
      showToast("Open a notation first.");
      return false;
    }
    state.ui.typewriterMode = !!enabled;
    if (!state.ui.typewriterMode) {
      resetTypewriterTracking();
    }
    saveState();
    scheduleTypewriterSync({ force: true });
    if (showFeedback) {
      showToast(state.ui.typewriterMode ? "Typewriter Mode on." : "Typewriter Mode off.");
    }
    return true;
  }

  function toggleTypewriterModeCommand() {
    setTypewriterMode(!state.ui.typewriterMode);
  }

  function toggleTagsVisibilityCommand() {
    state.ui.tagsHidden = !state.ui.tagsHidden;
    saveState();
    renderTagsVisibility();
    showToast(state.ui.tagsHidden ? "Tags hidden." : "Tags shown.");
  }

  function addTagToActiveSheetCommand(rawTag) {
    const sheet = getActiveSheet();
    if (!sheet) {
      showToast("Open a notation first.");
      return false;
    }
    const tag = normalizeTagName(rawTag);
    if (!tag) {
      showToast("Enter a tag name.");
      return false;
    }
    if (!addTagToSheet(sheet, tag)) {
      showToast(`Tag "${tag}" already exists.`);
      return false;
    }
    sheet.updatedAt = nowISO();
    saveState();
    renderTags();
    updateMeta();
    showToast(`Added tag "${tag}".`);
    return true;
  }

  function removeTagFromActiveSheetCommand(tag) {
    const sheet = getActiveSheet();
    if (!sheet) {
      showToast("Open a notation first.");
      return false;
    }
    if (!removeTagFromSheet(sheet, tag)) {
      showToast(`Tag "${tag}" not found.`);
      return false;
    }
    sheet.updatedAt = nowISO();
    saveState();
    renderTags();
    updateMeta();
    showToast(`Removed tag "${tag}".`);
    return true;
  }

  function renameActiveSheetCommand(rawTitle) {
    const sheet = getActiveSheet();
    if (!sheet) {
      showToast("Open a notation first.");
      return false;
    }
    const nextTitle = safeName(String(rawTitle || "").trim());
    if (sheet.title === nextTitle) {
      showToast("Name is unchanged.");
      return true;
    }
    sheet.title = nextTitle;
    sheet.updatedAt = nowISO();
    saveState();
    syncHashToState();
    renderEditor();
    showToast(`Renamed to "${nextTitle}".`);
    return true;
  }

  function moveActiveSheetToStack(targetStackId) {
    const sheet = getActiveSheet();
    const target = state.containers[targetStackId];
    if (!sheet || sheet.kind !== "sheet" || !target || target.kind !== "stack") return;

    const currentParentId = sheet.parentId || state.rootId;
    if (currentParentId === target.id) {
      showToast(`"${safeName(sheet.title)}" is already in ${safeName(target.title)}.`);
      return;
    }

    const source = state.containers[currentParentId];
    const now = nowISO();

    if (source && source.kind === "stack" && Array.isArray(source.children)) {
      source.children = source.children.filter((id) => id !== sheet.id);
      source.updatedAt = now;
    }

    if (!Array.isArray(target.children)) {
      target.children = [];
    }
    target.children = target.children.filter((id) => id !== sheet.id);
    target.children.unshift(sheet.id);
    target.updatedAt = now;

    sheet.parentId = target.id;
    sheet.updatedAt = now;
    state.currentStackId = target.id;
    state.ui.selectedCardId = sheet.id;
    saveState();
    navigateToSheet(sheet.id, { syncHash: true, persist: false });
    showToast(`Moved "${safeName(sheet.title)}" to ${safeName(target.title)}.`);
  }

  function getCommandItemsForMode(mode, query = "") {
    const activeSheet = getActiveSheet();
    const hasSheet = !!activeSheet;
    const canUseAppCommands = !!state.auth.loggedIn;
    const queryText = String(query || "").trim();

    if (mode === "move-stack") {
      if (!hasSheet) return [];
      const currentParentId = activeSheet.parentId || state.rootId;
      return getAllStacksSorted().map((stack) => ({
        id: `move:${stack.id}`,
        label: buildStackPathLabel(stack.id),
        meta: stack.id === currentParentId ? "Current stack" : "Move here",
        keywords: [stack.title, stack.id],
        disabled: false,
        closeOnRun: true,
        action: () => moveActiveSheetToStack(stack.id)
      }));
    }

    if (mode === "export-as") {
      if (!hasSheet) return [];
      return [
        {
          id: "export:print",
          label: "Print...",
          meta: "System print dialog",
          closeOnRun: true,
          action: runPrintCommand
        },
        {
          id: "export:pdf",
          label: "PDF file (.pdf)",
          meta: "Export current notation",
          closeOnRun: true,
          action: runPdfExportCommand
        },
        {
          id: "export:txt",
          label: "Text file (.txt)",
          meta: "Export plain text",
          closeOnRun: true,
          action: runTextExportCommand
        }
      ];
    }

    if (mode === "add-tag") {
      if (!hasSheet) return [];
      if (!queryText) return [];
      return [
        {
          id: `tag:add:${queryText.toLocaleLowerCase()}`,
          label: `Add tag "${queryText}"`,
          meta: "Press Enter to apply",
          keywords: [queryText, "tag", "add"],
          disabled: false,
          closeOnRun: false,
          action: () => {
            const added = addTagToActiveSheetCommand(queryText);
            if (added) {
              closeCommandPalette({ restoreFocus: false });
            }
            return added;
          }
        }
      ];
    }

    if (mode === "remove-tag") {
      if (!hasSheet) return [];
      const tags = ensureSheetTags(activeSheet);
      return tags.map((tag) => ({
        id: `tag:remove:${tag.toLocaleLowerCase()}`,
        label: tag,
        meta: "Remove tag",
        keywords: [tag, "tag", "remove", "delete"],
        disabled: false,
        closeOnRun: true,
        action: () => removeTagFromActiveSheetCommand(tag)
      }));
    }

    if (mode === "rename-note") {
      if (!hasSheet) return [];
      if (!queryText) return [];
      return [
        {
          id: "note:rename",
          label: `Rename to "${queryText}"`,
          meta: "Press Enter to rename",
          keywords: [queryText, "rename", "title"],
          disabled: false,
          closeOnRun: false,
          action: () => {
            const renamed = renameActiveSheetCommand(queryText);
            if (renamed) {
              closeCommandPalette({ restoreFocus: false });
            }
            return renamed;
          }
        }
      ];
    }

    return [
      {
        id: "root:move",
        label: "Move to stack...",
        meta: hasSheet ? "Pick destination stack" : "Open a notation first",
        keywords: ["move", "stack"],
        disabled: !canUseAppCommands || !hasSheet,
        closeOnRun: false,
        action: () => setCommandPaletteMode("move-stack")
      },
      {
        id: "root:export",
        label: "Export as...",
        meta: hasSheet ? "Print / PDF / TXT" : "Open a notation first",
        keywords: ["export", "pdf", "txt", "print"],
        disabled: !canUseAppCommands || !hasSheet,
        closeOnRun: false,
        action: () => setCommandPaletteMode("export-as")
      },
      {
        id: "root:reset-margins",
        label: "Reset Margins",
        meta: hasSheet ? "Reset current notation margins" : "Open a notation first",
        keywords: ["reset", "margins", "margin"],
        disabled: !canUseAppCommands || !hasSheet,
        closeOnRun: true,
        action: resetActiveMarginsCommand
      },
      {
        id: "root:reset-defaults",
        label: "Reset all Defaults",
        meta: canUseAppCommands ? "Reset font, font-size and margins" : "Log in first",
        keywords: ["reset", "defaults", "font", "size", "margins"],
        disabled: !canUseAppCommands,
        closeOnRun: true,
        action: resetTypographyDefaultsCommand
      },
      {
        id: "root:font-sans",
        label: "Font: Sans Serif",
        meta: canUseAppCommands ? "Switch editor font family" : "Log in first",
        keywords: ["font", "sans", "family", "serif"],
        disabled: !canUseAppCommands,
        closeOnRun: true,
        action: () => setFontFamilyCommand("sans", "Sans Serif")
      },
      {
        id: "root:font-serif",
        label: "Font: Serif",
        meta: canUseAppCommands ? "Switch editor font family" : "Log in first",
        keywords: ["font", "serif", "family"],
        disabled: !canUseAppCommands,
        closeOnRun: true,
        action: () => setFontFamilyCommand("serif", "Serif")
      },
      {
        id: "root:font-monospace",
        label: "Font: Monospace",
        meta: canUseAppCommands ? "Switch editor font family" : "Log in first",
        keywords: ["font", "mono", "monospace", "family"],
        disabled: !canUseAppCommands,
        closeOnRun: true,
        action: () => setFontFamilyCommand("monospace", "Monospace")
      },
      {
        id: "root:font-size-bigger",
        label: "Font Size: Bigger +4",
        meta: canUseAppCommands ? "Increase font size by 4" : "Log in first",
        keywords: ["font", "size", "bigger", "increase", "+4"],
        disabled: !canUseAppCommands,
        closeOnRun: true,
        action: () => adjustFontSizeCommand(4)
      },
      {
        id: "root:font-size-smaller",
        label: "Font Size: Smaller -4",
        meta: canUseAppCommands ? "Decrease font size by 4" : "Log in first",
        keywords: ["font", "size", "smaller", "decrease", "-4"],
        disabled: !canUseAppCommands,
        closeOnRun: true,
        action: () => adjustFontSizeCommand(-4)
      },
      {
        id: "root:toggle-options",
        label: "Show/Hide Options",
        meta: canUseAppCommands
          ? state.ui.settingsOpen
            ? "Options are visible"
            : "Options are hidden"
          : "Log in first",
        keywords: ["show", "hide", "options", "settings", "toggle"],
        disabled: !canUseAppCommands,
        closeOnRun: true,
        action: toggleOptionsCommand
      },
      {
        id: "root:zen-mode",
        label: "Zen Mode",
        meta: canUseAppCommands ? (state.ui.zenMode ? "Enabled" : "Disabled") : "Log in first",
        keywords: ["zen", "focus", "hide", "ui"],
        disabled: !canUseAppCommands || !hasSheet,
        closeOnRun: true,
        action: toggleZenModeCommand
      },
      {
        id: "root:typewriter-mode",
        label: "Typewriter Mode",
        meta: canUseAppCommands ? (state.ui.typewriterMode ? "Enabled" : "Disabled") : "Log in first",
        keywords: ["typewriter", "focus", "center", "wrap", "line"],
        disabled: !canUseAppCommands || !hasSheet,
        closeOnRun: true,
        action: toggleTypewriterModeCommand
      },
      {
        id: "root:toggle-tags",
        label: "Show/Hide Tags",
        meta: canUseAppCommands ? (state.ui.tagsHidden ? "Tags are hidden" : "Tags are visible") : "Log in first",
        keywords: ["show", "hide", "tags", "toggle"],
        disabled: !canUseAppCommands,
        closeOnRun: true,
        action: toggleTagsVisibilityCommand
      },
      {
        id: "root:add-tag",
        label: "Add Tag...",
        meta: hasSheet ? "Type tag name in the command bar" : "Open a notation first",
        keywords: ["add", "tag", "label"],
        disabled: !canUseAppCommands || !hasSheet,
        closeOnRun: false,
        action: () => setCommandPaletteMode("add-tag")
      },
      {
        id: "root:remove-tag",
        label: "Remove Tag...",
        meta: hasSheet ? "Select from applied tags" : "Open a notation first",
        keywords: ["remove", "tag", "delete"],
        disabled: !canUseAppCommands || !hasSheet,
        closeOnRun: false,
        action: () => setCommandPaletteMode("remove-tag")
      },
      {
        id: "root:rename",
        label: "Rename",
        meta: hasSheet ? "Rename current notation" : "Open a notation first",
        keywords: ["rename", "title", "note"],
        disabled: !canUseAppCommands || !hasSheet,
        closeOnRun: false,
        action: () =>
          setCommandPaletteMode("rename-note", { value: safeName(activeSheet ? activeSheet.title : ""), select: true })
      }
    ];
  }

  function getCommandSubtitle(mode) {
    if (mode === "move-stack") {
      const sheet = getActiveSheet();
      const title = sheet ? safeName(sheet.title) : "notation";
      return `Move "${title}" to stack. UP/DOWN to pick, ENTER to select, ESC to go back.`;
    }
    if (mode === "export-as") {
      return "Choose export target. UP/DOWN to pick, ENTER to select, ESC to go back.";
    }
    if (mode === "add-tag") {
      return "Type the tag name, then press ENTER to add it.";
    }
    if (mode === "remove-tag") {
      return "Pick an applied tag and press ENTER to remove it.";
    }
    if (mode === "rename-note") {
      return "Type the new note name, then press ENTER to rename it.";
    }
    return "Type to filter commands. UP/DOWN to pick, ENTER to run.";
  }

  function renderCommandPalette() {
    if (!commandPaletteState.open) return;
    const mode = commandPaletteState.mode;
    const query = String(el.commandPaletteInput.value || "")
      .trim()
      .toLocaleLowerCase();
    const allItems = getCommandItemsForMode(mode, query);
    const items = allItems.filter((item) => {
      if (!query) return true;
      const haystack = [item.label, item.meta]
        .concat(Array.isArray(item.keywords) ? item.keywords : [])
        .filter(Boolean)
        .join(" ")
        .toLocaleLowerCase();
      return haystack.includes(query);
    });

    commandPaletteState.items = items;
    if (!items.length) {
      commandPaletteState.selectedIndex = 0;
    } else {
      commandPaletteState.selectedIndex = clamp(commandPaletteState.selectedIndex, 0, items.length - 1);
    }

    el.commandPaletteSubtitle.textContent = getCommandSubtitle(mode);
    el.commandPaletteList.innerHTML = "";
    el.commandPaletteEmpty.classList.toggle("hidden", !!items.length);
    if (mode === "move-stack") {
      el.commandPaletteEmpty.textContent = "No matching stacks";
    } else if (mode === "export-as") {
      el.commandPaletteEmpty.textContent = "No matching export targets";
    } else if (mode === "remove-tag") {
      el.commandPaletteEmpty.textContent = query ? "No matching tags" : "No applied tags";
    } else if (mode === "add-tag") {
      el.commandPaletteEmpty.textContent = "Type a tag name";
    } else if (mode === "rename-note") {
      el.commandPaletteEmpty.textContent = "Type a new note name";
    } else {
      el.commandPaletteEmpty.textContent = "No matching commands";
    }

    items.forEach((item, index) => {
      const listItem = document.createElement("li");
      const button = document.createElement("button");
      button.type = "button";
      button.className = "command-palette-item";
      if (index === commandPaletteState.selectedIndex) {
        button.classList.add("active");
      }
      if (item.disabled) {
        button.disabled = true;
      }
      button.setAttribute("role", "option");
      button.setAttribute("aria-selected", index === commandPaletteState.selectedIndex ? "true" : "false");
      const main = document.createElement("span");
      main.className = "command-palette-item-main";
      main.textContent = String(item.label || "");
      button.appendChild(main);
      const meta = document.createElement("span");
      meta.className = "command-palette-item-meta";
      meta.textContent = String(item.meta || "");
      button.appendChild(meta);
      button.addEventListener("mouseenter", () => {
        commandPaletteState.selectedIndex = index;
        renderCommandPalette();
      });
      button.addEventListener("click", () => {
        runCommandItem(item);
      });
      listItem.appendChild(button);
      el.commandPaletteList.appendChild(listItem);
    });
  }

  function setCommandPaletteMode(mode, options = {}) {
    const nextValue = typeof options.value === "string" ? options.value : "";
    const select = !!options.select;
    commandPaletteState.mode = mode;
    commandPaletteState.selectedIndex = 0;
    el.commandPaletteInput.value = nextValue;
    renderCommandPalette();
    requestAnimationFrame(() => {
      el.commandPaletteInput.focus();
      if (select) {
        el.commandPaletteInput.select();
      }
    });
  }

  function openCommandPalette() {
    if (commandPaletteState.open) return;
    commandPaletteState.open = true;
    commandPaletteState.mode = "root";
    commandPaletteState.selectedIndex = 0;
    commandPaletteState.previousActiveElement =
      document.activeElement && typeof document.activeElement.focus === "function" ? document.activeElement : null;
    el.commandPaletteInput.value = "";
    el.commandPalette.classList.remove("hidden");
    el.commandPalette.setAttribute("aria-hidden", "false");
    renderCommandPalette();
    requestAnimationFrame(() => el.commandPaletteInput.focus());
  }

  function openAddTagCommand() {
    if (!state.auth.loggedIn) {
      showToast("Log in first.");
      return;
    }
    if (!getActiveSheet()) {
      showToast("Open a notation first.");
      return;
    }
    if (state.ui.zenMode) {
      showToast("Exit Zen Mode to add tags.");
      return;
    }

    let shouldSave = false;
    if (!state.ui.settingsOpen) {
      state.ui.settingsOpen = true;
      shouldSave = true;
    }
    if (state.ui.tagsHidden) {
      state.ui.tagsHidden = false;
      shouldSave = true;
    }

    if (commandPaletteState.open) {
      closeCommandPalette({ restoreFocus: false });
    }

    if (shouldSave) {
      saveState();
      renderSettingsVisibility();
      renderTags();
      requestAnimationFrame(() => applyTypographyAndGeometry());
    }

    requestAnimationFrame(() => {
      if (typeof tagComposerState.open === "function") {
        tagComposerState.open();
        return;
      }
      renderTags();
      if (typeof tagComposerState.open === "function") {
        tagComposerState.open();
      }
    });
  }

  function closeCommandPalette(options = {}) {
    if (!commandPaletteState.open) return;
    const { restoreFocus = true } = options;
    commandPaletteState.open = false;
    commandPaletteState.mode = "root";
    commandPaletteState.selectedIndex = 0;
    commandPaletteState.items = [];
    el.commandPalette.classList.add("hidden");
    el.commandPalette.setAttribute("aria-hidden", "true");
    el.commandPaletteList.innerHTML = "";
    if (restoreFocus && commandPaletteState.previousActiveElement) {
      commandPaletteState.previousActiveElement.focus();
    }
    commandPaletteState.previousActiveElement = null;
  }

  function moveCommandSelection(step) {
    const items = commandPaletteState.items;
    if (!items.length) return;
    const maxIndex = items.length - 1;
    const next = commandPaletteState.selectedIndex + step;
    if (next > maxIndex) {
      commandPaletteState.selectedIndex = 0;
    } else if (next < 0) {
      commandPaletteState.selectedIndex = maxIndex;
    } else {
      commandPaletteState.selectedIndex = next;
    }
    renderCommandPalette();
  }

  function getSelectedCommandItem() {
    const items = commandPaletteState.items;
    if (!items.length) return null;
    return items[commandPaletteState.selectedIndex] || null;
  }

  async function runCommandItem(item) {
    if (!item || item.disabled || typeof item.action !== "function") return;
    const shouldClose = item.closeOnRun !== false;
    if (shouldClose) {
      closeCommandPalette({ restoreFocus: false });
    }

    try {
      await item.action();
    } catch (_) {
      showToast("Command failed.");
    }

    if (shouldClose) {
      commandPaletteState.previousActiveElement = null;
    } else if (commandPaletteState.open) {
      renderCommandPalette();
    }
  }

  function normalizeRouteToken(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "");
  }

  function titleToRouteSegment(title) {
    const cleaned = String(title || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^\w\s-]/g, "")
      .trim()
      .replace(/\s+/g, "-")
      .replace(/-+/g, "-");
    return cleaned || "untitled";
  }

  function routeMatchesSegment(node, segment) {
    const target = normalizeRouteToken(segment);
    if (!target || !node) return false;
    const byId = normalizeRouteToken(node.id);
    const byTitle = normalizeRouteToken(node.title);
    const bySlug = normalizeRouteToken(titleToRouteSegment(node.title));
    if (byId === target || byTitle === target || bySlug === target) return true;
    return byTitle.startsWith(`${target}-`) || bySlug.startsWith(`${target}-`);
  }

  function findChildBySegment(parentStack, kind, segment) {
    if (!parentStack || parentStack.kind !== "stack" || !Array.isArray(parentStack.children)) {
      return null;
    }
    for (let i = 0; i < parentStack.children.length; i += 1) {
      const node = state.containers[parentStack.children[i]];
      if (!node || node.kind !== kind) continue;
      if (routeMatchesSegment(node, segment)) {
        return node;
      }
    }
    return null;
  }

  function getStackRouteSegments(stackId) {
    const segments = [];
    let cursor = state.containers[stackId];
    while (cursor && cursor.id !== state.rootId) {
      segments.push(titleToRouteSegment(cursor.title));
      cursor = cursor.parentId ? state.containers[cursor.parentId] : null;
      if (cursor && cursor.kind !== "stack") {
        break;
      }
    }
    return segments.reverse();
  }

  function getSheetRouteSegments(sheetId) {
    const sheet = state.containers[sheetId];
    if (!sheet || sheet.kind !== "sheet") return [];
    const parentSegments = getStackRouteSegments(sheet.parentId || state.rootId);
    parentSegments.push(titleToRouteSegment(sheet.title));
    return parentSegments;
  }

  function getStackTrail(stackId) {
    const trail = [];
    let cursor = state.containers[stackId];
    while (cursor && cursor.id !== state.rootId) {
      if (cursor.kind === "stack") {
        trail.push(cursor);
      }
      cursor = cursor.parentId ? state.containers[cursor.parentId] : null;
    }
    return trail.reverse();
  }

  function renderStackBreadcrumb(container, trail, options = {}) {
    if (!container) return;
    const { clickable = false } = options;
    container.innerHTML = "";
    if (!Array.isArray(trail) || !trail.length) {
      return;
    }

    const frag = document.createDocumentFragment();
    trail.forEach((stackNode) => {
      const segment = document.createElement(clickable ? "button" : "span");
      segment.className = "breadcrumb-segment";
      segment.textContent = safeName(stackNode.title);
      if (clickable) {
        segment.type = "button";
        segment.classList.add("breadcrumb-link");
        segment.addEventListener("click", () => {
          navigateToLibrary(stackNode.id, { syncHash: true, persist: true });
        });
      }
      frag.appendChild(segment);
    });
    container.appendChild(frag);
  }

  function readRouteSegmentsFromPath(pathValue) {
    const raw = String(pathValue || "").trim();
    if (!raw || raw === "/") return [];
    return raw
      .split("/")
      .filter(Boolean)
      .map((segment) => {
        try {
          return decodeURIComponent(segment);
        } catch (_) {
          return segment;
        }
      });
  }

  function readRouteSegmentsFromLocation() {
    return readRouteSegmentsFromPath(String(window.location.hash || "").replace(/^#/, ""));
  }

  function buildHashFromSegments(segments) {
    if (!segments || !segments.length) return "#/";
    return `#/${segments.map((s) => encodeURIComponent(s)).join("/")}`;
  }

  function getCurrentRouteHash() {
    if (!state.auth.loggedIn) return buildHashFromSegments(["login"]);
    if (state.activeSheetId) {
      return buildHashFromSegments(getSheetRouteSegments(state.activeSheetId));
    }
    const current = getCurrentStack();
    if (!current || current.id === state.rootId) {
      return "#/";
    }
    return buildHashFromSegments(getStackRouteSegments(current.id));
  }

  function syncHashToState() {
    const nextHash = getCurrentRouteHash();
    if (window.location.hash === nextHash) return;
    suppressHashChange = true;
    window.location.hash = nextHash;
  }

  function resolveRoute(segments) {
    if (!segments.length) {
      return { kind: "library", stackId: state.rootId };
    }

    const first = normalizeRouteToken(segments[0]);
    if (first === "login") {
      return { kind: "login" };
    }

    const path = first === "library" ? segments.slice(1) : segments.slice();
    if (!path.length) {
      return { kind: "library", stackId: state.rootId };
    }

    let currentStack = state.containers[state.rootId];
    if (!currentStack || currentStack.kind !== "stack") {
      return null;
    }

    for (let i = 0; i < path.length; i += 1) {
      const segment = path[i];
      const isLast = i === path.length - 1;
      const stackMatch = findChildBySegment(currentStack, "stack", segment);

      if (!isLast) {
        if (!stackMatch) return null;
        currentStack = stackMatch;
        continue;
      }

      if (stackMatch) {
        return { kind: "library", stackId: stackMatch.id };
      }

      const sheetMatch = findChildBySegment(currentStack, "sheet", segment);
      if (sheetMatch) {
        return { kind: "editor", sheetId: sheetMatch.id };
      }

      return null;
    }

    return { kind: "library", stackId: state.rootId };
  }

  function getPaper() {
    return PAPER_PRESETS[state.settings.paperSize] || PAPER_PRESETS.letter;
  }

  function getEditorLayoutWidthPx() {
    const pagesWidth = el.editorPages ? Math.round(el.editorPages.getBoundingClientRect().width || 0) : 0;
    if (pagesWidth > 0) return pagesWidth;

    if (el.editorMain) {
      const mainStyle = getComputedStyle(el.editorMain);
      const paddingLeft = Number.parseFloat(mainStyle.paddingLeft || "0") || 0;
      const paddingRight = Number.parseFloat(mainStyle.paddingRight || "0") || 0;
      const contentWidth = Math.round((el.editorMain.clientWidth || 0) - paddingLeft - paddingRight);
      if (contentWidth > 0) return contentWidth;
    }

    return Math.max(320, Math.round(window.innerWidth - 80));
  }

  function syncMarginInputBounds(bounds = {}) {
    const paper = getPaper();
    const widthIn = Number.isFinite(bounds.widthIn) ? bounds.widthIn : getEditorLayoutWidthPx() / PAGE_DPI;
    const heightIn = Number.isFinite(bounds.heightIn) ? bounds.heightIn : paper.heightIn;
    const widthMax = String(Number(widthIn.toFixed(2)));
    const heightMax = String(Number(heightIn.toFixed(2)));
    el.marginTopInput.max = heightMax;
    el.marginRightInput.max = widthMax;
    el.marginBottomInput.max = heightMax;
    el.marginLeftInput.max = widthMax;
  }

  function getActiveSheet() {
    if (!state.activeSheetId) return null;
    return state.containers[state.activeSheetId] || null;
  }

  function getActiveSheetMargins() {
    const sheet = getActiveSheet();
    if (sheet) {
      return ensureSheetMargins(sheet, state.settings.margins || DEFAULT_SETTINGS.margins);
    }
    return normalizeMargins(state.settings.margins, DEFAULT_SETTINGS.margins);
  }

  function getCurrentStack() {
    return state.containers[state.currentStackId] || state.containers[state.rootId];
  }

  function computePageMetrics() {
    const paper = getPaper();
    const pageWidthPx = getEditorLayoutWidthPx();
    const pageHeightPx = Math.round(paper.heightIn * PAGE_DPI);
    const pageWidthIn = pageWidthPx / PAGE_DPI;
    const pageHeightIn = paper.heightIn;
    const m = getActiveSheetMargins();
    const marginIn = {
      top: clamp(Number(m.top || 0), 0, pageHeightIn),
      right: clamp(Number(m.right || 0), 0, pageWidthIn),
      bottom: clamp(Number(m.bottom || 0), 0, pageHeightIn),
      left: clamp(Number(m.left || 0), 0, pageWidthIn)
    };
    const contentWidthIn = clamp(pageWidthIn - (marginIn.left + marginIn.right), 0, pageWidthIn);
    const contentHeightIn = clamp(pageHeightIn - (marginIn.top + marginIn.bottom), 0, pageHeightIn);
    const contentWidthPx = Math.round(contentWidthIn * PAGE_DPI);
    const contentHeightPx = Math.round(contentHeightIn * PAGE_DPI);

    return {
      paper,
      pageWidthPx,
      pageHeightPx,
      pageWidthIn,
      pageHeightIn,
      marginIn,
      contentWidthPx,
      contentHeightPx,
      marginPx: {
        top: Math.round(marginIn.top * PAGE_DPI),
        right: Math.round(marginIn.right * PAGE_DPI),
        bottom: Math.round(marginIn.bottom * PAGE_DPI),
        left: Math.round(marginIn.left * PAGE_DPI)
      }
    };
  }

  function alignZenTextColumn(metrics) {
    if (!el.editorView.classList.contains("editor-zen")) return;
    const textareaRect = el.sheetBody.getBoundingClientRect();
    if (!textareaRect.width) return;

    const viewportWidth = Math.max(window.innerWidth || 0, document.documentElement.clientWidth || 0, textareaRect.width);
    const targetColumnWidth = clamp(
      Number(zenTargetTextColumnWidthPx || metrics.contentWidthPx),
      0,
      Math.max(0, textareaRect.width)
    );
    const desiredColumnLeft = (viewportWidth - targetColumnWidth) / 2;
    const totalHorizontalPadding = Math.max(0, textareaRect.width - targetColumnWidth);
    const rawLeftPadding = desiredColumnLeft - textareaRect.left;
    const nextLeftPadding = clamp(rawLeftPadding, 0, totalHorizontalPadding);
    const nextRightPadding = totalHorizontalPadding - nextLeftPadding;

    el.sheetBody.style.paddingLeft = `${nextLeftPadding.toFixed(2)}px`;
    el.sheetBody.style.paddingRight = `${nextRightPadding.toFixed(2)}px`;
  }

  function applyTypographyAndGeometry() {
    const metrics = computePageMetrics();
    syncMarginInputBounds({ widthIn: metrics.pageWidthIn, heightIn: metrics.pageHeightIn });
    const rootStyle = document.documentElement.style;
    rootStyle.setProperty("--paper-width-px", `${metrics.pageWidthPx}px`);
    rootStyle.setProperty("--paper-height-px", `${metrics.pageHeightPx}px`);
    rootStyle.setProperty("--content-width-px", `${metrics.contentWidthPx}px`);
    rootStyle.setProperty("--content-height-px", `${metrics.contentHeightPx}px`);
    rootStyle.setProperty("--editor-font-size", `${state.settings.fontSize}px`);
    rootStyle.setProperty("--editor-line-height", String(state.settings.lineHeight));
    rootStyle.setProperty("--editor-font-family", FONT_PRESETS[state.settings.fontFamily] || FONT_PRESETS.sans);

    el.editorPages.style.minHeight = `${EDITOR_MIN_HEIGHT_PX}px`;

    el.sheetBody.style.paddingTop = `${metrics.marginPx.top}px`;
    el.sheetBody.style.paddingRight = `${metrics.marginPx.right}px`;
    el.sheetBody.style.paddingBottom = `${metrics.marginPx.bottom}px`;
    el.sheetBody.style.paddingLeft = `${metrics.marginPx.left}px`;
    el.sheetBody.style.width = `${metrics.pageWidthPx}px`;
    el.sheetBody.style.fontSize = `${state.settings.fontSize}px`;
    el.sheetBody.style.lineHeight = String(state.settings.lineHeight);
    el.sheetBody.style.fontFamily = FONT_PRESETS[state.settings.fontFamily] || FONT_PRESETS.sans;

    el.sheetBody.style.whiteSpace = "pre-wrap";
    el.sheetBody.style.overflowWrap = "break-word";
    el.sheetBody.style.overflowX = "hidden";
    if (!el.editorView.classList.contains("editor-zen")) {
      zenTargetTextColumnWidthPx = metrics.contentWidthPx;
    } else {
      alignZenTextColumn(metrics);
    }

    updatePrintRules(metrics);
    refreshPageFrames();
    scheduleTypewriterSync({ force: true });
    if (!hasInitializedSheetLayout) {
      hasInitializedSheetLayout = true;
      window.requestAnimationFrame(() => {
        el.sheetBody.classList.add("layout-ready");
      });
    }
  }

  function getSafePrintMarginsInches(metrics) {
    const sheetMargins = getActiveSheetMargins();
    const minContentIn = 1 / PAGE_DPI;
    const safe = {
      top: clamp(Number(sheetMargins.top || 0), 0, metrics.paper.heightIn),
      right: clamp(Number(sheetMargins.right || 0), 0, metrics.paper.widthIn),
      bottom: clamp(Number(sheetMargins.bottom || 0), 0, metrics.paper.heightIn),
      left: clamp(Number(sheetMargins.left || 0), 0, metrics.paper.widthIn)
    };
    const maxVerticalMargins = Math.max(0, metrics.paper.heightIn - minContentIn);
    const maxHorizontalMargins = Math.max(0, metrics.paper.widthIn - minContentIn);

    const verticalSum = safe.top + safe.bottom;
    if (verticalSum > maxVerticalMargins && verticalSum > 0) {
      const scale = maxVerticalMargins / verticalSum;
      safe.top *= scale;
      safe.bottom *= scale;
    }

    const horizontalSum = safe.left + safe.right;
    if (horizontalSum > maxHorizontalMargins && horizontalSum > 0) {
      const scale = maxHorizontalMargins / horizontalSum;
      safe.left *= scale;
      safe.right *= scale;
    }

    return safe;
  }

  function updatePrintRules(metrics) {
    const printMargins = getSafePrintMarginsInches(metrics);
    let styleEl = document.getElementById("printRules");
    if (!styleEl) {
      styleEl = document.createElement("style");
      styleEl.id = "printRules";
      document.head.appendChild(styleEl);
    }
    styleEl.textContent = `
      @page {
        size: ${metrics.paper.printKeyword} portrait;
        margin: ${printMargins.top}in ${printMargins.right}in ${printMargins.bottom}in ${printMargins.left}in;
      }
      @media print {
        #printDocument {
          font-family: ${FONT_PRESETS[state.settings.fontFamily]};
          font-size: ${state.settings.fontSize}px;
          line-height: ${state.settings.lineHeight};
          tab-size: 4;
        }
      }
    `;
  }

  function renderFloatingMarginGuides(metrics, totalHeight) {
    Array.from(el.editorView.querySelectorAll(".margin-guide-floating")).forEach((guide) => guide.remove());
    if (!el.editorMain || !el.editorPages) return;

    const mainRect = el.editorMain.getBoundingClientRect();
    const pagesRect = el.editorPages.getBoundingClientRect();
    if (mainRect.height <= 0 || pagesRect.width <= 0) return;

    const guideTop = Math.round(mainRect.top);
    const guideHeight = Math.max(0, Math.round(mainRect.height));
    const pageLeft = Math.round(pagesRect.left);

    const rightGuide = document.createElement("span");
    rightGuide.className = "margin-guide margin-guide-right margin-guide-floating";
    rightGuide.dataset.marginSide = "right";
    rightGuide.style.position = "fixed";
    rightGuide.style.top = `${guideTop}px`;
    rightGuide.style.left = `${pageLeft + metrics.pageWidthPx - metrics.marginPx.right}px`;
    rightGuide.style.height = `${guideHeight}px`;
    rightGuide.setAttribute("role", "slider");
    rightGuide.setAttribute("aria-label", "Right margin marker");
    rightGuide.setAttribute("aria-orientation", "vertical");
    rightGuide.addEventListener("pointerdown", (event) => startLineBreakDrag(event, "right"));
    el.editorView.appendChild(rightGuide);

    const leftGuide = document.createElement("span");
    leftGuide.className = "margin-guide margin-guide-left margin-guide-floating";
    leftGuide.dataset.marginSide = "left";
    leftGuide.style.position = "fixed";
    leftGuide.style.top = `${guideTop}px`;
    leftGuide.style.left = `${pageLeft + metrics.marginPx.left}px`;
    leftGuide.style.height = `${guideHeight}px`;
    leftGuide.setAttribute("role", "slider");
    leftGuide.setAttribute("aria-label", "Left margin marker");
    leftGuide.setAttribute("aria-orientation", "vertical");
    leftGuide.addEventListener("pointerdown", (event) => startLineBreakDrag(event, "left"));
    el.editorView.appendChild(leftGuide);
  }

  function refreshPageFrames() {
    const metrics = computePageMetrics();
    el.sheetBody.style.height = "auto";
    const requiredHeight = Math.max(EDITOR_MIN_HEIGHT_PX, el.sheetBody.scrollHeight + EDITOR_SCROLL_BUFFER_PX);
    const pageCount = Math.max(1, Math.ceil(requiredHeight / metrics.pageHeightPx));
    const totalHeight = requiredHeight;

    el.editorPages.style.height = "100%";
    el.sheetBody.style.height = `${totalHeight}px`;

    el.pageFrames.innerHTML = "";
    for (let i = 0; i < pageCount; i += 1) {
      const frame = document.createElement("div");
      frame.className = "page-frame";
      frame.style.top = `${i * metrics.pageHeightPx}px`;
      const remainingHeight = totalHeight - i * metrics.pageHeightPx;
      frame.style.height = `${Math.max(0, Math.min(metrics.pageHeightPx, remainingHeight))}px`;

      const label = document.createElement("span");
      label.className = "page-label";
      label.textContent = `Page ${i + 1}`;
      frame.appendChild(label);
      el.pageFrames.appendChild(frame);
    }

    renderFloatingMarginGuides(metrics, totalHeight);
  }

  function ensureTypewriterMirrorElement() {
    if (typewriterMirrorEl) return typewriterMirrorEl;
    const mirror = document.createElement("div");
    mirror.setAttribute("aria-hidden", "true");
    Object.assign(mirror.style, {
      position: "fixed",
      left: "-10000px",
      top: "-10000px",
      visibility: "hidden",
      pointerEvents: "none",
      zIndex: "-1",
      whiteSpace: "pre-wrap",
      overflowWrap: "break-word",
      wordBreak: "break-word"
    });
    typewriterMirrorBeforeNode = document.createTextNode("");
    typewriterMirrorMarker = document.createElement("span");
    typewriterMirrorMarker.textContent = "\u200b";
    mirror.appendChild(typewriterMirrorBeforeNode);
    mirror.appendChild(typewriterMirrorMarker);
    document.body.appendChild(mirror);
    typewriterMirrorEl = mirror;
    return mirror;
  }

  function getTypewriterLineHeightPx(style) {
    const rawLineHeight = Number.parseFloat(style.lineHeight || "");
    if (Number.isFinite(rawLineHeight) && rawLineHeight > 0) {
      return rawLineHeight;
    }
    const fontSize = Number.parseFloat(style.fontSize || "16");
    return Math.max(1, fontSize * 1.2);
  }

  function getTypewriterCaretMetrics(textarea) {
    if (!textarea) return null;
    const style = window.getComputedStyle(textarea);
    const mirror = ensureTypewriterMirrorElement();
    if (!mirror || !typewriterMirrorBeforeNode || !typewriterMirrorMarker) return null;

    const mirroredProperties = [
      "boxSizing",
      "fontFamily",
      "fontSize",
      "fontWeight",
      "fontStyle",
      "fontVariant",
      "fontStretch",
      "lineHeight",
      "letterSpacing",
      "wordSpacing",
      "textTransform",
      "textIndent",
      "tabSize",
      "paddingTop",
      "paddingRight",
      "paddingBottom",
      "paddingLeft",
      "borderTopWidth",
      "borderRightWidth",
      "borderBottomWidth",
      "borderLeftWidth"
    ];

    const nextMirrorWidth = textarea.clientWidth;
    const nextStyleKey = mirroredProperties.map((name) => style[name]).join("|");
    if (nextStyleKey !== typewriterMirrorStyleKey || nextMirrorWidth !== typewriterMirrorWidth) {
      mirror.style.width = `${nextMirrorWidth}px`;
      mirroredProperties.forEach((name) => {
        mirror.style[name] = style[name];
      });
      mirror.style.whiteSpace = "pre-wrap";
      mirror.style.overflowWrap = "break-word";
      mirror.style.wordBreak = "break-word";
      mirror.style.borderStyle = "solid";
      typewriterMirrorStyleKey = nextStyleKey;
      typewriterMirrorWidth = nextMirrorWidth;
    }

    const value = String(textarea.value || "");
    const caretIndex = clamp(Number(textarea.selectionEnd || 0), 0, value.length);
    const before = value.slice(0, caretIndex);
    typewriterMirrorBeforeNode.nodeValue = before;

    const lineHeightPx = getTypewriterLineHeightPx(style);
    const caretTop = typewriterMirrorMarker.offsetTop;
    const visualLine = Math.floor(caretTop / Math.max(1, lineHeightPx));

    return {
      caretTop,
      lineHeightPx,
      visualLine
    };
  }

  function syncTypewriterPosition(options = {}) {
    const { force = false } = options;
    if (!state.ui.typewriterMode) {
      resetTypewriterTracking();
      return;
    }
    if (!state.activeSheetId || !el.sheetBody || !el.editorMain || document.activeElement !== el.sheetBody) return;
    const value = String(el.sheetBody.value || "");
    const caretIndex = clamp(Number(el.sheetBody.selectionEnd || 0), 0, value.length);
    if (
      !force &&
      caretIndex === lastTypewriterMeasuredCaretIndex &&
      typewriterValueRevision === lastTypewriterMeasureRevision
    ) {
      return;
    }
    const metrics = getTypewriterCaretMetrics(el.sheetBody);
    if (!metrics) return;
    lastTypewriterMeasuredCaretIndex = caretIndex;
    lastTypewriterMeasureRevision = typewriterValueRevision;

    const lineChanged = metrics.visualLine !== lastTypewriterVisualLine;
    if (!force && !lineChanged) return;

    lastTypewriterVisualLine = metrics.visualLine;
    const editorRect = el.editorMain.getBoundingClientRect();
    const textareaRect = el.sheetBody.getBoundingClientRect();
    const caretViewportTop = textareaRect.top - el.sheetBody.scrollTop + metrics.caretTop;
    const caretTopInEditorContent = caretViewportTop - editorRect.top + el.editorMain.scrollTop;
    const targetScrollTop = caretTopInEditorContent - el.editorMain.clientHeight / 2 + metrics.lineHeightPx / 2;
    const maxScrollTop = Math.max(0, el.editorMain.scrollHeight - el.editorMain.clientHeight);
    const nextScrollTop = clamp(Math.round(targetScrollTop), 0, maxScrollTop);
    if (Math.abs(el.editorMain.scrollTop - nextScrollTop) > 0.5) {
      el.editorMain.scrollTop = nextScrollTop;
    }
  }

  function scheduleTypewriterSync(options = {}) {
    if (!state.ui.typewriterMode) {
      resetTypewriterTracking();
      return;
    }
    const { force = false } = options;
    typewriterPendingForce = typewriterPendingForce || !!force;
    if (typewriterSyncRaf != null) return;
    typewriterSyncRaf = window.requestAnimationFrame(() => {
      typewriterSyncRaf = null;
      const runForce = typewriterPendingForce;
      typewriterPendingForce = false;
      syncTypewriterPosition({ force: runForce });
    });
  }

  function startLineBreakDrag(event, side = null) {
    const dragSide = side || getMarginGuideSideAtPointer(event);
    if (dragSide !== "left" && dragSide !== "right") return;
    event.preventDefault();
    dragState.active = true;
    dragState.side = dragSide;
    document.body.classList.add("dragging-linebreak");
    if (event.currentTarget && typeof event.currentTarget.setPointerCapture === "function" && event.pointerId != null) {
      try {
        event.currentTarget.setPointerCapture(event.pointerId);
      } catch (_) {
        // no-op
      }
    }
    window.addEventListener("pointermove", onLineBreakDrag);
    window.addEventListener("pointerup", stopLineBreakDrag, { once: true });
    onLineBreakDrag(event);
  }

  function onLineBreakDrag(event) {
    if (!dragState.active) return;
    const metrics = computePageMetrics();
    const rect = el.editorPages.getBoundingClientRect();
    const pointerX = clamp(event.clientX - rect.left, 0, metrics.pageWidthPx);
    const margins = getActiveSheetMargins();
    const activeSheet = getActiveSheet();

    if (dragState.side === "left") {
      const leftPx = clamp(pointerX, 0, metrics.pageWidthPx);
      margins.left = Number((leftPx / PAGE_DPI).toFixed(2));
      el.marginLeftInput.value = String(margins.left);
    } else {
      const rightPx = clamp(metrics.pageWidthPx - pointerX, 0, metrics.pageWidthPx);
      margins.right = Number((rightPx / PAGE_DPI).toFixed(2));
      el.marginRightInput.value = String(margins.right);
    }

    if (activeSheet) {
      activeSheet.updatedAt = nowISO();
    }

    applyTypographyAndGeometry();
  }

  function stopLineBreakDrag() {
    dragState.active = false;
    dragState.side = null;
    document.body.classList.remove("dragging-linebreak");
    window.removeEventListener("pointermove", onLineBreakDrag);
    saveState();
  }

  function getMarginGuideSideAtPointer(event) {
    const metrics = computePageMetrics();
    const rect = el.editorPages.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const leftGuideX = metrics.marginPx.left;
    const rightGuideX = metrics.pageWidthPx - metrics.marginPx.right;
    const leftDistance = Math.abs(x - leftGuideX);
    const rightDistance = Math.abs(x - rightGuideX);
    const threshold = 14;

    if (leftDistance > threshold && rightDistance > threshold) {
      return null;
    }
    if (leftDistance <= rightDistance) {
      return "left";
    }
    return "right";
  }

  function formatDateTime(iso) {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return "-";
    return d.toLocaleString();
  }

  function setView(name) {
    el.loginView.classList.toggle("hidden", name !== "login");
    el.libraryView.classList.toggle("hidden", name !== "library");
    el.editorView.classList.toggle("hidden", name !== "editor");
    renderZenMode();
  }

  function updateEditorVisualState() {
    const sheet = getActiveSheet();
    const hasContent = !!(sheet && sheet.body && sheet.body.length);
    el.editorView.classList.toggle("editor-has-content", hasContent);
    el.editorView.classList.toggle("editor-empty", !hasContent);
  }

  function renderZenMode() {
    const shouldEnable = !!state.ui.zenMode && !!state.activeSheetId && !el.editorView.classList.contains("hidden");
    el.editorView.classList.toggle("editor-zen", shouldEnable);
  }

  function renderTagsVisibility() {
    if (!el.tagsPanel) return;
    const shouldShow =
      !!state.auth.loggedIn &&
      !!state.activeSheetId &&
      !!state.ui.settingsOpen &&
      !state.ui.tagsHidden &&
      !state.ui.zenMode;
    el.tagsPanel.classList.toggle("hidden", !shouldShow);
  }

  function updateMeta() {
    const sheet = getActiveSheet();
    if (!sheet) {
      el.editorMeta.innerHTML = "";
      return;
    }
    const words = sheet.body.trim() ? sheet.body.trim().split(/\s+/).length : 0;
    const chars = sheet.body.length;
    el.editorMeta.innerHTML = `
      <div class="meta-block">
        <strong>${words} words</strong>
        <span>${chars} chars</span>
        <span>created ${formatDateTime(sheet.createdAt)}</span>
        <span>last edited ${formatDateTime(sheet.updatedAt)}</span>
      </div>
    `;
  }

  function renderTags() {
    const sheet = getActiveSheet();
    renderTagsVisibility();
    el.tagChips.innerHTML = "";
    tagComposerState.open = null;
    tagComposerState.close = null;
    if (!sheet) {
      return;
    }

    ensureSheetTags(sheet).forEach((tag) => {
      const chip = document.createElement("span");
      chip.className = "chip";
      const label = document.createElement("span");
      label.className = "chip-label";
      label.textContent = tag;
      chip.appendChild(label);

      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "chip-remove";
      remove.textContent = "×";
      remove.setAttribute("aria-label", `Remove tag ${tag}`);
      remove.addEventListener("click", () => {
        if (!removeTagFromSheet(sheet, tag)) return;
        sheet.updatedAt = nowISO();
        saveState();
        renderTags();
      });
      chip.appendChild(remove);
      el.tagChips.appendChild(chip);
    });

    const addTag = document.createElement("button");
    addTag.type = "button";
    addTag.className = "chip-add";
    addTag.textContent = "add tag";
    addTag.setAttribute("aria-label", "Add a tag");

    const addForm = document.createElement("form");
    addForm.className = "chip-add-form hidden";

    const addInput = document.createElement("input");
    addInput.type = "text";
    addInput.className = "chip-input";
    addInput.placeholder = "+tag / -tag";
    addInput.setAttribute("aria-label", "Add or remove tag");
    addForm.appendChild(addInput);

    const applyBtn = document.createElement("button");
    applyBtn.type = "submit";
    applyBtn.className = "chip-apply";
    applyBtn.textContent = "ok";
    addForm.appendChild(applyBtn);

    function closeAddForm() {
      addInput.value = "";
      addForm.classList.add("hidden");
      addTag.classList.remove("hidden");
    }

    function openAddForm() {
      addTag.classList.add("hidden");
      addForm.classList.remove("hidden");
      addInput.focus();
      addInput.select();
    }

    tagComposerState.open = openAddForm;
    tagComposerState.close = closeAddForm;

    addTag.addEventListener("click", () => {
      openAddForm();
    });

    addInput.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        closeAddForm();
      }
    });

    addInput.addEventListener("blur", () => {
      window.setTimeout(() => {
        if (document.activeElement && addForm.contains(document.activeElement)) {
          return;
        }
        closeAddForm();
      }, 0);
    });

    addForm.addEventListener("submit", (event) => {
      event.preventDefault();
      if (applyTagOperation(sheet, addInput.value)) {
        renderTags();
        return;
      }
      addInput.focus();
      addInput.select();
    });

    el.tagChips.appendChild(addTag);
    el.tagChips.appendChild(addForm);
  }

  function renderLibrary() {
    const stack = getCurrentStack();
    el.libraryView.classList.toggle("library-stack", stack.id !== state.rootId);
    renderStackBreadcrumb(el.libraryBreadcrumb, getStackTrail(stack.id), { clickable: true });
    el.libraryGrid.innerHTML = "";

    const newSheetCard = makeActionCard("New Notation >", () => {
      const parent = getCurrentStack();
      const sheet = createSheet({ title: "Untitled", body: "", parentId: parent.id });
      state.containers[sheet.id] = sheet;
      parent.children.unshift(sheet.id);
      parent.updatedAt = nowISO();
      state.ui.selectedCardId = sheet.id;
      saveState();
      renderLibrary();
    });
    const newStackCard = makeActionCard("New Stack >", () => {
      const parent = getCurrentStack();
      const stackNode = createStack({ title: "Untitled Stack", parentId: parent.id });
      state.containers[stackNode.id] = stackNode;
      parent.children.unshift(stackNode.id);
      parent.updatedAt = nowISO();
      state.ui.selectedCardId = stackNode.id;
      saveState();
      renderLibrary();
    });

    el.libraryGrid.appendChild(newSheetCard);
    el.libraryGrid.appendChild(newStackCard);

    const ids = Array.isArray(stack.children) ? stack.children : [];
    ids.forEach((id) => {
      const node = state.containers[id];
      if (!node) return;
      const card = node.kind === "stack" ? makeStackCard(node) : makeSheetCard(node);
      if (state.ui.selectedCardId === node.id) {
        card.classList.add("selected");
      }
      el.libraryGrid.appendChild(card);
    });
  }

  function makeActionCard(label, onClick) {
    const card = document.createElement("button");
    card.type = "button";
    card.className = "card new";
    card.innerHTML = `<h3 class="card-title">${label}</h3>`;
    card.addEventListener("click", onClick);
    return card;
  }

  function countSheetsInStack(stackId) {
    const visitedStackIds = new Set();

    function walk(currentStackId) {
      if (visitedStackIds.has(currentStackId)) return 0;
      visitedStackIds.add(currentStackId);

      const current = state.containers[currentStackId];
      if (!current || current.kind !== "stack") return 0;

      const childIds = Array.isArray(current.children) ? current.children : [];
      let total = 0;
      childIds.forEach((childId) => {
        const child = state.containers[childId];
        if (!child) return;
        if (child.kind === "sheet") {
          total += 1;
          return;
        }
        if (child.kind === "stack") {
          total += walk(child.id);
        }
      });
      return total;
    }

    return walk(stackId);
  }

  function makeStackCard(node) {
    const card = document.createElement("button");
    card.type = "button";
    card.className = "card card-stack";
    const visibleCount = countSheetsInStack(node.id);
    card.innerHTML = `
      <div class="stack-icon">
        <span class="stack-count">${visibleCount}</span>
      </div>
      <div class="stack-meta">${node.title}</div>
    `;

    const del = document.createElement("button");
    del.type = "button";
    del.className = "delete-card";
    del.textContent = "×";
    del.addEventListener("click", (e) => {
      e.stopPropagation();
      removeNode(node.id);
    });
    card.appendChild(del);

    card.addEventListener("click", () => {
      navigateToLibrary(node.id, { syncHash: true, persist: true });
    });
    return card;
  }

  function makeSheetCard(node) {
    const card = document.createElement("button");
    card.type = "button";
    card.className = "card";
    const body = node.body || "";
    const preview = body ? body.slice(0, 160) : "";
    const isDateCard = /^\d{2}-\d{2}-\d{2,4}$/.test(node.title.trim());
    if (!body && isDateCard) {
      card.classList.add("card-date");
      card.innerHTML = `<h3 class="card-title">${node.title}</h3>`;
    } else if (!body) {
      card.classList.add("card-title-only");
      if (node.title.trim().length <= 4) {
        card.classList.add("card-title-jumbo");
      }
      card.innerHTML = `<h3 class="card-title">${node.title}</h3>`;
    } else {
      card.innerHTML = `
        <h3 class="card-title">${node.title}</h3>
        <div class="card-preview">${preview}</div>
      `;
    }

    applyDynamicCardTitleSizing(card, node.title);

    const del = document.createElement("button");
    del.type = "button";
    del.className = "delete-card";
    del.textContent = "×";
    del.addEventListener("click", (e) => {
      e.stopPropagation();
      removeNode(node.id);
    });
    card.appendChild(del);

    card.addEventListener("click", () => {
      navigateToSheet(node.id, { syncHash: true, persist: true });
    });
    return card;
  }

  async function removeNode(nodeId) {
    const node = state.containers[nodeId];
    if (!node) return;

    if (node.kind === "stack") {
      const stackDeleteChoice = await showChoiceDialog({
        title: `Delete "${node.title}"?`,
        message: "Delete Stack Only keeps the notes and moves them to Unstacked.",
        cancelValue: "cancel",
        actions: [
          { id: "delete-stack-only", label: "Delete Stack Only", variant: "primary", autoFocus: true },
          { id: "delete-stack-and-notes", label: "Delete Stack + Notes", variant: "danger" },
          { id: "cancel", label: "Cancel", variant: "secondary" }
        ]
      });

      if (stackDeleteChoice === "delete-stack-and-notes") {
        const confirmed = await showChoiceDialog({
          title: "Are you sure?",
          message: `This will delete "${node.title}" and every note inside it.`,
          cancelValue: "cancel",
          actions: [
            { id: "confirm", label: "Delete Everything", variant: "danger", autoFocus: true },
            { id: "cancel", label: "Cancel", variant: "secondary" }
          ]
        });
        if (confirmed !== "confirm") return;
        deleteNodeRecursive(nodeId);
      } else if (stackDeleteChoice === "delete-stack-only") {
        deleteStackAndUnstackNotes(nodeId);
      } else {
        return;
      }
      saveState();
      renderLibrary();
      return;
    }

    const confirmed = await showChoiceDialog({
      title: `Delete "${node.title}"?`,
      message: "This note will be permanently deleted.",
      cancelValue: "cancel",
      actions: [
        { id: "confirm", label: "Delete Note", variant: "danger", autoFocus: true },
        { id: "cancel", label: "Cancel", variant: "secondary" }
      ]
    });
    if (confirmed !== "confirm") return;

    deleteNodeRecursive(nodeId);
    saveState();
    renderLibrary();
  }

  function deleteStackAndUnstackNotes(stackId) {
    const stack = state.containers[stackId];
    const root = state.containers[state.rootId];
    if (!stack || stack.kind !== "stack" || !root || root.kind !== "stack" || stackId === state.rootId) return;

    const descendantStackIds = [];
    const descendantSheetIds = [];

    function collectDescendants(currentStackId) {
      const current = state.containers[currentStackId];
      if (!current || current.kind !== "stack") return;

      descendantStackIds.push(currentStackId);
      const childIds = Array.isArray(current.children) ? current.children : [];
      childIds.forEach((childId) => {
        const child = state.containers[childId];
        if (!child) return;
        if (child.kind === "stack") {
          collectDescendants(child.id);
          return;
        }
        descendantSheetIds.push(child.id);
      });
    }

    collectDescendants(stackId);
    const stackIdSet = new Set(descendantStackIds);
    const seenSheets = new Set();
    const sheetsToMove = [];
    descendantSheetIds.forEach((sheetId) => {
      if (seenSheets.has(sheetId)) return;
      seenSheets.add(sheetId);
      sheetsToMove.push(sheetId);
    });

    sheetsToMove.forEach((sheetId) => {
      const sheet = state.containers[sheetId];
      if (!sheet || sheet.kind !== "sheet") return;
      sheet.parentId = state.rootId;
    });

    descendantStackIds.forEach((id) => {
      const descendant = state.containers[id];
      if (!descendant || descendant.kind !== "stack") return;
      descendant.children = (Array.isArray(descendant.children) ? descendant.children : []).filter((childId) => {
        const child = state.containers[childId];
        return child && child.kind === "stack";
      });
    });

    root.children = (Array.isArray(root.children) ? root.children : []).filter((id) => !stackIdSet.has(id));
    const rootChildSet = new Set(root.children);
    const movedToRoot = sheetsToMove.filter((sheetId) => !rootChildSet.has(sheetId));
    root.children = movedToRoot.concat(root.children);

    deleteNodeRecursive(stackId);
  }

  function deleteNodeRecursive(nodeId) {
    const node = state.containers[nodeId];
    if (!node) return;

    const parent = node.parentId ? state.containers[node.parentId] : null;
    if (parent && Array.isArray(parent.children)) {
      parent.children = parent.children.filter((id) => id !== nodeId);
    }

    if (node.kind === "stack") {
      node.children.forEach((childId) => deleteNodeRecursive(childId));
    }

    if (state.activeSheetId === nodeId) {
      state.activeSheetId = null;
    }
    if (state.currentStackId === nodeId) {
      state.currentStackId = state.rootId;
    }
    delete state.containers[nodeId];
  }

  function renderEditor() {
    const sheet = getActiveSheet();
    if (!sheet) {
      navigateToLibrary(state.currentStackId || state.rootId, { syncHash: true, persist: true });
      return;
    }

    renderStackBreadcrumb(el.editorBreadcrumbTrail, getStackTrail(sheet.parentId || state.rootId), { clickable: true });
    el.sheetTitle.value = sheet.title;
    el.sheetTitle.placeholder = "Untitled";
    el.sheetSubtitle.textContent = sheet.subtitle || "";
    el.sheetSubtitle.style.display = sheet.subtitle ? "inline" : "none";
    el.sheetBody.value = sheet.body;
    el.sheetBody.placeholder = EMPTY_PLACEHOLDER;
    state.settings.fontSize = normalizeFontSize(state.settings.fontSize);
    el.quickFontFamily.value = state.settings.fontFamily;
    el.quickFontSize.value = String(state.settings.fontSize);

    el.paperSizeSelect.value = state.settings.paperSize;
    el.fontFamilySelect.value = state.settings.fontFamily;
    el.fontSizeInput.value = String(state.settings.fontSize);
    const margins = getActiveSheetMargins();
    syncMarginInputBounds();
    el.marginTopInput.value = String(margins.top);
    el.marginRightInput.value = String(margins.right);
    el.marginBottomInput.value = String(margins.bottom);
    el.marginLeftInput.value = String(margins.left);

    applyTypographyAndGeometry();
    renderTags();
    updateMeta();
    updateEditorVisualState();
    typewriterValueRevision += 1;
    resetTypewriterTracking();
    scheduleTypewriterSync({ force: true });
  }

  function navigateToLogin(options = {}) {
    const { syncHash = true, persist = true } = options;
    if (commandPaletteState.open) {
      closeCommandPalette({ restoreFocus: false });
    }
    state.activeSheetId = null;
    state.currentStackId = state.rootId;
    state.ui.selectedCardId = null;
    if (persist) {
      saveState();
    }
    setView("login");
    if (syncHash) {
      syncHashToState();
    }
    return true;
  }

  function navigateToLibrary(stackId, options = {}) {
    const { syncHash = true, persist = true } = options;
    if (commandPaletteState.open) {
      closeCommandPalette({ restoreFocus: false });
    }
    const stack = state.containers[stackId];
    if (!stack || stack.kind !== "stack") {
      return false;
    }
    state.currentStackId = stack.id;
    state.activeSheetId = null;
    state.ui.selectedCardId = stack.id === state.rootId ? null : stack.id;
    if (persist) {
      saveState();
    }
    setView("library");
    renderLibrary();
    renderSettingsVisibility();
    if (syncHash) {
      syncHashToState();
    }
    return true;
  }

  function navigateToSheet(sheetId, options = {}) {
    const { syncHash = true, persist = true } = options;
    if (commandPaletteState.open) {
      closeCommandPalette({ restoreFocus: false });
    }
    const sheet = state.containers[sheetId];
    if (!sheet || sheet.kind !== "sheet") {
      return false;
    }
    state.activeSheetId = sheet.id;
    state.currentStackId = sheet.parentId || state.rootId;
    state.ui.selectedCardId = sheet.id;
    if (persist) {
      saveState();
    }
    setView("editor");
    renderEditor();
    renderSettingsVisibility();
    if (syncHash) {
      syncHashToState();
    }
    return true;
  }

  function applyRouteSegments(segments, options = {}) {
    const { syncHash = true, persist = true } = options;
    const resolved = resolveRoute(segments);
    if (!resolved) {
      return false;
    }

    if (resolved.kind === "login") {
      if (!state.auth.loggedIn) {
        if (syncHash) {
          const loginHash = buildHashFromSegments(["login"]);
          if (window.location.hash !== loginHash) {
            suppressHashChange = true;
            window.location.hash = loginHash;
          }
        }
        setView("login");
        return true;
      }
      return navigateToLibrary(state.rootId, { syncHash, persist });
    }

    if (!state.auth.loggedIn) {
      if (syncHash) {
        const nextHash = buildHashFromSegments(segments);
        if (window.location.hash !== nextHash) {
          suppressHashChange = true;
          window.location.hash = nextHash;
        }
      }
      setView("login");
      return true;
    }

    if (resolved.kind === "library") {
      return navigateToLibrary(resolved.stackId, { syncHash, persist });
    }
    if (resolved.kind === "editor") {
      return navigateToSheet(resolved.sheetId, { syncHash, persist });
    }
    return false;
  }

  function applyRouteFromLocation() {
    return applyRouteSegments(readRouteSegmentsFromLocation(), { syncHash: true, persist: true });
  }

  function goHome() {
    state.ui.settingsOpen = false;
    navigateToLibrary(state.rootId, { syncHash: true, persist: true });
  }

  function renderSettingsVisibility() {
    const isOpen = !!state.ui.settingsOpen && !state.ui.zenMode;
    el.settingsPanel.classList.add("hidden");
    el.editorView.classList.toggle("editor-controls-open", isOpen);
    renderTagsVisibility();
  }

  function updateSettingsFromForm() {
    const nextPaper = PAPER_PRESETS[el.paperSizeSelect.value] ? el.paperSizeSelect.value : "letter";
    const paper = PAPER_PRESETS[nextPaper] || PAPER_PRESETS.letter;
    const layoutWidthIn = getEditorLayoutWidthPx() / PAGE_DPI;
    const nextFamily = FONT_PRESETS[el.fontFamilySelect.value] ? el.fontFamilySelect.value : "monospace";
    const nextSize = normalizeFontSize(el.fontSizeInput.value || DEFAULT_SETTINGS.fontSize);

    state.settings.paperSize = nextPaper;
    state.settings.fontFamily = nextFamily;
    state.settings.fontSize = nextSize;
    const margins = getActiveSheetMargins();
    margins.top = clamp(Number(el.marginTopInput.value || DEFAULT_SETTINGS.margins.top), 0, paper.heightIn);
    margins.right = clamp(Number(el.marginRightInput.value || DEFAULT_SETTINGS.margins.right), 0, layoutWidthIn);
    margins.bottom = clamp(Number(el.marginBottomInput.value || DEFAULT_SETTINGS.margins.bottom), 0, paper.heightIn);
    margins.left = clamp(Number(el.marginLeftInput.value || DEFAULT_SETTINGS.margins.left), 0, layoutWidthIn);
    syncMarginInputBounds({ widthIn: layoutWidthIn, heightIn: paper.heightIn });
    el.marginTopInput.value = String(margins.top);
    el.marginRightInput.value = String(margins.right);
    el.marginBottomInput.value = String(margins.bottom);
    el.marginLeftInput.value = String(margins.left);

    el.quickFontFamily.value = state.settings.fontFamily;
    el.quickFontSize.value = String(state.settings.fontSize);
    el.fontSizeInput.value = String(state.settings.fontSize);
    const activeSheet = getActiveSheet();
    if (activeSheet) {
      activeSheet.updatedAt = nowISO();
    }
    saveState();
    applyTypographyAndGeometry();
  }

  function setupEvents() {
    window.addEventListener("keydown", (event) => {
      if (shouldOpenCommandPalette(event)) {
        event.preventDefault();
        openCommandPalette();
      }
    });

    el.loginEnter.addEventListener("click", () => {
      state.auth.loggedIn = true;
      state.auth.username = el.usernameInput.value.trim();
      state.ui.selectedCardId = null;
      saveState();
      if (!applyRouteFromLocation()) {
        navigateToLibrary(state.currentStackId || state.rootId, { syncHash: true, persist: true });
      }
    });

    [el.usernameInput, el.passwordInput].forEach((input) => {
      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          el.loginEnter.click();
        }
      });
    });

    homeTriggers.forEach((btn) => btn.addEventListener("click", goHome));

    settingsTriggers.forEach((btn) =>
      btn.addEventListener("click", () => {
        state.ui.settingsOpen = !state.ui.settingsOpen;
        saveState();
        renderSettingsVisibility();
        if (state.activeSheetId) {
          requestAnimationFrame(() => applyTypographyAndGeometry());
        }
      })
    );

    el.commandPaletteScrim.addEventListener("click", () => {
      closeCommandPalette();
    });

    el.commandPaletteInput.addEventListener("input", () => {
      commandPaletteState.selectedIndex = 0;
      renderCommandPalette();
    });

    el.commandPaletteInput.addEventListener("keydown", (event) => {
      if (event.key === "ArrowDown") {
        event.preventDefault();
        moveCommandSelection(1);
        return;
      }
      if (event.key === "ArrowUp") {
        event.preventDefault();
        moveCommandSelection(-1);
        return;
      }
      if (event.key === "Enter") {
        event.preventDefault();
        const item = getSelectedCommandItem();
        if (!item || item.disabled) return;
        runCommandItem(item);
        return;
      }
      if (event.key === "Escape") {
        event.preventDefault();
        if (commandPaletteState.mode !== "root") {
          setCommandPaletteMode("root");
          return;
        }
        closeCommandPalette();
        return;
      }
      if (event.key === "Backspace" && !el.commandPaletteInput.value && commandPaletteState.mode !== "root") {
        event.preventDefault();
        setCommandPaletteMode("root");
      }
    });

    el.choiceDialog.addEventListener("click", (event) => {
      if (typeof activeDialogResolve !== "function") return;
      if (event.target === el.choiceDialog || event.target === el.choiceDialogScrim) {
        closeChoiceDialog(activeDialogCancelValue);
      }
    });

    window.addEventListener("keydown", (event) => {
      if (event.metaKey && !event.ctrlKey && !event.altKey && event.code === "Period") {
        if (typeof activeDialogResolve === "function") {
          return;
        }
        event.preventDefault();
        toggleOptionsCommand();
        return;
      }

      if (event.metaKey && !event.ctrlKey && !event.altKey && event.key === "Enter") {
        if (typeof activeDialogResolve === "function") {
          return;
        }
        event.preventDefault();
        toggleZenModeCommand();
        return;
      }

      if (event.metaKey && !event.ctrlKey && !event.altKey && !event.shiftKey && event.key.toLowerCase() === "t") {
        if (typeof activeDialogResolve === "function") {
          return;
        }
        event.preventDefault();
        openAddTagCommand();
        return;
      }

      if (event.key !== "Escape") return;
      if (state.ui.zenMode && !commandPaletteState.open && typeof activeDialogResolve !== "function") {
        event.preventDefault();
        setZenMode(false);
        return;
      }
      if (typeof activeDialogResolve !== "function") return;
      event.preventDefault();
      closeChoiceDialog(activeDialogCancelValue);
    });

    el.sheetTitle.addEventListener("input", () => {
      const sheet = getActiveSheet();
      if (!sheet) return;
      sheet.title = safeName(el.sheetTitle.value);
      sheet.updatedAt = nowISO();
      saveState();
      syncHashToState();
    });

    el.sheetBody.addEventListener("input", () => {
      const sheet = getActiveSheet();
      if (!sheet) return;
      sheet.body = el.sheetBody.value;
      sheet.updatedAt = nowISO();
      saveState();
      refreshPageFrames();
      updateMeta();
      updateEditorVisualState();
      typewriterValueRevision += 1;
      scheduleTypewriterSync();
    });

    el.sheetBody.addEventListener("keydown", (e) => {
      if (e.key === "Tab") {
        e.preventDefault();
        const start = el.sheetBody.selectionStart;
        const end = el.sheetBody.selectionEnd;
        const tab = "    ";
        el.sheetBody.setRangeText(tab, start, end, "end");
        el.sheetBody.dispatchEvent(new Event("input", { bubbles: true }));
      }
    });

    el.sheetBody.addEventListener("keyup", () => {
      scheduleTypewriterSync();
    });
    el.sheetBody.addEventListener("click", () => {
      scheduleTypewriterSync({ force: true });
    });
    el.sheetBody.addEventListener("focus", () => {
      resetTypewriterTracking();
      scheduleTypewriterSync({ force: true });
    });
    el.sheetBody.addEventListener("select", () => {
      scheduleTypewriterSync();
    });

    el.quickFontFamily.addEventListener("change", () => {
      state.settings.fontFamily = el.quickFontFamily.value;
      el.fontFamilySelect.value = state.settings.fontFamily;
      saveState();
      applyTypographyAndGeometry();
    });

    el.quickFontSize.addEventListener("change", () => {
      const n = normalizeFontSize(el.quickFontSize.value || DEFAULT_SETTINGS.fontSize);
      state.settings.fontSize = n;
      el.fontSizeInput.value = String(n);
      saveState();
      applyTypographyAndGeometry();
    });

    [
      el.paperSizeSelect,
      el.fontFamilySelect,
      el.fontSizeInput,
      el.marginTopInput,
      el.marginRightInput,
      el.marginBottomInput,
      el.marginLeftInput
    ].forEach((input) => input.addEventListener("change", updateSettingsFromForm));

    el.resetDefaults.addEventListener("click", () => {
      state.settings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
      const activeSheet = getActiveSheet();
      if (activeSheet) {
        activeSheet.margins = { ...DEFAULT_SETTINGS.margins };
        activeSheet.updatedAt = nowISO();
      }
      saveState();
      renderEditor();
    });

    el.printBtn.addEventListener("click", async () => {
      await runPrintCommand();
    });

    el.pdfBtn.addEventListener("click", async () => {
      await runPdfExportCommand();
    });

    el.txtBtn.addEventListener("click", async () => {
      await runTextExportCommand();
    });

    window.addEventListener("resize", () => {
      if (!state.activeSheetId) return;
      applyTypographyAndGeometry();
    });
    el.editorPages.addEventListener("pointerdown", (event) => {
      if (!state.ui.settingsOpen) return;
      if (event.target && event.target.closest && event.target.closest(".margin-guide")) return;
      const nearGuideSide = getMarginGuideSideAtPointer(event);
      if (!nearGuideSide) return;
      startLineBreakDrag(event, nearGuideSide);
    });
    window.addEventListener("hashchange", () => {
      if (suppressHashChange) {
        suppressHashChange = false;
        return;
      }
      if (!applyRouteFromLocation()) {
        navigateToLibrary(state.rootId, { syncHash: true, persist: true });
      }
    });

    if (window.notations && typeof window.notations.onDeepLink === "function") {
      window.notations.onDeepLink((routePath) => {
        const segments = readRouteSegmentsFromPath(routePath);
        if (!applyRouteSegments(segments, { syncHash: true, persist: true })) {
          navigateToLibrary(state.rootId, { syncHash: true, persist: true });
        }
      });
    }

    if (window.notations && typeof window.notations.onMenuPrintRequest === "function") {
      window.notations.onMenuPrintRequest(() => {
        runPrintCommand();
      });
    }
  }

  function safeFilename(name) {
    return safeName(name).replace(/[^\w.-]+/g, "_").toLowerCase();
  }

  function preparePrintDocument() {
    const sheet = getActiveSheet();
    if (!sheet) return;
    el.printDocument.textContent = sheet.body;
  }

  function buildPrintRequestPayload() {
    const metrics = computePageMetrics();
    const marginsIn = getSafePrintMarginsInches(metrics);
    const pageWidthPx = Math.round(metrics.paper.widthIn * PAGE_DPI);
    const pageHeightPx = Math.round(metrics.paper.heightIn * PAGE_DPI);
    let topPx = Math.max(0, Math.round(marginsIn.top * PAGE_DPI));
    let rightPx = Math.max(0, Math.round(marginsIn.right * PAGE_DPI));
    let bottomPx = Math.max(0, Math.round(marginsIn.bottom * PAGE_DPI));
    let leftPx = Math.max(0, Math.round(marginsIn.left * PAGE_DPI));

    const maxVerticalMarginsPx = Math.max(0, pageHeightPx - 1);
    const maxHorizontalMarginsPx = Math.max(0, pageWidthPx - 1);

    if (topPx + bottomPx > maxVerticalMarginsPx) {
      const scale = maxVerticalMarginsPx / (topPx + bottomPx);
      topPx = Math.floor(topPx * scale);
      bottomPx = Math.floor(bottomPx * scale);
    }
    if (leftPx + rightPx > maxHorizontalMarginsPx) {
      const scale = maxHorizontalMarginsPx / (leftPx + rightPx);
      leftPx = Math.floor(leftPx * scale);
      rightPx = Math.floor(rightPx * scale);
    }

    return {
      pageSize: metrics.paper.printKeyword,
      margins: {
        top: topPx,
        right: rightPx,
        bottom: bottomPx,
        left: leftPx
      }
    };
  }

  async function boot() {
    state = await loadState();
    populateFontSizeOptions();
    setupEvents();
    renderSettingsVisibility();

    if (window.notations && typeof window.notations.consumeInitialDeepLink === "function") {
      try {
        const initial = await window.notations.consumeInitialDeepLink();
        if (initial && initial.ok && initial.routePath) {
          const segments = readRouteSegmentsFromPath(initial.routePath);
          if (applyRouteSegments(segments, { syncHash: true, persist: true })) {
            return;
          }
        }
      } catch (_) {
        // continue with normal route resolution
      }
    }

    if (window.location.hash && applyRouteFromLocation()) {
      return;
    }

    if (!state.auth.loggedIn) {
      setView("login");
      return;
    }

    if (state.activeSheetId && state.containers[state.activeSheetId]) {
      navigateToSheet(state.activeSheetId, { syncHash: true, persist: false });
      return;
    }

    const fallbackStack =
      state.containers[state.currentStackId] && state.containers[state.currentStackId].kind === "stack"
        ? state.currentStackId
        : state.rootId;
    navigateToLibrary(fallbackStack, { syncHash: true, persist: false });
  }

  boot();
})();
