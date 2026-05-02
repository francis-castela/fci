const eventForm = document.getElementById("eventForm");
const eventsTableBody = document.querySelector("#eventsTable tbody");
const exportButton = document.getElementById("exportPng");
const exportCurrentButton = document.getElementById("exportCurrentPng");
const downloadTemplateButton = document.getElementById("downloadTemplate");
const importSheetButton = document.getElementById("importSheet");
const resetColorsButton = document.getElementById("resetColors");
const sheetFileInput = document.getElementById("sheetFile");
const importStatus = document.getElementById("importStatus");
const clearFormButton = document.getElementById("clearForm");
const loadSampleButton = document.getElementById("loadSample");
const previewCanvas = document.getElementById("previewCanvas");
const previewCtx = previewCanvas.getContext("2d");
const eventCount = document.getElementById("eventCount");
const emptyState = document.getElementById("emptyState");
const sortFieldSelect = document.getElementById("sortField");
const sortEventsButton = document.getElementById("sortEvents");
const prevPageButton = document.getElementById("prevPage");
const nextPageButton = document.getElementById("nextPage");
const pageInfo = document.getElementById("pageInfo");
const themeToggleButton = document.getElementById("themeToggle");
const themeToggleLabel = document.getElementById("themeToggleLabel");
const themeToggleIcon = document.getElementById("themeToggleIcon");

const titleInput = document.getElementById("agendaTitle");
const primaryColorInput = document.getElementById("primaryColor");
const darkColorInput = document.getElementById("darkColor");
const lightColorInput = document.getElementById("lightColor");

const fields = {
  date: document.getElementById("date"),
  time: document.getElementById("time"),
  name: document.getElementById("name"),
  producer: document.getElementById("producer"),
  ticket: document.getElementById("ticket"),
};

const state = {
  events: [],
  previewPage: 0,
};

let draggedRowIndex = null;

const OUTPUT_WIDTH = 1080;
const OUTPUT_HEIGHT = 1440;
const logoTeatro = new Image();
const logoFundacao = new Image();
const REQUIRED_COLUMNS = ["Dia/Mês", "Horário", "Nome do evento", "Produtor", "Onde comprar"];
const DEFAULT_THEME = {
  primary: "#b86848",
  dark: "#7f442f",
  light: "#c98971",
};
const COOKIE_PREFIX = "agenda_builder_";
const LOCAL_STORAGE_STATE_KEY = `${COOKIE_PREFIX}state_v2`;
const COOKIE_COUNT_KEY = `${COOKIE_PREFIX}chunks`;
const COOKIE_CHUNK_KEY = `${COOKIE_PREFIX}data_`;
const COOKIE_CHUNK_SIZE = 3500;
const COOKIE_MAX_AGE_DAYS = 180;
const COOKIE_THEME_KEY = `${COOKIE_PREFIX}theme`;
const XLSX_LIB_URLS = [
  "assets/vendor/xlsx.full.min.js",
  "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
];
const JSZIP_LIB_URLS = [
  "assets/vendor/jszip.min.js",
  "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js",
];

let xlsxLoadPromise = null;
let jszipLoadPromise = null;

function drawImageContain(ctx, image, x, y, maxWidth, maxHeight) {
  const ratio = Math.min(maxWidth / image.naturalWidth, maxHeight / image.naturalHeight);
  const drawWidth = image.naturalWidth * ratio;
  const drawHeight = image.naturalHeight * ratio;
  const drawX = x + (maxWidth - drawWidth) / 2;
  const drawY = y + (maxHeight - drawHeight) / 2;
  ctx.drawImage(image, drawX, drawY, drawWidth, drawHeight);
}

function loadLogos() {
  const triggerRender = () => renderPreview();

  logoTeatro.onload = triggerRender;
  logoFundacao.onload = triggerRender;

  logoTeatro.src = "assets/img/logo-teatro.png";
  logoFundacao.src = "assets/img/logo-fundacao-cultural.jpg";
}

function getTheme() {
  return {
    primary: primaryColorInput.value,
    dark: darkColorInput.value,
    light: lightColorInput.value,
  };
}

function getCurrentTheme() {
  return document.documentElement.dataset.theme === "dark" ? "dark" : "light";
}

function updateThemeToggleUi(theme) {
  if (theme === "dark") {
    themeToggleLabel.textContent = "Modo claro";
    themeToggleIcon.textContent = "☾";
  } else {
    themeToggleLabel.textContent = "Modo escuro";
    themeToggleIcon.textContent = "◐";
  }
}

function applyTheme(theme) {
  const normalizedTheme = theme === "dark" ? "dark" : "light";
  document.documentElement.dataset.theme = normalizedTheme;
  updateThemeToggleUi(normalizedTheme);
}

function restoreThemeFromCookie() {
  const savedTheme = getCookie(COOKIE_THEME_KEY);
  if (savedTheme === "dark" || savedTheme === "light") {
    applyTheme(savedTheme);
  } else {
    applyTheme("light");
  }
}

function persistThemeToCookie() {
  setCookie(COOKIE_THEME_KEY, getCurrentTheme());
}

function resetThemeToDefault() {
  primaryColorInput.value = DEFAULT_THEME.primary;
  darkColorInput.value = DEFAULT_THEME.dark;
  lightColorInput.value = DEFAULT_THEME.light;
}

function normalizeAgendaTitleInput() {
  const caretStart = titleInput.selectionStart;
  const caretEnd = titleInput.selectionEnd;
  titleInput.value = titleInput.value.toLocaleUpperCase("pt-BR");

  if (caretStart !== null && caretEnd !== null) {
    titleInput.setSelectionRange(caretStart, caretEnd);
  }
}

function setCookie(name, value, days = COOKIE_MAX_AGE_DAYS) {
  const maxAge = days * 24 * 60 * 60;
  document.cookie = `${name}=${value}; max-age=${maxAge}; path=/; samesite=lax`;
}

function getCookie(name) {
  const key = `${name}=`;
  const cookies = document.cookie.split(";");

  for (const rawCookie of cookies) {
    const cookie = rawCookie.trim();
    if (cookie.startsWith(key)) {
      return cookie.slice(key.length);
    }
  }

  return "";
}

function deleteCookie(name) {
  document.cookie = `${name}=; max-age=0; path=/; samesite=lax`;
}

function clearStateCookies() {
  const previousCount = Number.parseInt(getCookie(COOKIE_COUNT_KEY), 10) || 0;

  for (let i = 0; i < previousCount; i += 1) {
    deleteCookie(`${COOKIE_CHUNK_KEY}${i}`);
  }

  deleteCookie(COOKIE_COUNT_KEY);
}

function toSafeEvent(event) {
  return {
    date: String(event.date || "").trim(),
    time: String(event.time || "").trim(),
    name: String(event.name || "").trim(),
    producer: String(event.producer || "").trim(),
    ticket: String(event.ticket || "").trim(),
  };
}

function persistStateToCookies() {
  try {
    const snapshot = {
      title: titleInput.value,
      theme: getTheme(),
      themeMode: getCurrentTheme(),
      events: state.events.map(toSafeEvent),
    };

    localStorage.setItem(LOCAL_STORAGE_STATE_KEY, JSON.stringify(snapshot));
  } catch (error) {
    console.warn("Não foi possível salvar o estado em localStorage.", error);
  }
}

function restoreStateFromCookies() {
  try {
    let snapshot = null;
    const localSnapshot = localStorage.getItem(LOCAL_STORAGE_STATE_KEY);

    if (localSnapshot) {
      snapshot = JSON.parse(localSnapshot);
    } else {
      // Migração simples do formato antigo em cookies para localStorage.
      const chunkCount = Number.parseInt(getCookie(COOKIE_COUNT_KEY), 10);
      if (!Number.isInteger(chunkCount) || chunkCount <= 0) {
        return;
      }

      let encoded = "";
      for (let i = 0; i < chunkCount; i += 1) {
        const chunk = getCookie(`${COOKIE_CHUNK_KEY}${i}`);
        if (!chunk) {
          return;
        }
        encoded += chunk;
      }

      snapshot = JSON.parse(decodeURIComponent(encoded));
      localStorage.setItem(LOCAL_STORAGE_STATE_KEY, JSON.stringify(snapshot));
      clearStateCookies();
    }

    if (!snapshot || typeof snapshot !== "object") {
      return;
    }

    if (typeof snapshot.title === "string") {
      titleInput.value = snapshot.title;
    }

    const theme = snapshot.theme || {};
    const isHexColor = (value) => /^#[0-9a-fA-F]{6}$/.test(String(value || ""));
    if (isHexColor(theme.primary)) {
      primaryColorInput.value = theme.primary;
    }
    if (isHexColor(theme.dark)) {
      darkColorInput.value = theme.dark;
    }
    if (isHexColor(theme.light)) {
      lightColorInput.value = theme.light;
    }

    if (Array.isArray(snapshot.events)) {
      state.events = snapshot.events.map(toSafeEvent).filter((event) => {
        return event.date || event.time || event.name || event.producer || event.ticket;
      });
    }

    if (snapshot.themeMode === "dark" || snapshot.themeMode === "light") {
      applyTheme(snapshot.themeMode);
    }

    normalizeAgendaTitleInput();
  } catch (error) {
    console.warn("Não foi possível restaurar o estado salvo.", error);
  }
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

const EDITABLE_FIELD_KEYS = ["date", "time", "name", "producer", "ticket"];

function startCellEdit(td, eventIndex, field) {
  if (td.querySelector("input")) {
    return;
  }

  const original = state.events[eventIndex][field];
  const input = document.createElement("input");
  input.type = "text";
  input.value = original;
  input.className = "cell-edit-input";
  input.setAttribute("aria-label", `Editar campo`);

  td.textContent = "";
  td.appendChild(input);
  input.focus();
  input.select();

  let committed = false;

  const commit = () => {
    if (committed) return;
    committed = true;
    const newValue = input.value.trim();
    state.events[eventIndex][field] = newValue !== "" ? newValue : original;
    td.textContent = state.events[eventIndex][field];
    renderPreview();
    persistStateToCookies();
  };

  const cancel = () => {
    if (committed) return;
    committed = true;
    td.textContent = original;
  };

  input.addEventListener("blur", commit);

  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      input.blur();
    } else if (e.key === "Escape") {
      input.removeEventListener("blur", commit);
      cancel();
    }
  });

  input.addEventListener("dblclick", (e) => {
    e.stopPropagation();
  });
}

function normalizeHeader(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function expandSemicolonRows(rows) {
  if (!rows.length || rows[0].length !== 1) {
    return rows;
  }

  if (!String(rows[0][0]).includes(";")) {
    return rows;
  }

  return rows.map((row) => String(row[0] || "").split(";").map((cell) => cell.trim()));
}

function toEventFromRow(row) {
  return {
    date: String(row[0] || "").trim(),
    time: String(row[1] || "").trim(),
    name: String(row[2] || "").trim(),
    producer: String(row[3] || "").trim(),
    ticket: String(row[4] || "").trim(),
  };
}

function isBlankEvent(event) {
  return !event.date && !event.time && !event.name && !event.producer && !event.ticket;
}

function isCompleteEvent(event) {
  return Boolean(event.date && event.time && event.name && event.producer && event.ticket);
}

function normalizeSortableText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLocaleLowerCase("pt-BR");
}

function parseSortableDayMonth(value) {
  const match = String(value || "").trim().match(/^(\d{1,2})\/(\d{1,2})$/);
  if (!match) {
    return null;
  }

  const day = Number.parseInt(match[1], 10);
  const month = Number.parseInt(match[2], 10);
  if (Number.isNaN(day) || Number.isNaN(month) || day < 1 || day > 31 || month < 1 || month > 12) {
    return null;
  }

  return month * 100 + day;
}

function parseSortableTime(value) {
  const match = String(value || "").trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!match) {
    return null;
  }

  const hour = Number.parseInt(match[1], 10);
  const minute = Number.parseInt(match[2], 10);
  if (Number.isNaN(hour) || Number.isNaN(minute) || hour < 0 || hour > 23 || minute < 0 || minute > 59) {
    return null;
  }

  return hour * 60 + minute;
}

function compareTextAsc(a, b) {
  return normalizeSortableText(a).localeCompare(normalizeSortableText(b), "pt-BR", {
    sensitivity: "base",
    numeric: true,
  });
}

function compareNumberLikeAsc(a, b, parser) {
  const parsedA = parser(a);
  const parsedB = parser(b);

  if (parsedA !== null && parsedB !== null) {
    return parsedA - parsedB;
  }

  if (parsedA !== null) {
    return -1;
  }

  if (parsedB !== null) {
    return 1;
  }

  return compareTextAsc(a, b);
}

function sortEventsAscending(sortField) {
  const comparators = {
    date: (a, b) => compareNumberLikeAsc(a.date, b.date, parseSortableDayMonth),
    time: (a, b) => compareNumberLikeAsc(a.time, b.time, parseSortableTime),
    name: (a, b) => compareTextAsc(a.name, b.name),
    producer: (a, b) => compareTextAsc(a.producer, b.producer),
    ticket: (a, b) => compareTextAsc(a.ticket, b.ticket),
  };

  const comparator = comparators[sortField];
  if (!comparator || state.events.length <= 1) {
    return;
  }

  state.events.sort(comparator);
  state.previewPage = 0;
  renderTable();
  renderPreview();
  persistStateToCookies();
}

function parseEventsFromRows(rawRows) {
  const rows = expandSemicolonRows(rawRows).filter((row) => Array.isArray(row) && row.some((cell) => String(cell || "").trim() !== ""));
  if (!rows.length) {
    return { events: [], ignored: 0 };
  }

  let startIndex = 0;
  const firstHeader = rows[0].slice(0, 5).map(normalizeHeader);
  const expectedHeader = REQUIRED_COLUMNS.map(normalizeHeader);
  if (firstHeader.every((value, index) => value === expectedHeader[index])) {
    startIndex = 1;
  }

  const events = [];
  let ignored = 0;

  for (let i = startIndex; i < rows.length; i += 1) {
    const event = toEventFromRow(rows[i]);

    if (isBlankEvent(event)) {
      continue;
    }

    if (isCompleteEvent(event)) {
      events.push(event);
    } else {
      ignored += 1;
    }
  }

  return { events, ignored };
}

function setImportStatus(message, isError = false) {
  importStatus.textContent = message;
  importStatus.style.color = isError ? "#9c2f2f" : "#4f4e74";
}

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[data-src="${src}"]`);
    if (existing) {
      if (existing.dataset.loaded === "true") {
        resolve();
      } else {
        existing.addEventListener("load", () => resolve(), { once: true });
        existing.addEventListener("error", () => reject(new Error(`Falha ao carregar ${src}`)), { once: true });
      }
      return;
    }

    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.dataset.src = src;
    script.addEventListener("load", () => {
      script.dataset.loaded = "true";
      resolve();
    }, { once: true });
    script.addEventListener("error", () => reject(new Error(`Falha ao carregar ${src}`)), { once: true });
    document.head.appendChild(script);
  });
}

async function loadLibraryWithFallback(urls, isReady) {
  if (isReady()) {
    return;
  }

  let lastError = null;
  for (const url of urls) {
    try {
      await loadScript(url);
      if (isReady()) {
        return;
      }
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("Biblioteca indisponível");
}

function ensureXlsxLoaded() {
  if (!xlsxLoadPromise) {
    xlsxLoadPromise = loadLibraryWithFallback(XLSX_LIB_URLS, () => typeof window.XLSX !== "undefined");
  }

  return xlsxLoadPromise;
}

function ensureJsZipLoaded() {
  if (!jszipLoadPromise) {
    jszipLoadPromise = loadLibraryWithFallback(JSZIP_LIB_URLS, () => typeof window.JSZip !== "undefined");
  }

  return jszipLoadPromise;
}

function exportTemplateWorkbook() {
  if (typeof window.XLSX === "undefined") {
    throw new Error("Biblioteca XLSX indisponível.");
  }

  const workbook = XLSX.utils.book_new();

  const eventsSheetData = [
    REQUIRED_COLUMNS,
    ["02/04", "19:30", "Espetáculo A Menina que Gostava de Chuva", "Secretaria de Educação de Itajaí", "Gratuito"],
    ["11/04", "21:00", "Musical Broadway Night's", "Mantovani Promoções", "Blueticket"],
    ["23/04", "20:00", "Concerto Rock ao Piano - Especial Pink Floyd", "Bruno Hrabovsky", "Guichê Web"],
  ];
  const eventsSheet = XLSX.utils.aoa_to_sheet(eventsSheetData);
  eventsSheet["!cols"] = [{ wch: 12 }, { wch: 10 }, { wch: 45 }, { wch: 32 }, { wch: 22 }];

  XLSX.utils.book_append_sheet(workbook, eventsSheet, "Eventos");
  XLSX.writeFile(workbook, "modelo-importacao-agenda.xlsx", { compression: true });
}

async function importSheetFile() {
  const [file] = sheetFileInput.files;

  if (!file) {
    setImportStatus("Selecione um arquivo CSV, XLS ou XLSX para importar.", true);
    return;
  }

  try {
    await ensureXlsxLoaded();
    setImportStatus("Importando planilha...");
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, {
      type: "array",
      raw: false,
      cellText: true,
      codepage: 65001,
    });

    const preferredSheet = workbook.SheetNames.find((name) => normalizeHeader(name) === "eventos");
    const firstSheetName = preferredSheet || workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
      blankrows: false,
      raw: false,
    });

    const { events, ignored } = parseEventsFromRows(rows);
    if (!events.length) {
      setImportStatus("Nenhum evento válido encontrado. Verifique se as colunas A-E foram preenchidas corretamente.", true);
      return;
    }

    state.events = [...state.events, ...events];
    renderTable();
    renderPreview();
    persistStateToCookies();

    const ignoredText = ignored ? ` ${ignored} linha(s) incompleta(s) foram ignoradas.` : "";
    setImportStatus(`${events.length} evento(s) importado(s) com sucesso.${ignoredText}`);
    sheetFileInput.value = "";
  } catch (error) {
    console.error(error);
    setImportStatus("Não foi possível importar a planilha. Confira o formato do arquivo.", true);
  }
}

function renderTable(focusRowIndex = null) {
  eventsTableBody.innerHTML = "";

  const clearDropIndicators = () => {
    eventsTableBody.querySelectorAll("tr").forEach((tableRow) => {
      tableRow.classList.remove("drop-before", "drop-after", "is-dragging");
    });
  };

  const moveEvent = (fromIndex, toIndex) => {
    if (fromIndex === toIndex || fromIndex < 0 || toIndex < 0) {
      return;
    }

    const [movedEvent] = state.events.splice(fromIndex, 1);
    state.events.splice(toIndex, 0, movedEvent);
  };

  const commitReorder = (fromIndex, toIndex, rowToFocus = null) => {
    moveEvent(fromIndex, toIndex);
    draggedRowIndex = null;
    clearDropIndicators();
    renderTable(rowToFocus);
    renderPreview();
    persistStateToCookies();
  };

  state.events.forEach((event, index) => {
    const row = document.createElement("tr");
    row.className = "draggable-row";
    row.draggable = true;
    row.tabIndex = 0;
    row.dataset.index = String(index);
    row.setAttribute("aria-label", `Evento ${index + 1}. Use Alt + seta para cima ou para baixo para reorganizar.`);
    row.innerHTML = `
      <td>${escapeHtml(event.date)}</td>
      <td>${escapeHtml(event.time)}</td>
      <td>${escapeHtml(event.name)}</td>
      <td>${escapeHtml(event.producer)}</td>
      <td>${escapeHtml(event.ticket)}</td>
      <td>
        <div class="row-actions">
          <button type="button" class="ghost move-btn move-up" aria-label="Mover evento para cima" ${index === 0 ? "disabled" : ""}>
            <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
              <path d="M12 19V7"></path>
              <path d="M7.5 11.5L12 7l4.5 4.5"></path>
            </svg>
          </button>
          <button type="button" class="ghost move-btn move-down" aria-label="Mover evento para baixo" ${index === state.events.length - 1 ? "disabled" : ""}>
            <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
              <path d="M12 5v12"></path>
              <path d="M7.5 12.5L12 17l4.5-4.5"></path>
            </svg>
          </button>
          <button type="button" class="ghost remove-btn" aria-label="Remover evento">Remover</button>
        </div>
      </td>
    `;

    const removeButton = row.querySelector(".remove-btn");
    const moveUpButton = row.querySelector(".move-up");
    const moveDownButton = row.querySelector(".move-down");
    row.querySelectorAll("button").forEach((button) => {
      button.draggable = false;
    });

    const editableCells = Array.from(row.querySelectorAll("td")).slice(0, 5);
    editableCells.forEach((td, colIndex) => {
      td.classList.add("editable-cell");
      td.setAttribute("title", "Clique duas vezes para editar");
      td.addEventListener("dblclick", (e) => {
        e.stopPropagation();
        startCellEdit(td, index, EDITABLE_FIELD_KEYS[colIndex]);
      });
    });

    row.addEventListener("dragstart", (dragEvent) => {
      if (dragEvent.target instanceof Element && dragEvent.target.closest("button")) {
        dragEvent.preventDefault();
        return;
      }

      draggedRowIndex = index;
      clearDropIndicators();
      row.classList.add("is-dragging");

      if (dragEvent.dataTransfer) {
        dragEvent.dataTransfer.effectAllowed = "move";
        dragEvent.dataTransfer.setData("text/plain", String(index));
      }
    });

    row.addEventListener("dragover", (dragEvent) => {
      if (draggedRowIndex === null || draggedRowIndex === index) {
        return;
      }

      dragEvent.preventDefault();
      if (dragEvent.dataTransfer) {
        dragEvent.dataTransfer.dropEffect = "move";
      }

      clearDropIndicators();
      const rowRect = row.getBoundingClientRect();
      const shouldDropBefore = dragEvent.clientY < rowRect.top + rowRect.height / 2;
      row.classList.add(shouldDropBefore ? "drop-before" : "drop-after");
    });

    row.addEventListener("drop", (dragEvent) => {
      dragEvent.preventDefault();

      if (draggedRowIndex === null || draggedRowIndex === index) {
        clearDropIndicators();
        return;
      }

      const rowRect = row.getBoundingClientRect();
      const shouldDropBefore = dragEvent.clientY < rowRect.top + rowRect.height / 2;
      const rawTargetIndex = index + (shouldDropBefore ? 0 : 1);
      const adjustedTargetIndex = draggedRowIndex < rawTargetIndex ? rawTargetIndex - 1 : rawTargetIndex;

      commitReorder(draggedRowIndex, adjustedTargetIndex, adjustedTargetIndex);
    });

    row.addEventListener("dragend", () => {
      draggedRowIndex = null;
      clearDropIndicators();
    });

    row.addEventListener("keydown", (keyboardEvent) => {
      if (keyboardEvent.target !== row || !keyboardEvent.altKey) {
        return;
      }

      let targetIndex = null;
      if (keyboardEvent.key === "ArrowUp") {
        targetIndex = Math.max(0, index - 1);
      } else if (keyboardEvent.key === "ArrowDown") {
        targetIndex = Math.min(state.events.length - 1, index + 1);
      }

      if (targetIndex === null || targetIndex === index) {
        return;
      }

      keyboardEvent.preventDefault();
      commitReorder(index, targetIndex, targetIndex);
    });

    removeButton.addEventListener("click", () => {
      state.events.splice(index, 1);
      renderTable();
      renderPreview();
      persistStateToCookies();
    });

    moveUpButton.addEventListener("click", () => {
      const targetIndex = Math.max(0, index - 1);
      if (targetIndex === index) {
        return;
      }
      commitReorder(index, targetIndex, targetIndex);
    });

    moveDownButton.addEventListener("click", () => {
      const targetIndex = Math.min(state.events.length - 1, index + 1);
      if (targetIndex === index) {
        return;
      }
      commitReorder(index, targetIndex, targetIndex);
    });

    eventsTableBody.appendChild(row);
  });

  if (Number.isInteger(focusRowIndex)) {
    const rowToFocus = eventsTableBody.querySelector(`tr[data-index="${focusRowIndex}"]`);
    if (rowToFocus instanceof HTMLElement) {
      rowToFocus.focus();
    }
  }

  updateInterfaceState();
}

function updateInterfaceState() {
  const total = state.events.length;
  eventCount.textContent = `${total} evento${total === 1 ? "" : "s"}`;
  emptyState.hidden = total > 0;
  sortEventsButton.disabled = total === 0;
  exportButton.disabled = total === 0;
  exportCurrentButton.disabled = total === 0;
}

function downloadPageImage(pageEvents, pageIndex, totalPages) {
  const outputCanvas = document.createElement("canvas");
  outputCanvas.width = OUTPUT_WIDTH;
  outputCanvas.height = OUTPUT_HEIGHT;
  const outputCtx = outputCanvas.getContext("2d");

  drawSchedule(outputCtx, OUTPUT_WIDTH, OUTPUT_HEIGHT, pageEvents);

  const link = document.createElement("a");
  const pageSuffix = totalPages > 1 ? `-p${String(pageIndex + 1).padStart(2, "0")}` : "";
  link.download = `agenda-1080x1440${pageSuffix}.png`;
  link.href = outputCanvas.toDataURL("image/png");
  link.click();
}

function getPagePngDataUrl(pageEvents) {
  const outputCanvas = document.createElement("canvas");
  outputCanvas.width = OUTPUT_WIDTH;
  outputCanvas.height = OUTPUT_HEIGHT;
  const outputCtx = outputCanvas.getContext("2d");

  drawSchedule(outputCtx, OUTPUT_WIDTH, OUTPUT_HEIGHT, pageEvents);
  return outputCanvas.toDataURL("image/png");
}

function dataUrlToUint8Array(dataUrl) {
  const base64 = dataUrl.split(",")[1] || "";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

async function downloadAllPagesAsZip(pages) {
  if (typeof window.JSZip === "undefined") {
    throw new Error("Biblioteca JSZip indisponível.");
  }

  const zip = new JSZip();
  pages.forEach((pageEvents, index) => {
    const pageSuffix = String(index + 1).padStart(2, "0");
    const filename = `agenda-1080x1440-p${pageSuffix}.png`;
    const pngDataUrl = getPagePngDataUrl(pageEvents);
    zip.file(filename, dataUrlToUint8Array(pngDataUrl));
  });

  const zipBlob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(zipBlob);
  link.download = "agenda-1080x1440.zip";
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 4000);
}

function wrapText(ctx, text, maxWidth) {
  const words = text.trim().split(/\s+/);
  const lines = [];
  let currentLine = "";

  function breakLongWord(word) {
    const chunks = [];
    let chunk = "";

    for (const char of word) {
      const next = chunk + char;
      if (ctx.measureText(next).width <= maxWidth) {
        chunk = next;
      } else {
        if (chunk) {
          chunks.push(chunk);
        }
        chunk = char;
      }
    }

    if (chunk) {
      chunks.push(chunk);
    }

    return chunks;
  }

  for (const word of words) {
    if (ctx.measureText(word).width > maxWidth) {
      if (currentLine) {
        lines.push(currentLine);
        currentLine = "";
      }

      const chunks = breakLongWord(word);
      for (let i = 0; i < chunks.length; i += 1) {
        const chunk = chunks[i];
        if (i < chunks.length - 1) {
          lines.push(chunk);
        } else {
          currentLine = chunk;
        }
      }
      continue;
    }

    const testLine = currentLine ? `${currentLine} ${word}` : word;
    if (ctx.measureText(testLine).width <= maxWidth) {
      currentLine = testLine;
    } else {
      if (currentLine) {
        lines.push(currentLine);
      }
      currentLine = word;
    }
  }

  if (currentLine) {
    lines.push(currentLine);
  }

  return lines.length ? lines : [""];
}

function drawCenteredLineItems(ctx, items, centerX, centerY) {
  if (!items.length) {
    return;
  }

  const totalSpan = items.slice(1).reduce((sum, item) => sum + item.lineHeight, 0);
  let currentY = centerY - totalSpan / 2;

  items.forEach((item, index) => {
    if (index > 0) {
      currentY += item.lineHeight;
    }
    ctx.font = item.font;
    ctx.fillText(item.text, centerX, currentY);
  });
}

function getPageCapacity(height = OUTPUT_HEIGHT) {
  const brandTop = 24;
  const brandHeight = 146;
  const titleHeight = 76;
  const sectionGap = 16;
  const headerHeight = 52;
  const rowHeight = 89;
  const titleTop = brandTop + brandHeight + sectionGap;
  const headerTop = titleTop + titleHeight + sectionGap;
  const availableHeight = height - (headerTop + headerHeight) - 24;
  return Math.max(1, Math.floor(availableHeight / rowHeight));
}

function getPagedEvents() {
  const capacity = getPageCapacity(OUTPUT_HEIGHT);

  if (state.events.length === 0) {
    return [[]];
  }

  const pages = [];
  for (let i = 0; i < state.events.length; i += capacity) {
    pages.push(state.events.slice(i, i + capacity));
  }

  return pages;
}

function updatePageControls(pages) {
  const totalPages = pages.length;
  state.previewPage = Math.min(state.previewPage, totalPages - 1);
  state.previewPage = Math.max(state.previewPage, 0);

  pageInfo.textContent = `Página ${state.previewPage + 1} de ${totalPages}`;
  prevPageButton.disabled = state.previewPage === 0;
  nextPageButton.disabled = state.previewPage >= totalPages - 1;
}

function drawSchedule(ctx, width, height, events) {
  const theme = getTheme();
  const title = titleInput.value.trim() || "AGENDA";

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);

  const margin = 28;
  const contentX = margin;
  const contentWidth = width - margin * 2;

  const brandTop = 24;
  const brandHeight = 146;
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(contentX, brandTop, contentWidth, brandHeight);

  const leftAreaX = contentX + 18;
  const leftAreaWidth = contentWidth * 0.44;
  const rightAreaWidth = contentWidth * 0.32;
  const rightAreaX = contentX + contentWidth - rightAreaWidth - 18;
  const centerDividerX = contentX + contentWidth * 0.52;

  ctx.strokeStyle = "rgba(127, 68, 47, 0.18)";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(centerDividerX, brandTop + 18);
  ctx.lineTo(centerDividerX, brandTop + brandHeight - 18);
  ctx.stroke();

  if (logoTeatro.complete && logoTeatro.naturalWidth > 0) {
    drawImageContain(ctx, logoTeatro, leftAreaX, brandTop + 10, leftAreaWidth, brandHeight - 20);
  } else {
    ctx.fillStyle = "#111";
    ctx.textAlign = "left";
    ctx.font = "700 42px Montserrat";
    ctx.fillText("TEATRO", contentX + 22, 70);
    ctx.fillText("MUNICIPAL", contentX + 22, 112);
    ctx.fillText("DE ITAJAÍ", contentX + 22, 154);
  }

  if (logoFundacao.complete && logoFundacao.naturalWidth > 0) {
    drawImageContain(ctx, logoFundacao, rightAreaX, brandTop + 14, rightAreaWidth, brandHeight - 28);
  } else {
    ctx.fillStyle = "#7f6d5a";
    ctx.textAlign = "left";
    ctx.font = "700 30px Montserrat";
    ctx.fillText("Fundação Cultural", rightAreaX + 6, 106);
    ctx.fillText("de Itajaí", rightAreaX + 6, 144);
  }

  const titleTop = brandTop + brandHeight + 16;
  const titleHeight = 76;
  ctx.fillStyle = theme.primary;
  ctx.fillRect(contentX, titleTop, contentWidth, titleHeight);

  ctx.fillStyle = "#fff";
  ctx.textAlign = "center";
  ctx.font = "800 50px Montserrat";
  ctx.fillText(title.toUpperCase(), contentX + contentWidth / 2, titleTop + 52);

  const headerTop = titleTop + titleHeight + 16;
  const tableRightW = Math.floor(contentWidth * 0.23);
  const tableLeftW = Math.floor(contentWidth * 0.14);
  const tableMiddleW = contentWidth - tableLeftW - tableRightW;
  const rowHeight = 89;
  const rowCount = events.length;
  const tableBodyHeight = rowHeight * rowCount;

  ctx.textAlign = "left";
  ctx.fillStyle = theme.dark;
  ctx.fillRect(contentX, headerTop, tableLeftW, 52);
  ctx.fillRect(contentX + tableLeftW, headerTop, tableMiddleW, 52);
  ctx.fillRect(contentX + tableLeftW + tableMiddleW, headerTop, tableRightW, 52);

  ctx.strokeStyle = "#ffffff";
  ctx.lineWidth = 4;
  ctx.beginPath();
  ctx.moveTo(contentX + tableLeftW, headerTop);
  ctx.lineTo(contentX + tableLeftW, headerTop + 52 + tableBodyHeight);
  ctx.moveTo(contentX + tableLeftW + tableMiddleW, headerTop);
  ctx.lineTo(contentX + tableLeftW + tableMiddleW, headerTop + 52 + tableBodyHeight);
  ctx.stroke();

  const bodyFontSize = 24;
  const headerFontSize = bodyFontSize + 2;

  ctx.fillStyle = "#fff";
  ctx.font = `700 ${headerFontSize}px Montserrat`;
  ctx.textAlign = "center";
  ctx.fillText("Data", contentX + tableLeftW / 2, headerTop + 34);
  ctx.fillText("Evento", contentX + tableLeftW + tableMiddleW / 2, headerTop + 34);
  ctx.fillText("Onde comprar", contentX + tableLeftW + tableMiddleW + tableRightW / 2, headerTop + 34);

  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  for (let i = 0; i < rowCount; i += 1) {
    const y = headerTop + 52 + i * rowHeight;
    const rowCenterY = y + rowHeight / 2;
    const event = events[i];
    ctx.fillStyle = i % 2 === 0 ? theme.primary : theme.light;
    ctx.fillRect(contentX, y, contentWidth, rowHeight);

    ctx.strokeStyle = "#fff";
    ctx.lineWidth = 4;
    ctx.beginPath();
    ctx.moveTo(contentX, y);
    ctx.lineTo(contentX + contentWidth, y);
    ctx.stroke();

    ctx.fillStyle = "#fff";

    const leftItems = [event.date, event.time]
      .map((text) => (text || "").trim())
      .filter(Boolean)
      .map((text) => ({
        text,
        font: `700 ${bodyFontSize}px Montserrat`,
        lineHeight: 30,
      }));
    drawCenteredLineItems(ctx, leftItems, contentX + tableLeftW / 2, rowCenterY);

    const eventCenterX = contentX + tableLeftW + tableMiddleW / 2;
    const nameMaxWidth = tableMiddleW - 22;
    const producerMaxWidth = tableMiddleW - 22;

    ctx.font = `700 ${bodyFontSize}px Montserrat`;
    const nameLines = wrapText(ctx, event.name, nameMaxWidth).slice(0, 2).filter(Boolean);

    ctx.font = `500 ${bodyFontSize}px Montserrat`;
    const producerLines = wrapText(ctx, event.producer, producerMaxWidth).slice(0, 1).filter(Boolean);

    const middleItems = [
      ...nameLines.map((text) => ({
        text,
        font: `700 ${bodyFontSize}px Montserrat`,
        lineHeight: 30,
      })),
      ...producerLines.map((text) => ({
        text,
        font: `500 ${bodyFontSize}px Montserrat`,
        lineHeight: 28,
      })),
    ];
    drawCenteredLineItems(ctx, middleItems, eventCenterX, rowCenterY);

    ctx.font = `700 ${bodyFontSize}px Montserrat`;
    const ticketLines = wrapText(ctx, event.ticket, tableRightW - 26).slice(0, 2).filter(Boolean);
    const ticketItems = ticketLines.map((text) => ({
      text,
      font: `700 ${bodyFontSize}px Montserrat`,
      lineHeight: 26,
    }));
    drawCenteredLineItems(ctx, ticketItems, contentX + tableLeftW + tableMiddleW + tableRightW / 2, rowCenterY);
  }

  ctx.textBaseline = "alphabetic";

}

function renderPreview() {
  const pages = getPagedEvents();
  updatePageControls(pages);
  drawSchedule(previewCtx, previewCanvas.width, previewCanvas.height, pages[state.previewPage]);
}

function addEvent(eventData) {
  state.events.push(eventData);
  renderTable();
  renderPreview();
  persistStateToCookies();
}

eventForm.addEventListener("submit", (event) => {
  event.preventDefault();

  addEvent({
    date: fields.date.value.trim(),
    time: fields.time.value.trim(),
    name: fields.name.value.trim(),
    producer: fields.producer.value.trim(),
    ticket: fields.ticket.value.trim(),
  });

  eventForm.reset();
  fields.date.focus();
});

clearFormButton.addEventListener("click", () => {
  eventForm.reset();
  fields.date.focus();
});

downloadTemplateButton.addEventListener("click", async () => {
  try {
    await ensureXlsxLoaded();
    exportTemplateWorkbook();
    setImportStatus("Modelo XLSX baixado com sucesso.");
  } catch (error) {
    console.error(error);
    setImportStatus("Não foi possível carregar a biblioteca de planilhas para gerar o modelo.", true);
  }
});

importSheetButton.addEventListener("click", () => {
  importSheetFile();
});

sortEventsButton.addEventListener("click", () => {
  sortEventsAscending(sortFieldSelect.value);
});

resetColorsButton.addEventListener("click", () => {
  resetThemeToDefault();
  renderPreview();
  persistStateToCookies();
});

loadSampleButton.addEventListener("click", () => {
  state.events = [
    { date: "02/04", time: "19:30", name: "Espetáculo A Menina que Gostava de Chuva", producer: "Secretaria de Educação de Itajaí", ticket: "Gratuito" },
    { date: "09/04", time: "*", name: "9º Encontro Internacional das Etnias", producer: "Grupo Folclórico Tropeiros do Litoral", ticket: "*" },
    { date: "11/04", time: "16:00", name: "Musical O Rei Leão", producer: "Mantovani Promoções", ticket: "Blueticket" },
    { date: "11/04", time: "21:00", name: "Musical Broadway Night's", producer: "Mantovani Promoções", ticket: "Blueticket" },
    { date: "12/04", time: "19:30", name: "Stand-up Dete Pexera - VIAJADA!", producer: "Rizzih Arte e Entretenimento", ticket: "Sympla" },
    { date: "16/04", time: "21:00", name: "Concerto Starlight - Clássicos do Cinema", producer: "Starlight Concert", ticket: "Ingresso Digital" },
    { date: "18/04", time: "15:00", name: "Espetáculo K-Pop As Guerreiras O Show - Tributo", producer: "Canesso Produções", ticket: "Meaple" },
    { date: "19/04", time: "19:00", name: "Stand-up Cris Pereira Show 5.0", producer: "Art in Palco Produções e Eventos", ticket: "Minha Entrada" },
    { date: "23/04", time: "20:00", name: "Concerto Rock ao Piano - Especial Pink Floyd", producer: "Bruno Hrabovsky", ticket: "Guichê Web" },
    { date: "25/04", time: "18:00", name: "Musical O Mágico de Oz Em Forma de Cordel", producer: "Espaço Cena", ticket: "TicketCenter" },
    { date: "29/04", time: "19:30", name: "Festival de Dança", producer: "Fundação Cultural de Itajaí", ticket: "Gratuito" },
    { date: "30/04", time: "19:30", name: "Festival de Dança", producer: "Fundação Cultural de Itajaí", ticket: "Gratuito" },
  ];

  state.previewPage = 0;
  renderTable();
  renderPreview();
  persistStateToCookies();
});

prevPageButton.addEventListener("click", () => {
  state.previewPage = Math.max(0, state.previewPage - 1);
  renderPreview();
});

nextPageButton.addEventListener("click", () => {
  const pages = getPagedEvents();
  state.previewPage = Math.min(pages.length - 1, state.previewPage + 1);
  renderPreview();
});

[primaryColorInput, darkColorInput, lightColorInput].forEach((input) => {
  input.addEventListener("input", () => {
    renderPreview();
    persistStateToCookies();
  });
});

titleInput.addEventListener("input", () => {
  normalizeAgendaTitleInput();
  renderPreview();
  persistStateToCookies();
});

themeToggleButton.addEventListener("click", () => {
  const nextTheme = getCurrentTheme() === "dark" ? "light" : "dark";
  applyTheme(nextTheme);
  persistThemeToCookie();
  persistStateToCookies();
});

exportButton.addEventListener("click", async () => {
  const pages = getPagedEvents();

  if (pages.length <= 1) {
    downloadPageImage(pages[0], 0, 1);
    return;
  }

  const originalText = exportButton.textContent;
  exportButton.disabled = true;
  exportButton.textContent = "Preparando arquivo...";

  try {
    await ensureJsZipLoaded();
    await downloadAllPagesAsZip(pages);
  } catch (error) {
    console.warn("Falha ao gerar ZIP. Fazendo download individual das páginas.", error);
    pages.forEach((pageEvents, index) => {
      downloadPageImage(pageEvents, index, pages.length);
    });
  } finally {
    exportButton.disabled = false;
    exportButton.textContent = originalText;
  }
});

exportCurrentButton.addEventListener("click", () => {
  const pages = getPagedEvents();
  downloadPageImage(pages[state.previewPage], state.previewPage, pages.length);
});

loadLogos();
restoreThemeFromCookie();
restoreStateFromCookies();
normalizeAgendaTitleInput();
renderTable();
renderPreview();
