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
const prevPageButton = document.getElementById("prevPage");
const nextPageButton = document.getElementById("nextPage");
const pageInfo = document.getElementById("pageInfo");

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
const COOKIE_COUNT_KEY = `${COOKIE_PREFIX}chunks`;
const COOKIE_CHUNK_KEY = `${COOKIE_PREFIX}data_`;
const COOKIE_CHUNK_SIZE = 3500;
const COOKIE_MAX_AGE_DAYS = 180;

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
      events: state.events.map(toSafeEvent),
    };

    const encoded = encodeURIComponent(JSON.stringify(snapshot));
    const chunks = encoded.match(new RegExp(`.{1,${COOKIE_CHUNK_SIZE}}`, "g")) || [""];

    clearStateCookies();
    setCookie(COOKIE_COUNT_KEY, String(chunks.length));
    chunks.forEach((chunk, index) => {
      setCookie(`${COOKIE_CHUNK_KEY}${index}`, chunk);
    });
  } catch (error) {
    console.warn("Não foi possível salvar o estado em cookies.", error);
  }
}

function restoreStateFromCookies() {
  try {
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

    const snapshot = JSON.parse(decodeURIComponent(encoded));
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

    normalizeAgendaTitleInput();
  } catch (error) {
    console.warn("Não foi possível restaurar o estado salvo em cookies.", error);
  }
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
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

function exportTemplateWorkbook() {
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

function renderTable() {
  eventsTableBody.innerHTML = "";

  state.events.forEach((event, index) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${escapeHtml(event.date)}</td>
      <td>${escapeHtml(event.time)}</td>
      <td>${escapeHtml(event.name)}</td>
      <td>${escapeHtml(event.producer)}</td>
      <td>${escapeHtml(event.ticket)}</td>
      <td><button data-index="${index}" class="ghost remove-btn">Remover</button></td>
    `;

    row.querySelector("button").addEventListener("click", () => {
      state.events.splice(index, 1);
      renderTable();
      renderPreview();
      persistStateToCookies();
    });

    eventsTableBody.appendChild(row);
  });

  updateInterfaceState();
}

function updateInterfaceState() {
  const total = state.events.length;
  eventCount.textContent = `${total} evento${total === 1 ? "" : "s"}`;
  emptyState.hidden = total > 0;
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

downloadTemplateButton.addEventListener("click", () => {
  exportTemplateWorkbook();
  setImportStatus("Modelo XLSX baixado com sucesso.");
});

importSheetButton.addEventListener("click", () => {
  importSheetFile();
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

exportButton.addEventListener("click", () => {
  const pages = getPagedEvents();

  pages.forEach((pageEvents, index) => {
    downloadPageImage(pageEvents, index, pages.length);
  });
});

exportCurrentButton.addEventListener("click", () => {
  const pages = getPagedEvents();
  downloadPageImage(pages[state.previewPage], state.previewPage, pages.length);
});

loadLogos();
restoreStateFromCookies();
normalizeAgendaTitleInput();
renderTable();
renderPreview();
