/****************************************************************/
/******************      CONFIGURAZIONE      ******************/
/****************************************************************/

// IMPORTANTE: Incolla qui l'URL della tua Web App di Google Apps Script
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw_RClf2p3aQ2A2YkXJ2sYwYJ-p-gC4g5f6hJkK-eP7tN8X9bQ/exec";

/****************************************************************/
/******************      LOGICA PRINCIPALE      ******************/
/****************************************************************/

// Funzione centrale per chiamare il backend
async function callAppsScript(functionName, payload = {}) {
  const response = await fetch(SCRIPT_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'text/plain;charset=utf-8', // Necessario per Google Apps Script
    },
    body: JSON.stringify({ action: functionName, data: payload }),
    mode: 'cors', // Abilita CORS
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Errore di rete: ${response.status} ${response.statusText} - ${errorText}`);
  }

  const result = await response.json();
  if (result.error) {
    throw new Error(result.error);
  }
  return result.data;
}

// Funzioni di utilità per codifica e decodifica Base64
function b64Encode(str) {
  return btoa(encodeURIComponent(str).replace(/%([0-9A-F]{2})/g,
    function toSolidBytes(match, p1) {
      return String.fromCharCode('0x' + p1);
    }));
}

function b64Decode(str) {
  return decodeURIComponent(atob(str).split('').map(function (c) {
    return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
  }).join(''));
}

// Funzioni per i pulsanti di azione principali
async function runAction(functionName) {
  document.querySelectorAll('button').forEach(btn => btn.disabled = true);
  if (confirm("Sei sicuro di voler avviare l'azione: " + functionName + "?")) {
    const loader = document.getElementById('loader');
    const resultArea = document.getElementById('result-area');
    loader.style.display = 'block';
    resultArea.style.display = 'block';
    resultArea.innerText = 'Esecuzione in corso...';

    try {
      const result = await callAppsScript(functionName);
      onSuccess(result);
    } catch (error) {
      onFailure(error);
    }
  } else {
    document.querySelectorAll('button').forEach(btn => btn.disabled = false);
  }
}

function onSuccess(result) {
  document.querySelectorAll('button').forEach(btn => btn.disabled = false);
  document.getElementById('loader').style.display = 'none';
  const resultArea = document.getElementById('result-area');
  resultArea.innerText = result;
  resultArea.style.display = 'block';
  resultArea.style.backgroundColor = '#28A745'; // Success color
  resultArea.style.color = 'white';
}

function onFailure(error) {
  document.querySelectorAll('button').forEach(btn => btn.disabled = false);
  document.getElementById('loader').style.display = 'none';
  const resultArea = document.getElementById('result-area');
  resultArea.innerText = 'Si è verificato un errore: ' + error.message;
  resultArea.style.display = 'block';
  resultArea.style.backgroundColor = '#EF4444'; // Error color
  resultArea.style.color = 'white';
}

// Funzioni per la visualizzazione dei fogli
async function mostraFoglio(functionName) {
  document.getElementById('loader').style.display = 'block';
  document.getElementById('foglio-container').style.display = 'none';
  document.getElementById('result-area').style.display = 'none';

  try {
    const data = await callAppsScript(functionName);
    document.getElementById('loader').style.display = 'none';
    const container = document.getElementById('foglio-container');
    container.innerHTML = ''; // Clear previous content

    let buttonsHtml = '';
    if (functionName === 'eseguiMostraFoglio') {
      buttonsHtml = `<div class="inline-action-buttons"><button class="management-button" onclick="runAction('eseguiAzioneScarica')">📥 Scarica Nuove Notizie</button><button class="management-button" onclick="runAction('eseguiAzioneElabora')">✨ Elabora Selezionate</button><div class="dropdown-container"><button id="btn-gestisci-fonti" class="management-button" onclick="toggleDropdown()">📡 Gestisci Fonti</button><div id="fonti-dropdown" class="dropdown-panel"><div id="fonti-loader" style="padding: 12px;">Caricamento...</div><ul id="fonti-list"></ul><div class="dropdown-footer"><span id="fonti-status"></span><button id="fonti-save-button" onclick="saveFontiChanges()">Salva</button></div></div></div></div>`;
      container.innerHTML = buttonsHtml + createTableFoglio1(data);
    } else if (functionName === 'eseguiMostraFoglio2') {
      buttonsHtml = `<div class="inline-action-buttons"><button class="management-button" onclick="runAction('eseguiAzioneImmagini')">🖼️ Scarica Link Immagini</button></div>`;
      container.innerHTML = buttonsHtml + createTableFoglio2(data);
    } else if (functionName === 'eseguiMostraFoglio3') {
      container.innerHTML = createTableFoglio3(data);
    } else if (functionName === 'eseguiMostraFoglioInAttesa') {
      container.innerHTML = '<div class="table-wrapper-in-attesa">' + createTableFoglioInAttesa(data) + '</div>';
    }

    container.style.display = 'block';
  } catch (error) {
    onFailure(error);
  }
}

function createTable(data, renderRow) {
  if (!data || data.length <= 1) {
    const message = (data && data.length > 0 && data[0].length > 0) ? data[0].join(": ") : 'Nessun dato da mostrare.';
    return `<div class="result-area-inline">${message}</div>`;
  }
  let tableHtml = '<table>';
  const headers = data[0];
  tableHtml += '<thead><tr>';
  headers.forEach(header => { tableHtml += `<th>${escapeHtml(header)}</th>`; });
  if (renderRow.name.includes('Foglio2') || renderRow.name.includes('Foglio3')) { tableHtml += '<th>Azione</th>'; }
  tableHtml += '</tr></thead>';
  tableHtml += '<tbody>';
  for (let i = 1; i < data.length; i++) {
    tableHtml += renderRow(data[i], i + 1, headers);
  }
  tableHtml += '</tbody></table>';
  return tableHtml;
}

function createTableFoglio1(data) {
  return createTable(data, (rowData, sheetRowNumber, headers) => {
    let rowHtml = `<tr>`;
    rowData.forEach((cell, index) => {
      const label = headers[index] ? escapeHtml(headers[index]) : '';
      if (index === 3) { // Colonna "Selezionato"
        const isChecked = cell.toString().toLowerCase() === 'sì';
        rowHtml += `<td data-label="${label}"><input type="checkbox" onchange="inviaModifica(this, ${sheetRowNumber})" ${isChecked ? 'checked' : ''}><span id="status-${sheetRowNumber}" class="update-status"></span></td>`;
      } else {
        rowHtml += `<td data-label="${label}">${escapeHtml(cell)}</td>`;
      }
    });
    return rowHtml + '</tr>';
  });
}

function createTableFoglio2(data) {
  return createTable(data, (rowData, sheetRowNumber, headers) => {
    let rowHtml = `<tr>`;
    rowData.forEach((cell, index) => {
      const label = headers[index] ? escapeHtml(headers[index]) : '';
      if (index === 3 || index === 4) { // Colonne editabili
        rowHtml += `<td data-label="${label}"><div id="cell-${sheetRowNumber}-${index}" contenteditable="true" class="editable-cell">${escapeHtml(cell)}</div></td>`;
      } else {
        rowHtml += `<td data-label="${label}">${escapeHtml(cell)}</td>`;
      }
    });
    rowHtml += `<td data-label="Azione"><button class="management-button" onclick="salvaModificheFoglio2(${sheetRowNumber})">Salva</button><span id="status-foglio2-${sheetRowNumber}" class="update-status"></span></td>`;
    return rowHtml + '</tr>';
  });
}

function createTableFoglio3(data) {
  return createTable(data, (rowData, sheetRowNumber, headers) => {
    let rowHtml = `<tr>`;
    const gameTitle = rowData[0];
    const imageUrl = rowData[1];
    const postTitle = rowData[2] || '';
    const postBody = rowData[3] || '';

    rowHtml += `<td data-label="${headers[0]}">${escapeHtml(gameTitle)}</td>`;
    rowHtml += `<td data-label="${headers[1]}"><img src="${escapeHtml(imageUrl)}" style="max-width: 150px; max-height: 100px; object-fit: cover; border-radius: 4px;" loading="lazy"></td>`;

    rowHtml += `<td data-label="Azione"><button class="management-button" 
            data-image-url="${escapeHtml(imageUrl)}"
            data-post-title="${b64Encode(postTitle)}"
            data-post-body="${b64Encode(postBody)}"
            onclick="openImageEditor(this)">Modifica</button></td>`;

    return rowHtml + '</tr>';
  });
}

function createTableFoglioInAttesa(data) {
  const displayHeaders = [data[0][0], data[0][5], "Azione"]; // Titolo, Stato, Azione
  const displayData = [displayHeaders, ...data.slice(1)];

  return createTable(displayData, (rowData, _, headers) => {
    const titolo = rowData[0];
    const testo = rowData[1];
    const imageUrl = rowData[2];
    const dateValue = rowData[3];
    const timeValue = rowData[4];
    const stato = rowData[5];
    const sheetRowNumber = rowData[6];

    let statusClass = '';
    const lowerCaseStato = stato.toLowerCase();
    if (lowerCaseStato.includes('pubblicato')) {
      statusClass = 'status-published';
    } else if (lowerCaseStato.includes('programmato')) {
      statusClass = 'status-scheduled';
    } else if (lowerCaseStato.includes('da programmare')) {
      statusClass = 'status-pending';
    } else if (lowerCaseStato.includes('errore')) {
      statusClass = 'status-error';
    }

    let rowHtml = `<tr>`;
    rowHtml += `<td data-label="${headers[0]}">${escapeHtml(titolo)}</td>`;
    rowHtml += `<td data-label="${headers[1]}"><span class="status-pill ${statusClass}">${escapeHtml(stato)}</span></td>`;
    rowHtml += `<td data-label="${headers[2]}" class="action-cell">
                        <button class="management-button preview-button"
                                data-titolo="${b64Encode(titolo)}"
                                data-testo="${b64Encode(testo)}"
                                data-image-url="${escapeHtml(imageUrl)}"
                                data-date-value="${escapeHtml(dateValue)}"
                                data-time-value="${escapeHtml(timeValue)}"
                                data-row-number="${sheetRowNumber}">Anteprima</button>
                        <button class="management-button delete-button" onclick="confermaEliminaPost(${sheetRowNumber})">
                          &times;
                        </button>
                      </td>`;
    return rowHtml + '</tr>';
  });
}

async function confermaEliminaPost(riga) {
  if (confirm(`Sei sicuro di voler eliminare il post alla riga ${riga}? L'azione è irreversibile.`)) {
    const rowElement = document.querySelector(`button[data-row-number="${riga}"]`).closest('tr');
    if (rowElement) {
      rowElement.style.opacity = '0.5';
      rowElement.style.pointerEvents = 'none';
    }

    try {
      await callAppsScript('eliminaPostInAttesa', { riga });
      mostraFoglio('eseguiMostraFoglioInAttesa');
    } catch (error) {
      alert('Errore durante l\'eliminazione: ' + error.message);
      if (rowElement) {
        rowElement.style.opacity = '1';
        rowElement.style.pointerEvents = 'auto';
      }
    }
  }
}

function openPreviewModal(titolo, testo, imageUrl, data, ora, rowNumber) {
  const modal = document.getElementById('preview-modal');
  modal.dataset.rowNumber = rowNumber;

  document.getElementById('preview-title').innerText = b64Decode(titolo);
  document.getElementById('preview-text').innerText = b64Decode(testo).replace(/\n/g, '\n');

  document.getElementById('modal-schedule-date').value = data || '';
  document.getElementById('modal-schedule-time').value = ora || '';

  document.getElementById('modal-status').innerText = '';
  document.getElementById('modal-schedule-button').disabled = false;

  const imageElement = document.getElementById('preview-image');

  modal.style.display = 'flex';

  if (imageUrl) {
    imageElement.src = imageUrl;
    imageElement.style.display = 'block';
  } else {
    imageElement.src = '';
    imageElement.style.display = 'none';
  }
}

function closePreviewModal() {
  document.getElementById('preview-modal').style.display = 'none';
}

async function salvaModificheFoglio2(riga) {
  const postInstagram = document.getElementById(`cell-${riga}-3`).innerText;
  const postSlide = document.getElementById(`cell-${riga}-4`).innerText;
  const statusSpan = document.getElementById(`status-foglio2-${riga}`);
  statusSpan.innerText = 'Salvataggio...';

  try {
    await callAppsScript('aggiornaPostFoglio2', { riga, postInstagram, postSlide });
    statusSpan.innerText = '✅';
    setTimeout(() => { statusSpan.innerText = ''; }, 2000);
  } catch (error) {
    statusSpan.innerText = '❌ Errore';
  }
}

async function inviaModifica(checkboxElement, riga) {
  const nuovoStato = checkboxElement.checked ? 'Sì' : 'No';
  const statusSpan = document.getElementById('status-' + riga);
  statusSpan.innerText = 'Salvataggio...';

  try {
    await callAppsScript('aggiornaStatoNotizia', { riga, stato: nuovoStato });
    statusSpan.innerText = '✅';
    setTimeout(() => { statusSpan.innerText = ''; }, 2000);
  } catch (error) {
    statusSpan.innerText = '❌ Errore';
    checkboxElement.checked = !checkboxElement.checked;
  }
}

let fontiData = [];
let fontiCaricate = false;

async function toggleDropdown() {
  const dropdown = document.getElementById('fonti-dropdown');
  const isVisible = dropdown.classList.toggle('show');
  if (isVisible && !fontiCaricate) {
    try {
      const fonti = await callAppsScript('getFonti');
      onFontiLoaded(fonti);
    } catch (error) {
      onFontiFailure(error);
    }
  }
}

function onFontiLoaded(fonti) {
  fontiCaricate = true;
  fontiData = fonti;
  const list = document.getElementById('fonti-list');
  const loader = document.getElementById('fonti-loader');
  list.innerHTML = '';
  if (fonti.length === 0) {
    list.innerHTML = '<li style="padding: 0 10px;">Nessuna fonte configurata.</li>';
  } else {
    fonti.forEach((source, index) => {
      const li = document.createElement('li');
      li.innerHTML = `<input type="checkbox" id="fonte-${index}" ${source.active ? 'checked' : ''}><label for="fonte-${index}">${escapeHtml(source.url)}</label>`;
      list.appendChild(li);
    });
  }
  loader.style.display = 'none';
}

function onFontiFailure(error) {
  document.getElementById('fonti-loader').innerText = 'Errore nel caricamento.';
}

async function saveFontiChanges() {
  const saveButton = document.getElementById('fonti-save-button');
  const status = document.getElementById('fonti-status');
  saveButton.disabled = true;
  status.textContent = 'Salvataggio...';
  const newSources = fontiData.map((source, index) => {
    const checkbox = document.getElementById('fonte-' + index);
    return { url: source.url, active: checkbox.checked };
  });

  try {
    await callAppsScript('salvaStatoFonti', { fonti: newSources });
    status.textContent = '✅ Salvato!';
    fontiCaricate = false; // Force reload on next open
    setTimeout(() => {
      document.getElementById('fonti-dropdown').classList.remove('show');
      saveButton.disabled = false;
      status.textContent = '';
    }, 1500);
  } catch (error) {
    status.textContent = '❌ Errore';
    saveButton.disabled = false;
  }
}

window.onclick = function (event) {
  if (!event.target.closest('.dropdown-container') && !event.target.matches('.management-button')) {
    const dropdown = document.getElementById('fonti-dropdown');
    if (dropdown.classList.contains('show')) {
      dropdown.classList.remove('show');
    }
  }
}

function escapeHtml(text) {
  if (text === null || text === undefined) { return ''; }
  var map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' };
  return text.toString().replace(/[&<>"]/g, function (m) { return map[m]; });
}

async function updateDashboard() {
  const refreshButton = document.getElementById('refresh-dashboard');
  if (refreshButton) {
    refreshButton.classList.add('loading');
    refreshButton.disabled = true;
  }

  try {
    const stats = await callAppsScript('getDashboardStats');
    document.getElementById('stat-notizie').innerText = stats.notizie;
    document.getElementById('stat-post').innerText = stats.post;
    document.getElementById('stat-immagini').innerText = stats.immagini;
    document.getElementById('stat-esporta').innerText = stats.esporta;
  } catch (err) {
    document.getElementById('stat-notizie').innerText = '⚠️';
    document.getElementById('stat-post').innerText = '⚠️';
    document.getElementById('stat-immagini').innerText = '⚠️';
    document.getElementById('stat-esporta').innerText = '⚠️';
  } finally {
    if (refreshButton) {
      refreshButton.classList.remove('loading');
      refreshButton.disabled = false;
    }
  }
}

document.addEventListener('DOMContentLoaded', () => {
  updateDashboard();

  document.getElementById('preview-modal-close-btn').addEventListener('click', closePreviewModal);
  document.getElementById('preview-modal').addEventListener('click', function (event) {
    if (event.target === this) {
      closePreviewModal();
    }
  });

  document.getElementById('foglio-container').addEventListener('click', function (event) {
    const previewButton = event.target.closest('.preview-button');
    if (previewButton) {
      const data = previewButton.dataset;
      openPreviewModal(
        data.titolo,
        data.testo,
        data.imageUrl,
        data.dateValue,
        data.timeValue,
        data.rowNumber
      );
    }
  });

  document.getElementById('modal-schedule-button').addEventListener('click', async function () {
    const modal = document.getElementById('preview-modal');
    const rowNumber = modal.dataset.rowNumber;
    const dateValue = document.getElementById('modal-schedule-date').value;
    const timeValue = document.getElementById('modal-schedule-time').value;

    const statusSpan = document.getElementById('modal-status');
    const scheduleButton = document.getElementById('modal-schedule-button');

    if (!rowNumber || !dateValue || !timeValue) {
      statusSpan.innerText = 'Data e ora sono obbligatorie.';
      statusSpan.style.color = 'red';
      return;
    }

    statusSpan.innerText = 'Programmazione in corso...';
    statusSpan.style.color = '';
    scheduleButton.disabled = true;

    try {
      const response = await callAppsScript('programmaPostDaModale', { riga: rowNumber, data: dateValue, ora: timeValue });
      statusSpan.innerText = '✅ ' + response;
      scheduleButton.disabled = false;
      setTimeout(() => {
        closePreviewModal();
        mostraFoglio('eseguiMostraFoglioInAttesa');
      }, 2000);
    } catch (error) {
      statusSpan.innerText = '❌ ' + error.message;
      scheduleButton.disabled = false;
    }
  });
});

/******************************************************************/
/******************      LOGICA EDITOR MODALE      ******************/
/******************************************************************/

let editorAbortController = new AbortController();

function openImageEditor(buttonElement) {
  const imageUrl = buttonElement.getAttribute('data-image-url');
  const postTitle = b64Decode(buttonElement.getAttribute('data-post-title'));
  const postBody = b64Decode(buttonElement.getAttribute('data-post-body'));

  initImageEditor(imageUrl, postTitle, postBody);
  document.getElementById('editor-modal-overlay').style.display = 'flex';
}

function closeImageEditor() {
  document.getElementById('editor-modal-overlay').style.display = 'none';
}

async function initImageEditor(imageUrl, postTitle, postBody) {
  document.getElementById('font-size').value = 50;
  document.getElementById('line-height').value = 60;
  document.getElementById('text-color').value = '#FFFFFF';
  document.getElementById('font-weight').value = '700';
  document.getElementById('font-style').value = 'normal';
  document.getElementById('text-align').value = 'center';
  document.getElementById('gradient-color').value = '#000000';
  document.getElementById('gradient-opacity').value = '0.7';
  document.getElementById('shadow-blur').value = '0';

  editorAbortController.abort();
  editorAbortController = new AbortController();
  const signal = editorAbortController.signal;

  const canvas = document.getElementById('image-editor');
  const ctx = canvas.getContext('2d');
  const colorPicker = document.getElementById('gradient-color');
  const opacitySlider = document.getElementById('gradient-opacity');
  const textInput = document.getElementById('text-input');
  const fontSizeInput = document.getElementById('font-size');
  const textAlignInput = document.getElementById('text-align');
  const fontWeightInput = document.getElementById('font-weight');
  const fontStyleInput = document.getElementById('font-style');
  const textColorInput = document.getElementById('text-color');
  const shadowBlurInput = document.getElementById('shadow-blur');
  const postBodyPreview = document.getElementById('post-body-preview');
  const saveBtn = document.getElementById('save-btn');

  let img = new Image();
  let offsetX = 0, offsetY = 0, scale = 1, isDragging = false, startX, startY;
  let gradientColor = '#000000', gradientOpacity = 0.7, fontSize = 50, textAlign = 'center', fontWeight = 700, fontStyle = 'normal', textColor = '#FFFFFF', shadowBlur = 0, lineHeight = 60;
  let editablePostTitle = postTitle || '';

  postBodyPreview.innerHTML = (postBody || '').replace(/\n/g, '<br>');
  textInput.innerHTML = editablePostTitle;

  // Use a proxy to fetch the image if it's from an external domain to avoid CORS issues on canvas
  const proxiedImageUrl = `${SCRIPT_URL}?action=getImage&url=${encodeURIComponent(imageUrl)}`;
  img.crossOrigin = "Anonymous";
  img.src = proxiedImageUrl;

  img.onload = () => {
    const canvasAspect = canvas.width / canvas.height;
    const imgAspect = img.width / img.height;
    if (imgAspect > canvasAspect) {
      scale = canvas.height / img.height;
      offsetX = (canvas.width - img.width * scale) / 2;
      offsetY = 0;
    } else {
      scale = canvas.width / img.width;
      offsetX = 0;
      offsetY = (canvas.height - img.height * scale) / 2;
    }
    drawImage();
  };

  img.onerror = () => {
      console.error("Failed to load image from", proxiedImageUrl);
      ctx.fillStyle = '#f0f0f0';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.fillStyle = '#666';
      ctx.textAlign = 'center';
      ctx.font = '24px Arial';
      ctx.fillText('Errore caricamento immagine', canvas.width / 2, canvas.height / 2);
  }

  function hexToRgba(hex, opacity) {
    hex = hex.replace('#', '');
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);
    return `rgba(${r}, ${g}, ${b}, ${opacity})`;
  }

  function drawImage() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#000';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, offsetX, offsetY, img.width * scale, img.height * scale);

    const gradient = ctx.createLinearGradient(0, canvas.height, 0, canvas.height / 2);
    gradient.addColorStop(0, hexToRgba(gradientColor, gradientOpacity));
    gradient.addColorStop(1, hexToRgba(gradientColor, 0));
    ctx.fillStyle = gradient;
    ctx.fillRect(0, canvas.height / 2, canvas.width, canvas.height);

    ctx.fillStyle = textColor;
    ctx.font = `${fontStyle} ${fontWeight} ${fontSize}px Roboto, Arial, sans-serif`;
    ctx.textAlign = textAlign;
    ctx.textBaseline = 'bottom';
    const maxWidth = canvas.width - 100;

    let x;
    if (textAlign === 'left') x = 50;
    else if (textAlign === 'right') x = canvas.width - 50;
    else x = canvas.width / 2;

    const y = canvas.height - 90;
    ctx.shadowColor = 'black';
    ctx.shadowBlur = shadowBlur;
    ctx.shadowOffsetX = 2;
    ctx.shadowOffsetY = 2;
    wrapText(ctx, editablePostTitle, x, y, maxWidth, lineHeight);
    ctx.shadowColor = 'transparent';
    ctx.shadowBlur = 0;
    ctx.shadowOffsetX = 0;
    ctx.shadowOffsetY = 0;
  }

  function wrapText(context, text, x, y, maxWidth, lineHeight) {
    const textInput = document.getElementById('text-input');
    let lines = [];
    let currentLine = [];
    let currentLineWidth = 0;

    function processNodeForMeasurement(node) {
      if (node.nodeType === Node.TEXT_NODE) {
        const words = node.textContent.split(' ');
        words.forEach(word => {
          if (!word) return;

          let finalWeight = fontWeight;
          let finalStyle = fontStyle;
          let itemColor = textColor; // Default to the main text color
          let parent = node.parentNode;

          while (parent && parent !== textInput) {
            const tagName = parent.tagName;
            if (tagName === 'B' || tagName === 'STRONG') finalWeight = '700';
            if (tagName === 'I' || tagName === 'EM') finalStyle = 'italic';
            if (tagName === 'FONT' && parent.hasAttribute('color')) {
              itemColor = parent.getAttribute('color');
            }
            parent = parent.parentNode;
          }

          const tempFont = `${finalStyle} ${finalWeight} ${fontSize}px Roboto, Arial, sans-serif`;
          context.font = tempFont;

          const upperWord = word.toUpperCase();
          const wordWidth = context.measureText(upperWord + ' ').width;

          if (currentLineWidth + wordWidth > maxWidth && currentLine.length > 0) {
            lines.push({ items: currentLine, width: currentLineWidth });
            currentLine = [];
            currentLineWidth = 0;
          }
          currentLine.push({ word: upperWord, font: context.font, color: itemColor, width: wordWidth });
          currentLineWidth += wordWidth;
        });
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        if (node.tagName === 'DIV' || node.tagName === 'P' || node.tagName === 'BR') {
          if (currentLine.length > 0) {
            lines.push({ items: currentLine, width: currentLineWidth });
            currentLine = [];
            currentLineWidth = 0;
          }
          if (node.tagName !== 'BR') {
             node.childNodes.forEach(processNodeForMeasurement);
             if (currentLine.length > 0) {
                lines.push({ items: currentLine, width: currentLineWidth });
                currentLine = [];
                currentLineWidth = 0;
             }
          }
        } else {
          node.childNodes.forEach(processNodeForMeasurement);
        }
      }
    }

    textInput.childNodes.forEach(processNodeForMeasurement);
    if (currentLine.length > 0) {
      lines.push({ items: currentLine, width: currentLineWidth });
    }

    context.textAlign = 'left';
    let currentY = y - (lines.length - 1) * lineHeight;

    lines.forEach(line => {
      let currentX;
      if (textAlign === 'center') {
        currentX = x - (line.width / 2);
      } else if (textAlign === 'right') {
        currentX = x - line.width;
      } else { // left
        currentX = x;
      }

      line.items.forEach(item => {
        context.font = item.font;
        context.fillStyle = item.color;
        context.fillText(item.word, currentX, currentY);
        currentX += item.width;
      });
      currentY += lineHeight;
    });

    context.fillStyle = textColor;
  }

  colorPicker.addEventListener('input', (e) => { gradientColor = e.target.value; drawImage(); }, { signal });
  opacitySlider.addEventListener('input', (e) => { gradientOpacity = parseFloat(e.target.value); drawImage(); }, { signal });
  textInput.addEventListener('input', (e) => { editablePostTitle = e.target.innerHTML; drawImage(); }, { signal });
  fontSizeInput.addEventListener('input', (e) => { fontSize = parseInt(e.target.value, 10); drawImage(); }, { signal });
  document.getElementById('line-height').addEventListener('input', (e) => { lineHeight = parseInt(e.target.value, 10); drawImage(); }, { signal });
  textAlignInput.addEventListener('input', (e) => { textAlign = e.target.value; drawImage(); }, { signal });
  fontWeightInput.addEventListener('input', (e) => { fontWeight = parseInt(e.target.value, 10); drawImage(); }, { signal });
  fontStyleInput.addEventListener('input', (e) => { fontStyle = e.target.value; drawImage(); }, { signal });
  textColorInput.addEventListener('input', (e) => { textColor = e.target.value; drawImage(); }, { signal });
  shadowBlurInput.addEventListener('input', (e) => { shadowBlur = parseInt(e.target.value, 10); drawImage(); }, { signal });

  document.getElementById('modal-close-btn').addEventListener('click', closeImageEditor, { signal });

  saveBtn.onclick = async () => {
    saveBtn.disabled = true;
    saveBtn.textContent = 'Salvataggio...';
    const imageData = canvas.toDataURL('image/png');
    const finalTitle = textInput.innerHTML;
    const finalBody = postBody;

    try {
      const response = await callAppsScript('_salvaImmagineEPostInAttesa', { immagineBase64: imageData, titolo: finalTitle, testo: finalBody });
      alert(response);
      closeImageEditor();
      mostraFoglio('eseguiMostraFoglioInAttesa');
    } catch (error) {
      alert("Errore: " + error.message);
    } finally {
      saveBtn.disabled = false;
      saveBtn.textContent = 'Salva';
    }
  };

  function getCanvasCoordinates(e) {
    const rect = canvas.getBoundingClientRect();
    const clientX = e.clientX || e.touches[0].clientX;
    const clientY = e.clientY || e.touches[0].clientY;
    const canvasX = clientX - rect.left;
    const canvasY = clientY - rect.top;
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    return { x: canvasX * scaleX, y: canvasY * scaleY };
  }

  function onDragStart(e) { isDragging = true; const coords = getCanvasCoordinates(e); startX = coords.x - offsetX; startY = coords.y - offsetY; canvas.style.cursor = 'grabbing'; }
  function onDragEnd() { isDragging = false; canvas.style.cursor = 'grab'; }
  function onDrag(e) { if (!isDragging) return; const coords = getCanvasCoordinates(e); handleMove(coords.x, coords.y); }

  canvas.addEventListener('mousedown', onDragStart, { signal });
  canvas.addEventListener('mouseup', onDragEnd, { signal });
  canvas.addEventListener('mouseleave', onDragEnd, { signal });
  canvas.addEventListener('mousemove', onDrag, { signal });
  canvas.addEventListener('touchstart', (e) => { e.preventDefault(); onDragStart(e); }, { signal });
  canvas.addEventListener('touchend', (e) => { e.preventDefault(); onDragEnd(e); }, { signal });
  canvas.addEventListener('touchmove', (e) => { e.preventDefault(); onDrag(e); }, { signal });

  function handleMove(currentX, currentY) {
    const newOffsetX = currentX - startX;
    const newOffsetY = currentY - startY;
    const maxOffsetX = 0;
    const minOffsetX = canvas.width - img.width * scale;
    const maxOffsetY = 0;
    const minOffsetY = canvas.height - img.height * scale;
    offsetX = Math.max(minOffsetX, Math.min(newOffsetX, maxOffsetX));
    offsetY = Math.max(minOffsetY, Math.min(newOffsetY, maxOffsetY));
    drawImage();
  }
}
