const fileInput = document.getElementById('templateFile');
const dropZone = document.getElementById('dropZone');
const statusBadge = document.getElementById('statusBadge');
const templateInfo = document.getElementById('templateInfo');
const fieldForm = document.getElementById('fieldForm');
const previewBtn = document.getElementById('previewBtn');
const processBtn = document.getElementById('processBtn');
const previewPanel = document.getElementById('previewPanel');
const previewContent = document.getElementById('previewContent');
const resetBtn = document.getElementById('resetBtn');

const placeholderRegex = /\{\{([^{}]+)\}\}/g;
let currentZip = null;
let baseDocumentXml = '';
let placeholders = [];
let currentFileName = '';

function init() {
  fileInput.addEventListener('change', handleFileSelection);

  ['dragenter', 'dragover'].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.add('active');
    });
  });

  ['dragleave', 'dragend', 'drop'].forEach((eventName) => {
    dropZone.addEventListener(eventName, () => {
      dropZone.classList.remove('active');
    });
  });

  dropZone.addEventListener('drop', (event) => {
    event.preventDefault();
    const [file] = event.dataTransfer.files;
    if (file) {
      handleFile(file);
    }
  });

  previewBtn.addEventListener('click', () => {
    previewPanel.classList.add('active');
    renderPreview();
  });

  fieldForm.addEventListener('input', renderPreview);
  fieldForm.addEventListener('change', renderPreview);
  processBtn.addEventListener('click', processAndDownload);
  resetBtn.addEventListener('click', resetApp);
}

async function handleFileSelection(event) {
  const [file] = event.target.files;
  if (file) {
    await handleFile(file);
  }
}

async function handleFile(file) {
  if (!file.name.toLowerCase().endsWith('.docx')) {
    setStatus('Solo se aceptan archivos .docx', false);
    return;
  }

  try {
    setStatus('Leyendo plantilla...', true);
    const zip = await JSZip.loadAsync(file);
    const documentXml = await zip.file('word/document.xml').async('string');

    if (!documentXml) {
      throw new Error('No se encontró el contenido del documento en la plantilla.');
    }

    currentZip = zip;
    baseDocumentXml = documentXml;
    currentFileName = file.name.replace(/\.docx$/i, '');
    placeholders = extractPlaceholdersFromXml(documentXml);

    if (!placeholders.length) {
      setStatus('La plantilla no contiene marcadores.', false);
      templateInfo.textContent = 'No se encontraron campos tipo {{EJEMPLO}}. Prueba con otra plantilla.';
      fieldForm.innerHTML = '';
      processBtn.disabled = true;
      previewContent.innerHTML = '<p>No hay campos para completar.</p>';
      return;
    }

    renderForm(placeholders);
    previewPanel.classList.add('active');
    renderPreview();
    setStatus(`Plantilla lista: ${file.name}`, true);
    templateInfo.textContent = `Campos detectados: ${placeholders.join(', ')}`;
    processBtn.disabled = false;
  } catch (error) {
    console.error(error);
    setStatus('No se pudo leer la plantilla.', false);
    templateInfo.textContent = 'Ocurrió un error al leer el archivo. Intenta con otra plantilla.';
  }
}

function renderForm(items) {
  const formFields = items.map((placeholder) => {
    const label = humanize(placeholder);
    const isAutomatic = placeholder === 'RECONOCER_PERSONERIA';

    if (isAutomatic) {
      return `
        <div class="field toggle-field">
          <label class="toggle-row">
            <input id="toggle-${placeholder}" name="${placeholder}" type="checkbox" />
            <span>${label}</span>
          </label>
          <input
            name="APODERADO_DEMANDANTE"
            type="text"
            placeholder="NO LLENAR - texto automático"
          />
        </div>
      `;
    }

    if (placeholder === 'CAUSALES') {
      return `
        <div class="field full-width">
          <span>${label}</span>
          <div class="causal-picker">
            <input id="causalSearch" type="text" placeholder="Buscar causal..." />
            <div id="causalList" class="causal-list"></div>
          </div>
          <textarea name="${placeholder}" rows="6" placeholder="Las causales seleccionadas aparecerán aquí"></textarea>
        </div>
      `;
    }

    return `
      <label class="field">
        <span>${label}</span>
        <input name="${placeholder}" type="text" placeholder="Ingresa ${label.toLowerCase()}" />
      </label>
    `;
  });

  fieldForm.innerHTML = formFields.join('');
  if (items.includes('CAUSALES')) {
    initCausalPicker();
  }
}

function renderPreview() {
  const previewText = buildPreviewText(baseDocumentXml);
  if (!previewText) {
    previewContent.innerHTML = '<p>Tu contenido aparecerá aquí cuando completes los campos.</p>';
    return;
  }

  const paragraphs = previewText
    .split(/\n{2,}/)
    .map((paragraph) => `<p>${escapeHtml(paragraph)}</p>`)
    .join('');

  previewContent.innerHTML = paragraphs;
}

function buildPreviewText(xml) {
  if (!xml) return '';

  const values = getFieldValues();
  const rawText = xml
    .replace(/<w:br\b[^>]*\/?>/gi, '\n')
    .replace(/<\/w:p>/gi, '\n\n')
    .replace(/<w:p\b[^>]*>/gi, '')
    .replace(/<w:t\b[^>]*>(.*?)<\/w:t>/gis, (_, content) => decodeXmlEntities(content))
    .replace(/<[^>]+>/g, '');

  const previewText = rawText.replace(placeholderRegex, (_, key) => {
    const normalizedKey = key.trim();
    if (normalizedKey === 'RECONOCER_PERSONERIA') {
      return getSpecialPlaceholderValue(normalizedKey, values);
    }
    if (normalizedKey === 'CAUSALES') {
      const raw = values[normalizedKey] || '';
      const items = raw
        .split(/\r?\n\s*\r?\n|[\r\n]+/)
        .map((s) => s.trim())
        .filter(Boolean)
        .map((s) => s.replace(/^\s*\d+[\.)]\s*/, ''));
      return items.map((it, i) => `${i + 1}. ${it}`).join('\n\n');
    }
    return values[normalizedKey] || '';
  });

  return previewText.trim().replace(/\n{3,}/g, '\n\n');
}

function extractPlaceholdersFromXml(xml) {
  const placeholders = new Set();
  const textNodes = xml.matchAll(/<w:t\b[^>]*>(.*?)<\/w:t>/gs);

  for (const match of textNodes) {
    const textContent = match[1];
    for (const placeholderMatch of textContent.matchAll(placeholderRegex)) {
      const key = placeholderMatch[1].trim();
      if (key) {
        placeholders.add(key);
      }
    }
  }

  if (placeholders.has('RECONOCER_PERSONERIA')) {
    placeholders.add('APODERADO_DEMANDANTE');
  }

  return [...placeholders];
}

function replacePlaceholdersInXml(xml, values) {
  const causalesRaw = values['CAUSALES'] || '';
  xml = replaceCausalesPlaceholder(xml, causalesRaw);

  return xml.replace(/<w:t\b[^>]*>(.*?)<\/w:t>/gs, (fullMatch) => {
    return fullMatch.replace(placeholderRegex, (_, key) => {
      const normalizedKey = key.trim();
      if (normalizedKey === 'RECONOCER_PERSONERIA') {
        return getSpecialPlaceholderValue(normalizedKey, values);
      }
      if (normalizedKey === 'CAUSALES') {
        return '';
      }
      return values[normalizedKey] || '';
    });
  });
}

function replaceCausalesPlaceholder(xml, raw) {
  const items = raw
    .split(/\r?\n\s*\r?\n|[\r\n]+/)
    .map((s) => s.trim())
    .filter(Boolean)
    .map((s) => s.replace(/^\s*\d+[\.)]\s*/, ''));

  if (!items.length) return xml;

  const numbered = items.map((it, i) => escapeHtml(`${i + 1}. ${it}`));

  return xml.replace(/(<w:t\b[^>]*>)([^<]*?)\{\{CAUSALES\}\}([^<]*?)(<\/w:t>)/gs, (_, open, before, after) => {
    const parts = [`${open}${before}</w:t>`];
    numbered.forEach((item, index) => {
      if (index > 0) {
        parts.push('<w:br/><w:br/>');
      }
      parts.push(`<w:t>${item}</w:t>`);
    });
    if (after) {
      parts.push(`<w:br/><w:t>${after}</w:t>`);
    }
    return parts.join('');
  });
}

function getFieldValues() {
  return Array.from(fieldForm.querySelectorAll('input, textarea')).reduce((acc, input) => {
    if (input.type === 'checkbox') {
      acc[input.name] = input.checked;
    } else {
      acc[input.name] = input.value.trim();
    }
    return acc;
  }, {});
}

function getFieldValue(key) {
  return getFieldValues()[key] || '';
}

function getSpecialPlaceholderValue(key, values) {
  if (key === 'RECONOCER_PERSONERIA') {
    const enabled = values.RECONOCER_PERSONERIA === true || values.RECONOCER_PERSONERIA === 'on';
    if (!enabled) {
      return '';
    }

    const apoderado = values.APODERADO_DEMANDANTE || values.apoderado_demandante || '';
    return `SEXTO: Se reconoce personería para actuar a ${apoderado} como apoderado(a) judicial de la parte demandante, en los términos y para los efectos del poder conferido.`;
  }

  return '';
}

async function processAndDownload() {
  if (!currentZip) {
    setStatus('Primero carga una plantilla.', false);
    return;
  }

  try {
    setStatus('Procesando documento...', true);
    const outputZip = currentZip;
    const processedXml = replacePlaceholdersInXml(baseDocumentXml, getFieldValues());
    outputZip.file('word/document.xml', processedXml);

    const blob = await outputZip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });

    const radicado = getFieldValue('RADICADO') || getFieldValue('radicado') || '';
    const safeRadicado = radicado.replace(/[^a-z0-9._-]+/gi, '-').replace(/-+/g, '-').replace(/^-|-$/g, '');
    const safeTemplateName = currentFileName.replace(/[^a-z0-9._-]+/gi, '-').replace(/-+/g, '-').replace(/^-|-$/g, '');
    const downloadName = [safeRadicado, safeTemplateName].filter(Boolean).join('_') || 'documento-procesado';

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${downloadName}.docx`;
    link.click();
    URL.revokeObjectURL(url);

    setStatus('Documento descargado', true);
  } catch (error) {
    console.error(error);
    setStatus('No se pudo generar el documento.', false);
  }
}

function resetApp() {
  fileInput.value = '';
  fieldForm.innerHTML = '';
  previewContent.innerHTML = '<p>Tu contenido aparecerá aquí cuando cargues una plantilla y completes los campos.</p>';
  previewPanel.classList.remove('active');
  processBtn.disabled = true;
  currentZip = null;
  baseDocumentXml = '';
  placeholders = [];
  currentFileName = '';
  templateInfo.textContent = 'Aún no has cargado ninguna plantilla.';
  setStatus('Esperando archivo', false);
}

function humanize(value) {
  return value
    .replace(/_/g, ' ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^./, (char) => char.toUpperCase());
}

function decodeXmlEntities(value) {
  return value
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function escapeHtml(value) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\"/g, '&quot;');
}

function initCausalPicker() {
  const searchInput = document.getElementById('causalSearch');
  const list = document.getElementById('causalList');
  const textarea = fieldForm.querySelector('textarea[name="CAUSALES"]');

  let selectedOrder = [];

  const renderCausales = () => {
    const query = searchInput.value.toLowerCase();
    const filtered = (window.causalesData || []).filter((causal) => {
      return causal.texto.toLowerCase().includes(query);
    });

    list.innerHTML = filtered.slice(0, 20).map((causal) => `
      <div class="causal-item">
        <label>
          <input type="checkbox" value="${causal.texto}" />
          <span><strong>${causal.id}</strong> — ${causal.texto}</span>
        </label>
      </div>
    `).join('');

    list.querySelectorAll('input[type="checkbox"]').forEach((checkbox) => {
      // restore previous checked state if this value is in selectedOrder
      if (selectedOrder.includes(checkbox.value)) checkbox.checked = true;

      checkbox.addEventListener('change', () => {
        if (checkbox.checked) {
          if (!selectedOrder.includes(checkbox.value)) selectedOrder.push(checkbox.value);
        } else {
          selectedOrder = selectedOrder.filter((v) => v !== checkbox.value);
        }

        // store raw items in the textarea (no numbering) so we don't double-number later
        textarea.value = selectedOrder.join('\n\n');
        renderPreview();
      });
    });
  };

  searchInput.addEventListener('input', renderCausales);
  renderCausales();
}

function setStatus(message, ready) {
  statusBadge.textContent = message;
  statusBadge.className = `status-pill ${ready ? 'ready' : ''}`.trim();
}

init();
