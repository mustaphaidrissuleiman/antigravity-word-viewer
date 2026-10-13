// Word Document Viewer & Editor - Client Application
(function () {
  'use strict';

  // -------------------------------------------------------------------------
  // State
  // -------------------------------------------------------------------------
  const state = {
    currentPath: null,
    currentName: null,
    currentExtension: '.docx',
    currentArrayBuffer: null,
    currentRawText: '',
    currentHtml: '',
    zoom: 1.0,
    currentPage: 1,
    totalPages: 1,
    viewMode: 'page', // 'page' | 'flow'
    isEditing: false,
    hasUnsavedChanges: false,
    originalDocXml: '',
    originalSectPr: '',
    workspaceUris: [],
    searchResults: [],
    currentSearchIndex: -1,
    docMetadata: {},
  };

  // -------------------------------------------------------------------------
  // DOM Elements
  // -------------------------------------------------------------------------
  const $ = (id) => document.getElementById(id);

  const els = {
    fileInput: $('fileInput'),
    docToolbar: $('docToolbar'),
    toolbarOpenBtn: $('toolbarOpenBtn'),
    docBadge: $('docBadge'),
    docName: $('docName'),
    modePageBtn: $('modePageBtn'),
    modeFlowBtn: $('modeFlowBtn'),
    pageControls: $('pageControls'),
    prevPageBtn: $('prevPageBtn'),
    nextPageBtn: $('nextPageBtn'),
    pageNumInput: $('pageNumInput'),
    pageCount: $('pageCount'),
    tocToggleBtn: $('tocToggleBtn'),
    tocSidebar: $('tocSidebar'),
    closeTocBtn: $('closeTocBtn'),
    tocList: $('tocList'),
    searchToggleBtn: $('searchToggleBtn'),
    searchDrawer: $('searchDrawer'),
    searchInput: $('searchInput'),
    searchMatchCount: $('searchMatchCount'),
    searchPrevBtn: $('searchPrevBtn'),
    searchNextBtn: $('searchNextBtn'),
    closeSearchBtn: $('closeSearchBtn'),
    zoomControls: $('zoomControls'),
    zoomOutBtn: $('zoomOutBtn'),
    zoomResetBtn: $('zoomResetBtn'),
    zoomInBtn: $('zoomInBtn'),
    fitWidthBtn: $('fitWidthBtn'),
    docInfoBtn: $('docInfoBtn'),
    docInfoDrawer: $('docInfoDrawer'),
    closeInfoDrawerBtn: $('closeInfoDrawerBtn'),
    editToggleBtn: $('editToggleBtn'),
    editorRibbon: $('editorRibbon'),
    formatBlockSelect: $('formatBlockSelect'),
    boldBtn: $('boldBtn'),
    italicBtn: $('italicBtn'),
    underlineBtn: $('underlineBtn'),
    strikeBtn: $('strikeBtn'),
    bulletListBtn: $('bulletListBtn'),
    numListBtn: $('numListBtn'),
    alignLeftBtn: $('alignLeftBtn'),
    alignCenterBtn: $('alignCenterBtn'),
    alignRightBtn: $('alignRightBtn'),
    undoBtn: $('undoBtn'),
    redoBtn: $('redoBtn'),
    saveStatus: $('saveStatus'),
    saveDocBtn: $('saveDocBtn'),
    doneEditBtn: $('doneEditBtn'),
    askAgentBtn: $('askAgentBtn'),
    agentPromptDrawer: $('agentPromptDrawer'),
    closePromptDrawerBtn: $('closePromptDrawerBtn'),
    promptDocName: $('promptDocName'),
    agentPromptInput: $('agentPromptInput'),
    submitPromptBtn: $('submitPromptBtn'),
    viewerContainer: $('viewerContainer'),
    emptyState: $('emptyState'),
    dropzone: $('dropzone'),
    emptyOpenBtn: $('emptyOpenBtn'),
    workspaceList: $('workspaceList'),
    workspaceItems: $('workspaceItems'),
    refreshWorkspaceBtn: $('refreshWorkspaceBtn'),
    loadingOverlay: $('loadingOverlay'),
    loadingMessage: $('loadingMessage'),
    errorState: $('errorState'),
    errorMessage: $('errorMessage'),
    errorRetryBtn: $('errorRetryBtn'),
    documentViewWrapper: $('documentViewWrapper'),
    docLegacyBanner: $('docLegacyBanner'),
    docxContainer: $('docxContainer'),
    flowContainer: $('flowContainer'),
    // Info elements
    infoFileName: $('infoFileName'),
    infoFileSize: $('infoFileSize'),
    infoWordCount: $('infoWordCount'),
    infoCharCount: $('infoCharCount'),
    infoParaCount: $('infoParaCount'),
    infoPageCount: $('infoPageCount'),
    infoReadingTime: $('infoReadingTime'),
    infoAuthor: $('infoAuthor'),
    infoLastModifiedBy: $('infoLastModifiedBy'),
    infoModifiedDate: $('infoModifiedDate'),
    copyTextBtn: $('copyTextBtn'),
    copyMarkdownBtn: $('copyMarkdownBtn'),
    printDocBtn: $('printDocBtn'),
  };

  // -------------------------------------------------------------------------
  // Helper Utilities
  // -------------------------------------------------------------------------

  function formatBytes(bytes) {
    if (!bytes || bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  function escapeXml(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  function base64ToArrayBuffer(base64) {
    const binaryString = window.atob(base64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
  }

  async function apiFetch(endpoint, data = {}, method = 'POST') {
    if (window.sidecar && typeof window.sidecar.fetch === 'function') {
      const res = await window.sidecar.fetch(endpoint, {
        method,
        body: method === 'GET' ? undefined : JSON.stringify(data),
      });
      return await res.json();
    }
    const res = await fetch(endpoint, {
      method,
      headers: { 'Content-Type': 'application/json' },
      body: method === 'GET' ? undefined : JSON.stringify(data),
    });
    return await res.json();
  }

  function showLoading(msg = 'Loading document...') {
    els.loadingMessage.textContent = msg;
    els.loadingOverlay.hidden = false;
    els.errorState.hidden = true;
  }

  function hideLoading() {
    els.loadingOverlay.hidden = true;
  }

  function showError(msg) {
    hideLoading();
    els.errorMessage.textContent = msg;
    els.errorState.hidden = false;
    els.documentViewWrapper.hidden = true;
    els.emptyState.hidden = true;
  }

  // -------------------------------------------------------------------------
  // OpenXML Serializer (Converts edited HTML DOM back into OpenXML <w:body>)
  // -------------------------------------------------------------------------

  function htmlToOpenXmlBody(htmlString, existingSectPrXml = '') {
    const parser = new DOMParser();
    const doc = parser.parseFromString(`<body>${htmlString}</body>`, 'text/html');
    const body = doc.body;

    let bodyXml = '<w:body>';

    function serializeBlock(node) {
      if (node.nodeType === 3) {
        const txt = node.nodeValue.trim();
        if (txt) {
          bodyXml += `<w:p><w:r><w:t xml:space="preserve">${escapeXml(txt)}</w:t></w:r></w:p>`;
        }
        return;
      }
      if (node.nodeType !== 1) return;

      const tag = node.tagName.toLowerCase();

      // Headings (H1 - H6)
      if (/^h[1-6]$/.test(tag)) {
        const level = tag[1];
        const runs = serializeInline(node);
        bodyXml += `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>${runs}</w:p>`;
        return;
      }

      // Paragraphs & Divs
      if (tag === 'p' || tag === 'div') {
        const runs = serializeInline(node);
        let jc = '';
        if (node.style.textAlign === 'center') jc = '<w:jc w:val="center"/>';
        else if (node.style.textAlign === 'right') jc = '<w:jc w:val="right"/>';
        else if (node.style.textAlign === 'justify') jc = '<w:jc w:val="both"/>';

        bodyXml += `<w:p>${jc ? `<w:pPr>${jc}</w:pPr>` : ''}${runs}</w:p>`;
        return;
      }

      // Blockquotes
      if (tag === 'blockquote') {
        const runs = serializeInline(node);
        bodyXml += `<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr>${runs}</w:p>`;
        return;
      }

      // Lists (UL & OL)
      if (tag === 'ul' || tag === 'ol') {
        node.childNodes.forEach((li) => {
          if (li.nodeType === 1 && li.tagName.toLowerCase() === 'li') {
            const runs = serializeInline(li);
            bodyXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>${runs}</w:p>`;
          }
        });
        return;
      }

      // Tables
      if (tag === 'table') {
        bodyXml += '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>';
        const rows = node.querySelectorAll('tr');
        rows.forEach((tr) => {
          bodyXml += '<w:tr>';
          const cells = tr.querySelectorAll('th, td');
          cells.forEach((cell) => {
            bodyXml += '<w:tc>';
            const cellRuns = serializeInline(cell);
            bodyXml += `<w:p>${cellRuns}</w:p>`;
            bodyXml += '</w:tc>';
          });
          bodyXml += '</w:tr>';
        });
        bodyXml += '</w:tbl>';
        return;
      }

      // Fallback: traverse children
      node.childNodes.forEach((child) => serializeBlock(child));
    }

    function serializeInline(element, styles = {}) {
      let runsXml = '';

      function walk(node, curStyles) {
        if (node.nodeType === 3) {
          const text = node.nodeValue;
          if (!text) return;
          let rPr = '';
          if (curStyles.bold) rPr += '<w:b/>';
          if (curStyles.italic) rPr += '<w:i/>';
          if (curStyles.underline) rPr += '<w:u w:val="single"/>';
          if (curStyles.strike) rPr += '<w:strike/>';
          runsXml += `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
          return;
        }

        if (node.nodeType === 1) {
          const tag = node.tagName.toLowerCase();
          const nextStyles = { ...curStyles };
          if (tag === 'b' || tag === 'strong') nextStyles.bold = true;
          if (tag === 'i' || tag === 'em') nextStyles.italic = true;
          if (tag === 'u') nextStyles.underline = true;
          if (tag === 's' || tag === 'strike' || tag === 'del') nextStyles.strike = true;
          if (tag === 'br') {
            runsXml += '<w:r><w:br/></w:r>';
            return;
          }

          node.childNodes.forEach((c) => walk(c, nextStyles));
        }
      }

      element.childNodes.forEach((c) => walk(c, styles));
      return runsXml || '<w:r><w:t></w:t></w:r>';
    }

    body.childNodes.forEach((child) => serializeBlock(child));

    // Preserve original section properties (page margins, header/footer references)
    if (existingSectPrXml) {
      bodyXml += existingSectPrXml;
    } else {
      bodyXml += '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
    }

    bodyXml += '</w:body>';
    return bodyXml;
  }

  // -------------------------------------------------------------------------
  // Legacy .doc Text Scanner
  // -------------------------------------------------------------------------

  function extractTextFromBinaryDoc(arrayBuffer) {
    const uint8 = new Uint8Array(arrayBuffer);
    const textChunks = [];
    let currentAscii = '';

    for (let i = 0; i < uint8.length; i++) {
      const b = uint8[i];
      if ((b >= 32 && b <= 126) || b === 10 || b === 13 || b === 9) {
        currentAscii += String.fromCharCode(b);
      } else {
        if (currentAscii.trim().length >= 4) {
          textChunks.push(currentAscii.trim());
        }
        currentAscii = '';
      }
    }
    if (currentAscii.trim().length >= 4) {
      textChunks.push(currentAscii.trim());
    }

    return textChunks.join('\n\n');
  }

  // -------------------------------------------------------------------------
  // Open & Load Document
  // -------------------------------------------------------------------------

  async function openDocumentFile(path) {
    if (!path) return;
    showLoading('Locating and reading Word document...');
    try {
      const resp = await apiFetch('/api/load-doc', {
        path,
        workspaceUris: state.workspaceUris,
      });

      if (!resp.success) {
        showError(resp.error || 'Failed to load document');
        return;
      }

      state.currentPath = resp.path;
      state.currentName = resp.name;
      state.currentExtension = resp.extension;

      let buffer;
      if (resp.base64) {
        buffer = base64ToArrayBuffer(resp.base64);
      } else {
        const url = `/api/raw-doc?path=${encodeURIComponent(resp.path)}`;
        const binRes = await fetch(url);
        buffer = await binRes.arrayBuffer();
      }

      state.currentArrayBuffer = buffer;
      await renderLoadedDocument(resp.size, resp.mtime);

      apiFetch('/api/save-state', { lastOpenedPath: resp.path });
    } catch (err) {
      console.error('Failed to open document:', err);
      showError(err.message || 'Error parsing document');
    }
  }

  async function renderLoadedDocument(fileSize = 0, mtime = Date.now()) {
    showLoading('Rendering Word document...');

    els.emptyState.hidden = true;
    els.errorState.hidden = true;
    els.documentViewWrapper.hidden = false;
    els.docToolbar.hidden = false;

    els.docName.textContent = state.currentName;
    els.docName.title = state.currentPath;
    els.promptDocName.textContent = state.currentName;
    els.docBadge.textContent = state.currentExtension.toUpperCase().replace('.', '');
    els.docBadge.classList.toggle('doc-legacy', state.currentExtension === '.doc');

    const isDocx = state.currentExtension === '.docx' || state.currentExtension === '.dotx';

    if (isDocx) {
      els.docLegacyBanner.hidden = true;
      els.pageControls.hidden = false;
      els.zoomControls.hidden = false;

      // 1. Render Page Layout via docx-preview
      try {
        els.docxContainer.innerHTML = '';
        await window.docx.renderAsync(state.currentArrayBuffer, els.docxContainer, null, {
          className: 'docx',
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          ignoreLastRenderedPageBreak: false,
          experimental: true,
          trimXmlDeclaration: true,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true,
          renderEndnotes: true,
        });
      } catch (err) {
        console.warn('docx-preview notice:', err);
      }

      // 2. Render Semantic HTML via Mammoth for Flow / Edit mode
      try {
        if (window.mammoth) {
          const mammothResult = await window.mammoth.convertToHtml({ arrayBuffer: state.currentArrayBuffer });
          state.currentHtml = mammothResult.value;
          els.flowContainer.innerHTML = mammothResult.value || '<p>Empty document</p>';

          const rawTextResult = await window.mammoth.extractRawText({ arrayBuffer: state.currentArrayBuffer });
          state.currentRawText = rawTextResult.value;
        }
      } catch (err) {
        console.warn('Mammoth notice:', err);
        state.currentRawText = els.docxContainer.innerText || '';
        state.currentHtml = `<div style="white-space: pre-wrap;">${state.currentRawText}</div>`;
        els.flowContainer.innerHTML = state.currentHtml;
      }

      // 3. Extract core XML metadata and original sectPr via JSZip
      let metadata = {};
      try {
        if (window.JSZip) {
          const zip = await window.JSZip.loadAsync(state.currentArrayBuffer);
          if (zip.file('word/document.xml')) {
            state.originalDocXml = await zip.file('word/document.xml').async('text');
            const sectMatch = state.originalDocXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
            state.originalSectPr = sectMatch ? sectMatch[0] : '';
          }
          if (zip.file('docProps/core.xml')) {
            const coreXml = await zip.file('docProps/core.xml').async('text');
            const parser = new DOMParser();
            const doc = parser.parseFromString(coreXml, 'application/xml');
            metadata.title = doc.querySelector('title')?.textContent || '';
            metadata.creator = doc.querySelector('creator')?.textContent || '';
            metadata.lastModifiedBy = doc.querySelector('lastModifiedBy')?.textContent || '';
            metadata.created = doc.querySelector('created')?.textContent || '';
            metadata.modified = doc.querySelector('modified')?.textContent || '';
          }
          if (zip.file('docProps/app.xml')) {
            const appXml = await zip.file('docProps/app.xml').async('text');
            const parser = new DOMParser();
            const doc = parser.parseFromString(appXml, 'application/xml');
            metadata.pages = doc.querySelector('Pages')?.textContent || '';
            metadata.words = doc.querySelector('Words')?.textContent || '';
            metadata.paragraphs = doc.querySelector('Paragraphs')?.textContent || '';
          }
        }
      } catch (err) {
        console.warn('Metadata extraction notice:', err);
      }

      state.docMetadata = metadata;
      updateDocumentStats(fileSize, mtime);
      updatePageCount();
      setupPageObserver();
      buildDocumentOutline();

    } else {
      // Legacy .doc
      els.docLegacyBanner.hidden = false;
      els.pageControls.hidden = true;
      els.zoomControls.hidden = true;

      const extracted = extractTextFromBinaryDoc(state.currentArrayBuffer);
      state.currentRawText = extracted;
      state.currentHtml = `<div class="legacy-doc-content" style="white-space: pre-wrap;">${escapeHtml(extracted)}</div>`;

      els.docxContainer.innerHTML = '';
      els.flowContainer.innerHTML = state.currentHtml;
      switchViewMode('flow');

      updateDocumentStats(fileSize, mtime);
      buildDocumentOutline();
    }

    applyZoom();
    hideLoading();
  }

  // -------------------------------------------------------------------------
  // In-Place Editing & Saving Engine
  // -------------------------------------------------------------------------

  function toggleEditMode(forceState) {
    const next = forceState !== undefined ? forceState : !state.isEditing;
    state.isEditing = next;

    const appEl = document.getElementById('app');

    if (state.isEditing) {
      els.editToggleBtn.classList.add('active');
      els.editorRibbon.hidden = false;
      appEl.classList.add('is-editing');

      // Switch to Flow view for continuous, fluent document editing
      switchViewMode('flow');

      els.flowContainer.contentEditable = 'true';
      els.flowContainer.focus();

      els.saveStatus.textContent = state.hasUnsavedChanges ? '● Unsaved changes' : 'Saved';
      els.saveStatus.className = state.hasUnsavedChanges ? 'save-status unsaved' : 'save-status';
    } else {
      els.editToggleBtn.classList.remove('active');
      els.editorRibbon.hidden = true;
      appEl.classList.remove('is-editing');
      els.flowContainer.contentEditable = 'false';
    }
  }

  function execFormat(command, value = null) {
    if (!state.isEditing) toggleEditMode(true);
    document.execCommand(command, false, value);
    els.flowContainer.focus();
    markUnsaved();
  }

  function markUnsaved() {
    state.hasUnsavedChanges = true;
    els.saveStatus.textContent = '● Unsaved changes';
    els.saveStatus.className = 'save-status unsaved';
  }

  async function saveDocumentChanges() {
    if (!state.currentPath) return;

    els.saveStatus.textContent = 'Saving...';
    els.saveStatus.className = 'save-status saving';
    els.saveDocBtn.disabled = true;

    try {
      const isDocx = state.currentExtension === '.docx' || state.currentExtension === '.dotx';

      if (isDocx && window.JSZip && state.currentArrayBuffer) {
        // 1. Generate updated OpenXML <w:body> from edited DOM
        const editedHtml = els.flowContainer.innerHTML;
        const newBodyXml = htmlToOpenXmlBody(editedHtml, state.originalSectPr);

        // 2. Load zip archive, replace word/document.xml
        const zip = await window.JSZip.loadAsync(state.currentArrayBuffer);
        let currentDocXml = state.originalDocXml;
        if (!currentDocXml && zip.file('word/document.xml')) {
          currentDocXml = await zip.file('word/document.xml').async('text');
        }

        const updatedDocXml = currentDocXml.replace(/<w:body[\s\S]*?<\/w:body>/, newBodyXml);
        zip.file('word/document.xml', updatedDocXml);
        state.originalDocXml = updatedDocXml;

        // 3. Re-zip archive into updated binary base64
        const newBase64 = await zip.generateAsync({ type: 'base64' });
        const newArrayBuffer = base64ToArrayBuffer(newBase64);

        // 4. Save to disk via backend API
        const resp = await apiFetch('/api/save-doc', {
          path: state.currentPath,
          base64: newBase64,
          workspaceUris: state.workspaceUris,
        });

        if (!resp.success) {
          throw new Error(resp.error || 'Failed to write changes to disk');
        }

        // 5. Update cached state and re-render Page layout
        state.currentArrayBuffer = newArrayBuffer;
        state.currentRawText = els.flowContainer.innerText || '';
        state.hasUnsavedChanges = false;

        // Sync back to docxContainer in background
        try {
          els.docxContainer.innerHTML = '';
          await window.docx.renderAsync(state.currentArrayBuffer, els.docxContainer, null, {
            className: 'docx',
            inWrapper: true,
            breakPages: true,
            renderHeaders: true,
            renderFooters: true,
          });
          updatePageCount();
        } catch (err) {
          console.warn('Sync notice:', err);
        }

        updateDocumentStats(resp.size, resp.mtime);
        buildDocumentOutline();

        els.saveStatus.textContent = '✓ Saved';
        els.saveStatus.className = 'save-status';

      } else {
        // For plain text / fallback
        alert('Document saved in memory. For legacy .doc files, please export as .docx for full fidelity persistence.');
        els.saveStatus.textContent = 'Saved in memory';
        els.saveStatus.className = 'save-status';
      }
    } catch (err) {
      console.error('Save document error:', err);
      alert('Could not save document: ' + err.message);
      els.saveStatus.textContent = '⚠️ Save failed';
      els.saveStatus.className = 'save-status unsaved';
    } finally {
      els.saveDocBtn.disabled = false;
    }
  }

  // -------------------------------------------------------------------------
  // Page Count & Scroll Tracking
  // -------------------------------------------------------------------------

  function updatePageCount() {
    const pages = els.docxContainer.querySelectorAll('section.docx');
    state.totalPages = Math.max(1, pages.length || (state.docMetadata.pages ? parseInt(state.docMetadata.pages) : 1));
    els.pageCount.textContent = state.totalPages;
    els.pageNumInput.max = state.totalPages;
    els.pageNumInput.value = state.currentPage;
  }

  function setupPageObserver() {
    const pages = els.docxContainer.querySelectorAll('section.docx');
    if (!pages || pages.length === 0) return;

    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting) {
            const idx = Array.from(pages).indexOf(entry.target);
            if (idx !== -1) {
              state.currentPage = idx + 1;
              els.pageNumInput.value = state.currentPage;
            }
          }
        }
      },
      { root: els.viewerContainer, threshold: 0.3 }
    );

    pages.forEach((page) => observer.observe(page));
  }

  function scrollToPage(pageNum) {
    const pageIndex = Math.max(0, Math.min(pageNum - 1, state.totalPages - 1));
    const pages = els.docxContainer.querySelectorAll('section.docx');
    if (pages && pages[pageIndex]) {
      pages[pageIndex].scrollIntoView({ behavior: 'smooth', block: 'start' });
      state.currentPage = pageIndex + 1;
      els.pageNumInput.value = state.currentPage;
    }
  }

  // -------------------------------------------------------------------------
  // Zoom Controls
  // -------------------------------------------------------------------------

  const ZOOM_LEVELS = [0.5, 0.65, 0.75, 0.85, 1.0, 1.15, 1.25, 1.5, 1.75, 2.0];

  function applyZoom() {
    els.zoomResetBtn.textContent = `${Math.round(state.zoom * 100)}%`;
    const wrapper = els.docxContainer.querySelector('.docx-wrapper');
    if (wrapper) {
      wrapper.style.transform = `scale(${state.zoom})`;
      wrapper.style.transformOrigin = 'top center';
    }
  }

  function zoomIn() {
    const next = ZOOM_LEVELS.find((z) => z > state.zoom + 0.05);
    if (next) {
      state.zoom = next;
      applyZoom();
    }
  }

  function zoomOut() {
    const prev = [...ZOOM_LEVELS].reverse().find((z) => z < state.zoom - 0.05);
    if (prev) {
      state.zoom = prev;
      applyZoom();
    }
  }

  function resetZoom() {
    state.zoom = 1.0;
    applyZoom();
  }

  function fitWidth() {
    const firstPage = els.docxContainer.querySelector('section.docx');
    if (!firstPage) return;
    const pageWidth = firstPage.offsetWidth || 816;
    const containerWidth = els.viewerContainer.clientWidth - 40;
    if (containerWidth > 0 && pageWidth > 0) {
      state.zoom = Math.min(1.5, Math.max(0.4, parseFloat((containerWidth / pageWidth).toFixed(2))));
      applyZoom();
    }
  }

  // -------------------------------------------------------------------------
  // View Modes: Page vs Flow
  // -------------------------------------------------------------------------

  function switchViewMode(mode) {
    state.viewMode = mode;
    if (mode === 'page') {
      els.modePageBtn.classList.add('active');
      els.modeFlowBtn.classList.remove('active');
      els.docxContainer.hidden = false;
      els.flowContainer.hidden = true;
      els.pageControls.hidden = false;
      els.zoomControls.hidden = false;
      applyZoom();
    } else {
      els.modePageBtn.classList.remove('active');
      els.modeFlowBtn.classList.add('active');
      els.docxContainer.hidden = true;
      els.flowContainer.hidden = false;
      els.pageControls.hidden = true;
      els.zoomControls.hidden = true;
    }
  }

  // -------------------------------------------------------------------------
  // Document Outline (Table of Contents)
  // -------------------------------------------------------------------------

  function buildDocumentOutline() {
    els.tocList.innerHTML = '';
    const container = state.viewMode === 'page' ? els.docxContainer : els.flowContainer;
    const headings = container.querySelectorAll('h1, h2, h3, h4, h5, h6, [class*="heading-"], [class*="heading_"]');

    if (!headings || headings.length === 0) {
      els.tocList.innerHTML = '<div class="toc-empty">No headings found in this document.</div>';
      return;
    }

    headings.forEach((heading) => {
      const text = heading.textContent.trim();
      if (!text) return;

      const item = document.createElement('a');
      item.className = 'toc-item';
      const level = heading.tagName.startsWith('H') ? parseInt(heading.tagName[1]) : 2;
      item.classList.add(`toc-level-${level}`);
      item.textContent = text;
      item.title = text;

      item.addEventListener('click', (e) => {
        e.preventDefault();
        heading.scrollIntoView({ behavior: 'smooth', block: 'center' });
      });

      els.tocList.appendChild(item);
    });
  }

  // -------------------------------------------------------------------------
  // In-Document Search
  // -------------------------------------------------------------------------

  function clearSearchHighlights() {
    const marks = els.viewerContainer.querySelectorAll('.docx-search-match');
    marks.forEach((mark) => {
      const parent = mark.parentNode;
      if (parent) {
        parent.replaceChild(document.createTextNode(mark.textContent), mark);
        parent.normalize();
      }
    });
    state.searchResults = [];
    state.currentSearchIndex = -1;
    els.searchMatchCount.textContent = '0 of 0';
  }

  function performSearch(query) {
    clearSearchHighlights();
    if (!query || query.trim().length === 0) return;

    const searchTerm = query.trim().toLowerCase();
    const activeContainer = state.viewMode === 'page' ? els.docxContainer : els.flowContainer;

    const walker = document.createTreeWalker(activeContainer, NodeFilter.SHOW_TEXT, null, false);
    const textNodes = [];
    let node;
    while ((node = walker.nextNode())) {
      if (node.nodeValue && node.nodeValue.toLowerCase().includes(searchTerm)) {
        textNodes.push(node);
      }
    }

    const matches = [];
    textNodes.forEach((tNode) => {
      const val = tNode.nodeValue;
      const lower = val.toLowerCase();
      let index = lower.indexOf(searchTerm);

      if (index !== -1) {
        const frag = document.createDocumentFragment();
        let lastIdx = 0;

        while (index !== -1) {
          if (index > lastIdx) {
            frag.appendChild(document.createTextNode(val.substring(lastIdx, index)));
          }
          const mark = document.createElement('mark');
          mark.className = 'docx-search-match';
          mark.textContent = val.substring(index, index + searchTerm.length);
          frag.appendChild(mark);
          matches.push(mark);

          lastIdx = index + searchTerm.length;
          index = lower.indexOf(searchTerm, lastIdx);
        }

        if (lastIdx < val.length) {
          frag.appendChild(document.createTextNode(val.substring(lastIdx)));
        }

        if (tNode.parentNode) {
          tNode.parentNode.replaceChild(frag, tNode);
        }
      }
    });

    state.searchResults = matches;
    if (matches.length > 0) {
      state.currentSearchIndex = 0;
      updateSearchActive();
    } else {
      els.searchMatchCount.textContent = '0 matches';
    }
  }

  function updateSearchActive() {
    state.searchResults.forEach((m, i) => {
      m.classList.toggle('active', i === state.currentSearchIndex);
    });

    if (state.searchResults.length > 0) {
      const current = state.searchResults[state.currentSearchIndex];
      els.searchMatchCount.textContent = `${state.currentSearchIndex + 1} of ${state.searchResults.length}`;
      current.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }

  function nextSearchMatch() {
    if (state.searchResults.length === 0) return;
    state.currentSearchIndex = (state.currentSearchIndex + 1) % state.searchResults.length;
    updateSearchActive();
  }

  function prevSearchMatch() {
    if (state.searchResults.length === 0) return;
    state.currentSearchIndex = (state.currentSearchIndex - 1 + state.searchResults.length) % state.searchResults.length;
    updateSearchActive();
  }

  // -------------------------------------------------------------------------
  // Document Stats & Metadata Drawer
  // -------------------------------------------------------------------------

  function updateDocumentStats(fileSize, mtime) {
    const text = state.currentRawText || '';
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    const chars = text.length;
    const paras = text.split(/\n\s*\n/).filter((p) => p.trim().length > 0).length;
    const readingMinutes = Math.max(1, Math.ceil(words / 200));

    els.infoFileName.textContent = state.currentName || '-';
    els.infoFileSize.textContent = formatBytes(fileSize);
    els.infoWordCount.textContent = (state.docMetadata.words || words).toLocaleString();
    els.infoCharCount.textContent = chars.toLocaleString();
    els.infoParaCount.textContent = (state.docMetadata.paragraphs || paras).toLocaleString();
    els.infoPageCount.textContent = state.totalPages || 1;
    els.infoReadingTime.textContent = `~${readingMinutes} min`;

    els.infoAuthor.textContent = state.docMetadata.creator || '-';
    els.infoLastModifiedBy.textContent = state.docMetadata.lastModifiedBy || '-';

    const dateVal = state.docMetadata.modified ? new Date(state.docMetadata.modified) : new Date(mtime);
    els.infoModifiedDate.textContent = isNaN(dateVal.getTime()) ? '-' : dateVal.toLocaleDateString();
  }

  // -------------------------------------------------------------------------
  // Agent Prompt & Chat Integration
  // -------------------------------------------------------------------------

  async function sendPromptToAgent(promptText) {
    if (!promptText || !promptText.trim()) return;

    const fullMessage = `${promptText.trim()}\n\n[Document: ${state.currentName || 'Word Document'}](${state.currentPath ? 'file://' + state.currentPath : ''})`;

    if (window.sidecar && window.sidecar.agent) {
      try {
        if (window.sidecar.conversationId) {
          await window.sidecar.agent.sendMessage(fullMessage, window.sidecar.conversationId);
        } else {
          await window.sidecar.agent.startConversation(fullMessage);
        }
        els.agentPromptDrawer.hidden = true;
        els.agentPromptInput.value = '';
      } catch (err) {
        console.error('Agent message error:', err);
        alert('Could not deliver message to Antigravity: ' + err.message);
      }
    } else {
      alert('Antigravity agent interface is not connected.');
    }
  }

  // -------------------------------------------------------------------------
  // Workspace Discovery
  // -------------------------------------------------------------------------

  async function loadWorkspaceDocs() {
    try {
      const resp = await apiFetch('/api/workspace-docs', {
        workspaceUris: state.workspaceUris,
      });

      if (resp.docs && resp.docs.length > 0) {
        els.workspaceItems.innerHTML = '';
        resp.docs.forEach((doc) => {
          const item = document.createElement('div');
          item.className = 'workspace-item';
          const isLegacy = doc.extension === '.doc';

          item.innerHTML = `
            <div class="workspace-item-icon ${isLegacy ? 'doc-legacy' : ''}">W</div>
            <div class="workspace-item-info">
              <div class="workspace-item-name">${escapeHtml(doc.name)}</div>
              <div class="workspace-item-meta">${formatBytes(doc.size)} • ${escapeHtml(doc.path)}</div>
            </div>
          `;

          item.addEventListener('click', () => {
            openDocumentFile(doc.path);
          });

          els.workspaceItems.appendChild(item);
        });
        els.workspaceList.hidden = false;
      } else {
        els.workspaceList.hidden = true;
      }
    } catch (err) {
      console.warn('Could not load workspace documents:', err);
    }
  }

  // -------------------------------------------------------------------------
  // Event Listeners
  // -------------------------------------------------------------------------

  function setupEventListeners() {
    // File open buttons
    const triggerPicker = () => els.fileInput.click();
    els.toolbarOpenBtn.addEventListener('click', triggerPicker);
    els.emptyOpenBtn.addEventListener('click', triggerPicker);

    els.fileInput.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;

      showLoading(`Opening ${file.name}...`);
      state.currentName = file.name;
      state.currentPath = file.name;
      state.currentExtension = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();

      const reader = new FileReader();
      reader.onload = async (evt) => {
        state.currentArrayBuffer = evt.target.result;
        await renderLoadedDocument(file.size, file.lastModified);
      };
      reader.onerror = () => showError('Failed to read local file');
      reader.readAsArrayBuffer(file);
    });

    // Drag and Drop
    const dz = els.dropzone;
    ['dragenter', 'dragover'].forEach((eventName) => {
      dz.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dz.classList.add('drag-over');
      });
    });

    ['dragleave', 'drop'].forEach((eventName) => {
      dz.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dz.classList.remove('drag-over');
      });
    });

    dz.addEventListener('drop', (e) => {
      const file = e.dataTransfer.files[0];
      if (file) {
        showLoading(`Opening ${file.name}...`);
        state.currentName = file.name;
        state.currentPath = file.name;
        state.currentExtension = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();

        const reader = new FileReader();
        reader.onload = async (evt) => {
          state.currentArrayBuffer = evt.target.result;
          await renderLoadedDocument(file.size, file.lastModified);
        };
        reader.onerror = () => showError('Failed to read dropped file');
        reader.readAsArrayBuffer(file);
      }
    });

    // View Mode Switcher
    els.modePageBtn.addEventListener('click', () => switchViewMode('page'));
    els.modeFlowBtn.addEventListener('click', () => switchViewMode('flow'));

    // Edit Mode Toggles & Ribbon Buttons
    els.editToggleBtn.addEventListener('click', () => toggleEditMode());
    els.doneEditBtn.addEventListener('click', () => toggleEditMode(false));
    els.saveDocBtn.addEventListener('click', () => saveDocumentChanges());

    // Content Editable Input Listener (Tracks unsaved changes)
    els.flowContainer.addEventListener('input', () => {
      markUnsaved();
    });

    // Formatting Toolbar Actions
    els.boldBtn.addEventListener('click', () => execFormat('bold'));
    els.italicBtn.addEventListener('click', () => execFormat('italic'));
    els.underlineBtn.addEventListener('click', () => execFormat('underline'));
    els.strikeBtn.addEventListener('click', () => execFormat('strikeThrough'));
    els.bulletListBtn.addEventListener('click', () => execFormat('insertUnorderedList'));
    els.numListBtn.addEventListener('click', () => execFormat('insertOrderedList'));
    els.alignLeftBtn.addEventListener('click', () => execFormat('justifyLeft'));
    els.alignCenterBtn.addEventListener('click', () => execFormat('justifyCenter'));
    els.alignRightBtn.addEventListener('click', () => execFormat('justifyRight'));
    els.undoBtn.addEventListener('click', () => execFormat('undo'));
    els.redoBtn.addEventListener('click', () => execFormat('redo'));

    els.formatBlockSelect.addEventListener('change', (e) => {
      execFormat('formatBlock', e.target.value);
    });

    // Page controls
    els.prevPageBtn.addEventListener('click', () => scrollToPage(state.currentPage - 1));
    els.nextPageBtn.addEventListener('click', () => scrollToPage(state.currentPage + 1));
    els.pageNumInput.addEventListener('change', (e) => {
      const val = parseInt(e.target.value);
      if (!isNaN(val)) scrollToPage(val);
    });

    // Zoom controls
    els.zoomInBtn.addEventListener('click', zoomIn);
    els.zoomOutBtn.addEventListener('click', zoomOut);
    els.zoomResetBtn.addEventListener('click', resetZoom);
    els.fitWidthBtn.addEventListener('click', fitWidth);

    // Outline / TOC toggle
    els.tocToggleBtn.addEventListener('click', () => {
      els.tocSidebar.hidden = !els.tocSidebar.hidden;
      if (!els.tocSidebar.hidden) buildDocumentOutline();
    });
    els.closeTocBtn.addEventListener('click', () => (els.tocSidebar.hidden = true));

    // In-document search
    els.searchToggleBtn.addEventListener('click', () => {
      els.searchDrawer.hidden = !els.searchDrawer.hidden;
      if (!els.searchDrawer.hidden) {
        els.searchInput.focus();
        els.searchInput.select();
      } else {
        clearSearchHighlights();
      }
    });

    els.closeSearchBtn.addEventListener('click', () => {
      els.searchDrawer.hidden = true;
      clearSearchHighlights();
    });

    els.searchInput.addEventListener('input', (e) => performSearch(e.target.value));
    els.searchNextBtn.addEventListener('click', nextSearchMatch);
    els.searchPrevBtn.addEventListener('click', prevSearchMatch);
    els.searchInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        if (e.shiftKey) prevSearchMatch();
        else nextSearchMatch();
      } else if (e.key === 'Escape') {
        els.searchDrawer.hidden = true;
        clearSearchHighlights();
      }
    });

    // Info Drawer
    els.docInfoBtn.addEventListener('click', () => {
      els.docInfoDrawer.hidden = !els.docInfoDrawer.hidden;
    });
    els.closeInfoDrawerBtn.addEventListener('click', () => (els.docInfoDrawer.hidden = true));

    // Info Actions
    els.copyTextBtn.addEventListener('click', () => {
      navigator.clipboard.writeText(state.currentRawText || '');
      const prev = els.copyTextBtn.textContent;
      els.copyTextBtn.textContent = 'Copied!';
      setTimeout(() => (els.copyTextBtn.textContent = prev), 1500);
    });

    els.copyMarkdownBtn.addEventListener('click', () => {
      const md = (state.currentRawText || '').replace(/\n{2,}/g, '\n\n');
      navigator.clipboard.writeText(md);
      const prev = els.copyMarkdownBtn.textContent;
      els.copyMarkdownBtn.textContent = 'Copied Markdown!';
      setTimeout(() => (els.copyMarkdownBtn.textContent = prev), 1500);
    });

    els.printDocBtn.addEventListener('click', () => window.print());

    // Agent Prompt Drawer
    els.askAgentBtn.addEventListener('click', () => {
      els.agentPromptDrawer.hidden = !els.agentPromptDrawer.hidden;
      if (!els.agentPromptDrawer.hidden) els.agentPromptInput.focus();
    });
    els.closePromptDrawerBtn.addEventListener('click', () => (els.agentPromptDrawer.hidden = true));

    // Quick chips
    document.querySelectorAll('.chip').forEach((chip) => {
      chip.addEventListener('click', () => {
        const text = chip.getAttribute('data-prompt');
        sendPromptToAgent(text);
      });
    });

    // Submit Prompt
    els.submitPromptBtn.addEventListener('click', () => {
      sendPromptToAgent(els.agentPromptInput.value);
    });

    els.agentPromptInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        sendPromptToAgent(els.agentPromptInput.value);
      }
    });

    // Refresh workspace docs
    els.refreshWorkspaceBtn.addEventListener('click', loadWorkspaceDocs);

    // Global keyboard shortcuts
    window.addEventListener('keydown', (e) => {
      // Cmd/Ctrl + S to Save
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        if (state.currentPath) {
          saveDocumentChanges();
        }
      }
      // Cmd/Ctrl + E to Toggle Edit Mode
      else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'e') {
        e.preventDefault();
        if (state.currentPath) {
          toggleEditMode();
        }
      }
      // Cmd/Ctrl + F to Search
      else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
        e.preventDefault();
        els.searchDrawer.hidden = false;
        els.searchInput.focus();
        els.searchInput.select();
      } else if (e.key === 'Escape') {
        els.searchDrawer.hidden = true;
        els.agentPromptDrawer.hidden = true;
        els.docInfoDrawer.hidden = true;
        clearSearchHighlights();
      } else if (e.key === 'ArrowLeft' && state.viewMode === 'page' && !state.isEditing && !e.target.matches('input, textarea')) {
        scrollToPage(state.currentPage - 1);
      } else if (e.key === 'ArrowRight' && state.viewMode === 'page' && !state.isEditing && !e.target.matches('input, textarea')) {
        scrollToPage(state.currentPage + 1);
      }
    });

    // Window resize -> adjust zoom if needed
    window.addEventListener('resize', () => {
      if (state.viewMode === 'page') {
        // keep zoom stable
      }
    });
  }

  // -------------------------------------------------------------------------
  // Initialization
  // -------------------------------------------------------------------------

  async function init() {
    setupEventListeners();

    // Check workspace URIs
    if (window.sidecar && typeof window.sidecar.getWorkspaceUris === 'function') {
      try {
        state.workspaceUris = await window.sidecar.getWorkspaceUris();
      } catch (err) {
        console.warn('Workspace URIs error:', err);
      }

      if (typeof window.sidecar.onWorkspaceChange === 'function') {
        window.sidecar.onWorkspaceChange((uris) => {
          state.workspaceUris = uris || [];
          loadWorkspaceDocs();
        });
      }
    }

    // Check URL parameters for ?file=
    const params = new URLSearchParams(window.location.search);
    const fileParam = params.get('file');

    if (fileParam) {
      await openDocumentFile(fileParam);
    } else {
      // Check last opened file state
      try {
        const stateResp = await apiFetch('/api/get-state', {}, 'GET');
        if (stateResp && stateResp.lastOpenedPath) {
          await openDocumentFile(stateResp.lastOpenedPath);
        } else {
          loadWorkspaceDocs();
        }
      } catch {
        loadWorkspaceDocs();
      }
    }
  }

  // Run on DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
