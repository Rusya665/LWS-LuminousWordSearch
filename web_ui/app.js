// LWS 2.0 Web UI Logic - Communicates with Python backend via window.pywebview.api

let state = {
  useSynonyms: true,
  isRegex: false,
  activeRole: 'ALL',
  activeCategory: 'ALL',
  selectedDocId: null,
  selectedBundleId: null,
  documents: [],
  watchFolders: [],
  folderStats: [],
  categories: {},
  currentBundleDocs: [],
  currentSiDoc: null,
  figures: [],
  latencyMs: 0
};

// DOM Elements
const searchInput = document.getElementById('searchInput');
const toggleSynonymBtn = document.getElementById('toggleSynonymBtn');
const toggleRegexBtn = document.getElementById('toggleRegexBtn');
const searchLatency = document.getElementById('searchLatency');

// Sidebar
const folderList = document.getElementById('folderList');
const categoryTree = document.getElementById('categoryTree');
const catCountBadge = document.getElementById('catCountBadge');
const documentList = document.getElementById('documentList');

// Center Tabs
const sentenceResults = document.getElementById('sentenceResults');
const hitCountBadge = document.getElementById('hitCountBadge');
const figCountBadge = document.getElementById('figCountBadge');
const figuresGrid = document.getElementById('figuresGrid');
const mainReaderTitle = document.getElementById('mainReaderTitle');
const mainReaderBody = document.getElementById('mainReaderBody');
const siReaderTitle = document.getElementById('siReaderTitle');
const siReaderBody = document.getElementById('siReaderBody');
const openMainNativeBtn = document.getElementById('openMainNativeBtn');
const openSiNativeBtn = document.getElementById('openSiNativeBtn');
const siDocSelect = document.getElementById('siDocSelect');

// Right Panel
const metaTitle = document.getElementById('metaTitle');
const metaAuthors = document.getElementById('metaAuthors');
const metaCategoryFolder = document.getElementById('metaCategoryFolder');
const metaYearDoi = document.getElementById('metaYearDoi');
const metaBundle = document.getElementById('metaBundle');
const synonymCloud = document.getElementById('synonymCloud');
const notesInput = document.getElementById('notesInput');
const saveNoteBtn = document.getElementById('saveNoteBtn');

// Modals
const manageFoldersBtn = document.getElementById('manageFoldersBtn');
const browseQuickFolderBtn = document.getElementById('browseQuickFolderBtn');
const folderModal = document.getElementById('folderModal');
const closeFolderModalBtn = document.getElementById('closeFolderModalBtn');
const closeFolderModalX = document.getElementById('closeFolderModalX');
const modalBrowseBtn = document.getElementById('modalBrowseBtn');
const modalAddPathBtn = document.getElementById('modalAddPathBtn');
const manualFolderPath = document.getElementById('manualFolderPath');
const modalFoldersList = document.getElementById('modalFoldersList');
const modalSyncNowBtn = document.getElementById('modalSyncNowBtn');

const syncBtn = document.getElementById('syncBtn');
const syncModal = document.getElementById('syncModal');
const syncStatus = document.getElementById('syncStatus');
const syncProgressBar = document.getElementById('syncProgressBar');
const closeSyncModalBtn = document.getElementById('closeSyncModalBtn');

function getApi() {
  return window.pywebview ? window.pywebview.api : null;
}

window.addEventListener('pywebviewready', () => {
  initApp();
});

document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    if (!window.pywebview) {
      initApp();
    }
  }, 250);
});

async function initApp() {
  setupEventListeners();
  await loadState();
}

function setupEventListeners() {
  // Tabs
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      const targetId = btn.dataset.tab;
      document.getElementById(targetId).classList.add('active');
    });
  });

  // Role filters
  document.querySelectorAll('.role-tag').forEach(tag => {
    tag.addEventListener('click', () => {
      document.querySelectorAll('.role-tag').forEach(t => t.classList.remove('active'));
      tag.classList.add('active');
      state.activeRole = tag.dataset.role;
      filterAndRenderDocs();
      performSearch();
    });
  });

  // Synonyms toggle
  toggleSynonymBtn.addEventListener('click', () => {
    state.useSynonyms = !state.useSynonyms;
    toggleSynonymBtn.classList.toggle('active', state.useSynonyms);
    toggleSynonymBtn.textContent = state.useSynonyms ? 'Syn: ON' : 'Syn: OFF';
    performSearch();
  });

  // Regex toggle
  if (toggleRegexBtn) {
    toggleRegexBtn.addEventListener('click', () => {
      state.isRegex = !state.isRegex;
      toggleRegexBtn.classList.toggle('active', state.isRegex);
      performSearch();
    });
  }

  // Ctrl+F global keyboard shortcut
  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
      e.preventDefault();
      if (searchInput) {
        searchInput.focus();
        searchInput.select();
      }
    }
  });

  // Search input with debounce
  let searchTimer = null;
  searchInput.addEventListener('input', () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(performSearch, 150);
  });

  // Folder Modals
  manageFoldersBtn.addEventListener('click', () => openFolderModal());
  browseQuickFolderBtn.addEventListener('click', () => triggerNativeFolderPicker());
  closeFolderModalBtn.addEventListener('click', () => closeFolderModal());
  closeFolderModalX.addEventListener('click', () => closeFolderModal());

  modalBrowseBtn.addEventListener('click', () => triggerNativeFolderPicker());
  modalAddPathBtn.addEventListener('click', () => handleManualAddFolder());
  modalSyncNowBtn.addEventListener('click', () => {
    closeFolderModal();
    startSync();
  });

  // Sync
  syncBtn.addEventListener('click', startSync);
  closeSyncModalBtn.addEventListener('click', () => {
    syncModal.classList.remove('active');
  });

  // Notes & file open
  saveNoteBtn.addEventListener('click', saveNote);
  openMainNativeBtn.addEventListener('click', () => {
    if (!Array.isArray(state.currentBundleDocs) || state.currentBundleDocs.length === 0) return;
    const mainDoc = state.currentBundleDocs.find(d => d.doc_role === 'MAIN') || state.currentBundleDocs[0];
    if (mainDoc && getApi()) {
      getApi().open_native_file(mainDoc.watch_folder + '/' + mainDoc.rel_path);
    }
  });

  openSiNativeBtn.addEventListener('click', () => {
    const targetDoc = state.currentSiDoc || (state.currentBundleDocs && state.currentBundleDocs.find(d => d.doc_role === 'SI'));
    if (targetDoc && getApi()) {
      getApi().open_native_file(targetDoc.watch_folder + '/' + targetDoc.rel_path);
    }
  });

  if (siDocSelect) {
    siDocSelect.addEventListener('change', async () => {
      const selectedId = parseInt(siDocSelect.value, 10);
      if (!selectedId) {
        state.currentSiDoc = null;
        siReaderBody.innerHTML = '<p style="color: var(--text-dim); text-align: center; margin-top: 40px;">No document selected. Choose a paper from the dropdown above to compare side-by-side.</p>';
        return;
      }
      const chosenDoc = state.documents.find(d => d.id === selectedId);
      if (chosenDoc) {
        state.currentSiDoc = chosenDoc;
        const api = getApi();
        if (api) {
          siReaderBody.innerHTML = '<p style="color: var(--text-dim); text-align: center; margin-top: 40px;">Loading document text...</p>';
          try {
            const bundle = await api.get_document_bundle(chosenDoc.bundle_id || chosenDoc.rel_path);
            const text = (bundle.contents && bundle.contents[chosenDoc.id]) || '(Text empty or not cached)';
            siReaderBody.innerHTML = `<pre style="white-space: pre-wrap; font-family: inherit;">${escapeHtml(text)}</pre>`;
          } catch (e) {
            siReaderBody.innerHTML = `<p style="color: var(--accent-pink); text-align: center; margin-top: 40px;">Error loading text: ${escapeHtml(String(e))}</p>`;
          }
        }
      }
    });
  }
}

async function loadState() {
  const api = getApi();
  if (api) {
    try {
      const resp = await api.get_workspace_state();
      state.watchFolders = resp.watch_folders || [];
      state.folderStats = resp.folder_stats || [];
      state.documents = resp.documents || [];

      renderSidebarFolders();
      buildHierarchicalCategories();
      filterAndRenderDocs();

      if (state.documents.length > 0) {
        selectDocument(state.documents[0]);
      }
      loadFigures("");
    } catch (e) {
      console.error("Failed to load workspace state", e);
    }
  }
}

function renderSidebarFolders() {
  folderList.innerHTML = '';
  if (state.watchFolders.length === 0) {
    folderList.innerHTML = '<div style="color: var(--text-dim); font-size: 11px; padding: 4px;">No watched folders. Click "+ Browse"</div>';
    return;
  }

  state.watchFolders.forEach(f => {
    const div = document.createElement('div');
    div.className = 'folder-item';
    const baseName = f.split(/[\\/]/).filter(Boolean).pop() || f;
    const stat = state.folderStats.find(s => s.path === f);
    const count = stat ? stat.doc_count : 0;

    div.innerHTML = `
      <span style="overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">📁 ${escapeHtml(baseName)}</span>
      <span style="margin-left: auto; font-size: 10px; color: var(--accent-indigo);">${count}</span>
    `;
    div.title = f;
    div.addEventListener('click', () => openFolderModal());
    folderList.appendChild(div);
  });
}

function buildHierarchicalCategories() {
  // Aggregate nested paths
  const catCounts = {};
  state.documents.forEach(d => {
    const fullCat = d.rel_path.includes('/') ? d.rel_path.substring(0, d.rel_path.lastIndexOf('/')) : 'Root';
    catCounts[fullCat] = (catCounts[fullCat] || 0) + 1;

    // Also increment parent categories if nested (e.g. A/B increments A)
    const segments = fullCat.split('/');
    if (segments.length > 1) {
      let cur = '';
      for (let i = 0; i < segments.length - 1; i++) {
        cur = cur ? cur + '/' + segments[i] : segments[i];
        catCounts[cur] = (catCounts[cur] || 0) + 1;
      }
    }
  });

  categoryTree.innerHTML = '';

  // "All Categories" option
  const allItem = document.createElement('div');
  allItem.className = `folder-item ${state.activeCategory === 'ALL' ? 'active' : ''}`;
  allItem.innerHTML = `<span>📂 All Categories</span> <span style="margin-left: auto; font-size: 10px; color: var(--text-dim);">${state.documents.length}</span>`;
  allItem.addEventListener('click', () => {
    document.querySelectorAll('#categoryTree .folder-item').forEach(el => el.classList.remove('active'));
    allItem.classList.add('active');
    state.activeCategory = 'ALL';
    catCountBadge.textContent = 'All';
    filterAndRenderDocs();
  });
  categoryTree.appendChild(allItem);

  // Sort categories alphabetically
  const sortedCats = Object.keys(catCounts).sort();
  sortedCats.forEach(cat => {
    const item = document.createElement('div');
    item.className = `folder-item ${state.activeCategory === cat ? 'active' : ''}`;

    // Compute indentation depth for nested subfolders
    const depth = (cat.match(/\//g) || []).length;
    const displayName = cat.split('/').pop() || cat;
    const indentPx = depth * 12 + 6;

    item.style.paddingLeft = `${indentPx}px`;
    item.innerHTML = `
      <span style="overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">📁 ${escapeHtml(displayName)}</span>
      <span style="margin-left: auto; font-size: 10px; color: var(--accent-cyan);">${catCounts[cat]}</span>
    `;
    item.title = cat;
    item.addEventListener('click', () => {
      document.querySelectorAll('#categoryTree .folder-item').forEach(el => el.classList.remove('active'));
      item.classList.add('active');
      state.activeCategory = cat;
      catCountBadge.textContent = displayName;
      filterAndRenderDocs();
    });
    categoryTree.appendChild(item);
  });
}

function filterAndRenderDocs() {
  const filtered = state.documents.filter(d => {
    const matchesRole = state.activeRole === 'ALL' || d.doc_role === state.activeRole;

    let matchesCat = true;
    if (state.activeCategory !== 'ALL') {
      const fullCat = d.rel_path.includes('/') ? d.rel_path.substring(0, d.rel_path.lastIndexOf('/')) : 'Root';
      matchesCat = (fullCat === state.activeCategory) || fullCat.startsWith(state.activeCategory + '/');
    }

    return matchesRole && matchesCat;
  });

  documentList.innerHTML = '';
  if (filtered.length === 0) {
    documentList.innerHTML = '<div style="color: var(--text-dim); text-align: center; margin-top: 24px; font-size: 11px;">No documents match filter.</div>';
    return;
  }

  filtered.forEach(d => {
    const card = document.createElement('div');
    card.className = `glass-card doc-card ${state.selectedDocId === d.id ? 'active' : ''}`;

    // Extract subfolder category breadcrumb
    const subfolder = d.rel_path.includes('/') ? d.rel_path.substring(0, d.rel_path.lastIndexOf('/')) : '';

    card.innerHTML = `
      <div class="doc-card-title">${escapeHtml(d.title || d.filename)}</div>
      <div class="doc-card-meta">
        <span class="role-badge ${d.doc_role}">${d.doc_role}</span>
        ${subfolder ? `<span style="color: var(--accent-cyan); font-size: 10px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 130px;">📁 ${escapeHtml(subfolder)}</span>` : ''}
        <span>${d.year ? d.year : ''}</span>
        <span>${d.page_count ? d.page_count + 'p' : ''}</span>
      </div>
    `;
    card.addEventListener('click', () => selectDocument(d));
    documentList.appendChild(card);
  });
}

async function selectDocument(doc) {
  state.selectedDocId = doc.id;
  state.selectedBundleId = doc.bundle_id;
  filterAndRenderDocs();

  // Populate Right detail pane
  metaTitle.textContent = doc.title || doc.filename;
  metaAuthors.textContent = doc.authors || 'Unknown Authors';

  const folderCategory = doc.rel_path.includes('/') ? doc.rel_path.substring(0, doc.rel_path.lastIndexOf('/')) : 'Root';
  metaCategoryFolder.textContent = `📁 Folder: ${folderCategory}`;
  metaYearDoi.textContent = `${doc.year ? doc.year + ' • ' : ''}${doc.doi ? 'DOI: ' + doc.doi : 'No DOI'}`;
  metaBundle.textContent = `Bundle: ${doc.bundle_id || 'single'}`;

  // Fetch bundle documents & annotations
  const api = getApi();
  if (api) {
    try {
      const bundle = await api.get_document_bundle(doc.bundle_id || doc.rel_path);
      state.currentBundleDocs = bundle.documents || [doc];
      renderSplitReader(bundle.documents || [doc], bundle.contents || {});
      notesInput.value = bundle.notes || '';
    } catch (e) {
      console.error(e);
    }
  }
}

function renderSplitReader(docs, contents) {
  const mainDoc = docs.find(d => d.doc_role === 'MAIN') || docs[0];
  const siDoc = docs.find(d => d.doc_role === 'SI');
  state.currentSiDoc = siDoc || null;

  if (mainDoc) {
    mainReaderTitle.textContent = `Main: ${mainDoc.title || mainDoc.filename}`;
    const mainTxt = contents[mainDoc.id] || '(Text not loaded or empty)';
    mainReaderBody.innerHTML = `<pre style="white-space: pre-wrap; font-family: inherit;">${escapeHtml(mainTxt)}</pre>`;
  } else {
    mainReaderTitle.textContent = 'Main Document';
    mainReaderBody.innerHTML = '<p style="color: var(--text-dim); text-align: center; margin-top: 40px;">No main paper found.</p>';
  }

  // Populate SI / Comparison dropdown
  if (siDocSelect) {
    siDocSelect.innerHTML = '';
    if (siDoc) {
      const opt = document.createElement('option');
      opt.value = siDoc.id;
      opt.textContent = `[Companion SI] ${siDoc.filename}`;
      opt.selected = true;
      siDocSelect.appendChild(opt);
    } else {
      const opt = document.createElement('option');
      opt.value = '';
      opt.textContent = '-- Select paper to compare side-by-side --';
      siDocSelect.appendChild(opt);
    }

    state.documents.forEach(d => {
      if (mainDoc && d.id === mainDoc.id) return;
      if (siDoc && d.id === siDoc.id) return;
      const opt = document.createElement('option');
      opt.value = d.id;
      opt.textContent = `[${d.doc_role}] ${d.title || d.filename}`;
      siDocSelect.appendChild(opt);
    });
  }

  if (siDoc) {
    siReaderTitle.textContent = 'SI / Compare:';
    const siTxt = contents[siDoc.id] || '(Supplementary text empty)';
    siReaderBody.innerHTML = `<pre style="white-space: pre-wrap; font-family: inherit;">${escapeHtml(siTxt)}</pre>`;
  } else {
    siReaderTitle.textContent = 'SI / Compare:';
    siReaderBody.innerHTML = '<p style="color: var(--text-dim); text-align: center; margin-top: 40px;">No companion SI file linked to this bundle.<br><br>Choose any document from the dropdown above to compare side-by-side.</p>';
  }
}

async function performSearch() {
  const query = searchInput.value.trim();
  const api = getApi();
  if (!api) return;

  const t0 = performance.now();
  try {
    const resp = await api.search_query(query, state.useSynonyms, state.activeRole, state.isRegex);
    const latency = Math.round(performance.now() - t0);
    searchLatency.textContent = `${latency} ms`;

    // Render synonyms
    synonymCloud.innerHTML = '';
    if (resp.synonyms) {
      for (const [term, syns] of Object.entries(resp.synonyms)) {
        syns.slice(0, 6).forEach(s => {
          const pill = document.createElement('span');
          pill.className = 'syn-pill';
          pill.textContent = s;
          pill.addEventListener('click', () => {
            searchInput.value = s;
            performSearch();
          });
          synonymCloud.appendChild(pill);
        });
      }
    }

    renderSentenceHits(resp.hits || []);
    hitCountBadge.textContent = resp.hits ? resp.hits.length : 0;
    loadFigures(query);

  } catch (e) {
    console.error("Search failed", e);
  }
}

function renderSentenceHits(hits) {
  sentenceResults.innerHTML = '';
  if (hits.length === 0) {
    sentenceResults.innerHTML = '<p style="color: var(--text-dim); text-align: center; margin-top: 40px;">No matching sentences found.</p>';
    return;
  }

  hits.forEach(hit => {
    const card = document.createElement('div');
    card.className = 'glass-card hit-card';
    card.style.cursor = 'pointer';
    const subfolder = hit.rel_path && hit.rel_path.includes('/') ? hit.rel_path.substring(0, hit.rel_path.lastIndexOf('/')) : '';
    const fullDocPath = (hit.watch_folder ? hit.watch_folder + '/' : '') + (hit.rel_path || '');

    card.innerHTML = `
      <div class="hit-header">
        <div class="hit-source">
          <span class="role-badge ${hit.doc_role}">${hit.doc_role}</span>
          ${subfolder ? `<span style="color: var(--accent-cyan); font-weight: 500;">[${escapeHtml(subfolder)}]</span>` : ''}
          <span>${escapeHtml(hit.title || hit.filename)}</span>
          <span style="color: var(--text-dim);">p. ${hit.page_number}</span>
        </div>
        <div class="hit-actions" style="display: flex; gap: 6px; align-items: center;">
          <button class="btn btn-secondary btn-hit-open-pdf" title="Open source document in default viewer (Acrobat/Word)" style="padding: 2px 8px; font-size: 11px; display: inline-flex; align-items: center; gap: 4px;">
            📄 Open File
          </button>
          <button class="btn btn-secondary btn-hit-split-view" title="Open side-by-side in Split Paper + SI Reader" style="padding: 2px 8px; font-size: 11px; background: rgba(99, 102, 241, 0.15); border-color: var(--accent-indigo); display: inline-flex; align-items: center; gap: 4px;">
            📖 Split Reader
          </button>
        </div>
      </div>
      <div class="hit-text">${hit.highlighted_snippet || escapeHtml(hit.sentence_text)}</div>
    `;

    const openPdfBtn = card.querySelector('.btn-hit-open-pdf');
    openPdfBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const api = getApi();
      if (api && fullDocPath) {
        api.open_native_file(fullDocPath);
      }
    });

    const openReaderAction = async () => {
      const doc = state.documents.find(d => d.id === hit.doc_id);
      if (doc) {
        await selectDocument(doc);
        const readerTabBtn = document.querySelector('[data-tab="tab-reader"]');
        if (readerTabBtn) {
          readerTabBtn.click();
        }
      }
    };

    const splitViewBtn = card.querySelector('.btn-hit-split-view');
    splitViewBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      openReaderAction();
    });

    card.addEventListener('click', openReaderAction);
    sentenceResults.appendChild(card);
  });
}

async function loadFigures(query) {
  const api = getApi();
  if (!api) return;
  try {
    const figs = await api.get_figures(query);
    state.figures = figs || [];
    figCountBadge.textContent = state.figures.length;
    renderFigures();
  } catch (e) {
    console.error(e);
  }
}

function renderFigures() {
  figuresGrid.innerHTML = '';
  if (state.figures.length === 0) {
    figuresGrid.innerHTML = '<p style="color: var(--text-dim); text-align: center; grid-column: 1 / -1; margin-top: 40px;">No extracted figures found. Run Sync to harvest figures from watched papers.</p>';
    return;
  }

  state.figures.forEach(fig => {
    const card = document.createElement('div');
    card.className = 'glass-card figure-card';
    card.innerHTML = `
      <div class="figure-img-wrap" style="cursor: zoom-in;" title="Click to view full image in default viewer">
        <img src="${fig.data_url || fig.rel_image_path}" alt="Figure" loading="lazy" />
      </div>
      <div class="figure-info">
        <div style="font-weight: 700; color: #FFFFFF; display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
          <span style="overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 170px;" title="${escapeHtml(fig.title || fig.filename)}">
            ${escapeHtml(fig.title || fig.filename)}
          </span>
          <span style="color: var(--accent-indigo); font-size: 11px;">p. ${fig.page_number}</span>
        </div>
        <div class="figure-caption" title="${escapeHtml(fig.caption || '')}">${fig.caption_snippet || escapeHtml(fig.caption || '(No caption extracted)')}</div>
        <div style="display: flex; gap: 6px; margin-top: 8px;">
          <button class="btn btn-secondary btn-fig-view" style="flex: 1; padding: 3px 6px; font-size: 10px; justify-content: center;" title="Open image in system photo viewer">
            🔍 View Image
          </button>
          <button class="btn btn-secondary btn-fig-doc" style="flex: 1; padding: 3px 6px; font-size: 10px; justify-content: center; background: rgba(99, 102, 241, 0.15); border-color: var(--accent-indigo);" title="Open source PDF in native viewer">
            📄 Open Paper
          </button>
        </div>
      </div>
    `;

    const imgWrap = card.querySelector('.figure-img-wrap');
    const viewBtn = card.querySelector('.btn-fig-view');
    const docBtn = card.querySelector('.btn-fig-doc');
    const imgEl = card.querySelector('img');

    // On-demand fetch if data_url wasn't bundled
    if (!fig.data_url && getApi()) {
      getApi().get_figure_data(fig.id).then(res => {
        if (res && res.data_url) {
          imgEl.src = res.data_url;
          fig.data_url = res.data_url;
        }
      }).catch(() => {});
    }

    const openImageAction = () => {
      const api = getApi();
      if (api && fig.full_image_path) {
        api.open_native_file(fig.full_image_path);
      }
    };

    imgWrap.addEventListener('click', openImageAction);
    viewBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      openImageAction();
    });

    docBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const api = getApi();
      if (api && fig.full_doc_path) {
        api.open_native_file(fig.full_doc_path);
      }
    });

    figuresGrid.appendChild(card);
  });
}

// FOLDER MANAGEMENT MODAL
function openFolderModal() {
  renderModalFolders();
  folderModal.classList.add('active');
}

function closeFolderModal() {
  folderModal.classList.remove('active');
}

function renderModalFolders() {
  modalFoldersList.innerHTML = '';
  if (state.watchFolders.length === 0) {
    modalFoldersList.innerHTML = '<div style="color: var(--text-dim); font-size: 12px; padding: 8px;">No root directories currently watched.</div>';
    return;
  }

  state.watchFolders.forEach(folder => {
    const row = document.createElement('div');
    row.style.display = 'flex';
    row.style.alignItems = 'center';
    row.style.justifyContent = 'space-between';
    row.style.padding = '8px 12px';
    row.style.background = 'rgba(255, 255, 255, 0.04)';
    row.style.borderRadius = '6px';
    row.style.border = '1px solid var(--border-subtle)';

    const stat = state.folderStats.find(s => s.path === folder);
    const docCount = stat ? stat.doc_count : 0;

    row.innerHTML = `
      <div style="overflow: hidden; text-overflow: ellipsis; white-space: nowrap; margin-right: 12px;">
        <div style="font-weight: 600; font-size: 12px; color: #FFFFFF;">${escapeHtml(folder)}</div>
        <div style="font-size: 11px; color: var(--text-dim);">${docCount} indexed papers &bull; Recursive scanning active</div>
      </div>
      <button class="btn btn-secondary" style="color: var(--accent-pink); padding: 4px 8px; font-size: 11px;" title="Remove this directory">
        🗑️ Remove
      </button>
    `;

    const removeBtn = row.querySelector('button');
    removeBtn.addEventListener('click', async () => {
      if (getApi()) {
        await getApi().remove_watch_folder(folder);
        await loadState();
        renderModalFolders();
      }
    });

    modalFoldersList.appendChild(row);
  });
}

async function triggerNativeFolderPicker() {
  const api = getApi();
  if (api) {
    try {
      const resp = await api.choose_folder_dialog();
      if (resp && resp.added_folders && resp.added_folders.length > 0) {
        await loadState();
        renderModalFolders();
      }
    } catch (e) {
      console.error(e);
    }
  }
}

async function handleManualAddFolder() {
  const p = manualFolderPath.value.trim();
  if (!p) return;
  const api = getApi();
  if (api) {
    try {
      const res = await api.add_watch_folder(p);
      if (res.status === 'ok') {
        manualFolderPath.value = '';
        await loadState();
        renderModalFolders();
      } else {
        alert(res.message || "Failed to add folder");
      }
    } catch (e) {
      alert("Error: " + e);
    }
  }
}

// SAVE NOTE
async function saveNote() {
  const api = getApi();
  if (!api || !state.selectedDocId) return;
  const note = notesInput.value;
  try {
    await api.save_document_note(state.selectedDocId, note);
    saveNoteBtn.textContent = '✅ Saved';
    setTimeout(() => { saveNoteBtn.textContent = '💾 Save Annotation'; }, 1500);
  } catch (e) {
    console.error(e);
  }
}

// SYNC
async function startSync() {
  const api = getApi();
  if (!api) return;
  syncModal.classList.add('active');
  syncStatus.textContent = 'Scanning and indexing all watched folders recursively...';
  syncProgressBar.style.width = '20%';
  closeSyncModalBtn.style.display = 'none';

  try {
    const res = await api.sync_folders();
    syncProgressBar.style.width = '100%';
    syncStatus.textContent = `Completed! ${res.processed} updated, ${res.skipped} up to date, ${res.errors} errors.`;
    closeSyncModalBtn.style.display = 'block';
    await loadState();
  } catch (e) {
    syncStatus.textContent = `Sync failed: ${e}`;
    closeSyncModalBtn.style.display = 'block';
  }
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/[&<>'"]/g, 
    tag => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[tag] || tag));
}
