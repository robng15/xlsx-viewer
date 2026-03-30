// ─── Filter Module ────────────────────────────────────────────────────────────
// Attaches Excel-style per-column filter dropdowns to a rendered sheet table.
// Relies on data-c attributes on <th> and <td> elements set by renderer.js.

let filterState   = new Map(); // colIndex (number) → Set<string> of shown values
let activeDropdown = null;     // currently open dropdown element
let activeTable    = null;

// ─── Public API ───────────────────────────────────────────────────────────────

export function setupFilters(tableWrapper) {
  teardownFilters();
  filterState  = new Map();
  activeTable  = tableWrapper.querySelector('.sheet-table');
  if (!activeTable) return;

  for (const arrow of activeTable.querySelectorAll('.filter-arrow')) {
    arrow.addEventListener('click', onArrowClick);
  }
  document.addEventListener('click', onDocClick, true);
}

export function teardownFilters() {
  closeDropdown();
  if (activeTable) {
    for (const arrow of activeTable.querySelectorAll('.filter-arrow')) {
      arrow.removeEventListener('click', onArrowClick);
    }
  }
  document.removeEventListener('click', onDocClick, true);
  filterState  = new Map();
  activeTable  = null;
}

// ─── Event Handlers ───────────────────────────────────────────────────────────

function onArrowClick(e) {
  e.stopPropagation();
  const colIdx = Number(e.currentTarget.dataset.c);
  if (activeDropdown?.dataset.col === String(colIdx)) {
    closeDropdown();
  } else {
    openDropdown(colIdx, e.currentTarget);
  }
}

function onDocClick(e) {
  if (activeDropdown && !activeDropdown.contains(e.target)) {
    closeDropdown();
  }
}

// ─── Dropdown ─────────────────────────────────────────────────────────────────

function openDropdown(colIdx, arrowEl) {
  closeDropdown();

  const allValues    = collectColumnValues(colIdx);
  const activeValues = filterState.get(colIdx) ?? new Set(allValues);

  const dropdown = document.createElement('div');
  dropdown.className  = 'filter-dropdown';
  dropdown.dataset.col = colIdx;
  dropdown.addEventListener('click', e => e.stopPropagation());

  // ── Search ────────────────────────────────────────────────────────────────
  const searchBox = document.createElement('input');
  searchBox.type        = 'text';
  searchBox.placeholder = 'Search values…';
  searchBox.className   = 'filter-search';
  dropdown.appendChild(searchBox);

  // ── Select All ────────────────────────────────────────────────────────────
  const selectAllRow  = makeLabel('filter-option filter-all', '(Select All)');
  const selectAllCb   = makeCheckbox(activeValues.size === allValues.length);
  if (activeValues.size > 0 && activeValues.size < allValues.length) {
    selectAllCb.indeterminate = true;
  }
  selectAllRow.prepend(selectAllCb);
  dropdown.appendChild(selectAllRow);
  dropdown.appendChild(Object.assign(document.createElement('hr'), { className: 'filter-hr' }));

  // ── Value List ────────────────────────────────────────────────────────────
  const listEl   = document.createElement('div');
  listEl.className = 'filter-list';
  const checkboxes = [];

  for (const val of allValues) {
    const row  = makeLabel('filter-option', val === '' ? '(Blank)' : val);
    const cb   = makeCheckbox(activeValues.has(val));
    cb.dataset.val = val;
    checkboxes.push(cb);
    cb.addEventListener('change', () => syncSelectAll(selectAllCb, checkboxes));
    row.prepend(cb);
    listEl.appendChild(row);
  }
  dropdown.appendChild(listEl);

  selectAllCb.addEventListener('change', () => {
    const checked = selectAllCb.checked;
    for (const cb of visibleCheckboxes(listEl)) cb.checked = checked;
    syncSelectAll(selectAllCb, checkboxes);
  });

  searchBox.addEventListener('input', () => {
    const q = searchBox.value.toLowerCase();
    for (const row of listEl.querySelectorAll('.filter-option')) {
      row.hidden = q !== '' && !row.textContent.toLowerCase().includes(q);
    }
    syncSelectAll(selectAllCb, checkboxes);
  });

  // ── Buttons ───────────────────────────────────────────────────────────────
  const btnRow = document.createElement('div');
  btnRow.className = 'filter-btn-row';

  const clearBtn = Object.assign(document.createElement('button'), {
    className: 'filter-btn filter-btn-clear',
    textContent: 'Clear',
  });
  clearBtn.addEventListener('click', () => {
    applyFilter(colIdx, null);
    closeDropdown();
  });

  const applyBtn = Object.assign(document.createElement('button'), {
    className: 'filter-btn filter-btn-apply',
    textContent: 'Apply',
  });
  applyBtn.addEventListener('click', () => {
    const selected = new Set(
      checkboxes.filter(cb => cb.checked).map(cb => cb.dataset.val)
    );
    applyFilter(colIdx, selected.size === allValues.length ? null : selected);
    closeDropdown();
  });

  btnRow.appendChild(clearBtn);
  btnRow.appendChild(applyBtn);
  dropdown.appendChild(btnRow);

  // ── Position ──────────────────────────────────────────────────────────────
  const container = activeTable.closest('.table-container');
  container.appendChild(dropdown);

  const arrowRect     = arrowEl.getBoundingClientRect();
  const containerRect = container.getBoundingClientRect();
  let top  = arrowRect.bottom - containerRect.top  + container.scrollTop;
  let left = arrowRect.left   - containerRect.left + container.scrollLeft;

  // Flip left if dropdown would overflow viewport
  dropdown.style.visibility = 'hidden';
  dropdown.style.top  = `${top}px`;
  dropdown.style.left = `${left}px`;

  requestAnimationFrame(() => {
    const ddRect = dropdown.getBoundingClientRect();
    if (ddRect.right > window.innerWidth - 8) {
      left = Math.max(0, left - (ddRect.right - window.innerWidth + 8));
    }
    dropdown.style.left = `${left}px`;
    dropdown.style.visibility = '';
    searchBox.focus();
  });

  activeDropdown = dropdown;
}

function closeDropdown() {
  activeDropdown?.remove();
  activeDropdown = null;
}

// ─── Filter Logic ─────────────────────────────────────────────────────────────

function applyFilter(colIdx, selectedValues) {
  if (selectedValues === null) {
    filterState.delete(colIdx);
  } else {
    filterState.set(colIdx, selectedValues);
  }
  applyAllFilters();
  updateArrowState(colIdx);
}

function applyAllFilters() {
  if (!activeTable) return;
  for (const tr of activeTable.querySelectorAll('tbody tr')) {
    let show = true;
    for (const [col, allowed] of filterState) {
      const cell = tr.querySelector(`td[data-c="${col}"]`);
      const val  = cell ? cell.textContent.trim() : '';
      if (!allowed.has(val)) { show = false; break; }
    }
    tr.hidden = !show;
  }
}

function updateArrowState(colIdx) {
  const arrow = activeTable?.querySelector(`.filter-arrow[data-c="${colIdx}"]`);
  if (arrow) arrow.classList.toggle('active', filterState.has(colIdx));
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function collectColumnValues(colIdx) {
  const seen   = new Set();
  const values = [];
  for (const cell of activeTable.querySelectorAll(`td[data-c="${colIdx}"]`)) {
    const val = cell.textContent.trim();
    if (!seen.has(val)) { seen.add(val); values.push(val); }
  }
  return values.sort((a, b) => {
    const na = Number(a), nb = Number(b);
    if (a !== '' && b !== '' && !isNaN(na) && !isNaN(nb)) return na - nb;
    if (a === '') return 1;
    if (b === '') return -1;
    return a.localeCompare(b, undefined, { sensitivity: 'base' });
  });
}

function makeLabel(className, text) {
  const lbl = document.createElement('label');
  lbl.className = className;
  lbl.appendChild(document.createTextNode(text));
  return lbl;
}

function makeCheckbox(checked) {
  const cb   = document.createElement('input');
  cb.type    = 'checkbox';
  cb.checked = checked;
  return cb;
}

function visibleCheckboxes(listEl) {
  return [...listEl.querySelectorAll('input[type=checkbox]')]
    .filter(cb => !cb.closest('.filter-option').hidden);
}

function syncSelectAll(selectAllCb, allCbs) {
  const visible = allCbs.filter(cb => !cb.closest('.filter-option').hidden);
  const checkedCount = visible.filter(cb => cb.checked).length;
  selectAllCb.checked       = checkedCount === visible.length;
  selectAllCb.indeterminate = checkedCount > 0 && checkedCount < visible.length;
}
