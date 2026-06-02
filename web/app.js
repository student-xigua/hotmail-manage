const state = {
  accounts: [],
  settings: { delimiter: '----', defaultTop: 1, apiMode: 'Graph' },
  selected: new Set(),
  search: '',
  status: 'all',
  page: 1,
  pageSize: 10,
  lastMessages: new Map(),
};

const el = {
  delimiterInput: document.querySelector('#delimiterInput'),
  defaultTopSelect: document.querySelector('#defaultTopSelect'),
  apiModeSelect: document.querySelector('#apiModeSelect'),
  dropZone: document.querySelector('#dropZone'),
  fileInput: document.querySelector('#fileInput'),
  pasteImportButton: document.querySelector('#pasteImportButton'),
  importDialog: document.querySelector('#importDialog'),
  importText: document.querySelector('#importText'),
  cancelImportButton: document.querySelector('#cancelImportButton'),
  confirmImportButton: document.querySelector('#confirmImportButton'),
  searchInput: document.querySelector('#searchInput'),
  statusFilter: document.querySelector('#statusFilter'),
  selectAllToolbar: document.querySelector('#selectAllToolbar'),
  batchFetchButton: document.querySelector('#batchFetchButton'),
  exportButton: document.querySelector('#exportButton'),
  deleteSelectedButton: document.querySelector('#deleteSelectedButton'),
  clearAllButton: document.querySelector('#clearAllButton'),
  headCheckbox: document.querySelector('#headCheckbox'),
  accountRows: document.querySelector('#accountRows'),
  emptyState: document.querySelector('#emptyState'),
  totalCount: document.querySelector('#totalCount'),
  todayCount: document.querySelector('#todayCount'),
  selectedCount: document.querySelector('#selectedCount'),
  totalRowsText: document.querySelector('#totalRowsText'),
  pageSizeSelect: document.querySelector('#pageSizeSelect'),
  prevPage: document.querySelector('#prevPage'),
  nextPage: document.querySelector('#nextPage'),
  pageNumber: document.querySelector('#pageNumber'),
  pageJump: document.querySelector('#pageJump'),
  messagesDialog: document.querySelector('#messagesDialog'),
  messagesTitle: document.querySelector('#messagesTitle'),
  messagesMeta: document.querySelector('#messagesMeta'),
  messagesList: document.querySelector('#messagesList'),
  toast: document.querySelector('#toast'),
};

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function showToast(message) {
  el.toast.textContent = message;
  el.toast.classList.add('show');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => el.toast.classList.remove('show'), 2800);
}

async function requestJson(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || payload.ok === false) {
    throw new Error(payload.error || `请求失败: ${response.status}`);
  }
  return payload;
}

function syncFromPayload(payload) {
  if (Array.isArray(payload.accounts)) {
    state.accounts = payload.accounts;
    const validIds = new Set(state.accounts.map((item) => item.id));
    state.selected = new Set([...state.selected].filter((id) => validIds.has(id)));
  }
  if (payload.settings) {
    state.settings = payload.settings;
    el.delimiterInput.value = state.settings.delimiter || '----';
    el.defaultTopSelect.value = String(state.settings.defaultTop || 1);
    el.apiModeSelect.value = state.settings.apiMode || 'Graph';
  }
}

async function loadAccounts() {
  const payload = await requestJson('/api/accounts');
  syncFromPayload(payload);
  render();
}

function filteredAccounts() {
  const query = state.search.trim().toLowerCase();
  return state.accounts.filter((account) => {
    const status = account.used ? 'used' : account.status;
    const matchesStatus = state.status === 'all' || status === state.status;
    const haystack = [
      account.email,
      account.status,
      account.clientId,
      account.refreshToken,
      account.lastError,
    ].join(' ').toLowerCase();
    return matchesStatus && (!query || haystack.includes(query));
  });
}

function currentPageRows() {
  const rows = filteredAccounts();
  const pageCount = Math.max(1, Math.ceil(rows.length / state.pageSize));
  state.page = Math.min(Math.max(1, state.page), pageCount);
  const start = (state.page - 1) * state.pageSize;
  return { rows, visible: rows.slice(start, start + state.pageSize), pageCount };
}

function statusLabel(account) {
  if (account.used) return ['已用', 'used'];
  if (account.status === 'error') return ['异常', 'error'];
  if (account.status === 'fetching') return ['获取中', 'fetching'];
  if (account.status === 'success') return ['剩余 90 天', 'success'];
  return ['待取件', 'success'];
}

function renderRows() {
  const { rows, visible, pageCount } = currentPageRows();
  el.emptyState.hidden = state.accounts.length > 0;
  el.accountRows.innerHTML = visible.map((account, index) => {
    const [label, statusClass] = statusLabel(account);
    const rowNumber = (state.page - 1) * state.pageSize + index + 1;
    return `
      <tr>
        <td><input type="checkbox" data-select-id="${escapeHtml(account.id)}" ${state.selected.has(account.id) ? 'checked' : ''}></td>
        <td>${rowNumber}</td>
        <td><span class="email-link">${escapeHtml(account.email)}</span></td>
        <td class="mono" title="${escapeHtml(account.password)}">${escapeHtml(account.password || '-')}</td>
        <td class="mono" title="${escapeHtml(account.clientId)}">${escapeHtml(account.clientId || '-')}</td>
        <td class="mono" title="${escapeHtml(account.refreshToken)}">${escapeHtml(account.refreshToken || '-')}</td>
        <td><span class="status-pill ${statusClass}">${label}</span></td>
        <td>
          <div class="row-actions">
            <button class="button view" data-action="view" data-id="${escapeHtml(account.id)}" type="button">查看邮件</button>
            <button class="button fetch" data-action="fetch" data-id="${escapeHtml(account.id)}" type="button">刷新邮件</button>
            <button class="button light" data-action="toggle-used" data-id="${escapeHtml(account.id)}" type="button">${account.used ? '恢复' : '已用'}</button>
            <button class="button delete" data-action="delete" data-id="${escapeHtml(account.id)}" type="button">删除</button>
          </div>
        </td>
      </tr>
    `;
  }).join('');

  const visibleIds = visible.map((item) => item.id);
  el.headCheckbox.checked = visibleIds.length > 0 && visibleIds.every((id) => state.selected.has(id));
  el.totalRowsText.textContent = `共 ${rows.length} 条`;
  el.pageNumber.textContent = String(state.page);
  el.pageJump.value = String(state.page);
  el.prevPage.disabled = state.page <= 1;
  el.nextPage.disabled = state.page >= pageCount;
}

function renderStats() {
  const today = new Date().toISOString().slice(0, 10);
  el.totalCount.textContent = String(state.accounts.length);
  el.todayCount.textContent = String(state.accounts.filter((item) => String(item.createdAt || '').startsWith(today)).length);
  el.selectedCount.textContent = String(state.selected.size);
}

function render() {
  renderStats();
  renderRows();
}

async function importText(text) {
  const payload = await requestJson('/api/accounts/import', {
    method: 'POST',
    body: JSON.stringify({
      text,
      settings: {
        delimiter: el.delimiterInput.value.trim() || '----',
        defaultTop: Number(el.defaultTopSelect.value) || 1,
        apiMode: el.apiModeSelect.value,
      },
    }),
  });
  syncFromPayload(payload);
  state.page = 1;
  render();
  showToast(`导入 ${payload.imported || 0} 条，更新 ${payload.updated || 0} 条`);
}

async function saveSettings() {
  try {
    const payload = await requestJson('/api/settings', {
      method: 'POST',
      body: JSON.stringify({
        delimiter: el.delimiterInput.value.trim() || '----',
        defaultTop: Number(el.defaultTopSelect.value) || 1,
        apiMode: el.apiModeSelect.value,
      }),
    });
    state.settings = payload.settings || state.settings;
  } catch (error) {
    showToast(error.message);
  }
}

async function fetchAccount(accountId, showDialog = true) {
  const payload = await requestJson(`/api/accounts/${encodeURIComponent(accountId)}/messages`, {
    method: 'POST',
    body: JSON.stringify({
      top: Number(el.defaultTopSelect.value) || 1,
      mailboxes: ['INBOX', 'Junk'],
    }),
  });
  if (payload.account) {
    const index = state.accounts.findIndex((item) => item.id === payload.account.id);
    if (index >= 0) state.accounts[index] = payload.account;
  }
  if (Array.isArray(payload.messages)) {
    state.lastMessages.set(accountId, payload.messages);
  }
  render();
  if (showDialog) {
    const account = state.accounts.find((item) => item.id === accountId);
    openMessages(account, payload.messages || [], payload);
  }
  return payload;
}

function openMessages(account, messages, meta = {}) {
  el.messagesTitle.textContent = account?.email ? `${account.email} 的邮件` : '邮件列表';
  el.messagesMeta.textContent = `${messages.length} 条 | ${meta.transport || account?.transport || 'unknown'} | ${meta.tokenEndpoint || account?.tokenEndpoint || ''}`;
  el.messagesList.innerHTML = messages.length ? messages.map((message) => {
    const sender = message?.from?.emailAddress?.address || '';
    return `
      <article class="message-item">
        <strong>${escapeHtml(message.subject || '(无主题)')}</strong>
        <p>发件人: ${escapeHtml(sender || '-')}</p>
        <p>邮箱夹: ${escapeHtml(message.mailbox || '-')} | 时间: ${escapeHtml(message.receivedDateTime || '-')}</p>
        <p>${escapeHtml(message.bodyPreview || '')}</p>
      </article>
    `;
  }).join('') : '<div class="empty-state">没有读取到邮件。</div>';
  el.messagesDialog.showModal();
}

async function deleteAccounts(ids, mode = '') {
  const payload = await requestJson('/api/accounts/delete', {
    method: 'POST',
    body: JSON.stringify({ ids, mode }),
  });
  syncFromPayload(payload);
  render();
  showToast(`已删除 ${payload.deleted || 0} 条`);
}

async function exportAccounts() {
  const payload = await requestJson('/api/accounts/export');
  const blob = new Blob([payload.content || ''], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = payload.filename || 'hotmail-accounts.txt';
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function bindEvents() {
  el.pasteImportButton.addEventListener('click', () => el.importDialog.showModal());
  el.cancelImportButton.addEventListener('click', () => el.importDialog.close());
  el.confirmImportButton.addEventListener('click', async () => {
    try {
      await importText(el.importText.value);
      el.importText.value = '';
      el.importDialog.close();
    } catch (error) {
      showToast(error.message);
    }
  });

  el.fileInput.addEventListener('change', async () => {
    const file = el.fileInput.files?.[0];
    if (!file) return;
    try {
      await importText(await file.text());
    } catch (error) {
      showToast(error.message);
    } finally {
      el.fileInput.value = '';
    }
  });

  ['dragenter', 'dragover'].forEach((eventName) => {
    el.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      el.dropZone.classList.add('is-dragging');
    });
  });
  ['dragleave', 'drop'].forEach((eventName) => {
    el.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      el.dropZone.classList.remove('is-dragging');
    });
  });
  el.dropZone.addEventListener('drop', async (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;
    try {
      await importText(await file.text());
    } catch (error) {
      showToast(error.message);
    }
  });

  el.searchInput.addEventListener('input', () => {
    state.search = el.searchInput.value;
    state.page = 1;
    renderRows();
  });
  el.statusFilter.addEventListener('change', () => {
    state.status = el.statusFilter.value;
    state.page = 1;
    renderRows();
  });
  el.defaultTopSelect.addEventListener('change', saveSettings);
  el.apiModeSelect.addEventListener('change', saveSettings);
  el.delimiterInput.addEventListener('change', saveSettings);

  el.headCheckbox.addEventListener('change', () => {
    const { visible } = currentPageRows();
    for (const account of visible) {
      if (el.headCheckbox.checked) state.selected.add(account.id);
      else state.selected.delete(account.id);
    }
    render();
  });

  el.selectAllToolbar.addEventListener('click', () => {
    const { rows } = currentPageRows();
    const allSelected = rows.length > 0 && rows.every((item) => state.selected.has(item.id));
    for (const account of rows) {
      if (allSelected) state.selected.delete(account.id);
      else state.selected.add(account.id);
    }
    render();
  });

  el.accountRows.addEventListener('change', (event) => {
    const id = event.target?.dataset?.selectId;
    if (!id) return;
    if (event.target.checked) state.selected.add(id);
    else state.selected.delete(id);
    render();
  });

  el.accountRows.addEventListener('click', async (event) => {
    const button = event.target.closest('[data-action]');
    if (!button) return;
    const id = button.dataset.id;
    const account = state.accounts.find((item) => item.id === id);
    try {
      button.disabled = true;
      if (button.dataset.action === 'view') {
        const cached = state.lastMessages.get(id);
        if (cached) openMessages(account, cached, {});
        else await fetchAccount(id, true);
      } else if (button.dataset.action === 'fetch') {
        await fetchAccount(id, true);
      } else if (button.dataset.action === 'toggle-used') {
        const payload = await requestJson(`/api/accounts/${encodeURIComponent(id)}`, {
          method: 'POST',
          body: JSON.stringify({ updates: { used: !account.used, status: account.used ? 'ready' : 'used' } }),
        });
        const index = state.accounts.findIndex((item) => item.id === payload.account.id);
        if (index >= 0) state.accounts[index] = payload.account;
        render();
      } else if (button.dataset.action === 'delete') {
        if (confirm(`确认删除 ${account.email} 吗？`)) await deleteAccounts([id]);
      }
    } catch (error) {
      showToast(error.message);
    } finally {
      button.disabled = false;
    }
  });

  el.batchFetchButton.addEventListener('click', async () => {
    const ids = [...state.selected];
    if (!ids.length) {
      showToast('请先选择要取件的邮箱');
      return;
    }
    el.batchFetchButton.disabled = true;
    try {
      const payload = await requestJson('/api/accounts/batch-fetch', {
        method: 'POST',
        body: JSON.stringify({
          ids,
          top: Number(el.defaultTopSelect.value) || 1,
          mailboxes: ['INBOX', 'Junk'],
        }),
      });
      syncFromPayload(payload);
      render();
      const okCount = (payload.results || []).filter((item) => item.ok).length;
      showToast(`批量取件完成：成功 ${okCount} 个，失败 ${(payload.results || []).length - okCount} 个`);
    } catch (error) {
      showToast(error.message);
    } finally {
      el.batchFetchButton.disabled = false;
    }
  });

  el.exportButton.addEventListener('click', () => exportAccounts().catch((error) => showToast(error.message)));
  el.deleteSelectedButton.addEventListener('click', async () => {
    const ids = [...state.selected];
    if (!ids.length) return showToast('请先选择要删除的数据');
    if (confirm(`确认删除选中的 ${ids.length} 条数据吗？`)) {
      await deleteAccounts(ids);
    }
  });
  el.clearAllButton.addEventListener('click', async () => {
    if (confirm('确认清空所有邮箱数据吗？')) {
      await deleteAccounts([], 'all');
    }
  });

  el.pageSizeSelect.addEventListener('change', () => {
    state.pageSize = Number(el.pageSizeSelect.value) || 10;
    state.page = 1;
    renderRows();
  });
  el.prevPage.addEventListener('click', () => {
    state.page -= 1;
    renderRows();
  });
  el.nextPage.addEventListener('click', () => {
    state.page += 1;
    renderRows();
  });
  el.pageJump.addEventListener('change', () => {
    state.page = Number(el.pageJump.value) || 1;
    renderRows();
  });
}

bindEvents();
loadAccounts().catch((error) => showToast(error.message));
