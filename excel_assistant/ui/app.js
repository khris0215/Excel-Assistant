const formFields = Array.from(document.querySelectorAll('[data-path]'));
const resultsBody = document.getElementById('resultsBody');
const resultCount = document.getElementById('resultCount');
const monitorState = document.getElementById('monitorState');

const btnSave = document.getElementById('btnSave');
const btnOpenSettings = document.getElementById('btnOpenSettings');
const btnRunNow = document.getElementById('btnRunNow');
const btnRefresh = document.getElementById('btnRefresh');
const btnPauseResume = document.getElementById('btnPauseResume');
const btnBackDashboard = document.getElementById('btnBackDashboard');
const tabDashboard = document.getElementById('tabDashboard');
const tabSettings = document.getElementById('tabSettings');
const panelDashboard = document.getElementById('panelDashboard');
const panelSettings = document.getElementById('panelSettings');
const watchMode = document.getElementById('watchMode');
const AUTO_REFRESH_MS = 15000;

function getApi() {
  if (window.pywebview && window.pywebview.api) {
    return window.pywebview.api;
  }

  const localKey = 'excel-assistant-settings';
  const fallbackSettings = {
    excel_file_path: '',
    poll_minutes: 5,
    autostart: false,
    email_sent_column: '',
    watch: {
      mode: 'range',
      sheet_name: '',
      start_row: 2,
      end_row: 200,
      start_col: 'A',
      end_col: 'A',
      row_list: '',
      column_list: '',
    },
    thresholds: {
      good_max: 20,
      soft_max: 28,
      medium_max: 35,
      hard_max: 44,
      due_at: 45,
    },
    email: {
      enabled: false,
      recipient_mode: 'global',
      global_recipient: '',
      email_column: '',
      sender_email: '',
      smtp_host: '',
      smtp_port: 587,
      smtp_username: '',
      smtp_password: '',
      use_tls: true,
      subject_template: '[Excel Assistant] {status} item at {cell}',
      body_template: 'Cell: {cell}\\nDays: {days}\\nStatus: {status}',
    },
  };

  return {
    async get_settings() {
      return JSON.parse(localStorage.getItem(localKey) || JSON.stringify(fallbackSettings));
    },
    async save_settings(payload) {
      localStorage.setItem(localKey, JSON.stringify(payload));
      return { ok: true, settings: payload };
    },
    async run_check_now() {
      return { ok: true, results: [] };
    },
    async get_last_results() {
      return [];
    },
    async pause_monitor() {
      return { ok: true, paused: true };
    },
    async resume_monitor() {
      return { ok: true, paused: false };
    },
    async monitor_state() {
      return { paused: false };
    },
  };
}

let currentSettings = null;
let monitorPaused = false;
let isRefreshingResults = false;
let autoRefreshTimer = null;

function liveApi() {
  // Resolve API at call time so we use pywebview once it is ready.
  return getApi();
}

function setByPath(obj, path, value) {
  const keys = path.split('.');
  let target = obj;
  for (let i = 0; i < keys.length - 1; i += 1) {
    if (!target[keys[i]]) {
      target[keys[i]] = {};
    }
    target = target[keys[i]];
  }
  target[keys[keys.length - 1]] = value;
}

function getByPath(obj, path) {
  return path.split('.').reduce((acc, key) => (acc ? acc[key] : undefined), obj);
}

function readValue(field) {
  if (field.type === 'checkbox') return field.checked;
  if (field.type === 'number') return Number(field.value || 0);
  return field.value;
}

function setFieldValue(field, value) {
  if (field.type === 'checkbox') {
    field.checked = Boolean(value);
    return;
  }
  field.value = value ?? '';
}

function switchTo(section) {
  const showSettings = section === 'settings';
  panelDashboard.classList.toggle('hidden', showSettings);
  panelSettings.classList.toggle('hidden', !showSettings);

  tabDashboard.classList.toggle('bg-ink', !showSettings);
  tabDashboard.classList.toggle('text-white', !showSettings);
  tabDashboard.classList.toggle('border', showSettings);
  tabDashboard.classList.toggle('border-slate-300', showSettings);
  tabDashboard.classList.toggle('bg-white', showSettings);

  tabSettings.classList.toggle('bg-ink', showSettings);
  tabSettings.classList.toggle('text-white', showSettings);
  tabSettings.classList.toggle('border', !showSettings);
  tabSettings.classList.toggle('border-slate-300', !showSettings);
  tabSettings.classList.toggle('bg-white', !showSettings);
}

function updateWatchModeUI() {
  const mode = watchMode.value;
  document.getElementById('modeRange').classList.toggle('hidden', mode !== 'range');
  document.getElementById('modeRows').classList.toggle('hidden', mode !== 'rows_all_columns');
  document.getElementById('modeCols').classList.toggle('hidden', mode !== 'columns_all_rows');
}

function renderThresholdLabels() {
  const t = currentSettings.thresholds;
  document.getElementById('lblGood').innerText = `0-${t.good_max}`;
  document.getElementById('lblSoft').innerText = `${t.good_max + 1}-${t.soft_max}`;
  document.getElementById('lblMedium').innerText = `${t.soft_max + 1}-${t.medium_max}`;
  document.getElementById('lblHard').innerText = `${t.medium_max + 1}-${t.hard_max}`;
  document.getElementById('lblDue').innerText = `>= ${t.due_at}`;
}

function renderRows(results) {
  resultCount.innerText = `${results.length} entries`;
  if (!results.length) {
    resultsBody.innerHTML = '<tr><td class="px-4 py-4 text-sm text-slate-500" colspan="6">No monitored entries found in current scan.</td></tr>';
    return;
  }

  resultsBody.innerHTML = results
    .map(
      (row) => `
      <tr class="border-t border-slate-200">
        <td class="px-4 py-3 font-semibold">${row.cell}</td>
        <td class="px-4 py-3">${row.entry_date}</td>
        <td class="px-4 py-3">${row.days}</td>
        <td class="px-4 py-3"><span class="chip status-${row.status}">${row.status.toUpperCase()}</span></td>
        <td class="px-4 py-3">${row.recipient || '-'}</td>
        <td class="px-4 py-3">${row.emailed ? 'Yes' : 'No'}</td>
      </tr>
    `,
    )
    .join('');
}

function hydrateForm(settings) {
  currentSettings = settings;
  formFields.forEach((field) => {
    setFieldValue(field, getByPath(settings, field.dataset.path));
  });
  updateWatchModeUI();
  renderThresholdLabels();
}

function collectForm() {
  const next = structuredClone(currentSettings);
  formFields.forEach((field) => {
    setByPath(next, field.dataset.path, readValue(field));
  });
  return next;
}

async function loadSettings() {
  const settings = await liveApi().get_settings();
  hydrateForm(settings);
}

async function refreshResults() {
  if (isRefreshingResults) {
    return;
  }

  isRefreshingResults = true;
  try {
    const rows = await liveApi().get_last_results();
    renderRows(rows);
  } finally {
    isRefreshingResults = false;
  }
}

function startAutoRefresh() {
  if (autoRefreshTimer) {
    clearInterval(autoRefreshTimer);
  }

  autoRefreshTimer = setInterval(async () => {
    if (document.hidden) {
      return;
    }
    await refreshResults();
  }, AUTO_REFRESH_MS);
}

async function saveSettings() {
  try {
    const payload = collectForm();
    const res = await liveApi().save_settings(payload);
    if (res.ok) {
      hydrateForm(res.settings);
      monitorState.innerText = 'Settings saved';
      setTimeout(async () => {
        const state = await liveApi().monitor_state();
        monitorState.innerText = state.paused ? 'Paused' : 'Running';
      }, 900);
    }
  } catch (error) {
    monitorState.innerText = `Save failed: ${error?.message || 'Unknown error'}`;
  }
}

async function runNow() {
  try {
    monitorState.innerText = 'Running manual check...';
    const res = await liveApi().run_check_now();
    if (!res.ok) {
      renderRows(res.results || []);
      monitorState.innerText = `Scan failed: ${res.error || 'Unknown error'}`;
      return;
    }
    renderRows(res.results || []);
    const state = await liveApi().monitor_state();
    monitorState.innerText = state.paused ? 'Paused' : 'Running';
  } catch (error) {
    monitorState.innerText = `Scan failed: ${error?.message || 'Unknown error'}`;
  }
}

async function togglePauseResume() {
  if (monitorPaused) {
    await liveApi().resume_monitor();
    monitorPaused = false;
  } else {
    await liveApi().pause_monitor();
    monitorPaused = true;
  }
  btnPauseResume.innerText = monitorPaused ? 'Resume' : 'Pause';
  monitorState.innerText = monitorPaused ? 'Paused' : 'Running';
}

watchMode.addEventListener('change', updateWatchModeUI);
formFields
  .filter((f) => f.dataset.path.startsWith('thresholds.'))
  .forEach((f) => f.addEventListener('input', () => {
    currentSettings = collectForm();
    renderThresholdLabels();
  }));

btnSave.addEventListener('click', saveSettings);
btnOpenSettings.addEventListener('click', () => switchTo('settings'));
btnRunNow.addEventListener('click', runNow);
btnRefresh.addEventListener('click', refreshResults);
btnPauseResume.addEventListener('click', togglePauseResume);
btnBackDashboard.addEventListener('click', () => switchTo('dashboard'));
tabDashboard.addEventListener('click', () => switchTo('dashboard'));
tabSettings.addEventListener('click', () => switchTo('settings'));

window.addEventListener('pywebviewready', async () => {
  await loadSettings();
  const state = await liveApi().monitor_state();
  monitorPaused = Boolean(state.paused);
  btnPauseResume.innerText = monitorPaused ? 'Resume' : 'Pause';
  monitorState.innerText = monitorPaused ? 'Paused' : 'Running';
  await refreshResults();
  startAutoRefresh();
});

(async function bootstrap() {
  switchTo('dashboard');
  await loadSettings();
  const state = await liveApi().monitor_state();
  monitorPaused = Boolean(state.paused);
  btnPauseResume.innerText = monitorPaused ? 'Resume' : 'Pause';
  monitorState.innerText = monitorPaused ? 'Paused' : 'Running';
  await refreshResults();
  startAutoRefresh();
})();

window.addEventListener('beforeunload', () => {
  if (autoRefreshTimer) {
    clearInterval(autoRefreshTimer);
    autoRefreshTimer = null;
  }
});
