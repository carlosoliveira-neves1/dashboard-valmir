const USERS_KEY = "simplicio_users_v1";
const SESSION_KEY = "simplicio_session_v1";

const fileInput = document.getElementById("fileInput");
const fileName = document.getElementById("fileName");
const kpisContainer = document.getElementById("kpis");
const filtersContainer = document.getElementById("filters");
const insightsContainer = document.getElementById("insights");
const reportsContainer = document.getElementById("reports");
const summaryTableBody = document.querySelector("#summaryTable tbody");
const dashboardArea = document.getElementById("dashboardArea");
const accessDenied = document.getElementById("accessDenied");

const searchInput = document.getElementById("searchInput");
const ufFilter = document.getElementById("ufFilter");
const statusFilter = document.getElementById("statusFilter");
const companyFilter = document.getElementById("companyFilter");
const companyList = document.getElementById("companyList");
const startDateFilter = document.getElementById("startDateFilter");
const endDateFilter = document.getElementById("endDateFilter");
const clearFilters = document.getElementById("clearFilters");
const filterResult = document.getElementById("filterResult");

const zoomModal = document.getElementById("zoomModal");
const modalBackdrop = document.getElementById("modalBackdrop");
const closeModalBtn = document.getElementById("closeModal");
const modalTitle = document.getElementById("modalTitle");
const modalBody = document.getElementById("modalBody");

const loginView = document.getElementById("loginView");
const appView = document.getElementById("appView");
const loginForm = document.getElementById("loginForm");
const loginUsername = document.getElementById("loginUsername");
const loginPassword = document.getElementById("loginPassword");
const loginMessage = document.getElementById("loginMessage");
const logoutBtn = document.getElementById("logoutBtn");
const sessionInfo = document.getElementById("sessionInfo");

const adminToggle = document.getElementById("adminToggle");
const adminSection = document.getElementById("adminSection");
const createUserForm = document.getElementById("createUserForm");
const newUsername = document.getElementById("newUsername");
const newPassword = document.getElementById("newPassword");
const adminMessage = document.getElementById("adminMessage");
const usersTableBody = document.querySelector("#usersTable tbody");

const charts = {
  ufChart: null,
  statusChart: null,
  timelineChart: null,
  monthlyValueChart: null,
  companyValueChart: null,
  ufValueChart: null,
  statusValueChart: null,
  valueRangeChart: null,
};

const dashboardState = {
  allRows: [],
  filteredRows: [],
  detected: null,
};

const authState = {
  currentUser: null,
};

const palette = ["#0f766e", "#0e7490", "#06b6d4", "#14b8a6", "#f59e0b", "#ef4444", "#6366f1", "#8b5cf6", "#84cc16", "#475569"];

const normalizeHeader = (header, index) => {
  const value = String(header ?? "").trim();
  return value || `COLUNA_${index + 1}`;
};

const cleanText = (value) => String(value ?? "").trim().replace(/\s+/g, " ");
const stripAccents = (value) => cleanText(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();

const toBase64 = (text) => btoa(unescape(encodeURIComponent(text)));
const fromBase64 = (text) => decodeURIComponent(escape(atob(text)));

const safeReadJSON = (key, fallback) => {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
};

const saveJSON = (key, value) => localStorage.setItem(key, JSON.stringify(value));

const ensureDefaultAdmin = () => {
  const users = safeReadJSON(USERS_KEY, []);
  const seedUsers = [
    {
      username: "admin",
      password: "admin123",
      permissions: ["view_dashboard", "manage_users"],
    },
    {
      username: "simplicio",
      password: "Sucesso@2026",
      permissions: ["view_dashboard", "manage_users"],
    },
  ];

  seedUsers.forEach((seed) => {
    const exists = users.some((u) => u.username.toLowerCase() === seed.username.toLowerCase());
    if (exists) return;

    users.push({
      id: crypto.randomUUID(),
      username: seed.username,
      password: toBase64(seed.password),
      active: true,
      permissions: seed.permissions,
      createdAt: new Date().toISOString(),
    });
  });

  saveUsers(users);
};

const getUsers = () => safeReadJSON(USERS_KEY, []);
const saveUsers = (users) => saveJSON(USERS_KEY, users);

const getSession = () => safeReadJSON(SESSION_KEY, null);
const setSession = (session) => saveJSON(SESSION_KEY, session);
const clearSession = () => localStorage.removeItem(SESSION_KEY);

const hasPermission = (permission) => Boolean(authState.currentUser?.permissions?.includes(permission));

const parseNumber = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value == null) return null;

  const text = cleanText(value);
  if (!text) return null;

  const numeric = text.replace(/[^\d,.-]/g, "");
  if (!numeric) return null;

  let normalized = numeric;
  if (numeric.includes(",") && numeric.includes(".")) {
    normalized = numeric.lastIndexOf(",") > numeric.lastIndexOf(".") ? numeric.replace(/\./g, "").replace(",", ".") : numeric.replace(/,/g, "");
  } else if (numeric.includes(",")) {
    normalized = numeric.replace(",", ".");
  }

  const result = Number(normalized);
  return Number.isFinite(result) ? result : null;
};

const excelSerialToDate = (serial) => {
  const utc = Math.round((serial - 25569) * 86400 * 1000);
  const d = new Date(utc);
  return Number.isNaN(d.getTime()) ? null : d;
};

const parseDate = (value) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && Number.isFinite(value) && value > 1000) return excelSerialToDate(value);
  if (value == null) return null;

  const text = cleanText(value);
  if (!text) return null;

  const br = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (br) {
    const day = Number(br[1]);
    const month = Number(br[2]) - 1;
    let year = Number(br[3]);
    if (year < 100) year += 2000;
    const d = new Date(year, month, day);
    if (!Number.isNaN(d.getTime())) return d;
  }

  const iso = new Date(text);
  return Number.isNaN(iso.getTime()) ? null : iso;
};

const inferColumns = (rows, headers) => {
  const keyMap = headers.map((h) => stripAccents(h));
  const byKeyword = (words) => {
    const idx = keyMap.findIndex((h) => words.some((w) => h.includes(w)));
    return idx >= 0 ? headers[idx] : null;
  };

  let dateCol = byKeyword(["data", "dt"]);
  let valueCol = byKeyword(["valor", "total", "consolidado", "faturamento"]);
  let ufCol = byKeyword([" uf", "uf ", "estado", "unidade federativa"]);
  if (!ufCol) {
    const exactUf = headers.find((h) => stripAccents(h) === "uf");
    if (exactUf) ufCol = exactUf;
  }
  let statusCol = byKeyword(["situacao", "status", "estatus"]);
  let companyCol = byKeyword(["empresa", "razao", "nome fantasia", "cliente", "pj"]);

  const candidates = headers.map((h) => ({
    header: h,
    numericCount: rows.reduce((acc, r) => (parseNumber(r[h]) != null ? acc + 1 : acc), 0),
    dateCount: rows.reduce((acc, r) => (parseDate(r[h]) ? acc + 1 : acc), 0),
    uniqueCount: new Set(rows.map((r) => cleanText(r[h])).filter(Boolean)).size,
  }));

  if (!valueCol) {
    const best = [...candidates].sort((a, b) => b.numericCount - a.numericCount)[0];
    if (best?.numericCount > 0) valueCol = best.header;
  }
  if (!dateCol) {
    const best = [...candidates].sort((a, b) => b.dateCount - a.dateCount)[0];
    if (best?.dateCount > 0) dateCol = best.header;
  }
  if (!statusCol) {
    const best = candidates.filter((c) => c.uniqueCount > 1 && c.uniqueCount <= 24).sort((a, b) => a.uniqueCount - b.uniqueCount)[0];
    if (best) statusCol = best.header;
  }
  if (!companyCol) {
    const best = candidates.filter((c) => c.uniqueCount > 24).sort((a, b) => b.uniqueCount - a.uniqueCount)[0];
    if (best) companyCol = best.header;
  }

  return { dateCol, valueCol, ufCol, statusCol, companyCol };
};

const fmtNumber = (v) => new Intl.NumberFormat("pt-BR").format(v ?? 0);
const fmtCurrency = (v) => new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 2 }).format(v ?? 0);
const fmtPercent = (v) => `${v.toFixed(1)}%`;

const groupCount = (rows, col, limit = 10) => {
  if (!col) return [];
  const map = new Map();
  rows.forEach((row) => {
    const key = cleanText(row[col]) || "Não informado";
    map.set(key, (map.get(key) ?? 0) + 1);
  });
  return [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, limit);
};

const groupSum = (rows, groupCol, valueCol, limit = 10) => {
  if (!groupCol || !valueCol) return [];
  const map = new Map();
  rows.forEach((row) => {
    const group = cleanText(row[groupCol]) || "Não informado";
    const value = parseNumber(row[valueCol]) ?? 0;
    map.set(group, (map.get(group) ?? 0) + value);
  });
  return [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, limit);
};

const byMonth = (rows, dateCol, valueCol = null) => {
  if (!dateCol) return [];
  const map = new Map();
  rows.forEach((row) => {
    const date = parseDate(row[dateCol]);
    if (!date) return;
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    map.set(key, (map.get(key) ?? 0) + (valueCol ? parseNumber(row[valueCol]) ?? 0 : 1));
  });
  return [...map.entries()].sort((a, b) => a[0].localeCompare(b[0]));
};

const valueRanges = (rows, valueCol) => {
  if (!valueCol) return [];
  const buckets = [
    { label: "Até 10 mil", min: 0, max: 10000, count: 0 },
    { label: "10 mil - 50 mil", min: 10000, max: 50000, count: 0 },
    { label: "50 mil - 100 mil", min: 50000, max: 100000, count: 0 },
    { label: "100 mil - 500 mil", min: 100000, max: 500000, count: 0 },
    { label: "500 mil - 1 milhão", min: 500000, max: 1000000, count: 0 },
    { label: "Acima de 1 milhão", min: 1000000, max: Number.POSITIVE_INFINITY, count: 0 },
  ];

  rows.forEach((row) => {
    const value = parseNumber(row[valueCol]);
    if (value == null || value < 0) return;
    const bucket = buckets.find((b) => value >= b.min && value < b.max);
    if (bucket) bucket.count += 1;
  });

  return buckets.map((b) => [b.label, b.count]);
};

const uniqueOptions = (rows, col) => {
  if (!col) return [];
  return [...new Set(rows.map((row) => cleanText(row[col])).filter(Boolean))].sort((a, b) => a.localeCompare(b, "pt-BR"));
};

const fillSelect = (select, values, firstLabel) => {
  select.innerHTML = "";
  const first = document.createElement("option");
  first.value = "";
  first.textContent = firstLabel;
  select.appendChild(first);
  values.forEach((value) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    select.appendChild(opt);
  });
};

const fillCompanyDatalist = (values) => {
  companyList.innerHTML = "";
  values.slice(0, 1000).forEach((value) => {
    const opt = document.createElement("option");
    opt.value = value;
    companyList.appendChild(opt);
  });
};

const initFilters = (rows, detected) => {
  fillSelect(ufFilter, uniqueOptions(rows, detected.ufCol), "Todas");
  fillSelect(statusFilter, uniqueOptions(rows, detected.statusCol), "Todos");
  fillCompanyDatalist(uniqueOptions(rows, detected.companyCol));

  searchInput.value = "";
  ufFilter.value = "";
  statusFilter.value = "";
  companyFilter.value = "";
  startDateFilter.value = "";
  endDateFilter.value = "";

  filtersContainer.classList.remove("hidden");
};

const destroyCharts = () => Object.values(charts).forEach((instance) => instance?.destroy());

const chartBaseOptions = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: { legend: { labels: { boxWidth: 12, usePointStyle: true } } },
};

const renderKpis = (rows, allRows, detected) => {
  const totalRows = rows.length;
  const allRowsCount = allRows.length;
  const totalCols = Object.keys(rows[0] ?? allRows[0] ?? {}).length;
  const numericValues = detected.valueCol ? rows.map((r) => parseNumber(r[detected.valueCol])).filter((n) => n != null) : [];
  const dateValues = detected.dateCol ? rows.map((r) => parseDate(r[detected.dateCol])).filter(Boolean) : [];

  const sum = numericValues.reduce((acc, n) => acc + n, 0);
  const avg = numericValues.length ? sum / numericValues.length : 0;
  const max = numericValues.length ? Math.max(...numericValues) : 0;
  const periodStart = dateValues.length ? dateValues.reduce((a, b) => (a < b ? a : b)) : null;
  const periodEnd = dateValues.length ? dateValues.reduce((a, b) => (a > b ? a : b)) : null;

  const cards = [
    { label: "Registros filtrados", value: fmtNumber(totalRows) },
    { label: "Registros totais", value: fmtNumber(allRowsCount) },
    { label: "Colunas", value: fmtNumber(totalCols) },
    { label: "Valor Total", value: fmtCurrency(sum) },
    { label: "Valor Médio", value: fmtCurrency(avg) },
    { label: "Maior Valor", value: fmtCurrency(max) },
    {
      label: "Período filtrado",
      value: periodStart && periodEnd ? `${new Intl.DateTimeFormat("pt-BR").format(periodStart)} - ${new Intl.DateTimeFormat("pt-BR").format(periodEnd)}` : "Não identificado",
    },
  ];

  kpisContainer.innerHTML = cards
    .map((card) => `<article class="kpi-card"><h3>${card.label}</h3><p>${card.value}</p></article>`)
    .join("");

  kpisContainer.classList.remove("hidden");
};

const renderInsights = (stats) => {
  insightsContainer.innerHTML = [
    { title: "Coluna de Data", value: stats.dateCol || "Não detectada" },
    { title: "Coluna de Valor", value: stats.valueCol || "Não detectada" },
    { title: "Maior UF (Qtd)", value: stats.topUf || "Sem dados" },
    { title: "Maior Status (Qtd)", value: stats.topStatus || "Sem dados" },
  ]
    .map((item) => `<article class="insight-card"><h4>${item.title}</h4><p>${item.value}</p></article>`)
    .join("");
  insightsContainer.classList.remove("hidden");
};

const renderSummaryTable = (countEntries, sumEntries) => {
  const total = countEntries.reduce((acc, [, count]) => acc + count, 0) || 1;
  const sumMap = new Map(sumEntries);

  summaryTableBody.innerHTML = countEntries
    .map(([category, count]) => {
      const share = (count / total) * 100;
      const value = sumMap.get(category) ?? 0;
      return `<tr><td>${category}</td><td>${fmtNumber(count)}</td><td>${fmtPercent(share)}</td><td>${fmtCurrency(value)}</td></tr>`;
    })
    .join("");
};

const createBarChart = (elementId, labels, values, label, color, horizontal = false, compactY = false) =>
  new Chart(document.getElementById(elementId), {
    type: "bar",
    data: { labels, datasets: [{ label, data: values, backgroundColor: color, borderRadius: 6 }] },
    options: {
      ...chartBaseOptions,
      indexAxis: horizontal ? "y" : "x",
      plugins: { ...chartBaseOptions.plugins, legend: { display: false } },
      scales: {
        y: {
          beginAtZero: true,
          ticks: compactY ? { callback: (value) => new Intl.NumberFormat("pt-BR", { notation: "compact" }).format(Number(value)) } : undefined,
        },
      },
    },
  });

const createCharts = (rows, detected) => {
  destroyCharts();

  const ufData = groupCount(rows, detected.ufCol || detected.statusCol, 12);
  const statusData = groupCount(rows, detected.statusCol || detected.ufCol, 8);
  const timelineData = byMonth(rows, detected.dateCol);
  const monthlyValueData = byMonth(rows, detected.dateCol, detected.valueCol);
  const companyValueData = groupSum(rows, detected.companyCol, detected.valueCol, 10);
  const ufValueData = groupSum(rows, detected.ufCol || detected.statusCol, detected.valueCol, 12);
  const statusValueData = groupSum(rows, detected.statusCol || detected.ufCol, detected.valueCol, 8);
  const rangeData = valueRanges(rows, detected.valueCol);

  charts.ufChart = createBarChart("ufChart", ufData.map(([k]) => k), ufData.map(([, v]) => v), "Registros", "#0f766ecc");
  charts.statusChart = new Chart(document.getElementById("statusChart"), {
    type: "doughnut",
    data: { labels: statusData.map(([k]) => k), datasets: [{ data: statusData.map(([, v]) => v), backgroundColor: palette }] },
    options: { ...chartBaseOptions, cutout: "60%" },
  });

  charts.timelineChart = new Chart(document.getElementById("timelineChart"), {
    type: "line",
    data: {
      labels: timelineData.map(([k]) => k),
      datasets: [{ label: "Registros por mês", data: timelineData.map(([, v]) => v), borderColor: "#0e7490", backgroundColor: "rgba(14,116,144,.15)", fill: true, tension: 0.25, pointRadius: 1.8 }],
    },
    options: chartBaseOptions,
  });

  charts.monthlyValueChart = new Chart(document.getElementById("monthlyValueChart"), {
    type: "line",
    data: {
      labels: monthlyValueData.map(([k]) => k),
      datasets: [{ label: "Valor por mês", data: monthlyValueData.map(([, v]) => v), borderColor: "#f59e0b", backgroundColor: "rgba(245,158,11,.16)", fill: true, tension: 0.25, pointRadius: 1.8 }],
    },
    options: {
      ...chartBaseOptions,
      scales: { y: { beginAtZero: true, ticks: { callback: (value) => new Intl.NumberFormat("pt-BR", { notation: "compact" }).format(Number(value)) } } },
    },
  });

  charts.companyValueChart = createBarChart("companyValueChart", companyValueData.map(([k]) => k), companyValueData.map(([, v]) => v), "Valor acumulado", "#115e59c9", true, true);
  charts.ufValueChart = createBarChart("ufValueChart", ufValueData.map(([k]) => k), ufValueData.map(([, v]) => v), "Valor por UF", "#6366f1c9", false, true);
  charts.statusValueChart = createBarChart("statusValueChart", statusValueData.map(([k]) => k), statusValueData.map(([, v]) => v), "Valor por status", "#ef4444b8", false, true);
  charts.valueRangeChart = createBarChart("valueRangeChart", rangeData.map(([k]) => k), rangeData.map(([, v]) => v), "Registros por faixa", "#14b8a6c9");

  renderSummaryTable(statusData.length ? statusData : ufData, statusValueData.length ? statusValueData : ufValueData);

  const topUf = ufData[0] ? `${ufData[0][0]} (${fmtNumber(ufData[0][1])})` : null;
  const topStatus = statusData[0] ? `${statusData[0][0]} (${fmtNumber(statusData[0][1])})` : null;
  renderInsights({ dateCol: detected.dateCol, valueCol: detected.valueCol, topUf, topStatus });

  reportsContainer.classList.remove("hidden");
};

const renderDashboard = () => {
  const { allRows, filteredRows, detected } = dashboardState;
  renderKpis(filteredRows, allRows, detected);
  createCharts(filteredRows, detected);
  filterResult.textContent = `Mostrando ${fmtNumber(filteredRows.length)} de ${fmtNumber(allRows.length)} registros.`;
};

const applyFiltersToRows = () => {
  const { allRows, detected } = dashboardState;
  if (!allRows.length || !detected) return;

  const query = stripAccents(searchInput.value);
  const selectedUf = ufFilter.value;
  const selectedStatus = statusFilter.value;
  const selectedCompany = stripAccents(companyFilter.value);

  const startDate = startDateFilter.value ? new Date(`${startDateFilter.value}T00:00:00`) : null;
  const endDate = endDateFilter.value ? new Date(`${endDateFilter.value}T23:59:59`) : null;

  dashboardState.filteredRows = allRows.filter((row) => {
    if (selectedUf && cleanText(row[detected.ufCol]) !== selectedUf) return false;
    if (selectedStatus && cleanText(row[detected.statusCol]) !== selectedStatus) return false;

    if (selectedCompany && !stripAccents(row[detected.companyCol]).includes(selectedCompany)) return false;
    if (query && !Object.values(row).some((value) => stripAccents(value).includes(query))) return false;

    if ((startDate || endDate) && detected.dateCol) {
      const currentDate = parseDate(row[detected.dateCol]);
      if (!currentDate) return false;
      if (startDate && currentDate < startDate) return false;
      if (endDate && currentDate > endDate) return false;
    }

    return true;
  });

  renderDashboard();
};

const resetModalBody = () => {
  modalBody.innerHTML = '<canvas id="modalCanvas"></canvas>';
};

const closeModal = () => {
  zoomModal.classList.add("hidden");
  zoomModal.setAttribute("aria-hidden", "true");
  resetModalBody();
};

const openChartModal = (title, chartKey) => {
  const chart = charts[chartKey];
  if (!chart || !chart.canvas) return;

  modalTitle.textContent = title;
  zoomModal.classList.remove("hidden");
  zoomModal.setAttribute("aria-hidden", "false");
  modalBody.innerHTML = `<img class="modal-image" alt="${title}" src="${chart.canvas.toDataURL("image/png", 1)}" />`;
};

const openTableModal = (title, tableId) => {
  const table = document.getElementById(tableId);
  if (!table) return;
  modalTitle.textContent = title;
  zoomModal.classList.remove("hidden");
  zoomModal.setAttribute("aria-hidden", "false");
  modalBody.innerHTML = `<div class="table-wrapper"><table class="modal-table">${table.innerHTML}</table></div>`;
};

const updateAuthView = () => {
  const user = authState.currentUser;
  if (!user) {
    loginView.classList.remove("hidden");
    appView.classList.add("hidden");
    return;
  }

  loginView.classList.add("hidden");
  appView.classList.remove("hidden");
  sessionInfo.textContent = `Usuário: ${user.username} | Permissões: ${user.permissions.join(", ")}`;

  const canManageUsers = hasPermission("manage_users");
  const canViewDashboard = hasPermission("view_dashboard");

  adminToggle.classList.toggle("hidden", !canManageUsers);
  adminSection.classList.add("hidden");

  dashboardArea.classList.toggle("hidden", !canViewDashboard);
  accessDenied.classList.toggle("hidden", canViewDashboard);

  if (canManageUsers) renderUsersTable();
};

const renderUsersTable = () => {
  const users = getUsers();
  const currentId = authState.currentUser?.id;

  usersTableBody.innerHTML = users
    .map((u) => {
      const perms = u.permissions.join(", ") || "-";
      const self = u.id === currentId;
      return `<tr>
        <td>${u.username}${self ? " (você)" : ""}</td>
        <td>${u.active ? "Ativo" : "Inativo"}</td>
        <td>${perms}</td>
        <td>
          <button class="zoom-btn" type="button" data-action="toggle" data-id="${u.id}">${u.active ? "Desativar" : "Ativar"}</button>
          <button class="zoom-btn" type="button" data-action="delete" data-id="${u.id}">Excluir</button>
        </td>
      </tr>`;
    })
    .join("");
};

const login = (username, password) => {
  const users = getUsers();
  const found = users.find((u) => u.username.toLowerCase() === username.toLowerCase());
  if (!found) return { ok: false, message: "Usuário não encontrado." };
  if (!found.active) return { ok: false, message: "Usuário inativo. Fale com o administrador." };

  const expected = fromBase64(found.password);
  if (expected !== password) return { ok: false, message: "Senha inválida." };

  const session = { id: found.id, username: found.username, permissions: found.permissions };
  setSession(session);
  authState.currentUser = session;
  return { ok: true };
};

const restoreSession = () => {
  const session = getSession();
  if (!session) return;
  const user = getUsers().find((u) => u.id === session.id && u.active);
  if (!user) {
    clearSession();
    return;
  }
  authState.currentUser = { id: user.id, username: user.username, permissions: user.permissions };
};

const logout = () => {
  clearSession();
  authState.currentUser = null;
  dashboardState.allRows = [];
  dashboardState.filteredRows = [];
  dashboardState.detected = null;
  kpisContainer.classList.add("hidden");
  filtersContainer.classList.add("hidden");
  insightsContainer.classList.add("hidden");
  reportsContainer.classList.add("hidden");
  fileName.textContent = "Nenhum arquivo selecionado.";
  updateAuthView();
};

const parseWorkbook = async (file) => {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: null, raw: false });
  if (!matrix.length) return [];

  const headers = matrix[0].map((header, index) => normalizeHeader(header, index));
  return matrix.slice(1).map((row) => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index] ?? null;
    });
    return item;
  });
};

const bindEvents = () => {
  loginForm.addEventListener("submit", (event) => {
    event.preventDefault();
    loginMessage.textContent = "";

    const result = login(loginUsername.value, loginPassword.value);
    if (!result.ok) {
      loginMessage.textContent = result.message;
      return;
    }

    loginForm.reset();
    updateAuthView();
  });

  logoutBtn.addEventListener("click", logout);

  adminToggle.addEventListener("click", () => {
    adminSection.classList.toggle("hidden");
    adminMessage.textContent = "";
  });

  createUserForm.addEventListener("submit", (event) => {
    event.preventDefault();
    adminMessage.textContent = "";

    const username = cleanText(newUsername.value);
    const password = newPassword.value;
    const permissions = [...createUserForm.querySelectorAll('input[name="permission"]:checked')].map((input) => input.value);

    if (!username) {
      adminMessage.textContent = "Informe o nome do usuário.";
      return;
    }
    if (password.length < 4) {
      adminMessage.textContent = "A senha deve ter no mínimo 4 caracteres.";
      return;
    }
    if (!permissions.length) {
      adminMessage.textContent = "Selecione ao menos uma permissão.";
      return;
    }

    const users = getUsers();
    const exists = users.some((u) => u.username.toLowerCase() === username.toLowerCase());
    if (exists) {
      adminMessage.textContent = "Este usuário já existe.";
      return;
    }

    users.push({
      id: crypto.randomUUID(),
      username,
      password: toBase64(password),
      active: true,
      permissions,
      createdAt: new Date().toISOString(),
    });

    saveUsers(users);
    createUserForm.reset();
    createUserForm.querySelector('input[value="view_dashboard"]').checked = true;
    adminMessage.style.color = "#166534";
    adminMessage.textContent = "Usuário criado com sucesso.";
    renderUsersTable();
  });

  usersTableBody.addEventListener("click", (event) => {
    const btn = event.target.closest("button[data-action]");
    if (!btn) return;

    const action = btn.dataset.action;
    const id = btn.dataset.id;
    const users = getUsers();
    const idx = users.findIndex((u) => u.id === id);
    if (idx < 0) return;

    if (id === authState.currentUser?.id) {
      adminMessage.style.color = "#b91c1c";
      adminMessage.textContent = "Você não pode alterar ou excluir seu próprio usuário por aqui.";
      return;
    }

    if (action === "toggle") users[idx].active = !users[idx].active;
    if (action === "delete") users.splice(idx, 1);

    saveUsers(users);
    renderUsersTable();
  });

  fileInput.addEventListener("change", async (event) => {
    const [file] = event.target.files || [];
    if (!file || !hasPermission("view_dashboard")) return;

    try {
      fileName.textContent = `Arquivo selecionado: ${file.name}`;
      const rows = await parseWorkbook(file);
      if (!rows.length) {
        fileName.textContent = "A planilha está vazia.";
        return;
      }

      const headers = Object.keys(rows[0]);
      const detected = inferColumns(rows, headers);

      dashboardState.allRows = rows;
      dashboardState.filteredRows = [...rows];
      dashboardState.detected = detected;

      initFilters(rows, detected);
      applyFiltersToRows();
    } catch (error) {
      fileName.textContent = `Falha ao processar arquivo: ${error.message}`;
    }
  });

  [searchInput, ufFilter, statusFilter, companyFilter, startDateFilter, endDateFilter].forEach((el) => {
    el.addEventListener("input", applyFiltersToRows);
    el.addEventListener("change", applyFiltersToRows);
  });

  clearFilters.addEventListener("click", () => {
    searchInput.value = "";
    ufFilter.value = "";
    statusFilter.value = "";
    companyFilter.value = "";
    startDateFilter.value = "";
    endDateFilter.value = "";
    applyFiltersToRows();
  });

  document.querySelectorAll(".zoomable").forEach((panel) => {
    panel.addEventListener("click", (event) => {
      const targetButton = event.target.closest("button");
      if (targetButton && !targetButton.classList.contains("zoom-btn")) return;

      const title = panel.dataset.title || "Visualização Ampliada";
      if (panel.dataset.chartKey) openChartModal(title, panel.dataset.chartKey);
      if (panel.dataset.table) openTableModal(title, panel.dataset.table);
    });
  });

  closeModalBtn.addEventListener("click", closeModal);
  modalBackdrop.addEventListener("click", closeModal);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !zoomModal.classList.contains("hidden")) closeModal();
  });
};

const init = () => {
  ensureDefaultAdmin();
  restoreSession();
  bindEvents();
  updateAuthView();
};

init();
