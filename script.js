const fileInput = document.getElementById('excelFile');
const dateInput = document.getElementById('lettingDate');
const generateBtn = document.getElementById('generateBtn');
const printBtn = document.getElementById('printBtn');
const statusBox = document.getElementById('status');

const FIELD_MAP = {
  callOrder: 'CALLORDER',
  proposal: 'PROPOSAL_NM',
  bidders: 'Bidders',
  holders: 'Proposal Holders',
  overUnder: '%OverUnder',
  percentGroup: 'Percent Group',
  estimate: 'PROPOSALITEMTOTAL',
  siteUse: 'Site Use'
};

const CATEGORY_META = [
  { key: 0, label: 'Greater Than 10%\nUnder the Estimate', color: '#006400' },
  { key: 1, label: 'Less Than 10%\nUnder the Estimate', color: '#23a023' },
  { key: 2, label: 'No Bids Received', color: '#d9d9d9' },
  { key: 3, label: 'Less Than 10% Over\nthe Estimate', color: '#d67882' },
  { key: 4, label: 'Greater Than 10%\nOver the Estimate', color: '#d00000' }
];

function setDefaultDate() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  dateInput.value = `${y}-${m}-${d}`;
}

function updateStatus(message, isError = false) {
  statusBox.textContent = message;
  statusBox.style.color = isError ? '#b42318' : '#3e4b58';
}

function formatDateForTitle(dateString) {
  if (!dateString) return 'Letting Summary';
  const date = new Date(`${dateString}T12:00:00`);
  return new Intl.DateTimeFormat('en-US', {
    month: 'long',
    day: 'numeric',
    year: 'numeric'
  }).format(date) + ' Letting Summary';
}

function formatCurrency(value) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0
  }).format(value || 0);
}

function formatPercent(value) {
  const pct = Math.round((Number(value) || 0) * 100);
  return `${pct}%`;
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  if (value == null || value === '') return 0;
  return Number(String(value).replace(/[$,%\s,]/g, '')) || 0;
}

function normalizeRows(rawRows) {
  return rawRows
    .map((row) => ({
      callOrder: row[FIELD_MAP.callOrder],
      proposal: String(row[FIELD_MAP.proposal] ?? '').trim(),
      bidders: parseNumber(row[FIELD_MAP.bidders]),
      holders: parseNumber(row[FIELD_MAP.holders]),
      overUnder: parseNumber(row[FIELD_MAP.overUnder]),
      percentGroup: parseNumber(row[FIELD_MAP.percentGroup]),
      estimate: parseNumber(row[FIELD_MAP.estimate]),
      siteUse: String(row[FIELD_MAP.siteUse] ?? '---').trim() || '---'
    }))
    .filter((row) => row.proposal);
}

function sortRows(rows) {
  return [...rows].sort((a, b) => String(a.callOrder).localeCompare(String(b.callOrder)));
}

function validateColumns(row) {
  const missing = Object.values(FIELD_MAP).filter((field) => !(field in row));
  if (missing.length) {
    throw new Error(`Missing required column(s): ${missing.join(', ')}`);
  }
}

function computeSummary(rows) {
  const projects = rows.length;
  const bidsReceived = rows.reduce((sum, row) => sum + row.bidders, 0);
  const proposalsIssued = rows.reduce((sum, row) => sum + row.holders, 0);
  const estimateTotal = rows.reduce((sum, row) => sum + row.estimate, 0);
  const lowBidTotal = rows.reduce((sum, row) => sum + row.estimate * (1 + row.overUnder), 0);

  const categoryCounts = CATEGORY_META.map((category) => rows.filter((row) => row.percentGroup === category.key).length);

  return {
    projects,
    bidsReceived,
    proposalsIssued,
    estimateTotal,
    lowBidTotal,
    categoryCounts
  };
}

function renderMetrics(summary, dateString) {
  document.getElementById('reportTitle').textContent = formatDateForTitle(dateString);
  document.getElementById('projectsValue').textContent = summary.projects;
  document.getElementById('bidsValue').textContent = summary.bidsReceived;
  document.getElementById('proposalsValue').textContent = summary.proposalsIssued;
  document.getElementById('lowBidValue').textContent = formatCurrency(summary.lowBidTotal);
  document.getElementById('estimateValue').textContent = formatCurrency(summary.estimateTotal);
  document.getElementById('grandProposals').textContent = summary.proposalsIssued;
  document.getElementById('grandBids').textContent = summary.bidsReceived;
}

function renderTable(rows) {
  const tbody = document.getElementById('summaryBody');
  tbody.innerHTML = '';

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(row.proposal)}</td>
      <td>${escapeHtml(row.siteUse)}</td>
      <td>${row.holders}</td>
      <td>${row.bidders}</td>
      <td>${formatPercent(row.overUnder)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function fitTableToPageBottom() {
  const reportPage = document.getElementById('reportPage');
  const table = document.querySelector('.summary-table');
  const tbody = document.getElementById('summaryBody');
  if (!reportPage || !table || !tbody) return;

  const rows = Array.from(tbody.querySelectorAll('tr'));
  if (!rows.length) return;

  rows.forEach((row) => {
    row.style.height = '';
    row.querySelectorAll('td').forEach((cell) => {
      cell.style.height = '';
    });
  });

  const pageRect = reportPage.getBoundingClientRect();
  const tableRect = table.getBoundingClientRect();
  const theadHeight = table.querySelector('thead')?.getBoundingClientRect().height || 0;
  const tfootHeight = table.querySelector('tfoot')?.getBoundingClientRect().height || 0;
  const bottomGap = 8;

  const availableBodyHeight =
    pageRect.bottom - tableRect.top - theadHeight - tfootHeight - bottomGap;

  if (availableBodyHeight <= 0) return;

  const rowHeight = Math.max(12, availableBodyHeight / rows.length);
  rows.forEach((row) => {
    row.style.height = `${rowHeight}px`;
    row.querySelectorAll('td').forEach((cell) => {
      cell.style.height = `${rowHeight}px`;
    });
  });
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function renderLegend() {
  const legend = document.getElementById('customLegend');
  legend.innerHTML = '';

  CATEGORY_META.forEach((item) => {
    const swatch = document.createElement('div');
    swatch.className = 'legend-swatch';
    swatch.style.background = item.color;

    const text = document.createElement('div');
    text.className = 'legend-text';
    text.innerHTML = item.label.replaceAll('\n', '<br>');

    legend.appendChild(swatch);
    legend.appendChild(text);
  });
}

function polarToCartesian(cx, cy, radius, angle) {
  return {
    x: cx + Math.cos(angle) * radius,
    y: cy + Math.sin(angle) * radius
  };
}

function makeSlicePath(cx, cy, radius, startAngle, endAngle) {
  const start = polarToCartesian(cx, cy, radius, startAngle);
  const end = polarToCartesian(cx, cy, radius, endAngle);
  const largeArc = endAngle - startAngle > Math.PI ? 1 : 0;
  return [
    `M ${cx} ${cy}`,
    `L ${start.x} ${start.y}`,
    `A ${radius} ${radius} 0 ${largeArc} 1 ${end.x} ${end.y}`,
    'Z'
  ].join(' ');
}

function addNumberLabel(parent, value, left, top) {
  if (!value) return;
  const div = document.createElement('div');
  div.className = 'pie-number';
  div.textContent = String(value);
  div.style.left = `${left}px`;
  div.style.top = `${top}px`;
  parent.appendChild(div);
}

function renderPie(categoryCounts) {
  const pieStage = document.getElementById('pieStage');
  if (!pieStage) return;

  pieStage.innerHTML = '';

  const width = 180;
  const height = 140;
  const cx = 84;
  const cy = 68;
  const r = 50;

  const total = categoryCounts.reduce((a, b) => a + b, 0);

  const svgNS = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(svgNS, 'svg');
  svg.setAttribute('width', width);
  svg.setAttribute('height', height);
  svg.setAttribute('viewBox', `0 0 ${width} ${height}`);

  function polarToCartesian(centerX, centerY, radius, angleInRadians) {
    return {
      x: centerX + radius * Math.cos(angleInRadians),
      y: centerY + radius * Math.sin(angleInRadians)
    };
  }

  function makeSlicePath(centerX, centerY, radius, startAngle, endAngle) {
    const start = polarToCartesian(centerX, centerY, radius, startAngle);
    const end = polarToCartesian(centerX, centerY, radius, endAngle);
    const largeArcFlag = endAngle - startAngle > Math.PI ? 1 : 0;

    return [
      `M ${centerX} ${centerY}`,
      `L ${start.x} ${start.y}`,
      `A ${radius} ${radius} 0 ${largeArcFlag} 1 ${end.x} ${end.y}`,
      'Z'
    ].join(' ');
  }

  if (total > 0) {
    let startAngle = -Math.PI / 2;

    // Tableau-like visual order:
    // dark red -> light red -> light green -> dark green -> gray
    const drawOrder = [4, 3, 1, 0, 2];

    drawOrder.forEach((index) => {
      const count = categoryCounts[index] || 0;
      if (count <= 0) return;

      const sweep = (count / total) * Math.PI * 2;
      const endAngle = startAngle + sweep;

      const path = document.createElementNS(svgNS, 'path');
      path.setAttribute('d', makeSlicePath(cx, cy, r, startAngle, endAngle));
      path.setAttribute('fill', CATEGORY_META[index].color);
      svg.appendChild(path);

      startAngle = endAngle;
    });
  } else {
    const circle = document.createElementNS(svgNS, 'circle');
    circle.setAttribute('cx', cx);
    circle.setAttribute('cy', cy);
    circle.setAttribute('r', r);
    circle.setAttribute('fill', '#d9d9d9');
    svg.appendChild(circle);
  }

  pieStage.appendChild(svg);

  const labelPositions = {
    4: { left: 10, top: 10 },   // dark red = 8, upper-left
    3: { left: 24, top: 86 },   // light red = 4, lower-left
    1: { left: 60, top: 118 },  // light green = 3, bottom
    0: { left: 154, top: 49 }   // dark green = 16, right
  };

  [4, 3, 1, 0].forEach((index) => {
    const value = categoryCounts[index] || 0;
    const pos = labelPositions[index];
    if (!value || !pos) return;

    const label = document.createElement('div');
    label.className = 'pie-label';
    label.textContent = value;
    label.style.left = `${pos.left}px`;
    label.style.top = `${pos.top}px`;
    pieStage.appendChild(label);
  });
}


async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

async function generateReport() {
  const file = fileInput.files[0];
  if (!file) {
    updateStatus('Please upload the Excel file first.', true);
    return;
  }

  try {
    updateStatus('Reading Excel file...');
    const rawRows = await readWorkbook(file);
    if (!rawRows.length) throw new Error('The worksheet is empty.');

    validateColumns(rawRows[0]);
    const rows = sortRows(normalizeRows(rawRows));
    const summary = computeSummary(rows);

    renderMetrics(summary, dateInput.value);
    renderTable(rows);
    fitTableToPageBottom();
    renderLegend();
    renderPie(summary.categoryCounts);

    updateStatus(`Generated page from ${file.name}.`);
  } catch (error) {
    console.error(error);
    updateStatus(error.message || 'Could not generate the report.', true);
  }
}

generateBtn.addEventListener('click', generateReport);
printBtn.addEventListener('click', () => window.print());
window.addEventListener('resize', fitTableToPageBottom);
window.addEventListener('beforeprint', fitTableToPageBottom);
setDefaultDate();
renderLegend();
renderPie([16, 3, 0, 4, 8]);
fitTableToPageBottom();
