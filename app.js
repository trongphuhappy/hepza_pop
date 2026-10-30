/**
 * HỘI NGHỊ DOANH NGHIỆP - MÀN HÌNH ĐỒNG BỘ TRỰC TIẾP VỚI GOOGLE FORM / SHEET
 * - Tự động nhận dữ liệu khi khách hàng gửi Form (https://forms.gle/fjCd1NZ43Mv89FGR6)
 * - Hỗ trợ đồng thời cả 2 phương thức:
 *   1. Google Apps Script Web App URL (khi giữ Sheet riêng tư)
 *   2. Đọc trực tiếp từ Google Sheet ID 1aYdZGk3-l-RuVWwrwQjJLoY1Jw38TbTbzuLYq6g3gOY (nếu mở quyền Xem công khai)
 * - Pop sáng tên doanh nghiệp kèm pháo hoa nhẹ (hỗ trợ nhiều người submit đồng thời)
 * - Thanh Progress Bar tăng tiến, đạt đủ 100 doanh nghiệp nổ pháo hoa lớn toàn màn hình và tiếp tục tăng tiến
 */

const DEFAULT_SHEET_ID = '1aYdZGk3-l-RuVWwrwQjJLoY1Jw38TbTbzuLYq6g3gOY';
const DEFAULT_FORM_URL = 'https://forms.gle/fjCd1NZ43Mv89FGR6';
const STORAGE_KEY_CONFIG = 'hepza_conf_config_v3';
const STORAGE_KEY_SEEN = 'hepza_conf_seen_ids_v3';

const APPS_SCRIPT_CODE = `const SHEET_ID = '1aYdZGk3-l-RuVWwrwQjJLoY1Jw38TbTbzuLYq6g3gOY';
const DEFAULT_COMPANY_COL = 2; // Mặc định Cột B (ngay sau Dấu thời gian)

function detectCompanyColIndex(headers, fallbackCol) {
  if (!headers || !headers.length) return fallbackCol - 1;
  const keywords = ['tên doanh nghiệp', 'doanh nghiệp', 'tên công ty', 'công ty', 'tên đơn vị', 'đơn vị', 'company'];
  for (let kw of keywords) {
    for (let i = 1; i < headers.length; i++) {
      const h = (headers[i] || '').toString().toLowerCase().trim();
      if (h.indexOf(kw) !== -1) return i;
    }
  }
  return Math.max(0, fallbackCol - 1);
}

function doGet(e) {
  try {
    let ss;
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (err) { ss = null; }
    if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);

    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const items = [];
    let headers = [];

    if (lastRow >= 1 && lastCol >= 1) {
      const data = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      headers = data[0] || [];
      let colIdx;
      if (e && e.parameter && e.parameter.col && e.parameter.col !== 'auto') {
        colIdx = Math.max(0, (parseInt(e.parameter.col, 10) || DEFAULT_COMPANY_COL) - 1);
      } else {
        colIdx = detectCompanyColIndex(headers, DEFAULT_COMPANY_COL);
      }

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const timestamp = (row[0] || '').toString().trim();
        let companyName = (row[colIdx] || '').toString().trim();
        if (!companyName) {
          for (let c = 1; c < row.length; c++) {
            if (row[c] && row[c].toString().trim() !== '') {
              companyName = row[c].toString().trim();
              break;
            }
          }
        }
        if (!companyName) continue;

        items.push({
          rowNumber: i + 1,
          id: 'row_' + (i + 1) + '_' + encodeURIComponent(timestamp) + '_' + encodeURIComponent(companyName),
          timestamp: timestamp,
          companyName: companyName
        });
      }
    }

    const jsonString = JSON.stringify({
      ok: true,
      total: items.length,
      items: items
    });

    if (e && e.parameter && e.parameter.callback) {
      const cb = e.parameter.callback.replace(/[^a-zA-Z0-9_$.]/g, '');
      return ContentService.createTextOutput(cb + '(' + jsonString + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(jsonString)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const errStr = JSON.stringify({ ok: false, error: error.toString() });
    if (e && e.parameter && e.parameter.callback) {
      const cb = e.parameter.callback.replace(/[^a-zA-Z0-9_$.]/g, '');
      return ContentService.createTextOutput(cb + '(' + errStr + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(errStr)
      .setMimeType(ContentService.MimeType.JSON);
  }
}`;

const defaultConfig = {
  scriptUrl: '',
  columnIndex: 'auto',
  pollIntervalSec: 2,
  targetCount: 100,
  eventSubtitle: 'HỘI NGHỊ KẾT NỐI & ĐỒNG HÀNH CÙNG DOANH NGHIỆP',
  eventTitle: 'CHÀO MỪNG CÁC DOANH NGHIỆP THAM DỰ',
  popExistingOnStart: false,
  soundEnabled: true,
  showQR: true
};

let config = loadConfig();
let seenIds = loadSeenIds();
let hasCompletedInitialSync = false;
let hasCelebratedMilestone = false;

let companies = [];
const activePoppingIds = new Set();

let pollTimer = null;
let isFetching = false;

const els = {
  connectionBadge: document.getElementById('connectionBadge'),
  connectionText: document.getElementById('connectionText'),
  lastSyncText: document.getElementById('lastSyncText'),
  eventSubtitleDisplay: document.getElementById('eventSubtitleDisplay'),
  eventTitleDisplay: document.getElementById('eventTitleDisplay'),

  // Progress bar
  totalCount: document.getElementById('totalCount'),
  targetCountDisplay: document.getElementById('targetCountDisplay'),
  progressPercent: document.getElementById('progressPercent'),
  progressBarFill: document.getElementById('progressBarFill'),
  milestoneStatusPill: document.getElementById('milestoneStatusPill'),
  milestoneBanner: document.getElementById('milestoneBanner'),
  milestoneTitle: document.getElementById('milestoneTitle'),

  // Board & QR
  mainStage: document.getElementById('mainStage'),
  boardSection: document.getElementById('boardSection'),
  emptyState: document.getElementById('emptyState'),
  companyGrid: document.getElementById('companyGrid'),
  qrCodeImg: document.getElementById('qrCodeImg'),
  qrFormLink: document.getElementById('qrFormLink'),
  soundIcon: document.getElementById('soundIcon'),

  // Settings Modal
  settingsModal: document.getElementById('settingsModal'),
  inputScriptUrl: document.getElementById('inputScriptUrl'),
  inputColumnIndex: document.getElementById('inputColumnIndex'),
  inputPollInterval: document.getElementById('inputPollInterval'),
  inputTargetCount: document.getElementById('inputTargetCount'),
  inputEventSubtitle: document.getElementById('inputEventSubtitle'),
  inputEventTitle: document.getElementById('inputEventTitle'),
  chkPopExistingOnStart: document.getElementById('chkPopExistingOnStart'),

  // Guide Modal
  guideModal: document.getElementById('guideModal'),
  scriptCodeDisplay: document.getElementById('scriptCodeDisplay'),
  btnCopyScriptCode: document.getElementById('btnCopyScriptCode')
};

/* =========================================================
   KHỞI TẠO
   ========================================================= */
function init() {
  els.scriptCodeDisplay.textContent = APPS_SCRIPT_CODE;
  applyConfigToUI();
  updateProgressBar(0, false);
  bindEvents();
  initFxCanvas();
  startPolling();
}

function loadConfig() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_CONFIG);
    if (!raw) return { ...defaultConfig };
    return { ...defaultConfig, ...JSON.parse(raw) };
  } catch (e) {
    return { ...defaultConfig };
  }
}

function saveConfig() {
  try {
    localStorage.setItem(STORAGE_KEY_CONFIG, JSON.stringify(config));
  } catch (e) {}
}

function loadSeenIds() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_SEEN);
    if (!raw) return new Set();
    return new Set(JSON.parse(raw));
  } catch (e) {
    return new Set();
  }
}

function saveSeenIds() {
  try {
    localStorage.setItem(STORAGE_KEY_SEEN, JSON.stringify(Array.from(seenIds)));
  } catch (e) {}
}

function applyConfigToUI() {
  els.eventSubtitleDisplay.textContent = config.eventSubtitle;
  els.eventTitleDisplay.textContent = config.eventTitle;
  els.targetCountDisplay.textContent = config.targetCount || 100;

  if (config.showQR) {
    els.mainStage.classList.remove('hide-qr');
  } else {
    els.mainStage.classList.add('hide-qr');
  }

  els.soundIcon.textContent = config.soundEnabled ? '🔊' : '🔇';
}

function setConnectionStatus(state, message) {
  els.connectionBadge.className = 'live-indicator';
  if (state === 'connected') {
    els.connectionBadge.classList.add('connected');
  } else if (state === 'error') {
    els.connectionBadge.classList.add('error');
  }
  els.connectionText.textContent = message;
}

/* =========================================================
   ĐỒNG BỘ THỜI GIAN THỰC VỚI GOOGLE SHEET / GOOGLE FORM
   ========================================================= */
function startPolling() {
  if (pollTimer) clearInterval(pollTimer);
  syncFromGoogleSheet();
  const intervalMs = Math.max(1200, (Number(config.pollIntervalSec) || 2) * 1000);
  pollTimer = setInterval(syncFromGoogleSheet, intervalMs);
}

async function syncFromGoogleSheet() {
  if (isFetching) return;
  isFetching = true;

  try {
    let payload = null;
    const scriptUrl = (config.scriptUrl || '').trim();

    if (scriptUrl) {
      payload = await requestAppsScriptData(scriptUrl, config.columnIndex);
    } else {
      // Nếu chưa gắn link Apps Script, tự động thử đọc trực tiếp từ Google Sheet (nếu Sheet đã bật Xem công khai)
      payload = await requestDirectGoogleSheetGviz(DEFAULT_SHEET_ID, config.columnIndex);
    }

    if (!payload || !payload.ok) {
      throw new Error((payload && payload.error) || 'Chưa kết nối được Google Sheet');
    }

    const nowStr = new Date().toLocaleTimeString('vi-VN');
    els.lastSyncText.textContent = `Đồng bộ trực tiếp: ${nowStr}`;
    setConnectionStatus('connected', 'ĐANG ĐỒNG BỘ FORM (LIVE)');

    processIncomingRows(payload.items || []);
  } catch (err) {
    if (!config.scriptUrl) {
      setConnectionStatus('idle', 'CẦN MỞ QUYỀN XEM SHEET HOẶC CÀI SCRIPT');
    } else {
      setConnectionStatus('error', 'KIỂM TRA LẠI LINK APPS SCRIPT');
    }
  } finally {
    isFetching = false;
  }
}

/**
 * Phương thức 1: Đọc qua Google Apps Script Web App (JSONP + Fetch fallback)
 */
function requestAppsScriptData(baseUrl, colSetting) {
  return new Promise((resolve, reject) => {
    const cbName = 'hepzaScriptCb_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
    const script = document.createElement('script');
    let isDone = false;

    const timer = setTimeout(() => {
      if (isDone) return;
      isDone = true;
      cleanup();
      reject(new Error('Timeout'));
    }, 8000);

    function cleanup() {
      clearTimeout(timer);
      delete window[cbName];
      if (script.parentNode) script.parentNode.removeChild(script);
    }

    window[cbName] = function (data) {
      if (isDone) return;
      isDone = true;
      cleanup();
      resolve(data);
    };

    script.onerror = function () {
      if (isDone) return;
      isDone = true;
      cleanup();
      const sep = baseUrl.includes('?') ? '&' : '?';
      fetch(`${baseUrl}${sep}col=${encodeURIComponent(colSetting)}&_t=${Date.now()}`)
        .then((r) => r.json())
        .then(resolve)
        .catch(reject);
    };

    const sep = baseUrl.includes('?') ? '&' : '?';
    script.src = `${baseUrl}${sep}col=${encodeURIComponent(colSetting)}&callback=${cbName}&_t=${Date.now()}`;
    document.body.appendChild(script);
  });
}

/**
 * Phương thức 2: Đọc trực tiếp từ Google Sheet bằng Google Visualization JSONP
 * (Hoạt động ngay lập tức khi Google Sheet để chế độ "Bất kỳ ai có đường liên kết đều có thể xem")
 */
function requestDirectGoogleSheetGviz(sheetId, colSetting) {
  return new Promise((resolve, reject) => {
    const cbName = 'hepzaGvizCb_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
    const script = document.createElement('script');
    let isDone = false;

    const timer = setTimeout(() => {
      if (isDone) return;
      isDone = true;
      cleanup();
      reject(new Error('Sheet đang ở chế độ riêng tư'));
    }, 6500);

    function cleanup() {
      clearTimeout(timer);
      delete window[cbName];
      if (script.parentNode) script.parentNode.removeChild(script);
    }

    window[cbName] = function (response) {
      if (isDone) return;
      isDone = true;
      cleanup();

      try {
        if (!response || response.status === 'error' || !response.table) {
          reject(new Error('Invalid gviz response'));
          return;
        }

        const cols = (response.table.cols || []).map((c) => (c && c.label ? c.label.trim() : ''));
        const rows = response.table.rows || [];

        // Xác định cột Tên doanh nghiệp
        let colIdx = 1; // Mặc định Cột B
        if (colSetting && colSetting !== 'auto') {
          colIdx = Math.max(0, (parseInt(colSetting, 10) || 2) - 1);
        } else {
          const keywords = ['tên doanh nghiệp', 'doanh nghiệp', 'tên công ty', 'công ty', 'tên đơn vị', 'đơn vị', 'company'];
          let found = -1;
          for (const kw of keywords) {
            for (let i = 1; i < cols.length; i++) {
              if (cols[i].toLowerCase().includes(kw)) {
                found = i;
                break;
              }
            }
            if (found !== -1) break;
          }
          if (found !== -1) colIdx = found;
        }

        const items = [];
        rows.forEach((r, idx) => {
          const cells = r.c || [];
          const getVal = (cell) => {
            if (!cell) return '';
            if (cell.f !== undefined && cell.f !== null) return String(cell.f).trim();
            if (cell.v !== undefined && cell.v !== null) return String(cell.v).trim();
            return '';
          };

          const timestamp = getVal(cells[0]);
          let companyName = getVal(cells[colIdx]);

          // Bỏ qua nếu hàng đầu tiên vô tình là dòng tiêu đề lặp lại
          if (idx === 0 && companyName.toLowerCase().includes('tên doanh nghiệp')) {
            return;
          }

          if (!companyName) {
            for (let c = 1; c < cells.length; c++) {
              const v = getVal(cells[c]);
              if (v) {
                companyName = v;
                break;
              }
            }
          }

          if (!companyName) return;

          items.push({
            rowNumber: idx + 2,
            id: `row_${idx + 2}_${encodeURIComponent(timestamp)}_${encodeURIComponent(companyName)}`,
            timestamp,
            companyName
          });
        });

        resolve({
          ok: true,
          total: items.length,
          items
        });
      } catch (e) {
        reject(e);
      }
    };

    script.onerror = function () {
      if (isDone) return;
      isDone = true;
      cleanup();
      reject(new Error('Không thể tải trực tiếp từ Google Sheet'));
    };

    script.src = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=responseHandler:${cbName}&_t=${Date.now()}`;
    document.body.appendChild(script);
  });
}

/* =========================================================
   XỬ LÝ DỮ LIỆU MỚI: POP ĐỒNG THỜI & TIẾN ĐỘ 100 DOANH NGHIỆP
   ========================================================= */
function processIncomingRows(items) {
  const validItems = items
    .filter((item) => item.companyName && item.companyName.trim() !== '')
    .map((item, idx) => ({
      ...item,
      orderNumber: idx + 1
    }));

  const prevCount = companies.length;
  const newEntries = [];

  for (const item of validItems) {
    if (!seenIds.has(item.id)) {
      seenIds.add(item.id);
      newEntries.push(item);
    }
  }

  if (newEntries.length > 0) {
    saveSeenIds();
  }

  companies = validItems;

  // Lần đồng bộ đầu tiên khi mới mở trang
  if (!hasCompletedInitialSync) {
    hasCompletedInitialSync = true;
    const target = Number(config.targetCount) || 100;
    if (companies.length >= target) {
      hasCelebratedMilestone = true;
    }

    if (config.popExistingOnStart && newEntries.length > 0) {
      triggerSimultaneousPop(newEntries, prevCount);
    } else {
      renderBoard();
      updateProgressBar(companies.length, false);
    }
    return;
  }

  // Khi khách mời submit form: có bao nhiêu doanh nghiệp mới sẽ pop sáng đồng thời bấy nhiêu
  if (newEntries.length > 0) {
    triggerSimultaneousPop(newEntries, prevCount);
  } else if (validItems.length !== prevCount) {
    renderBoard();
    updateProgressBar(companies.length, false);
  }
}

function triggerSimultaneousPop(newItems, previousTotal) {
  const newIds = newItems.map((item) => item.id);
  newIds.forEach((id) => activePoppingIds.add(id));

  renderBoard();
  els.boardSection.scrollTo({ top: 0, behavior: 'smooth' });

  const currentTotal = companies.length;
  updateProgressBar(currentTotal, true);

  const target = Number(config.targetCount) || 100;
  const crossedMilestone = previousTotal < target && currentTotal >= target && !hasCelebratedMilestone;

  if (crossedMilestone) {
    hasCelebratedMilestone = true;
    triggerMilestone100Celebration(target);
  } else if (config.soundEnabled) {
    playPopChime(Math.min(newItems.length, 5));
  }

  // Hiệu ứng pháo hoa nhẹ ngay tại từng thẻ doanh nghiệp vừa pop lên
  requestAnimationFrame(() => {
    newIds.forEach((id, idx) => {
      const cardEl = els.companyGrid.querySelector(`[data-id="${CSS.escape(id)}"]`);
      if (cardEl) {
        const rect = cardEl.getBoundingClientRect();
        setTimeout(() => {
          spawnLightCardFirework(rect.left + rect.width * 0.5, rect.top + rect.height * 0.45);
        }, idx * 65);
      }
    });
  });

  setTimeout(() => {
    newIds.forEach((id) => activePoppingIds.delete(id));
    newIds.forEach((id) => {
      const cardEl = els.companyGrid.querySelector(`[data-id="${CSS.escape(id)}"]`);
      if (cardEl) cardEl.classList.remove('is-popping');
    });
  }, 4500);
}

/* =========================================================
   THANH PROGRESS BAR & NỔ PHÁO LỚN MỐC 100
   ========================================================= */
function updateProgressBar(count, animateBump = false) {
  const target = Math.max(1, Number(config.targetCount) || 100);
  els.totalCount.textContent = count;
  els.targetCountDisplay.textContent = target;

  if (animateBump) {
    els.totalCount.classList.add('bump');
    setTimeout(() => els.totalCount.classList.remove('bump'), 300);
  }

  const rawPercent = Math.round((count / target) * 100);
  els.progressPercent.textContent = `${rawPercent}%`;

  const barWidth = Math.min(100, (count / target) * 100);
  els.progressBarFill.style.width = `${barWidth}%`;

  if (count >= target) {
    els.milestoneStatusPill.textContent = `🎉 Đã đạt mốc ${target} Doanh nghiệp & đang tiếp tục tăng tiến`;
    els.milestoneStatusPill.classList.add('reached');
  } else {
    els.milestoneStatusPill.textContent = `Mục tiêu: ${target} Doanh nghiệp`;
    els.milestoneStatusPill.classList.remove('reached');
  }
}

function triggerMilestone100Celebration(targetCount = 100) {
  els.milestoneTitle.textContent = `CHÚC MỪNG ĐÃ ĐẠT ${targetCount} DOANH NGHIỆP THAM DỰ!`;
  els.milestoneBanner.classList.add('show');

  if (config.soundEnabled) {
    playMilestoneGrandFanfare();
  }

  spawnGrandFullscreenFireworks();

  setTimeout(() => {
    els.milestoneBanner.classList.remove('show');
  }, 6500);
}

/* =========================================================
   HIỂN THỊ TÊN DOANH NGHIỆP TRÊN MÀN HÌNH
   ========================================================= */
function renderBoard() {
  if (companies.length === 0) {
    els.emptyState.style.display = 'flex';
    els.companyGrid.style.display = 'none';
    els.companyGrid.innerHTML = '';
    return;
  }

  els.emptyState.style.display = 'none';
  els.companyGrid.style.display = 'grid';

  // Doanh nghiệp mới nhất hiển thị lên đầu
  const reversed = [...companies].reverse();

  els.companyGrid.innerHTML = reversed
    .map((item, idx) => {
      const orderNum = item.orderNumber || companies.length - idx;
      const orderStr = String(orderNum).padStart(2, '0');
      const isPopping = activePoppingIds.has(item.id);

      return `
        <article class="enterprise-card ${isPopping ? 'is-popping' : ''}" data-id="${escapeHtml(item.id)}">
          <div class="card-index">#${orderStr}</div>
          <div class="card-body">
            <h3 class="card-title">${escapeHtml(item.companyName)}</h3>
          </div>
        </article>
      `;
    })
    .join('');
}

/* =========================================================
   HỆ THỐNG PHÁO HOA (PHÁO HOA NHẸ TẠI THẺ & PHÁO HOA LỚN MỐC 100)
   ========================================================= */
let fxParticles = [];
let fxAnimId = null;

function initFxCanvas() {
  const canvas = document.getElementById('fxCanvas');
  const resize = () => {
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
  };
  window.addEventListener('resize', resize);
  resize();
}

function spawnLightCardFirework(x, y) {
  const colors = ['#e60028', '#1d4ed8', '#f59e0b', '#10b981', '#38bdf8'];
  const count = 26;

  for (let i = 0; i < count; i++) {
    const angle = (Math.PI * 2 * i) / count + (Math.random() - 0.5) * 0.25;
    const speed = 2.2 + Math.random() * 5.2;
    fxParticles.push({
      x,
      y,
      vx: Math.cos(angle) * speed,
      vy: Math.sin(angle) * speed - 1.2,
      size: 2.5 + Math.random() * 3.2,
      color: colors[Math.floor(Math.random() * colors.length)],
      gravity: 0.12,
      drag: 0.96,
      alpha: 1,
      decay: 0.022 + Math.random() * 0.015,
      shape: Math.random() > 0.4 ? 'circle' : 'rect'
    });
  }

  startFxLoop();
}

function spawnGrandFullscreenFireworks() {
  const canvas = document.getElementById('fxCanvas');
  const palette = ['#e60028', '#2563eb', '#f59e0b', '#10b981', '#ec4899', '#06b6d4', '#fbbf24'];

  for (let wave = 0; wave < 14; wave++) {
    setTimeout(() => {
      const targetX = canvas.width * (0.15 + Math.random() * 0.7);
      const targetY = canvas.height * (0.15 + Math.random() * 0.45);
      createLargeBurst(targetX, targetY, palette);
    }, wave * 320);
  }
}

function createLargeBurst(x, y, palette) {
  const mainColor = palette[Math.floor(Math.random() * palette.length)];
  const secondColor = palette[Math.floor(Math.random() * palette.length)];
  const count = 78;

  for (let i = 0; i < count; i++) {
    const angle = (Math.PI * 2 * i) / count + (Math.random() - 0.5) * 0.15;
    const speed = 3 + Math.random() * 9.5;
    fxParticles.push({
      x,
      y,
      vx: Math.cos(angle) * speed,
      vy: Math.sin(angle) * speed,
      size: 3.5 + Math.random() * 4.5,
      color: i % 2 === 0 ? mainColor : secondColor,
      gravity: 0.11,
      drag: 0.975,
      alpha: 1,
      decay: 0.011 + Math.random() * 0.01,
      shape: i % 3 === 0 ? 'rect' : 'circle'
    });
  }

  startFxLoop();
}

function startFxLoop() {
  if (fxAnimId) return;
  const canvas = document.getElementById('fxCanvas');
  const ctx = canvas.getContext('2d');

  const loop = () => {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    fxParticles = fxParticles.filter((p) => p.alpha > 0.03);

    for (const p of fxParticles) {
      p.x += p.vx;
      p.y += p.vy;
      p.vx *= p.drag;
      p.vy *= p.drag;
      p.vy += p.gravity;
      p.alpha -= p.decay;

      ctx.save();
      ctx.globalAlpha = Math.max(0, p.alpha);
      ctx.fillStyle = p.color;

      if (p.shape === 'circle') {
        ctx.beginPath();
        ctx.arc(p.x, p.y, p.size, 0, Math.PI * 2);
        ctx.fill();
      } else {
        ctx.translate(p.x, p.y);
        ctx.rotate(p.alpha * 6);
        ctx.fillRect(-p.size, -p.size * 0.6, p.size * 2, p.size * 1.2);
      }
      ctx.restore();
    }

    if (fxParticles.length > 0) {
      fxAnimId = requestAnimationFrame(loop);
    } else {
      fxAnimId = null;
      ctx.clearRect(0, 0, canvas.width, canvas.height);
    }
  };

  fxAnimId = requestAnimationFrame(loop);
}

/* =========================================================
   ÂM THANH HỘI NGHỊ HIỆN ĐẠI (WEB AUDIO API)
   ========================================================= */
let audioCtx = null;

function ensureAudioCtx() {
  const AudioContextClass = window.AudioContext || window.webkitAudioContext;
  if (!AudioContextClass) return null;
  if (!audioCtx) audioCtx = new AudioContextClass();
  if (audioCtx.state === 'suspended') audioCtx.resume();
  return audioCtx;
}

function playPopChime(simultaneousCount = 1) {
  try {
    const ctx = ensureAudioCtx();
    if (!ctx) return;

    const now = ctx.currentTime;
    const baseNotes = [587.33, 880.0, 1174.66];
    const extraNotes = simultaneousCount > 1 ? [1479.98] : [];
    const allNotes = [...baseNotes, ...extraNotes];

    allNotes.forEach((freq, idx) => {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      const startTime = now + idx * 0.065;

      osc.type = 'sine';
      osc.frequency.setValueAtTime(freq, startTime);

      gain.gain.setValueAtTime(0.0001, startTime);
      gain.gain.exponentialRampToValueAtTime(0.14, startTime + 0.025);
      gain.gain.exponentialRampToValueAtTime(0.0001, startTime + 0.65);

      osc.connect(gain);
      gain.connect(ctx.destination);

      osc.start(startTime);
      osc.stop(startTime + 0.7);
    });
  } catch (e) {}
}

function playMilestoneGrandFanfare() {
  try {
    const ctx = ensureAudioCtx();
    if (!ctx) return;

    const now = ctx.currentTime;
    const notes = [
      { f: 523.25, t: 0.0, d: 0.4 },
      { f: 659.25, t: 0.12, d: 0.4 },
      { f: 783.99, t: 0.24, d: 0.4 },
      { f: 1046.5, t: 0.36, d: 0.9 },
      { f: 1318.51, t: 0.52, d: 1.4 }
    ];

    notes.forEach((n) => {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'triangle';
      osc.frequency.setValueAtTime(n.f, now + n.t);

      gain.gain.setValueAtTime(0.0001, now + n.t);
      gain.gain.exponentialRampToValueAtTime(0.2, now + n.t + 0.04);
      gain.gain.exponentialRampToValueAtTime(0.0001, now + n.t + n.d);

      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start(now + n.t);
      osc.stop(now + n.t + n.d + 0.05);
    });
  } catch (e) {}
}

/* =========================================================
   SỰ KIỆN ĐIỀU KHIỂN CÀI ĐẶT
   ========================================================= */
function bindEvents() {
  // Click vào trạng thái kết nối để mở cài đặt nếu chưa cấu hình
  els.connectionBadge.addEventListener('click', openSettingsModal);

  // Bật/Tắt âm thanh
  document.getElementById('btnToggleSound').addEventListener('click', () => {
    config.soundEnabled = !config.soundEnabled;
    saveConfig();
    applyConfigToUI();
    if (config.soundEnabled) playPopChime(1);
  });

  // Ẩn/Hiện QR
  document.getElementById('btnToggleQR').addEventListener('click', () => {
    config.showQR = !config.showQR;
    saveConfig();
    applyConfigToUI();
  });

  // Toàn màn hình
  document.getElementById('btnFullscreen').addEventListener('click', toggleFullscreen);

  // Mở/Đóng Cài đặt
  document.getElementById('btnSettings').addEventListener('click', openSettingsModal);
  document.getElementById('btnCloseSettings').addEventListener('click', closeSettingsModal);
  document.getElementById('btnCancelSettings').addEventListener('click', closeSettingsModal);

  // Lưu Cài đặt
  document.getElementById('btnSaveSettings').addEventListener('click', () => {
    config.scriptUrl = els.inputScriptUrl.value.trim();
    config.columnIndex = els.inputColumnIndex.value || 'auto';
    config.pollIntervalSec = parseFloat(els.inputPollInterval.value) || 2;
    config.targetCount = Math.max(1, parseInt(els.inputTargetCount.value, 10) || 100);
    config.eventSubtitle = els.inputEventSubtitle.value.trim() || defaultConfig.eventSubtitle;
    config.eventTitle = els.inputEventTitle.value.trim() || defaultConfig.eventTitle;
    config.popExistingOnStart = els.chkPopExistingOnStart.checked;

    saveConfig();
    applyConfigToUI();
    updateProgressBar(companies.length, false);
    closeSettingsModal();
    startPolling();
  });

  // Làm mới & Pop lại danh sách từ Sheet
  document.getElementById('btnResetSeenHistory').addEventListener('click', () => {
    seenIds.clear();
    saveSeenIds();
    hasCelebratedMilestone = false;
    hasCompletedInitialSync = true;
    closeSettingsModal();
    syncFromGoogleSheet();
  });

  // Mở/Đóng Hướng dẫn
  document.getElementById('btnSwitchToGuide').addEventListener('click', () => {
    closeSettingsModal();
    openGuideModal();
  });
  document.getElementById('btnCloseGuide').addEventListener('click', closeGuideModal);
  document.getElementById('btnGuideToSettings').addEventListener('click', () => {
    closeGuideModal();
    openSettingsModal();
  });

  // Copy mã Google Apps Script
  els.btnCopyScriptCode.addEventListener('click', () => {
    navigator.clipboard.writeText(APPS_SCRIPT_CODE).then(() => {
      const oldText = els.btnCopyScriptCode.textContent;
      els.btnCopyScriptCode.textContent = '✅ Đã Copy mã!';
      setTimeout(() => {
        els.btnCopyScriptCode.textContent = oldText;
      }, 2000);
    });
  });

  window.addEventListener('keydown', (e) => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') {
      return;
    }
    if (e.key === 'Escape') {
      closeSettingsModal();
      closeGuideModal();
    } else if (e.key === 'f' || e.key === 'F') {
      toggleFullscreen();
    }
  });
}

function openSettingsModal() {
  els.inputScriptUrl.value = config.scriptUrl || '';
  els.inputColumnIndex.value = String(config.columnIndex || 'auto');
  els.inputPollInterval.value = String(config.pollIntervalSec || 2);
  els.inputTargetCount.value = String(config.targetCount || 100);
  els.inputEventSubtitle.value = config.eventSubtitle;
  els.inputEventTitle.value = config.eventTitle;
  els.chkPopExistingOnStart.checked = !!config.popExistingOnStart;
  els.settingsModal.classList.add('open');
}

function closeSettingsModal() {
  els.settingsModal.classList.remove('open');
}

function openGuideModal() {
  els.guideModal.classList.add('open');
}

function closeGuideModal() {
  els.guideModal.classList.remove('open');
}

function toggleFullscreen() {
  if (!document.fullscreenElement) {
    document.documentElement.requestFullscreen().catch(() => {});
  } else {
    document.exitFullscreen().catch(() => {});
  }
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

document.addEventListener('DOMContentLoaded', init);
