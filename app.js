/**
 * HỘI NGHỊ DOANH NGHIỆP - BẢNG CHÀO MỪNG & TIẾN ĐỘ ĐĂNG KÝ THỜI GIAN THỰC
 * - Font Sarabun hiện đại, tích hợp Logo Nhà tài trợ (HEPZA, HBA, Techcombank, Viettel)
 * - Pop đồng thời nhiều doanh nghiệp khi submit cùng lúc (kèm hiệu ứng pop sáng & pháo hoa nhẹ tại thẻ)
 * - Thanh Progress Bar tăng tiến, đạt đủ 100 doanh nghiệp nổ pháo hoa lớn toàn màn hình và tiếp tục tăng tiến
 */

const DEFAULT_SHEET_ID = '1aYdZGk3-l-RuVWwrwQjJLoY1Jw38TbTbzuLYq6g3gOY';
const DEFAULT_FORM_URL = 'https://forms.gle/fjCd1NZ43Mv89FGR6';
const STORAGE_KEY_CONFIG = 'hepza_conf_config_v2';
const STORAGE_KEY_SEEN = 'hepza_conf_seen_ids_v2';

const APPS_SCRIPT_CODE = `const SHEET_ID = '1aYdZGk3-l-RuVWwrwQjJLoY1Jw38TbTbzuLYq6g3gOY';
const COMPANY_COLUMN_INDEX = 2; // Cột B (ngay sau Dấu thời gian)

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
      const colParam = e && e.parameter && e.parameter.col ? parseInt(e.parameter.col, 10) : COMPANY_COLUMN_INDEX;
      const colIdx = Math.max(0, (isNaN(colParam) ? COMPANY_COLUMN_INDEX : colParam) - 1);

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const timestamp = (row[0] || '').toString().trim();
        const companyName = (row[colIdx] || '').toString().trim();
        if (!companyName && !timestamp) continue;

        const extraFields = [];
        for (let c = 1; c < row.length; c++) {
          if (c !== colIdx && row[c] && row[c].toString().trim() !== '') {
            extraFields.push({
              label: headers[c] || ('Cột ' + (c + 1)),
              value: row[c].toString().trim()
            });
          }
        }

        items.push({
          rowNumber: i + 1,
          id: 'row_' + (i + 1) + '_' + encodeURIComponent(timestamp) + '_' + encodeURIComponent(companyName),
          timestamp: timestamp,
          companyName: companyName,
          extra: extraFields
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
  columnIndex: 2,
  pollIntervalSec: 2,
  targetCount: 100,
  eventSubtitle: 'HỘI NGHỊ KẾT NỐI & ĐỒNG HÀNH CÙNG DOANH NGHIỆP',
  eventTitle: 'CHÀO MỪNG CÁC DOANH NGHIỆP THAM DỰ',
  formUrl: DEFAULT_FORM_URL,
  popExistingOnStart: false,
  showExtraColumns: true,
  soundEnabled: true,
  showQR: true
};

let config = loadConfig();
let seenIds = loadSeenIds();
let hasCompletedInitialSync = false;
let hasCelebratedMilestone = false;

let companies = [];
// Lưu danh sách ID vừa mới pop để giữ hiệu ứng sáng trong vài giây
const activePoppingIds = new Set();

let pollTimer = null;
let isFetching = false;
let demoCounter = 0;

const DEMO_COMPANIES = [
  { companyName: 'TẬP ĐOÀN CÔNG NGHIỆP - VIỄN THÔNG QUÂN ĐỘI (VIETTEL)', extra: [{ label: 'Đơn vị', value: 'Viettel TP.HCM' }] },
  { companyName: 'NGÂN HÀNG TMCP KỸ THƯƠNG VIỆT NAM (TECHCOMBANK)', extra: [{ label: 'Khối', value: 'Khách hàng Doanh nghiệp' }] },
  { companyName: 'BAN QUẢN LÝ CÁC KCX VÀ CÔNG NGHIỆP TP.HCM (HEPZA)', extra: [{ label: 'Vai trò', value: 'Ban Tổ Chức' }] },
  { companyName: 'HIỆP HỘI CÁC DOANH NGHIỆP KCN TP.HCM (HBA)', extra: [{ label: 'Vai trò', value: 'Đơn vị Đồng hành' }] },
  { companyName: 'CÔNG TY CỔ PHẦN SỮA VIỆT NAM (VINAMILK)', extra: [{ label: 'KCN', value: 'KCN Hiệp Phước' }] },
  { companyName: 'CÔNG TY TNHH INTEL PRODUCTS VIỆT NAM', extra: [{ label: 'Lĩnh vực', value: 'Công nghệ cao' }] },
  { companyName: 'CÔNG TY TNHH NIDEC VIỆT NAM', extra: [{ label: 'Lĩnh vực', value: 'Điện tử chính xác' }] },
  { companyName: 'CÔNG TY CỔ PHẦN TẬP ĐOÀN HÒA PHÁT', extra: [{ label: 'Lĩnh vực', value: 'Công nghiệp sản xuất' }] },
  { companyName: 'CÔNG TY TNHH SAMSUNG ELECTRONICS HCMC CE COMPLEX', extra: [{ label: 'Lĩnh vực', value: 'Công nghệ điện tử' }] },
  { companyName: 'CÔNG TY CỔ PHẦN FPT (FPT CORPORATION)', extra: [{ label: 'Lĩnh vực', value: 'Chuyển đổi số & AI' }] },
  { companyName: 'TỔNG CÔNG TY TÂN CẢNG SÀI GÒN', extra: [{ label: 'Lĩnh vực', value: 'Logistics & Cảng biển' }] },
  { companyName: 'CÔNG TY TNHH JABIL VIỆT NAM', extra: [{ label: 'KCN', value: 'Khu Công Nghệ Cao' }] }
];

const els = {
  connectionBadge: document.getElementById('connectionBadge'),
  connectionText: document.getElementById('connectionText'),
  lastSyncText: document.getElementById('lastSyncText'),
  eventSubtitleDisplay: document.getElementById('eventSubtitleDisplay'),
  eventTitleDisplay: document.getElementById('eventTitleDisplay'),

  // Progress bar elements
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
  soundLabel: document.getElementById('soundLabel'),

  // Settings Modal
  settingsModal: document.getElementById('settingsModal'),
  inputScriptUrl: document.getElementById('inputScriptUrl'),
  inputColumnIndex: document.getElementById('inputColumnIndex'),
  inputPollInterval: document.getElementById('inputPollInterval'),
  inputTargetCount: document.getElementById('inputTargetCount'),
  inputEventSubtitle: document.getElementById('inputEventSubtitle'),
  inputEventTitle: document.getElementById('inputEventTitle'),
  inputFormUrl: document.getElementById('inputFormUrl'),
  chkPopExistingOnStart: document.getElementById('chkPopExistingOnStart'),
  chkShowExtraColumns: document.getElementById('chkShowExtraColumns'),

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

  const formUrl = (config.formUrl || DEFAULT_FORM_URL).trim();
  els.qrCodeImg.src = `https://api.qrserver.com/v1/create-qr-code/?size=280x280&margin=8&data=${encodeURIComponent(formUrl)}`;
  els.qrFormLink.href = formUrl;
  els.qrFormLink.textContent = formUrl.replace(/^https?:\/\//, '') + ' ↗';

  if (config.showQR) {
    els.mainStage.classList.remove('hide-qr');
  } else {
    els.mainStage.classList.add('hide-qr');
  }

  els.soundIcon.textContent = config.soundEnabled ? '🔊' : '🔇';
  els.soundLabel.textContent = config.soundEnabled ? 'Âm thanh: BẬT' : 'Âm thanh: TẮT';

  if (!config.scriptUrl) {
    setConnectionStatus('idle', 'CHƯA CẤU HÌNH SCRIPT');
  }
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
   ĐỒNG BỘ GOOGLE APPS SCRIPT (JSONP + FETCH)
   ========================================================= */
function startPolling() {
  if (pollTimer) clearInterval(pollTimer);
  fetchSheetData();
  const intervalMs = Math.max(1500, (Number(config.pollIntervalSec) || 2) * 1000);
  pollTimer = setInterval(fetchSheetData, intervalMs);
}

async function fetchSheetData() {
  const url = (config.scriptUrl || '').trim();
  if (!url || isFetching) return;
  isFetching = true;

  try {
    const payload = await requestScriptData(url, config.columnIndex);
    if (!payload || !payload.ok) {
      throw new Error((payload && payload.error) || 'Dữ liệu không hợp lệ');
    }

    const nowStr = new Date().toLocaleTimeString('vi-VN');
    els.lastSyncText.textContent = `Cập nhật: ${nowStr}`;
    setConnectionStatus('connected', 'TRỰC TIẾP (LIVE)');

    processIncomingRows(payload.items || []);
  } catch (err) {
    console.warn('Lỗi kết nối Google Apps Script:', err);
    setConnectionStatus('error', 'LỖI KẾT NỐI SCRIPT');
  } finally {
    isFetching = false;
  }
}

function requestScriptData(baseUrl, colIndex) {
  return new Promise((resolve, reject) => {
    const cbName = 'hepzaJsonpCb_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
    const script = document.createElement('script');
    let isDone = false;

    const timer = setTimeout(() => {
      if (isDone) return;
      isDone = true;
      cleanup();
      reject(new Error('Timeout'));
    }, 10000);

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
      fetch(`${baseUrl}${sep}col=${encodeURIComponent(colIndex)}&_t=${Date.now()}`)
        .then((r) => r.json())
        .then(resolve)
        .catch(reject);
    };

    const sep = baseUrl.includes('?') ? '&' : '?';
    script.src = `${baseUrl}${sep}col=${encodeURIComponent(colIndex)}&callback=${cbName}&_t=${Date.now()}`;
    document.body.appendChild(script);
  });
}

/* =========================================================
   XỬ LÝ DỮ LIỆU MỚI: POP ĐỒNG THỜI NHIỀU DOANH NGHIỆP & TIẾN ĐỘ
   ========================================================= */
function processIncomingRows(items) {
  const validItems = items
    .filter((item) => item.companyName && item.companyName.trim() !== '')
    .map((item, idx) => ({
      ...item,
      orderNumber: idx + 1
    }));

  // Nếu có dữ liệu thật từ Sheet, bỏ dữ liệu Demo
  const hadOnlyDemo = companies.length > 0 && companies.every((c) => c.isDemo);
  if (validItems.length > 0 && hadOnlyDemo) {
    companies = [];
    hasCelebratedMilestone = false;
  }

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

  // Lần tải đầu tiên khi mở trang
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

  // Các lần cập nhật tiếp theo: bao nhiêu người submit cùng lúc thì pop đồng thời bấy nhiêu!
  if (newEntries.length > 0) {
    triggerSimultaneousPop(newEntries, prevCount);
  } else if (validItems.length !== prevCount) {
    renderBoard();
    updateProgressBar(companies.length, false);
  }
}

/**
 * Kích hoạt hiệu ứng Pop sáng đồng thời cho 1 hoặc nhiều doanh nghiệp mới
 * kèm pháo hoa nhẹ tại từng thẻ và kiểm tra cột mốc 100
 */
function triggerSimultaneousPop(newItems, previousTotal) {
  const newIds = newItems.map((item) => item.id);
  newIds.forEach((id) => activePoppingIds.add(id));

  // Render bảng với các thẻ mới ở trạng thái .is-popping
  renderBoard();

  // Cuộn mượt lên đầu bảng để thấy các doanh nghiệp mới vừa xuất hiện
  els.boardSection.scrollTo({ top: 0, behavior: 'smooth' });

  // Cập nhật thanh Progress Bar
  const currentTotal = companies.length;
  updateProgressBar(currentTotal, true);

  // Phát âm thanh hiện đại
  const target = Number(config.targetCount) || 100;
  const crossedMilestone = previousTotal < target && currentTotal >= target && !hasCelebratedMilestone;

  if (crossedMilestone) {
    hasCelebratedMilestone = true;
    triggerMilestone100Celebration(target);
  } else if (config.soundEnabled) {
    playPopChime(Math.min(newItems.length, 5));
  }

  // Bắn pháo hoa nhẹ (Sparkle Burst) đồng thời ngay tại vị trí của từng thẻ mới vừa pop lên!
  requestAnimationFrame(() => {
    newIds.forEach((id, idx) => {
      const cardEl = els.companyGrid.querySelector(`[data-id="${CSS.escape(id)}"]`);
      if (cardEl) {
        const rect = cardEl.getBoundingClientRect();
        // Nếu có nhiều thẻ cùng lúc, lệch nhau 60ms tạo cảm giác nở rộ tự nhiên
        setTimeout(() => {
          spawnLightCardFirework(rect.left + rect.width * 0.5, rect.top + rect.height * 0.4);
        }, idx * 65);
      }
    });
  });

  // Sau 4.5 giây, gỡ class .is-popping để thẻ trở về trạng thái bình thường
  setTimeout(() => {
    newIds.forEach((id) => activePoppingIds.delete(id));
    newIds.forEach((id) => {
      const cardEl = els.companyGrid.querySelector(`[data-id="${CSS.escape(id)}"]`);
      if (cardEl) cardEl.classList.remove('is-popping');
    });
  }, 4500);
}

/* =========================================================
   THANH PROGRESS BAR & CỘT MỐC 100 DOANH NGHIỆP
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

  // Thanh fill tối đa 100% chiều rộng nhưng số liệu và % vẫn tiếp tục tăng tiến (101%, 102%...)
  const barWidth = Math.min(100, (count / target) * 100);
  els.progressBarFill.style.width = `${barWidth}%`;

  if (count >= target) {
    els.milestoneStatusPill.textContent = `🎉 Đã đạt & vượt mốc ${target} Doanh nghiệp! (Đang tiếp tục tăng tiến)`;
    els.milestoneStatusPill.classList.add('reached');
  } else {
    const remaining = target - count;
    els.milestoneStatusPill.textContent = `Mục tiêu: ${target} Doanh nghiệp (Còn ${remaining} đơn vị)`;
    els.milestoneStatusPill.classList.remove('reached');
  }
}

function triggerMilestone100Celebration(targetCount = 100) {
  els.milestoneTitle.textContent = `CHÚC MỪNG ĐÃ ĐẠT ${targetCount} DOANH NGHIỆP THAM DỰ!`;
  els.milestoneBanner.classList.add('show');

  if (config.soundEnabled) {
    playMilestoneGrandFanfare();
  }

  // Nổ pháo hoa lớn toàn màn hình liên hoàn trong 5.5 giây
  spawnGrandFullscreenFireworks();

  setTimeout(() => {
    els.milestoneBanner.classList.remove('show');
  }, 6500);
}

/* =========================================================
   HIỂN THỊ LƯỚI DOANH NGHIỆP
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

  // Hiển thị doanh nghiệp mới nhất ở đầu danh sách
  const reversed = [...companies].reverse();

  els.companyGrid.innerHTML = reversed
    .map((item, idx) => {
      const orderNum = item.orderNumber || companies.length - idx;
      const orderStr = String(orderNum).padStart(2, '0');
      const isPopping = activePoppingIds.has(item.id);
      const extraText =
        config.showExtraColumns && Array.isArray(item.extra) && item.extra.length > 0
          ? item.extra.map((e) => `${escapeHtml(e.label)}: ${escapeHtml(e.value)}`).join(' • ')
          : '';

      return `
        <article class="enterprise-card ${isPopping ? 'is-popping' : ''}" data-id="${escapeHtml(item.id)}" title="Nhấn để hiệu ứng sáng lại doanh nghiệp này">
          <div class="card-index">#${orderStr}</div>
          <div class="card-body">
            <h3 class="card-title">${escapeHtml(item.companyName)}</h3>
            ${extraText ? `<div class="card-sub">${extraText}</div>` : ''}
            <div class="card-meta-row">
              ${isPopping ? '<span class="badge-just-now">MỚI CẬP NHẬT</span>' : ''}
              <span>${escapeHtml(item.timestamp || 'Vừa đăng ký')}</span>
            </div>
          </div>
        </article>
      `;
    })
    .join('');

  // Click vào thẻ bất kỳ để kích hoạt lại hiệu ứng pop sáng + pháo hoa nhẹ tại thẻ đó
  els.companyGrid.querySelectorAll('.enterprise-card').forEach((card) => {
    card.addEventListener('click', () => {
      card.classList.remove('is-popping');
      void card.offsetWidth; // Reflow
      card.classList.add('is-popping');
      const rect = card.getBoundingClientRect();
      spawnLightCardFirework(rect.left + rect.width * 0.5, rect.top + rect.height * 0.4);
      if (config.soundEnabled) playPopChime(1);
      setTimeout(() => card.classList.remove('is-popping'), 3500);
    });
  });
}

/* =========================================================
   CHẾ ĐỘ MÔ PHỎNG (TEST 1 NGƯỜI, NHIỀU NGƯỜI CÙNG LÚC, MỐC 100)
   ========================================================= */
function addDemoEnterprises(batchSize = 1) {
  const prevCount = companies.length;
  const newItems = [];
  const nowStr = new Date().toLocaleTimeString('vi-VN');

  for (let i = 0; i < batchSize; i++) {
    const sample = DEMO_COMPANIES[demoCounter % DEMO_COMPANIES.length];
    demoCounter++;
    const orderNum = companies.length + 1;
    const item = {
      id: `demo_${Date.now()}_${demoCounter}_${Math.random().toString(36).slice(2, 6)}`,
      orderNumber: orderNum,
      companyName: sample.companyName,
      extra: sample.extra,
      timestamp: `${nowStr}`,
      isDemo: true
    };
    companies.push(item);
    newItems.push(item);
  }

  triggerSimultaneousPop(newItems, prevCount);
}

/* =========================================================
   HỆ THỐNG PHÁO HOA (PHÁO HOA NHẸ TẠI THẺ & PHÁO HOA LỚN MỐC 100)
   ========================================================= */
let fxParticles = [];
let fxRockets = [];
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

/**
 * Pháo hoa nhẹ & hạt sáng tinh tế ngay tại thẻ doanh nghiệp vừa pop lên
 */
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

/**
 * Pháo hoa lớn toàn màn hình khi đạt đủ cột mốc 100 Doanh nghiệp
 */
function spawnGrandFullscreenFireworks() {
  const canvas = document.getElementById('fxCanvas');
  const palette = ['#e60028', '#2563eb', '#f59e0b', '#10b981', '#ec4899', '#06b6d4', '#fbbf24'];

  // Bắn 14 quả pháo hoa lớn liên tiếp khắp màn hình trong 4.5 giây
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

/**
 * Âm thanh Chime trong trẻo, hiện đại khi doanh nghiệp pop lên
 */
function playPopChime(simultaneousCount = 1) {
  try {
    const ctx = ensureAudioCtx();
    if (!ctx) return;

    const now = ctx.currentTime;
    // Hợp âm trong trẻo, thanh lịch
    const baseNotes = [587.33, 880.0, 1174.66]; // D5 - A5 - D6
    const extraNotes = simultaneousCount > 1 ? [1479.98] : []; // F#6 khi có nhiều doanh nghiệp cùng lúc
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

/**
 * Âm thanh chúc mừng cột mốc 100 Doanh nghiệp
 */
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
   SỰ KIỆN NÚT BẤM & PHÍM TẮT
   ========================================================= */
function bindEvents() {
  // Mô phỏng 1 doanh nghiệp
  document.getElementById('btnTestSingle').addEventListener('click', () => addDemoEnterprises(1));
  document.getElementById('btnTestFromEmpty').addEventListener('click', () => addDemoEnterprises(1));

  // Mô phỏng 4 doanh nghiệp submit đồng thời
  document.getElementById('btnTestBatch').addEventListener('click', () => addDemoEnterprises(4));
  document.getElementById('btnTestMultiFromEmpty').addEventListener('click', () => addDemoEnterprises(4));

  // Mô phỏng nổ pháo hoa cột mốc 100
  document.getElementById('btnTestMilestone100').addEventListener('click', () => {
    triggerMilestone100Celebration(config.targetCount || 100);
  });

  // Đồng bộ ngay
  document.getElementById('btnSyncNow').addEventListener('click', () => {
    if (!config.scriptUrl) {
      openSettingsModal();
    } else {
      fetchSheetData();
    }
  });

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
  document.getElementById('btnOpenSettingsFromEmpty').addEventListener('click', openSettingsModal);
  document.getElementById('btnCloseSettings').addEventListener('click', closeSettingsModal);
  document.getElementById('btnCancelSettings').addEventListener('click', closeSettingsModal);

  // Lưu Cài đặt
  document.getElementById('btnSaveSettings').addEventListener('click', () => {
    config.scriptUrl = els.inputScriptUrl.value.trim();
    config.columnIndex = parseInt(els.inputColumnIndex.value, 10) || 2;
    config.pollIntervalSec = parseInt(els.inputPollInterval.value, 10) || 2;
    config.targetCount = Math.max(1, parseInt(els.inputTargetCount.value, 10) || 100);
    config.eventSubtitle = els.inputEventSubtitle.value.trim() || defaultConfig.eventSubtitle;
    config.eventTitle = els.inputEventTitle.value.trim() || defaultConfig.eventTitle;
    config.formUrl = els.inputFormUrl.value.trim() || DEFAULT_FORM_URL;
    config.popExistingOnStart = els.chkPopExistingOnStart.checked;
    config.showExtraColumns = els.chkShowExtraColumns.checked;

    saveConfig();
    applyConfigToUI();
    updateProgressBar(companies.length, false);
    closeSettingsModal();
    startPolling();
  });

  // Xóa dữ liệu chạy thử (Demo)
  document.getElementById('btnClearDemoData').addEventListener('click', () => {
    companies = companies.filter((c) => !c.isDemo);
    hasCelebratedMilestone = companies.length >= (Number(config.targetCount) || 100);
    renderBoard();
    updateProgressBar(companies.length, false);
    closeSettingsModal();
  });

  // Đặt lại lịch sử đã hiện
  document.getElementById('btnResetSeenHistory').addEventListener('click', () => {
    seenIds.clear();
    saveSeenIds();
    hasCelebratedMilestone = false;
    hasCompletedInitialSync = true;
    closeSettingsModal();
    if (config.scriptUrl) {
      fetchSheetData();
    } else if (companies.length > 0) {
      triggerSimultaneousPop(companies, 0);
    }
  });

  // Mở/Đóng Hướng dẫn Script
  document.getElementById('btnGuide').addEventListener('click', openGuideModal);
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

  // Phím tắt nhanh
  window.addEventListener('keydown', (e) => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') {
      return;
    }
    if (e.key === 'Escape') {
      closeSettingsModal();
      closeGuideModal();
    } else if (e.key === ' ' || e.code === 'Space') {
      e.preventDefault();
      addDemoEnterprises(1);
    } else if (e.key === 'm' || e.key === 'M') {
      e.preventDefault();
      addDemoEnterprises(4);
    } else if (e.key === 'f' || e.key === 'F') {
      toggleFullscreen();
    }
  });
}

function openSettingsModal() {
  els.inputScriptUrl.value = config.scriptUrl || '';
  els.inputColumnIndex.value = String(config.columnIndex || 2);
  els.inputPollInterval.value = String(config.pollIntervalSec || 2);
  els.inputTargetCount.value = String(config.targetCount || 100);
  els.inputEventSubtitle.value = config.eventSubtitle;
  els.inputEventTitle.value = config.eventTitle;
  els.inputFormUrl.value = config.formUrl || DEFAULT_FORM_URL;
  els.chkPopExistingOnStart.checked = !!config.popExistingOnStart;
  els.chkShowExtraColumns.checked = config.showExtraColumns !== false;
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
