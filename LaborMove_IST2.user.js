// ==UserScript==
// @name         Labour Move Assistant — IST2
// @namespace    IST2-Flow
// @version      4.3
// @description  FCLM permission sync + Pick Workforce highlight + Excel export (SLAM=Expert, V-Returns)
// @match        https://fclm-portal.amazon.com/utilities/employeesByPermissions*
// @match        https://picking-console.eu.picking.aft.a2z.com/fc/IST2/pick-workforce
// @grant        GM_addStyle
// @grant        GM_setValue
// @grant        GM_getValue
// @updateURL    https://raw.githubusercontent.com/meyhur777/labourmove-ist2/main/LaborMove_IST2.user.js
// @downloadURL  https://raw.githubusercontent.com/meyhur777/labourmove-ist2/main/LaborMove_IST2.user.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

(function () {
  'use strict';

  const IS_FCLM = location.hostname === 'fclm-portal.amazon.com';
  const IS_WORKFORCE = location.hostname === 'picking-console.eu.picking.aft.a2z.com';

  const PERM_KEYS = ['PICK','WRANGLE','PACK','REBIN','NOSLAM','MANUALSLAM','GIFTPACK','STOW','RECEIVE','ICQA','SHIP','VRETURNS'];

  // ─── FCLM SYNC PAGE ─────────────────────────────────────────────────────────
  if (IS_FCLM) {

    const PERMISSIONS = [
      { key: 'PICK',       processId: '1002909' },
      { key: 'WRANGLE',    processId: '1002939' },
      { key: 'PACK',       processId: '3180'    },
      { key: 'REBIN',      processId: '1002937' },
      { key: 'NOSLAM',     processId: '1548196124262', expertOnly: true },
      { key: 'MANUALSLAM', processId: '1002898', expertOnly: true },
      { key: 'GIFTPACK',   processId: '1002890' },
      { key: 'STOW',       processId: '1002940' },
      { key: 'RECEIVE',    processId: '1002929' },
      { key: 'ICQA',       processId: '1002882' },
      { key: 'SHIP',       processId: '1002935' },
      { key: 'VRETURNS',   processId: '1002951' },
    ];

    const BASE_URL = 'https://fclm-portal.amazon.com/utilities/employeesByPermissions';

    GM_addStyle(`
      #lm-sync-panel {
        position: fixed; top: 16px; right: 16px; z-index: 99999;
        background: #1a2233; color: #e8edf5; border-radius: 10px;
        padding: 16px 20px; font-family: Arial, sans-serif; font-size: 13px;
        width: 290px; box-shadow: 0 4px 20px rgba(0,0,0,0.4);
      }
      #lm-sync-panel h3 { margin: 0 0 12px 0; font-size: 14px; color: #ff9900; letter-spacing: 0.5px; }
      #lm-sync-btn {
        width: 100%; padding: 9px; background: #ff9900; color: #1a2233;
        border: none; border-radius: 6px; font-weight: bold; font-size: 13px;
        cursor: pointer; margin-top: 8px;
      }
      #lm-sync-btn:disabled { background: #555; color: #999; cursor: not-allowed; }
      #lm-export-btn {
        width: 100%; padding: 9px; background: #00c853; color: #fff;
        border: none; border-radius: 6px; font-weight: bold; font-size: 13px;
        cursor: pointer; margin-top: 8px; display: none;
      }
      #lm-export-btn:hover { background: #00e676; }
      #lm-sync-status { margin-top: 10px; font-size: 12px; color: #a0b0c8; min-height: 18px; }
      #lm-sync-progress { margin-top: 8px; height: 6px; background: #2a3a4a; border-radius: 3px; overflow: hidden; display: none; }
      #lm-sync-progress-bar { height: 100%; background: #ff9900; width: 0%; transition: width 0.3s; }
      #lm-sync-last { margin-top: 8px; font-size: 11px; color: #607080; }
    `);

    const panel = document.createElement('div');
    panel.id = 'lm-sync-panel';
    panel.innerHTML = `
      <h3>Labour Move — Sync</h3>
      <div id="lm-sync-last">Last sync: —</div>
      <button id="lm-sync-btn">Update Permissions</button>
      <button id="lm-export-btn">📥 Download Excel</button>
      <div id="lm-sync-progress"><div id="lm-sync-progress-bar"></div></div>
      <div id="lm-sync-status"></div>
    `;
    document.body.appendChild(panel);

    const lastSync = GM_getValue('lm_sync_time', null);
    if (lastSync) {
      const mins = Math.round((Date.now() - lastSync) / 60000);
      document.getElementById('lm-sync-last').textContent =
        mins < 60 ? `Last sync: ${mins} min ago` : `Last sync: ${Math.round(mins/60)} hr ago`;
      if (GM_getValue('lm_perm_data', null))
        document.getElementById('lm-export-btn').style.display = 'block';
    }

    function setStatus(msg, color) {
      const el = document.getElementById('lm-sync-status');
      el.textContent = msg;
      el.style.color = color || '#a0b0c8';
    }

    function setProgress(pct) {
      document.getElementById('lm-sync-progress').style.display = 'block';
      document.getElementById('lm-sync-progress-bar').style.width = pct + '%';
    }

    function parseEmployees(html) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');
      const employees = [];
      const table = doc.querySelectorAll('table')[1];
      if (!table) return employees;
      const headerRow = table.querySelector('tr');
      if (!headerRow) return employees;
      const headers = Array.from(headerRow.querySelectorAll('th')).map(th => th.textContent.trim());
      const nameIdx = headers.findIndex(h => h.toLowerCase().includes('employee name'));
      const mgrIdx  = headers.findIndex(h => h.toLowerCase().includes('manager name'));
      if (nameIdx === -1) return employees;
      Array.from(table.querySelectorAll('tr')).slice(1).forEach(row => {
        const cells = row.querySelectorAll('td');
        const name = cells[nameIdx] ? cells[nameIdx].textContent.trim() : '';
        const manager = (mgrIdx !== -1 && cells[mgrIdx]) ? cells[mgrIdx].textContent.trim() : '';
        if (name) employees.push({ name, manager });
      });
      return employees;
    }

    async function fetchPermission(perm) {
      const level = perm.expertOnly ? 'Expert' : 'Beginner';
      const params = [
        'warehouseId=IST2',
        '_processIds=1',
        `permissionLevel=${level}`,
        'employeeStatusActive=true',
        '_employeeStatusActive=on',
        'employeeStatusTerminated=true',
        '_employeeStatusTerminated=on',
        'employeeStatusLeaveOfAbsence=true',
        '_employeeStatusLeaveOfAbsence=on',
      ].join('&');
      const resp = await fetch(`${BASE_URL}?${params}&processIds=${perm.processId}`, { credentials: 'include' });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      return parseEmployees(await resp.text());
    }

    // ─── EXCEL EXPORT ────────────────────────────────────────────────────────
    function buildAndDownloadExcel(permMap) {
      const wb = XLSX.utils.book_new();

      const C_DARK     = 'FF232F3E';
      const C_WHITE    = 'FFFFFFFF';
      const C_ORANGE   = 'FFFF9900';
      const C_GRN_BG   = 'FFE8F5E9';
      const C_GRN_FG   = 'FF1B5E20';
      const C_ALT      = 'FFF5F5F5';
      const C_SLATE    = 'FF37474F';

      const hdrStyle = {
        font: { bold: true, color: { rgb: C_WHITE }, sz: 11 },
        fill: { fgColor: { rgb: C_DARK } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: { bottom: { style: 'thin', color: { rgb: C_ORANGE } } }
      };
      const permHdrStyle = {
        font: { bold: true, color: { rgb: C_WHITE }, sz: 10 },
        fill: { fgColor: { rgb: C_ORANGE } },
        alignment: { horizontal: 'center', vertical: 'center' }
      };
      const checkStyle = {
        font: { bold: true, color: { rgb: C_GRN_FG }, sz: 11 },
        fill: { fgColor: { rgb: C_GRN_BG } },
        alignment: { horizontal: 'center' }
      };
      const altStyle = { fill: { fgColor: { rgb: C_ALT } } };

      const employees = Object.entries(permMap).map(([name, data]) => ({
        name, manager: data.manager || '—', perms: new Set(data.perms)
      })).sort((a, b) => a.name.localeCompare(b.name));

      // ── SHEET 1: Employees ───────────────────────────────────────────────
      const ws1Data = [['Employee Name', 'Manager', ...PERM_KEYS]];
      employees.forEach(emp =>
        ws1Data.push([emp.name, emp.manager, ...PERM_KEYS.map(k => emp.perms.has(k) ? '✓' : '')])
      );
      const ws1 = XLSX.utils.aoa_to_sheet(ws1Data);
      ws1['!cols'] = [{ wch: 28 }, { wch: 26 }, ...PERM_KEYS.map(() => ({ wch: 11 }))];
      ws1['!freeze'] = { xSplit: 0, ySplit: 1 };
      const r1 = XLSX.utils.decode_range(ws1['!ref']);
      for (let C = r1.s.c; C <= r1.e.c; C++) {
        const cell = ws1[XLSX.utils.encode_cell({ r: 0, c: C })];
        if (cell) cell.s = C < 2 ? hdrStyle : permHdrStyle;
      }
      for (let R = 1; R <= r1.e.r; R++) {
        const isAlt = R % 2 === 0;
        for (let C = r1.s.c; C <= r1.e.c; C++) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          if (!ws1[addr]) ws1[addr] = { t: 's', v: '' };
          ws1[addr].s = C >= 2 && ws1[addr].v === '✓' ? checkStyle : (isAlt ? altStyle : {});
        }
      }
      XLSX.utils.book_append_sheet(wb, ws1, 'Employees');

      // ── SHEET 2: Permission Summary ──────────────────────────────────────
      const total = employees.length;
      const ws2Data = [['Permission', 'Employee Count', 'Rate (%)']];
      PERM_KEYS.forEach(k => {
        const count = employees.filter(e => e.perms.has(k)).length;
        ws2Data.push([k, count, Math.round((count / total) * 100)]);
      });
      const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
      ws2['!cols'] = [{ wch: 16 }, { wch: 18 }, { wch: 12 }];
      const r2 = XLSX.utils.decode_range(ws2['!ref']);
      for (let C = 0; C <= 2; C++) {
        const cell = ws2[XLSX.utils.encode_cell({ r: 0, c: C })];
        if (cell) cell.s = { ...hdrStyle, fill: { fgColor: { rgb: C_SLATE } } };
      }
      for (let R = 1; R <= r2.e.r; R++) {
        const isAlt = R % 2 === 0;
        for (let C = 0; C <= 2; C++) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          if (!ws2[addr]) ws2[addr] = { t: 's', v: '' };
          ws2[addr].s = {
            alignment: { horizontal: C === 0 ? 'left' : 'center' },
            fill: isAlt ? { fgColor: { rgb: C_ALT } } : {},
            font: C === 1 ? { bold: true, color: { rgb: C_GRN_FG } } : {}
          };
        }
      }
      XLSX.utils.book_append_sheet(wb, ws2, 'Permission Summary');

      // ── SHEET 3: By Manager ──────────────────────────────────────────────
      const mgrMap = {};
      employees.forEach(emp => {
        const mgr = emp.manager || '—';
        if (!mgrMap[mgr]) mgrMap[mgr] = [];
        mgrMap[mgr].push(emp);
      });
      const ws3Data = [['Manager', 'Employee Name', ...PERM_KEYS]];
      Object.keys(mgrMap).sort().forEach(mgr => {
        mgrMap[mgr].sort((a, b) => a.name.localeCompare(b.name)).forEach((emp, i) => {
          ws3Data.push([i === 0 ? mgr : '', emp.name, ...PERM_KEYS.map(k => emp.perms.has(k) ? '✓' : '')]);
        });
        ws3Data.push(Array(2 + PERM_KEYS.length).fill(''));
      });
      const ws3 = XLSX.utils.aoa_to_sheet(ws3Data);
      ws3['!cols'] = [{ wch: 26 }, { wch: 26 }, ...PERM_KEYS.map(() => ({ wch: 11 }))];
      ws3['!freeze'] = { xSplit: 0, ySplit: 1 };
      const r3 = XLSX.utils.decode_range(ws3['!ref']);
      for (let C = r3.s.c; C <= r3.e.c; C++) {
        const cell = ws3[XLSX.utils.encode_cell({ r: 0, c: C })];
        if (cell) cell.s = C < 2 ? hdrStyle : permHdrStyle;
      }
      for (let R = 1; R <= r3.e.r; R++) {
        for (let C = r3.s.c; C <= r3.e.c; C++) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          if (!ws3[addr]) ws3[addr] = { t: 's', v: '' };
          if (C === 0 && ws3[addr].v) {
            ws3[addr].s = { font: { bold: true, color: { rgb: C_WHITE } }, fill: { fgColor: { rgb: C_SLATE } }, alignment: { vertical: 'center' } };
          } else if (C >= 2) {
            ws3[addr].s = ws3[addr].v === '✓' ? checkStyle : {};
          }
        }
      }
      XLSX.utils.book_append_sheet(wb, ws3, 'By Manager');

      const d = new Date();
      const dateStr = `${d.getFullYear()}_${String(d.getMonth()+1).padStart(2,'0')}_${String(d.getDate()).padStart(2,'0')}`;
      XLSX.writeFile(wb, `IST2_Permissions_${dateStr}.xlsx`);
    }

    // ─── SYNC BUTTON ─────────────────────────────────────────────────────────
    document.getElementById('lm-sync-btn').addEventListener('click', async function () {
      const btn = this;
      btn.disabled = true;
      document.getElementById('lm-export-btn').style.display = 'none';
      setProgress(0);
      setStatus('Starting...', '#a0b0c8');

      try {
        const permMap = {};
        for (let i = 0; i < PERMISSIONS.length; i++) {
          const perm = PERMISSIONS[i];
          setStatus(`Fetching: ${perm.key} (${i + 1}/${PERMISSIONS.length})`, '#ff9900');
          setProgress(Math.round((i / PERMISSIONS.length) * 90));
          try {
            const employees = await fetchPermission(perm);
            employees.forEach(({ name, manager }) => {
              if (!permMap[name]) permMap[name] = { manager, perms: [] };
              permMap[name].perms.push(perm.key);
              if (manager && !permMap[name].manager) permMap[name].manager = manager;
            });
          } catch (e) {
            console.warn(`[LM Sync] ${perm.key} error:`, e);
          }
          await new Promise(r => setTimeout(r, 300));
        }

        setProgress(100);
        const count = Object.keys(permMap).length;
        const permByName = {};
        Object.entries(permMap).forEach(([name, data]) => { permByName[name] = data.perms; });
        GM_setValue('lm_perm_data', JSON.stringify(permByName));
        GM_setValue('lm_perm_full', JSON.stringify(permMap));
        GM_setValue('lm_sync_time', Date.now());
        setStatus(`✅ ${count} employees synced`, '#4caf50');
        document.getElementById('lm-sync-last').textContent = 'Last sync: Just now';
        document.getElementById('lm-export-btn').style.display = 'block';

      } catch (e) {
        setStatus(`❌ Error: ${e.message}`, '#f44336');
      }
      btn.disabled = false;
    });

    document.getElementById('lm-export-btn').addEventListener('click', function () {
      const raw = GM_getValue('lm_perm_full', null);
      if (!raw) { setStatus('Run sync first', '#ff4757'); return; }
      try {
        setStatus('Building Excel...', '#ff9900');
        buildAndDownloadExcel(JSON.parse(raw));
        setStatus('✅ Excel downloaded', '#4caf50');
      } catch (e) {
        setStatus(`❌ ${e.message}`, '#f44336');
      }
    });

    return;
  }

  // ─── PICK WORKFORCE PAGE ─────────────────────────────────────────────────────
  if (!IS_WORKFORCE) return;

  let PERM_BY_NAME = {};
  const permDataRaw = GM_getValue('lm_perm_data', null);
  const syncTime = GM_getValue('lm_sync_time', null);
  const syncAge = syncTime ? Math.round((Date.now() - syncTime) / 60000) : null;

  if (permDataRaw) {
    try {
      const parsed = JSON.parse(permDataRaw);
      Object.entries(parsed).forEach(([name, perms]) => {
        PERM_BY_NAME[name.toLowerCase().trim()] = new Set(perms);
      });
    } catch (e) {
      console.warn('[LM v4] permData parse error:', e);
    }
  }

  function getPermsForRow(row) {
    const name = getNameFromRow(row);
    if (!name) return new Set();
    return PERM_BY_NAME[name.toLowerCase().trim()] || new Set();
  }

  GM_addStyle(`
    #lm-btn {
      position: fixed; top: 80px; right: 16px; z-index: 9999;
      background: #0a0e1a; color: #00d4ff; border: 2px solid #00d4ff;
      border-radius: 8px; padding: 10px 16px; font-family: Arial, sans-serif;
      font-size: 13px; font-weight: bold; cursor: pointer; letter-spacing: 1px;
      box-shadow: 0 4px 20px rgba(0,212,255,0.3); transition: all 0.2s;
    }
    #lm-btn:hover { background: #00d4ff; color: #0a0e1a; }
    #lm-panel {
      position: fixed; top: 130px; right: 16px; z-index: 9998;
      background: #111827; border: 1px solid #1e2d45; border-radius: 12px;
      width: 320px; max-height: 80vh; overflow-y: auto; display: none;
      font-family: Arial, sans-serif; box-shadow: 0 8px 40px rgba(0,0,0,0.6);
    }
    #lm-panel.open { display: block; }
    .lm-header { padding: 14px 16px; border-bottom: 1px solid #1e2d45; display: flex; align-items: center; justify-content: space-between; }
    .lm-title { font-size: 16px; font-weight: bold; color: #00d4ff; letter-spacing: 1px; text-transform: uppercase; }
    .lm-close { cursor: pointer; color: #4a6080; font-size: 18px; }
    .lm-close:hover { color: #ff4757; }
    .lm-section { padding: 10px 16px; border-bottom: 1px solid #1e2d45; }
    .lm-section-title { font-size: 10px; color: #4a6080; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px; }
    .lm-perm-btns { display: flex; flex-wrap: wrap; gap: 6px; }
    .lm-perm-btn {
      padding: 5px 10px; border-radius: 6px; border: 1px solid #1e2d45;
      background: #1a2236; color: #4a6080; font-size: 11px; cursor: pointer;
      font-weight: bold; transition: all 0.15s;
    }
    .lm-perm-btn:hover { border-color: #00d4ff; color: #00d4ff; }
    .lm-perm-btn.active { background: #00d4ff; color: #0a0e1a; border-color: #00d4ff; }
    .lm-count { font-size: 11px; color: #4a6080; margin-top: 6px; }
    .lm-count span { color: #00ff88; font-weight: bold; }
    .lm-selected-list { padding: 10px 16px; max-height: 200px; overflow-y: auto; }
    .lm-selected-item { display: flex; align-items: center; justify-content: space-between; padding: 6px 8px; border-radius: 6px; margin-bottom: 4px; background: #1a2236; font-size: 12px; }
    .lm-sel-name { color: #e0e8f0; }
    .lm-sel-area { color: #4a6080; font-size: 10px; }
    .lm-sel-remove { color: #ff4757; cursor: pointer; font-size: 14px; padding: 0 4px; }
    .lm-sync-bar { padding: 8px 16px; font-size: 11px; border-bottom: 1px solid #1e2d45; }
    .lm-sync-ok { color: #4caf50; }
    .lm-sync-warn { color: #ff9900; }
    .lm-sync-none { color: #ff4757; }
    tr.lm-highlight { background-color: rgba(255,152,0,0.2) !important; }
    tr.lm-highlight td { background-color: rgba(255,152,0,0.2) !important; }
    tr.lm-selected { background-color: rgba(0,255,136,0.25) !important; outline: 2px solid #00ff88; }
    tr.lm-selected td { background-color: rgba(0,255,136,0.25) !important; }
    tr.lm-clickable { cursor: pointer; }
  `);

  let syncBarClass = 'lm-sync-none';
  let syncBarMsg = '⚠️ Not synced — run sync from FCLM page';
  if (syncAge !== null) {
    if (syncAge < 120) {
      syncBarClass = 'lm-sync-ok';
      syncBarMsg = `✅ Last sync: ${syncAge < 60 ? syncAge + ' min' : Math.round(syncAge/60) + ' hr'} ago`;
    } else {
      syncBarClass = 'lm-sync-warn';
      syncBarMsg = `⚠️ Last sync: ${Math.round(syncAge/60)} hr ago — update recommended`;
    }
  }

  const btn = document.createElement('button');
  btn.id = 'lm-btn';
  btn.textContent = '⚡ LABOUR MOVE';
  document.body.appendChild(btn);

  const panel = document.createElement('div');
  panel.id = 'lm-panel';
  panel.innerHTML = `
    <div class="lm-header">
      <div class="lm-title">Labour Move</div>
      <div class="lm-close" id="lm-close">✕</div>
    </div>
    <div class="lm-sync-bar ${syncBarClass}">${syncBarMsg}</div>
    <div class="lm-section">
      <div class="lm-section-title">Highlight by permission</div>
      <div class="lm-perm-btns">
        <button class="lm-perm-btn" data-perm="REBIN">REBIN</button>
        <button class="lm-perm-btn" data-perm="PACK">PACK</button>
        <button class="lm-perm-btn" data-perm="WRANGLE">WRANGLE</button>
        <button class="lm-perm-btn" data-perm="GIFTPACK">GIFTPACK</button>
        <button class="lm-perm-btn" data-perm="NOSLAM">NO SLAM</button>
        <button class="lm-perm-btn" data-perm="MANUALSLAM">MANUALSLAM</button>
        <button class="lm-perm-btn" data-perm="STOW">STOW</button>
        <button class="lm-perm-btn" data-perm="RECEIVE">RECEIVE</button>
        <button class="lm-perm-btn" data-perm="ICQA">ICQA</button>
        <button class="lm-perm-btn" data-perm="SHIP">SHIP</button>
        <button class="lm-perm-btn" data-perm="VRETURNS">V-RETURNS</button>
      </div>
      <div class="lm-count" id="lm-count"></div>
    </div>
    <div class="lm-section">
      <div class="lm-section-title">Selected pickers (<span id="lm-sel-count">0</span>)</div>
      <div class="lm-selected-list" id="lm-selected-list">
        <div style="color:#4a6080;font-size:11px">Select a permission, then click pickers in the table</div>
      </div>
    </div>
  `;
  document.body.appendChild(panel);

  btn.addEventListener('click', () => panel.classList.toggle('open'));
  document.getElementById('lm-close').addEventListener('click', () => panel.classList.remove('open'));

  let activePerm = null;
  let selectedPickers = [];

  panel.querySelectorAll('.lm-perm-btn').forEach(b => {
    b.addEventListener('click', function () {
      const perm = this.dataset.perm;
      if (activePerm === perm) {
        activePerm = null;
        this.classList.remove('active');
        clearHighlights();
        document.getElementById('lm-count').innerHTML = '';
      } else {
        activePerm = perm;
        panel.querySelectorAll('.lm-perm-btn').forEach(x => x.classList.remove('active'));
        this.classList.add('active');
        highlightByPermission(perm);
      }
    });
  });

  function getPickerRows() {
    return Array.from(document.querySelectorAll('tr')).filter(row =>
      row.querySelector('a[href*="/picker/"]')
    );
  }

  function getUserIdFromRow(row) {
    const link = row.querySelector('a[href*="/picker/"]');
    if (!link) return null;
    const parts = link.getAttribute('href').split('/');
    return parts[parts.length - 1].toLowerCase().trim();
  }

  function getNameFromRow(row) {
    const cells = row.querySelectorAll('td');
    if (cells.length > 4) return cells[4].textContent.trim();
    return '';
  }

  function getPickAreaFromRow(row) {
    const cells = row.querySelectorAll('td');
    if (cells.length > 6) return cells[6].textContent.trim();
    return '';
  }

  function clearHighlights() {
    getPickerRows().forEach(row => {
      row.classList.remove('lm-highlight', 'lm-clickable');
      row.removeEventListener('click', onRowClick);
    });
  }

  function highlightByPermission(perm) {
    clearHighlights();
    const rows = getPickerRows();
    let count = 0;
    rows.forEach(row => {
      if (getPermsForRow(row).has(perm)) {
        row.classList.add('lm-highlight', 'lm-clickable');
        row.addEventListener('click', onRowClick);
        count++;
      }
    });
    document.getElementById('lm-count').innerHTML =
      `<span>${count}</span> pickers with ${perm} permission`;
  }

  function onRowClick(e) {
    if (!activePerm) return;
    const row = e.currentTarget;
    const userId = getUserIdFromRow(row);
    if (!userId) return;
    const existing = selectedPickers.findIndex(p => p.userId === userId);
    if (existing > -1) {
      selectedPickers.splice(existing, 1);
      row.classList.remove('lm-selected');
    } else {
      selectedPickers.push({ userId, name: getNameFromRow(row) || userId, pickArea: getPickAreaFromRow(row) });
      row.classList.add('lm-selected');
    }
    updateSelectedList();
  }

  function updateSelectedList() {
    const list = document.getElementById('lm-selected-list');
    document.getElementById('lm-sel-count').textContent = selectedPickers.length;
    if (selectedPickers.length === 0) {
      list.innerHTML = '<div style="color:#4a6080;font-size:11px">Select a permission, then click pickers in the table</div>';
      return;
    }
    list.innerHTML = selectedPickers.map((p, i) => `
      <div class="lm-selected-item">
        <div>
          <div class="lm-sel-name">${p.name}</div>
          <div class="lm-sel-area">${p.userId} · ${p.pickArea}</div>
        </div>
        <div class="lm-sel-remove" data-idx="${i}">✕</div>
      </div>
    `).join('');
    list.querySelectorAll('.lm-sel-remove').forEach(b => {
      b.addEventListener('click', function (e) {
        e.stopPropagation();
        const removed = selectedPickers.splice(parseInt(this.dataset.idx), 1)[0];
        getPickerRows().forEach(row => {
          if (getUserIdFromRow(row) === removed.userId) row.classList.remove('lm-selected');
        });
        updateSelectedList();
      });
    });
  }

  function init() {
    const rows = getPickerRows();
    if (rows.length > 0) {
      console.log(`[LM v4.1] Ready — ${rows.length} pickers, ${Object.keys(PERM_BY_NAME).length} employees loaded`);
    } else {
      setTimeout(init, 1000);
    }
  }

  init();

})();
