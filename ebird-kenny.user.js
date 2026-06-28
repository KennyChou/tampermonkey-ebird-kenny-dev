// ==UserScript==
// @name         ebird-kenny
// @namespace    https://kennychou.github.io/
// @version      1.0.0
// @description  eBird 工具包：清單/行程下載、群組名片夾、GPS 軌跡下載
// @author       Kenny Chou
// @grant        GM_xmlhttpRequest
// @connect      api.ebird.org
// @connect      ebird.org
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @match        https://ebird.org/mychecklists*
// @match        https://ebird.org/*/mychecklists*
// @match        https://ebird.org/mytripreports
// @match        https://ebird.org/*/mytripreports
// @match        https://ebird.org/checklist/*
// @match        https://ebird.org/*/checklist/*
// ==/UserScript==

/**
 * @fileoverview ebird-kenny — eBird 非官方工具腳本
 *
 * 模組架構（依路由啟動）：
 *
 *   /mychecklists   → initChecklist()   批次勾選清單並下載為 Excel
 *   /mytripreports  → initTrip()        行程一鍵整包下載 Excel
 *   /checklist/:id  → initUsersGroup()  群組名片夾（快速分享）
 *                  → initTrack()        GPS 軌跡匯出 KML
 *
 * 資料流：
 *   localStorage ←→ createSetupPanel()  ← API Key、鳥名錄快取
 *   eBird API v2  → gmFetch()           → checklist / hotspot / taxonomy
 *   SheetJS       ← xlsxDownload()      → .xlsx 檔案下載
 *   Blob URL      ← blobDownload()      → .kml 檔案下載
 */

(function () {
  'use strict';

  // ── 繁殖行為代碼對照表 ────────────────────────────────────────────────────
  //
  // 對應 eBird 繁殖代碼至易讀描述，用於 Excel 匯出欄位。
  // 代碼依確認程度由高（NY）到低（FO）排列。
  // 參考：https://help.ebird.org/customer/portal/articles/1006519

  const BREEDING_CODE = {
    NY: 'NY Nest with Young',
    NE: 'NE Nest with Eggs',
    FS: 'CS Carrying Fecal Sac',
    FR: 'FY Feeding Young',
    CF: 'CF Carrying Food',
    FL: 'FL Recently Fledged Young',
    ON: 'ON Occupied Nest',
    UN: 'UN Used Nest',
    DD: 'DD Distraction Display',
    NB: 'NB Nest Building',
    CM: 'CN Carrying Nesting Material',
    BP: 'PE Physiological Evidence',
    DN: 'B Wren/Woodpecker Nest Building',
    AB: 'A Agitated Behavior',
    VS: 'N Visiting Probable Nest Site',
    CC: 'C Courtship, Display or Copulation',
    T7: 'T Territorial Defense',
    PO: 'P Pair in Suitable Habitat',
    SM: 'M Multiple (7+) Singing Birds',
    S7: 'S7 Singing Bird Present 7+ Days',
    S1: 'S Singing Bird',
    OS: 'H In Appropriate Habitat',
    FO: 'F Flyover',
  };

  // eBird 調查方法代碼對照表（protocolId → 人類可讀名稱）
  const PRO_CODE = {
    P20: 'eBird - Casual Observation',
    P21: 'eBird - Stationary Count',
    P22: 'eBird - Traveling Count',
    P23: 'eBird - Exhaustive Area Count',
  };

  // ── 網路請求層 ────────────────────────────────────────────────────────────
  //
  // 使用 GM_xmlhttpRequest 而非 fetch，原因：
  //   eBird API (api.ebird.org) 不開放 CORS，一般 fetch 會被瀏覽器封鎖；
  //   GM API 在 @connect 白名單內可繞過同源限制。

  /**
   * 透過 GM_xmlhttpRequest 發送 GET 請求並解析 JSON 回應。
   *
   * @param {string} url - 請求目標 URL
   * @param {string} [key] - eBird API Token，存入 X-eBirdApiToken header
   * @returns {Promise<any>} 解析後的 JSON 物件
   * @throws {Error} HTTP 4xx/5xx 或 JSON 解析失敗時拋出
   */
  function gmFetch(url, key) {
    return new Promise((resolve, reject) => {
      const headers = {};
      if (key) headers['X-eBirdApiToken'] = key;
      GM_xmlhttpRequest({
        method: 'GET',
        url,
        headers,
        onload(r) {
          if (r.status >= 400) { reject(new Error(`HTTP ${r.status}`)); return; }
          try { resolve(JSON.parse(r.responseText)); }
          catch { reject(new Error('JSON parse error')); }
        },
        onerror() { reject(new Error('Network error')); },
      });
    });
  }

  /**
   * 取得 eBird 鳥種名錄。
   *
   * @param {string} lang - 語系代碼，例如 'zh_TW'、'en'
   * @returns {Promise<Array<{speciesCode: string, comName: string, sciName: string, taxonOrder: number}>>}
   */
  function fetchTaxonomy(lang) {
    return gmFetch(`https://api.ebird.org/v2/ref/taxonomy/ebird?fmt=json&locale=${lang}`);
  }

  /**
   * 取得單筆清單的詳細資料（物種、數量、地點、時間等）。
   *
   * @param {string} subId - 清單 ID（例如 'S12345678'）
   * @param {string} key - eBird API Token
   * @returns {Promise<Object>} eBird checklist view 物件
   */
  function fetchChecklistData(subId, key) {
    return gmFetch(`https://api.ebird.org/v2/product/checklist/view/${subId}`, key);
  }

  /**
   * 取得熱點資訊（名稱、座標）。失敗時靜默回傳 null，讓呼叫端降級處理。
   *
   * @param {string} locId - 地點 ID（例如 'L12345'）
   * @param {string} key - eBird API Token
   * @returns {Promise<Object|null>} 熱點物件，或取得失敗時為 null
   */
  async function fetchHotspot(locId, key) {
    try { return await gmFetch(`https://api.ebird.org/v2/ref/hotspot/info/${locId}`, key); }
    catch { return null; }
  }

  /**
   * 取得行程報告下所有清單的摘要清單。
   * 使用 eBird 內部 API，不需要 API Token。
   *
   * @param {string} tripId - 行程 ID（對應頁面上 heading 的 id 屬性）
   * @returns {Promise<Array<{subId: string, loc: {locName: string, lat: number, lng: number}}>>}
   */
  function fetchTripChecklists(tripId) {
    return gmFetch(`https://ebird.org/tripreport-internal/v1/checklists/${tripId}`);
  }

  // ── localStorage 存取 ─────────────────────────────────────────────────────
  //
  // 三組資料各自獨立存放，方便單獨清除或更新：
  //   ebirdKey   → string（API Token）
  //   sp_info    → JSON 物件，key 為 speciesCode，value 為 {c, s, o}
  //   ebird-groups → JSON 陣列，每筆 {name, ebirdIds}

  const getKey    = () => localStorage.getItem('ebirdKey') || '';
  const setKey    = k  => localStorage.setItem('ebirdKey', k);
  const clearKey  = () => localStorage.removeItem('ebirdKey');
  const getSpInfo = () => JSON.parse(localStorage.getItem('sp_info') || '{}');
  const setSpInfo = d  => localStorage.setItem('sp_info', JSON.stringify(d));

  // ── Excel 輸出 ─────────────────────────────────────────────────────────────

  /**
   * 將單筆觀察紀錄組合成 Excel 一列資料。
   * 欄位對應 eBird 官方「Download My Data」匯出格式，方便與既有分析工具相容。
   *
   * @param {Object} ret - fetchChecklistData 回傳的清單物件
   * @param {Object} obs - 清單內單筆 obs 物件
   * @param {{c: string, s: string, o: number}} sp - 來自 sp_info 的鳥種資訊
   * @param {string} location - 地點名稱
   * @param {number} lat - 緯度
   * @param {number} lng - 經度
   * @param {string} county - 縣市代碼
   * @returns {Object} 一列 Excel 資料，key 為欄位名稱
   */
  function buildObsRow(ret, obs, sp, location, lat, lng, county) {
    const [date, time] = ret.obsDt.split(' ');
    return {
      'Submission ID':          ret.subId,
      'Common Name':            sp.c,
      'Scientific Name':        sp.s,
      'Taxonomic Order':        sp.o,
      Count:                    obs.howManyAtmost,
      'State/Province':         ret.subnational1Code,
      County:                   county || '',
      'Location ID':            ret.locId,
      Location:                 location || '',
      Latitude:                 lat || '',
      Longitude:                lng || '',
      Date:                     date,
      Time:                     time || '',
      Protocol:                 PRO_CODE[ret.protocolId] || ret.protocolId || '',
      'Duration (Min)':         ret.durationHrs ? Math.round(parseFloat(ret.durationHrs) * 60) : '',
      'All Obs Reported':       ret.allObsReported ? 1 : 0,
      'Distance Traveled (km)': ret.effortDistanceKm || '',
      'Area Covered (ha)':      ret.effortAreaHa || '',
      'Number of Observers':    ret.numObservers || '',
      'Breeding Code':          obs.obsAux?.[0] ? (BREEDING_CODE[obs.obsAux[0].auxCode] || '') : '',
      'Observation Details':    obs.comments || '',
      'Checklist Comments':     ret.comments || '',
    };
  }

  /**
   * 將資料列陣列寫成 xlsx 檔案並觸發下載。
   *
   * @param {Object[]} rows - buildObsRow 產生的資料列陣列
   * @param {string} filename - 下載檔名（含副檔名 .xlsx）
   */
  function xlsxDownload(rows, filename) {
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'sheet1');
    XLSX.writeFile(wb, filename);
  }

  // ── DOM 建構輔助 ──────────────────────────────────────────────────────────

  /**
   * 輕量 createElement 包裝函式，類似 React.createElement 的命令式版本。
   * 支援 className、style（cssText 字串）、事件監聽（onXxx）及一般 attribute。
   *
   * @param {string} tag - HTML 標籤名
   * @param {Object} [props] - 屬性物件
   * @param {Array<Node|string|null>} [children] - 子節點陣列，null 會自動跳過
   * @returns {HTMLElement}
   */
  function h(tag, props, children = []) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(props || {})) {
      if (k === 'className') node.className = v;
      else if (k === 'style') node.style.cssText = v;
      else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2), v);
      else node.setAttribute(k, v);
    }
    for (const child of children) {
      if (child == null) continue;
      node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child);
    }
    return node;
  }

  // ── 設定面板（清單與行程頁共用）─────────────────────────────────────────
  //
  // 負責 API Key 輸入／清除、鳥名錄下載兩步初始設定。
  // 採用「每次狀態異動重繪」策略，避免維護散落的 display 旗標。
  // onReady() 由呼叫端傳入，當設定完成（key + sp_info 均存在）後觸發，
  // 讓各模組注入自己的功能 UI。

  /**
   * 建立共用設定面板，並在設定完成時回呼 onReady。
   *
   * @param {Function} [onReady] - API Key 與鳥名錄皆就緒後的回呼
   * @returns {{ panel: HTMLElement, msgDiv: HTMLElement }}
   */
  function createSetupPanel(onReady) {
    const panel = document.createElement('div');
    panel.className = 'Page-section-inner Page-section-inner--md';

    const msgDiv = h('div', { style: 'margin-top:8px;' });
    panel.appendChild(msgDiv);

    function render() {
      // 保留 msgDiv，移除其餘子節點後重繪
      while (panel.firstChild && panel.firstChild !== msgDiv) {
        panel.removeChild(panel.firstChild);
      }

      const key    = getKey();
      const spInfo = getSpInfo();

      // 尚未輸入 API Key → 顯示輸入欄
      if (!key) {
        const keyInput = h('input', { type: 'text', style: 'margin-left:8px;width:260px;' });
        const saveBtn  = h('button', {
          className: 'Button Button--highlight',
          style: 'margin-left:8px;',
          onclick() {
            const val = keyInput.value.trim();
            if (!val) return;
            setKey(val);
            render();
            onReady?.();
          },
        }, ['儲存']);
        panel.insertBefore(
          h('div', { style: 'margin-bottom:8px;' }, [
            h('label', {}, [
              h('span', {}, [
                '輸入 ',
                h('a', { href: 'https://ebird.org/api/keygen', target: '_blank' }, ['eBird API Key']),
              ]),
              keyInput,
              saveBtn,
            ]),
          ]),
          msgDiv
        );
      }

      // 鳥名錄尚未快取 → 顯示下載選項
      if (Object.keys(spInfo).length === 0) {
        const langSel = h('select', { style: 'margin:0 8px;' }, [
          h('option', { value: 'zh_TW' }, ['中文俗名']),
          h('option', { value: 'en' },    ['英文俗名']),
        ]);
        const dlBtn = h('button', {
          className: 'Button Button--highlight',
          async onclick() {
            msgDiv.textContent = '正在下載鳥名錄...';
            try {
              const list = await fetchTaxonomy(langSel.value);
              // 將陣列轉成以 speciesCode 為 key 的查找表，加速後續名稱對照
              const info = {};
              for (const item of list) {
                info[item.speciesCode] = { c: item.comName, s: item.sciName, o: item.taxonOrder };
              }
              setSpInfo(info);
              msgDiv.textContent = '鳥名錄下載成功';
              render();
              onReady?.();
            } catch (e) {
              msgDiv.textContent = '下載失敗：' + e.message;
            }
          },
        }, ['下載鳥名錄']);
        panel.insertBefore(
          h('div', { style: 'margin-bottom:8px;' }, [
            h('label', {}, [h('span', {}, ['鳥名顯示方式']), langSel, dlBtn]),
          ]),
          msgDiv
        );
      }

      // 已設定 API Key → 提供清除按鈕
      if (key) {
        panel.insertBefore(
          h('button', {
            className: 'Button Button--secondary',
            style: 'font-size:0.8em;margin-bottom:8px;',
            onclick() { clearKey(); render(); onReady?.(); },
          }, ['清除 API Key']),
          msgDiv
        );
      }
    }

    render();
    return { panel, msgDiv };
  }

  // ── 路由 ──────────────────────────────────────────────────────────────────
  //
  // 依 pathname 決定啟動哪些模組。
  // /checklist/* 同時掛載群組名片夾與 GPS 軌跡兩個獨立模組。

  const path = location.pathname;

  if (/\/mychecklists/.test(path))  initChecklist();
  else if (/\/mytripreports/.test(path)) initTrip();
  else if (/\/checklist\//.test(path)) { initUsersGroup(); initTrack(); }

  // ── 清單下載模組 (/mychecklists) ─────────────────────────────────────────

  /**
   * 在「我的清單」頁面注入勾選框與批次下載功能。
   *
   * DOM 依賴：
   *   #myChecklistsForm                     → 面板插入錨點
   *   #place-species-observed-results       → 清單容器
   *   .ResultsStats--manageMyChecklists     → 每筆清單 item（id="result-{subId}-*"）
   *   .Color-text-neutral-4.u-text-3        → 勾選框插入位置
   */
  function initChecklist() {
    const anchor = document.getElementById('myChecklistsForm');
    if (!anchor) return;

    const root = document.createElement('div');
    anchor.after(root);

    const { panel, msgDiv } = createSetupPanel(injectCheckboxes);
    root.appendChild(panel);

    const dlBtn = h('button', {
      className: 'Button Button--highlight',
      async onclick() { await downloadSelected(); },
    }, ['下載選取的清單']);
    const dlBar = h('div', { style: 'margin:8px 0;display:none;' }, [dlBtn]);
    root.appendChild(dlBar);

    injectCheckboxes();

    /** 為每筆清單注入勾選框，設定完成前不執行。 */
    function injectCheckboxes() {
      document.querySelectorAll('.kenny-checkbox').forEach(n => n.remove());
      if (!getKey() || !Object.keys(getSpInfo()).length) { dlBar.style.display = 'none'; return; }

      try {
        const items = document
          .getElementById('place-species-observed-results')
          ?.getElementsByClassName('ResultsStats ResultsStats--manageMyChecklists');
        if (!items) return;
        for (const item of items) {
          // id 格式為 "result-{subId}-{idx}"，取第二段即為 subId
          const id = item.id.split('-')[1];
          const cb = document.createElement('input');
          cb.type      = 'checkbox';
          cb.value     = id;
          cb.className = 'kenny-checkbox';
          cb.style.marginRight = '6px';
          cb.addEventListener('change', updateBar);
          item.getElementsByClassName('Color-text-neutral-4 u-text-3')[0]?.prepend(cb);
        }
      } catch (e) {
        msgDiv.textContent = '注入失敗：' + e;
      }
    }

    /** 依勾選數量更新下載按鈕文字與顯示狀態。 */
    function updateBar() {
      const count = document.querySelectorAll('.kenny-checkbox:checked').length;
      dlBar.style.display  = count ? '' : 'none';
      dlBtn.textContent    = `下載選取的清單 (${count})`;
    }

    /** 逐筆呼叫 API 取得資料，合併後匯出 Excel。 */
    async function downloadSelected() {
      const ids    = [...document.querySelectorAll('.kenny-checkbox:checked')].map(c => c.value);
      if (!ids.length) return;
      const key    = getKey();
      const spInfo = getSpInfo();
      dlBtn.disabled   = true;
      msgDiv.innerHTML = '';
      const rows = [];
      for (const id of ids) {
        try {
          msgDiv.innerHTML += `下載 [${id}]...<br/>`;
          const ret = await fetchChecklistData(id, key);
          const loc = await fetchHotspot(ret.locId, key);
          for (const obs of ret.obs) {
            const sp = spInfo[obs.speciesCode];
            if (!sp) continue; // 未知鳥種（雜交、形態群等）略過
            rows.push(buildObsRow(ret, obs, sp, loc?.name, loc?.latitude, loc?.longitude, ret.countyCode));
          }
        } catch (e) {
          msgDiv.innerHTML += `失敗 [${id}]：${e.message}<br/>`;
        }
      }
      msgDiv.innerHTML += '產生 Excel...<br/>';
      xlsxDownload(rows, `checklist_${Date.now()}.xlsx`);
      msgDiv.innerHTML += '下載完成';
      dlBtn.disabled = false;
    }
  }

  // ── 行程下載模組 (/mytripreports) ─────────────────────────────────────────

  /**
   * 在「我的行程報告」頁面，為每筆行程注入下載按鈕。
   *
   * DOM 依賴：
   *   #my-reports-heading        → 面板插入錨點
   *   #my-reports-items li       → 行程列表項目
   *   .Heading.Heading--h4       → 行程 heading（id 即為 tripId）
   */
  function initTrip() {
    const anchor = document.getElementById('my-reports-heading');
    if (!anchor) return;

    const root = document.createElement('div');
    anchor.after(root);

    const { panel, msgDiv } = createSetupPanel(injectButtons);
    root.appendChild(panel);

    injectButtons();

    /** 為每筆行程注入「下載」按鈕，設定完成前不執行。 */
    function injectButtons() {
      document.querySelectorAll('.kenny-dl-trip').forEach(n => n.remove());
      if (!getKey() || !Object.keys(getSpInfo()).length) return;

      try {
        const lis = document.getElementById('my-reports-items')?.getElementsByTagName('li');
        if (!lis) return;
        for (const li of lis) {
          const heading = li.getElementsByClassName('Heading Heading--h4')[0];
          if (!heading) continue;
          const tripId = heading.id;
          const btn = h('button', {
            className: 'Button Button--highlight kenny-dl-trip',
            style: 'margin-left:8px;',
            async onclick() { await downloadTrip(tripId, btn); },
          }, ['下載']);
          li.appendChild(btn);
        }
      } catch (e) {
        msgDiv.textContent = '注入失敗：' + e;
      }
    }

    /**
     * 取得行程下所有清單並合併匯出 Excel。
     * loc 資訊直接從行程 API 取得，省去逐筆查詢熱點。
     *
     * @param {string} tripId - 行程 ID
     * @param {HTMLButtonElement} btn - 觸發的按鈕（用於禁用/啟用）
     */
    async function downloadTrip(tripId, btn) {
      const key    = getKey();
      const spInfo = getSpInfo();
      btn.disabled     = true;
      msgDiv.innerHTML = `開啟下載 Trip id=${tripId}<br/>`;
      try {
        const checklists = await fetchTripChecklists(tripId);
        const rows = [];
        for (const c of checklists) {
          try {
            msgDiv.innerHTML += `下載 [${c.subId}]...<br/>`;
            const ret = await fetchChecklistData(c.subId, key);
            for (const obs of ret.obs) {
              const sp = spInfo[obs.speciesCode];
              if (!sp) continue;
              rows.push(buildObsRow(ret, obs, sp, c.loc.locName, c.loc.lat, c.loc.lng, ''));
            }
          } catch (e) {
            msgDiv.innerHTML += `失敗 [${c.subId}]：${e.message}<br/>`;
          }
        }
        msgDiv.innerHTML += '產生 Excel...<br/>';
        xlsxDownload(rows, `trip_${tripId}.xlsx`);
        msgDiv.innerHTML += '下載完成';
      } catch (e) {
        msgDiv.innerHTML += `Trip 下載失敗：${e.message}`;
      }
      btn.disabled = false;
    }
  }

  // ── 群組名片夾模組 (/checklist/*) ────────────────────────────────────────

  /**
   * 在清單分享頁注入「群組通訊錄」UI，可建立、編輯、刪除聯絡人群組，
   * 並一鍵將群組成員填入分享欄位。
   *
   * 群組資料以 JSON 陣列存於 localStorage['ebird-groups']，格式：
   *   [{ name: string, ebirdIds: string }]  (ebirdIds 以逗號分隔)
   *
   * DOM 依賴：
   *   #share-contacts-fieldset   → 面板插入錨點
   *   #share-contacts .u-inset-squish-sm  → eBird 既有聯絡人列表
   *   #share-recipients          → 分享收件人輸入框
   */
  function initUsersGroup() {
    const anchor = document.getElementById('share-contacts-fieldset');
    if (!anchor) return;

    const state = {
      editMode: false,
      editId:   -1,                       // -1 表示新增模式，≥0 表示編輯既有群組
      editData: { name: '', ebirdIds: '' },
      users:    [],                        // 從頁面讀取的 eBird 既有聯絡人
      groups:   [],                        // localStorage 中的自訂群組
    };

    // 從 eBird 既有聯絡人列表讀取 name 與 eBird ID
    const contactItems = document.getElementById('share-contacts')
      ?.getElementsByClassName('u-inset-squish-sm');
    if (contactItems) {
      for (const item of contactItems) {
        // onclick 屬性格式：addShareContact('userId')
        const m = (item.getAttribute('onclick') || '').match(/'([^']+)'/);
        state.users.push({ name: item.textContent.trim(), id: m ? m[1] : '' });
      }
    }
    try {
      state.groups = JSON.parse(localStorage.getItem('ebird-groups') || '[]');
    } catch { /**/ }

    const root = document.createElement('div');
    anchor.after(root);

    /** 重繪根節點（切換清單視圖與編輯器）。 */
    function render() {
      root.innerHTML = '';
      root.appendChild(state.editMode ? renderEditor() : renderList());
    }

    /** 清單視圖：無群組時顯示「全部成員」，有群組時列出各群組。 */
    function renderList() {
      const listItems = state.groups.length === 0
        ? [h('a', {
            href: '#share-recipients',
            className: 'u-inset-squish-sm',
            style: 'display:block;border-bottom:1px solid #efefef;',
            onclick() {
              // 一鍵填入所有已知聯絡人 ID
              document.getElementById('share-recipients').value =
                state.users.map(u => u.id).join(',');
            },
          }, ['全部成員'])]
        : state.groups.map((g, i) =>
            h('span', {
              className: 'u-inset-squish-sm',
              style: 'display:block;border-bottom:1px solid #efefef;',
            }, [
              h('a', { href: '#share-recipients', onclick() { shareGroup(g.ebirdIds); } }, [g.name]),
              h('a', { href: '#', style: 'float:right;', onclick(e) { e.preventDefault(); openEditor(i); } }, ['編輯']),
            ])
          );

      return h('div', {}, [
        h('span', { className: 'Heading Heading--h6' }, [
          h('span', { className: 'Heading-main' }, ['群組通訊錄']),
        ]),
        h('div', {
          style: 'background:white;max-height:8rem;overflow-y:auto;' +
                 'box-shadow:inset 0 1px 2px rgba(33,33,33,.1);border:1px solid #d3d3d3;',
        }, listItems),
        h('p', { className: 'u-text-2' }, [
          h('a', { href: '#', onclick(e) { e.preventDefault(); openEditor(-1); } }, ['新增群組']),
        ]),
      ]);
    }

    /** 編輯器視圖：輸入群組名稱與成員 ID，並提供從既有聯絡人快速點選。 */
    function renderEditor() {
      const nameInput = h('input', { type: 'text', className: 'u-text-2' });
      nameInput.value = state.editData.name;

      const idsArea = h('textarea', { className: 'u-text-2', style: 'display:block;width:100%;' });
      idsArea.value = state.editData.ebirdIds;

      // 點擊聯絡人名稱 → 追加 ID 至 textarea（不重複）
      const userLinks = state.users.map(u =>
        h('a', {
          href: '#',
          className: 'u-inset-squish-sm',
          style: 'display:block;border-bottom:1px solid #efefef;',
          onclick(e) {
            e.preventDefault();
            const existing = idsArea.value
              ? idsArea.value.split(',').map(s => s.trim()).filter(Boolean)
              : [];
            if (!existing.includes(u.id)) existing.push(u.id);
            idsArea.value = existing.join(',');
          },
        }, [u.name])
      );

      const buttons = [
        h('button', {
          className: 'Button Button--small Button--highlight u-inline-sm',
          onclick() {
            const name = nameInput.value.trim();
            const ids  = idsArea.value.trim();
            if (!name || !ids) return;
            if (state.editId === -1) {
              state.groups.push({ name, ebirdIds: ids });
            } else {
              state.groups[state.editId] = { name, ebirdIds: ids };
            }
            localStorage.setItem('ebird-groups', JSON.stringify(state.groups));
            state.editMode = false;
            render();
          },
        }, ['存檔']),
        h('button', {
          className: 'Button Button--small Button--secondary u-inline-sm',
          onclick() { state.editMode = false; render(); },
        }, ['取消']),
      ];

      // 編輯既有群組時才顯示刪除按鈕
      if (state.editId !== -1) {
        buttons.push(h('button', {
          className: 'Button Button--small Button--secondary',
          onclick() {
            state.groups.splice(state.editId, 1);
            localStorage.setItem('ebird-groups', JSON.stringify(state.groups));
            state.editMode = false;
            render();
          },
        }, ['刪除']));
      }

      return h('div', {}, [
        h('span', { className: 'Heading Heading--h6' }, [
          h('span', { className: 'Heading-main' }, ['編輯群組']),
        ]),
        nameInput,
        h('span', { className: 'Heading-sub Heading-sub--inline' }, ['(用戶名稱或電子郵件地址，以逗號分隔)']),
        idsArea,
        h('div', { style: 'background:white;max-height:8rem;overflow-y:auto;border:1px solid #d3d3d3;' }, userLinks),
        h('div', { className: 'ButtonGroup', style: 'margin-bottom:20px;' }, buttons),
      ]);
    }

    /**
     * 進入編輯模式。
     * @param {number} id - 群組索引，-1 為新增
     */
    function openEditor(id) {
      state.editId   = id;
      state.editData = id === -1
        ? { name: '群組名稱', ebirdIds: '' }
        : { name: state.groups[id].name, ebirdIds: state.groups[id].ebirdIds };
      state.editMode = true;
      render();
    }

    /**
     * 將群組 ID 追加至 #share-recipients，已存在的 ID 不重複加入。
     * @param {string} ebirdIds - 逗號分隔的 eBird ID 字串
     */
    function shareGroup(ebirdIds) {
      const input    = document.getElementById('share-recipients');
      const existing = input.value
        ? input.value.split(',').map(s => s.trim()).filter(Boolean)
        : [];
      const newIds   = ebirdIds.split(',').map(s => s.trim()).filter(s => !existing.includes(s));
      input.value    = [...existing, ...newIds].join(',');
    }

    render();
  }

  // ── GPS 軌跡下載模組 (/checklist/*) ──────────────────────────────────────

  /**
   * 在含有 GPS 軌跡的清單頁，注入「下載軌跡」按鈕，輸出標準 KML 檔。
   *
   * 軌跡座標抓取策略（優先順序）：
   *   1. data-maptrack-data="lng,lat,lng,lat,..."  → 新版 eBird 頁面
   *   2. :path="[lng,lat,lng,lat,...]"             → Vue 元件 prop（備用）
   *
   * DOM 依賴：
   *   #tracks-map-mini 或 #tracks  → 按鈕插入錨點
   *   section[aria-labelledby="primary-details"] .Heading--h3 span → 地點名稱
   *   time[datetime]               → 日期時間
   */
  function initTrack() {
    if (document.getElementById('ebird-track-dl')) return; // 防止重複注入
    const anchor = document.getElementById('tracks-map-mini') || document.getElementById('tracks');
    if (!anchor) return;

    const btn = document.createElement('button');
    btn.id        = 'ebird-track-dl';
    btn.type      = 'button';
    btn.textContent = '下載軌跡';
    btn.className   = 'Button Button--highlight';
    btn.style.cssText = 'display:block;margin:8px 0;';
    btn.addEventListener('click', async function () {
      const label  = btn.textContent;
      btn.disabled = true;
      btn.textContent = '下載中…';
      try {
        const coords = await fetchTrackCoords();
        if (!coords.length) {
          alert('找不到 GPS 軌跡資料 (請確認已登入，且此清單有可見的 GPS 軌跡)');
          return;
        }
        downloadKML(coords);
      } catch (e) {
        alert('下載軌跡失敗：' + (e && e.message ? e.message : e));
      } finally {
        btn.disabled    = false;
        btn.textContent = label;
      }
    });
    anchor.after(btn);
  }

  /**
   * 重新 fetch 當前頁面 HTML，從中解析 GPS 座標陣列。
   * 需要登入 cookie，因此用原生 fetch 帶 credentials。
   *
   * @returns {Promise<Array<[number, number]>>} [lng, lat] 配對陣列
   */
  async function fetchTrackCoords() {
    const res = await fetch(location.href, { credentials: 'include' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const html = await res.text();

    let nums = null;

    // 策略一：新版 eBird 將座標序列化為 data 屬性
    let m = html.match(/data-maptrack-data="([^"]*)"/);
    if (m) {
      nums = m[1].split(',').map(Number);
    } else {
      // 策略二：Vue 元件 :path prop（格式為 JSON 陣列）
      m = html.match(/:path="(\[[^"]*\])"/);
      if (m) { try { nums = JSON.parse(m[1]); } catch { nums = null; } }
    }
    if (!nums) return [];

    // 數字序列格式：lng1, lat1, lng2, lat2, ...
    const coords = [];
    for (let i = 0; i + 1 < nums.length; i += 2) {
      const lng = nums[i], lat = nums[i + 1];
      if (Number.isFinite(lng) && Number.isFinite(lat)) coords.push([lng, lat]);
    }
    return coords;
  }

  /**
   * 將座標陣列組成 KML 字串並觸發下載。
   * 檔名取自頁面地點名稱，並過濾 Windows 不合法字元。
   *
   * @param {Array<[number, number]>} coords - [lng, lat] 配對陣列
   */
  function downloadKML(coords) {
    const sid     = (location.pathname.match(/\/(S\d+)\b/) || [])[1] || '';
    const timeEl  = document.querySelector('time[datetime]');
    const [date, time] = (timeEl ? timeEl.getAttribute('datetime') : 'T').split('T');
    const locname = getLocname();
    const coordStr = coords.map(c => `${c[0]},${c[1]},0`).join('\n');
    const kml = buildKML(locname, [sid, date, time].filter(Boolean).join(' '), coordStr);
    blobDownload(kml, kmlSanitize(locname) + '.kml', 'application/vnd.google-earth.kml+xml');
  }

  /**
   * 從頁面標題區域取得地點名稱，找不到時回傳 'track'。
   *
   * @returns {string} 地點名稱
   */
  function getLocname() {
    const base = 'section[aria-labelledby="primary-details"] .Heading.Heading--h3.u-margin-none';
    const link = document.querySelector(base + ' a span');
    if (link) return link.textContent.trim();
    const span = document.querySelector(base + ' span');
    return span ? span.textContent.trim() : 'track';
  }

  /**
   * 產生 KML 2.2 格式字串，包含單條紅色 LineString 軌跡。
   *
   * @param {string} docName       - KML Document 名稱（地點）
   * @param {string} placemarkName - Placemark 名稱（清單 ID + 日期）
   * @param {string} coordStr      - 座標字串（lng,lat,alt 以換行分隔）
   * @returns {string} KML XML 字串
   */
  function buildKML(docName, placemarkName, coordStr) {
    const x = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return (
      '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<kml xmlns="http://www.opengis.net/kml/2.2">\n' +
      '  <Document>\n' +
      `    <name>${x(docName)}</name>\n` +
      '    <Style id="s"><LineStyle><color>ff0000ff</color><width>4</width></LineStyle></Style>\n' +
      '    <Placemark>\n' +
      `      <name>${x(placemarkName)}</name>\n` +
      '      <styleUrl>#s</styleUrl>\n' +
      '      <LineString><tessellate>1</tessellate>\n' +
      `        <coordinates>${coordStr}</coordinates>\n` +
      '      </LineString>\n' +
      '    </Placemark>\n' +
      '  </Document>\n' +
      '</kml>'
    );
  }

  /**
   * 移除 Windows 檔案系統不允許的字元，避免下載時失敗。
   *
   * @param {string} s - 原始字串
   * @returns {string} 合法檔名字串
   */
  function kmlSanitize(s) {
    return String(s).replace(/[\\/:*?"<>|]/g, '_').trim() || 'track';
  }

  /**
   * 將字串內容包成 Blob，建立暫時的物件 URL 並觸發瀏覽器下載。
   * 使用 setTimeout 延遲釋放 URL，確保下載完成前 URL 有效。
   *
   * @param {string} content  - 檔案內容
   * @param {string} filename - 下載檔名
   * @param {string} mime     - MIME 類型
   */
  function blobDownload(content, filename, mime) {
    const blob = new Blob([content], { type: mime });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

})();
