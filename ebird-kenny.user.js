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

(function () {
  'use strict';

  // ── Constants ─────────────────────────────────────────────────────────────

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

  const PRO_CODE = {
    P20: 'eBird - Casual Observation',
    P21: 'eBird - Stationary Count',
    P22: 'eBird - Traveling Count',
    P23: 'eBird - Exhaustive Area Count',
  };

  // ── Network ───────────────────────────────────────────────────────────────

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

  function fetchTaxonomy(lang) {
    return gmFetch(`https://api.ebird.org/v2/ref/taxonomy/ebird?fmt=json&locale=${lang}`);
  }

  function fetchChecklistData(subId, key) {
    return gmFetch(`https://api.ebird.org/v2/product/checklist/view/${subId}`, key);
  }

  async function fetchHotspot(locId, key) {
    try { return await gmFetch(`https://api.ebird.org/v2/ref/hotspot/info/${locId}`, key); }
    catch { return null; }
  }

  function fetchTripChecklists(tripId) {
    return gmFetch(`https://ebird.org/tripreport-internal/v1/checklists/${tripId}`);
  }

  // ── Storage ───────────────────────────────────────────────────────────────

  const getKey = () => localStorage.getItem('ebirdKey') || '';
  const setKey = k => localStorage.setItem('ebirdKey', k);
  const clearKey = () => localStorage.removeItem('ebirdKey');
  const getSpInfo = () => JSON.parse(localStorage.getItem('sp_info') || '{}');
  const setSpInfo = d => localStorage.setItem('sp_info', JSON.stringify(d));

  // ── Excel ─────────────────────────────────────────────────────────────────

  function buildObsRow(ret, obs, sp, location, lat, lng, county) {
    const [date, time] = ret.obsDt.split(' ');
    return {
      'Submission ID': ret.subId,
      'Common Name': sp.c,
      'Scientific Name': sp.s,
      'Taxonomic Order': sp.o,
      Count: obs.howManyAtmost,
      'State/Province': ret.subnational1Code,
      County: county || '',
      'Location ID': ret.locId,
      Location: location || '',
      Latitude: lat || '',
      Longitude: lng || '',
      Date: date,
      Time: time || '',
      Protocol: PRO_CODE[ret.protocolId] || ret.protocolId || '',
      'Duration (Min)': ret.durationHrs ? Math.round(parseFloat(ret.durationHrs) * 60) : '',
      'All Obs Reported': ret.allObsReported ? 1 : 0,
      'Distance Traveled (km)': ret.effortDistanceKm || '',
      'Area Covered (ha)': ret.effortAreaHa || '',
      'Number of Observers': ret.numObservers || '',
      'Breeding Code': obs.obsAux && obs.obsAux[0] ? (BREEDING_CODE[obs.obsAux[0].auxCode] || '') : '',
      'Observation Details': obs.comments || '',
      'Checklist Comments': ret.comments || '',
    };
  }

  function xlsxDownload(rows, filename) {
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'sheet1');
    XLSX.writeFile(wb, filename);
  }

  // ── DOM helper ────────────────────────────────────────────────────────────

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

  // ── Setup panel (shared by checklist + trip) ──────────────────────────────
  //
  // Returns { panel, msgDiv }.
  // onReady() is called whenever key or sp_info changes so the caller can
  // re-inject its buttons/checkboxes.

  function createSetupPanel(onReady) {
    const panel = document.createElement('div');
    panel.className = 'Page-section-inner Page-section-inner--md';

    const msgDiv = h('div', { style: 'margin-top:8px;' });
    panel.appendChild(msgDiv);

    function render() {
      // Remove all children except the persistent msgDiv
      while (panel.firstChild && panel.firstChild !== msgDiv) {
        panel.removeChild(panel.firstChild);
      }

      const key = getKey();
      const spInfo = getSpInfo();

      if (!key) {
        const keyInput = h('input', { type: 'text', style: 'margin-left:8px;width:260px;' });
        const saveBtn = h('button', {
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

      if (Object.keys(spInfo).length === 0) {
        const langSel = h('select', { style: 'margin:0 8px;' }, [
          h('option', { value: 'zh_TW' }, ['中文俗名']),
          h('option', { value: 'en' }, ['英文俗名']),
        ]);
        const dlBtn = h('button', {
          className: 'Button Button--highlight',
          async onclick() {
            msgDiv.textContent = '正在下載鳥名錄...';
            try {
              const list = await fetchTaxonomy(langSel.value);
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

  // ─────────────────────────────────────────────────────────────────────────
  // ROUTER
  // ─────────────────────────────────────────────────────────────────────────

  const path = location.pathname;

  if (/\/mychecklists/.test(path)) {
    initChecklist();
  } else if (/\/mytripreports/.test(path)) {
    initTrip();
  } else if (/\/checklist\//.test(path)) {
    initUsersGroup();
    initTrack();
  }

  // ─────────────────────────────────────────────────────────────────────────
  // CHECKLIST MODULE  /mychecklists
  // ─────────────────────────────────────────────────────────────────────────

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

    function injectCheckboxes() {
      document.querySelectorAll('.kenny-checkbox').forEach(n => n.remove());
      if (!getKey() || !Object.keys(getSpInfo()).length) { dlBar.style.display = 'none'; return; }

      try {
        const items = document
          .getElementById('place-species-observed-results')
          ?.getElementsByClassName('ResultsStats ResultsStats--manageMyChecklists');
        if (!items) return;
        for (const item of items) {
          const id = item.id.split('-')[1];
          const cb = document.createElement('input');
          cb.type = 'checkbox';
          cb.value = id;
          cb.className = 'kenny-checkbox';
          cb.style.marginRight = '6px';
          cb.addEventListener('change', updateBar);
          item.getElementsByClassName('Color-text-neutral-4 u-text-3')[0]?.prepend(cb);
        }
      } catch (e) {
        msgDiv.textContent = '注入失敗：' + e;
      }
    }

    function updateBar() {
      const count = document.querySelectorAll('.kenny-checkbox:checked').length;
      dlBar.style.display = count ? '' : 'none';
      dlBtn.textContent = `下載選取的清單 (${count})`;
    }

    async function downloadSelected() {
      const ids = [...document.querySelectorAll('.kenny-checkbox:checked')].map(c => c.value);
      if (!ids.length) return;
      const key = getKey();
      const spInfo = getSpInfo();
      dlBtn.disabled = true;
      msgDiv.innerHTML = '';
      const rows = [];
      for (const id of ids) {
        try {
          msgDiv.innerHTML += `下載 [${id}]...<br/>`;
          const ret = await fetchChecklistData(id, key);
          const loc = await fetchHotspot(ret.locId, key);
          for (const obs of ret.obs) {
            const sp = spInfo[obs.speciesCode];
            if (!sp) continue;
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

  // ─────────────────────────────────────────────────────────────────────────
  // TRIP MODULE  /mytripreports
  // ─────────────────────────────────────────────────────────────────────────

  function initTrip() {
    const anchor = document.getElementById('my-reports-heading');
    if (!anchor) return;

    const root = document.createElement('div');
    anchor.after(root);

    const { panel, msgDiv } = createSetupPanel(injectButtons);
    root.appendChild(panel);

    injectButtons();

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

    async function downloadTrip(tripId, btn) {
      const key = getKey();
      const spInfo = getSpInfo();
      btn.disabled = true;
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

  // ─────────────────────────────────────────────────────────────────────────
  // USERS GROUP MODULE  /checklist/*
  // ─────────────────────────────────────────────────────────────────────────

  function initUsersGroup() {
    const anchor = document.getElementById('share-contacts-fieldset');
    if (!anchor) return;

    const state = {
      editMode: false,
      editId: -1,
      editData: { name: '', ebirdIds: '' },
      users: [],
      groups: [],
    };

    // Read eBird's existing contact list
    const contactItems = document.getElementById('share-contacts')
      ?.getElementsByClassName('u-inset-squish-sm');
    if (contactItems) {
      for (const item of contactItems) {
        const m = (item.getAttribute('onclick') || '').match(/'([^']+)'/);
        state.users.push({ name: item.textContent.trim(), id: m ? m[1] : '' });
      }
    }
    try {
      state.groups = JSON.parse(localStorage.getItem('ebird-groups') || '[]');
    } catch { /**/ }

    const root = document.createElement('div');
    anchor.after(root);

    function render() {
      root.innerHTML = '';
      root.appendChild(state.editMode ? renderEditor() : renderList());
    }

    function renderList() {
      const listItems = state.groups.length === 0
        ? [h('a', {
            href: '#share-recipients',
            className: 'u-inset-squish-sm',
            style: 'display:block;border-bottom:1px solid #efefef;',
            onclick() {
              document.getElementById('share-recipients').value =
                state.users.map(u => u.id).join(',');
            },
          }, ['全部成員'])]
        : state.groups.map((g, i) =>
            h('span', {
              className: 'u-inset-squish-sm',
              style: 'display:block;border-bottom:1px solid #efefef;',
            }, [
              h('a', { href: '#share-recipients', onclick() { shareGroup(g.ebirdIds); }}, [g.name]),
              h('a', { href: '#', style: 'float:right;', onclick(e) { e.preventDefault(); openEditor(i); }}, ['編輯']),
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
          h('a', { href: '#', onclick(e) { e.preventDefault(); openEditor(-1); }}, ['新增群組']),
        ]),
      ]);
    }

    function renderEditor() {
      const nameInput = h('input', { type: 'text', className: 'u-text-2' });
      nameInput.value = state.editData.name;

      const idsArea = h('textarea', { className: 'u-text-2', style: 'display:block;width:100%;' });
      idsArea.value = state.editData.ebirdIds;

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
            const ids = idsArea.value.trim();
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

    function openEditor(id) {
      state.editId = id;
      state.editData = id === -1
        ? { name: '群組名稱', ebirdIds: '' }
        : { name: state.groups[id].name, ebirdIds: state.groups[id].ebirdIds };
      state.editMode = true;
      render();
    }

    function shareGroup(ebirdIds) {
      const input = document.getElementById('share-recipients');
      const existing = input.value
        ? input.value.split(',').map(s => s.trim()).filter(Boolean)
        : [];
      const newIds = ebirdIds.split(',').map(s => s.trim()).filter(s => !existing.includes(s));
      input.value = [...existing, ...newIds].join(',');
    }

    render();
  }

  // ─────────────────────────────────────────────────────────────────────────
  // TRACK MODULE  /checklist/*
  // ─────────────────────────────────────────────────────────────────────────

  function initTrack() {
    if (document.getElementById('ebird-track-dl')) return;
    const anchor = document.getElementById('tracks-map-mini') || document.getElementById('tracks');
    if (!anchor) return;

    const btn = document.createElement('button');
    btn.id = 'ebird-track-dl';
    btn.type = 'button';
    btn.textContent = '下載軌跡';
    btn.className = 'Button Button--highlight';
    btn.style.cssText = 'display:block;margin:8px 0;';
    btn.addEventListener('click', async function () {
      const label = btn.textContent;
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
        btn.disabled = false;
        btn.textContent = label;
      }
    });
    anchor.after(btn);
  }

  async function fetchTrackCoords() {
    const res = await fetch(location.href, { credentials: 'include' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const html = await res.text();

    let nums = null;
    let m = html.match(/data-maptrack-data="([^"]*)"/);
    if (m) {
      nums = m[1].split(',').map(Number);
    } else {
      m = html.match(/:path="(\[[^"]*\])"/);
      if (m) { try { nums = JSON.parse(m[1]); } catch { nums = null; } }
    }
    if (!nums) return [];

    const coords = [];
    for (let i = 0; i + 1 < nums.length; i += 2) {
      const lng = nums[i], lat = nums[i + 1];
      if (Number.isFinite(lng) && Number.isFinite(lat)) coords.push([lng, lat]);
    }
    return coords;
  }

  function downloadKML(coords) {
    const sid = (location.pathname.match(/\/(S\d+)\b/) || [])[1] || '';
    const timeEl = document.querySelector('time[datetime]');
    const [date, time] = (timeEl ? timeEl.getAttribute('datetime') : 'T').split('T');
    const locname = getLocname();
    const coordStr = coords.map(c => `${c[0]},${c[1]},0`).join('\n');
    const kml = buildKML(locname, [sid, date, time].filter(Boolean).join(' '), coordStr);
    blobDownload(kml, kmlSanitize(locname) + '.kml', 'application/vnd.google-earth.kml+xml');
  }

  function getLocname() {
    const base = 'section[aria-labelledby="primary-details"] .Heading.Heading--h3.u-margin-none';
    const link = document.querySelector(base + ' a span');
    if (link) return link.textContent.trim();
    const span = document.querySelector(base + ' span');
    return span ? span.textContent.trim() : 'track';
  }

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

  function kmlSanitize(s) {
    return String(s).replace(/[\\/:*?"<>|]/g, '_').trim() || 'track';
  }

  function blobDownload(content, filename, mime) {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

})();
