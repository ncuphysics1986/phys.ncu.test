/* ════════════════════════════════════════════════════════════════
   國立中央大學物理學系網站 · 渲染引擎（app.js）· 單一權威版本
   最後整理：2026-07-02（vb-text 防呆＋註解校正）
   ────────────────────────────────────────────────────────────────
   搭配 index.html（外殼＋CSS）與 8 個 *.xlsx（資料）。三者一起部署。
   ────────────────────────────────────────────────────────────────
   運作模型：
     init() → 讀 website.xlsx 設定 SECTIONS → 平行載入各分區 xlsx（LOADP）
     → buildNav()/buildMobileNav() 建選單 → navigateTo() → renderPage()
     renderPage() 以單一 switch 依「版面類型」分派給對應 renderXxx()（30+ 種）。
   資料契約：所有工作頁一律「依欄名」讀取；開關欄 isOn() 同時支援 1/0 與布林；
     順序欄 orderNum()；公告總庫由 parseNewsSheet() 依內容定位表頭。
   安全/無障礙：無行內事件處理器，全部事件委派（data-act/data-goto/data-onerr）
     ＋ JS 指派 .onclick / addEventListener（相容嚴格 CSP script-src 'self'）；
     javascript: 連結白名單（gotoArg）；Email 反爬（emailSpan/hydrateEmails）；
     行動版漢堡抽屜目錄（buildMobileNav）；每頁更新 document.title。
   ⚠ 整檔取代部署，勿手動合併片段。
═════════════════════════════════════════════════════════════════ */
'use strict';

/* ════════════════════════════════════════════════════
   SITE CONFIG  — overridable by website.xlsx
════════════════════════════════════════════════════ */
const SECTIONS = [
  { key:'intro',    label:'系所簡介', xlsxFile:'introduction.xlsx', show:true },
  { key:'faculty',  label:'系所人員', xlsxFile:'faculty.xlsx',      show:true },
  { key:'research', label:'學術成果', xlsxFile:'research.xlsx',     show:true },
  { key:'student',  label:'學生專區', xlsxFile:'student.xlsx',      show:true },
  { key:'emi',      label:'EMI專區',  xlsxFile:'emi.xlsx',           show:true },
  { key:'highschool', label:'高中生專區', xlsxFile:'highschool.xlsx', show:true },
  { key:'news',     label:'所有公告', xlsxFile:'news.xlsx',          show:true },
];

/* ── App state ── */
const DATA = {};
const LOADP = {};   /* LOADP[sectionKey] = 載入中的 Promise（漸進式載入） */
let curSection  = '';
let curPageId   = '';
let returnState = null;
let researchRows = [];

/* ════════════════════════════════════════════════════
   XLSX LOADER — always bypass cache
════════════════════════════════════════════════════ */
async function loadXlsx(filename) {
  /* cache:'no-cache' = 每次都向伺服器驗證 ETag：檔案沒變回 304（不重新下載），
     檔案更新自動抓新版 — 兼顧速度與即時性 */
  const resp = await fetch(filename, { cache:'no-cache' });
  if (!resp.ok) throw new Error('無法讀取 ' + filename);
  const lastMod = resp.headers.get('Last-Modified') || resp.headers.get('Date') || '';
  const wb = XLSX.read(await resp.arrayBuffer(), { type:'array' });
  const out = { _lastMod: lastMod };
  wb.SheetNames.forEach(name => {
    out[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval:'' });
  });
  return out;
}

/* website.xlsx — override key / xlsxFile / show */
async function applyWebsiteConfig() {
  try {
    const wb = await loadXlsx('website.xlsx');
    setVersionBadge(wb._lastMod);
    (wb['網頁架構'] || []).forEach(row => {
      const sec = SECTIONS.find(s => s.label === row['選單分類']);
      if (!sec) return;
      if (row['選單Key'])  sec.key      = row['選單Key'];
      if (row['xlsx檔名']) sec.xlsxFile = row['xlsx檔名'];
      if (row['是否顯示'] !== undefined && row['是否顯示'] !== '')
        sec.show = isOn(row['是否顯示']);
      const ord = parseFloat(row['顯示順序']);
      if (!isNaN(ord)) sec.order = ord;
    });
    /* 依「顯示順序」排序導覽列（未填者排最後，維持原相對順序） */
    SECTIONS.sort((a, b) => (a.order ?? 999) - (b.order ?? 999));
  } catch(_) { setVersionBadge(''); }
}

function setVersionBadge(lastMod) {
  const el = document.getElementById('vb-text');
  if (!el) return;   /* 外殼缺 #vb-text 時安靜略過（防呆：避免例外中斷啟動流程） */
  if (!lastMod) { el.textContent = '已載入'; return; }
  try {
    const d = new Date(lastMod);
    if (isNaN(d)) { el.textContent = '已載入'; return; }
    const p = n => String(n).padStart(2,'0');
    el.textContent = `已更新　${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  } catch(_) { el.textContent = '已載入'; }
}

/* ════════════════════════════════════════════════════
   UTILITIES
════════════════════════════════════════════════════ */
function esc(v) {
  return String(v ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* 統一的顯示開關判斷 — 取代散落各處的 on==='TRUE'||'True'||1… 容錯碼。
   接受真布林、'TRUE'/'True'、1、'是'、'Y' 等為「顯示」；
   僅 false、'FALSE'、'否'、'0'、'N'、空白 視為「隱藏」。 */
function isOn(v) {
  if (v === true)  return true;
  if (v === false) return false;
  const s = String(v ?? '').trim().toUpperCase();
  return !(s === '' || s === 'FALSE' || s === '否' || s === '0' || s === 'N' || s === 'NO');
}
/* 人員工作頁的「啟用」顯示旗標：1=顯示、0/否/N/FALSE=隱藏、空白或整欄缺漏=顯示
   （安全預設，與手冊「空=顯示」一致，避免新增列漏填而整批消失）。 */
function rowEnabled(r) {
  if (!('啟用' in r)) return true;
  const v = String(r['啟用'] ?? '').trim();
  return v === '' || isOn(v);
}
/* 統一的順序值 — 數字優先，缺值或非數字排最後（99） */
function orderNum(v) {
  const n = parseFloat(v);
  return isNaN(n) ? 99 : n;
}
/* 安全解碼網址片段：location.hash 可能被瀏覽器百分比編碼（如中文分頁ID），
   解碼後才能對上 getPages 的分頁ID；解碼失敗則原樣返回 */
function decSeg(s) {
  try { return decodeURIComponent(s || ''); } catch (_) { return s || ''; }
}
/* Email 反爬：輸出的 HTML 不含 @ 字面（拆成 data-u / data-d，顯示用「[at]」），
   bot 的正規表示式抓不到完整信箱；載入後由 hydrateEmails() 在前端組回為純文字顯示
   （不建立 mailto 連結，反爬效果更佳——無 href 可供爬蟲擷取） */
function emailSpan(addr) {
  const s = String(addr || '').trim();
  const at = s.indexOf('@');
  if (at < 1) return esc(s);
  const u = s.slice(0, at), d = s.slice(at + 1);
  return `<span class="eml" data-u="${esc(u)}" data-d="${esc(d)}">✉ ${esc(u)}<span class="eml-sep"> [at] </span>${esc(d)}</span>`;
}
function hydrateEmails(root) {
  (root || document).querySelectorAll('span.eml:not([data-eml-done])').forEach(el => {
    const u = el.getAttribute('data-u'), d = el.getAttribute('data-d');
    el.setAttribute('data-eml-done', '1');
    if (!u || !d) return;
    el.textContent = '✉ ' + u + '@' + d;   /* 純文字呈現，不建立 mailto 連結 */
  });
}

function toBullets(str) {
  if (!str) return '';
  const lines = String(str).split('\n').map(s => s.trim()).filter(Boolean);
  if (lines.length === 1) return `<p>${esc(lines[0])}</p>`;
  return '<ul>' + lines.map(l => `<li>${esc(l)}</li>`).join('') + '</ul>';
}

function fileBadge(type) {
  const t = String(type || '').toUpperCase().trim();
  if (t === 'PDF')                                             return `<span class="badge b-pdf">PDF</span>`;
  if (t === 'DOC' || t === 'DOCX')                            return `<span class="badge b-doc">${esc(t)}</span>`;
  if (t === 'YOUTUBE')                                        return `<span class="badge b-yt">YouTube</span>`;
  if (t === '' || t === '連結' || t === 'LINK' || t === 'URL') return `<span class="badge b-link">連結</span>`;
  return `<span class="badge b-other">${esc(type)}</span>`;
}

function getPages(sectionKey) {
  const sheets = DATA[sectionKey];
  if (!sheets) return [];
  return (sheets['頁面設定'] || [])
    .filter(r => {
      const id = r['分頁ID'];
      return id && id !== '［說明］' && isOn(r['啟用']);
    })
    .sort((a,b) => orderNum(a['排列順序']) - orderNum(b['排列順序']))
    .map(r => ({ id:r['分頁ID'], label:r['分頁標題'], cfg:r }));
}

/* ── goTo: jump to any page from content links
   Usage in xlsx: javascript:goTo('faculty/faculty') ── */
let newsPendingParam = '';  /* goTo 第三段參數：數字=WP-ID（展開該則公告），文字=分類頁籤 */

function goTo(hash) {
  const parts = hash.replace('#','').split('/');
  const sec = parts[0], pg = parts[1];
  if (parts[2]) newsPendingParam = decodeURIComponent(parts[2]);
  const section = SECTIONS.find(s => s.show && s.key === sec);
  if (section && pg) {
    navigateTo(sec, pg);
  }
}

/* ════════════════════════════════════════════════════
   NAV — built dynamically from SECTIONS + getPages
════════════════════════════════════════════════════ */
function buildNav() {
  const menu = document.getElementById('main-menu');
  menu.innerHTML = '';
  SECTIONS.filter(s => s.show).forEach(sec => {
    const li = document.createElement('li');
    li.dataset.section = sec.key;
    const a  = document.createElement('a');
    a.href = '#';
    a.textContent = sec.label;
    a.onclick = async e => {
      e.preventDefault();
      if (!DATA[sec.key] && LOADP[sec.key]) {
        document.getElementById('list-view').innerHTML = '<div class="loading">載入資料中</div>';
        try { await LOADP[sec.key]; } catch(_) {}
        buildNav();
      }
      const pg = getPages(sec.key);
      if (pg.length) navigateTo(sec.key, pg[0].id);
    };
    const ul = document.createElement('ul');
    getPages(sec.key).forEach(p => {
      const sli = document.createElement('li');
      const sa  = document.createElement('a');
      sa.href = '#'; sa.textContent = p.label;
      sa.onclick = e => { e.preventDefault(); navigateTo(sec.key, p.id); };
      sli.appendChild(sa); ul.appendChild(sli);
    });
    li.appendChild(a); li.appendChild(ul);
    menu.appendChild(li);
  });
  buildMobileNav();   /* 行動版目錄與桌機選單同步重建 */
}

/* ── 行動版抽屜式網頁目錄（與 buildNav 同源資料）── */
function buildMobileNav() {
  const body = document.getElementById('mobile-nav-body');
  if (!body) return;
  body.innerHTML = '';
  SECTIONS.filter(s => s.show).forEach(sec => {
    const secDiv = document.createElement('div');
    secDiv.className = 'mnav-sec';
    const title = document.createElement('button');
    title.type = 'button';
    title.className = 'mnav-sec-title';
    title.textContent = sec.label;
    title.onclick = async () => {
      if (!DATA[sec.key] && LOADP[sec.key]) { try { await LOADP[sec.key]; } catch (_) {} buildNav(); }
      const pg = getPages(sec.key);
      if (pg.length) { navigateTo(sec.key, pg[0].id); closeMobileNav(); }
    };
    secDiv.appendChild(title);
    const ul = document.createElement('ul');
    ul.className = 'mnav-sub';
    getPages(sec.key).forEach(p => {
      const li = document.createElement('li');
      const a  = document.createElement('a');
      a.href = '#';
      a.textContent = p.label;
      a.onclick = e => { e.preventDefault(); navigateTo(sec.key, p.id); closeMobileNav(); };
      li.appendChild(a); ul.appendChild(li);
    });
    secDiv.appendChild(ul);
    body.appendChild(secDiv);
  });
}

function openMobileNav() {
  const d = document.getElementById('mobile-nav');
  const o = document.getElementById('mobile-nav-overlay');
  const t = document.getElementById('nav-toggle');
  if (!d) return;
  d.classList.add('open'); o.classList.add('open');
  d.setAttribute('aria-hidden', 'false');
  if (t) t.setAttribute('aria-expanded', 'true');
  document.body.classList.add('nav-open');
  const c = document.getElementById('mobile-nav-close'); if (c) c.focus();
}
function closeMobileNav() {
  const d = document.getElementById('mobile-nav');
  const o = document.getElementById('mobile-nav-overlay');
  const t = document.getElementById('nav-toggle');
  if (!d) return;
  d.classList.remove('open'); o.classList.remove('open');
  d.setAttribute('aria-hidden', 'true');
  if (t) t.setAttribute('aria-expanded', 'false');
  document.body.classList.remove('nav-open');
}
/* 一次性綁定漢堡鈕 / 關閉鈕 / 遮罩 / Esc（靜態元素，載入即可綁定） */
(function initMobileNav() {
  const t = document.getElementById('nav-toggle');
  const c = document.getElementById('mobile-nav-close');
  const o = document.getElementById('mobile-nav-overlay');
  if (t) t.addEventListener('click', openMobileNav);
  if (c) c.addEventListener('click', closeMobileNav);
  if (o) o.addEventListener('click', closeMobileNav);
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeMobileNav(); });
})();

/* ════════════════════════════════════════════════════
   NAVIGATION
════════════════════════════════════════════════════ */
async function navigateTo(sectionKey, pageId) {
  document.getElementById('detail-view').style.display = 'none';
  document.getElementById('list-view').style.display   = '';

  /* 該分區資料尚在背景載入 → 顯示載入動畫並等待 */
  if (!DATA[sectionKey] && LOADP[sectionKey]) {
    document.getElementById('list-view').innerHTML = '<div class="loading">載入資料中</div>';
    try { await LOADP[sectionKey]; } catch(_) {}
  }

  curSection = sectionKey; curPageId = pageId;

  document.querySelectorAll('#main-menu > li')
    .forEach(li => li.classList.toggle('active', li.dataset.section === sectionKey));

  const pages = getPages(sectionKey);
  const page  = pages.find(p => p.id === pageId);
  if (!page) return;

  document.getElementById('hero-title').textContent = page.cfg['頁首大標題'] || page.label;
  document.getElementById('hero-sub').textContent   = page.cfg['頁首副標題'] || '';
  const bg = page.cfg['頁首圖URL'] || '';
  document.getElementById('hero-bg').style.backgroundImage = bg ? `url('${esc(bg)}')` : '';
  document.getElementById('bc-sec').textContent  = SECTIONS.find(s => s.key === sectionKey)?.label || sectionKey;
  document.getElementById('bc-page').textContent = page.label;
  /* 無障礙：SPA 換頁時更新分頁標題，利於螢幕報讀器與檢測工具辨識當前頁面 */
  document.title = `${page.label}｜國立中央大學物理學系`;

  const subNav = document.getElementById('sub-nav');
  subNav.innerHTML = '';
  pages.forEach(p => {
    const btn = document.createElement('button');
    btn.textContent = p.label;
    btn.classList.toggle('active', p.id === pageId);
    btn.onclick = () => navigateTo(sectionKey, p.id);
    subNav.appendChild(btn);
  });

  /* Update URL hash */
  const newHash = '#' + sectionKey + '/' + pageId;
  if (location.hash !== newHash) history.pushState(null, '', newHash);

  renderPage(sectionKey, page);
}

/* ════════════════════════════════════════════════════
   PAGE DISPATCHER  (single switch, no patches)
════════════════════════════════════════════════════ */
function renderPage(sectionKey, page) {
  const area      = document.getElementById('list-view');
  area.innerHTML  = '<div class="loading">載入中</div>';
  const sheets    = DATA[sectionKey];
  const layout    = page.cfg['版面類型']   || '';
  const sheetName = page.cfg['資料工作頁'] || '';
  const desc      = page.cfg['說明文字']   || '';

  setTimeout(() => {
    try {
      switch (layout) {
        case 'timeline':       renderTimeline(area, sheets, sheetName, desc);      break;
        case 'text':           renderText(area, desc);                             break;
        case 'cards':          renderChairCards(area, sheets, sheetName, page);    break;
        case 'table':          renderCompetency(area, sheets, sheetName);          break;
        case 'links':          renderRegulations(area, sheets, sheetName, page);   break;
        case 'events':         renderActivities(area, sheets, sheetName);          break;
        case 'faculty':        renderFaculty(area, sheets, sheetName, sectionKey, page); break;
        case 'staff':          renderStaff(area, sheets, sheetName);               break;
        case 'assistant':      renderAssistant(area, sheets, sheetName);           break;
        case 'phd':            renderPhd(area, sheets, sheetName);                 break;
        case 'research':       renderResearch(area, sheets, sheetName);            break;
        case 'lab':            renderLab(area, sheets, sheetName);                 break;
        case 'awards':         renderAwards(area, sheets, sheetName);              break;
        case 'visitors':       renderVisitors(area, sheets, sheetName);            break;
        case 'course_ug':      renderCourseUG(area, sheets, sheetName);            break;
        case 'course_pg':      renderCoursePG(area, sheets, sheetName);            break;
        case 'annualmeeting':  renderAnnualMeeting(area, sheets);                  break;
        case 'scholarship':    renderScholarship(area, sheets, sheetName);         break;
        case 'space':          renderSpaces(area, sheets, sheetName);              break;
        case 'student_awards': renderStudentAwards(area, sheets, sheetName);       break;
        case 'theses':         renderTheses(area, sheets, sheetName);              break;
        case 'travel':         renderStudentTravel(area, sheets, sheetName);       break;
        case 'program':        renderPrograms(area, sheets, sheetName);            break;
        case 'journal':        renderJournals(area, sheets, sheetName);            break;
        case 'emi':            renderEmi(area, sheets, sheetName);                 break;
        case 'portal':         renderPortal(area, sheets, sheetName);              break;
        case 'linkcards':      renderLinkCards(area, sheets, sheetName, page);     break;
        case 'newslist':       renderNewsListPage(area, sheets, sheetName, page);  break;
        case 'news':           renderNews(area, sheets, sheetName);                break;
        default:
          if (sheetName && sheets[sheetName]) renderGenericTable(area, sheets, sheetName);
          else renderText(area, desc);
      }
    } catch(e) {
      area.innerHTML = `<div class="error-msg">⚠ 頁面載入失敗：${esc(e.message)}</div>`;
    }
  }, 0);
}

/* ════════════════════════════════════════════════════
   RENDERERS
════════════════════════════════════════════════════ */

function renderTimeline(area, sheets, sheetName, desc) {
  const data = sheetDataRows(sheets[sheetName] || [], ['類型','時期']);
  const eras = [], events = [];
  data.rows.forEach(r => {
    const type = String(data.at(r,'類型')).trim();
    const rec = {
      era:   String(data.at(r,'時期')).trim() || '1',
      year:  String(data.at(r,'年份')).trim(),
      title: String(data.at(r,'標題')).trim(),
      sub:   String(data.at(r,'副標題')).trim(),
      body:  String(data.at(r,'內文')).trim(),
    };
    if (type === 'era') eras.push(rec);
    else if (type === 'event') events.push(rec);
  });

  const eraColor = { '1':'var(--gold)', '2':'var(--blue)', '3':'#1d9e75' };
  const eraIcon  = { '1':'🌱', '2':'🚀', '3':'⚛' };

  let html = '<h2 class="section-heading">歷史沿革</h2>';

  /* 依時期分組：左側該時期的事件群 + 右側該時期卡片，同列對齊 */
  const eraList = eras.length ? eras : [...new Set(events.map(e=>e.era))].sort().map(era=>({era,title:'',sub:'',body:''}));

  html += '<div class="tl-grid">';
  eraList.forEach(e => {
    const color = eraColor[e.era] || 'var(--gold)';
    const evs = events.filter(ev => ev.era === e.era);

    /* 左欄：此時期的時間節點群 */
    html += '<div class="tlg-left">';
    evs.forEach(it => {
      html += `<div class="tlv-item">`
        + `<div class="tlv-marker"><span class="tlv-dot" style="background:${color}"></span></div>`
        + `<div class="tlv-content">`
        + `<div class="tlv-year" style="color:${color}">${esc(it.year)}</div>`
        + `<div class="tlv-title">${esc(it.title)}</div>`
        + (it.body ? `<div class="tlv-note">${esc(it.body)}</div>` : '')
        + `</div></div>`;
    });
    html += '</div>';

    /* 右欄：此時期卡片 */
    html += `<div class="tlg-right"><div class="tl-era" style="border-left:4px solid ${color}">`
      + `<div class="tl-era-head"><span class="tl-era-icon">${esc(eraIcon[e.era]||'')}</span>`
      + `<span class="tl-era-title">${esc(e.title)}</span>`
      + (e.sub ? `<span class="tl-era-sub">${esc(e.sub)}</span>` : '')
      + `</div>`;
    if (e.body) {
      html += '<div class="text-content tl-era-body">';
      e.body.split('\n\n').forEach(p => { if (p.trim()) html += `<p>${esc(p.trim())}</p>`; });
      html += '</div>';
    }
    html += '</div></div>';
  });
  html += '</div>';

  /* 相容：無 era 段落且有說明文字時，附在末尾 */
  if (!eras.length && desc) {
    html += '<div class="text-content" style="margin-top:24px">';
    desc.split('\n\n').forEach(p => { if (p.trim()) html += `<p>${esc(p.trim())}</p>`; });
    html += '</div>';
  }

  area.innerHTML = html;
}

function renderText(area, desc) {
  if (!desc) { area.innerHTML = '<p style="color:var(--muted)">（此頁面暫無內容）</p>'; return; }
  let html = '<div class="text-content">';
  desc.split('\n\n').forEach(p => { if (p.trim()) html += `<p>${esc(p.trim())}</p>`; });
  area.innerHTML = html + '</div>';
}

function renderChairCards(area, sheets, sheetName, page) {
  const rows = sheets[sheetName] || [];
  let html = `<h2 class="section-heading">${esc(page.cfg['頁首大標題']||'')} <span class="section-count">${rows.length} 位</span></h2><div class="chair-grid">`;
  rows.forEach(r => {
    const photo = r['照片URL'] || '';
    const tags  = (r['專長']||'').split(/[,，]/).map(s=>s.trim()).filter(Boolean);
    html += '<div class="chair-card"><div class="chair-photo-wrap">';
    if (photo) {
      html += `<img loading="lazy" decoding="async" src="${esc(photo)}" alt="${esc(r['姓名']||'')}" data-onerr="hideflex">`;
      html += '<div class="chair-no-photo" style="display:none">👤</div>';
    } else { html += '<div class="chair-no-photo">👤</div>'; }
    html += `</div><div class="chair-body"><h3>${esc(r['姓名']||'')}</h3>`;
    html += `<div class="chair-dates">📅 ${esc(r['服務起']||'')} — ${esc(r['服務迄']||'')}</div>`;
    if (r['最高學歷']) html += `<div class="chair-edu">🎓 ${esc(r['最高學歷'])}</div>`;
    if (tags.length) html += '<div class="chair-tags">' + tags.map(t=>`<span class="chair-tag">${esc(t)}</span>`).join('') + '</div>';
    html += '</div></div>';
  });
  area.innerHTML = html + '</div>';
}

function renderFaculty(area, sheets, sheetName, sectionKey, page) {
  const rows = sheets[sheetName] || [];
  /* 選用篩選：頁面設定『篩選』欄，格式「欄名=值」（例：在職狀態=退休）。
     空白＝不篩選。用於把多個分頁共用同一張工作頁（如專任/退休共用「師資」）。 */
  const spec = (page && page.cfg && page.cfg['篩選']) ? String(page.cfg['篩選']).trim() : '';
  let fcol = '', fval = '';
  if (spec.includes('=')) { const i = spec.indexOf('='); fcol = spec.slice(0, i).trim(); fval = spec.slice(i + 1).trim(); }
  /* 保留原始列索引（idx）——詳情頁靠它回查整張工作頁，篩選後不可重新編號 */
  const visible = [];
  rows.forEach((r, idx) => {
    if (fcol && String(r[fcol] ?? '').trim() !== fval) return;
    if (!rowEnabled(r)) return;
    visible.push({ r, idx });
  });
  const heading = (page && page.label) || sheetName;
  // 只有「專任師資」頁的卡片可點開詳情；其餘師資頁（退休/訪問/兼任…）為純顯示卡（與靜態版一致）
  const detailable = sectionKey === 'faculty' && (!page || page.id === 'faculty' || /專任/.test(page.label || ''));
  let html = `<h2 class="section-heading">${esc(heading)} <span class="section-count">${visible.length} 位</span></h2><div class="faculty-grid">`;
  visible.forEach(({ r, idx }) => {
    const photo = r['照片URL'] || '';
    const detailAttrs = detailable
      ? ` role="button" tabindex="0" aria-label="${esc(r['姓名']||'')} 詳細資料" data-act="openFaculty" data-sec="${esc(sectionKey)}" data-idx="${idx}" data-sheet="${esc(sheetName)}"`
      : '';
    html += `<div class="faculty-card${detailable ? '' : ' faculty-card-static'}"${detailAttrs}>`;
    if (photo) {
      html += `<img class="fc-photo" loading="lazy" decoding="async" src="${esc(photo)}" alt="${esc(r['姓名']||'')}" data-onerr="hideflex">`;
      html += '<div class="fc-no-photo" style="display:none">👤</div>';
    } else { html += '<div class="fc-no-photo">👤</div>'; }
    html += `<div class="fc-info"><h3>${esc(r['姓名']||'')}</h3><div class="fc-title">${esc(r['職稱']||'')}</div>`;
    const areaLine = (r['研究領域 Research area']||'').split('\n')[0];
    if (areaLine) html += `<div class="fc-area">${esc(areaLine)}</div>`;
    if (r['Email']) html += `<div class="fc-email">${emailSpan(r['Email'])}</div>`;
    html += '</div></div>';
  });
  area.innerHTML = html + '</div>';
}

function openFacultyDetail(sectionKey, idx, sheetName) {
  returnState = { sectionKey, pageId: curPageId };
  const sheets = DATA[sectionKey];
  const r      = (sheets[sheetName]||[])[idx];
  if (!r) return;
  const name   = r['姓名']  || '';
  const title  = r['職稱']  || '';
  const photo  = r['照片URL'] || '';
  const pubs   = (sheets['研究著作'] ||[]).filter(p => p['姓名'] === name);
  const projs  = (sheets['研究計畫']||[]).filter(p => p['姓名'] === name);
  const labRaw = r['實驗室 Laboratory'] || r['實驗室'] || '';
  const labURL = (labRaw.match(/https?:\/\/[^\s\n]+/)||[])[0] || '';
  const labName= labRaw.replace(/\n?https?:\/\/[^\s\n]+/g,'').trim();

  document.getElementById('detail-view').innerHTML = `
<button class="fd-back" data-act="closeFaculty">← 返回列表</button>
<div class="fd-layout">
  <div class="fd-sidebar">
    ${photo
      ? `<img class="fd-photo" src="${esc(photo)}" alt="${esc(name)}" data-onerr="hideflex"><div class="fd-no-photo" style="display:none">👤</div>`
      : '<div class="fd-no-photo">👤</div>'}
    <div class="fd-contact">
      ${r['辦公室']     ? `<div>🏢 辦公室：${esc(r['辦公室'])}</div>`       : ''}
      ${r['研究室電話'] ? `<div>📞 研究室：${esc(r['研究室電話'])}</div>`   : ''}
      ${r['實驗室電話'] ? `<div>🔬 實驗室：${esc(r['實驗室電話'])}</div>`   : ''}
      ${r['Email']      ? `<div>${emailSpan(r['Email'])}</div>` : ''}
    </div>
  </div>
  <div class="fd-main">
    <h1>${esc(name)}</h1><div class="fd-rank">${esc(title)}</div>
    <hr class="fd-divider">
    ${(labName||labURL) ? `<div class="fd-section"><h3>實驗室 Laboratory</h3><p>${esc(labName)}${labURL?` <a href="${esc(labURL)}" target="_blank" rel="noopener noreferrer">${esc(labURL)}</a>`:''}</p></div>` : ''}
    ${r['現職 Position']         ? `<div class="fd-section"><h3>現職 Position</h3>${toBullets(r['現職 Position'])}</div>`                   : ''}
    ${r['學歷 Education']         ? `<div class="fd-section"><h3>學歷 Education</h3>${toBullets(r['學歷 Education'])}</div>`                   : ''}
    ${r['經歷 Experience']        ? `<div class="fd-section"><h3>經歷 Experience</h3>${toBullets(r['經歷 Experience'])}</div>`                  : ''}
    ${r['學術榮譽 Academic honor'] ? `<div class="fd-section"><h3>學術榮譽 Academic honor</h3>${toBullets(r['學術榮譽 Academic honor'])}</div>` : ''}
    ${r['研究領域 Research area']  ? `<div class="fd-section"><h3>研究領域 Research area</h3>${toBullets(r['研究領域 Research area'])}</div>`   : ''}
    <div class="fd-tabs">
      <button class="fd-tab-btn active" data-act="switchFdTab" data-tab="fd-projs">研究計畫 (${projs.length})</button>
      <button class="fd-tab-btn"        data-act="switchFdTab" data-tab="fd-pubs">研究著作 (${pubs.length})</button>
    </div>
    <div id="fd-projs">
      ${projs.length
        ? `<table class="proj-table"><thead><tr><th>計畫年度</th><th>計畫名稱</th><th>執行起迄</th></tr></thead><tbody>`
          + projs.map(p=>`<tr><td>${esc(String(p['計畫年度']||''))}</td><td>${esc(p['計畫名稱']||'')}</td><td>${esc(p['執行起迄']||'')}</td></tr>`).join('')
          + '</tbody></table>'
        : '<p style="color:var(--muted);font-size:.9rem">暫無研究計畫資料</p>'}
    </div>
    <div id="fd-pubs" style="display:none">
      ${pubs.length
        ? pubs.map(p=>`<div class="pub-item"><div class="pub-title">${esc(p['論文標題']||'')}</div><div class="pub-meta">${esc(p['出版期刊']||'')}${p['出版年度']?' · '+esc(String(p['出版年度']).replace('.0','')):''}</div></div>`).join('')
        : '<p style="color:var(--muted);font-size:.9rem">暫無著作資料</p>'}
    </div>
  </div>
</div>`;
  document.getElementById('list-view').style.display   = 'none';
  document.getElementById('detail-view').style.display = 'block';
  window.scrollTo(0,0);
  document.getElementById('hero-title').textContent = name;
  document.getElementById('hero-sub').textContent   = title;
  document.getElementById('bc-page').textContent    = name;
}

function closeFacultyDetail() {
  if (returnState) navigateTo(returnState.sectionKey, returnState.pageId);
}

function switchFdTab(btn, tabId) {
  document.querySelectorAll('.fd-tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  ['fd-projs','fd-pubs'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = id === tabId ? 'block' : 'none';
  });
}

function renderStaff(area, sheets, sheetName) {
  const cols = ['姓名','職稱','辦公室','Email','辦公室電話','工作執掌'];
  const rows = (sheets[sheetName] || []).filter(rowEnabled);
  let html = `<h2 class="section-heading">${esc(sheetName)} <span class="section-count">${rows.length} 位</span></h2><div class="table-wrap"><table class="data-table"><thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>`;
  rows.forEach(r => {
    html += '<tr>' + cols.map(c => `<td>${c==='Email'?emailSpan(r[c]):esc(r[c]||'')}</td>`).join('') + '</tr>';
  });
  area.innerHTML = html + '</tbody></table></div>';
}

function renderAssistant(area, sheets, sheetName) {
  const cols = ['姓名','職稱','辦公室','實驗室','Email','研究室電話','實驗室電話','實驗室 Laboratory'];
  const last = cols.length - 1;
  const rows = (sheets[sheetName] || []).filter(rowEnabled);
  let html = `<h2 class="section-heading">${esc(sheetName)} <span class="section-count">${rows.length} 位</span></h2><div class="table-wrap"><table class="data-table nowrap-table"><thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>`;
  rows.forEach(r => { html += '<tr>' + cols.map((c,i)=>`<td${i===last?' class="wrap-col"':''}>${c==='Email'?emailSpan(r[c]):esc(String(r[c]||''))}</td>`).join('') + '</tr>'; });
  area.innerHTML = html + '</tbody></table></div>';
}

function renderPhd(area, sheets, sheetName) {
  const cols = ['中文姓名','英文姓名','指導教授','研究室分機'];
  const rows = (sheets[sheetName] || []).filter(rowEnabled);
  let html = `<h2 class="section-heading">${esc(sheetName)} <span class="section-count">${rows.length} 位</span></h2><div class="table-wrap"><table class="data-table"><thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>`;
  rows.forEach(r => { html += '<tr>' + cols.map(c=>`<td>${r[c]!==''?esc(String(r[c])):''}</td>`).join('') + '</tr>'; });
  area.innerHTML = html + '</tbody></table></div>';
}

function renderGenericTable(area, sheets, sheetName) {
  const rows = sheets[sheetName] || [];
  if (!rows.length) { area.innerHTML = '<div class="error-msg">此工作頁無資料</div>'; return; }
  const cols = Object.keys(rows[0]).filter(k => k !== '_lastMod');
  let html = `<h2 class="section-heading">${esc(sheetName)}</h2><div class="table-wrap"><table class="data-table"><thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>`;
  rows.forEach(r => { html += '<tr>' + cols.map(c=>`<td>${esc(String(r[c]||''))}</td>`).join('') + '</tr>'; });
  area.innerHTML = html + '</tbody></table></div>';
}

function renderCompetency(area, sheets, sheetName) {
  const rows = sheets[sheetName] || [];
  const groupOrder = [], groupMap = {};
  let cur = '';
  rows.forEach(r => {
    if (r['學制']) cur = r['學制'];
    if (!groupMap[cur]) { groupMap[cur]=[]; groupOrder.push(cur); }
    groupMap[cur].push(r);
  });
  let html = '<h2 class="section-heading">教育目標與核心能力</h2>';
  [...new Set(groupOrder)].forEach(g => {
    html += `<div class="comp-group"><h3>${esc(g)}</h3>`;
    groupMap[g].forEach(r => {
      html += `<div class="comp-row"><div class="comp-zh">${esc(r['核心能力（中）']||'')}</div><div class="comp-en">${esc(r['核心能力（英）']||'')}</div><div class="comp-desc">${esc(r['說明']||'')}</div></div>`;
    });
    html += '</div>';
  });
  area.innerHTML = html;
}

async function renderRegulations(area, sheets, sheetName, page) {
  const data = sheetDataRows(sheets[sheetName] || [], ['分類','名稱']);
  const heading = (page && page.cfg && page.cfg['頁首大標題']) || '系所規章';
  const catOrder = [], catMap = {};
  data.rows.forEach(r => {
    const cat = String(data.at(r,'分類')).trim() || '其他';
    if (!catMap[cat]) { catMap[cat]=[]; catOrder.push(cat); }
    catMap[cat].push(r);
  });
  /* 含 goTo 公告路徑的列以公告頁面式呈現（內容由公告總庫帶出） */
  const newsRef = r => String(data.at(r,'URL / 相對路徑') || data.at(r,'URL')).match(/goTo\('[^']*\/(\d+)'\)/);
  let byId = null;
  if (data.rows.some(newsRef)) {
    area.innerHTML = '<div class="loading">載入公告資料中</div>';
    const db = await getNewsDB();
    byId = {};
    db.forEach(x => { if (x.wpid) byId[x.wpid] = x; });
  }
  let html = `<h2 class="section-heading">${esc(heading)}</h2>`;
  catOrder.forEach(cat => {
    html += `<div class="reg-wrap"><div class="reg-cat-title">${esc(cat)}</div>`;
    let list = '';
    catMap[cat].forEach(r => {
      const name = String(data.at(r,'名稱')).trim();
      const url  = String(data.at(r,'URL / 相對路徑') || data.at(r,'URL')).trim();
      const type = String(data.at(r,'檔案類型') || data.at(r,'類型')).trim();
      const sub  = String(data.at(r,'子分類')).trim();
      if (!name) return;
      const m = byId && newsRef(r);
      const item = m ? byId[m[1]] : null;
      if (item) {
        if (list) { html += `<ul class="reg-list">${list}</ul>`; list = ''; }
        html += newsItemHTML(item, false);
        return;
      }
      list += `<li class="reg-item">${fileBadge(type)}`;
      if (sub) list += `<span class="b-sub">${esc(sub)}</span>`;
      list += url ? `<a ${cardLinkAttrs(url)}>${esc(name)}</a>` : `<span class="name">${esc(name)}</span>`;
      list += '</li>';
    });
    if (list) html += `<ul class="reg-list">${list}</ul>`;
    html += '</div>';
  });
  area.innerHTML = html;
}

function renderActivities(area, sheets, sheetName) {
  const rows = sheets[sheetName] || [];
  const catOrder = [], catMap = {};
  rows.forEach(r => {
    const cat = r['分類']||'活動';
    if (!catMap[cat]) { catMap[cat]=[]; catOrder.push(cat); }
    catMap[cat].push(r);
  });
  let html = '<h2 class="section-heading">系所活動</h2>';
  catOrder.forEach(cat => {
    html += `<div class="reg-wrap"><div class="reg-cat-title">${esc(cat)}</div><ul class="reg-list">`;
    catMap[cat].forEach(r => {
      const title = r['標題']||'', url = r['URL']||'', type = r['類型']||'';
      html += `<li class="reg-item">${fileBadge(type)}`;
      html += url ? `<a href="${esc(url)}" target="_blank" rel="noopener noreferrer">${esc(title)}</a>` : `<span class="name">${esc(title)}</span>`;
      html += '</li>';
    });
    html += '</ul></div>';
  });
  area.innerHTML = html;
}

function renderResearch(area, sheets, sheetName) {
  researchRows = sheets[sheetName] || [];
  let html = `<h2 class="section-heading">研究成果 <span class="section-count">${researchRows.length} 項</span></h2><div class="research-grid">`;
  researchRows.forEach((r, idx) => {
    const photo = r['照片URL']||'';
    html += `<div class="research-card" role="button" tabindex="0" aria-label="${esc(r['標題']||'研究成果')} 詳細資料" data-act="openResearch" data-idx="${idx}">`;
    if (photo) {
      html += `<img class="rc-img" loading="lazy" decoding="async" src="${esc(photo)}" alt="${esc(r['標題']||'')}" data-onerr="hideflex">`;
      html += '<div class="rc-no-img" style="display:none">🔬</div>';
    } else { html += '<div class="rc-no-img">🔬</div>'; }
    const authors = r['作者']||'';
    html += `<div class="rc-body"><div class="rc-journal">${esc(r['期刊']||'')}</div><div class="rc-title">${esc(r['標題']||'')}</div><div class="rc-authors">${esc(authors.length>80?authors.slice(0,80)+'…':authors)}</div></div></div>`;
  });
  area.innerHTML = html + '</div>';
}

function openResearchModal(idx) {
  const r = researchRows[idx]; if (!r) return;
  const photo = r['照片URL']||'';
  document.getElementById('modal-body').innerHTML = `
<div class="rmodal-head">
  <div class="rj">${esc(r['期刊']||'')}</div>
  <h2>${esc(r['標題']||'')}</h2>
  <div class="ra">${esc(r['作者']||'')}</div>
</div>
<div class="rmodal-body">
  ${photo?`<img src="${esc(photo)}" alt="${esc(r['標題']||'')}" data-onerr="hide"><p class="rmodal-caption">${esc(r['圖說明']||'')}</p>`:''}
  ${r['研究說明(中文)']?`<div class="rmodal-desc"><strong>摘要（中）</strong><br>${esc(r['研究說明(中文)'])}</div>`:''}
  ${r['研究說明(英文)']?`<div class="rmodal-desc"><strong>Abstract</strong><br>${esc(r['研究說明(英文)'])}</div>`:''}
  ${r['題目']?`<div class="rmodal-desc" style="font-style:italic">${esc(r['題目'])}</div>`:''}
  ${r['網頁連結']?`<a class="rmodal-link" href="${esc(r['網頁連結'])}" target="_blank" rel="noopener noreferrer">🔗 查看完整內容</a>`:''}
</div>`;
  document.getElementById('modal-overlay').classList.add('open');
}

function closeModal() { document.getElementById('modal-overlay').classList.remove('open'); }
document.getElementById('modal-overlay').addEventListener('click', e => {
  if (e.target === document.getElementById('modal-overlay')) closeModal();
});

function renderLab(area, sheets, sheetName) {
  const rows = sheets[sheetName]||[];
  let html = `<h2 class="section-heading">實驗室 <span class="section-count">${rows.length} 間</span></h2><div class="table-wrap"><table class="data-table"><thead><tr><th>中文名稱</th><th>英文名稱</th><th>實驗室主持人</th><th>實驗室網頁</th></tr></thead><tbody>`;
  rows.forEach(r => {
    const web = r['實驗室網頁']||'';
    html += `<tr><td>${esc(r['中文名稱']||'')}</td><td>${esc(r['英文名稱']||'')}</td><td>${esc(r['實驗室主持人']||'')}</td><td>${web?`<a href="${esc(web)}" target="_blank" rel="noopener noreferrer">🔗 連結</a>`:''}</td></tr>`;
  });
  area.innerHTML = html + '</tbody></table></div>';
}

function renderAwards(area, sheets, sheetName) {
  const rows = [...(sheets[sheetName]||[])].sort((a,b)=>(b['獲獎年度']||0)-(a['獲獎年度']||0));
  let html = `<h2 class="section-heading">教師獲獎紀錄 <span class="section-count">${rows.length} 筆</span></h2><div class="table-wrap"><table class="data-table"><thead><tr><th>獲獎年度</th><th>主辦單位</th><th>榮譽名稱</th><th>獲獎人</th></tr></thead><tbody>`;
  rows.forEach(r => { html += `<tr><td>${esc(String(r['獲獎年度']||''))}</td><td>${esc(r['主辦單位']||'')}</td><td>${esc(r['榮譽名稱']||'')}</td><td>${esc(r['獲獎人']||'')}</td></tr>`; });
  area.innerHTML = html + '</tbody></table></div>';
}

function renderVisitors(area, sheets, sheetName) {
  const rows = sheets[sheetName]||[];
  let html = `<h2 class="section-heading">國外學者來訪 <span class="section-count">${rows.length} 筆</span></h2><div class="table-wrap"><table class="data-table visitor-table"><colgroup><col style="width:20%"><col style="width:10%"><col style="width:10%"><col style="width:34%"><col style="width:11%"><col style="width:15%"></colgroup><thead><tr><th>來訪者</th><th>來訪國家</th><th>職稱</th><th>來訪單位</th><th>邀請人</th><th>日期</th></tr></thead><tbody>`;
  rows.forEach(r => { html += `<tr><td>${esc(r['來訪者']||'')}</td><td>${esc(r['來訪國家']||'')}</td><td>${esc(r['來訪者職稱']||'')}</td><td>${esc(r['來訪單位']||'')}</td><td>${esc(r['邀請人']||'')}</td><td>${esc(r['日期']||'')}</td></tr>`; });
  area.innerHTML = html + '</tbody></table></div>';
}

/* ════════════════════════════════════════════════════
   STUDENT RENDERERS
════════════════════════════════════════════════════ */

/* ── Unified tab system ── */
function initTabs(container) {
  container.querySelectorAll('.tab-btn').forEach(btn => {
    btn.onclick = () => {
      const group   = btn.dataset.group;
      const panelId = btn.dataset.panel;
      const style   = btn.dataset.style || 'tab-navy';
      container.querySelectorAll(`.tab-btn[data-group="${group}"]`).forEach(b => b.classList.remove('tab-navy','tab-blue'));
      btn.classList.add(style);
      container.querySelectorAll(`.tab-panel[data-group="${group}"]`).forEach(p => p.classList.remove('active'));
      const panel = container.querySelector('#' + panelId);
      if (panel) panel.classList.add('active');
    };
  });
}

function renderCourseUG(area, sheets, sheetName) {
  const catColors = {
    '理論物理':{ bg:'#2563a820', border:'#2563a8', text:'#2563a8' },
    '實驗物理':{ bg:'#c8960c20', border:'#c8960c', text:'#c8960c' },
    '數學課程':{ bg:'#2a7a4b20', border:'#2a7a4b', text:'#2a7a4b' },
  };
  const gradeOrder = [], gradeMap = {};
  let curGrade = '';
  (sheets[sheetName]||[]).forEach(r => {
    const g = String(r['年級']||'').replace('▌ ','').trim();
    if (!g) return;
    if (g !== curGrade) { curGrade = g; if (!gradeMap[g]) { gradeMap[g]=[]; gradeOrder.push(g); } }
    if (r['科目名']) gradeMap[curGrade].push(r);
  });
  const grades = [...new Set(gradeOrder)].filter(g => (gradeMap[g]||[]).length > 0);

  let html = '<h2 class="section-heading">大學部課程</h2><div class="tab-bar">';
  grades.forEach((g,i) => {
    const lbl = (g.includes('三')||g.includes('四')) ? '三、四年級' : g;
    html += `<button class="tab-btn${i===0?' tab-navy':''}" data-group="ug" data-panel="ug-${i}">${esc(lbl)}</button>`;
  });
  html += '</div>';
  grades.forEach((g,i) => {
    html += `<div id="ug-${i}" class="tab-panel${i===0?' active':''}" data-group="ug"><div class="table-wrap"><table class="data-table ug-course-table"><colgroup><col class="ug-c-cat"><col class="ug-c-name"><col class="ug-c-course"><col class="ug-c-code"><col class="ug-c-term"></colgroup><thead><tr><th>分類</th><th>課程名稱</th><th>Course</th><th>課號</th><th>全/半年</th></tr></thead><tbody>`;
      (gradeMap[g]||[]).forEach(r => {
      const cat = r['分類']||'', cc = catColors[cat];
      const catBadge = cc ? `<span class="cat-inline" style="background:${cc.bg};border:1px solid ${cc.border};color:${cc.text}">${esc(cat)}</span>` : '';
      html += `<tr><td class="ug-c-cat">${catBadge}</td><td>${esc(r['科目名']||'')}</td><td>${esc(r['Course']||'')}</td><td><code>${esc(r['課號']||'')}</code></td><td class="ug-c-term">${esc(r['全/半年']||'')}</td></tr>`;
    });
    html += '</tbody></table></div>';
    html += '</div>';
  });
  html += buildEmiSection(sheets['大學部EMI課程地圖'], '大學部 EMI 課程地圖');   // EMI 課程地圖移到全頁最下方（不再夾在一年級分頁內）
  area.innerHTML = html;
  initTabs(area);
}

function renderCoursePG(area, sheets, sheetName) {
  const rows = sheets[sheetName]||[];
  let html = '<h2 class="section-heading">研究所課程</h2><div class="table-wrap"><table class="data-table nowrap-table"><thead><tr><th>課程名稱</th><th>Course</th><th>課號</th><th>必/選</th><th>學分</th></tr></thead><tbody>';
  rows.forEach(r => {
    const req = String(r['必/選']||''), bc = req.includes('必')?'req-must':'req-elec';
    html += `<tr><td>${esc(r['科目名']||'')}</td><td>${esc(r['Course']||'')}</td><td><code>${esc(r['課號']||'')}</code></td><td><span class="req-badge ${bc}">${esc(req)}</span></td><td>${esc(String(r['學分數']||''))}</td></tr>`;
  });
  html += '</tbody></table></div>';
  html += buildEmiSection(sheets['研究所EMI課程地圖'], '研究所 EMI 課程地圖');
  area.innerHTML = html;
}

function buildEmiSection(rows, title) {
  if (!rows || !rows.length) return '';
  const header = Object.values(rows[0]);
  let noteText = '';
  const data = rows.slice(1).filter(r => {
    const vals = Object.values(r);
    if (!vals.some(v => v !== '')) return false;
    /* 整列只有第一格有內容 → 視為說明文字，移出表格 */
    const first = String(vals[0]||'').trim();
    if (first && vals.slice(1).every(v => String(v ?? '').trim() === '')) {
      noteText += (noteText ? '<br>' : '') + esc(first);
      return false;
    }
    return true;
  });
  let html = `<div class="emi-section-block"><div class="emi-section-title">${esc(title)} <span class="emi-section-subtitle">English as Medium of Instruction</span></div>`;
  html += '<div class="table-wrap"><table class="data-table" style="table-layout:fixed;width:100%"><thead><tr>';
  header.forEach(h => { html += `<th>${esc(String(h||'').split('\n')[0])}</th>`; });
  html += '</tr></thead><tbody>';
  /* 必/選修欄依表頭偵測（大學部地圖有此欄，研究所無） */
  const reqIdx = header.findIndex(h => String(h||'').startsWith('必/選'));
  data.forEach(r => {
    const vals = Object.values(r);
    html += '<tr>';
    vals.forEach((v,vi) => {
      const s = String(v||'');
      if (vi===reqIdx && s) html += `<td><span class="req-badge ${s.includes('必')?'req-must':'req-elec'}">${esc(s)}</span></td>`;
      else if (vi===4||vi===5) html += `<td style="text-align:center">${esc(s)}</td>`;
      else html += `<td>${esc(s)}</td>`;
    });
    html += '</tr>';
  });
  html += `</tbody></table></div><p class="emi-note">${noteText || '說明：「x」表示該學期不開課　｜　上/下學期欄數字為學分數'}</p></div>`;
  return html;
}

function renderAnnualMeeting(area, sheets) {
  const meetInfo  = sheets['物理小年會']  || [];
  const albumDefs = sheets['小年會相簿']  || [];
  const histRows  = sheets['歷年小年會']  || [];
  const descRow   = meetInfo.find(r => String(r['欄位']||'').includes('活動說明'));
  const desc      = descRow ? String(descRow['內容/路徑']||'') : '';

  let html = `<div class="meeting-desc">${esc(desc).replace(/\\n/g,'<br>')}</div>`;

  if (albumDefs.length) {
    html += '<div class="section-heading mt-24">活動相簿</div><div class="tab-bar">';
    albumDefs.forEach((r,i) => {
      html += `<button class="tab-btn${i===0?' tab-navy':''}" data-group="alb" data-panel="alb-${i}">${esc(r['分類']||'')}</button>`;
    });
    html += '</div>';
    albumDefs.forEach((r,i) => {
      html += `<div id="alb-${i}" class="tab-panel${i===0?' active':''}" data-group="alb">`;
      html += buildCarousel(`alb${i}`, r['圖片跑馬燈路徑']||'', 28);
      html += '</div>';
    });
  }

  html += '<div class="section-heading mt-24">歷年小年會</div>';
  const yearOrder = [], yearMap = {};
  histRows.forEach(r => {
    const yr = String(r['年份']||''), cls = r['實驗班']||'';
    if (!yr) return;
    if (!yearMap[yr]) { yearMap[yr]={}; yearOrder.push(yr); }
    if (!yearMap[yr][cls]) yearMap[yr][cls] = [];
    yearMap[yr][cls].push(r);
  });
  const years = [...new Set(yearOrder)];
  html += '<div class="tab-bar">';
  years.forEach((y,i) => { html += `<button class="tab-btn${i===0?' tab-navy':''}" data-group="yr" data-panel="yr-${y}">${esc(y)}</button>`; });
  html += '</div>';
  years.forEach((y,i) => {
    const classes = Object.keys(yearMap[y]);
    html += `<div id="yr-${y}" class="tab-panel${i===0?' active':''}" data-group="yr">`;
    if (classes.length > 1) {
      html += '<div class="tab-bar">';
      classes.forEach((c,ci) => { html += `<button class="tab-btn${ci===0?' tab-blue':''}" data-group="cls-${y}" data-panel="cls-${y}-${ci}" data-style="tab-blue">${esc(c)}</button>`; });
      html += '</div>';
    }
    classes.forEach((c,ci) => {
      html += `<div id="cls-${y}-${ci}" class="tab-panel${ci===0?' active':''}" data-group="cls-${y}">`;
      // 合照路徑：優先用 Excel「照片URL」欄（取該實驗班任一列有填者）；未填則退回自動路徑
      const rowsOfClass = yearMap[y][c];
      const customRow = rowsOfClass.find(r => String(r['照片URL']||'').trim());
      const custom = customRow ? String(customRow['照片URL']).trim() : '';
      const photoSrc  = custom || `files/student/annualmeeting/previousmeeting/${y}/${c}.webp`;
      const photoBase = custom ? custom.replace(/\.[^.\/]+$/,'') : `files/student/annualmeeting/previousmeeting/${y}/${c}`;
      html += '<div class="ameet-row">';
      html += `<div class="ameet-photo"><img src="${esc(photoSrc)}" data-base="${esc(photoBase)}" loading="lazy" alt="${esc(y)} ${esc(c)}" data-onerr="ameet"></div>`;
      html += '<table class="data-table ameet-table"><thead><tr><th>組員</th><th>專題題目</th></tr></thead><tbody>';
      yearMap[y][c].forEach(r => {
        const title = r['專題題目']||'', url = r['海報URL']||'';
        html += `<tr><td>${esc(r['組員']||'')}</td><td>${url?`<a href="${esc(url)}" target="_blank" rel="noopener noreferrer">${esc(title)}</a>`:esc(title)}</td></tr>`;
      });
      html += '</tbody></table>';
      html += '</div></div>';
    });
    html += '</div>';
  });
  area.innerHTML = html;
  initTabs(area);
}

function buildCarousel(id, folder, maxCount) {
  if (!folder) return '<p style="color:var(--muted);font-size:.88rem;padding:12px">（圖片資料夾未設定）</p>';
  let slides = '', dots = '';
  for (let i = 1; i <= maxCount; i++) {
    const active = i===1?' cs-active':'';
    const base = `${esc(folder)}/${i}`;
    slides += `<div class="cs-slide${active}"><img src="${base}.webp" data-base="${base}" data-fallback="1" loading="lazy" alt="" data-onerr="cs"></div>`;
    dots   += `<span class="cs-dot${active}" data-idx="${i-1}"></span>`;
  }
  return `<div class="carousel" id="cs-${esc(id)}"><div class="cs-track">${slides}</div><button class="cs-btn cs-prev" data-id="${esc(id)}" data-dir="-1">&#8249;</button><button class="cs-btn cs-next" data-id="${esc(id)}" data-dir="1">&#8250;</button><div class="cs-dots" data-id="${esc(id)}">${dots}</div></div>`;
}

document.getElementById('content').addEventListener('click', e => {
  const btn = e.target.closest('.cs-btn');
  const dot = e.target.closest('.cs-dot');
  if (btn) {
    const id = btn.dataset.id, dir = parseInt(btn.dataset.dir);
    const c  = document.getElementById('cs-' + id); if (!c) return;
    const slides = c.querySelectorAll('.cs-slide');
    let cur = [...slides].findIndex(s => s.classList.contains('cs-active'));
    goCarousel(c, (cur + dir + slides.length) % slides.length);
  }
  if (dot) {
    const wrap = dot.closest('.cs-dots');
    const c = wrap ? document.getElementById('cs-' + wrap.dataset.id) : null;
    if (c) goCarousel(c, parseInt(dot.dataset.idx));
  }
});

function goCarousel(container, idx) {
  container.querySelectorAll('.cs-slide').forEach((s,i) => s.classList.toggle('cs-active', i===idx));
  container.querySelectorAll('.cs-dot').forEach((d,i)   => d.classList.toggle('active',    i===idx));
}

/* 圖片載入失敗 → 整張幻燈片連同一顆圓點移除，輪播只剩實際存在的照片 */
function csImgFail(img) {
  // 先嘗試其他副檔名（.webp → .jpg → .png → .jpeg）再放棄
  const base = img.dataset.base;
  if (base) {
    const tried = (img.dataset.tried || 'webp').split(',');
    const order = ['webp','jpg','png','jpeg'];
    const next = order.find(ext => !tried.includes(ext));
    if (next) {
      img.dataset.tried = tried.concat(next).join(',');
      img.src = `${base}.${next}`;
      return;   // 換副檔名再試一次，先不剔除
    }
  }
  const slide    = img.closest('.cs-slide');
  const carousel = img.closest('.carousel');
  if (!slide || !carousel) return;
  const wasActive = slide.classList.contains('cs-active');
  slide.remove();
  const dots = carousel.querySelectorAll('.cs-dot');
  if (dots.length) dots[dots.length - 1].remove();
  carousel.querySelectorAll('.cs-dot').forEach((d,i) => d.dataset.idx = i);
  const remaining = carousel.querySelectorAll('.cs-slide');
  if (!remaining.length) {
    carousel.outerHTML = '<p style="color:var(--muted);font-size:.88rem;padding:12px">（此分類暫無照片）</p>';
    return;
  }
  if (wasActive) goCarousel(carousel, 0);
}

/* 小年會合照載入失敗：先試其他副檔名，全失敗則隱藏整個照片區（表格仍正常顯示） */
function ameetImgFail(img) {
  const base = img.dataset.base;
  const tried = (img.dataset.tried || 'webp').split(',');
  const order = ['webp','jpg','png','jpeg'];
  const next = order.find(ext => !tried.includes(ext));
  if (base && next) {
    img.dataset.tried = tried.concat(next).join(',');
    img.src = `${base}.${next}`;
    return;
  }
  const box = img.closest('.ameet-photo');
  if (box) box.remove();
}

function renderScholarship(area, sheets, sheetName) {
  const rows = sheets[sheetName]||[];
  let html = '<h2 class="section-heading">獎學金</h2><div class="link-cards">';
  rows.forEach(r => {
    html += `<a href="${esc(r['URL']||'#')}" target="_blank" rel="noopener noreferrer" class="link-card"><div class="lc-title">${esc(r['連結標籤']||'')}</div><div class="lc-desc">${esc(r['說明']||'')}</div></a>`;
  });
  html += '</div>';
  /* 獎學金公告：自公告總庫匯入（分類=獎學金），與下方「所有公告」同源 */
  html += `<h2 class="section-heading mt-32">獎學金公告 <span class="section-count" id="schol-news-count"></span></h2>
  <div id="schol-news"><div class="loading">載入獎學金公告中</div></div>`;
  area.innerHTML = html;
  loadScholNews();
}

/* 獎學金公告區塊：獨立於主公告頁的 newsState，避免互相干擾 */
const SCHOL_PAGE_SIZE = 15;
let scholNewsState = { all: [], shown: SCHOL_PAGE_SIZE };

async function loadScholNews() {
  const box = document.getElementById('schol-news');
  if (!box) return;
  try { if (!DATA['news'] && LOADP['news']) await LOADP['news']; } catch (_) {}
  if (!document.getElementById('schol-news')) return;   /* 使用者已切換頁面 */
  const newsSheets = DATA['news'];
  const all = (newsSheets && newsSheets['公告總庫']) ? parseNewsSheet(newsSheets['公告總庫']) : [];
  const schol = all.filter(r => r.cat === '獎學金' || r.subs.includes('獎學金'));
  schol.sort((a, b) => (newsToDate(b.date) || 0) - (newsToDate(a.date) || 0));
  scholNewsState = { all: schol, shown: SCHOL_PAGE_SIZE };
  renderScholNews();
}

function renderScholNews() {
  const box = document.getElementById('schol-news');
  if (!box) return;
  const { all, shown } = scholNewsState;
  const cntEl = document.getElementById('schol-news-count');
  if (cntEl) cntEl.textContent = all.length + ' 則';
  if (!all.length) { box.innerHTML = '<div class="error-msg">目前沒有獎學金公告</div>'; return; }
  const slice = all.slice(0, shown);
  let html = slice.map(r => newsItemHTML(r, false)).join('');
  if (all.length > shown) {
    html += `<button class="news-more-btn" data-act="scholNewsMore">載入更多（已顯示 ${slice.length} / ${all.length} 則）</button>`;
  } else {
    html += `<div class="news-count-info">已顯示全部 ${all.length} 則獎學金公告</div>`;
  }
  box.innerHTML = html;
}

function scholNewsMore() {
  scholNewsState.shown += SCHOL_PAGE_SIZE;
  renderScholNews();
}

function renderSpaces(area, sheets, sheetName) {
  const raw  = sheets[sheetName]||[];
  const rows = raw.slice(1).filter(r => Object.values(r)[0] && String(Object.values(r)[0]).trim());
  let html = '<h2 class="section-heading">學生活動空間</h2><div class="spaces-grid">';
  rows.forEach((r, idx) => {
    const v = Object.values(r);
    const name=String(v[0]||'').replace(/\n/g,' '), loc=String(v[1]||''), hours=String(v[2]||'');
    const desc=String(v[3]||''), folder=String(v[4]||'');
    html += `<div class="space-card"><div class="sp-header"><div class="sp-name">${esc(name)}</div><div class="sp-meta"><span class="sp-loc">📍 ${esc(loc)}</span><span class="sp-hours">🕐 ${esc(hours)}</span></div></div><div class="sp-desc">${esc(desc).replace(/\n/g,'<br>')}</div>`;
    if (folder) html += buildCarousel('sp'+idx, folder, 17);
    html += '</div>';
  });
  area.innerHTML = html + '</div>';
}

function renderStudentAwards(area, sheets, sheetName) {
  const rows = (sheets[sheetName]||[]).filter(r => r['姓名'] && !String(r['姓名']).startsWith('※'));
  let html = `<h2 class="section-heading">學生獲獎紀錄 <span class="section-count">${rows.length} 筆</span></h2><div class="table-wrap"><table class="data-table nowrap-table"><thead><tr><th>姓名</th><th>身份</th><th>指導教授</th><th>得獎日期</th><th>得獎名稱</th></tr></thead><tbody>`;
  rows.forEach(r => { html += `<tr><td>${esc(r['姓名']||'')}</td><td>${esc(r['職稱']||'')}</td><td>${esc(r['指導教授']||'')}</td><td>${esc(r['得獎日期']||'')}</td><td class="wrap-col">${esc(r['得獎名稱']||'')}</td></tr>`; });
  area.innerHTML = html + '</tbody></table></div>';
}

function renderTheses(area, sheets, sheetName) {
  const rows = (sheets[sheetName]||[]).filter(r => r['學生']);
  let html = `<h2 class="section-heading">研究生論文 <span class="section-count">${rows.length} 筆</span></h2><div class="table-wrap"><table class="data-table" style="table-layout:fixed;width:100%"><colgroup><col style="width:6em"><col style="width:7em"><col style="width:5.5em"><col style="width:5em"><col></colgroup><thead><tr><th>學生</th><th>指導老師</th><th style="white-space:nowrap">學位</th><th>學年度</th><th>論文題目</th></tr></thead><tbody>`;
  rows.forEach(r => {
    const deg = String(r['碩/博士']||''), bc = deg.includes('博')?'badge-phd':'badge-ms';
    html += `<tr><td>${esc(r['學生']||'')}</td><td>${esc(r['指導老師']||'')}</td><td><span class="deg-badge ${bc}">${esc(deg)}</span></td><td>${esc(String(r['學年度']||'').replace('.0',''))}</td><td>${esc(r['論文題目']||'')}</td></tr>`;
  });
  area.innerHTML = html + '</tbody></table></div>';
}

function renderStudentTravel(area, sheets, sheetName) {
  const rows = sheets[sheetName]||[];
  let html = `<h2 class="section-heading">學生國外差旅 <span class="section-count">${rows.length} 筆</span></h2><div class="table-wrap"><table class="data-table nowrap-table"><thead><tr><th>學生</th><th>教師</th><th>身份</th><th>出國期間</th><th>國家</th><th>經費來源</th><th>出國目的</th></tr></thead><tbody>`;
  rows.forEach(r => { html += `<tr><td>${esc(r['學生']||'')}</td><td>${esc(r['教師']||'')}</td><td>${esc(r['身份']||'')}</td><td>${esc(r['出國期間']||'')}</td><td>${esc(r['國家']||'')}</td><td>${esc(r['經費來源']||'')}</td><td class="purpose-cell wrap-col">${esc(r['出國目的']||'')}</td></tr>`; });
  area.innerHTML = html + '</tbody></table></div>';
}

function renderPrograms(area, sheets, sheetName) {
  const rows = sheets[sheetName]||[];
  let html = '<h2 class="section-heading">相關學程</h2><div class="program-list">';
  rows.forEach(r => {
    html += `<div class="prog-row"><div class="prog-name">${esc(r['學程名稱']||'')}</div><div class="prog-desc">${esc(r['說明']||'')}</div>${r['URL']?`<a class="prog-link" href="${esc(r['URL'])}" target="_blank" rel="noopener noreferrer">查看網頁 →</a>`:''}</div>`;
  });
  area.innerHTML = html + '</div>';
}

function renderJournals(area, sheets, sheetName) {
  const rows = sheets[sheetName]||[];
  let html = '<h2 class="section-heading">物理系刊</h2><div class="journal-grid">';
  rows.forEach(r => {
    const pdf=r['PDF路徑或URL']||'', cover=r['封面路徑']||'', note=r['備註']||'';
    html += `<div class="journal-card"><div class="jc-cover">${cover?`<img src="${esc(cover)}" alt="${esc(r['期次/標題']||'')}" loading="lazy" data-onerr="hide">`:' 📖'}</div><div class="jc-body"><div class="jc-issue">${esc(r['期次/標題']||'')}</div><div class="jc-year">${esc(String(r['出版年份']||''))}</div><div class="jc-links">${pdf?`<a class="jc-btn" href="${esc(pdf)}" target="_blank" rel="noopener noreferrer">${pdf.startsWith('http')?'🔗 線上閱讀':'📄 PDF'}</a>`:''}${note&&note.startsWith('http')?`<a class="jc-btn jc-yt" href="${esc(note)}" target="_blank" rel="noopener noreferrer">▶ YouTube</a>`:''}</div></div></div>`;
  });
  area.innerHTML = html + '</div>';
}

/* ── 所有公告 ── */
const NEWS_PAGE_SIZE = 30;
let newsState = { all:[], cat:'全部', year:'', shown:NEWS_PAGE_SIZE };

/* 公告總庫解析（欄位定位＋顯示過濾），renderNews 與 newslist 共用 */
function parseNewsSheet(ws) {
  if (!ws || !ws.length) return [];
  const keys = Object.keys(ws[0] || {}).filter(k => k !== '_lastMod');
  let headerIdx = -1;
  const colMap = {};
  for (let i = 0; i < Math.min(ws.length, 8); i++) {
    const vals = keys.map(k => String(ws[i][k] ?? '').trim());
    if (vals.includes('標題') && vals.includes('分類')) {
      headerIdx = i;
      keys.forEach(k => {
        const name = String(ws[i][k] ?? '').trim();
        if (name) colMap[name] = k;
      });
      break;
    }
  }
  if (headerIdx === -1) keys.forEach(k => { colMap[k] = k; });
  const dataRows = headerIdx === -1 ? ws : ws.slice(headerIdx + 1);
  const at = (r, name) => r[colMap[name]] ?? '';

  return dataRows.filter(r => {
    const title = String(at(r,'標題')).trim();
    if (!title || title === '標題') return false;
    return isOn(at(r,'顯示'));
  }).map(r => ({
    wpid:  String(at(r,'WP-ID')).trim(),
    cat:   String(at(r,'分類')).trim() || '其他',
    subs:  String(at(r,'次分類')).split('、').map(s=>s.trim()).filter(Boolean),
    date:  at(r,'刊登日期'),
    title: String(at(r,'標題')).trim(),
    body:  String(at(r,'內文(純文字)')).trim(),
    links: [
      { name:String(at(r,'連結1名稱')).trim(), url:String(at(r,'連結1網址')).trim() },
      { name:String(at(r,'連結2名稱')).trim(), url:String(at(r,'連結2網址')).trim() },
      { name:String(at(r,'連結3名稱')).trim(), url:String(at(r,'連結3網址')).trim() },
    ].filter(l => l.url),
    src:   String(at(r,'原始連結')).trim(),
    pub:   String(at(r,'發布者')).trim(),
  }));
}

/* 單則公告的條列卡片 HTML（展開式），newsRenderList 與 newslist 共用 */
function newsItemHTML(r, showBadge) {
  return `<div class="news-item">
  <div class="news-item-row" role="button" tabindex="0" aria-expanded="false" data-act="newsToggle">
    ${showBadge ? `<span class="news-badge">${esc(r.cat)}</span>` : ''}
    <div class="news-item-title">${esc(r.title)}</div>
    <div class="news-item-date">${esc(newsFmtDate(r.date))}</div>
  </div>
  <div class="news-item-body">
    <div class="news-meta-line">🗓 ${esc(newsFmtDate(r.date))}${r.pub ? '　📋 發布者：' + esc(r.pub) : ''}${r.subs.length ? '　🏷 次分類：' + esc(r.subs.join('、')) : ''}</div>
    ${r.body ? `<div class="news-body-text">${esc(r.body)}</div>` : '<div class="news-body-text" style="color:var(--muted)">（無內文）</div>'}
    ${r.links.length ? `<div class="news-links-block">${r.links.map(l => `<a href="${esc(l.url)}" target="_blank" rel="noopener">🔗 ${esc(l.name || l.url)}</a>`).join('')}</div>` : ''}
    ${r.src ? `<div class="news-src-link"><a href="${esc(r.src)}" target="_blank" rel="noopener">查看原始公告 →</a></div>` : ''}
  </div>
</div>`;
}

function renderNews(area, sheets, sheetName) {
  const ws = sheets[sheetName];
  if (!ws || !ws.length) { area.innerHTML = '<div class="error-msg">找不到公告資料工作頁</div>'; return; }

  const rows = parseNewsSheet(ws);

  const catOrder = [];
  rows.forEach(r => { if (!catOrder.includes(r.cat)) catOrder.push(r.cat); });
  const years = [...new Set(rows.map(r => newsYear(r.date)).filter(Boolean))].sort((a,b) => b-a);

  newsState = { all:rows, cat:'全部', year:'', shown:NEWS_PAGE_SIZE };

  area.innerHTML = `
<h2 class="section-heading">所有公告 <span class="section-count" id="news-count">${rows.length} 則</span></h2>
<div class="news-toolbar">
  <select class="news-year-sel" id="news-year" aria-label="依年份篩選公告"><option value="">全部年份</option>${years.map(y=>`<option value="${y}">${y}</option>`).join('')}</select>
</div>
<div class="news-tabs" id="news-tabs"></div>
<div id="news-list"></div>`;

  const tabsEl = document.getElementById('news-tabs');
  ['全部', ...catOrder].forEach(cat => {
    const count = cat === '全部' ? rows.length : rows.filter(r => r.cat === cat || r.subs.includes(cat)).length;
    const btn = document.createElement('button');
    btn.className = 'news-tab' + (cat === '全部' ? ' active' : '');
    btn.innerHTML = `${esc(cat)}<span class="n">${count}</span>`;
    btn.onclick = () => {
      newsState.cat = cat;
      newsState.shown = NEWS_PAGE_SIZE;
      tabsEl.querySelectorAll('.news-tab').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      newsRenderList();
    };
    tabsEl.appendChild(btn);
  });

  document.getElementById('news-year').addEventListener('change', e => {
    newsState.year = e.target.value;
    newsState.shown = NEWS_PAGE_SIZE;
    newsRenderList();
  });

  /* goTo 第三段參數：數字 = WP-ID 直達該則公告；文字 = 切換分類頁籤 */
  if (newsPendingParam) {
    const param = newsPendingParam;
    newsPendingParam = '';
    if (/^\d+$/.test(param)) {
      /* WP-ID 模式：搜尋該則公告並自動展開 */
      const idx = rows.findIndex(r => r.wpid === param);
      if (idx !== -1) {
        newsState.cat = '全部';
        newsState.shown = Math.max(NEWS_PAGE_SIZE, idx + 1);
        newsRenderList();
        setTimeout(() => {
          const items = document.querySelectorAll('#news-list .news-item');
          const el = items[idx];
          if (el) {
            el.classList.add('open');
            el.scrollIntoView({ behavior:'smooth', block:'center' });
          }
        }, 50);
      } else {
        newsRenderList();
      }
    } else {
      /* 分類模式 */
      const target = [...tabsEl.querySelectorAll('.news-tab')]
        .find(b => b.textContent.replace(/\d+$/,'').trim() === param);
      if (target) target.click();
      else newsRenderList();
    }
  } else {
    newsRenderList();
  }
}

/* Excel 日期序號或字串/Date 皆可解析 */
function newsToDate(d) {
  if (!d && d !== 0) return null;
  if (d instanceof Date) return isNaN(d) ? null : d;
  if (typeof d === 'number') {
    /* Excel serial number (days since 1899-12-30) */
    const dt = new Date(Math.round((d - 25569) * 86400 * 1000));
    return isNaN(dt) ? null : dt;
  }
  const dt = new Date(d);
  return isNaN(dt) ? null : dt;
}

function newsYear(d) {
  const dt = newsToDate(d);
  return dt ? String(dt.getFullYear()) : '';
}

function newsFmtDate(d) {
  const dt = newsToDate(d);
  if (!dt) return String(d ?? '').slice(0,10);
  const p = n => String(n).padStart(2,'0');
  return `${dt.getFullYear()}-${p(dt.getMonth()+1)}-${p(dt.getDate())}`;
}

function newsFiltered() {
  const { all, cat, year } = newsState;
  return all.filter(r => {
    if (cat !== '全部' && r.cat !== cat && !r.subs.includes(cat)) return false;
    if (year && newsYear(r.date) !== year) return false;
    return true;
  });
}

function newsRenderList() {
  const listEl = document.getElementById('news-list');
  if (!listEl) return;
  const filtered = newsFiltered();
  const countEl = document.getElementById('news-count');
  if (countEl) countEl.textContent = filtered.length + ' 則';

  if (!filtered.length) {
    listEl.innerHTML = '<div class="news-empty">查無符合的公告</div>';
    return;
  }

  const slice = filtered.slice(0, newsState.shown);
  let html = '';
  slice.forEach(r => {
    html += newsItemHTML(r, newsState.cat === '全部');
  });

  if (filtered.length > newsState.shown) {
    html += `<button class="news-more-btn" data-act="newsLoadMore">載入更多（已顯示 ${slice.length} / ${filtered.length} 則）</button>`;
  } else {
    html += `<div class="news-count-info">已顯示全部 ${filtered.length} 則公告</div>`;
  }
  listEl.innerHTML = html;
}

function newsToggle(rowEl) {
  const open = rowEl.closest('.news-item').classList.toggle('open');
  rowEl.setAttribute('aria-expanded', open ? 'true' : 'false');
}

function newsLoadMore() {
  newsState.shown += NEWS_PAGE_SIZE;
  newsRenderList();
}

/* ── 共用：橫幅式工作頁定位（前 2 列為標題與說明，依表頭名稱對應欄位） ── */
function sheetDataRows(ws, requiredCols) {
  if (!ws || !ws.length) return { rows: [], at: () => '' };
  const keys = Object.keys(ws[0]).filter(k => k !== '_lastMod');
  let headerIdx = -1;
  const colMap = {};
  for (let i = 0; i < Math.min(ws.length, 8); i++) {
    const vals = keys.map(k => String(ws[i][k] ?? '').trim());
    if (requiredCols.every(c => vals.includes(c))) {
      headerIdx = i;
      keys.forEach(k => { const n = String(ws[i][k] ?? '').trim(); if (n) colMap[n] = k; });
      break;
    }
  }
  if (headerIdx === -1) keys.forEach(k => { colMap[k] = k; });
  const rows = (headerIdx === -1 ? ws : ws.slice(headerIdx + 1))
    .filter(r => Object.values(r).some(v => String(v ?? '').trim()));
  const at = (r, name) => r[colMap[name]] ?? '';
  return { rows, at };
}

/* 嚴格解析 javascript:goTo('...')：只取出字串參數，連反斜線/引號都排除以防跳脫；
   任何其他 javascript: 內容一律視為無效、不執行（白名單） */
function gotoArg(url) {
  const m = /^javascript:\s*goTo\(\s*['"]([^'"\\]*)['"]\s*\)\s*;?\s*$/.exec(String(url || '').trim());
  return m ? m[1] : null;
}
/* 卡片連結屬性：javascript:goTo(...) → 以 data-goto 站內跳頁（事件委派，不執行任意 JS）；其餘開新分頁 */
function cardLinkAttrs(url) {
  const u = String(url || '').trim();
  if (!u) return 'href="#" data-noop="1"';
  if (u.startsWith('javascript:')) {
    const arg = gotoArg(u);
    return arg !== null ? `href="#" data-goto="${esc(arg)}"` : 'href="#" data-noop="1"';
  }
  return `href="${esc(u)}" target="_blank" rel="noopener noreferrer"`;
}
/* ════════════════════════════════════════════════════
   事件委派（取代所有行內 on* 處理器，符合嚴格 CSP script-src 'self'）
   click：data-goto 站內跳頁、data-noop 無作用連結、data-act 各類動作
   error（捕獲階段）：圖片載入失敗的處置（hide / hideflex / 換副檔名）
════════════════════════════════════════════════════ */
document.addEventListener('click', e => {
  const t = e.target.closest && e.target.closest('[data-goto],[data-noop],[data-act]');
  if (!t) return;
  if (t.hasAttribute('data-goto')) { e.preventDefault(); goTo(t.getAttribute('data-goto')); return; }
  if (t.hasAttribute('data-noop')) { e.preventDefault(); return; }
  e.preventDefault();
  switch (t.getAttribute('data-act')) {
    case 'closeModal':    closeModal(); break;
    case 'closeFaculty':  closeFacultyDetail(); break;
    case 'newsToggle':    newsToggle(t); break;
    case 'newsLoadMore':  newsLoadMore(); break;
    case 'scholNewsMore': scholNewsMore(); break;
    case 'openResearch':  openResearchModal(parseInt(t.dataset.idx, 10)); break;
    case 'openFaculty':   openFacultyDetail(t.dataset.sec, parseInt(t.dataset.idx, 10), t.dataset.sheet); break;
    case 'switchFdTab':   switchFdTab(t, t.dataset.tab); break;
  }
});
window.addEventListener('error', e => {
  const el = e.target;
  if (!el || el.tagName !== 'IMG') return;
  switch (el.getAttribute('data-onerr')) {
    case 'hide':     el.style.display = 'none'; break;
    case 'hideflex': el.style.display = 'none';
                     if (el.nextElementSibling) el.nextElementSibling.style.display = 'flex'; break;
    case 'cs':       csImgFail(el); break;
    case 'ameet':    ameetImgFail(el); break;
  }
}, true);  /* 捕獲階段：img error 不冒泡，必須在 capture 攔截 */

/* 無障礙：role=button 的 div（公告列、師資/研究卡）以 Enter/Space 啟動（原生 button/a 不重複處理） */
document.addEventListener('keydown', e => {
  if (e.key !== 'Enter' && e.key !== ' ' && e.key !== 'Spacebar') return;
  const t = e.target;
  if (!t || !t.matches || !t.matches('[data-act],[data-goto],[data-noop]')) return;
  if (/^(BUTTON|A|INPUT|SELECT|TEXTAREA)$/.test(t.tagName)) return;
  e.preventDefault();
  t.click();
});

/* Email 反爬：監看內容區，任何渲染（list-view / detail-view）後自動把信箱組回純文字（不建 mailto） */
(function () {
  const content = document.getElementById('content');
  if (!content) return;
  hydrateEmails(content);
  new MutationObserver(() => hydrateEmails(content))
    .observe(content, { childList: true, subtree: true });
})();


/* ── 高中生專區主頁（portal） ── */
/* 比例長條：依該組最大值等比縮放，最大項填滿、其餘按比例 */
function renderBars(rows, data) {
  const items = rows.map(r => ({
    name: String(data.at(r,'標題')).trim(),
    pctRaw: String(data.at(r,'分類')).trim(),
    note: String(data.at(r,'內容/路徑')).trim(),
    pct: parseFloat(String(data.at(r,'分類')).replace(/[^0-9.]/g,'')) || 0,
  })).filter(it => it.name);
  const max = Math.max(1, ...items.map(it => it.pct));
  let html = '<div class="major-bars">';
  items.forEach(it => {
    const w = Math.round(it.pct / max * 100);
    html += `<div class="major-row">`
      + `<div class="major-name">${esc(it.name)}</div>`
      + `<div class="major-track"><div class="major-fill" style="width:${w}%"></div></div>`
      + `<div class="major-pct">${esc(it.pctRaw)}</div>`
      + `</div>`
      + (it.note ? `<div class="major-note">${esc(it.note)}</div>` : '');
  });
  return html + '</div>';
}

async function renderPortal(area, sheets, sheetName) {
  const data = sheetDataRows(sheets[sheetName], ['區塊','標題']);
  const blocks = { '設定':[], '數據':[], '影片':[], '課程地圖':[], '雙軌':[], '出路數據':[], '升學領域':[], '出路':[], '產業對照':[], '問答':[], '入口':[], '公告':[] };
  data.rows.forEach(r => {
    const b = String(data.at(r,'區塊')).trim();
    if (blocks[b]) blocks[b].push(r);
  });
  const get = name => {
    const row = blocks['設定'].find(r => String(data.at(r,'標題')).trim() === name);
    return row ? String(data.at(row,'內容/路徑')).trim() : '';
  };
  const about = get('關於我們');

  /* 影片清單（標題＋YouTube URL，由 Excel 控制；相容舊「設定/介紹影片」） */
  const ytId = u => { const m = String(u).match(/(?:youtube\.com\/watch\?v=|youtu\.be\/|youtube\.com\/embed\/)([\w-]+)/); return m ? m[1] : ''; };
  let videos = blocks['影片'].map(r => ({ title:String(data.at(r,'標題')).trim(), id:ytId(data.at(r,'內容/路徑')) })).filter(v => v.id);
  if (!videos.length) { const legacy = get('介紹影片'); if (ytId(legacy)) videos = [{title:'介紹影片', id:ytId(legacy)}]; }

  let html = '';

  /* 影片：左側標題列表 + 右側嵌入播放器 */
  if (videos.length) {
    const single = videos.length === 1;
    html += `<div class="vid-block${single?' vid-single':''}">`;
    if (!single) {
      html += '<ul class="vid-list">' + videos.map((v,i) =>
        `<li class="vid-item${i===0?' active':''}" data-vid="${esc(v.id)}" data-i="${i}">${esc(v.title)}</li>`
      ).join('') + '</ul>';
    }
    html += `<div class="vid-player"><div class="portal-video"><iframe id="vid-frame" src="https://www.youtube.com/embed/${esc(videos[0].id)}" title="${esc(videos[0].title)}" allowfullscreen loading="lazy"></iframe></div></div>`;
    html += '</div>';
  }

  /* 數據亮點卡（為什麼選中央物理）：標題=大數字、分類=標籤、內容/路徑=說明 */
  if (blocks['數據'].length) {
    html += '<h2 class="section-heading">為什麼選中央物理</h2>';
    html += '<div class="stat-grid">';
    blocks['數據'].forEach(r => {
      const num   = String(data.at(r,'標題')).trim();        // 大數字 / 主標
      const label = String(data.at(r,'分類')).trim();        // 標籤
      const note  = String(data.at(r,'內容/路徑')).trim();   // 一句說明
      if (!num && !label) return;
      html += `<div class="stat-card">`
        + `<div class="stat-num">${esc(num)}</div>`
        + (label ? `<div class="stat-label">${esc(label)}</div>` : '')
        + (note ? `<div class="stat-note">${esc(note)}</div>` : '')
        + `</div>`;
    });
    html += '</div>';
  }

  /* 關於我們 */
  if (about) {
    html += '<h2 class="section-heading">關於我們</h2>';
    html += `<div class="portal-about">${esc(about).replace(/\\n|\n/g,'<br>')}</div>`;
  }

  /* 雙軌學程課程地圖：水平三節點時間線（大一→大二→大三大四雙軌） */
  if (blocks['課程地圖'].length || blocks['雙軌'].length) {
    html += `<h2 class="section-heading">${esc(get('課程標題') || '四年怎麼學？認識甲乙制')}</h2>`;
    const intro = get('課程說明') || '中大物理獨創「甲乙制」雙軌學程：大三起依興趣選擇路徑——想廣泛修課、保留彈性走「甲制」；想及早投入研究、做專題寫論文走「乙制」。兩條路都通往升學與就業。';
    if (intro) html += `<p class="career-lead">${esc(intro)}</p>`;

    /* 三個年級節點（標題=年級、分類=課程清單以頓號或換行分隔、內容/路徑=備註） */
    if (blocks['課程地圖'].length) {
      html += '<div class="cmap-track">';
      blocks['課程地圖'].forEach((r) => {
        const year = String(data.at(r,'標題')).trim();
        const courses = String(data.at(r,'分類')).trim();
        const note = String(data.at(r,'內容/路徑')).trim();
        const list = courses.split(/[、,\n]/).map(s=>s.trim()).filter(Boolean);
        const isEvent = /小年會|盛事|專題成果|發表/.test(year + courses);
        html += `<div class="cmap-node${isEvent ? ' cmap-event' : ''}">`
          + `<div class="cmap-dot"></div>`
          + `<div class="cmap-stem"></div>`
          + `<div class="cmap-card"><div class="cmap-year">${esc(year)}</div>`;
        if (isEvent) {
          html += `<div class="cmap-event-title">${esc(courses)}</div>`
            + (note ? `<div class="cmap-event-note">${esc(note)}</div>` : '');
        } else if (list.length) {
          html += '<div class="cmap-courses">'+list.map(c=>`<span class="cmap-course">${esc(c)}</span>`).join('')+'</div>';
        }
        html += `</div></div>`;
      });
      html += '</div>';
    }

    /* 雙軌對照卡（標題=制度名、分類=適合誰一句話、內容/路徑=課程內容以頓號或換行分隔） */
    if (blocks['雙軌'].length) {
      html += '<div class="track-grid">';
      blocks['雙軌'].forEach((r, idx) => {
        const name = String(data.at(r,'標題')).trim();
        const who  = String(data.at(r,'分類')).trim();
        const items = String(data.at(r,'內容/路徑')).trim().split(/[、,\n]/).map(s=>s.trim()).filter(Boolean);
        const cls = idx===0 ? 'track-a' : 'track-b';
        html += `<div class="track-card ${cls}">`
          + `<div class="track-name">${esc(name)}</div>`
          + (who ? `<div class="track-who">${esc(who)}</div>` : '')
          + (items.length ? '<ul class="track-list">'+items.map(i=>`<li>${esc(i)}</li>`).join('')+'</ul>' : '')
          + `</div>`;
      });
      html += '</div>';
    }
  }

  /* 畢業出路：上方數據條 + 下方比例條（標題/說明可由「設定」區塊控制） */
  if (blocks['出路數據'].length || blocks['出路'].length || blocks['升學領域'].length || blocks['產業對照'].length) {
    html += `<h2 class="section-heading">${esc(get('出路標題') || '畢業出路')}</h2>`;

    /* 出路數據條（標題=數字、分類=標籤、內容/路徑=說明） */
    if (blocks['出路數據'].length) {
      html += '<div class="career-stats">';
      blocks['出路數據'].forEach(r => {
        const num   = String(data.at(r,'標題')).trim();
        const label = String(data.at(r,'分類')).trim();
        const note  = String(data.at(r,'內容/路徑')).trim();
        if (!num && !label) return;
        html += `<div class="career-stat">`
          + `<div class="career-stat-num">${esc(num)}</div>`
          + (label ? `<div class="career-stat-label">${esc(label)}</div>` : '')
          + (note ? `<div class="career-stat-note">${esc(note)}</div>` : '')
          + `</div>`;
      });
      html += '</div>';
    }

    /* 升學深造領域（水平比例條：標題=領域、分類=百分比數字、內容/路徑=說明） */
    if (blocks['升學領域'].length) {
      html += `<h3 class="career-subhead">${esc(get('升學標題') || '升學深造，他們讀什麼？')}</h3>`;
      html += renderBars(blocks['升學領域'], data);
    }

    /* 出路產業比例條（單組，標題=產業、分類=百分比） */
    if (blocks['出路'].length) {
      html += `<h3 class="career-subhead">${esc(get('就業標題') || '就業，他們進哪些產業？')}</h3>`;
      html += renderBars(blocks['出路'], data);
    }

    /* 產業對照（碩博 vs 大學部 雙色並排：標題=產業、分類=碩博%、內容/路徑=大學部%） */
    if (blocks['產業對照'].length) {
      html += `<h3 class="career-subhead">${esc(get('對照標題') || '學歷越高，越進尖端科技')}</h3>`;
      const lead = get('對照說明') || '中大物理系畢業生高度集中於半導體與光電產業——畢業 2–5 年的產業分布顯示，學歷越高、越往核心研發走。';
      html += `<p class="career-lead">${esc(lead)}</p>`;
      const item1 = esc(get('對照圖例1') || '碩博班');
      const item2 = esc(get('對照圖例2') || '大學部');
      const items = blocks['產業對照'].map(r => ({
        name: String(data.at(r,'標題')).trim(),
        pg:   parseFloat(String(data.at(r,'分類')).replace(/[^0-9.]/g,'')) || 0,
        ug:   parseFloat(String(data.at(r,'內容/路徑')).replace(/[^0-9.]/g,'')) || 0,
      })).filter(it => it.name);
      const max = Math.max(1, ...items.map(it => Math.max(it.pg, it.ug)));
      html += `<div class="cmp-legend"><span class="cmp-key"><i class="cmp-pg"></i>${item1}</span><span class="cmp-key"><i class="cmp-ug"></i>${item2}</span></div>`;
      html += '<div class="major-bars">';
      items.forEach(it => {
        html += `<div class="cmp-row"><div class="cmp-name">${esc(it.name)}</div><div class="cmp-bars">`
          + `<div class="cmp-line"><div class="cmp-fill cmp-pg" style="width:${Math.round(it.pg/max*100)}%"></div><span class="cmp-val">${it.pg}%</span></div>`
          + `<div class="cmp-line"><div class="cmp-fill cmp-ug" style="width:${Math.round(it.ug/max*100)}%"></div><span class="cmp-val">${it.ug}%</span></div>`
          + `</div></div>`;
      });
      html += '</div>';
    }
  }

  /* 常見問題 Q&A（手風琴：標題=問題、內容/路徑=答案；標題文字由「問答標題」設定控制） */
  if (blocks['問答'].length) {
    html += `<h2 class="section-heading">${esc(get('問答標題') || '常見問題 Q&A')}</h2>`;
    html += '<div class="faq-list">';
    blocks['問答'].forEach(r => {
      const q = String(data.at(r,'標題')).trim();
      const a = String(data.at(r,'內容/路徑')).trim();
      if (!q) return;
      const aHtml = esc(a).replace(/\\n|\n/g,'<br>')
        .replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>')
        .replace(/(^|[\s>])(https?:\/\/[^\s<]+)/g, '$1<a href="$2" target="_blank" rel="noopener">$2</a>');
      html += `<details class="faq-item">`
        + `<summary class="faq-q">${esc(q)}</summary>`
        + `<div class="faq-a">${aHtml}</div>`
        + `</details>`;
    });
    html += '</div>';
  }

  /* 入口（系所規章式分類連結清單，依「分類」欄分組） */
  if (blocks['入口'].length) {
    html += '<h2 class="section-heading">入口</h2>';
    const catOrder = [], catMap = {};
    blocks['入口'].forEach(r => {
      const cat = String(data.at(r,'分類')).trim() || '入口';
      if (!catMap[cat]) { catMap[cat]=[]; catOrder.push(cat); }
      catMap[cat].push(r);
    });
    catOrder.forEach(cat => {
      html += `<div class="reg-wrap"><div class="reg-cat-title">${esc(cat)}</div><ul class="reg-list">`;
      catMap[cat].forEach(r => {
        const title = String(data.at(r,'標題')).trim();
        const url   = String(data.at(r,'內容/路徑')).trim();
        if (!title) return;
        html += `<li class="reg-item"><a ${cardLinkAttrs(url)}>${esc(title)}</a></li>`;
      });
      html += '</ul></div>';
    });
  }

  area.innerHTML = html || '<p style="color:var(--muted)">（此頁面暫無內容）</p>';

  /* 影片列表：點標題切換播放器 */
  if (videos.length > 1) {
    const frame = area.querySelector('#vid-frame');
    area.querySelectorAll('.vid-item').forEach(li => {
      li.addEventListener('click', () => {
        area.querySelectorAll('.vid-item').forEach(x => x.classList.remove('active'));
        li.classList.add('active');
        if (frame) frame.src = `https://www.youtube.com/embed/${li.dataset.vid}`;
      });
    });
  }

  /* 公告區（公告頁面式條列，內容由公告總庫帶出；標題由「公告標題」設定列控制） */
  if (blocks['公告'].length) {
    const annTitle = get('公告標題') || '入學須知';
    const annArea = document.createElement('div');
    area.appendChild(annArea);
    annArea.innerHTML = `<h2 class="section-heading">${esc(annTitle)}</h2><div class="loading">載入公告資料中</div>`;
    const db = await getNewsDB();
    const byId = {};
    db.forEach(x => { if (x.wpid) byId[x.wpid] = x; });
    let ahtml = `<h2 class="section-heading">${esc(annTitle)} <span class="section-count">${blocks['公告'].length} 則</span></h2>`;
    blocks['公告'].forEach(r => {
      const title = String(data.at(r,'標題')).trim();
      const url   = String(data.at(r,'內容/路徑')).trim();
      if (!title) return;
      const m = url.match(/goTo\('[^']*\/(\d+)'\)/);
      const item = m ? byId[m[1]] : null;
      if (item) {
        ahtml += newsItemHTML(item, false);
      } else {
        ahtml += `<div class="news-item"><a class="news-item-row" style="text-decoration:none" ${cardLinkAttrs(url)}>
  <div class="news-item-title">${esc(title)}</div>
  <div class="news-item-date">→</div>
</a></div>`;
      }
    });
    annArea.innerHTML = ahtml;
  }
}

/* ── 連結卡片頁（linkcards） ── */
function renderLinkCards(area, sheets, sheetName, page) {
  const data = sheetDataRows(sheets[sheetName], ['標題','URL']);
  const heading = page.cfg['頁首大標題'] || page.label || sheetName;
  const desc = page.cfg['說明文字'] || '';

  let html = `<h2 class="section-heading">${esc(heading)} <span class="section-count">${data.rows.length} 項</span></h2>`;
  if (desc) html += `<p style="color:var(--muted);font-size:.88rem;margin:-12px 0 20px">${esc(desc)}</p>`;
  html += '<div class="portal-cards">';
  data.rows.forEach(r => {
    const title = String(data.at(r,'標題')).trim();
    const url   = String(data.at(r,'URL')).trim();
    const note  = String(data.at(r,'說明')).trim();
    if (!title) return;
    const isPdf = /\.pdf(\?|$)/i.test(url);
    html += `<a class="portal-card" ${cardLinkAttrs(url)}>
  <div><div class="portal-card-title">${isPdf ? '📄 ' : ''}${esc(title)}</div>${note ? `<div class="portal-card-desc">${esc(note)}</div>` : ''}</div>
  <div class="portal-card-arrow">→</div>
</a>`;
  });
  html += '</div>';
  area.innerHTML = html;
}

/* 取得公告總庫資料（任一已載入分區中含「公告總庫」工作頁者） */
async function getNewsDB() {
  const find = () => {
    for (const sec of SECTIONS.filter(s => s.show)) {
      const d = DATA[sec.key];
      if (d && d['公告總庫']) return d['公告總庫'];
    }
    return null;
  };
  let ws = find();
  if (!ws) {
    await Promise.allSettled(Object.values(LOADP));
    ws = find();
  }
  return ws ? parseNewsSheet(ws) : [];
}

/* ── newslist：條列式公告清單頁（依工作頁的 goTo WP-ID 從公告總庫撈完整內容） ── */
async function renderNewsListPage(area, sheets, sheetName, page) {
  const data = sheetDataRows(sheets[sheetName], ['標題','URL']);
  const heading = page.cfg['頁首大標題'] || page.label || sheetName;
  const desc = page.cfg['說明文字'] || '';

  area.innerHTML = '<div class="loading">載入公告資料中</div>';
  const db = await getNewsDB();
  const byId = {};
  db.forEach(r => { if (r.wpid) byId[r.wpid] = r; });

  let html = `<h2 class="section-heading">${esc(heading)} <span class="section-count">${data.rows.length} 則</span></h2>`;
  if (desc) html += `<p style="color:var(--muted);font-size:.88rem;margin:-12px 0 20px">${esc(desc)}</p>`;

  data.rows.forEach(r => {
    const title = String(data.at(r,'標題')).trim();
    const url   = String(data.at(r,'URL')).trim();
    if (!title) return;
    const m = url.match(/goTo\('[^']*\/(\d+)'\)/);
    const item = m ? byId[m[1]] : null;
    if (item) {
      /* 公告總庫有資料 → 完整條列卡片（含日期、可展開內文與連結） */
      html += newsItemHTML(item, false);
    } else {
      /* 無對應公告 → 簡單列，點擊依 URL 行為 */
      html += `<div class="news-item"><a class="news-item-row" style="text-decoration:none" ${cardLinkAttrs(url)}>
  <div class="news-item-title">${esc(title)}</div>
  <div class="news-item-date">→</div>
</a></div>`;
    }
  });

  area.innerHTML = html;
}

function renderEmi(area, sheets, sheetName) {
  const ws = sheets[sheetName];
  if (!ws || !ws.length) { area.innerHTML = '<div class="error-msg">找不到 EMI 資料工作頁</div>'; return; }

  const CAT_CLS = { '教師專區':'teacher', '學生專區':'student', '助教專區':'ta', '資源連結':'resource' };

  const rows = ws.filter(r => {
    const cat = String(r['分類']||'').trim();
    if (!cat || cat === '分類' || cat === '［說明］') return false;
    return isOn(r['啟用']);
  });

  const groupOrder = [], groupMap = {};
  rows.forEach(r => {
    const cat = String(r['分類']||'').trim();
    const name= String(r['名稱']||'').trim();
    const url = String(r['URL']  ||'').trim();
    if (!cat || !name || !url) return;
    if (!groupMap[cat]) { groupMap[cat]=[]; groupOrder.push(cat); }
    groupMap[cat].push({ name, url });
  });

  if (!groupOrder.length) {
    area.innerHTML = '<div class="error-msg">EMI 資料為空，請確認工作頁格式正確。</div>';
    return;
  }

  let html = '<div class="emi-wrap">';
  groupOrder.forEach(cat => {
    const cls = CAT_CLS[cat] || 'resource';
    html += `<div class="emi-cat-section emi-cat-${esc(cls)}">`;
    html += `<div class="emi-cat-title">${esc(cat)}</div>`;
    html += `<div class="emi-cat-div"></div>`;
    html += '<div class="emi-links-grid">';
    groupMap[cat].forEach(({ name, url }) => {
      html += `<a class="emi-link-card" ${cardLinkAttrs(url)}><div class="emi-link-name">${esc(name)}</div><div class="emi-link-arrow">→</div></a>`;
    });
    html += '</div></div>';
  });
  html += '</div>';
  area.innerHTML = html;
}

/* ════════════════════════════════════════════════════
   INIT
════════════════════════════════════════════════════ */
async function init() {
  /* 補捉 app.js 載入前就已失敗的靜態圖（如 logo）：補發一次 error 事件交由委派處理 */
  document.querySelectorAll('img[data-onerr]').forEach(img => {
    if (img.complete && img.naturalWidth === 0) img.dispatchEvent(new Event('error'));
  });
  document.getElementById('list-view').innerHTML = '<div class="loading">讀取設定檔與資料中</div>';

  await applyWebsiteConfig();

  /* 全部 xlsx 同時開始載入（不互相等待）；Promise 存入 LOADP */
  SECTIONS.filter(s => s.show).forEach(sec => {
    LOADP[sec.key] = loadXlsx(sec.xlsxFile)
      .then(d => {
        DATA[sec.key] = d;
        const vb = document.getElementById('vb-text');   /* 防呆：徽章缺失不可拖垮資料載入 */
        if (vb && vb.textContent === '載入中…')
          setVersionBadge(d._lastMod);
      })
      .catch(e => {
        DATA[sec.key] = { '頁面設定':[] };
        console.warn('Failed to load', sec.xlsxFile, e.message);
      });
  });

  /* 決定首頁目標分區（hash 優先，否則第一個分區） */
  const hash = location.hash.replace('#','');
  const [rawSec, rawPg, hashSub] = hash.split('/');
  const hashSec = decSeg(rawSec), hashPg = decSeg(rawPg);
  if (hashSub) newsPendingParam = decodeURIComponent(hashSub);
  let targetKey = SECTIONS.find(s => s.show && s.key === hashSec)?.key
               || SECTIONS.find(s => s.show)?.key;

  /* 只等目標分區載完就先渲染（小檔案幾百 ms 內完成，不被公告 1.7MB 拖慢） */
  if (targetKey) {
    try { await LOADP[targetKey]; } catch(_) {}
    buildNav();
    const pages = getPages(targetKey);
    const pg = (hashSec === targetKey && hashPg && pages.find(p => p.id === hashPg)) ? hashPg
             : (pages[0]?.id);
    if (pg) navigateTo(targetKey, pg);
  }

  /* 其餘分區於背景載完後重建導覽列（補齊下拉子選單） */
  Promise.allSettled(Object.values(LOADP)).then(() => buildNav());

  window.addEventListener('popstate', () => {
    const h = location.hash.replace('#','');
    const [rawSec, rawPg, sub] = h.split('/');
    const sec = decSeg(rawSec), pg = decSeg(rawPg);
    if (sub) newsPendingParam = decodeURIComponent(sub);
    const s = SECTIONS.find(s => s.show && s.key === sec);
    if (s && pg) navigateTo(sec, pg);
  });
}

init();
