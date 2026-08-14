/* ===================== 스키마 ===================== */
const FIELDS = [
  '회원번호','등록구분','최초등록구분','회원구분','회원상태','성명','나이','성별',
  '생일/설립일','휴대전화번호','이메일','최초납부년월','최종납부년월','납부여부',
  '납부방법','탈퇴여부','Web아이디','메모'
];
const STATUS_ORDER = ['정착', '잠재후원자', '신규(정기)', '신규(일시)', '이탈'];
const STATUS_COLOR = {
  '정착': { fill: '#1B4B43', badge: 'badge-정착' },
  '잠재후원자': { fill: '#B87F23', badge: 'badge-잠재후원자' },
  '신규(정기)': { fill: '#3B6EA5', badge: 'badge-신규' },
  '신규(일시)': { fill: '#6B8FB8', badge: 'badge-신규' },
  '이탈': { fill: '#A8443B', badge: 'badge-이탈' },
};
const SELECT_OPTIONS = {
  '등록구분': ['개인', '기업/단체', '외국인'],
  '최초등록구분': ['정기', '일시', '비후원'],
  '회원구분': ['개인', '의료종사자', '교회', '기업/단체', '병원', '제약회사', '의료기기회사', '정부기관'],
  '회원상태': STATUS_ORDER,
  '성별': ['남', '여'],
  '납부여부': ['Y', 'N'],
  '납부방법': ['CMS', '계좌이체', '신용카드', 'KakaoPay', 'PAYCO'],
  '탈퇴여부': ['N', 'Y'],
};
const PAGE_SIZE = 50;
const TODAY = new Date();
const TODAY_LABEL = `${TODAY.getFullYear()}-${String(TODAY.getMonth()+1).padStart(2,'0')}-${String(TODAY.getDate()).padStart(2,'0')}`;
const STORAGE_KEY = 'donor_crm_working_v1';

let DONORS = [];          // 작업중인 전체 배열 (메모리)
let sourceFileName = '';
let filters = { status: new Set(), firstReg: new Set(), regType: new Set(), memberType: new Set(), payMethod: new Set(), gender: new Set() };
let segment = null;
let searchTerm = '';
let sortMode = 'default';
let page = 1;
let activeUid = null;   // 현재 drawer 대상 _uid (신규 추가시 null)

/* ===================== 유틸 ===================== */
function uid() { return 'd_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36); }
function uniqueValues(field) {
  const set = new Set();
  DONORS.forEach(d => { if (d[field]) set.add(d[field]); });
  return Array.from(set).sort();
}
function hasMemo(d) { return !!(d['메모'] && String(d['메모']).trim()); }

/* ===================== 파일 파싱 (SheetJS) ===================== */
function excelSerialToDateStr(v, isMonth) {
  // v가 Date 객체인 경우
  if (v instanceof Date && !isNaN(v)) {
    const y = v.getFullYear(), m = String(v.getMonth()+1).padStart(2,'0'), day = String(v.getDate()).padStart(2,'0');
    return isMonth ? `${y}-${m}` : `${y}-${m}-${day}`;
  }
  return v;
}

function normalizeRecord(raw) {
  const rec = {};
  FIELDS.forEach(f => { rec[f] = null; });
  Object.keys(raw).forEach(k => {
    const key = k.trim();
    if (!(key in rec) && FIELDS.includes(key)) rec[key] = raw[k];
    else if (key in rec) rec[key] = raw[k];
  });
  // 타입 정리
  ['최초납부년월','최종납부년월'].forEach(f => { rec[f] = excelSerialToDateStr(rec[f], true); });
  ['생일/설립일'].forEach(f => { rec[f] = excelSerialToDateStr(rec[f], false); });
  if (rec['나이'] != null && rec['나이'] !== '') rec['나이'] = Math.round(Number(rec['나이'])) || null;
  else rec['나이'] = null;
  FIELDS.forEach(f => {
    if (typeof rec[f] === 'string') {
      const t = rec[f].trim();
      rec[f] = t === '' ? null : t;
    }
  });
  rec['_uid'] = rec['회원번호'] ? String(rec['회원번호']) : uid();
  rec['_createdAt'] = null;
  rec['_updatedAt'] = null;
  rec['_isNew'] = false;
  return rec;
}

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
  return rows.map(normalizeRecord);
}

/* ===================== localStorage ===================== */
function saveLocal() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ data: DONORS, savedAt: new Date().toISOString(), fileName: sourceFileName }));
    return true;
  } catch (e) {
    console.error('저장 실패', e);
    return false;
  }
}
function loadLocalMeta() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed;
  } catch (e) { return null; }
}

/* ===================== 오늘 기준 상태 규칙 ===================== */
function monthsBetween(ymStr, refDate) {
  if (!ymStr) return null;
  const parts = String(ymStr).split('-');
  const y = parseInt(parts[0], 10), m = parseInt(parts[1], 10);
  if (!y || !m) return null;
  const refY = refDate.getFullYear(), refM = refDate.getMonth() + 1;
  return (refY - y) * 12 + (refM - m);
}
function computeRecommended(d) {
  const first = d['최초납부년월'], last = d['최종납부년월'];
  const firstType = d['최초등록구분'], payYN = d['납부여부'];
  if (!first) return { status: '잠재후원자', reason: '후원 이력 없음', confidence: 'confirmed' };
  const msf = monthsBetween(first, TODAY);
  const msl = last ? monthsBetween(last, TODAY) : msf;
  if (firstType === '정기') {
    if (payYN === 'Y') {
      if (msf < 6) return { status: '신규(정기)', reason: `정기후원 시작 ${msf}개월째(6개월 미만)`, confidence: 'confirmed' };
      return { status: '정착', reason: `정기후원 ${msf}개월째 유지중`, confidence: 'confirmed' };
    }
    return { status: '이탈', reason: '정기후원 해지(납부여부 N)', confidence: 'confirmed' };
  }
  if (firstType === '일시') {
    const isSingle = !last || first === last;
    if (msl >= 12) return { status: '이탈', reason: `최종 후원 후 ${msl}개월 경과(12개월 내 추가 후원 없음)`, confidence: 'confirmed' };
    if (isSingle) return { status: '신규(일시)', reason: `첫 일시후원 후 ${msf}개월(12개월 이내)`, confidence: 'confirmed' };
    return { status: '정착', reason: '일시후원 반복 중 · 12개월 내 추가 후원 있음(연간 빈도는 원자료로 확정 불가)', confidence: 'estimated' };
  }
  return { status: d['회원상태'], reason: '규칙 미해당(원본 상태 유지)', confidence: 'confirmed' };
}
function needsChange(d) { return computeRecommended(d).status !== d['회원상태']; }

/* ===================== 세그먼트 ===================== */
const SEGMENTS = [
  { key: 'all', label: '전체 후원자', test: () => true },
  { key: 'status_change', label: `오늘(${TODAY_LABEL}) 기준 변경 대상`, test: d => needsChange(d) },
  { key: 'convert_target', label: '정기후원 전환 대상', test: d => d['회원상태'] === '잠재후원자' && d['휴대전화번호'] },
  { key: 'potential', label: '잠재후원자 전체', test: d => d['회원상태'] === '잠재후원자' },
  { key: 'settled', label: '정착 후원자', test: d => d['회원상태'] === '정착' },
  { key: 'churned', label: '이탈 후원자', test: d => d['회원상태'] === '이탈' },
  { key: 'new_recent', label: '신규 후원자', test: d => d['회원상태'] === '신규(정기)' || d['회원상태'] === '신규(일시)' },
  { key: 'no_contact', label: '연락처 정보 없음', test: d => !d['휴대전화번호'] && !d['이메일'] },
  { key: 'memo_only', label: '메모 등록된 후원자', test: d => hasMemo(d) },
  { key: 'newly_added', label: '내가 추가한 항목', test: d => d['_isNew'] },
];

/* ===================== 필터/정렬 ===================== */
function applyFilters() {
  let list = DONORS;
  if (segment && segment !== 'all') {
    const seg = SEGMENTS.find(s => s.key === segment);
    if (seg) list = list.filter(seg.test);
  }
  if (filters.status.size) list = list.filter(d => filters.status.has(d['회원상태']));
  if (filters.firstReg.size) list = list.filter(d => filters.firstReg.has(d['최초등록구분']));
  if (filters.regType.size) list = list.filter(d => filters.regType.has(d['등록구분']));
  if (filters.memberType.size) list = list.filter(d => filters.memberType.has(d['회원구분']));
  if (filters.payMethod.size) list = list.filter(d => filters.payMethod.has(d['납부방법']));
  if (filters.gender.size) list = list.filter(d => filters.gender.has(d['성별']));

  if (searchTerm.trim()) {
    const q = searchTerm.trim().toLowerCase();
    list = list.filter(d =>
      (d['성명'] && d['성명'].toLowerCase().includes(q)) ||
      (d['휴대전화번호'] && d['휴대전화번호'].replace(/-/g, '').includes(q.replace(/-/g, ''))) ||
      (d['이메일'] && d['이메일'].toLowerCase().includes(q)) ||
      (d['회원번호'] && String(d['회원번호']).includes(q))
    );
  }
  return sortList(list);
}
function sortList(list) {
  const arr = list.slice();
  switch (sortMode) {
    case 'name': arr.sort((a, b) => (a['성명'] || '').localeCompare(b['성명'] || '', 'ko')); break;
    case 'age_desc': arr.sort((a, b) => (b['나이'] || -1) - (a['나이'] || -1)); break;
    case 'age_asc': arr.sort((a, b) => (a['나이'] == null ? 999 : a['나이']) - (b['나이'] == null ? 999 : b['나이'])); break;
    case 'last_pay_desc': arr.sort((a, b) => (b['최종납부년월'] || '').localeCompare(a['최종납부년월'] || '')); break;
    case 'last_pay_asc': arr.sort((a, b) => (a['최종납부년월'] || '9999-99').localeCompare(b['최종납부년월'] || '9999-99')); break;
    case 'recent_edit': arr.sort((a, b) => (b['_updatedAt'] || '').localeCompare(a['_updatedAt'] || '')); break;
    default: arr.sort((a, b) => String(a['회원번호'] || '').localeCompare(String(b['회원번호'] || '')));
  }
  return arr;
}

/* ===================== 렌더: 사이드바 ===================== */
function renderSegmentList() {
  const el = document.getElementById('segmentList');
  el.innerHTML = '';
  SEGMENTS.forEach(seg => {
    const count = seg.key === 'all' ? DONORS.length : DONORS.filter(seg.test).length;
    const btn = document.createElement('button');
    btn.className = 'side-btn' + ((segment === seg.key || (seg.key === 'all' && !segment)) ? ' active' : '');
    btn.innerHTML = `<span>${seg.label}</span><span class="n tabular">${count.toLocaleString()}</span>`;
    btn.onclick = () => { segment = (seg.key === 'all') ? null : seg.key; page = 1; renderAll(); };
    el.appendChild(btn);
  });
}
function renderChipGroup(containerId, values, filterKey) {
  const el = document.getElementById(containerId);
  el.innerHTML = '';
  values.forEach(v => {
    const chip = document.createElement('span');
    chip.className = 'chip' + (filters[filterKey].has(v) ? ' active' : '');
    chip.textContent = v;
    chip.onclick = () => {
      if (filters[filterKey].has(v)) filters[filterKey].delete(v); else filters[filterKey].add(v);
      page = 1; renderAll();
    };
    el.appendChild(chip);
  });
}

/* ===================== 렌더: 상태 변경 배너 ===================== */
function renderStatusBanner() {
  const el = document.getElementById('statusChangeBanner');
  const flagged = DONORS.filter(needsChange).map(d => ({ d, rec: computeRecommended(d) }));
  const confirmedList = flagged.filter(x => x.rec.confidence === 'confirmed');
  const estimatedList = flagged.filter(x => x.rec.confidence === 'estimated');
  if (!flagged.length) { el.innerHTML = ''; return; }
  el.innerHTML = `
    <div class="status-banner">
      <div class="sb-main">
        <div class="sb-title">오늘(${TODAY_LABEL}) 기준 회원상태 변경 대상 ${flagged.length.toLocaleString()}명</div>
        <div class="sb-sub">확정 규칙 적용 ${confirmedList.length.toLocaleString()}명 · 추정 판단 필요 ${estimatedList.length.toLocaleString()}명</div>
      </div>
      <div class="sb-actions">
        <button class="btn-ghost" id="sbReview">목록 확인</button>
        <button class="btn-primary" id="sbBulkApply" ${confirmedList.length ? '' : 'disabled'}>확정 대상 일괄 반영(${confirmedList.length})</button>
      </div>
    </div>`;
  document.getElementById('sbReview').onclick = () => { segment = 'status_change'; page = 1; renderAll(); window.scrollTo({top:0,behavior:'smooth'}); };
  document.getElementById('sbBulkApply').onclick = () => {
    if (!confirmedList.length) return;
    const ok = confirm(`확정 규칙에 해당하는 ${confirmedList.length}명의 회원상태를 오늘 기준 권장 상태로 일괄 반영합니다.\n(추정 판단이 필요한 ${estimatedList.length}명은 제외되며 개별 확인이 필요합니다)\n계속할까요?`);
    if (!ok) return;
    const now = new Date().toISOString().slice(0, 16).replace('T', ' ');
    confirmedList.forEach(({ d, rec }) => { d['회원상태'] = rec.status; d['_updatedAt'] = now; });
    saveLocal();
    renderAll();
  };
}

/* ===================== 렌더: 리본/통계 ===================== */
function renderRibbon() {
  const ribbon = document.getElementById('ribbon');
  const legend = document.getElementById('ribbonLegend');
  ribbon.innerHTML = ''; legend.innerHTML = '';
  const total = DONORS.length;
  if (!total) return;
  const counts = {};
  STATUS_ORDER.forEach(s => counts[s] = 0);
  DONORS.forEach(d => { counts[d['회원상태']] = (counts[d['회원상태']] || 0) + 1; });
  STATUS_ORDER.forEach(s => {
    const count = counts[s] || 0;
    if (!count) return;
    const pct = (count / total) * 100;
    const seg = document.createElement('div');
    seg.className = 'ribbon-seg';
    seg.style.width = pct + '%';
    seg.style.background = STATUS_COLOR[s].fill;
    seg.title = `${s}: ${count.toLocaleString()}명 (${pct.toFixed(1)}%)`;
    seg.onclick = () => { filters.status.clear(); filters.status.add(s); page = 1; renderAll(); };
    ribbon.appendChild(seg);
    const item = document.createElement('div');
    item.className = 'legend-item';
    item.innerHTML = `<span class="legend-dot" style="background:${STATUS_COLOR[s].fill}"></span>${s} · ${count.toLocaleString()}명`;
    item.onclick = seg.onclick;
    legend.appendChild(item);
  });
}
function renderStatRow(filteredCount) {
  const el = document.getElementById('statRow');
  const total = DONORS.length;
  const activePay = DONORS.filter(d => d['납부여부'] === 'Y').length;
  const withMemo = DONORS.filter(hasMemo).length;
  const cards = [
    { num: total.toLocaleString(), lbl: '전체 후원자' },
    { num: filteredCount.toLocaleString(), lbl: '현재 필터 결과' },
    { num: activePay.toLocaleString(), lbl: '납부 진행중(Y)' },
    { num: withMemo.toLocaleString(), lbl: '메모 등록됨' },
  ];
  el.innerHTML = cards.map(c => `<div class="stat-card"><div class="num tabular">${c.num}</div><div class="lbl">${c.lbl}</div></div>`).join('');
}

/* ===================== 렌더: 테이블 ===================== */
function renderTable() {
  const list = applyFilters();
  const totalPages = Math.max(1, Math.ceil(list.length / PAGE_SIZE));
  if (page > totalPages) page = totalPages;
  const start = (page - 1) * PAGE_SIZE;
  const pageItems = list.slice(start, start + PAGE_SIZE);

  const tbody = document.getElementById('tableBody');
  const emptyState = document.getElementById('emptyState');
  document.getElementById('resultCount').textContent = `${list.length.toLocaleString()}명 검색됨`;

  if (!list.length) {
    tbody.innerHTML = '';
    emptyState.style.display = 'block';
  } else {
    emptyState.style.display = 'none';
    tbody.innerHTML = pageItems.map(d => {
      const status = d['회원상태'];
      const badgeClass = STATUS_COLOR[status] ? STATUS_COLOR[status].badge : '';
      const memoFlag = hasMemo(d) ? ' 📝' : '';
      const newTag = d['_isNew'] ? '<span class="new-tag">신규입력</span>' : '';
      const rec = computeRecommended(d);
      const changeFlag = rec.status !== status
        ? `<span class="change-flag ${rec.confidence}" title="권장: ${rec.status} · ${rec.reason}">→ ${rec.status}</span>` : '';
      return `
      <tr data-uid="${d['_uid']}">
        <td class="tabular muted">${d['회원번호'] || '-'}</td>
        <td class="name-cell">${d['성명'] || '-'}${memoFlag}${newTag}</td>
        <td><span class="badge ${badgeClass}">${status || '-'}</span>${changeFlag}</td>
        <td class="tabular">${d['나이'] != null ? d['나이'] : '-'}</td>
        <td>${d['성별'] || '-'}</td>
        <td class="tabular">${d['휴대전화번호'] || '<span class="muted">-</span>'}</td>
        <td class="muted">${d['이메일'] || '-'}</td>
        <td class="tabular">${d['최종납부년월'] || '-'}</td>
        <td>${d['납부방법'] || '<span class="muted">-</span>'}</td>
      </tr>`;
    }).join('');
    Array.from(tbody.querySelectorAll('tr')).forEach(tr => { tr.onclick = () => openDrawer(tr.dataset.uid); });
  }
  renderPager(list.length, totalPages);
  renderStatRow(list.length);
}
function renderPager(count, totalPages) {
  const el = document.getElementById('pager');
  if (!count) { el.innerHTML = ''; return; }
  el.innerHTML = `<button id="prevPg" ${page <= 1 ? 'disabled' : ''}>← 이전</button><span class="pg-info">${page} / ${totalPages} 페이지</span><button id="nextPg" ${page >= totalPages ? 'disabled' : ''}>다음 →</button>`;
  const prev = document.getElementById('prevPg'), next = document.getElementById('nextPg');
  if (prev) prev.onclick = () => { page--; renderTable(); window.scrollTo({top:0,behavior:'smooth'}); };
  if (next) next.onclick = () => { page++; renderTable(); window.scrollTo({top:0,behavior:'smooth'}); };
}

/* ===================== 드로어: 상세/수정/추가 ===================== */
function fieldInput(f, d) {
  const val = (d && d[f] != null) ? d[f] : '';
  if (SELECT_OPTIONS[f]) {
    const opts = SELECT_OPTIONS[f].map(o => `<option value="${o}" ${val === o ? 'selected' : ''}>${o}</option>`).join('');
    return `<select data-field="${f}"><option value="">-</option>${opts}</select>`;
  }
  if (f === '메모') return `<textarea data-field="${f}">${val}</textarea>`;
  let placeholder = '';
  if (f === '생일/설립일') placeholder = 'YYYY-MM-DD';
  if (f === '최초납부년월' || f === '최종납부년월') placeholder = 'YYYY-MM';
  if (f === '나이') return `<input type="number" data-field="${f}" value="${val}" placeholder="${placeholder}" />`;
  return `<input type="text" data-field="${f}" value="${val}" placeholder="${placeholder}" />`;
}
function renderForm(d) {
  const grid = document.getElementById('formGrid');
  const fullWidth = new Set(['메모']);
  grid.innerHTML = FIELDS.map(f => `
    <div class="form-field ${fullWidth.has(f) ? 'full' : ''}">
      <label>${f}${f === '성명' ? ' *' : ''}</label>
      ${fieldInput(f, d)}
    </div>`).join('');
}
function readForm() {
  const out = {};
  document.querySelectorAll('#formGrid [data-field]').forEach(el => {
    let v = el.value.trim();
    out[el.dataset.field] = v === '' ? null : v;
  });
  if (out['나이'] != null) out['나이'] = Math.round(Number(out['나이'])) || null;
  return out;
}

function openDrawer(uidVal) {
  activeUid = uidVal;
  const d = DONORS.find(x => x['_uid'] === uidVal);
  if (!d) return;

  document.getElementById('formTitle').textContent = '정보 수정';
  document.getElementById('drawerName').textContent = d['성명'] || '-';
  document.getElementById('drawerSub').textContent = `회원번호 ${d['회원번호'] || '-'} · ${d['등록구분'] || '-'} · ${d['회원구분'] || '-'}`;
  document.getElementById('deleteBtn').style.display = 'inline-block';

  const metaNote = document.getElementById('metaNote');
  const metaBits = [];
  if (d['_isNew']) metaBits.push('직접 추가한 항목');
  if (d['_updatedAt']) metaBits.push(`마지막 수정: ${d['_updatedAt']}`);
  metaNote.textContent = metaBits.join(' · ');

  const rec = computeRecommended(d);
  const recBox = document.getElementById('recommendBox');
  if (rec.status !== d['회원상태']) {
    recBox.innerHTML = `
      <div class="rec-title">오늘(${TODAY_LABEL}) 기준 권장 상태${rec.confidence === 'estimated' ? ' · 추정' : ''}</div>
      <div class="rec-body">${d['회원상태']} → <strong>${rec.status}</strong></div>
      <div class="rec-reason">${rec.reason}</div>
      <button class="btn-ghost sm" id="applyRecBtn">권장 상태로 채우기</button>`;
    recBox.classList.add('show');
  } else { recBox.innerHTML = ''; recBox.classList.remove('show'); }

  renderForm(d);
  const applyBtn = document.getElementById('applyRecBtn');
  if (applyBtn) applyBtn.onclick = () => { document.querySelector('[data-field="회원상태"]').value = rec.status; };

  document.getElementById('saveMsg').classList.remove('show');
  document.getElementById('overlay').classList.add('open');
  document.getElementById('drawer').classList.add('open');
}

function openNewDrawer() {
  activeUid = null;
  document.getElementById('formTitle').textContent = '새 후원자 추가';
  document.getElementById('drawerName').textContent = '새 후원자';
  document.getElementById('drawerSub').textContent = '필수: 성명';
  document.getElementById('deleteBtn').style.display = 'none';
  document.getElementById('recommendBox').classList.remove('show');
  document.getElementById('recommendBox').innerHTML = '';
  document.getElementById('metaNote').textContent = '';
  renderForm(null);
  document.getElementById('saveMsg').classList.remove('show');
  document.getElementById('overlay').classList.add('open');
  document.getElementById('drawer').classList.add('open');
}

function closeDrawer() {
  document.getElementById('overlay').classList.remove('open');
  document.getElementById('drawer').classList.remove('open');
  activeUid = null;
}

function handleSave() {
  const values = readForm();
  if (!values['성명']) { alert('성명은 필수 입력 항목입니다.'); return; }
  const now = new Date().toISOString().slice(0, 16).replace('T', ' ');

  if (activeUid) {
    const d = DONORS.find(x => x['_uid'] === activeUid);
    if (!d) return;
    FIELDS.forEach(f => { d[f] = values[f]; });
    d['_updatedAt'] = now;
  } else {
    const rec = { ...values };
    FIELDS.forEach(f => { if (!(f in rec)) rec[f] = null; });
    if (!rec['회원번호']) rec['회원번호'] = generateMemberNo();
    rec['_uid'] = uid();
    rec['_createdAt'] = now;
    rec['_updatedAt'] = now;
    rec['_isNew'] = true;
    DONORS.unshift(rec);
  }
  saveLocal();

  const msg = document.getElementById('saveMsg');
  msg.textContent = '저장됨';
  msg.classList.add('show');
  setTimeout(() => msg.classList.remove('show'), 1500);

  renderAll();
  setTimeout(closeDrawer, 350);
}

function handleDelete() {
  if (!activeUid) return;
  const d = DONORS.find(x => x['_uid'] === activeUid);
  if (!d) return;
  const ok = confirm(`${d['성명'] || '이 후원자'} 정보를 삭제합니다. 되돌릴 수 없습니다. 계속할까요?`);
  if (!ok) return;
  DONORS = DONORS.filter(x => x['_uid'] !== activeUid);
  saveLocal();
  closeDrawer();
  renderAll();
}

function generateMemberNo() {
  let max = 0;
  DONORS.forEach(d => {
    const n = parseInt(String(d['회원번호'] || '').replace(/\D/g, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  return String(max + 1).padStart(8, '0');
}

/* ===================== 로그 박스 ===================== */
function renderLog() {
  const total = DONORS.length;
  const log =
`[source] ${sourceFileName || '-'} (브라우저 내 처리, 서버 미전송)
[rows] ${total.toLocaleString()}건 · 필드 ${FIELDS.length}개
[storage] localStorage(이 브라우저 전용) · 자동 저장
[rule] 상태 재계산 기준일 TODAY=${TODAY_LABEL}`;
  document.getElementById('logBox').textContent = log;
  document.getElementById('brandSub').textContent = `${total.toLocaleString()}명 · CRM 세그먼트`;
}

/* ===================== 전체 렌더 ===================== */
function renderAll() {
  renderLog();
  renderStatusBanner();
  renderSegmentList();
  renderChipGroup('filterStatus', STATUS_ORDER, 'status');
  renderChipGroup('filterFirstReg', uniqueValues('최초등록구분'), 'firstReg');
  renderChipGroup('filterRegType', uniqueValues('등록구분'), 'regType');
  renderChipGroup('filterMemberType', uniqueValues('회원구분'), 'memberType');
  renderChipGroup('filterPayMethod', uniqueValues('납부방법'), 'payMethod');
  renderChipGroup('filterGender', uniqueValues('성별'), 'gender');
  renderRibbon();
  renderTable();
}

/* ===================== 내보내기 ===================== */
function exportXlsx() {
  const rows = DONORS.map(d => {
    const row = {};
    FIELDS.forEach(f => { row[f] = d[f]; });
    return row;
  });
  const ws = XLSX.utils.json_to_sheet(rows, { header: FIELDS });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const ts = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `후원자관리_${ts}.xlsx`);
}

/* ===================== 앱 시작/업로드 흐름 ===================== */
function showApp() {
  document.getElementById('uploadScreen').style.display = 'none';
  document.getElementById('app').classList.add('show');
  page = 1; segment = null; searchTerm = ''; sortMode = 'default';
  filters = { status: new Set(), firstReg: new Set(), regType: new Set(), memberType: new Set(), payMethod: new Set(), gender: new Set() };
  renderAll();
}

function handleFile(file) {
  const err = document.getElementById('uploadError');
  err.classList.remove('show');
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      DONORS = parseWorkbook(e.target.result);
      sourceFileName = file.name;
      saveLocal();
      showApp();
    } catch (ex) {
      err.textContent = '파일을 읽는 중 오류가 발생했습니다. 엑셀(.xlsx) 형식을 확인해주세요.';
      err.classList.add('show');
      console.error(ex);
    }
  };
  reader.readAsArrayBuffer(file);
}

function initUploadScreen() {
  const meta = loadLocalMeta();
  if (meta && meta.data && meta.data.length) {
    const box = document.getElementById('resumeBox');
    const savedDate = new Date(meta.savedAt);
    document.getElementById('resumeText').textContent =
      `이전 작업 있음 · ${meta.fileName || '파일'} · ${meta.data.length.toLocaleString()}건 · ${savedDate.toLocaleString('ko-KR')} 저장됨`;
    box.style.display = 'flex';
    document.getElementById('resumeBtn').onclick = () => {
      DONORS = meta.data;
      sourceFileName = meta.fileName || '(이전 작업)';
      showApp();
    };
  }

  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  dropzone.onclick = () => fileInput.click();
  fileInput.onchange = () => { if (fileInput.files[0]) handleFile(fileInput.files[0]); };
  ['dragenter', 'dragover'].forEach(evt => dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach(evt => dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.remove('drag'); }));
  dropzone.addEventListener('drop', (e) => { if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });
}

function init() {
  initUploadScreen();

  document.getElementById('searchInput').addEventListener('input', e => { searchTerm = e.target.value; page = 1; renderTable(); });
  document.getElementById('sortSelect').addEventListener('change', e => { sortMode = e.target.value; page = 1; renderTable(); });
  document.getElementById('resetFilters').addEventListener('click', () => {
    filters = { status: new Set(), firstReg: new Set(), regType: new Set(), memberType: new Set(), payMethod: new Set(), gender: new Set() };
    segment = null; searchTerm = ''; document.getElementById('searchInput').value = ''; page = 1; renderAll();
  });
  document.getElementById('drawerClose').addEventListener('click', closeDrawer);
  document.getElementById('overlay').addEventListener('click', closeDrawer);
  document.getElementById('saveBtn').addEventListener('click', handleSave);
  document.getElementById('deleteBtn').addEventListener('click', handleDelete);
  document.getElementById('addDonorBtn').addEventListener('click', openNewDrawer);
  document.getElementById('exportBtn').addEventListener('click', exportXlsx);
  document.getElementById('reuploadBtn').addEventListener('click', () => {
    const ok = confirm('현재 작업 화면을 닫고 새 파일을 업로드합니다. (자동 저장된 현재 데이터는 유지되며, 다음에 "이어서 열기"로 복구할 수 있습니다)');
    if (!ok) return;
    document.getElementById('app').classList.remove('show');
    document.getElementById('uploadScreen').style.display = 'flex';
    initUploadScreen();
  });
}

init();
