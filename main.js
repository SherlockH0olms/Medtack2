// ─── DATA ──────────────────────────────────────────
const T = () => { const d = new Date(); return [d.getHours(), d.getMinutes(), d.getSeconds()].map(n => String(n).padStart(2, '0')).join(':') };

const PATIENTS = [
    { id: 1, name: 'Əliyev Kamran', age: 54, room: 'A-201', uid: 'NF:A3:21:9C', diagnosis: 'Kardioloji', allergies: ['Penisilin'], meds: [1, 3], status: 'Stabil', color: '#0EA5E9', bg: 'rgba(14,165,233,0.15)' },
    { id: 2, name: 'Həsənova Aytən', age: 38, room: 'B-104', uid: 'NF:B7:45:2D', diagnosis: 'Nevroloji', allergies: [], meds: [2, 4], status: 'İzləmə', color: '#10B981', bg: 'rgba(16,185,129,0.15)' },
    { id: 3, name: 'Babayev Tural', age: 71, room: 'C-307', uid: 'NF:C1:8F:77', diagnosis: 'Ortopediya', allergies: ['İbuprofen'], meds: [3, 5], status: 'Kritik', color: '#EF4444', bg: 'rgba(239,68,68,0.15)' },
    { id: 4, name: 'Qasımova Nigar', age: 29, room: 'A-105', uid: 'NF:D9:3A:61', diagnosis: 'Daxili Xəstəliklər', allergies: [], meds: [1, 2], status: 'Stabil', color: '#8B5CF6', bg: 'rgba(139,92,246,0.15)' },
];

const MEDS = [
    { id: 1, name: 'Metoprolol 50mg', uid: 'RF:M1:4A:2B', dose: '1x1', expires: '2026-08', type: 'Ürək' },
    { id: 2, name: 'Levetirasetam 500mg', uid: 'RF:M2:9C:F3', dose: '2x1', expires: '2025-09', type: 'Nevroloji' },
    { id: 3, name: 'Paracetamol 500mg', uid: 'RF:M3:2D:87', dose: '3x1', expires: '2027-03', type: 'Ağrıkəsici' },
    { id: 4, name: 'Gabapentin 300mg', uid: 'RF:M4:7E:15', dose: '1x2', expires: '2026-11', type: 'Sinir' },
    { id: 5, name: 'Tramadol 100mg', uid: 'RF:M5:B2:44', dose: '1x1', expires: '2026-06', type: 'Ağrıkəsici' },
    { id: 6, name: 'Amoksisilin 500mg', uid: 'RF:M6:C8:39', dose: '3x1', expires: '2026-09', type: 'Antibiotik' },
];

const ALERTS = [
    { type: 'danger', icon: 'ti-alert-triangle', msg: 'Babayev Tural — İbuprofen tərkibli dərman təyin olundu! Allergiya xəbərdarlığı.', time: '09:14' },
    { type: 'danger', icon: 'ti-clock', msg: 'Həsənova Aytən — 08:00 dərman dozası verilməyib (2 saat gecikmə).', time: '10:02' },
    { type: 'warn', icon: 'ti-map-pin', msg: 'Xəstə Babayev Tural icazə verilən zonadan kənara çıxdı (Təcili Çıxış yanı).', time: '10:31' },
    { type: 'info', icon: 'ti-pill', msg: 'Metoprolol ehtiyat xəbərdarlığı — 8 qutu qalıb, yenilənmə tələb olunur.', time: '10:45' },
];

let selMed = null, nfcLogs = [], mainLogs = [], givenToday = {};

// ─── NAV ───────────────────────────────────────────
function nav(id) {
    document.querySelectorAll('.nav-btn').forEach((b, i) => {
        const ids = ['patients', 'meds', 'track', 'alerts', 'log'];
        b.classList.toggle('active', ids[i] === id);
    });
    document.querySelectorAll('.section').forEach(s => s.classList.toggle('on', s.id === 'sec-' + id));
}

// ─── LOGGING ───────────────────────────────────────
function addLog(msg, type = 'info') {
    mainLogs.unshift({ t: T(), msg, type });
    if (mainLogs.length > 60) mainLogs.pop();
    renderMainLog();
}
function renderMainLog() {
    const el = document.getElementById('main-log');
    el.innerHTML = mainLogs.length ? mainLogs.map(l => `<div class="log-row"><span class="lt">${l.t}</span><span class="lm ${l.type}">${l.msg}</span></div>`).join('') : '<div style="color:var(--text-tertiary);font-size:11px">Qeyd yoxdur</div>';
}
function clearLog() { mainLogs = []; renderMainLog() }

// ─── PATIENTS ──────────────────────────────────────
function renderPatients() {
    const el = document.getElementById('patient-list');
    el.innerHTML = PATIENTS.map(p => `
    <div class="patient-row" onclick="showPatient(${p.id})">
      <div class="avatar" style="background:${p.bg};color:${p.color}">${p.name.split(' ').map(w => w[0]).join('').slice(0, 2)}</div>
      <div class="pinfo">
        <div class="pname">${p.name}</div>
        <div class="psub">${p.uid} · ${p.diagnosis}</div>
      </div>
      <div class="proom">
        <div style="font-size:13px;font-weight:500">${p.room}</div>
        <span class="badge ${p.status === 'Kritik' ? 'badge-red' : p.status === 'İzleme' ? 'badge-amber' : 'badge-green'}">${p.status}</span>
      </div>
    </div>`).join('');

    const sel = document.getElementById('nfc-patient-select');
    sel.innerHTML = PATIENTS.map(p => `<option value="${p.id}">${p.name} — ${p.room}</option>`).join('');
    const sel2 = document.getElementById('med-patient-sel');
    sel2.innerHTML = PATIENTS.map(p => `<option value="${p.id}">${p.name} (${p.diagnosis})</option>`).join('');
}

function showPatient(id) {
    const p = PATIENTS.find(x => x.id === id);
    const meds = p.meds.map(mid => MEDS.find(m => m.id === mid));
    const history = [
        { time: '08:00', text: 'Dərman verildi', sub: meds[0]?.name || '-' },
        { time: '07:30', text: 'NFC giriş — A Blok', sub: 'Tibb bacısı: Anar Rüstəmov' },
        { time: '06:45', text: 'Həyati göstəricilər', sub: 'QT:120/80 · N:72 · T:36.7' },
        { time: 'Dünən 20:00', text: 'Dərman verildi', sub: meds[1]?.name || '-' },
    ];
    document.getElementById('patient-detail').innerHTML = `
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px">
      <div class="avatar" style="background:${p.bg};color:${p.color};width:48px;height:48px;font-size:16px">${p.name.split(' ').map(w => w[0]).join('').slice(0, 2)}</div>
      <div>
        <div style="font-size:15px;font-weight:600">${p.name}</div>
        <div style="font-size:12px;color:var(--text-secondary)">${p.age} yaş · ${p.diagnosis} · ${p.room}</div>
        <code style="font-size:10px;color:var(--text-tertiary);font-family:var(--font-mono)">${p.uid}</code>
      </div>
    </div>
    <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:14px">
      ${p.allergies.length ? p.allergies.map(a => `<span class="badge badge-red"><i class="ti ti-alert-triangle"></i> ${a}</span>`).join('') : '<span class="badge badge-green">Allergiya yoxdur</span>'}
      <span class="badge ${p.status === 'Kritik' ? 'badge-red' : p.status === 'İzləmə' ? 'badge-amber' : 'badge-green'}">${p.status}</span>
    </div>
    <div style="font-size:12px;font-weight:600;color:var(--text-secondary);margin-bottom:8px;text-transform:uppercase;letter-spacing:0.5px">Təyin edilmiş dərmanlar</div>
    ${meds.map(m => `<div style="font-size:12px;padding:6px 0;border-bottom:1px solid var(--border);color:var(--text-primary)">${m.name} — <span style="color:var(--text-secondary)">${m.dose}</span></div>`).join('')}
    <div style="font-size:12px;font-weight:600;color:var(--text-secondary);margin:14px 0 8px;text-transform:uppercase;letter-spacing:0.5px">Keçmiş (RFID/NFC)</div>
    <div class="timeline">${history.map(h => `<div class="tl-item"><div class="tl-time">${h.time}</div><div class="tl-text">${h.text}</div><div class="tl-sub">${h.sub}</div></div>`).join('')}</div>`;
    addLog(`Xəstə faylı açıldı: ${p.name} [${p.uid}]`, 'info');
}

// ─── NFC SCAN ──────────────────────────────────────
function nfcScan() {
    const pid = parseInt(document.getElementById('nfc-patient-select').value);
    const p = PATIENTS.find(x => x.id === pid);
    const zone = document.getElementById('nfc-zone');
    const ico = document.getElementById('nfc-ico');
    const txt = document.getElementById('nfc-txt');
    const sub = document.getElementById('nfc-sub');
    const res = document.getElementById('nfc-result');
    zone.className = 'nfc-zone scanning';
    ico.className = 'nfc-icon amber'; ico.innerHTML = '<i class="ti ti-refresh"></i>';
    txt.textContent = 'NFC bilərzik oxunur...'; sub.textContent = p.uid; res.className = 'result-box';
    setTimeout(() => {
        zone.className = 'nfc-zone ok';
        ico.className = 'nfc-icon green'; ico.innerHTML = '<i class="ti ti-circle-check"></i>';
        txt.textContent = 'Kimlik Doğrulandı'; sub.textContent = `${p.name} · ${p.room}`;
        res.className = 'result-box show ok';
        res.innerHTML = `<i class="ti ti-check"></i> ${p.name} — ${p.room} giriş icazəsi verildi (${p.diagnosis})`;
        const log = `${T()} · ${p.name} [${p.uid}] giriş — ${p.room}`;
        nfcLogs.unshift(log);
        document.getElementById('nfc-log-box').innerHTML = nfcLogs.slice(0, 8).map(l => `<div class="log-row"><span class="lm ok">${l}</span></div>`).join('');
        addLog(`NFC kimlik təsdiqləndi: ${p.name} — ${p.room}`, 'ok');
        setTimeout(() => { zone.className = 'nfc-zone'; ico.className = 'nfc-icon blue'; ico.innerHTML = '<i class="ti ti-wifi"></i>'; txt.textContent = 'Xəstənin NFC bilərziyini oxut'; sub.textContent = 'Klik et → simulyasiya et'; }, 2500);
    }, 1100);
}

// ─── MEDS ──────────────────────────────────────────
function renderMeds() {
    document.getElementById('med-list').innerHTML = MEDS.map(m => `
    <div class="med-card ${selMed === m.id ? 'selected' : ''}" onclick="selectMed(${m.id})">
      <div class="med-name">${m.name}</div>
      <div class="med-uid">${m.uid}</div>
      <div class="med-dose">${m.dose} · ${m.type}</div>
    </div>`).join('');

    const pid = parseInt(document.getElementById('med-patient-sel')?.value) || 1;
    const p = PATIENTS.find(x => x.id === pid);
    document.getElementById('schedule-list').innerHTML = p ? p.meds.map(mid => {
        const m = MEDS.find(x => x.id === mid);
        const done = givenToday[`${pid}-${mid}`];
        return `<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid var(--border);font-size:12px">
      <i class="ti ${done ? 'ti-circle-check' : 'ti-circle'}" style="color:${done ? 'var(--green-text)' : 'var(--text-tertiary)'}"></i>
      <span style="color:var(--text-primary)">${m.name}</span><span style="color:var(--text-secondary);margin-left:auto">${m.dose}</span>
    </div>`;
    }).join('') : '';
}

function selectMed(id) { selMed = id; renderMeds(); resetChecks() }

function resetChecks() {
    ['match', 'dose', 'allergy', 'expire', 'double'].forEach(k => {
        const el = document.getElementById('chk-' + k);
        el.className = 'check-row';
        el.querySelector('i').className = 'ti ti-circle';
        el.querySelector('i').style.color = 'var(--text-tertiary)';
    });
    document.getElementById('med-result').className = 'result-box';
}

function setCheck(id, ok) {
    const el = document.getElementById('chk-' + id);
    el.className = 'check-row ' + (ok ? 'pass' : 'fail') + ' animated';
    const icon = el.querySelector('i');
    icon.className = ok ? 'ti ti-circle-check' : 'ti ti-x';
    icon.style.color = ok ? 'var(--green-text)' : 'var(--red-text)';
}

function verifyMed() {
    if (!selMed) { alert('Əvvəlcə bir dərman seçin'); return }
    const pid = parseInt(document.getElementById('med-patient-sel').value);
    const p = PATIENTS.find(x => x.id === pid);
    const m = MEDS.find(x => x.id === selMed);
    const res = document.getElementById('med-result');
    resetChecks();
    const checks = {};
    setTimeout(() => { checks.match = p.meds.includes(m.id); setCheck('match', checks.match); }, 300);
    setTimeout(() => { checks.dose = true; setCheck('dose', true); }, 600);
    setTimeout(() => {
        const allergyConflict = (m.name.includes('Amoks') && p.allergies.includes('Penisilin')) || (m.name.includes('İbupro') && p.allergies.includes('İbuprofen')) || (m.name.includes('Tramad') && p.allergies.includes('İbuprofen'));
        checks.allergy = !allergyConflict; setCheck('allergy', checks.allergy);
    }, 900);
    setTimeout(() => {
        const exp = new Date(m.expires + '-01'); checks.expire = exp > new Date(); setCheck('expire', checks.expire);
    }, 1200);
    setTimeout(() => {
        const key = `${pid}-${selMed}`; checks.double = !givenToday[key]; setCheck('double', checks.double);
    }, 1500);
    setTimeout(() => {
        const allOk = Object.values(checks).every(Boolean);
        const key = `${pid}-${selMed}`;
        if (allOk) {
            givenToday[key] = true;
            res.className = 'result-box show ok';
            res.innerHTML = `<i class="ti ti-shield-check"></i> Təsdiqləndi — ${m.name} xəstəyə verilə bilər`;
            addLog(`Dərman təsdiqləndi: ${m.name} → ${p.name} [${m.uid}]`, 'ok');
        } else {
            res.className = 'result-box show fail';
            const why = !checks.match ? 'Dərman bu xəstəyə təyin olunmayıb' : !checks.allergy ? 'Allergiya uyğunsuzluğu' : !checks.expire ? 'Son istifadə tarixi keçib' : !checks.double ? 'Bu dərman bu gün artıq verilib' : 'Yoxlama xətası';
            res.innerHTML = `<i class="ti ti-ban"></i> RƏDD EDİLDİ — ${why}`;
            addLog(`Dərman RƏDD EDİLDİ: ${m.name} → ${p.name} — ${why}`, 'err');
        }
        renderMeds();
    }, 1800);
}

// ─── LOCATION MAP ──────────────────────────────────
function renderTrack() {
    const map = document.getElementById('floor-map');
    const rooms = [
        { label: 'A Palatası', x: 5, y: 10, w: 30, h: 35, color: 'rgba(14,165,233,0.12)', stroke: '#0EA5E9' },
        { label: 'B Palatası', x: 38, y: 10, w: 28, h: 35, color: 'rgba(16,185,129,0.12)', stroke: '#10B981' },
        { label: 'C Palatası', x: 69, y: 10, w: 28, h: 35, color: 'rgba(245,158,11,0.12)', stroke: '#F59E0B' },
        { label: 'Əməliyyat', x: 5, y: 55, w: 28, h: 30, color: 'rgba(239,68,68,0.12)', stroke: '#EF4444' },
        { label: 'Aptek', x: 36, y: 55, w: 26, h: 30, color: 'rgba(139,92,246,0.12)', stroke: '#8B5CF6' },
        { label: 'Təcili', x: 65, y: 55, w: 32, h: 30, color: 'rgba(249,115,22,0.12)', stroke: '#F97316' },
    ];
    const dots = [
        { name: 'Kamran Ə.', x: 18, y: 25, color: '#0EA5E9' },
        { name: 'Aytən H.', x: 52, y: 22, color: '#10B981' },
        { name: 'Tural B.', x: 80, y: 62, color: '#EF4444' },
        { name: 'Nigar Q.', x: 15, y: 63, color: '#8B5CF6' },
        { name: 'Narkoz cih.', x: 40, y: 65, color: '#F59E0B', shape: 'square' },
    ];
    map.innerHTML = `<svg width="100%" height="100%" viewBox="0 0 100 90" preserveAspectRatio="xMidYMid meet">
    ${rooms.map(r => `<rect x="${r.x}" y="${r.y}" width="${r.w}" height="${r.h}" rx="2" fill="${r.color}" stroke="${r.stroke}" stroke-width="0.5" stroke-dasharray="2,1"/>
    <text x="${r.x + r.w / 2}" y="${r.y + 6}" text-anchor="middle" font-size="3.5" fill="${r.stroke}" font-family="Inter,sans-serif" font-weight="600">${r.label}</text>`).join('')}
    ${dots.map(d => d.shape === 'square' ?
        `<rect x="${d.x - 2.5}" y="${d.y - 2.5}" width="5" height="5" rx="1" fill="${d.color}" opacity="0.9"/>
       <text x="${d.x}" y="${d.y + 7}" text-anchor="middle" font-size="2.8" fill="${d.color}" font-family="Inter,sans-serif">${d.name}</text>` :
        `<circle cx="${d.x}" cy="${d.y}" r="3" fill="${d.color}" opacity="0.9" stroke="${d.color}" stroke-width="0.3" stroke-opacity="0.4">
        <animate attributeName="r" values="3;4;3" dur="2s" repeatCount="indefinite"/>
        <animate attributeName="opacity" values="0.9;0.5;0.9" dur="2s" repeatCount="indefinite"/>
      </circle>
       <text x="${d.x}" y="${d.y + 6}" text-anchor="middle" font-size="2.8" fill="${d.color}" font-family="Inter,sans-serif" font-weight="500">${d.name}</text>`).join('')}
    <text x="50" y="88" text-anchor="middle" font-size="2.5" fill="#64748B" font-family="Inter,sans-serif">● Xəstə  ■ Avadanlıq — RFID Anten Şəbəkəsi</text>
  </svg>`;

    document.getElementById('location-cards').innerHTML = PATIENTS.map(p => `
    <div class="stat">
      <div style="font-size:12px;font-weight:600;color:var(--text-primary)">${p.name.split(' ')[0]}</div>
      <div style="font-size:11px;color:var(--text-secondary)">${p.room}</div>
      <div style="font-size:10px;color:${p.status === 'Kritik' ? 'var(--red-text)' : p.status === 'İzleme' ? 'var(--amber-text)' : 'var(--green-text)'};margin-top:3px;font-weight:500">${p.status}</div>
    </div>`).join('');
}

// ─── ALERTS ────────────────────────────────────────
function renderAlerts() {
    document.getElementById('alerts-list').innerHTML = ALERTS.map(a => `
    <div class="alert-item ${a.type}">
      <i class="ti ${a.icon}" style="font-size:20px;flex-shrink:0;opacity:0.9"></i>
      <div>
        <div style="font-size:13px;color:var(--text-primary);font-weight:500">${a.msg}</div>
        <div style="font-size:11px;color:var(--text-tertiary);margin-top:2px;font-family:var(--font-mono)">${a.time}</div>
      </div>
    </div>`).join('');
}

// ─── INIT ──────────────────────────────────────────
renderPatients(); renderMeds(); renderAlerts(); renderTrack();
addLog('MedTrack sistemi başladıldı — NFC + RFID modulları aktivdir', 'ok');
addLog('12 aktiv xəstə izlənmədədir, 3 kritik siqnal mövcuddur', 'warn');

document.getElementById('med-patient-sel').addEventListener('change', () => { selMed = null; resetChecks(); renderMeds() });

// Expose to global scope for onclick handlers (module scripts are scoped)
window.nav = nav;
window.showPatient = showPatient;
window.nfcScan = nfcScan;
window.selectMed = selectMed;
window.verifyMed = verifyMed;
window.clearLog = clearLog;

