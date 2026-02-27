/**
 * LPR Registration System — v4.6.1
 *
 * Major overhaul replacing v2.5.23 architecture:
 *
 * ─── Core Engine / Logic ────────────────────────────────────────────────
 * • Full context-based session engine (ctx_date, ctx_jour, ctx_cours)
 * • Universal backdating: ISSUE, SPEND, ATTEND, DROP-IN all respect ctx_date
 * • Unified audit log: AuditAt = real timestamp; Date = session context
 * • Duplicate-safe execution via submission_id (idempotency)
 * • Validity logic refined for cards & trimestres
 * • Drop-in supports multi-class AND correct expected-price accounting
 *
 * ─── UI / UX Overhaul (Patch 4) ──────────────────────────────────────────
 * • New sticky mini-bar (context chip + reset + staff identity)
 * • Main header no longer sticky → larger working area for all devices
 * • Search filters inline (Members · PreReg · Contacts)
 * • Doorlist readability: zebra, spacing, partial/highlight, explicit OK — Cours
 * • French sorting + naming: NOM Prénom, sorted by NOM
 * • Smart button enabling (doorlist + search)
 * • Auto-prefill highlighting for tonight classes in Issue forms
 *
 * ─── Data / Mapping / Search ────────────────────────────────────────────
 * • Data-driven Jour↔Cours mapping from Settings sheet (A for jours, D:E for cours)
 * • Search: compact, real filters, 5-item cap + “Voir plus”
 * • Student mode: clean card/trim display
 *
 * ─── Admin Features ─────────────────────────────────────────────────────
 * • Recap report (jour + cours + date) made admin-only
 * • Recap fully independent of context (ctx_date unaffected)
 *
 * ─── Automation / Emails ────────────────────────────────────────────────
 * • Email on card usage (single SPEND1, multi-class CHECKIN_MULTI)
 *
 * Maintainer: Shehan
 * Release date: 2025-12-02
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('LPR Registration')
    .addItem('Generate Staff Tokens', 'generateStaffTokens_')
    .addItem('Rebuild Student URLs (Members)', 'rebuildStudentUrls_')
    .addToUi();
}

const CFG = {
  SHEET: {
    SETTINGS: 'Settings',
    MEMBERS: 'Members',
    LOGS: 'Logs',
    CONTACTS: 'Contacts',
    PREREG: 'PreReg',
    PAYMENTS: 'Payments',
    TRIMS: 'Trimestres',        // Code | Nom affiché | Début | Fin
  },
  ID_PREFIX: 'DW-',
  WEBAPP_BASE: 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec',
  EMAIL_ON_ISSUE: true,
  FROM_NAME: 'FromName',
  SIGN_NAME: 'SignName',
  CC_EMAIL: '',
  PRICE_DROPIN: 12,
  PRICE_CARD_6: 70,
  PRICE_CARD_12: 135,
  // Keys are exact labels shown in the Forfait select
  TRIM_LABELS: [
    '1 cours / semaine – 120€',
    'Pack 2 styles (Lindy Hop + Blues) – 200 €',
    '2 cours / semaine – 215€',
    '3 cours / semaine – 275€',
    'Cours illimités – 300€'
  ],
  TRIM_PRICES: {
    '1 cours / semaine – 120€': 120,
    'Pack 2 styles (Lindy Hop + Blues) – 200 €': 200,
    '2 cours / semaine – 215€': 215,
    '3 cours / semaine – 275€': 275,
    'Cours illimités – 300€': 300
  },
};

function ss_(){ return SpreadsheetApp.getActive(); }
function sh_(n){ return ss_().getSheetByName(n); }
function todayMid_(){ const d=new Date(); d.setHours(0,0,0,0); return d; }
function norm_(s){ return (s||'').toString().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim(); }
function token_(len=28){ const c='abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'; let o=''; for (let i=0;i<len;i++) o+=c[Math.floor(Math.random()*c.length)]; return o; }
function computeLeft_(total, used){ total=+total||0; used=+used||0; return Math.max(total-used,0); }
function isExpired_(d){ if(!d) return false; const a=todayMid_().getTime(); const b=new Date(d); b.setHours(0,0,0,0); return a>b.getTime(); }

function fmtFR_(d){
  if (!d) return '';
  var date = (d instanceof Date) ? new Date(d) : new Date(d);
  if (isNaN(date)) return '';

  var day = date.getDate(); // 1–31
  var months = ['janv.','févr.','mars','avr.','mai','juin','juil.','août','sept.','oct.','nov.','déc.'];
  var m = months[date.getMonth()];
  var y2 = String(date.getFullYear()).slice(-2); // '25'

  // 24 févr. ’25  (French apostrophe)
  return day + ' ' + m + ' \u2019' + y2;
}

function getNameFromRow_(row){ return (row[2]||'')+' '+(row[1]||''); }
function amountOrFallback_(raw, fallback){
  // Treat "" or null/undefined as "no input" → use fallback.
  // Treat "0" as a valid number 0.
  if (raw === '' || raw === null || raw === undefined) return +fallback || 0;
  const n = parseFloat(String(raw).replace(',', '.'));
  return isNaN(n) ? (+fallback || 0) : n;
}

// Format name for emails: LASTNAME in ALL CAPS, first names normal
function formatNameEmail_(name){
  const raw = String(name || '').trim();
  if (!raw) return '';
  const parts = raw.split(/\s+/);
  if (parts.length === 1) {
    return raw.toUpperCase();
  }
  const last  = parts[0].toUpperCase();
  const first = parts.slice(1).join(' ');
  return last + ' ' + first;
}


// Parse HTML date input (YYYY-MM-DD) to midnight Date
function parseYMD_(s){
  if (!s) return null;
  const parts = String(s).split('-');
  if (parts.length !== 3) return null;
  const y = +parts[0], m = +parts[1]-1, d = +parts[2];
  if (!y || m<0 || m>11 || !d) return null;
  const dt = new Date(y,m,d);
  dt.setHours(0,0,0,0);
  return dt;
}

// Normalize forfait labels (handles NBSP, dash variants, double spaces)
function normForfaitKey_(s){
  return String(s||'')
    .replace(/\u00A0/g, ' ')     // NBSP -> space
    .replace(/[–—]/g, '-')      // long dashes -> hyphen
    .replace(/\s+/g, ' ')       // collapse spaces
    .replace(/\s*€\s*/g, '€')   // normalize euro spacing
    .trim()
    .toLowerCase();
}

// Map Date -> French jour label used in Settings
function jourFromDate_(dateObj){
  const days = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
  return days[new Date(dateObj).getDay()];
}

// French weekday label from a Date object
function labelJourFR_(d){
  if (!d) return '';
  const days = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
  return days[d.getDay()];
}

// Duplicate check-ins for an arbitrary date
function alreadyCheckedOnDate_(memberId, action, jour, cours, dateObj){
  const s=sh_(CFG.SHEET.LOGS); const last=s.getLastRow(); if(last<2) return false;
  const t0=new Date(dateObj); t0.setHours(0,0,0,0);
  const tms=t0.getTime();
  const vals=s.getRange(2,1,last-1,14).getValues();
  for (let i=0;i<vals.length;i++){
    const v=vals[i];
    const dt=v[0]; if(!dt) continue;
    const d=new Date(dt); d.setHours(0,0,0,0);
    if(d.getTime()!==tms) continue;
    if(v[1]!==memberId) continue;
    if(v[3]!==action) continue;
    if((v[9]||'')!==jour) continue;
    if((v[10]||'')!==cours) continue;
    return true;
  }
  return false;
}

function readSettings_(){
  const s = sh_(CFG.SHEET.SETTINGS);
  if(!s) throw new Error('Settings sheet missing');

  // Jours (A)
  const jours=[];
  for (let r=2; r<=s.getLastRow(); r++){
    const v=s.getRange(r,1).getValue();
    if(!v) break;
    jours.push(String(v));
  }

  // Cours table (D:E) => mapping by jour
  const mapping = {}; // { "Mardi": ["Mardi – Nature (Sud)", ...], ... }
  for (let r=2; r<=s.getLastRow(); r++){
    const cours = s.getRange(r,4).getValue(); // D
    const jour  = s.getRange(r,5).getValue(); // E
    if (!cours || !jour) continue;
    const J = String(jour).trim();
    const C = String(cours).trim();
    if (!mapping[J]) mapping[J] = [];
    mapping[J].push(C);
  }

  // Équipe (G:J)
  const team=[];
  for (let r=2; r<=s.getLastRow(); r++){
    const name=s.getRange(r,7).getValue();
    if(!name) break;
    team.push({
      row:r, name:String(name),
      token:String(s.getRange(r,8).getValue()||''),
      pin:String(s.getRange(r,9).getValue()||''),
      active:String(s.getRange(r,10).getValue()||'Oui').toLowerCase().startsWith('o')
    });
  }

  // Trimestres sheet (optional)
  const trimsSheet = sh_(CFG.SHEET.TRIMS);
  const trims=[];
  if (trimsSheet){
    const last=trimsSheet.getLastRow();
    for (let r=2; r<=last; r++){
      const code = String(trimsSheet.getRange(r,1).getValue()||'').trim();
      if(!code) break;
      trims.push({
        code,
        label: String(trimsSheet.getRange(r,2).getValue()||'').trim(),
        start: trimsSheet.getRange(r,3).getValue()||'',
        end:   trimsSheet.getRange(r,4).getValue()||'',
      });
    }
  }

  return { jours, mapping, team, trims, settingsSheet:s };
}
function generateStaffTokens_(){
  const { team, settingsSheet } = readSettings_();
  let made=0;
  team.forEach(t=>{
    if(t.name && !t.token){
      settingsSheet.getRange(t.row,8).setValue(token_());
      made++;
    }
  });
  SpreadsheetApp.getUi().alert('Generated tokens for '+made+' team member(s).');
}

// -------- Payments summary helpers --------
function sumPaidForId_(id){
  if (!id) return 0;
  const s = sh_(CFG.SHEET.PAYMENTS); if(!s) return 0;
  const last = s.getLastRow(); if (last < 2) return 0;
  const rows = s.getRange(2,1,last-1,8).getValues(); // A:H (Date, Montant, Mode, Ref, Payeur, Liés à, Saisi par, Note)
  let sum = 0;
  for (const r of rows){
    const amount = +r[1] || 0;
    const linked = String(r[5]||'');
    if (!linked) continue;
    const ids = linked.split(',').map(x=>x.trim()).filter(Boolean);
    if (ids.includes(id)) sum += amount;
  }
  return sum;
}

function expectedDropinForId_(memberId){
  const s = sh_(CFG.SHEET.LOGS);
  const lastRow = s.getLastRow();
  if (lastRow < 2) return 0;

  const lastCol = s.getLastColumn();
  const header = s.getRange(1,1,1,lastCol).getValues()[0]
    .map(h => String(h||'').toLowerCase().trim());

  // Try to locate columns by header name, fallback if not found
  const idxAction =
    header.indexOf('action') !== -1 ? header.indexOf('action') :
    header.indexOf('actions') !== -1 ? header.indexOf('actions') :
    3; // fallback ≈ col D

  const idxMember =
    header.indexOf('member_id') !== -1 ? header.indexOf('member_id') :
    header.indexOf('membre_id') !== -1 ? header.indexOf('membre_id') :
    header.indexOf('id') !== -1 ? header.indexOf('id') :
    1; // fallback ≈ col B/C

  const idxAmount =
    header.indexOf('amount') !== -1 ? header.indexOf('amount') :
    header.indexOf('montant') !== -1 ? header.indexOf('montant') :
    -1;

  const vals = s.getRange(2,1,lastRow-1,lastCol).getValues();

  let count = 0;
  let sum = 0;

  vals.forEach(r=>{
    const act = String(r[idxAction]||'').toUpperCase();
    const mid = String(r[idxMember]||'');
    if (act === 'DROPIN' && mid === String(memberId)){
      count++;
      if (idxAmount >= 0){
        sum += Number(r[idxAmount] || 0);
      }
    }
  });

  // If amount column exists, trust it. Otherwise count × price.
  return (idxAmount >= 0) ? sum : (count * CFG.PRICE_DROPIN);
}


// Expected “price” of an offer, for Solde calculation
function expectedForMemberRow_(row){
  const type = String(row[15]||'');      // Type d’offre
  const tl = type.toLowerCase();

  if (tl.startsWith('trimestre')){
    const forfaitRaw = String(row[17]||'').trim();  // R = forfait trimestre
    if (!forfaitRaw) return 0;

    if (CFG.TRIM_PRICES[forfaitRaw] != null) return +CFG.TRIM_PRICES[forfaitRaw] || 0;

    const nf = normForfaitKey_(forfaitRaw);
    for (const k in CFG.TRIM_PRICES){
      if (normForfaitKey_(k) === nf) return +CFG.TRIM_PRICES[k] || 0;
    }
    return 0;
  }

  if (tl.startsWith('carte')){
    if (/12/.test(type)) return CFG.PRICE_CARD_12;
    if (/6/.test(type))  return CFG.PRICE_CARD_6;
    return 0;
  }

  if (tl.startsWith('drop-in')) {
    const id = String(row[0] || ''); // Members col A = member id
    return expectedDropinForId_(id);
  }

  return 0;
}

function getValidUntilForId_(id){
  const it = findMemberById_(id);
  if (!it.row) return '';
  const until = it.data[8];   // Column I = "Valide jusqu’au"
  return until ? fmtFR_(until) : '';
}


// Members helpers
function findMemberById_(id){
  const s=sh_(CFG.SHEET.MEMBERS); if(!s) return {row:null};
  const last=s.getLastRow(); if(last<2) return {row:null};
  const ids=s.getRange(2,1,last-1,1).getValues().flat();
  const i=ids.indexOf(id); if(i===-1) return {row:null};
  return { row:i+2, data:s.getRange(i+2,1,1,21).getValues()[0], sheet:s }; // through U
}
function findActiveCardByEmail_(email){
  const s = sh_(CFG.SHEET.MEMBERS); if(!s) return null;
  const last=s.getLastRow(); if(last<2) return null;
  const vals = s.getRange(2,1,last-1,21).getValues();
  const today = todayMid_().getTime();
  for (let r=0; r<vals.length; r++){
    const row = vals[r];
    const emailRow = String(row[4]||'').toLowerCase();
    if (emailRow && emailRow === String(email).toLowerCase()){
      const type = String(row[15]||'');         // Type d’offre
      const total = +row[9]||0;                 // total
      const used  = +row[10]||0;                // used
      const left  = Math.max(total-used,0);
      const vu    = row[8];                     // until
      const notExpired = vu ? (new Date(vu).setHours(0,0,0,0) >= today) : true;
      if (type.toLowerCase().startsWith('carte') && left>0 && notExpired){
        return { rowIndex:r+2, id:row[0], name:getNameFromRow_(row), left, total, until:vu };
      }
    }
  }
  return null;
}
function canRenewCard_(rowData){
  const type = String(rowData[15]||'').toLowerCase();
  if(!type.startsWith('carte')) return false;
  const total=+rowData[9]||0, used=+rowData[10]||0, left=Math.max(total-used,0);
  const vu=rowData[8];
  return left===0 || isExpired_(vu);
}

// Trimestres: prevent duplicate active registration for the same trimester code
function findActiveTrimByEmailAndCode_(email, trimCode){
  const s = sh_(CFG.SHEET.MEMBERS); if(!s) return null;
  const last = s.getLastRow(); if (last < 2) return null;
  const vals = s.getRange(2,1,last-1,21).getValues();
  const e = String(email||'').toLowerCase();
  const code = String(trimCode||'').trim();
  for (const row of vals){
    const type = String(row[15]||'').toLowerCase();     // Type d’offre
    const em   = String(row[4]||'').toLowerCase();      // Email
    const c    = String(row[20]||'').trim();            // Code Trimestre
    if (type.startsWith('trimestre') && em && e && em===e && c && code && c===code){
      return { id: row[0], name: getNameFromRow_(row) };
    }
  }
  return null;
}

// Search (Members + PreReg + Contacts)
function extractIdFromUrlish_(q){
  const m = String(q||'').match(/(?:\bid=)(DW-\d+)/i);
  return m ? m[1] : null;
}
function performSearch_(q, filters, limit, includeMore){
  const out = [];
  const idFromUrl = extractIdFromUrlish_(q);
  const term = norm_(q);
  const wantMembers  = !!filters.members;
  const wantPreReg   = !!filters.prereg;
  const wantContacts = !!filters.contacts;

  if (wantMembers){
    const s=sh_(CFG.SHEET.MEMBERS); if(s){
      const last=s.getLastRow(); if(last>=2){
        const vals=s.getRange(2,1,last-1,21).getValues();
        vals.forEach((v)=>{
          const [id,nom,prenom,tel,email,, paid, months, vu, total, used,, tok, url,, typeOffre, , forfaitTrim, coursIncl, , trimCode] = v;
          let match=false;
          if(idFromUrl && id===idFromUrl) match=true;
          const hay = [id,nom,prenom,tel,email].map(norm_).join(' ');
          const termDigits = term.replace(/\D/g,'');
          const hayDigits  = String(tel||'').replace(/\D/g,'');

          if(
            !term ||
            hay.includes(term) ||
            (termDigits && hayDigits && hayDigits.includes(termDigits))
          ) match=true;
          if(match){
          const exp = expectedForMemberRow_(v);
          const paidSum = id ? sumPaidForId_(String(id)) : 0;

          // Type d’offre (ex: "Trimestre", "Carte 6 cours…", "Drop-in")
          const typeOffreLow = String(typeOffre||'').toLowerCase();

          let owed;
          if (typeOffreLow.startsWith('drop-in')) {
            // Pour les drop-ins, on veut TOUJOURS le vrai solde, même si exp = 0
            // (0 cours → exp=0 → solde = -montant payé = avoir complet, ce qui est logique).
            owed = exp - paidSum;
          } else {
            // Pour trimestres / cartes :
            // si exp = 0 (forfait inconnu, anciennes lignes), on masque le solde.
            owed = exp > 0 ? (exp - paidSum) : 0;
          }

            out.push({
              source:'Members', id:String(id||''), nom:String(nom||''), prenom:String(prenom||''),
              email:String(email||''), tel:String(tel||''), typeOffre:String(typeOffre||''),
              total:+(total||0), used:+(used||0), left:computeLeft_(total,used),
              until: vu? new Date(vu) : null, expired: isExpired_(vu), url:String(url||''),
              forfaitTrim:String(forfaitTrim||''), coursIncl:String(coursIncl||''), trimCode:String(trimCode||''),
              expected: exp||0, paid: paidSum||0, owed: owed||0
            });
          }
        });
      }
    }
  }

  if (wantPreReg){
    const s=sh_(CFG.SHEET.PREREG); if(s){
      const last=s.getLastRow(); if(last>=2){
        const vals=s.getRange(2,1,last-1,13).getValues(); // A..M
        vals.forEach((v)=>{
          const nom=v[2], prenom=v[3], tel=v[4], emailAddr=v[1];
          const typeInscr=v[6], forfaitTrim=v[7], typeCarte=v[8], choixCours=v[9];
          const hay=[nom,prenom,emailAddr,tel,(typeInscr||''),(forfaitTrim||''),(typeCarte||'')].map(norm_).join(' ');
          if(!term || hay.includes(term)){
            out.push({
              source:'PreReg', id:'', nom:String(nom||''), prenom:String(prenom||''), email:String(emailAddr||''), tel:String(tel||''),
              prereg:{ typeInscr, forfaitTrim, typeCarte, choixCours }
            });
          }
        });
      }
    }
  }

  if (wantContacts){
    const s=sh_(CFG.SHEET.CONTACTS); if(s){
      const last=s.getLastRow(); if(last>=2){
        const vals=s.getRange(2,1,last-1,8).getValues();
        vals.forEach((v)=>{
          const [nom,prenom,email,tel] = [v[0],v[1],v[2],v[3]];
          const hay=[nom,prenom,email,tel].map(norm_).join(' ');
          if(!term || hay.includes(term)){
            out.push({ source:'Contacts', id:'', nom:String(nom||''), prenom:String(prenom||''), email:String(email||''), tel:String(tel||'') });
          }
        });
      }
    }
  }

  const order = {Members:0, PreReg:1, Contacts:2};
  out.sort((a,b)=> (order[a.source]-order[b.source]));
  if (!includeMore && out.length>limit) return { shown: out.slice(0,limit), more: out.length-limit, allCount: out.length };
  return { shown: out, more: 0, allCount: out.length };
}

// Logging
function log_(o){
  const s = sh_(CFG.SHEET.LOGS);
  if (!s) throw new Error('Logs sheet missing');
  s.appendRow([
    o.timestamp || new Date(),   // A: Session date (backdated)
    o.member_id || '',           // B
    o.name || '',                // C
    o.action || '',              // D
    o.amount || '',              // E
    o.classes_left_after || '',  // F
    o.by || '',                  // G
    o.type_offre || '',          // H
    o.mode_paiement || '',       // I
    o.jour || '',                // J
    o.cours || '',               // K
    o.classe_libre || '',        // L
    o.email || '',               // M
    o.note || '',                // N
    o.submission_id || '',       // O
    new Date()                   // P: AuditAt (real time of entry)
  ]);
}

// Anti double-submit: check if a submission_id already exists in Logs!O
function isDuplicateSubmission_(submissionId){
  if (!submissionId) return false;
  const s = sh_(CFG.SHEET.LOGS);
  if (!s) return false;
  const last = s.getLastRow();
  if (last < 2) return false;

  const vals = s.getRange(2, 15, last - 1, 1).getValues();
  const sid = String(submissionId);
  for (let i = 0; i < vals.length; i++){
    if (String(vals[i][0] || '') === sid) return true;
  }
  return false;
}

// Submission de-duplication (Payments-only)
function hasSubmissionBeenHandled_(submissionId){
  if (!submissionId) return false;
  const s = sh_(CFG.SHEET.PAYMENTS);
  if (!s) return false;
  const last = s.getLastRow();
  if (last < 2) return false;
  const ids = s.getRange(2, 9, last - 1, 1).getValues().flat(); // col I = 9th
  return ids.includes(submissionId);
}


// Payments
function ensurePaymentsSheet_(){
  const s = sh_(CFG.SHEET.PAYMENTS);
  if (!s) throw new Error('Payments sheet missing. Create a tab named "Payments" with headers.');
}

// Append one payment row, including optional SubmissionID for de-duplication.
function logPayment_(p){
  const s = sh_(CFG.SHEET.PAYMENTS);
  s.appendRow([
    new Date(),                // A: Date
    +p.amount || 0,            // B: Montant
    p.mode || '',              // C: Mode
    p.ref || '',               // D: Réf.
    p.payer || '',             // E: Payeur
    p.linked || '',            // F: Liés à
    p.by || '',                // G: Saisi par
    p.note || '',              // H: Note
    p.submissionId || ''       // I: SubmissionID
  ]);
}

// Duplicate check-ins (same person / same day / same class)
function alreadyCheckedToday_(memberId, action, jour, cours){
  return alreadyCheckedOnDate_(memberId, action, jour, cours, todayMid_());
}

// Expected trimestres for an arbitrary date
function getExpectedTrimsOnDate_(jour, cours, dateObj){
  const s=sh_(CFG.SHEET.MEMBERS); if(!s) return [];
  const last = s.getLastRow(); if (last < 2) return [];
  const vals = s.getRange(2,1,last-1,21).getValues();
  const settings = readSettings_();
  const t = new Date(dateObj); t.setHours(0,0,0,0);
  const out=[];
  for (const v of vals){
    const type = String(v[15]||'').toLowerCase();
    if (!type.startsWith('trimestre')) continue;
    const name = getNameFromRow_(v);
    const memberId = v[0];
    const coursIncl = String(v[18]||''); // col S = index 18
    if (!coursIncl || coursIncl.indexOf(cours)===-1) continue;
    const code = String(v[20]||'');
    let inWindow = true;
    if (code){
      const f = settings.trims.find(x=>x.code===code);
      if (f && f.start && f.end){
        const sdt = new Date(f.start); sdt.setHours(0,0,0,0);
        const edt = new Date(f.end);   edt.setHours(0,0,0,0);
        inWindow = (t>=sdt && t<=edt);
      }
    }
    if (inWindow) out.push({ memberId, name });
  }
  return out;
}

// Expected trimestres for TODAY (kept for existing behaviour)
function getExpectedTrimsToday_(jour, cours){
  const today = todayMid_();
  return getExpectedTrimsOnDate_(jour, cours, today);
}

// “Voir la classe (aujourd’hui)”
function getClassToday_(jour, cours){
  if(!jour || !cours) return {present:[], totals:{}, expected:[], absent:[]};
  const s=sh_(CFG.SHEET.LOGS); const last=s.getLastRow(); 
  const t0=todayMid_().getTime();
  const vals = last>=2 ? s.getRange(2,1,last-1,14).getValues() : [];
  const present=[], totals={dropin:0, cartes:0, trims:0, total:0};
  vals.forEach(v=>{
    const dt=v[0]; if(!dt) return;
    const d=new Date(dt); d.setHours(0,0,0,0);
    if(d.getTime()!==t0) return;
    if((v[9]||'')!==jour) return;
    if((v[10]||'')!==cours) return;
    const action=v[3]||'';
    const name=v[2]||'';
    const type=v[7]||'';
    present.push({name, action, type, member_id:v[1]});
    totals.total++;
    if(action==='DROPIN') totals.dropin++;
    if(action==='SPEND') totals.cartes++;
    if(action==='ATTEND') totals.trims++;
  });
  const expected = getExpectedTrimsToday_(jour, cours);
  const presentIds = new Set(present.map(x=>x.member_id));
  const absent = expected.filter(x=> !presentIds.has(x.memberId));
  return {present, totals, expected, absent};
}

// NEW: “Voir la classe” for an arbitrary date (reports)
function getClassSession_(dateStr, jour, cours){
  if(!jour || !cours || !dateStr) return {present:[], totals:{}, expected:[], absent:[]};
  const dateObj = parseYMD_(dateStr);
  if (!dateObj) return {present:[], totals:{}, expected:[], absent:[]};

  const s=sh_(CFG.SHEET.LOGS); const last=s.getLastRow();
  if (last < 2) return {present:[], totals:{}, expected:[], absent:[]};
  const t0 = dateObj.getTime();
  const vals = s.getRange(2,1,last-1,14).getValues();
  const present=[], totals={dropin:0, cartes:0, trims:0, total:0};
  vals.forEach(v=>{
    const dt=v[0]; if(!dt) return;
    const d=new Date(dt); d.setHours(0,0,0,0);
    if(d.getTime()!==t0) return;
    if((v[9]||'')!==jour) return;
    if((v[10]||'')!==cours) return;
    const action=v[3]||'';
    const name=v[2]||'';
    const type=v[7]||'';
    present.push({name, action, type, member_id:v[1]});
    totals.total++;
    if(action==='DROPIN') totals.dropin++;
    if(action==='SPEND') totals.cartes++;
    if(action==='ATTEND') totals.trims++;
  });

  const expected = getExpectedTrimsOnDate_(jour, cours, dateObj);
  const presentIds = new Set(present.map(x=>x.member_id));
  const absent = expected.filter(x=> !presentIds.has(x.memberId));
  return {present, totals, expected, absent};
}

// NEW: door checklist data for TODAY, by jour (all classes on that day)
function buildDoorListOnDate_(jour, dateObj) {
  const dObj = dateObj || todayMid_();

  // Read mapping + trims from settings
  const st = readSettings_();
  const mapping = st.mapping || {};
  const trims   = st.trims   || [];

  // Classes offered on this jour
  const classesTonight = (mapping[jour] || []).slice();
  if (!classesTonight.length) {
    return { classes: [], rows: [] };
  }

  const s = sh_(CFG.SHEET.MEMBERS);
  const lastRow = s.getLastRow();
  if (lastRow < 2) {
    return { classes: classesTonight, rows: [] };
  }

  const lastCol = s.getLastColumn();
  const data = s.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Logs map: which classes already attended / spent for this jour + date
  const actionsMap = getActionsByMemberOnDate_(jour, dObj); // { memberId: ['Cours A','Cours B', ...] }

  const rows = [];

  data.forEach(function (row) {
    const id = String(row[0] || '');
    if (!id) return;

    const nom    = row[1] || '';
    const prenom = row[2] || '';
    const tel    = row[3] || '';
    const email  = row[4] || '';

    const typeOffreRaw = String(row[15] || '').trim(); // P: Type d’offre
    if (!typeOffreRaw) return;

    const lowType = typeOffreRaw.toLowerCase();
    const isCard  = lowType.indexOf('carte') === 0;
    const isTrim  = lowType.indexOf('trimestre') === 0;

    // Only cards + trimestres are relevant for the door list
    if (!isCard && !isTrim) return;

    // Common fields
    const validUntil = row[8] || '';          // I: Valide jusqu’au
    const total      = +row[9]  || 0;         // J: Total cours
    const used       = +row[10] || 0;         // K: Utilisés
    const left       = computeLeft_(total, used);
    const expired    = isExpired_(validUntil);

    const forfaitTrim = isTrim ? String(row[17] || '') : ''; // R: Forfait
    const rawList     = String(row[18] || '');                // S: Cours inclus / préférés
    const listArr = rawList
      ? rawList.split(',').map(function (s) { return s.trim(); }).filter(Boolean)
      : [];

    // ── Trimestre window (based on trimCode + settings.trims) ──
    let inTrimWindow = true;
    let trimCode     = '';

    if (isTrim) {
      trimCode = String(row[20] || ''); // U: Code trimestre
      if (trimCode) {
        const found = trims.find(function (t) {
          return String(t.code || '') === trimCode;
        });
        if (found && found.start && found.end) {
          const start = found.start;
          const end   = found.end;
          inTrimWindow = (dObj >= start && dObj <= end);
        }
      }
    }

    // ─────────────────────────────────
    // Row-level inclusion rules
    // ─────────────────────────────────
    if (isCard) {
      // Preferred classes are ONLY for inclusion in the list:
      // if they have at least one preferred class that is offered tonight, they appear.
      const hasPrefTonight = listArr.some(function (c) {
        return classesTonight.indexOf(c) !== -1;
      });

      if (!hasPrefTonight) return;  // do not show this card at all
      if (expired) return;          // expired card → skip row

      // NEW RULE:
      // - If card has 0 left and has NOT been used today → hide from door list.
      // - If it has 0 left but WAS used today → keep it on the list (for tonight’s record).
      const actionsTodayForMember = actionsMap[id] || [];   // classes used for this jour + date
      const usedToday = actionsTodayForMember.length > 0;

      if (left <= 0 && !usedToday) {
        return; // card was already empty before today → don't show
      }
    }

    if (isTrim) {
      // Must have at least one included class that is offered tonight
      const hasIncludedTonight = listArr.some(function (c) {
        return classesTonight.indexOf(c) !== -1;
      });
      if (!hasIncludedTonight) return;
      if (!inTrimWindow) return;
    }

    const alreadyList = actionsMap[id] || [];

    const classCells = [];
    let eligibleCount = 0;
    let doneCount     = 0;

    // ─────────────────────────────────
    // Per-class cells
    // ─────────────────────────────────
    classesTonight.forEach(function (coursName) {
      const already = alreadyList.indexOf(coursName) !== -1;
      let canAttend = false;
      let canSpend  = false;

      if (isTrim) {
        // Trimestre: ONLY classes included at issuing time
        const isIncluded = listArr.indexOf(coursName) !== -1;
        if (isIncluded && !already && inTrimWindow) {
          canAttend = true;
        }
      }

      if (isCard) {
        // CARD RULE:
        // Once they are on the list, ANY class offered that soir is spendable
        // as long as the card is valid and they still have classes left.
        if (!expired && left > 0 && !already) {
          canSpend = true;
        }
      }

      if (canAttend || canSpend) {
        eligibleCount++;
      }
      if (already) {
        doneCount++;
      }

      classCells.push({
        cours: coursName,
        already: already,
        canAttend: canAttend,
        canSpend: canSpend
      });
    });

    // Solde (expected vs paid, consistent with search chips)
    const expected = expectedForMemberRow_(row);
    const paid     = sumPaidForId_(id);
    const owed     = expected - paid;

    const r = {
      id: id,
      nom: nom,
      prenom: prenom,
      email: email,
      tel: tel,

      typeOffre: typeOffreRaw,
      isCard: isCard,
      isTrim: isTrim,

      forfaitTrim: forfaitTrim,
      coursIncl: isTrim ? rawList : '',

      // Card balance info (only meaningful for cards, but harmless otherwise)
      left: isCard ? left : 0,
      total: isCard ? total : 0,
      until: isCard ? validUntil : '',
      expired: isCard ? expired : false,

      owed: owed,

      classCells: classCells,
      eligibleCount: eligibleCount,
      doneCount: doneCount,
      allDone: (eligibleCount > 0 && doneCount === eligibleCount),
      partiallyDone: (doneCount > 0 && doneCount < eligibleCount)
    };

    rows.push(r);
  });
  
  // Sort door list rows by NOM (surname), then Prénom, using French collation
  rows.sort(function(a, b) {
    const an = String(a.nom || '').toUpperCase();
    const bn = String(b.nom || '').toUpperCase();
    const lastCmp = an.localeCompare(bn, 'fr', { sensitivity: 'base' });
    if (lastCmp !== 0) return lastCmp;

    const ap = String(a.prenom || '');
    const bp = String(b.prenom || '');
    return ap.localeCompare(bp, 'fr', { sensitivity: 'base' });
  });

  return {
    classes: classesTonight,
    rows: rows
  };
}


function buildDoorListToday_(jour){
  return buildDoorListOnDate_(jour, todayMid_());
}

// Actions already done on a given date/jour, grouped by member
function getActionsByMemberOnDate_(jour, dateObj){
  const out = {};
  const s = sh_(CFG.SHEET.LOGS);
  if (!s) return out;
  const last = s.getLastRow();
  if (last < 2) return out;

  const t0 = new Date(dateObj); t0.setHours(0,0,0,0);
  const tms = t0.getTime();
  const vals = s.getRange(2,1,last-1,14).getValues();

  vals.forEach(v=>{
    const dt = v[0]; if(!dt) return;
    const d = new Date(dt); d.setHours(0,0,0,0);
    if (d.getTime() !== tms) return;
    if ((v[9]||'') !== jour) return;

    const memberId = String(v[1]||'');
    const action   = String(v[3]||'');
    const cours    = String(v[10]||'');
    if (!memberId || !cours) return;
    if (action !== 'ATTEND' && action !== 'SPEND') return;

    if (!out[memberId]) out[memberId] = [];
    if (out[memberId].indexOf(cours) === -1) out[memberId].push(cours);
  });

  return out;
}

// ─────────────────────────────────────────────
// GLOBAL UNIQUE ID GENERATOR (v4.6.1)
// ─────────────────────────────────────────────
function generateUniqueMemberId_() {
  var props = PropertiesService.getScriptProperties();
  var seqRaw = props.getProperty('memberIdSeq');
  var seq = parseInt(seqRaw, 10);

  // First-time bootstrap: look at existing IDs in MEMBERS to avoid collisions
  if (!seq || isNaN(seq)) {
    var s = sh_(CFG.SHEET.MEMBERS);
    var lastRow = s.getLastRow();
    var maxNum = 0;

    if (lastRow >= 2) {
      var ids = s.getRange(2, 1, lastRow - 1, 1).getValues(); // col A
      ids.forEach(function (row) {
        var raw = String(row[0] || '');
        var m = raw.match(/^DW-(\d+)$/);
        if (m) {
          var n = parseInt(m[1], 10);
          if (!isNaN(n) && n > maxNum) maxNum = n;
        }
      });
    }

    // Start at least at 2000 to stay well above your earlier manual DW-10xx, DW-20xx etc.
    seq = Math.max(maxNum, 2000);
  }

  // Next ID
  seq += 1;
  props.setProperty('memberIdSeq', String(seq));

  return 'DW-' + seq;
}


// Issue Card
function issueMemberCard_(p, by, submissionId, dateObj){
  if (!p.email || !/@/.test(String(p.email))) return { ok:false, err:'EMAIL_REQUIRED' };

  if (!p.renew){
    const act = findActiveCardByEmail_(p.email);
    if (act){ return { ok:false, err:'ACTIVE_CARD_EXISTS', active: act }; }
  }

  const s = sh_(CFG.SHEET.MEMBERS);
  const row = s.getLastRow() + 1;
  const id  = generateUniqueMemberId_();
  const tok = token_();
  const total  = (String(p.cardType) === '6') ? 6 : 12;
  const months = (total === 6) ? 4 : 6;
  const paid   = p.paidDate || new Date();
  
  // p.preferredCours is a comma-separated string of “preferred classes” (for check-in rapide)
  s.appendRow([
    id, p.nom||'', p.prenom||'', p.telephone||'', p.email||'',
    String(p.cardType)||'', paid, months, '', total, 0, '',
    tok, '', '', 'Carte '+total, p.modePaiement||'', '', p.coursPref||'', p.commentaire||'', '' // U=trim code blank
  ]);
  // French separators for formulas
  s.getRange(row,9).setFormula('=IF(G'+row+'="";;EDATE(G'+row+';H'+row+'))');       // I
  s.getRange(row,12).setFormula('=MAX(J'+row+'-K'+row+';0)');                       // L

  if (CFG.WEBAPP_BASE.startsWith('http')){
    const url = CFG.WEBAPP_BASE+'?id='+encodeURIComponent(id)+'&t='+encodeURIComponent(tok);
    s.getRange(row,14).setValue(url); // N
    s.getRange(row,15).setFormula('=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=240x240&data=" & ENCODEURL(N'+row+'))'); // O
  }

  // LOG (tagged with submission_id)
  log_({
    timestamp: dateObj || todayMid_(),   // ✅ backdated session date
    member_id:id,
    name:(p.prenom||'')+' '+(p.nom||''),
    action:'ISSUE',
    amount:total,
    classes_left_after:total,
    by,
    type_offre:'Carte',
    mode_paiement:p.modePaiement||'',
    email:p.email||'',
    submission_id: submissionId || ''
  });


  // Payment auto-log (use captured amountEncaisse) + SubmissionID
  ensurePaymentsSheet_();
  const paidAmt = (p.amountEncaisse === '' || p.amountEncaisse == null) ? 0 : +p.amountEncaisse;
  logPayment_({
    amount: paidAmt,
    mode: p.modePaiement||'',
    ref: p.refPaiement || '',
    payer:(p.prenom||'')+' '+(p.nom||''),
    linked:id,
    by,
    note:'Auto (carte)',
    submissionId: submissionId || ''
  });

  // Email with QR (add ID in subject + recap & reminder)
  if (CFG.EMAIL_ON_ISSUE && p.email && CFG.WEBAPP_BASE.startsWith('http')){
    const url           = s.getRange(row,14).getValue();
    const untilDate     = s.getRange(row, 9).getValue();
    const untilLabel    = untilDate ? fmtFR_(untilDate) : '';
    const totalCourses  = s.getRange(row,10).getValue();
    const qrUrl         = 'https://api.qrserver.com/v1/create-qr-code/?size=480x480&data=' + encodeURIComponent(url);

    const expected      = (String(p.cardType)==='12') ? CFG.PRICE_CARD_12 : CFG.PRICE_CARD_6;
    const owed          = expected - paidAmt;
    let   owedLineText  = '';
    if (owed > 0)      owedLineText = 'À payer : '+owed+'€';
    else if (owed < 0) owedLineText = 'Avoir : '+Math.abs(owed)+'€';
    else               owedLineText = 'Solde : 0€';

    const firstName = (p.prenom || '').trim();
    const lastRaw   = (p.nom || '').trim();
    const lastUp    = lastRaw ? lastRaw.toUpperCase() : '';
    const greetName = firstName || lastUp || 'Panthère';
    const fullName  = [firstName, lastUp].filter(Boolean).join(' ');

    const issueDateLabel = fmtFR_(paid);

    MailApp.sendEmail({
      to: p.email,
      subject: 'Ta carte danse (ID: '+id+') — ' + (CFG.FROM_NAME || 'La Panthère Rose'),
      htmlBody:
        'Salut ' + greetName + ' ! 😉<br><br>' +
        '🎫 BOOM ! Ta Flexi Carte Panthère est prête 🐾🔥<br><br>' +
        '<strong>Récapitulatif</strong> :<br>' +
        '<ul>' +
          '<li><strong>ID</strong> : ' + id + '</li>' +
          '<li><strong>Nom complet</strong> : ' + fullName + '</li>' +
          '<li><strong>Type de carte</strong> : Carte ' + totalCourses + ' cours</li>' +
          '<li><strong>Cours préférés</strong> : ' + (p.coursPref || '—') + '</li>' +
          '<li><strong>Date d’émission</strong> : ' + issueDateLabel + '</li>' +
          (untilLabel ? ('<li><strong>Valide jusqu’au</strong> : ' + untilLabel + '</li>') : '') +
          '<li><strong>Montant</strong> : ' + expected + '€</li>' +
          '<li><strong>Payé</strong> : ' + paidAmt + '€</li>' +
          '<li><strong>Solde</strong> : ' + owedLineText + '</li>' +
          '<li><strong>Mode de paiement</strong> : ' + (p.modePaiement || '') + '</li>' +
        '</ul>' +
        '<a href="' + url + '" style="display:inline-block;padding:10px 16px;background:#1a73e8;color:#fff;text-decoration:none;border-radius:8px" target="_blank">Voir ma carte / mon solde</a>' +
        '<br><small>(Lien direct vers ton espace élève.)</small><br><br>' +
        'Ou scanne ce QR :<br>' +
        '<img src="' + qrUrl + '" alt="QR carte" style="max-width:240px;width:240px;height:auto;border:1px solid #eee;border-radius:8px"><br>' +
        '<div style="margin-top:8px">Pense à ton <strong>ID</strong> ou à ce QR pour utiliser ta carte et valider ta présence à l’arrivée en cours.</div>' +
        '<br>See you on ze floor! ;)<br>' + (CFG.SIGN_NAME || 'SignName'),
      name: (CFG.FROM_NAME || 'La Panthère Rose'),
      cc: CFG.CC_EMAIL || ''
    });
  }
  return { ok:true, id };
}

// Issue Trimestre
function issueTrimestre_(p, by, submissionId, dateObj){
  if (!p.email || !/@/.test(String(p.email))) return { ok:false, err:'EMAIL_REQUIRED' };
  if (!p.trimCode) return { ok:false, err:'TRIM_CODE_REQUIRED' };

  // Block duplicate active trimestre (same email + same trimester code)
  const dup = findActiveTrimByEmailAndCode_(p.email, p.trimCode);
  if (dup) return { ok:false, err:'ACTIVE_TRIM_EXISTS', existing: dup };

  const s   = sh_(CFG.SHEET.MEMBERS);
  const row = s.getLastRow() + 1;
  const id  = generateUniqueMemberId_();
  const tok = token_();
  const paid = p.paidDate || new Date();

  // p.coursIncl is the comma-separated list of “included classes”
  s.appendRow([
    id, p.nom||'', p.prenom||'', p.telephone||'', p.email||'',
    '', paid, '', '', 0, 0, '', tok, '', '',
    'Trimestre', p.modePaiement||'', (p.forfait||''), (p.coursIncl||''), p.commentaire||'', (p.trimCode||'')
  ]);

  if (CFG.WEBAPP_BASE.startsWith('http')){
    const url = CFG.WEBAPP_BASE+'?id='+encodeURIComponent(id)+'&t='+encodeURIComponent(tok);
    s.getRange(row,14).setValue(url);
    s.getRange(row,15).setFormula('=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=240x240&data=" & ENCODEURL(N'+row+'))');
  }

  // LOG (tagged)
  log_({
    timestamp: dateObj || todayMid_(),   // ✅ backdated session date
    member_id:id,
    name:(p.prenom||'')+' '+(p.nom||''),
    action:'ISSUE',
    by,
    type_offre:'Trimestre',
    mode_paiement:p.modePaiement||'',
    email:p.email||'',
    note:(p.forfait||''),
    submission_id: submissionId || ''
  });

  // Payment auto-log (use amountEncaisse) + SubmissionID
  ensurePaymentsSheet_();
  const amount = (p.amountEncaisse === '' || p.amountEncaisse == null) ? 0 : +p.amountEncaisse;
  logPayment_({
    amount,
    mode: p.modePaiement||'',
    ref:p.refPaiement||'',
    payer:(p.prenom||'')+' '+(p.nom||''),
    linked:id,
    by,
    note:'Auto (trimestre)',
    submissionId: submissionId || ''
  });

  // Email (with ID) — no link / QR for trimestres
  if (CFG.EMAIL_ON_ISSUE && p.email && CFG.WEBAPP_BASE.startsWith('http')){
    const trims    = readSettings_().trims || [];
    let   period   = '';
    const expected = CFG.TRIM_PRICES[p.forfait] || 0;
    const paidAmt  = amount;
    const owed     = expected - paidAmt;
    const owedLine = (owed>0)
      ? ('<li><strong>Reste à payer</strong> : '+owed+'€</li>')
      : (owed<0 ? ('<li><strong>Avoir</strong> : '+Math.abs(owed)+'€</li>') : '');

    if (p.trimCode){
      const found = trims.find(x=>x.code===p.trimCode);
      if (found && found.start && found.end){
        period = fmtFR_(found.start)+' → '+fmtFR_(found.end);
      }
    }

    const firstName = (p.prenom || '').trim();
    const lastRaw   = (p.nom || '').trim();
    const lastUp    = lastRaw ? lastRaw.toUpperCase() : '';
    const greetName = firstName || lastUp || 'Panthère';
    const fullName  = [firstName, lastUp].filter(Boolean).join(' ');

    MailApp.sendEmail({
      to: p.email,
      subject: 'Ton inscription au trimestre (ID: '+id+') — ' + (CFG.FROM_NAME || 'La Panthère Rose'),
      htmlBody:
        'Salut ' + greetName + ' ! 😉<br><br>' +
        '📝 BOOM ! Ton inscription au trimestre est confirmée 🔥🐾<br><br>' +
        '<strong>Récapitulatif</strong> :<br>' +
        '<ul>' +
          '<li><strong>ID</strong> : ' + id + '</li>' +
          '<li><strong>Nom complet</strong> : ' + fullName + '</li>' +
          '<li><strong>Forfait</strong> : ' + (p.forfait||'') + '</li>' +
          (p.trimCode ? ('<li><strong>Trimestre</strong> : ' + p.trimCode + (period?(' — '+period):'') + '</li>') : '') +
          (p.coursIncl ? ('<li><strong>Cours inclus</strong> : ' + p.coursIncl + '</li>') : '') +
          '<li><strong>Montant</strong> : ' + expected + '€</li>' +
          '<li><strong>Payé</strong> : ' + paidAmt + '€</li>' +
          owedLine +
          '<li><strong>Mode de paiement</strong> : ' + (p.modePaiement || '') + '</li>' +
        '</ul>' +
        'See you on ze floor! ;)<br>' + (CFG.SIGN_NAME || 'SignName'),
      name: (CFG.FROM_NAME || 'La Panthère Rose'),
      cc: CFG.CC_EMAIL || ''
    });
  }

  return { ok:true, id };
}


// Update info (fill missing only)
function updateMemberInfoMissing_(id, fields){
  const {row, sheet}=findMemberById_(id); if(!row) return {ok:false,msg:'Membre introuvable'};
  let changed=0;
  const map = { nom:2, prenom:3, telephone:4, email:5 };
  Object.keys(fields).forEach(k=>{
    const col = map[k]; if(!col) return;
    const cur = sheet.getRange(row,col).getValue();
    if(!cur && fields[k]){ sheet.getRange(row,col).setValue(fields[k]); changed++; }
  });
  return { ok:true, changed };
}

function checkinMulti_(memberId, by, jour, coursList, dateObj, submissionId){
  if (!coursList || !coursList.length) return {ok:false,msg:'Aucun cours sélectionné'};
  const it = findMemberById_(memberId);
  if (!it.row) return {ok:false,msg:'Membre introuvable'};

  const data = it.data;
  const type = String(data[15]||'').toLowerCase();
  const isTrim = type.startsWith('trimestre');
  const isCard = type.startsWith('carte');
  if (!isTrim && !isCard) return {ok:false,msg:'Offre non compatible'};

  const dObj = dateObj || todayMid_();
  const msgs = [];
  const spentCourses = [];

  for (const c of coursList){
    if (isTrim){
      const r = attendById_(memberId, by, jour, c, submissionId, dObj);
      if (!r.ok) return r;
      msgs.push('ATTEND '+c);
    } else if (isCard){
      const r = spendOneById_(memberId, by, jour, c, submissionId, dObj);
      if (!r.ok) return r;
      msgs.push('SPEND '+c);
      spentCourses.push(c);
    }
  }

  // One aggregated email for card usage
  if (isCard && spentCourses.length){
    const refreshed = findMemberById_(memberId);
    const total = +refreshed.data[9] || 0;
    const used  = +refreshed.data[10] || 0;
    const leftAfter = computeLeft_(total, used);
    const email = String(refreshed.data[4]||'');
    const name  = getNameFromRow_(refreshed.data);
    const idRef = String(refreshed.data[0]||'');
    const urlRef= String(refreshed.data[13]||'');

    sendCardUsageEmail_(
      email,
      name,
      dObj,
      spentCourses,
      leftAfter,
      total,
      by,
      idRef,
      urlRef
    );
  }

  return {ok:true,msgs};
}

// Spend / Attend (with duplicate guard support)
function spendOneById_(memberId, by, jour, cours, submissionId, dateObj){
  const {row, data, sheet} = findMemberById_(memberId); if(!row) return {ok:false,msg:'Membre introuvable'};
  const typeOffre = String(data[15]||'');
  const total=+data[9]||0; let used=+data[10]||0;
  const vu=data[8]; const email=data[4]||''; const name=getNameFromRow_(data);
  if(!jour || !cours) return {ok:false,msg:'Choisir Jour et Cours'};
  if(!typeOffre.toLowerCase().startsWith('carte')) return {ok:false,msg:'Pas une carte'};
  if(isExpired_(vu)) return {ok:false,msg:'Carte expirée'};

  const dObj = dateObj || todayMid_();
  if (alreadyCheckedOnDate_(memberId, 'SPEND', jour, cours, dObj)) {
    return {ok:false,msg:'Déjà validé pour ce cours à cette date'};
  }

  const left=computeLeft_(total,used); if(left<1) return {ok:false,msg:'Solde insuffisant'};
  used+=1; sheet.getRange(row,11).setValue(used);
  const newLeft=computeLeft_(total,used);

  log_({
    timestamp: dObj,
    member_id:data[0],
    name,
    action:'SPEND',
    amount:1,
    classes_left_after:newLeft,
    by,
    type_offre:'Carte',
    jour,
    cours,
    email,
    submission_id: submissionId || ''
  });
  return { ok:true, name, left:newLeft, total, vu };
}

function attendById_(memberId, by, jour, cours, submissionId, dateObj){
  const {row, data} = findMemberById_(memberId); 
  if(!row) return {ok:false,msg:'Membre introuvable'};

  const typeOffre = String(data[15]||'');
  if(!jour || !cours) return {ok:false,msg:'Choisir Jour et Cours'};
  if(!typeOffre.toLowerCase().startsWith('trimestre')) return {ok:false,msg:'Pas un trimestre'};

  const dObj = dateObj || todayMid_();

  // Enforce trimester window if code exists
  const trimCode = String(data[20]||'').trim();
  if (trimCode){
    const trims = readSettings_().trims || [];
    const found = trims.find(x=>String(x.code||'').trim()===trimCode);
    if (found && found.start && found.end){
      const sdt = new Date(found.start); sdt.setHours(0,0,0,0);
      const edt = new Date(found.end);   edt.setHours(0,0,0,0);
      if (dObj < sdt || dObj > edt){
        return {ok:false,msg:'Trimestre hors période'};
      }
    }
  }

  // Included-classes guard
  const inclRaw = String(data[18] || ''); // S = cours inclus
  if (!inclRaw){
    return {ok:false,msg:'Aucun cours inclus enregistré pour ce trimestre'};
  }
  const inclList = inclRaw.split(',').map(x=>x.trim()).filter(Boolean);
  if (inclList.indexOf(cours) === -1){
    return {ok:false,msg:'Ce cours n’est pas inclus dans son abonnement'};
  }

  if (alreadyCheckedOnDate_(memberId, 'ATTEND', jour, cours, dObj)) {
    return {ok:false,msg:'Déjà validé pour ce cours à cette date'};
  }

  const email = data[4]||''; 
  const name  = getNameFromRow_(data);

  log_({
    timestamp: dObj,
    member_id:data[0],
    name,
    action:'ATTEND',
    by,
    type_offre:'Trimestre',
    jour,
    cours,
    email,
    submission_id: submissionId || ''
  });
  return { ok:true, name };
}

// Student URL rebuild
function rebuildStudentUrls_(){
  const s=sh_(CFG.SHEET.MEMBERS); const last=s.getLastRow();
  if(last<2) { SpreadsheetApp.getUi().alert('No members yet'); return; }
  if(!CFG.WEBAPP_BASE.startsWith('http')){ SpreadsheetApp.getUi().alert('Set CFG.WEBAPP_BASE first'); return; }
  let made=0;
  for(let row=2; row<=last; row++){
    let id=s.getRange(row,1).getValue();
    if(!id){
      id=CFG.ID_PREFIX+String(100000+row).slice(-6);
      s.getRange(row,1).setValue(id);
    }
    let tok=s.getRange(row,13).getValue();
    if(!tok){
      tok=token_();
      s.getRange(row,13).setValue(tok);
    }
    const url=CFG.WEBAPP_BASE+'?id='+encodeURIComponent(id)+'&t='+encodeURIComponent(tok);
    s.getRange(row,14).setValue(url);
    s.getRange(row,15).setFormula('=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=240x240&data=" & ENCODEURL(N'+row+'))');
    made++;
  }
  SpreadsheetApp.getUi().alert('Rebuilt URLs for '+made+' member(s).');
}

// Web App
function doGet(e) {
  try {
    const pSingle = (e && e.parameter) || {};
    const token = (pSingle.staff_token || '').trim();
    const pin   = (pSingle.pin || '').trim();
    const mode  = (pSingle.mode || '').trim();

    // Access to multi-valued fields if POSTed (e.parameters)
    const pMulti = (e && e.parameters) || {};

    if (mode === 'staff') {
      const submissionId = pSingle.submission_id || '';

      const staff = getTeamByToken_(token);
      const showRecap = norm_(staff.name).indexOf('adminlpr') !== -1;
      if (!staff) return HtmlService.createHtmlOutput('Accès refusé (token inconnu).');
      if (!staff.active) return HtmlService.createHtmlOutput('Accès refusé (profil inactif).');

      // No / wrong PIN → show PIN form
      if (!/^\d{4}$/.test(staff.pin) || pin !== staff.pin) {
        const html = ''
          + '<!doctype html><html><body style="font-family:sans-serif;padding:16px">'
          + '<h1>LPR Registration — Équipe</h1>'
          + '<h2>Salut, <strong>'+staff.name+'</strong> 👋🏾</h2>'
          + '<form method="POST" action="' + ScriptApp.getService().getUrl() + '" target="_top">'
          + '  <input type="hidden" name="mode" value="staff">'
          + '  <input type="hidden" name="staff_token" value="' + token + '">'
          + '  <label>PIN (4 chiffres): <input name="pin" inputmode="numeric" maxlength="4" required></label>'
          + '  <button type="submit">GO !</button>'
          + '</form>'
          + '</body></html>';
        return HtmlService.createHtmlOutput(html);
      }

      // Context is now DATE-only
      const ctxDateStr = pSingle.ctx_date || '';
      const ctxDateObj = ctxDateStr ? parseYMD_(ctxDateStr) : null;

      // If no ctx_date provided, ctxJour stays empty → front will open modal
      let ctxJour  = ctxDateObj ? jourFromDate_(ctxDateObj) : '';
      let ctxCours = ''; // no single focus class anymore

      // Internal safe fallback for actions/logs
      const effectiveCtxDateObj = ctxDateObj || todayMid_();

      // Reporting: “classe + date que je veux voir”
      const repJour  = pSingle.rep_jour || ctxJour || '';
      const repCours = pSingle.rep_cours || ctxCours || '';

      // Filters: default on first load (members+prereg), otherwise respect checkboxes
      const isSearch = (String(pSingle.do||'').toUpperCase()==='SEARCH');
      const filtMembers  = isSearch ? !!pSingle.f_members : true;
      const filtPreReg   = isSearch ? !!pSingle.f_prereg  : true;
      const filtContacts = isSearch ? !!pSingle.f_contacts: false;

      const limit = 5;
      const includeMore = (pSingle.more==='1');

      // Actions
      let flashHtml = '';
      const action = (pSingle.do || '').toUpperCase();

      if (action === 'ISSUE') {

        if (submissionId && hasSubmissionBeenHandled_(submissionId)) {
          flashHtml = flashBox_('Cette opération a déjà été enregistrée (rechargement ignoré).', 'warn');
        } else {
          // Collect preferred classes (can be multi-valued)
          let coursPrefArr = [];
          if (pMulti.coursPref){
            coursPrefArr = Array.isArray(pMulti.coursPref) ? pMulti.coursPref : [pMulti.coursPref];
          }
          const coursPref = coursPrefArr.join(', ');

          const renew = (pSingle.renew==='1');
          const r1 = issueMemberCard_({
            nom: pSingle.nom||'', prenom: pSingle.prenom||'', telephone: pSingle.telephone||'',
            email: pSingle.email||'', cardType: pSingle.cardType||'6',
            modePaiement: pSingle.modePaiement||'', commentaire: pSingle.commentaire||'',
            amountEncaisse: pSingle.amountEncaisse || '',
            refPaiement: pSingle.refPaiement || '',
            coursPref,
            renew,
            paidDate: ctxDateObj   // ✅ backdate payment + ISSUE log
          }, staff.name, submissionId, ctxDateObj);

          if (r1.ok) {
            const expected = (pSingle.cardType==='12') ? CFG.PRICE_CARD_12 : CFG.PRICE_CARD_6;
            const paidAmt  = (pSingle.amountEncaisse === '' || pSingle.amountEncaisse == null) ? 0 : +pSingle.amountEncaisse;
            const owed     = expected - paidAmt;
            const chip = (owed>0)
              ? '<span class="chip" style="border-color:#d33;background:#fdecea">À payer : '+owed+'€</span>'
              : (owed<0 ? '<span class="chip" style="border-color:#f5a623;background:#fff5e6">À rembourser : '+Math.abs(owed)+'€</span>' : '');

            // Use-now MULTI for cards (hidden field "useNow_multi_card" like "Cours A|Cours B")
            let nowClasses = [];
            if (pSingle.useNow_multi_card){
              nowClasses = String(pSingle.useNow_multi_card)
                .split('|').map(s=>s.trim()).filter(Boolean);
            }

            if (nowClasses.length && ctxJour){
              const rM = checkinMulti_(r1.id, staff.name, ctxJour, nowClasses, ctxDateObj, submissionId);
              if (!rM.ok){
                flashHtml = flashBox_(
                  'Carte créée ✅ (ID: '+r1.id+') mais check-in non effectué : '+rM.msg+
                  '<div class="small">'+(pSingle.cardType==='12'?'Carte 12':'Carte 6')+
                  ' — Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>',
                  'warn'
                );
              } else {
                flashHtml = flashBox_(
                  'Carte créée ✅ (ID: '+r1.id+') & check-in enregistré — '+
                  (pSingle.prenom||'')+' '+(pSingle.nom||'')+
                  ' — '+rM.msgs.join(' · ') +
                  '<div class="small">Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>',
                  'ok'
                );
              }
            } else {
              flashHtml = flashBox_(
                'Carte créée ✅ (ID: '+r1.id+') : ' + (pSingle.prenom||'')+' '+(pSingle.nom||'') +
                ' — ' + (pSingle.cardType==='12'?'Carte 12':'Carte 6') +
                ' — ' + (pSingle.modePaiement||'') +
                '<div class="small">Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>',
                'ok'
              );
            }
          } else {
            if (r1.err === 'ACTIVE_CARD_EXISTS' && r1.active) {
              const a = r1.active;
              flashHtml = flashBox_(
                'Impossible de créer : une carte active existe déjà pour cet email ' +
                '(ID: ' + a.id + ', ' + a.name + ', solde ' + a.left + '/' + a.total +
                (a.until ? (' — jusqu’au ' + fmtFR_(a.until)) : '') + '). ' +
                'Cochez “Renouveler carte” si elle est expirée ou le solde est 0.',
                'err'
              );
            } else if (r1.err === 'EMAIL_REQUIRED') {
              flashHtml = flashBox_('Email requis pour créer une carte.', 'err');
            } else {
              flashHtml = flashBox_('Création de carte impossible.', 'err');
            }
          }
        }
      }

      if (action === 'TRIM') {

        if (submissionId && hasSubmissionBeenHandled_(submissionId)) {
          flashHtml = flashBox_('Cette inscription a déjà été enregistrée (rechargement ignoré).', 'warn');
        } else {
          // coursIncl can be single string or array (via e.parameters)
          let coursInclArr = [];
          if (pMulti.coursIncl){
            coursInclArr = Array.isArray(pMulti.coursIncl) ? pMulti.coursIncl : [pMulti.coursIncl];
          }
          const coursIncl = coursInclArr.join(', ');

          const rT = issueTrimestre_({
            nom: pSingle.nom||'', prenom: pSingle.prenom||'', telephone: pSingle.telephone||'',
            email: pSingle.email||'', forfait: pSingle.forfait||'',
            coursIncl, trimCode: pSingle.trimCode || '',
            modePaiement: pSingle.modePaiement||'', commentaire: pSingle.commentaire||'',
            amountEncaisse: pSingle.amountEncaisse || '', refPaiement: pSingle.refPaiement || '',
            paidDate: ctxDateObj   // ✅ backdate payment + ISSUE log
          }, staff.name, submissionId, ctxDateObj);

          if (rT.ok) {
            const expected = CFG.TRIM_PRICES[pSingle.forfait] || 0;
            const paidAmt  = (pSingle.amountEncaisse === '' || pSingle.amountEncaisse == null) ? 0 : +pSingle.amountEncaisse;
            const owed     = expected - paidAmt;
            const chip = (owed>0)
              ? '<span class="chip" style="border-color:#d33;background:#fdecea">À payer : '+owed+'€</span>'
              : (owed<0 ? '<span class="chip" style="border-color:#f5a623;background:#fff5e6">À rembourser : '+Math.abs(owed)+'€</span>' : '');
            const incl = (coursIncl && coursIncl.trim())
              ? ('<div class="small">Cours inclus : ' + coursIncl + '</div>')
              : '';

// Use-now MULTI for trimestres (hidden field "useNow_multi")
let nowClassesT = [];
if (pSingle.useNow_multi){
  nowClassesT = String(pSingle.useNow_multi)
    .split('|').map(s=>s.trim()).filter(Boolean);
}

            if (nowClassesT.length && ctxJour){
              const rM = checkinMulti_(rT.id, staff.name, ctxJour, nowClassesT, ctxDateObj, submissionId);
              if (!rM.ok){
                flashHtml = flashBox_(
                  'Trimestre créé ✅ (ID: '+rT.id+') mais présence non marquée : '+rM.msg+
                  '<div class="small">Forfait : '+(pSingle.forfait||'')+
                  ' — Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>'+incl,
                  'warn'
                );
              } else {
                flashHtml = flashBox_(
                  'Trimestre créé ✅ (ID: '+rT.id+') & présence enregistrée — '+
                  (pSingle.prenom||'')+' '+(pSingle.nom||'')+
                  ' — '+rM.msgs.join(' · ') +
                  '<div class="small">Forfait : '+(pSingle.forfait||'')+
                  ' — Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>'+incl,
                  'ok'
                );
              }
            } else {
              flashHtml = flashBox_(
                'Trimestre créé ✅ (ID: '+rT.id+') : ' +
                (pSingle.prenom||'')+' '+(pSingle.nom||'') +
                ' — ' + (pSingle.forfait||'') +
                ' — ' + (pSingle.modePaiement||'') +
                '<div class="small">Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>'+incl,
                'ok'
              );
            }
          } else {
            if (rT.err === 'ACTIVE_TRIM_EXISTS' && rT.existing) {
              flashHtml = flashBox_(
                'Inscription refusée : une inscription est déjà enregistrée pour ce trimestre ' +
                '(ID: ' + rT.existing.id + ', ' + rT.existing.name + ').',
                'err'
              );
            } else if (rT.err === 'EMAIL_REQUIRED') {
              flashHtml = flashBox_('Email requis pour une inscription au trimestre.', 'err');
            } else if (rT.err === 'TRIM_CODE_REQUIRED') {
              flashHtml = flashBox_('Choisissez le trimestre (code requis).', 'err');
            } else {
              flashHtml = flashBox_('Impossible de créer le trimestre.', 'err');
            }
          }
        }
      }

      if (action === 'UPDATEINFO' && pSingle.member_id) {
        const rU = updateMemberInfoMissing_(pSingle.member_id, {
          nom:pSingle.nom||'',
          prenom:pSingle.prenom||'',
          telephone:pSingle.telephone||'',
          email:pSingle.email||''
        });
        flashHtml = flashBox_(
          rU.ok ? ('Coordonnées mises à jour ('+rU.changed+')') : 'Impossible de mettre à jour',
          rU.ok?'ok':'err'
        );
      }

      if (action === 'DROPIN') {
        // Option A coherence: drop-in follows ctx_date + its ctxJour
        const jour = ctxJour || '';  // ignore form day to avoid mismatches

        // Multi-classes sent in hidden field "dropin_multi" as "Cours A|Cours B|Cours C"
        let coursList = [];
        const rawMulti = pSingle.dropin_multi || '';
        if (rawMulti) {
          coursList = rawMulti
            .split('|')
            .map(function (s) { return s.trim(); })
            .filter(function (s) { return s; });
        }

        if (!jour || !coursList.length) {
          flashHtml = flashBox_('Jour et au moins 1 cours obligatoires pour le Drop-in', 'err');
        } else if (submissionId && hasSubmissionBeenHandled_(submissionId)) {
          flashHtml = flashBox_('Ce drop-in a déjà été enregistré (rechargement ignoré).', 'warn');
        } else {

          const r2 = addOrUpdateMemberForDropin_({
            nom: pSingle.nom||'',
            prenom: pSingle.prenom||'',
            telephone: pSingle.telephone||'',
            email: pSingle.email||'',
            modePaiement: pSingle.modePaiement||'',
            jour: jour,
            coursList: coursList,
            by: staff.name
          }, submissionId, ctxDateObj);

          ensurePaymentsSheet_();
          // Default = 12€ x nb de cours cochés (ou 12€ si jamais liste vide par sécurité)
          const defaultTotal = CFG.PRICE_DROPIN * (coursList.length || 1);
          const amount = amountOrFallback_(pSingle.amountEncaisse, defaultTotal);

          logPayment_({
            amount: amount,
            mode: pSingle.modePaiement||'',
            ref: pSingle.refPaiement||'',
            payer: (pSingle.prenom||'')+' '+(pSingle.nom||''),
            linked: r2.id,
            by: staff.name,
            note: 'Auto (drop-in)',
            submissionId: submissionId || ''
          });

          const expected = defaultTotal;
          const paidAmt  = amount;
          const owed     = expected - paidAmt;
          const chip = (owed>0)
            ? '<span class="chip" style="border-color:#d33;background:#fdecea">À payer : '+owed+'€</span>'
            : (owed<0
                ? '<span class="chip" style="border-color:#f5a623;background:#fff5e6">À rembourser : '+Math.abs(owed)+'€</span>'
                : '<span class="chip" style="border-color:#9fddb2;background:#e6f7ea">OK payé</span>');

          flashHtml = flashBox_(
            r2.ok
              ? ('Présence drop-in enregistrée (ID: '+r2.id+') — '+
                 (pSingle.prenom||'')+' '+(pSingle.nom||'')+
                 ' — '+coursList.length+' cours'+
                 '<div class="small">Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>')
              : 'Erreur drop-in',
            r2.ok ? 'ok' : 'err'
          );
        }
      }


      if (action === 'SPEND1' && pSingle.member_id) {
        const r3 = spendOneById_(pSingle.member_id, staff.name, ctxJour, ctxCours, submissionId, effectiveCtxDateObj);

        // ✅ Send usage email on successful single spend
        if (r3.ok){
          const it2 = findMemberById_(pSingle.member_id);
          const email2 = it2.row ? String(it2.data[4]||'') : '';
          const id2    = it2.row ? String(it2.data[0]||'') : '';
          const url2   = it2.row ? String(it2.data[13]||'') : '';

          sendCardUsageEmail_(
            email2,
            r3.name,
            effectiveCtxDateObj,
            [ctxCours],
            r3.left,
            r3.total,
            staff.name,
            id2,
            url2
          );
        }

        // warn if owes money
        let oweNote = '';
        if (r3.ok && pSingle.member_id){
          const it = findMemberById_(pSingle.member_id);
          const exp = it.row ? expectedForMemberRow_(it.data) : 0;
          const paid = sumPaidForId_(pSingle.member_id);
          const owed = exp > 0 ? (exp - paid) : 0;
          if (owed > 0) oweNote = ' <span class="chip" style="border-color:#f5a623;background:#fff5e6">Solde dû: '+owed+'€</span>';
        }

        flashHtml = flashBox_(
          (r3.ok
            ? ('1 cours débité — ' + r3.name + ' — solde ' + r3.left + '/' + r3.total +
               (r3.vu?(' — jusqu’au '+fmtFR_(r3.vu)):'') + ' — ' + (ctxJour||'') + ' – ' + (ctxCours||''))
            : ('Impossible : '+r3.msg)
          ) + oweNote,
          r3.ok?'ok':'err'
        );
      }

      if (action === 'ATTEND' && pSingle.member_id) {
        const r4 = attendById_(pSingle.member_id, staff.name, ctxJour, ctxCours, submissionId, effectiveCtxDateObj);


        let oweNote = '';
        if (r4.ok && pSingle.member_id){
          const it = findMemberById_(pSingle.member_id);
          const exp = it.row ? expectedForMemberRow_(it.data) : 0;
          const paid = sumPaidForId_(pSingle.member_id);
          const owed = exp > 0 ? (exp - paid) : 0;
          if (owed > 0) oweNote = ' <span class="chip" style="border-color:#f5a623;background:#fff5e6">Solde dû: '+owed+'€</span>';
        }

        flashHtml = flashBox_(
          (r4.ok
            ? ('Présence enregistrée — ' + r4.name + ' — ' + (ctxJour||'') + ' – ' + (ctxCours||''))
            : ('Impossible : ' + r4.msg)
          ) + oweNote,
          r4.ok ? 'ok' : 'err'
        );
      }

      // DOORROW: multi-classe depuis la liste rapide (un seul submit par personne)
      if (action === 'DOORROW' && pSingle.member_id) {
        const jour = ctxJour || '';
        if (!jour) {
          flashHtml = flashBox_('Contexte manquant : choisis le jour / cours en haut.', 'err');
        } else {
          // Récupère les cours cochés dans la ligne
          let doorCoursArr = [];
          if (pMulti.doorCours) {
            doorCoursArr = Array.isArray(pMulti.doorCours) ? pMulti.doorCours : [pMulti.doorCours];
          }
          doorCoursArr = doorCoursArr.map(c => String(c || '').trim()).filter(Boolean);

          if (!doorCoursArr.length) {
            flashHtml = flashBox_('Aucun cours sélectionné pour cette personne.', 'warn');
          } else {
            const memberId = pSingle.member_id;
            const it = findMemberById_(memberId);
            if (!it.row) {
              flashHtml = flashBox_('Membre introuvable.', 'err');
            } else {
              const typeOffreRaw = String(it.data[15] || '');
              const typeOffre    = typeOffreRaw.toLowerCase();
              const isTrim       = typeOffre.startsWith('trimestre');
              const isCard       = typeOffre.startsWith('carte');

              if (!isTrim && !isCard) {
                flashHtml = flashBox_('Offre non gérée dans la liste rapide (ni carte ni trimestre).', 'err');
              } else {
                let okCount = 0;
                const errMsgs = [];
                const spentCourses = [];

                doorCoursArr.forEach(function(cours){
                  if (!cours) return;
                  if (isTrim) {
                    const rA = attendById_(memberId, staff.name, jour, cours, submissionId, ctxDateObj);
                    if (rA.ok) {
                      okCount++;
                    } else {
                      errMsgs.push(cours + ' : ' + rA.msg);
                    }
                  } else if (isCard) {
                    const rS = spendOneById_(memberId, staff.name, jour, cours, submissionId, ctxDateObj);
                    if (rS.ok) {
                      okCount++;
                      spentCourses.push(cours);
                    } else {
                      errMsgs.push(cours + ' : ' + rS.msg);
                    }
                  }
                });

                // One aggregated email if card spends happened
                if (isCard && spentCourses.length){
                  const refreshed = findMemberById_(memberId);
                  const total = +refreshed.data[9] || 0;
                  const used  = +refreshed.data[10] || 0;
                  const leftAfter = computeLeft_(total, used);
                  const email = String(refreshed.data[4]||'');
                  const name  = getNameFromRow_(refreshed.data);
                  const idRef = String(refreshed.data[0]||'');
                  const urlRef= String(refreshed.data[13]||'');

                  sendCardUsageEmail_(
                    email,
                    name,
                    ctxDateObj,
                    spentCourses,
                    leftAfter,
                    total,
                    staff.name,
                    idRef,
                    urlRef
                  );
                }

                if (okCount > 0) {
                  const baseMsg = (isTrim ? 'Présence enregistrée' : 'Cours débités') +
                    ' (' + okCount + ' cours) — ' + getNameFromRow_(it.data) +
                    ' — ' + jour;
                  const errPart = errMsgs.length
                    ? '<br><span class="small">Erreurs : ' + errMsgs.join(' ; ') + '</span>'
                    : '';
                  flashHtml = flashBox_(baseMsg + errPart, errMsgs.length ? 'warn' : 'ok');
                } else {
                  flashHtml = flashBox_(
                    'Aucune action effectuée : ' + (errMsgs[0] || 'vérifie la sélection'),
                    'err'
                  );
                }
              }
            }
          }
        }
      }

      if (action === 'PAY') {
        try{
          if (submissionId && hasSubmissionBeenHandled_(submissionId)) {
            flashHtml = flashBox_('Ce paiement a déjà été enregistré (rechargement ignoré).', 'warn');
          } else {
            ensurePaymentsSheet_();
            const ids = (pSingle.linked_ids||'').split(',').map(s=>s.trim()).filter(Boolean).join(',');
            logPayment_({
              amount:pSingle.amount||0,
              mode:pSingle.modePaiement||'',
              ref:pSingle.ref||'',
              payer:pSingle.payer||'',
              linked:ids,
              by:staff.name,
              note:pSingle.note||'',
              submissionId: submissionId || ''
            });
            flashHtml = flashBox_('Paiement enregistré ('+(pSingle.modePaiement||'')+' — '+(pSingle.amount||'')+'€)', 'ok');
          }
        }catch(e){
          flashHtml = flashBox_('Erreur paiement','err');
        }
      }

      if (action === 'CHECKIN_MULTI' && pSingle.member_id){
        let coursSel = [];
        if (pMulti.cours_sel){
          coursSel = Array.isArray(pMulti.cours_sel) ? pMulti.cours_sel : [pMulti.cours_sel];
        }
        const rM = checkinMulti_(pSingle.member_id, staff.name, ctxJour, coursSel, effectiveCtxDateObj, submissionId);

        flashHtml = flashBox_(
          rM.ok ? ('Check-in enregistré : '+rM.msgs.join(' · ')) : ('Impossible : '+rM.msg),
          rM.ok?'ok':'err'
        );
      }

      // Build search results
      let results = [];
      let q = pSingle.q || '';
      const filters = { members:filtMembers, prereg:filtPreReg, contacts:filtContacts };
      let moreCount = 0, allCount = 0;
      if (isSearch && q) {
        const res = performSearch_(q, filters, limit, includeMore);
        results = res.shown;
        moreCount = res.more;
        allCount = res.allCount;

        // Attach already-validated classes for ctx_date (for pre-disable in search UI)
        const actionsMap = (ctxJour && ctxDateObj) ? getActionsByMemberOnDate_(ctxJour, ctxDateObj) : {};
        results.forEach(r=>{
          if (r.source === 'Members' && r.id){
            r.alreadyCourses = actionsMap[String(r.id)] || [];
          }
        });
      }

      // Reporting / class recap (separate from check-in context)
      let recap = null;
      const actUpper = (pSingle.do || '').toUpperCase();
      const sessionDate = pSingle.session_date || '';

      function todayYmd_() {
        const d = todayMid_();
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const da = String(d.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + da;
      }

      const effectiveDate = sessionDate || todayYmd_();

      if (actUpper === 'SESSION' && repJour && repCours) {
        recap = getClassSession_(effectiveDate, repJour, repCours);
      }

      // Prefill
      const pre = {
        nom: pSingle.prefill_nom||'',
        prenom: pSingle.prefill_prenom||'',
        email: pSingle.prefill_email||'',
        telephone: pSingle.prefill_tel||'',
        forfait: pSingle.prefill_forfait||'',
        pf: pSingle.pf || ''
      };

      // Settings payloads
      const st = readSettings_();

      // Door checklist (for today), based on ctxJour only
      const door = (ctxJour) ? buildDoorListOnDate_(ctxJour, effectiveCtxDateObj) : { classes:[], rows:[] };


      const t = HtmlService.createTemplateFromFile('page');
      t.view = {
        kind:'staff',
        showRecap: showRecap,
        staffName: staff.name,
        staffToken: staff.token,
        pin: pin,
        flashHtml,
        ctxJour, ctxCours,

        ctxDate: ctxDateStr || '',



        repJour,
        repCours,

        fMembers: filtMembers?'1':'0',
        fPreReg:  filtPreReg ?'1':'0',
        fContacts:filtContacts?'1':'0',
        searchQ: q,
        resultsJson: JSON.stringify(results),
        moreCount: moreCount||0,
        allCount: allCount||0,
        prefill: pre,
        recapJson: JSON.stringify(recap||{present:[], totals:{}, expected:[], absent:[]}),
        trimsJson: JSON.stringify(st.trims||[]),
        mappingJson: JSON.stringify({ jours: st.jours, mapping: st.mapping }),
        sessionDate: sessionDate || '',
        doorJson: JSON.stringify(door || { classes:[], rows:[] })
      };

      return t.evaluate()
              .setTitle('LPR Registration — Équipe')
              .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // STUDENT VIEW
    if (pSingle.id && pSingle.t) {
      const v = buildStudentView_(pSingle.id, pSingle.t);
      const t2 = HtmlService.createTemplateFromFile('page');
      t2.view = v;
      return t2.evaluate()
               .setTitle('La Panthère Rose — Espace élève')
               .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    return HtmlService.createHtmlOutput('Unknown mode.');

  } catch (err) {
    return HtmlService.createHtmlOutput('ERROR: ' + (err && err.message ? err.message : String(err)));
  }
}
function doPost(e){ return doGet(e); }


// Student View builder
function buildStudentView_(id,t){
  if(!id||!t) return { kind:'student', ok:false, msg:'Lien invalide' };
  const it = findMemberById_(id);
  if(!it.row || !it.data || t!==it.data[12]) return { kind:'student', ok:false, msg:'Lien invalide' };

  const row = it.data;
  const typeOffre = String(row[15]||'');
  if (typeOffre.toLowerCase().startsWith('trimestre')){
    const paid = row[6] || '';
    const forfait = row[17] || '';   // R
    const coursIncl = row[18] || ''; // S
    const trimCode = row[20] || '';  // U
    let trimWindow = '';
    const trims = readSettings_().trims || [];
    if (trimCode){
      const found = trims.find(x=> (x.code||'')===trimCode);
      if (found && found.start && found.end){
        trimWindow = fmtFR_(found.start)+' → '+fmtFR_(found.end);
      }
    }
    return {
      kind:'student', ok:true, isTrim:true,
      id, t,
      nom:row[1]||'', prenom:row[2]||'',
      trim: { forfait, paid: (paid? fmtFR_(paid):''), coursIncl, code:trimCode, window:trimWindow }
    };
  }

  const total=+row[9]||0, used=+row[10]||0, left=computeLeft_(total,used), vu=row[8]||'';
  return {
    kind:'student', ok:true, id, t, nom:row[1]||'', prenom:row[2]||'',
    total, used, left, until: (vu? fmtFR_(vu) : ''), expired: isExpired_(vu)
  };
}

// Drop-in (backdated to ctx_date, supports multiple classes)
function addOrUpdateMemberForDropin_(p, submissionId, dateObj){
  const s   = sh_(CFG.SHEET.MEMBERS);
  const row = s.getLastRow() + 1;
  const id  = generateUniqueMemberId_();
  
  // ✅ session date = ctx_date (backdating)
  const dObj = dateObj || todayMid_();
  // ✅ weekday derived from ctx_date
  const jourFromCtx = labelJourFR_(dObj);

  s.appendRow([
    id,
    p.nom||'',
    p.prenom||'',
    p.telephone||'',
    p.email||'',
    '',        // F: Type carte (not used)
    '',        // G: Date paiement (not used for drop-in)
    '',        // H: Mois (not used)
    '',        // I: Valide jusqu’au
    0,         // J: Total cours
    0,         // K: Utilisés
    '',        // L: Restants
    '',        // M: token
    '',        // N: URL
    '',        // O: QR
    'Drop-in', // P: Type d’offre
    p.modePaiement||'',
    '',        // R: Forfait
    '',        // S: Cours inclus / préférés (not used)
    '',        // T: Commentaire
    ''         // U: Code Trimestre
  ]);

  const coursList = Array.isArray(p.coursList)
      ? p.coursList
      : (p.coursList ? [p.coursList] : []);

  if (!coursList.length){
    log_({
      timestamp: dObj,                 // ✅ backdated ctx_date
      member_id: id,
      name: (p.prenom||'')+' '+(p.nom||''),
      action: 'DROPIN',
      amount: CFG.PRICE_DROPIN,
      by: p.by||'',
      type_offre: 'Drop-in',
      mode_paiement: p.modePaiement||'',
      jour: jourFromCtx,              // ✅ jour from ctx_date
      cours: '',
      email: p.email||'',
      submission_id: submissionId || ''
    });
  } else {
    coursList.forEach(c=>{
      log_({
        timestamp: dObj,               // ✅ backdated ctx_date
        member_id: id,
        name: (p.prenom||'')+' '+(p.nom||''),
        action: 'DROPIN',
        amount: CFG.PRICE_DROPIN,
        by: p.by||'',
        type_offre: 'Drop-in',
        mode_paiement: p.modePaiement||'',
        jour: jourFromCtx,            // ✅ jour from ctx_date
        cours: c||'',
        email: p.email||'',
        submission_id: submissionId || ''
      });
    });
  }

  return { ok:true, id };
}

// Staff lookup
function getTeamByToken_(staff_token){
  const { team } = readSettings_();
  for (const t of team){ if(t.token && t.token===staff_token) return t; }
  return null;
}

function sendCardUsageEmail_(email, name, dateObj, classList, leftAfter, total, by, idRef, urlRef){
  if (!email || !/@/.test(String(email))) return;

  // Name formatting
  const parts = String(name || '').trim().split(/\s+/);
  const firstName = parts[0] || '';
  const lastName  = (parts.slice(1).join(' ') || '').toUpperCase();
  const greetName = firstName || lastName || 'Panthère';
  const fullName  = [firstName, lastName].filter(Boolean).join(' ');

  const dayLabel = fmtFR_(dateObj);
  const classes  = (classList || []).map(c => '• '+c).join('<br>');
  const validUntil = getValidUntilForId_(idRef);


  const qrUrl = urlRef
    ? ('https://api.qrserver.com/v1/create-qr-code/?size=480x480&data=' + encodeURIComponent(urlRef))
    : '';

  MailApp.sendEmail({
    to: email,
    subject: 'Utilisation de ta carte danse (ID: ' + idRef + ') — ' + (CFG.FROM_NAME || 'La Panthère Rose'),
    htmlBody:
      'Salut ' + greetName + ' ! 😉<br><br>' +
      '✅ On a bien enregistré l’utilisation de ta FlexiCarte Panthère pour la séance du <strong>' + dayLabel + '</strong> :<br>' +
      classes + '<br><br>' +

      '<strong>Récap de ta carte :</strong><br>' +
      '• <strong>ID</strong> : ' + idRef + '<br>' +
      '• <strong>Type</strong> : Carte ' + total + ' cours<br>' +
      (validUntil ? ('• <strong>Valide jusqu’au</strong> : ' + validUntil + '<br>') : '') +
      '• <strong>Solde</strong> : ' + leftAfter + ' / ' + total + '<br><br>' +

      '<a href="' + urlRef + '" style="display:inline-block;padding:10px 16px;background:#1a73e8;color:#fff;text-decoration:none;border-radius:8px" target="_blank">Voir ma carte / mon solde</a>' +
      '<br><small>(Lien direct vers ton espace élève.)</small><br><br>' +

      (qrUrl ? ('Scanne ce QR si tu préfères :<br><img src="' + qrUrl + '" alt="QR" style="max-width:240px;width:240px;height:auto;border:1px solid #eee;border-radius:8px"><br><br>') : '') +

      'Souviens-toi de ton <strong>ID</strong> ou montre ton <strong>QR</strong> à l’accueil pour utiliser ta carte et valider ta présence à chaque cours.<br><br>' +
      'See you on ze floor! 😉<br>' +
      (CFG.SIGN_NAME || 'La Panthère Rose'),
    name: (CFG.FROM_NAME || 'La Panthère Rose'),
    cc: CFG.CC_EMAIL || ''
  });
}


// Flash HTML helper
function flashBox_(msg, kind){
  const color = kind==='ok' ? '#e6f7ea' : (kind==='warn' ? '#fff7e6' : '#fdecea');
  const border = kind==='ok' ? '#9fddb2' : (kind==='warn' ? '#ffd591' : '#f5c2c7');
  return '<div style="border:1px solid '+border+';background:'+color+';padding:10px;border-radius:10px;margin:10px 0">'+ msg +'</div>';
}

// ─────────────────────────────────────────────────────────────
// AJAX STAFF API — Chunk A
// Search + Issue Card + Issue Trim + Drop-in without full reload
// Called from page.html via google.script.run
// ─────────────────────────────────────────────────────────────

function staffAuthFromAjax_(p){
  const token = String(p.staff_token || '').trim();
  const pin   = String(p.pin || '').trim();
  const staff = getTeamByToken_(token);
  if (!staff) return { ok:false, err:'TOKEN_INVALID' };
  if (!staff.active) return { ok:false, err:'TOKEN_INACTIVE' };
  if (!/^\d{4}$/.test(staff.pin) || pin !== staff.pin) return { ok:false, err:'PIN_INVALID' };
  return { ok:true, staff };
}

function staffSearchAjax_(p){
  try{
    const auth = staffAuthFromAjax_(p);
    if (!auth.ok) return { ok:false, err:auth.err };

    const staff = auth.staff;

    const ctxDateStr = String(p.ctx_date || '').trim();
    const ctxDateObj = ctxDateStr ? parseYMD_(ctxDateStr) : null;
    const ctxJour    = ctxDateObj ? jourFromDate_(ctxDateObj) : '';
    const effectiveCtxDateObj = ctxDateObj || todayMid_();

    const q = String(p.q || '');

    // ─────────────────────────────────────
    // Filters with sensible defaults
    // Default = Members + PreReg ON, Contacts OFF
    // ─────────────────────────────────────
    let filtMembers  = (p.f_members  == null) ? true  : !!p.f_members;
    let filtPreReg   = (p.f_prereg   == null) ? true  : !!p.f_prereg;
    let filtContacts = (p.f_contacts == null) ? false : !!p.f_contacts;

    // Safety net: if all three are false (eg. weird frontend state),
    // fall back to the default again.
    if (!filtMembers && !filtPreReg && !filtContacts){
      filtMembers  = true;
      filtPreReg   = true;
      filtContacts = false;
    }

    const filters = {
      members:  filtMembers,
      prereg:   filtPreReg,
      contacts: filtContacts
    };

    const limit = 5;
    const includeMore = (String(p.more||'') === '1');

    let results = [];
    let moreCount = 0, allCount = 0;

    if (q){
      const res = performSearch_(q, filters, limit, includeMore);
      results   = res.shown;
      moreCount = res.more;
      allCount  = res.allCount;

      // Attach already-validated classes for ctx_date
      const actionsMap = (ctxJour && ctxDateObj)
        ? getActionsByMemberOnDate_(ctxJour, ctxDateObj)
        : {};
      results.forEach(r=>{
        if (r.source === 'Members' && r.id){
          r.alreadyCourses = actionsMap[String(r.id)] || [];
        }
      });
    }

    return {
      ok: true,
      ctxDate: ctxDateStr || '',
      ctxJour,
      results,
      moreCount,
      allCount,
      staffName: staff.name
    };
  } catch(e){
    return {
      ok:false,
      err:'SEARCH_ERROR',
      msg: e && e.message ? e.message : String(e)
    };
  }
}

function staffIssueCardAjax_(p){
  try{
    const auth = staffAuthFromAjax_(p);
    if (!auth.ok) return { ok:false, err:auth.err };
    const staff = auth.staff;

    const submissionId = String(p.submission_id || '');

    if (submissionId && (isDuplicateSubmission_(submissionId) || hasSubmissionBeenHandled_(submissionId))) {
      return { ok:true, kind:'warn',
        flashHtml: flashBox_('Cette opération a déjà été enregistrée (double-clic ignoré).', 'warn')
      };
    }

    const ctxDateStr = String(p.ctx_date || '').trim();
    const ctxDateObj = ctxDateStr ? parseYMD_(ctxDateStr) : null;
    const ctxJour  = ctxDateObj ? jourFromDate_(ctxDateObj) : '';
    const effectiveCtxDateObj = ctxDateObj || todayMid_();

    // Multi-valued coursPref
    let coursPrefArr = [];
    if (p.coursPref){
      coursPrefArr = Array.isArray(p.coursPref) ? p.coursPref : [p.coursPref];
    }
    const coursPref = coursPrefArr.join(', ');

    const renew = (String(p.renew||'')==='1');

    const r1 = issueMemberCard_({
      nom: p.nom||'', prenom: p.prenom||'', telephone: p.telephone||'',
      email: p.email||'', cardType: p.cardType||'6',
      modePaiement: p.modePaiement||'', commentaire: p.commentaire||'',
      amountEncaisse: p.amountEncaisse || '',
      refPaiement: p.refPaiement || '',
      coursPref,
      renew,
      paidDate: ctxDateObj
    }, staff.name, submissionId, ctxDateObj);

    if (!r1.ok){
      if (r1.err === 'ACTIVE_CARD_EXISTS' && r1.active){
        const a=r1.active;
        return {
          ok:false, kind:'err',
          flashHtml: flashBox_(
            'Impossible de créer : une carte active existe déjà pour cet email '+
            '(ID: '+a.id+', '+a.name+', solde '+a.left+'/'+a.total+
            (a.until?(' — jusqu’au '+fmtFR_(a.until)):'')+'). '+
            'Cochez “Renouveler carte” si elle est expirée ou le solde est 0.',
            'err'
          )
        };
      }
      if (r1.err === 'EMAIL_REQUIRED'){
        return { ok:false, kind:'err', flashHtml: flashBox_('Email requis pour créer une carte.', 'err') };
      }
      return { ok:false, kind:'err', flashHtml: flashBox_('Création de carte impossible.', 'err') };
    }

    const expected = (String(p.cardType)==='12') ? CFG.PRICE_CARD_12 : CFG.PRICE_CARD_6;
    const paidAmt  = (p.amountEncaisse === '' || p.amountEncaisse == null) ? 0 : +p.amountEncaisse;
    const owed     = expected - paidAmt;
    const chip = (owed>0)
      ? '<span class="chip" style="border-color:#d33;background:#fdecea">À payer : '+owed+'€</span>'
      : (owed<0 ? '<span class="chip" style="border-color:#f5a623;background:#fff5e6">À rembourser : '+Math.abs(owed)+'€</span>' : '');

    // Use-now multi for cards
    let nowClasses = [];
    if (p.useNow_multi_card){
      nowClasses = String(p.useNow_multi_card).split('|').map(s=>s.trim()).filter(Boolean);
    }

    let flashHtml = '';
    if (nowClasses.length && ctxJour){
      const rM = checkinMulti_(r1.id, staff.name, ctxJour, nowClasses, ctxDateObj, submissionId);
      if (!rM.ok){
        flashHtml = flashBox_(
          'Carte créée ✅ (ID: '+r1.id+') mais check-in non effectué : '+rM.msg+
          '<div class="small">'+(String(p.cardType)==='12'?'Carte 12':'Carte 6')+
          ' — Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>',
          'warn'
        );
        return { ok:true, kind:'warn', flashHtml, id:r1.id };
      }
      flashHtml = flashBox_(
        'Carte créée ✅ (ID: '+r1.id+') & check-in enregistré — '+
        (p.prenom||'')+' '+(p.nom||'')+
        ' — '+rM.msgs.join(' · ') +
        '<div class="small">Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>',
        'ok'
      );
      return { ok:true, kind:'ok', flashHtml, id:r1.id };
    }

    flashHtml = flashBox_(
      'Carte créée ✅ (ID: '+r1.id+') : '+(p.prenom||'')+' '+(p.nom||'')+
      ' — '+(String(p.cardType)==='12'?'Carte 12':'Carte 6')+
      ' — '+(p.modePaiement||'')+
      '<div class="small">Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>',
      'ok'
    );

    return { ok:true, kind:'ok', flashHtml, id:r1.id };

  }catch(e){
    return { ok:false, kind:'err', flashHtml: flashBox_('Création de carte impossible.', 'err'), msg:String(e) };
  }
}

function staffIssueTrimAjax_(p){
  try{
    const auth = staffAuthFromAjax_(p);
    if (!auth.ok) return { ok:false, err:auth.err };
    const staff = auth.staff;

    const submissionId = String(p.submission_id || '');

    if (submissionId && (isDuplicateSubmission_(submissionId) || hasSubmissionBeenHandled_(submissionId))) {
      return { ok:true, kind:'warn',
        flashHtml: flashBox_('Cette opération a déjà été enregistrée (double-clic ignoré).', 'warn')
      };
    }

    const ctxDateStr = String(p.ctx_date || '').trim();
    const ctxDateObj = ctxDateStr ? parseYMD_(ctxDateStr) : null;
    const ctxJour  = ctxDateObj ? jourFromDate_(ctxDateObj) : '';
    const effectiveCtxDateObj = ctxDateObj || todayMid_();

    let coursInclArr = [];
    if (p.coursIncl){
      coursInclArr = Array.isArray(p.coursIncl) ? p.coursIncl : [p.coursIncl];
    }
    const coursIncl = coursInclArr.join(', ');

    const rT = issueTrimestre_({
      nom: p.nom||'', prenom: p.prenom||'', telephone: p.telephone||'',
      email: p.email||'', forfait: p.forfait||'',
      coursIncl, trimCode: p.trimCode || '',
      modePaiement: p.modePaiement||'', commentaire: p.commentaire||'',
      amountEncaisse: p.amountEncaisse || '', refPaiement: p.refPaiement || '',
      paidDate: ctxDateObj
    }, staff.name, submissionId, ctxDateObj);

    if (!rT.ok){
      if (rT.err === 'ACTIVE_TRIM_EXISTS' && rT.existing){
        return {
          ok:false, kind:'err',
          flashHtml: flashBox_(
            'Inscription refusée : une inscription est déjà enregistrée pour ce trimestre '+
            '(ID: '+rT.existing.id+', '+rT.existing.name+').',
            'err'
          )
        };
      }
      if (rT.err === 'EMAIL_REQUIRED'){
        return { ok:false, kind:'err', flashHtml: flashBox_('Email requis pour une inscription au trimestre.', 'err') };
      }
      if (rT.err === 'TRIM_CODE_REQUIRED'){
        return { ok:false, kind:'err', flashHtml: flashBox_('Choisissez le trimestre (code requis).', 'err') };
      }
      return { ok:false, kind:'err', flashHtml: flashBox_('Impossible de créer le trimestre.', 'err') };
    }

    const expected = CFG.TRIM_PRICES[p.forfait] || 0;
    const paidAmt  = (p.amountEncaisse === '' || p.amountEncaisse == null) ? 0 : +p.amountEncaisse;
    const owed     = expected - paidAmt;
    const chip = (owed>0)
      ? '<span class="chip" style="border-color:#d33;background:#fdecea">À payer : '+owed+'€</span>'
      : (owed<0 ? '<span class="chip" style="border-color:#f5a623;background:#fff5e6">À rembourser : '+Math.abs(owed)+'€</span>' : '');
    const incl = (coursIncl && coursIncl.trim())
      ? ('<div class="small">Cours inclus : ' + coursIncl + '</div>')
      : '';

    // Use-now multi for trims
    let nowClassesT = [];
    if (p.useNow_multi){
      nowClassesT = String(p.useNow_multi).split('|').map(s=>s.trim()).filter(Boolean);
    }

    let flashHtml = '';
    if (nowClassesT.length && ctxJour){
      const rM = checkinMulti_(rT.id, staff.name, ctxJour, nowClassesT, ctxDateObj, submissionId);
      if (!rM.ok){
        flashHtml = flashBox_(
          'Trimestre créé ✅ (ID: '+rT.id+') mais présence non marquée : '+rM.msg+
          '<div class="small">Forfait : '+(p.forfait||'')+
          ' — Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>'+incl,
          'warn'
        );
        return { ok:true, kind:'warn', flashHtml, id:rT.id };
      }
      flashHtml = flashBox_(
        'Trimestre créé ✅ (ID: '+rT.id+') & présence enregistrée — '+
        (p.prenom||'')+' '+(p.nom||'')+
        ' — '+rM.msgs.join(' · ') +
        '<div class="small">Forfait : '+(p.forfait||'')+
        ' — Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>'+incl,
        'ok'
      );
      return { ok:true, kind:'ok', flashHtml, id:rT.id };
    }

    flashHtml = flashBox_(
      'Trimestre créé ✅ (ID: '+rT.id+') : '+(p.prenom||'')+' '+(p.nom||'')+
      ' — '+(p.forfait||'')+
      ' — '+(p.modePaiement||'')+
      '<div class="small">Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>'+incl,
      'ok'
    );

    return { ok:true, kind:'ok', flashHtml, id:rT.id };

  }catch(e){
    return { ok:false, kind:'err', flashHtml: flashBox_('Impossible de créer le trimestre.', 'err'), msg:String(e) };
  }
}

function staffDropinAjax_(p){
  try{
    const auth = staffAuthFromAjax_(p);
    if (!auth.ok) return { ok:false, err:auth.err };
    const staff = auth.staff;

    const submissionId = String(p.submission_id || '');

    if (submissionId && (isDuplicateSubmission_(submissionId) || hasSubmissionBeenHandled_(submissionId))) {
      return { ok:true, kind:'warn',
        flashHtml: flashBox_('Cette opération a déjà été enregistrée (double-clic ignoré).', 'warn')
      };
    }

    const ctxDateStr = String(p.ctx_date || '').trim();
    const ctxDateObj = ctxDateStr ? parseYMD_(ctxDateStr) : null;
    const ctxJour  = ctxDateObj ? jourFromDate_(ctxDateObj) : '';
    const effectiveCtxDateObj = ctxDateObj || todayMid_();

    // Multi-classes sent in hidden field "dropin_multi"
    let coursList = [];
    const rawMulti = p.dropin_multi || '';
    if (rawMulti) {
      coursList = String(rawMulti).split('|').map(s=>s.trim()).filter(Boolean);
    }

    if (!ctxJour || !coursList.length){
      return { ok:false, kind:'err', flashHtml: flashBox_('Jour et au moins 1 cours obligatoires pour le Drop-in', 'err') };
    }

    const r2 = addOrUpdateMemberForDropin_({
      nom: p.nom||'',
      prenom: p.prenom||'',
      telephone: p.telephone||'',
      email: p.email||'',
      modePaiement: p.modePaiement||'',
      jour: ctxJour,
      coursList: coursList,
      by: staff.name
    }, submissionId, ctxDateObj);

    ensurePaymentsSheet_();
    const defaultTotal = CFG.PRICE_DROPIN * (coursList.length || 1);
    const amount = amountOrFallback_(p.amountEncaisse, defaultTotal);

    logPayment_({
      amount: amount,
      mode: p.modePaiement||'',
      ref: p.refPaiement||'',
      payer: (p.prenom||'')+' '+(p.nom||''),
      linked: r2.id,
      by: staff.name,
      note: 'Auto (drop-in)',
      submissionId: submissionId || ''
    });

    const expected = defaultTotal;
    const paidAmt  = amount;
    const owed     = expected - paidAmt;
    const chip = (owed>0)
      ? '<span class="chip" style="border-color:#d33;background:#fdecea">À payer : '+owed+'€</span>'
      : (owed<0
          ? '<span class="chip" style="border-color:#f5a623;background:#fff5e6">À rembourser : '+Math.abs(owed)+'€</span>'
          : '<span class="chip" style="border-color:#9fddb2;background:#e6f7ea">OK payé</span>');

    const flashHtml = flashBox_(
      'Présence drop-in enregistrée (ID: '+r2.id+') — '+
      (p.prenom||'')+' '+(p.nom||'')+' — '+coursList.length+' cours'+
      '<div class="small">Attendu : '+expected+'€ — Reçu : '+paidAmt+'€ '+chip+'</div>',
      'ok'
    );

    return { ok:true, kind:'ok', flashHtml, id:r2.id };

  }catch(e){
    const msg = String(e);
    return {
      ok:false,
      kind:'err',
      flashHtml: flashBox_('Erreur drop-in : ' + msg, 'err'),
      msg: msg
    };
  }
}

function staffDoorrowAjax_(p){
  try{
    const auth = staffAuthFromAjax_(p);
    if (!auth.ok) {
      return { ok:false, err:auth.err, flashHtml: flashBox_('Accès refusé (token/PIN).', 'err') };
    }
    const staff = auth.staff;

    const submissionId = String(p.submission_id || '');
    // Optional extra idempotency – semantic guards in SPEND/ATTEND already protect us,
    // but this is cheap insurance against accidental resubmits.
    if (submissionId && (isDuplicateSubmission_(submissionId) || hasSubmissionBeenHandled_(submissionId))) {
      return {
        ok:true,
        kind:'warn',
        flashHtml: flashBox_('Cette opération a déjà été enregistrée (double-clic ignoré).', 'warn')
      };
    }

    const ctxDateStr = String(p.ctx_date || '').trim();
    const ctxDateObj = ctxDateStr ? parseYMD_(ctxDateStr) : null;
    const jour       = ctxDateObj ? jourFromDate_(ctxDateObj) : '';

    if (!jour) {
      return {
        ok:false,
        kind:'err',
        flashHtml: flashBox_('Contexte manquant : choisis le jour / la date en haut.', 'err')
      };
    }

    const memberId = String(p.member_id || '').trim();
    const cours    = String(p.cours || '').trim();

    if (!memberId || !cours) {
      return {
        ok:false,
        kind:'err',
        flashHtml: flashBox_('Paramètres manquants pour la liste rapide (membre ou cours).', 'err')
      };
    }

    const it = findMemberById_(memberId);
    if (!it.row) {
      return {
        ok:false,
        kind:'err',
        flashHtml: flashBox_('Membre introuvable pour la liste rapide.', 'err')
      };
    }

    const typeOffreRaw = String(it.data[15] || '');
    const typeOffre    = typeOffreRaw.toLowerCase();
    const isTrim       = typeOffre.startsWith('trimestre');
    const isCard       = typeOffre.startsWith('carte');

    if (!isTrim && !isCard) {
      return {
        ok:false,
        kind:'err',
        flashHtml: flashBox_('Offre non gérée dans la liste rapide (ni carte ni trimestre).', 'err')
      };
    }

    let res;
    if (isTrim) {
      // Trimestre → ATTEND (respecte fenêtre + cours inclus)
      res = attendById_(memberId, staff.name, jour, cours, submissionId, ctxDateObj);
    } else {
      // Carte → SPEND (respecte dates + solde + doublons)
      res = spendOneById_(memberId, staff.name, jour, cours, submissionId, ctxDateObj);
    }

    if (!res || !res.ok) {
      return {
        ok:false,
        kind:'err',
        flashHtml: flashBox_('Impossible : ' + (res && res.msg ? res.msg : 'action refusée.'), 'err')
      };
    }

    // Carte : email d’utilisation, comme pour SPEND1 / DOORROW existants
    if (isCard) {
      const refreshed = findMemberById_(memberId);
      const total = +refreshed.data[9] || 0;
      const used  = +refreshed.data[10] || 0;
      const leftAfter = computeLeft_(total, used);
      const email = String(refreshed.data[4]||'');
      const name  = getNameFromRow_(refreshed.data);
      const idRef = String(refreshed.data[0]||'');
      const urlRef= String(refreshed.data[13]||'');

      sendCardUsageEmail_(
        email,
        name,
        ctxDateObj || todayMid_(),
        [cours],
        leftAfter,
        total,
        staff.name,
        idRef,
        urlRef
      );
    }

    // Solde dû (copié de SPEND1 / ATTEND pour cohérence des chips)
    let oweNote = '';
    if (memberId){
      const it2  = findMemberById_(memberId);
      const exp  = it2.row ? expectedForMemberRow_(it2.data) : 0;
      const paid = sumPaidForId_(memberId);
      const owed = exp > 0 ? (exp - paid) : 0;
      if (owed > 0) {
        oweNote = ' <span class="chip" style="border-color:#f5a623;background:#fff5e6">Solde dû: '+owed+'€</span>';
      }
    }

    const baseMsg = isTrim
      ? ('Présence enregistrée — ' + res.name + ' — ' + jour + ' – ' + cours)
      : ('1 cours débité — ' + res.name + ' — solde ' + res.left + '/' + res.total +
         ' — ' + jour + ' – ' + cours);

    return {
      ok:true,
      kind:'ok',
      flashHtml: flashBox_(baseMsg + oweNote, 'ok')
    };

  } catch(e){
    return {
      ok:false,
      kind:'err',
      flashHtml: flashBox_('Erreur liste rapide : ' + (e && e.message ? e.message : String(e)), 'err')
    };
  }
}


// =========================================================
// Aliases (no underscore) for page.html
// =========================================================
function staffSearchAjax(p){ return staffSearchAjax_(p); }
function staffIssueCardAjax(p){ return staffIssueCardAjax_(p); }
function staffIssueTrimAjax(p){ return staffIssueTrimAjax_(p); }
function staffDropinAjax(p){ return staffDropinAjax_(p); }
function staffDoorrowAjax(p){ return staffDoorrowAjax_(p); }





