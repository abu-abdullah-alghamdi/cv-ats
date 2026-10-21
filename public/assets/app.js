/* CV builder app: form, preview, import, analysis, steps, account */
(function(){
const KEY = 'ats-cv-builder-v2';
const DOCX_URL = 'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js';
const $ = s => document.querySelector(s);
const esc = s => String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

const blankEdu = () => ({ degree: '', school: '', dates: '', link: '', linkText: '', details: '' });
const blankProj = () => ({ title: '', dates: '', link: '', linkText: '', format: 'bullets', details: '' });
const blankExp = () => ({ role: '', company: '', dates: '', link: '', linkText: '', format: 'bullets', details: '' });
const blankState = () => ({
  name: '', headline: '', city: '', cityOther: '', showNationality: false, expOrder: 'role', email: '', phone: '', linkedin: '', website: '',
  titles: Object.assign({}, CVEngine.DEFAULT_TITLES), summary: '',
  education: [blankEdu()], skills: '', projects: [blankProj()], experience: [blankExp()], certs: '', languages: ''
});

const CITY_LIST = KSA_CITIES.flatMap(r => r[1].map(c => c[0]));
function normalize(s) {
  const orig = s || {}, b = blankState();
  s = Object.assign(b, orig);
  s.titles = Object.assign({}, CVEngine.DEFAULT_TITLES, s.titles || {});
  ['education', 'projects', 'experience'].forEach(k => { if (!Array.isArray(s[k])) s[k] = []; if (k !== 'experience') s[k].forEach(it => delete it.logo); });
  // older saves kept a free-text address / nationality
  if (s.address && !s.city) {
    const raw = String(s.address).replace(/,?\s*(saudi arabia|ksa)\s*$/i, '').trim();
    const hit = CITY_LIST.find(c => c.toLowerCase() === raw.toLowerCase());
    if (hit) s.city = hit; else if (raw) { s.city = 'Other'; s.cityOther = raw; }
  }
  if (String(orig.nationality || '').trim() && !('showNationality' in orig)) s.showNationality = true;
  delete s.address; delete s.nationality;
  return s;
}
/* what actually goes on the CV */
function cvData() {
  const city = state.city === 'Other' ? String(state.cityOther || '').trim() : state.city;
  return Object.assign({}, state, {
    address: city ? city + ', Saudi Arabia' : '',
    nationality: state.showNationality ? 'Saudi' : ''
  });
}
let state = null;
try { const raw = localStorage.getItem(KEY); if (raw) state = JSON.parse(raw); } catch (e) {}
state = normalize(state);

/* ---------- text measuring (same metrics as the PDF) ---------- */
let mdoc = null;
try { mdoc = new window.jspdf.jsPDF({ unit: 'pt' }); } catch (e) {}
const cctx = document.createElement('canvas').getContext('2d');
function measure(t, size, style) {
  style = style || 'normal';
  if (mdoc) { mdoc.setFont('helvetica', style); mdoc.setFontSize(size); return mdoc.getTextWidth(t); }
  cctx.font = (style.includes('italic') ? 'italic ' : '') + (style.includes('bold') ? '700 ' : '400 ') + size + 'px Arimo, Helvetica, Arial, sans-serif';
  return cctx.measureText(t).width;
}

const getP = p => p.split('.').reduce((o, k) => (o == null ? o : o[k]), state);
function setP(p, v) { const ks = p.split('.'); let o = state; for (let i = 0; i < ks.length - 1; i++) o = o[ks[i]]; o[ks[ks.length - 1]] = v; }

let saveT;
function save() {
  clearTimeout(saveT);
  saveT = setTimeout(() => { try { localStorage.setItem(KEY, JSON.stringify(state)); } catch (e) {} }, 300);
  cloudSaveSoon();
}

/* ---------- form ---------- */
function field(path, label, ph, kind, opts) {
  opts = opts || {};
  const id = 'f-' + path.replace(/\./g, '-');
  const v = getP(path) == null ? '' : getP(path);
  const cls = opts.wide || kind === 'area' ? 'f wide' : 'f';
  const lab = `<label for="${id}">${label}${opts.optional ? ' <span>(اختياري)</span>' : ''}</label>`;
  let ctl;
  if (kind === 'area') ctl = `<textarea id="${id}" data-path="${path}" class="ltr" dir="auto" rows="${opts.rows || 4}" placeholder="${esc(ph)}">${esc(v)}</textarea>`;
  else if (kind === 'format') ctl = `<select id="${id}" data-path="${path}"><option value="bullets"${v !== 'paragraph' ? ' selected' : ''}>نقاط: كل سطر نقطة</option><option value="paragraph"${v === 'paragraph' ? ' selected' : ''}>فقرة واحدة متصلة</option></select>`;
  else if (opts.prefix) ctl = `<div class="prefixed" dir="ltr"><span dir="ltr" aria-hidden="true">${opts.prefix}</span><input id="${id}" data-path="${path}" type="${kind || 'text'}" inputmode="numeric" class="ltr" value="${esc(CVEngine.localPhone(v))}" placeholder="${esc(ph)}" autocomplete="off" aria-label="${label}، يبدأ بمفتاح السعودية ${opts.prefix}"></div>`;
  else ctl = `<input id="${id}" data-path="${path}" type="${kind || 'text'}" class="ltr" dir="auto" value="${esc(v)}" placeholder="${esc(ph)}" autocomplete="off">`;
  return `<div class="${cls}">${lab}${ctl}${opts.help ? `<small>${opts.help}</small>` : ''}</div>`;
}
function cityField() {
  const v = state.city || '';
  const groups = KSA_CITIES.map(([region, cities]) => `<optgroup label="${region}">${cities.map(([en, ar]) =>
    `<option value="${en}"${v === en ? ' selected' : ''}>${ar} · ${en}</option>`).join('')}</optgroup>`).join('');
  return `<div class="f"><label for="f-city">المدينة</label>
    <select id="f-city" data-path="city"><option value=""${v ? '' : ' selected'}>اختر مدينتك</option><option value="Other"${v === 'Other' ? ' selected' : ''}>Other · مدينة أخرى</option>${groups}</select>
    <div class="sub" id="city-other-wrap"${v === 'Other' ? '' : ' hidden'}><input id="f-cityOther" data-path="cityOther" type="text" class="ltr" dir="auto" value="${esc(state.cityOther || '')}" placeholder="City name" autocomplete="off" aria-label="اكتب اسم مدينتك"></div></div>`;
}
function nationalityField() {
  const on = !!state.showNationality;
  return `<div class="f"><label for="f-nationality">الجنسية</label>
    <input id="f-nationality" type="text" class="ltr" value="Saudi" readonly>
    <button type="button" class="btn toggle" data-act="nat" aria-pressed="${on}">${on ? 'إخفاء الجنسية' : 'إظهار الجنسية'}</button></div>`;
}
const titleField = key => field('titles.' + key, 'اسم القسم في السيرة', CVEngine.DEFAULT_TITLES[key], 'text', { wide: true, help: 'تقدر تتركه كما هو.' });

const LISTS = {
  education: {
    name: 'مؤهل', add: 'إضافة مؤهل آخر', blank: blankEdu,
    body: i => { const p = 'education.' + i + '.'; return `<div class="fields">
      ${field(p + 'degree', 'الشهادة والتخصص', 'Bachelor of Science in Electrical Engineering', 'text', { wide: true })}
      ${field(p + 'school', 'الجامعة', 'King Saud University')}
      ${field(p + 'dates', 'سنة البداية والتخرج', '2019 – 2023')}
      ${field(p + 'details', 'معلومات إضافية', 'GPA: 4.5 / 5\nMajor: Power Engineering', 'area', { rows: 3, optional: true, help: 'مثل المعدل أو التخصص الدقيق. كل معلومة في سطر.' })}
      ${field(p + 'link', 'رابط', 'https://…', 'url', { optional: true, help: 'لو عندك صفحة توضح شهادتك.' })}
      ${field(p + 'linkText', 'الكلمة اللي تظهر على الرابط', 'See more', 'text', { optional: true })}
    </div>`; }
  },
  projects: {
    name: 'مشروع', add: 'إضافة مشروع آخر', blank: blankProj,
    body: i => { const p = 'projects.' + i + '.'; return `<div class="fields">
      ${field(p + 'title', 'اسم المشروع', 'Graduation Project: Smart Energy Meter', 'text', { wide: true })}
      ${field(p + 'dates', 'تاريخ المشروع', 'Jan 2023 – Jun 2023')}
      ${field(p + 'format', 'شكل الوصف', '', 'format')}
      ${field(p + 'details', 'وش سويت في المشروع؟', 'Designed a smart meter that tracks home energy use.\nTested the device and improved its accuracy by 15%.', 'area', { rows: 4, help: 'اكتب كل شي سويته في سطر لحاله. مثال:<span class="ex">Designed a smart meter that tracks home energy use.<br>Tested the device and improved its accuracy.</span>' })}
      ${field(p + 'link', 'رابط المشروع', 'github.com/anas/project', 'url', { optional: true })}
      ${field(p + 'linkText', 'الكلمة اللي تظهر على الرابط', 'See more', 'text', { optional: true })}
    </div>`; }
  },
  experience: {
    name: 'خبرة', add: 'إضافة خبرة أخرى', blank: blankExp,
    body: i => { const p = 'experience.' + i + '.', lg = state.experience[i].logo; return `
      <div class="logo-row">
        <div class="logo-box${lg ? ' has' : ''}">${lg ? `<img src="${lg.src}" alt="شعار الشركة">` : 'بدون شعار'}</div>
        <span class="btn filebtn">${lg ? 'تغيير الشعار' : 'إضافة شعار الشركة'}<input type="file" accept="image/png,image/jpeg,image/webp" id="logo-${i}" data-logo="${i}" aria-label="شعار الشركة"></span>
        ${lg ? `<button type="button" class="btn" data-act="logo-del" data-i="${i}">إزالة الشعار</button>` : ''}
      </div>
      <p class="logo-rules">اختياري. صورة PNG أو JPG أو WEBP أقل من 2 ميجا، بخلفية بيضاء أو بدون خلفية.</p>
      <p class="logo-err" id="logo-err-${i}" hidden></p>
      <div class="fields">
      ${field(p + 'role', 'المسمى الوظيفي', 'Electrical Engineer')}
      ${field(p + 'company', 'اسم الشركة', 'ABC Company')}
      ${field(p + 'dates', 'من متى إلى متى', 'Jan 2024 – Present', 'text', { help: 'لو للحين تشتغل فيها اكتب Present' })}
      ${field(p + 'format', 'شكل الوصف', '', 'format')}
      ${field(p + 'details', 'وش كانت مهامك وإنجازاتك؟', 'Supervised electrical work for 5 buildings.\nReviewed drawings and approved materials.', 'area', { rows: 5, help: 'اكتب كل مهمة في سطر لحالها. مثال:<span class="ex">Supervised electrical work for 5 buildings.<br>Reviewed drawings and approved materials.</span>' })}
      ${field(p + 'link', 'رابط المشروع أو الشركة', 'https://…', 'url', { optional: true })}
      ${field(p + 'linkText', 'الكلمة اللي تظهر على الرابط', 'See project', 'text', { optional: true })}
    </div>`; }
  }
};

function renderList(key) {
  const L = LISTS[key], arr = state[key];
  document.getElementById('list-' + key).innerHTML = arr.map((it, i) => `
    <div class="item">
      <div class="item-head"><b>${L.name} ${i + 1}</b>
        <div class="tools">
          <button type="button" class="btn icon" data-act="up" data-k="${key}" data-i="${i}" ${i === 0 ? 'disabled' : ''} aria-label="نقل للأعلى">↑</button>
          <button type="button" class="btn icon" data-act="down" data-k="${key}" data-i="${i}" ${i === arr.length - 1 ? 'disabled' : ''} aria-label="نقل للأسفل">↓</button>
          <button type="button" class="btn danger" data-act="del" data-k="${key}" data-i="${i}">حذف</button>
        </div></div>
      ${L.body(i)}
    </div>`).join('') + `<button type="button" class="btn add" data-act="add" data-k="${key}">+ ${L.add}</button>`;
}
function sec(title, sub, body, open) {
  return `<details class="sec"${open ? ' open' : ''}><summary><span>${title}${sub ? `<small>${sub}</small>` : ''}</span></summary><div class="sec-body">${body}</div></details>`;
}
function renderEditor() {
  $('#editor').innerHTML =
    sec('معلوماتك الشخصية', 'تظهر بأعلى السيرة', `<div class="fields">
      ${field('name', 'اسمك', 'Anas', 'text', { wide: true })}
      ${field('headline', 'مسماك الوظيفي أو تخصصك', 'Electrical Engineer | Project Coordinator', 'text', { wide: true })}
      ${field('email', 'الإيميل', 'example@gmail.com', 'email')}
      ${field('phone', 'رقم الجوال', '5XXXXXXXX', 'tel', { prefix: '+966' })}
      ${field('linkedin', 'رابط حسابك في لنكدإن', 'linkedin.com/in/anas', 'url', { optional: true })}
      ${field('website', 'موقعك أو أعمالك', 'github.com/anas', 'url', { optional: true })}
      ${cityField()}
      ${nationalityField()}
    </div>`, true) +
    sec('نبذة عنك', 'فقرة قصيرة في بداية السيرة', `<div class="fields">
      ${titleField('summary')}
      ${field('summary', 'عرّف بنفسك في 3 إلى 5 جمل', 'Electrical engineer with 3 years of experience in site supervision. Skilled in reviewing drawings, leading teams, and delivering projects on time.', 'area', { rows: 5, help: 'مين أنت، وش خبرتك، وش تقدر تضيف للشركة.' })}</div>`) +
    sec('التعليم', 'شهاداتك الجامعية', `<div class="fields">${titleField('education')}</div><div class="items" id="list-education"></div>`) +
    sec('المهارات', 'الأشياء اللي تجيدها', `<div class="fields">
      ${titleField('skills')}
      ${field('skills', 'مهاراتك', 'Technical: AutoCAD, MATLAB, Python\nSoft Skills: Teamwork, Communication', 'area', { rows: 5, help: 'كل سطر يظهر كنقطة. لو كتبت كلمة وبعدها نقطتين ( : ) تظهر الكلمة بخط عريض.' })}</div>`) +
    sec('المشاريع', 'مثل مشروع التخرج', `<div class="fields">${titleField('projects')}</div><div class="items" id="list-projects"></div>`) +
    sec('الخبرات', 'وظائفك وتدريبك', `<div class="fields">${titleField('experience')}
      <div class="f wide"><label for="f-expOrder">ترتيب سطر الخبرة في السيرة</label>
        <select id="f-expOrder" data-path="expOrder"><option value="role"${state.expOrder !== 'company' ? ' selected' : ''}>المسمى | الشركة   (Engineer | ABC Company)</option><option value="company"${state.expOrder === 'company' ? ' selected' : ''}>الشركة | المسمى   (ABC Company | Engineer)</option></select></div></div><p class="note">ابدأ بأحدث وظيفة. تقدر ترتبها بالأسهم ↑ ↓.</p><div class="items" id="list-experience"></div>`) +
    sec('الشهادات والدورات', '', `<div class="fields">
      ${titleField('certs')}
      ${field('certs', 'شهاداتك ودوراتك', 'PMP: Project Management Professional\nCCNA: Cisco Certified Network Associate', 'area', { rows: 4, help: 'كل شهادة في سطر.' })}</div>`) +
    sec('اللغات', '', `<div class="fields">
      ${titleField('languages')}
      ${field('languages', 'اللغات ومستواك فيها', 'Arabic: Native\nEnglish: Fluent', 'area', { rows: 3, help: 'كل لغة في سطر.' })}</div>`);
  Object.keys(LISTS).forEach(renderList);
  document.querySelectorAll('#editor > details.sec').forEach((d, i) => { if (SEC_IDS[i]) d.id = 'sec-' + SEC_IDS[i]; });
}

/* ---------- preview ---------- */
function opToSvg(op) {
  if (op.k === 'text') {
    const fw = op.s.includes('bold') ? 700 : 400, fs = op.s.includes('italic') ? 'italic' : 'normal';
    return `<text x="${op.x.toFixed(2)}" y="${op.y.toFixed(2)}" font-size="${op.size}" font-weight="${fw}" font-style="${fs}" fill="${op.c}"${op.ws ? ` word-spacing="${op.ws.toFixed(3)}"` : ''}>${esc(op.t).replace(/^ +| +$/g, m => '\u00A0'.repeat(m.length)).replace(/  +/g, m => ' ' + '\u00A0'.repeat(m.length - 1))}</text>`;
  }
  if (op.k === 'line') return `<line x1="${op.x1}" y1="${op.y1}" x2="${op.x2}" y2="${op.y2}" stroke="${op.c}" stroke-width="${op.w}"/>`;
  if (op.k === 'dot') return `<circle cx="${op.x}" cy="${op.y}" r="${op.r}" fill="#141414"/>`;
  if (op.k === 'img') return `<image href="${op.src}" x="${op.x}" y="${op.y}" width="${op.w}" height="${op.h}" preserveAspectRatio="xMinYMid meet"/>`;
  if (op.k === 'link') return `<a href="${esc(op.url)}" target="_blank" rel="noopener"><rect x="${op.x}" y="${op.y}" width="${op.w}" height="${op.h}" fill="transparent"/></a>`;
  return '';
}
function pageSvg(ops, r) {
  return `<svg class="page" viewBox="0 0 ${r.PW} ${r.PH}" xmlns="http://www.w3.org/2000/svg" text-rendering="geometricPrecision" direction="ltr" xml:space="preserve" style="direction:ltr;unicode-bidi:isolate;white-space:pre;font-family:Arimo,Helvetica,Arial,sans-serif">
    <rect width="100%" height="100%" fill="#ffffff"/>${ops.map(opToSvg).join('')}</svg>`;
}
function renderPreview() {
  const cv = cvData();
  const empty = CVEngine.isEmpty(cv);
  const r = CVEngine.layout(cv, measure);
  $('#pages').innerHTML = r.pages.map(ops => `<div class="page-wrap">${pageSvg(ops, r)}${empty ? '<div class="empty-hint"><strong>الصفحة فاضية</strong><span>ابدأ بكتابة اسمك في البيانات، وبتظهر سيرتك هنا خطوة بخطوة.</span></div>' : ''}</div>`).join('');
  const n = r.pages.length;
  $('#pcount').textContent = empty ? '' : (n === 1 ? 'صفحة واحدة' : n === 2 ? 'صفحتان' : n + ' صفحات');
  renderAnalysis(empty, n);
}
/* ---------- CV analysis: what is missing compared with the example ---------- */
let linkedinHint = false;
const SEC_IDS = ['personal', 'summary', 'education', 'skills', 'projects', 'experience', 'certs', 'languages', 'backup'];
const words = s => String(s || '').trim().split(/\s+/).filter(Boolean).length;
function analyze(pages) {
  const S = state, out = [];
  const has = v => String(v || '').trim().length > 0;
  const add = (ok, level, text, sec, field) => out.push({ ok, level, text, sec, field });
  // personal
  add(has(S.name), 'must', 'أضف اسمك.', 'personal', 'f-name');
  add(has(S.headline), 'must', 'أضف مسماك الوظيفي تحت اسمك، مثل: ⟦ELECTRICAL ENGINEER | Electronics & Communications⟧.', 'personal', 'f-headline');
  if (!has(S.email)) add(false, 'must', 'أضف إيميلك عشان يتواصلون معك.', 'personal', 'f-email');
  else add(/^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(S.email.trim()), 'must', 'الإيميل مكتوب غلط. تأكد إنه مثل ⟦example@gmail.com⟧.', 'personal', 'f-email');
  const ph = CVEngine.localPhone(S.phone);
  if (!ph) add(false, 'must', 'أضف رقم جوالك.', 'personal', 'f-phone');
  else add(/^5\d{8}$/.test(ph), 'must', 'رقم الجوال لازم يكون 9 أرقام ويبدأ بـ5، مثل ⟦501234567⟧.', 'personal', 'f-phone');
  add(has(S.linkedin), 'must', linkedinHint ? 'سيرتك القديمة فيها لنكدإن بس بدون رابط واضح. انسخ رابط حسابك وحطه في خانة لنكدإن.' : 'أضف رابط حسابك في لنكدإن. أغلب الشركات تشوفه قبل المقابلة.', 'personal', 'f-linkedin');
  add(has(S.city) && (S.city !== 'Other' || has(S.cityOther)), 'must', 'اختر مدينتك.', 'personal', 'f-city');
  // summary
  const sw = words(S.summary);
  if (!sw) add(false, 'must', 'أضف نبذة عنك من 3 إلى 5 جمل: مين أنت، وش خبرتك، وش تقدر تضيف.', 'summary', 'f-summary');
  else if (sw < 25) add(false, 'tip', `النبذة قصيرة (${sw} كلمة). خلها بين 40 و90 كلمة مثل المثال.`, 'summary', 'f-summary');
  else if (sw > 120) add(false, 'tip', `النبذة طويلة (${sw} كلمة). اختصرها لأقل من 100 كلمة.`, 'summary', 'f-summary');
  else add(true, 'must', 'النبذة بطول مناسب.', 'summary', 'f-summary');
  // education
  const ed = (S.education || []).filter(e => has(e.degree) || has(e.school));
  if (!ed.length) add(false, 'must', 'أضف شهادتك الجامعية في قسم التعليم.', 'education', 'f-education-0-degree');
  else ed.forEach(e => {
    const i = S.education.indexOf(e);
    if (!has(e.school)) add(false, 'must', `أضف اسم الجامعة لشهادة «${e.degree}».`, 'education', `f-education-${i}-school`);
    if (!has(e.dates)) add(false, 'must', `أضف سنة البداية والتخرج لـ«${e.degree || e.school}»، مثل ⟦2019 – 2023⟧.`, 'education', `f-education-${i}-dates`);
    if (!has(e.details)) add(false, 'tip', 'أضف تخصصك الدقيق ومعدلك تحت الشهادة، مثل: ⟦GPA: 4.2 / 5.0⟧.', 'education', `f-education-${i}-details`);
  });
  if (ed.length && ed.every(e => has(e.school) && has(e.dates))) add(true, 'must', 'التعليم مكتمل.', 'education');
  // skills
  const sk = CVEngine.lines(S.skills);
  if (!sk.length) add(false, 'must', 'أضف مهاراتك. اكتب كل فئة في سطر، مثل: ⟦Technical: AutoCAD, Python⟧.', 'skills', 'f-skills');
  else {
    add(sk.length >= 3, 'tip', `عندك ${sk.length === 1 ? 'سطر واحد' : sk.length === 2 ? 'سطرين' : sk.length + ' أسطر'} مهارات بس. المثال فيه 4 فئات، أضف مهارات أكثر.`, 'skills', 'f-skills');
    add(sk.some(l => /^[^:]{2,50}:/.test(l)), 'tip', 'رتّب المهارات بفئات، اكتب اسم الفئة وبعدها نقطتين، مثل: ⟦Soft Skills: Teamwork, Communication⟧.', 'skills', 'f-skills');
  }
  // experience
  const ex = (S.experience || []).filter(e => has(e.role) || has(e.company) || has(e.details));
  if (!ex.length) add(false, 'must', 'أضف خبراتك أو تدريبك الصيفي. حتى التدريب القصير يفرق.', 'experience', 'f-experience-0-role');
  ex.forEach(e => {
    const i = S.experience.indexOf(e), nm = e.company || e.role || ('خبرة ' + (i + 1));
    if (!has(e.role)) add(false, 'must', `أضف المسمى الوظيفي في «${nm}».`, 'experience', `f-experience-${i}-role`);
    if (!has(e.company)) add(false, 'must', `أضف اسم الشركة للمسمى «${e.role}».`, 'experience', `f-experience-${i}-company`);
    if (!has(e.dates)) add(false, 'must', `أضف الفترة في «${nm}»، مثل ⟦Jan 2024 – Present⟧.`, 'experience', `f-experience-${i}-dates`);
    const pts = CVEngine.lines(e.details);
    if (!pts.length) add(false, 'must', `اكتب مهامك وإنجازاتك في «${nm}».`, 'experience', `f-experience-${i}-details`);
    else if (e.format !== 'paragraph' && pts.length < 2) add(false, 'tip', `في «${nm}» نقطة وحدة بس. أضف نقطتين أو ثلاث مثل المثال.`, 'experience', `f-experience-${i}-details`);
    if (pts.some(p => words(p) > 40)) add(false, 'tip', `في «${nm}» نقطة طويلة جدًا. قسّمها لنقطتين أقصر.`, 'experience', `f-experience-${i}-details`);
  });
  if (ex.length) {
    add(ex.some(e => /\d/.test(e.details || '')), 'tip', 'أضف أرقام توضح حجم شغلك، مثل: ⟦4 buildings, 120 units, 5 engineers⟧.', 'experience', `f-experience-${S.experience.indexOf(ex[0])}-details`);
    add(ex.some(e => e.logo), 'tip', 'تقدر تضيف شعار كل شركة فوق خبرتك مثل المثال (اختياري).', 'experience');
  }
  // projects, certs, languages
  const pr = (S.projects || []).filter(p => has(p.title));
  add(pr.length > 0, 'tip', 'أضف مشروع تخرجك أو أي مشروع سويته، خصوصًا لو خبرتك قليلة.', 'projects', 'f-projects-0-title');
  add(CVEngine.lines(S.certs).length > 0, 'tip', 'أضف شهاداتك المهنية أو الدورات اللي أخذتها.', 'certs', 'f-certs');
  add(CVEngine.lines(S.languages).length > 0, 'must', 'أضف اللغات ومستواك فيها، مثل: ⟦Arabic: Native⟧.', 'languages', 'f-languages');
  // whole CV
  const arabic = /[؀-ۿ]/.test(JSON.stringify(Object.assign({}, state, { titles: null })));
  add(!arabic, 'must', 'فيه كلام عربي ما راح يظهر في الملف. اكتب كل شي بالإنجليزي.', null);
  add(pages <= 2, 'must', `سيرتك ${pages} صفحات. اختصرها لصفحة أو صفحتين.`, null);
  return out;
}
const notesWord = n => n === 1 ? 'ملاحظة وحدة' : n === 2 ? 'ملاحظتين' : n + (n <= 10 ? ' ملاحظات' : ' ملاحظة');
let anHidden = false;
try { anHidden = localStorage.getItem('ats-cv-builder-hide-analysis') === '1'; } catch (e) {}
function renderAnalysis(empty, pages) {
  $('#an-off').hidden = !anHidden;
  $('#an-empty').hidden = anHidden || !empty;
  $('#an-main').hidden = anHidden || empty;
  if (anHidden || empty) { $('#score-chip').hidden = true; return; }
  const res = analyze(pages);
  const w = r => (r.level === 'must' ? 3 : 1);
  const total = res.reduce((a, r) => a + w(r), 0), got = res.reduce((a, r) => a + (r.ok ? w(r) : 0), 0);
  const score = total ? Math.round(100 * got / total) : 0;
  const bad = res.filter(r => !r.ok).sort((a, b) => (a.level === b.level ? 0 : a.level === 'must' ? -1 : 1));
  const good = res.filter(r => r.ok);
  $('#an-score').textContent = score + '%';
  $('#an-score').dataset.level = score >= 85 ? 'ok' : score >= 60 ? 'mid' : 'low';
  $('#an-bar').style.width = score + '%';
  $('#an-bar').dataset.level = $('#an-score').dataset.level;
  $('#an-sub').textContent = bad.length ? notesWord(bad.length) + ' عشان تصير مثل السيرة النموذجية' : 'سيرتك مكتملة مثل السيرة النموذجية';
  $('#an-list').innerHTML = bad.map(r => `<li class="an-item ${r.level}">
      <span class="an-tag">${r.level === 'must' ? 'ناقص' : 'تحسين'}</span>
      <span class="an-text">${esc(r.text).replace(/⟦(.*?)⟧/g, '<bdi dir="ltr" class="en">$1</bdi>')}</span>
      ${r.sec ? `<button type="button" class="btn an-go" data-act="goto" data-sec="${r.sec}" data-field="${r.field || ''}">عدّلها</button>` : ''}
    </li>`).join('') + (good.length ? `<li class="an-good">✓ ${good.length === 1 ? 'شي واحد مكتمل' : good.length === 2 ? 'شيئين مكتملين' : good.length + (good.length <= 10 ? ' أشياء مكتملة' : ' شي مكتمل')}</li>` : '');
  $('#score-chip').hidden = false;
  $('#chip-score').textContent = score + '%';
  $('#chip-score').dataset.level = $('#an-score').dataset.level;
  $('#chip-sub').textContent = bad.length ? notesWord(bad.length) : 'مكتملة';
}
/* ---------- steps ---------- */
let curStep = 1;
function goStep(n, opts) {
  opts = opts || {};
  curStep = n;
  [1, 2, 3].forEach(i => { $('#step-' + i).hidden = i !== n; });
  document.querySelectorAll('.stepper .st').forEach(b => {
    const i = +b.dataset.step;
    b.classList.toggle('done', i < n);
    if (i === n) b.setAttribute('aria-current', 'step'); else b.removeAttribute('aria-current');
  });
  if (n === 3) { renderPreview(); }
  if (!opts.noScroll) $('#stepper').scrollIntoView({ block: 'start', behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' });
  if (opts.focus) setTimeout(() => {
    if (opts.focus === 'f-name') { const d = document.getElementById('sec-personal'); if (d) d.open = true; }
    const f = document.getElementById(opts.focus); if (f) f.focus({ preventScroll: true });
  }, 350);
}
function goToField(secId, fieldId) {
  if (curStep !== 3) goStep(3, { noScroll: true });
  if (document.body.dataset.view === 'preview') { const tab = document.querySelector('.tabs [data-view="edit"]'); if (tab) tab.click(); }
  const d = document.getElementById('sec-' + secId);
  if (d) d.open = true;
  const f = fieldId && document.getElementById(fieldId);
  const target = f || d;
  if (!target) return;
  target.scrollIntoView({ block: 'center', behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' });
  if (f) setTimeout(() => f.focus({ preventScroll: true }), 250);
}
let pvT;
function schedule() { clearTimeout(pvT); pvT = setTimeout(renderPreview, 120); save(); }

const ZOOMS = [1, 1.25, 1.5, 2, 2.5];
let zi = 0;
function setZoom(i) {
  zi = Math.max(0, Math.min(ZOOMS.length - 1, i));
  $('#pages').style.setProperty('--z', ZOOMS[zi]);
  $('#zv').textContent = Math.round(ZOOMS[zi] * 100) + '%';
}

/* ---------- logo checks ---------- */
function logoError(msg) { const e = new Error(msg); e.userMsg = msg; return e; }
function loadImage(file) {
  return new Promise((res, rej) => {
    const url = URL.createObjectURL(file), img = new Image();
    img.onload = () => { URL.revokeObjectURL(url); res(img); };
    img.onerror = () => { URL.revokeObjectURL(url); rej(logoError('ما قدرنا نفتح الصورة. جرّب صورة PNG أو JPG.')); };
    img.src = url;
  });
}
async function processLogo(file) {
  if (!/^image\/(png|jpeg|webp)$/.test(file.type)) throw logoError('نوع الصورة غير مدعوم. استخدم PNG أو JPG أو WEBP.');
  if (file.size > 2 * 1024 * 1024) throw logoError('حجم الصورة أكبر من 2 ميجا. صغّرها وجرّب مرة ثانية.');
  const img = await loadImage(file);
  if (img.naturalHeight < 60) throw logoError('الصورة صغيرة جدًا وبتطلع مشوشة. استخدم صورة ارتفاعها 60 بكسل أو أكثر.');
  const k = Math.min(1, 1000 / Math.max(img.naturalWidth, img.naturalHeight));
  const W = Math.max(1, Math.round(img.naturalWidth * k)), H = Math.max(1, Math.round(img.naturalHeight * k));
  const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
  const g = cv.getContext('2d', { willReadFrequently: true }); g.drawImage(img, 0, 0, W, H);
  const px = g.getImageData(0, 0, W, H).data;
  const blank = i => px[i + 3] < 24 || (px[i] > 232 && px[i + 1] > 232 && px[i + 2] > 232);
  // 1) background: the outer edge must be white or transparent
  let edge = 0, bad = 0;
  const probe = (x, y) => { edge++; if (!blank((y * W + x) * 4)) bad++; };
  for (let x = 0; x < W; x++) { probe(x, 0); probe(x, H - 1); }
  for (let y = 1; y < H - 1; y++) { probe(0, y); probe(W - 1, y); }
  if (bad / edge > 0.12) throw logoError('الشعار لازم يكون بخلفية بيضاء أو بدون خلفية. جرّب نسخة ثانية من الشعار.');
  // 2) trim empty space around the mark
  let x0 = W, y0 = H, x1 = -1, y1 = -1;
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    if (!blank((y * W + x) * 4)) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  }
  if (x1 < 0) throw logoError('الصورة فاضية. اختر صورة فيها الشعار.');
  const pad = 2; x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad); x1 = Math.min(W - 1, x1 + pad); y1 = Math.min(H - 1, y1 + pad);
  const tw = x1 - x0 + 1, th = y1 - y0 + 1;
  // 3) proportions
  if (tw / th > 5) throw logoError('الشعار عريض جدًا. استخدم نسخة أقصر، مثل الرمز مع الاسم بسطر واحد.');
  if (th / tw > 1.5) throw logoError('الشعار طويل جدًا. استخدم نسخة أفقية أو مربعة من الشعار.');
  if (th < 24) throw logoError('الشعار نفسه صغير جدًا داخل الصورة. استخدم صورة أوضح.');
  // 4) one standard height, flattened on white
  const OH = 120, OW = Math.round(tw * OH / th);
  const out = document.createElement('canvas'); out.width = OW; out.height = OH;
  const o = out.getContext('2d'); o.fillStyle = '#ffffff'; o.fillRect(0, 0, OW, OH);
  o.imageSmoothingQuality = 'high'; o.drawImage(cv, x0, y0, tw, th, 0, 0, OW, OH);
  return { src: out.toDataURL('image/png'), w: OW, h: OH };
}

/* ---------- downloads ---------- */
const setStatus = t => { $('#status').textContent = t; };
function fileBase() {
  const n = String(state.name || '').trim().replace(/[^\w\- ]+/g, '').replace(/\s+/g, '_');
  return n || 'My';
}
async function saveFile(filename, data, okMsg) {
  try {
    const blob = data instanceof Blob ? data : new Blob([data]);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.rel = 'noopener';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    setStatus(okMsg);
  } catch (e) {
    setStatus('ما قدرنا نحمّل الملف. جرّب مرة ثانية.');
  }
}
function needContent() {
  if (CVEngine.isEmpty(cvData())) { setStatus('اكتب بياناتك أولًا، ثم حمّل السيرة.'); return true; }
  return false;
}
function loadScript(src) {
  return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = rej; document.head.appendChild(s); });
}
$('#dl-pdf').addEventListener('click', async () => {
  if (needContent()) return;
  if (!window.jspdf) { setStatus('ما اكتمل تحميل الصفحة. تأكد من الإنترنت وافتحها من جديد.'); return; }
  const b = $('#dl-pdf'); b.disabled = true; setStatus('نجهّز ملف PDF…');
  try {
    const r = CVEngine.layout(cvData(), measure);
    const doc = CVEngine.toPDF(r, window.jspdf.jsPDF, { title: (state.name || '') + ' – CV', author: state.name || '', keywords: CVEngine.lines(state.skills).join(', ').slice(0, 900) });
    await saveFile(fileBase() + '_CV.pdf', doc.output('blob'), 'تم تحميل السيرة PDF.');
  } catch (e) { setStatus('صار خطأ أثناء تجهيز الملف. جرّب مرة ثانية.'); }
  b.disabled = false;
});
$('#dl-word').addEventListener('click', async () => {
  if (needContent()) return;
  const b = $('#dl-word'); b.disabled = true; setStatus('نجهّز ملف Word…');
  try {
    if (!window.docx) await loadScript(DOCX_URL);
    const blob = await window.docx.Packer.toBlob(CVEngine.toDocx(cvData(), window.docx));
    await saveFile(fileBase() + '_CV.docx', blob, 'تم تحميل السيرة Word.');
  } catch (e) { setStatus('ما قدرنا نجهّز ملف Word. تأكد من الإنترنت وجرّب مرة ثانية.'); }
  b.disabled = false;
});

/* ---------- events ---------- */
const ed = $('#editor');
/* English only: Arabic digits become 0-9, Arabic letters are removed with a warning under the field */
const AR_TEST = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
const AR_ALL = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/g;
const latinDigits = s => s.replace(/[\u0660-\u0669]/g, d => String(d.charCodeAt(0) - 0x660)).replace(/[\u06F0-\u06F9]/g, d => String(d.charCodeAt(0) - 0x6F0));
function warnArabic(el) {
  const f = el.closest('.f') || el.parentElement;
  let w = f.querySelector(':scope > .ar-warn');
  if (!w) { w = document.createElement('small'); w.className = 'ar-warn'; w.setAttribute('role', 'alert'); w.textContent = 'اكتب بياناتك باللغة الإنجليزية'; f.appendChild(w); }
  w.hidden = false;
  el.classList.add('ar-bad');
  clearTimeout(w._t);
  w._t = setTimeout(() => { w.hidden = true; el.classList.remove('ar-bad'); }, 4000);
}
function englishOnly(el) {
  if (!('value' in el) || el.tagName === 'SELECT') return;
  const before = el.value;
  let v = latinDigits(before);
  if (AR_TEST.test(v)) { v = v.replace(AR_ALL, '').replace(/ {2,}/g, ' '); warnArabic(el); }
  if (v !== before) {
    let pos = null;
    try { pos = el.selectionStart - (before.length - v.length); } catch (err) {}
    el.value = v;
    if (pos != null) { try { el.setSelectionRange(Math.max(0, pos), Math.max(0, pos)); } catch (err) {} }
  }
}
ed.addEventListener('input', e => {
  const t = e.target, p = t.dataset && t.dataset.path;
  if (!p) return;
  englishOnly(t);
  if (p === 'phone') { const d = CVEngine.localPhone(t.value).slice(0, 9); if (t.value !== d) t.value = d; setP(p, d); }
  else setP(p, t.value);
  schedule();
});
ed.addEventListener('change', async e => {
  const t = e.target;
  if (t.tagName === 'SELECT' && t.dataset.path) {
    setP(t.dataset.path, t.value); schedule();
    if (t.id === 'f-city') { const w = $('#city-other-wrap'); w.hidden = t.value !== 'Other'; if (!w.hidden) $('#f-cityOther').focus(); }
  }
  if (t.dataset && t.dataset.logo != null && t.files && t.files[0]) {
    const i = +t.dataset.logo, file = t.files[0];
    try {
      state.experience[i].logo = await processLogo(file);
      renderList('experience'); schedule(); setStatus('تمت إضافة الشعار.');
    } catch (err) {
      t.value = '';
      const box = $('#logo-err-' + i);
      const msg = err && err.userMsg ? err.userMsg : 'ما قدرنا نقرأ الصورة. جرّب صورة PNG أو JPG.';
      if (box) { box.textContent = msg; box.hidden = false; }
      setStatus(msg);
    }
  }
  if (t.id === 'import-file' && t.files && t.files[0]) {
    try {
      const data = JSON.parse(await t.files[0].text());
      if (!data || typeof data !== 'object' || !('name' in data)) throw new Error('shape');
      state = normalize(data); renderEditor(); schedule(); setStatus('تم فتح النسخة.');
    } catch (err) { setStatus('هذا الملف ما ينفع. اختر ملف نسخة محفوظ من هذا الموقع.'); }
    t.value = '';
  }
});
function arm(btn, run) {
  if (btn.dataset.armed === '1') { run(); return; }
  btn.dataset.armed = '1'; const old = btn.textContent; btn.textContent = 'متأكد؟ اضغط مرة ثانية';
  setTimeout(() => { if (btn.isConnected) { btn.dataset.armed = ''; btn.textContent = old; } }, 3000);
}
document.addEventListener('click', e => {
  const b = e.target.closest('[data-act]');
  if (!b || b.disabled) return;
  const act = b.dataset.act, k = b.dataset.k, i = +b.dataset.i;
  if (act === 'view') {
    document.body.dataset.view = b.dataset.view;
    document.querySelectorAll('.tabs button').forEach(x => x.setAttribute('aria-selected', String(x === b)));
    window.scrollTo(0, 0);
  } else if (act === 'zoom') setZoom(zi + (+b.dataset.d));
  else if (act === 'add') {
    state[k].push(LISTS[k].blank()); renderList(k); schedule();
    const items = document.querySelectorAll('#list-' + k + ' .item'); const f = items[items.length - 1] && items[items.length - 1].querySelector('input,textarea'); f && f.focus();
  } else if (act === 'up' || act === 'down') {
    const j = act === 'up' ? i - 1 : i + 1, a = state[k];
    if (j < 0 || j >= a.length) return;
    [a[i], a[j]] = [a[j], a[i]]; renderList(k); schedule();
  } else if (act === 'del') arm(b, () => { state[k].splice(i, 1); renderList(k); schedule(); });
  else if (act === 'logo-del') { state.experience[i].logo = null; renderList('experience'); schedule(); }
  else if (act === 'nat') {
    state.showNationality = !state.showNationality;
    b.setAttribute('aria-pressed', String(state.showNationality));
    b.textContent = state.showNationality ? 'إخفاء الجنسية' : 'إظهار الجنسية';
    schedule();
  }
  else if (act === 'an-hide') { e.preventDefault(); openAnConfirm(b); }
  else if (act === 'an-hide-no') closeAnConfirm();
  else if (act === 'an-hide-yes') { closeAnConfirm(); setAnHidden(true); setStatus('تم إخفاء التحليل.'); }
  else if (act === 'an-show') { setAnHidden(false); goStep(2); }
  else if (act === 'goto') goToField(b.dataset.sec, b.dataset.field);
  else if (act === 'start') goStep(1);
  else if (act === 'step') goStep(+b.dataset.step, { focus: b.dataset.focus });
  else if (act === 'signout') signOutNow();
  else if (act === 'auth-resend') resendLink();
  else if (act === 'auth-change') showAuth('form');
  else if (act === 'sample') openSample(b);
  else if (act === 'sample-close') closeSample();
  else if (act === 'export') saveFile(fileBase() + '_CV_data.json', JSON.stringify(state, null, 1), 'تم حفظ النسخة.');
  else if (act === 'blank') arm(b, () => { state = blankState(); renderEditor(); schedule(); setStatus('تم مسح البيانات.'); });
});

/* ---------- example CV (fictional person; shown read-only) ---------- */
let sampleReturn = null;
function openSample(from) {
  const r = CVEngine.layout(SAMPLE_CV, measure);
  $('#sample-pages').innerHTML = r.pages.map(ops => `<div class="page-wrap">${pageSvg(ops, r)}</div>`).join('');
  sampleReturn = from || null;
  $('#sample').hidden = false;
  document.documentElement.style.overflow = 'hidden';
  $('#sample-close').focus();
}
function closeSample() {
  $('#sample').hidden = true;
  document.documentElement.style.overflow = '';
  if (sampleReturn) sampleReturn.focus();
}
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  if (!$('#an-confirm').hidden) closeAnConfirm();
  else if (!$('#sample').hidden) closeSample();
});
let anReturn = null;
function openAnConfirm(from) { anReturn = from; $('#an-confirm').hidden = false; document.documentElement.style.overflow = 'hidden'; $('#an-confirm-yes').focus(); }
function closeAnConfirm() { $('#an-confirm').hidden = true; document.documentElement.style.overflow = ''; if (anReturn && anReturn.isConnected) anReturn.focus(); }
function setAnHidden(v) {
  anHidden = v;
  try { localStorage.setItem('ats-cv-builder-hide-analysis', v ? '1' : '0'); } catch (e) {}
  const s = $('#an-show'); if (s) s.hidden = !v;
  renderPreview();
}
$('#an-confirm').addEventListener('click', e => { if (e.target.id === 'an-confirm') closeAnConfirm(); });
$('#sample').addEventListener('click', e => { if (e.target.id === 'sample') closeSample(); });
$('#viewer').addEventListener('dblclick', () => setZoom(zi === 0 ? 3 : 0));

/* ---------- import an old CV (PDF / DOCX) ---------- */
const LIB = {
  pdf: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js',
  pdfWorker: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js',
  mammoth: 'https://cdn.jsdelivr.net/npm/mammoth@1.6.0/mammoth.browser.min.js'
};
const BULLET_RE = /^[\s]*[•●▪◦‣∙·○■□➢➤►▶✓✔\-–*]\s*/;

/* progress ring: blue at the start, green when done */
const imp = {
  shown: 0, target: 0, raf: 0,
  set(p, step) {
    this.target = Math.max(this.target, Math.min(100, p));
    if (step) $('#imp-step').textContent = step;
    const reduce = window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (reduce) { this.shown = this.target; this.draw(); return; }
    if (!this.raf) this.raf = requestAnimationFrame(() => this.tick());
  },
  tick() {
    this.raf = 0;
    const d = this.target - this.shown;
    this.shown = Math.abs(d) < 0.4 ? this.target : this.shown + d * 0.18;
    this.draw();
    if (this.shown !== this.target) this.raf = requestAnimationFrame(() => this.tick());
  },
  draw() {
    const p = this.shown / 100, C = 2 * Math.PI * 24;
    const arc = $('#imp-arc');
    arc.style.strokeDashoffset = String(C * (1 - p));
    const h = 212 - (212 - 148) * p, s = 68 - 8 * p, l = 40 - 9 * p;
    arc.style.stroke = this.shown >= 100 ? 'var(--ok)' : `hsl(${h.toFixed(0)} ${s.toFixed(0)}% ${l.toFixed(0)}%)`;
    $('#imp-pct').textContent = Math.round(this.shown) + '%';
    $('#imp-pct').style.color = arc.style.stroke;
    if (this.shown >= 100) $('#imp-step').textContent = 'تم';
  },
  reset() { this.shown = 0; this.target = 0; this.draw(); }
};

const loaded = {};
function needScript(src) {
  if (!loaded[src]) loaded[src] = loadScript(src).catch(e => { delete loaded[src]; throw e; });
  return loaded[src];
}

/* ---- text extraction: returns {lines:[{t, bullet, x, xEnd}], links:[], source} ---- */
async function readPdf(buf) {
  await needScript(LIB.pdf);
  await needScript(LIB.pdfWorker);
  const lib = window.pdfjsLib;
  const doc = await lib.getDocument({ data: buf, isEvalSupported: false }).promise;
  const lines = [], links = [];
  for (let n = 1; n <= doc.numPages; n++) {
    imp.set(20 + 35 * (n - 1) / doc.numPages, `نسحب النص من الصفحة ${n} من ${doc.numPages}…`);
    const page = await doc.getPage(n);
    const tc = await page.getTextContent();
    try { (await page.getAnnotations()).forEach(a => { const u = a.url || a.unsafeUrl; if (u) links.push(u); }); } catch (e) {}
    const items = tc.items.filter(it => it.str && it.str.trim()).map(it => ({
      s: it.str, x: it.transform[4], y: it.transform[5], w: it.width, h: Math.abs(it.transform[3]) || it.height || 10
    }));
    items.sort((a, b) => b.y - a.y || a.x - b.x);
    const rows = [];
    items.forEach(it => {
      const r = rows.find(r => Math.abs(r.y - it.y) < Math.max(2, r.h * 0.45));
      if (r) r.items.push(it); else rows.push({ y: it.y, h: it.h, items: [it] });
    });
    rows.sort((a, b) => b.y - a.y);
    rows.forEach(r => {
      r.items.sort((a, b) => a.x - b.x);
      let t = '', prev = null;
      r.items.forEach(it => {
        if (prev) {
          const gap = it.x - (prev.x + prev.w);
          if (gap > 28) t += '\t';
          else if (gap > prev.h * 0.12 && !/\s$/.test(t) && !/^\s/.test(it.s)) t += ' ';
        }
        t += it.s; prev = it;
      });
      t = t.replace(/[  ]+/g, ' ').trim();
      if (!t) return;
      const first = r.items[0], last = r.items[r.items.length - 1];
      lines.push({ t, bullet: BULLET_RE.test(t), x: first.x, xEnd: last.x + last.w, size: Math.max(...r.items.map(i => i.h)), page: n });
    });
  }
  return { lines, links, source: 'pdf' };
}
async function readDocx(buf) {
  await needScript(LIB.mammoth);
  imp.set(35, 'نسحب النص من ملف Word…');
  const res = await window.mammoth.convertToHtml({ arrayBuffer: buf });
  const dom = new DOMParser().parseFromString('<div>' + res.value + '</div>', 'text/html');
  const lines = [], links = [];
  dom.querySelectorAll('a[href]').forEach(a => links.push(a.getAttribute('href')));
  dom.querySelectorAll('p, li, h1, h2, h3, h4, h5, h6').forEach(el => {
    if (el.tagName === 'P' && el.closest('li')) return;
    const t = (el.textContent || '').replace(/[  ]+/g, ' ').trim();
    if (!t) return;
    lines.push({ t, bullet: el.tagName === 'LI' || BULLET_RE.test(t), heading: /^H\d$/.test(el.tagName), x: 0, xEnd: 0 });
  });
  imp.set(55);
  return { lines, links, source: 'docx' };
}

/* ---- rule-based reader (works offline, no account needed) ---- */
const HEADINGS = {
  summary: ['summary', 'professional summary', 'profile', 'profile summary', 'professional profile', 'about me', 'about', 'objective', 'career objective', 'personal statement', 'career summary', 'overview'],
  education: ['education', 'academic background', 'academic qualifications', 'qualifications', 'education and qualifications', 'educational background', 'education and training'],
  skills: ['skills', 'core technical skills', 'technical skills', 'core skills', 'key skills', 'core competencies', 'competencies', 'skills and abilities', 'skills and competencies', 'technical skills and tools', 'soft skills', 'hard skills', 'tools and technologies', 'areas of expertise', 'expertise'],
  projects: ['projects', 'engineering projects', 'academic projects', 'key projects', 'personal projects', 'graduation project', 'project', 'selected projects'],
  experience: ['experience', 'work experience', 'professional experience', 'employment history', 'employment', 'work history', 'career history', 'internships', 'internship', 'training', 'experience and training', 'relevant experience', 'professional background'],
  certs: ['certifications', 'certification', 'certificates', 'courses', 'licenses and certifications', 'certifications and courses', 'courses and certifications', 'training and certifications', 'professional certifications', 'certifications and training'],
  languages: ['languages', 'language', 'language skills']
};
const AR_HEADINGS = { summary: ['نبذة', 'الملخص', 'نبذة عني', 'الهدف الوظيفي'], education: ['التعليم', 'المؤهلات', 'المؤهل العلمي'], skills: ['المهارات'], projects: ['المشاريع'], experience: ['الخبرات', 'الخبرة', 'الخبرات العملية', 'الخبرة العملية'], certs: ['الشهادات', 'الدورات', 'الشهادات والدورات'], languages: ['اللغات'] };
const OTHER_HEADS = ['volunteering', 'volunteer experience', 'volunteer work', 'awards', 'honors and awards', 'achievements', 'interests', 'hobbies', 'references', 'activities', 'extracurricular activities', 'publications', 'memberships', 'personal information', 'personal details', 'additional information'];
const normHead = s => s.toLowerCase().replace(/[:：]\s*$/, '').replace(/&/g, 'and').replace(/[^a-z؀-ۿ ]+/g, ' ').replace(/\s+/g, ' ').trim();
function headingKey(line, bodySize) {
  const raw = line.t.replace(BULLET_RE, '').trim();
  if (raw.length > 48 || /\d{3,}|@/.test(raw)) return null;
  const n = normHead(raw);
  for (const k in HEADINGS) if (HEADINGS[k].includes(n)) return k;
  for (const k in AR_HEADINGS) if (AR_HEADINGS[k].includes(n)) return k;
  if (OTHER_HEADS.includes(n)) return 'other';
  const words = n.split(' ').filter(Boolean).length;
  if (words && words <= 4 && bodySize && line.size >= bodySize * 1.18 && !/[.,;]$/.test(raw) && /^[A-Z\u0600-\u06FF]/.test(raw)) return 'other';
  // an unknown heading: short, all caps, ends with a colon (or a Word heading style)
  if (words && words <= 4 && ((raw === raw.toUpperCase() && /[A-Z]/.test(raw) && /:\s*$/.test(raw)) || line.heading)) return 'other';
  return null;
}

const MON = '(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)';
const DP = `(?:${MON}\\.?\\s*,?\\s*(?:19|20)\\d{2}|\\d{1,2}\\s*[/.-]\\s*(?:19|20)\\d{2}|(?:19|20)\\d{2})`;
const RANGE_RE = new RegExp(`(${DP}\\s*(?:–|—|-|to|until)\\s*(?:${DP}|present|current|now|till now|to date|ongoing)|${DP}\\s*(?:–|—|-)\\s*$)\\s*\\.?\\s*$`, 'i');
const ONE_RE = new RegExp(`(${MON}\\.?\\s*(?:19|20)\\d{2}|(?:19|20)\\d{2})\\s*\\.?\\s*$`, 'i');
function splitDate(text) {
  const t = text.replace(/\s+$/, '');
  let m = t.match(RANGE_RE);
  if (!m) {
    const tab = t.lastIndexOf('\t');
    const tail = tab >= 0 ? t.slice(tab + 1) : '';
    if (tail && ONE_RE.test(tail) && tail.trim().length <= 16) m = [tail, tail.trim()];
    else if (t.length < 90 && /[|]/.test(t) && ONE_RE.test(t)) m = t.match(ONE_RE);
  }
  if (!m) return null;
  const date = m[1].trim().replace(/\.$/, '').replace(/\s*[-–—]\s*/, ' – ').replace(/\s+to\s+/i, ' – ');
  const title = t.slice(0, t.length - m[0].length).replace(/[\t|,–—\-\s]+$/, '').trim();
  return { date, title };
}
const isDateOnly = t => { const d = splitDate(t); return d && !d.title; };
function cleanSegs(title) {
  return title.split(/\s*[|\t]\s*/).map(s => s.trim().replace(/,$/, '').replace(/(?<!\b(?:co|inc|ltd|corp|est|jr|sr))\.$/i, '')).filter(s => s && !/^see\s+(more|project|projects|details|here)$/i.test(s));
}
const SCHOOL_RE = /universit|college|institute|academy|school|جامعة|كلية|معهد/i;

/* group a section's lines into list items, joining lines that wrap */
function toItems(lines, source, doc) {
  const items = [];
  const maxEnd = Math.max(0, ...lines.map(l => l.xEnd || 0));
  const minX = Math.min(...lines.map(l => l.x || 0));
  const anyBullet = lines.some(l => l.bullet);
  const indented = source === 'pdf' && !anyBullet && lines.some(l => l.x > minX + 8);
  lines.forEach((l, i) => {
    const t = l.t.replace(BULLET_RE, '').replace(/\t+/g, ' ').trim();
    if (!t) return;
    if (!items.length || source === 'docx') { items.push(t); return; }
    const prev = lines[i - 1];
    let fresh;
    if (anyBullet) fresh = l.bullet;
    else {
      const right = (doc && doc.right) || maxEnd, width = (doc && doc.width) || 400;
      const prevShort = prev && prev.xEnd && right && prev.xEnd < right - Math.max(40, width * 0.2);
      fresh = /[.:;!?)]$/.test(items[items.length - 1]) || prevShort || (indented && l.x <= minX + 8) || /^[A-Z][^:]{1,40}:/.test(t);
      if (/^[a-z(,&]/.test(t)) fresh = false;
    }
    if (fresh) items.push(t); else items[items.length - 1] += ' ' + t;
  });
  return items;
}

const ROLE_WORDS = /\b(engineer|manager|representative|rep|specialist|analyst|intern|internship|trainee|officer|coordinator|consultant|developer|designer|accountant|assistant|director|supervisor|technician|lead|head|executive|nurse|pharmacist|doctor|physician|teacher|lecturer|instructor|sales|agent|administrator|admin|associate|advisor|architect|scientist|researcher|programmer|operator|planner|inspector|auditor|clerk|cashier|receptionist|secretary|chef|driver|founder|ceo|cto|cfo|owner|president|partner|volunteer|employee|staff|member|controller|buyer|marketer|writer|editor|translator|therapist|dentist|surgeon|resident)\b|مهندس|مدير|مندوب|أخصائي|اخصائي|محلل|متدرب|موظف|مشرف|فني|محاسب|مساعد|مستشار|معلم|ممرض|صيدلي/i;
const COMPANY_WORDS = /\b(company|co\.?|corp|corporation|inc|ltd|llc|est|establishment|group|holding|holdings|bank|hospital|clinic|university|college|institute|ministry|authority|agency|pharma|pharmaceuticals?|contracting|industries|industrial|factory|telecom|communications|solutions|services|systems|technologies|tech|consulting|consultants|trading|international|national|foundation|center|centre|organization|association|council|municipality|airlines|aramco|sabic|stc|neom)\b|شركة|مؤسسة|مستشفى|وزارة|جامعة|هيئة|بنك|مجموعة/i;
function looksCompany(s) {
  s = String(s || '');
  let sc = 0;
  if (COMPANY_WORDS.test(s)) sc += 2;
  if (ROLE_WORDS.test(s)) sc -= 2;
  if (/^[A-Z]{2,6}$/.test(s.trim())) sc += 1;
  return sc;
}
function parseEntries(lines, kind, source, doc) {
  const out = [];
  let cur = null, body = [];
  const close = () => {
    if (!cur) return;
    let b = body.slice();
    if (kind === 'education' && !cur.school && b.length && SCHOOL_RE.test(b[0].t) && !b[0].bullet) cur.school = cleanSegs(b.shift().t)[0] || '';
    const hasBullets = b.some(l => l.bullet) || (source === 'pdf' && b.some(l => l.x > Math.min(...lines.map(x => x.x)) + 8));
    if (!hasBullets && b.length && kind !== 'education') {
      cur.format = source === 'docx' && b.length > 1 ? 'bullets' : 'paragraph';
      cur.details = cur.format === 'paragraph' ? [b.map(l => l.t.replace(/\t+/g, ' ')).join(' ')] : toItems(b, source, doc);
    } else { cur.format = 'bullets'; cur.details = toItems(b, source, doc); }
    out.push(cur); cur = null; body = [];
  };
  const start = (title, date) => {
    close();
    const segs = cleanSegs(title);
    cur = { dates: date || '' };
    if (kind === 'projects') cur.title = segs.join(' | ');
    else {
      let a = segs[0] || '', b = segs.slice(1).join(' | ');
      if (!b) {
        const m = a.match(/^(.+?)\s+(?:at|@|–|—|-)\s+(.+)$/) || (kind === 'education' ? a.match(/^(.+?),\s*(.+)$/) : null);
        if (m) { a = m[1].trim(); b = m[2].trim(); }
      }
      if (kind === 'education') { cur.degree = a; cur.school = b; }
      else if (b && looksCompany(a) > looksCompany(b)) { cur.role = b; cur.company = a; cur._companyFirst = true; }
      else { cur.role = a; cur.company = b; if (b) cur._roleFirst = true; }
    }
  };
  lines.forEach((l, i) => {
    if (!l.bullet) {
      const d = splitDate(l.t);
      if (d && d.title) { start(d.title, d.date); return; }
      if (d && !d.title) {
        if (cur && !cur.dates && !body.length) { cur.dates = d.date; return; }
        const prevLine = body.length ? body[body.length - 1] : null;
        if (prevLine && !prevLine.bullet) { body.pop(); start(prevLine.t, d.date); return; }
        start('', d.date); return;
      }
    }
    if (!cur) { start(l.t, ''); return; }
    body.push(l);
  });
  close();
  return out;
}

const PHONE_RE = /(?:\+|00)?\s*(?:966)?[\s-]*\(?0?\)?\s*5\d(?:[\s-]?\d){7}/;
function contactLike(t) { return /@|https?:|www\.|linkedin|phone|mobile|email|e-mail|address|nationality|tel\b|\+\d|\d{7,}|location/i.test(t); }

function heuristicParse(doc) {
  const { lines, links, source } = doc;
  const all = lines.map(l => l.t).join('\n');
  const o = { titles: {}, education: [], projects: [], experience: [], skills: [], certs: [], languages: [], other_sections: [] };
  // split into sections
  const sizes = lines.map(l => l.size || 0).filter(Boolean).sort((a, b) => a - b);
  const bodySize = sizes.length ? sizes[Math.floor(sizes.length / 2)] : 0;
  if (source === 'pdf' && lines.length) {
    const xs = lines.map(l => l.x), xe = lines.map(l => l.xEnd);
    doc.left = Math.min(...xs); doc.right = Math.max(...xe); doc.width = doc.right - doc.left;
  }
  const secs = [{ key: 'head', lines: [] }];
  let seenKnown = false;
  lines.forEach((l, i) => {
    let k = i === 0 ? null : headingKey(l, bodySize);
    if (k === 'other' && !seenKnown) k = null;
    if (k && k !== 'other') seenKnown = true;
    if (k) {
      secs.push({ key: k, title: l.t.replace(/[:：]\s*$/, '').trim(), lines: [] });
      if (k === 'other') o.other_sections.push(l.t.replace(/[:：]\s*$/, '').trim());
    } else secs[secs.length - 1].lines.push(l);
  });
  // contacts
  const em = all.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/); o.email = em ? em[0] : '';
  const ph = all.match(PHONE_RE); o.phone = ph ? ph[0] : '';
  const li = (links.find(u => /linkedin\.com\/in\//i.test(u)) || (all.match(/(?:https?:\/\/)?(?:[a-z]{2,3}\.)?linkedin\.com\/in\/[A-Za-z0-9\-_%]+\/?/i) || [''])[0]);
  let liUrl = li;
  if (!liUrl) {
    const ln = lines.find(l => /linkedin\.com\/in\//i.test(l.t));
    if (ln) { const m = ln.t.match(/((?:https?:\/\/)?(?:www\.)?linkedin\.com\/in\/[^|\t]*)/i); if (m) liUrl = m[1].replace(/\s+/g, ''); }
  }
  o.linkedin = String(liUrl || '').replace(/^https?:\/\//i, '').replace(/\/$/, '');
  o.website = '';
  // header block
  const head = secs[0].lines.map(l => l.t);
  const plain = head.filter(t => !contactLike(t));
  o.name = (plain[0] || '').replace(/\t+/g, ' ').trim();
  const second = plain[1] || '';
  o.headline = second && second.split(/\s+/).length <= 16 ? second.replace(/\t+/g, ' ').trim() : '';
  const headText = head.join(' | ');
  const siteM = headText.match(/(?:https?:\/\/)?(?:www\.)?(?:[a-z0-9-]+\.)+[a-z]{2,}(?:\/[\w\-./%]*)?/gi) || [];
  o.website = (siteM.find(u => !/linkedin\.com|@/i.test(u) && !/^[^/]*\.(?:com|net|org)$/i.test(u.replace(/^https?:\/\/(www\.)?/i, '')) || /github|gitlab|behance|dribbble/i.test(u)) || '').replace(/^https?:\/\//i, '').replace(/\/$/, '');
  if (o.email && o.website && o.email.toLowerCase().includes(o.website.toLowerCase())) o.website = '';
  o.city = findCity(headText) || '';
  if (!o.city) { const m = headText.match(/(?:address|location)\s*:\s*([^|\t]+)/i); if (m) o.city = m[1].trim(); }
  o.nationality = /nationality\s*:/i.test(headText) ? 'yes' : '';
  const longHead = plain.slice(2).filter(t => t.split(/\s+/).length > 12);
  o.summary = longHead.join(' ');
  // sections
  secs.slice(1).forEach(s => {
    if (s.key === 'other' || !s.lines.length) return;
    o.titles[s.key] = s.title.toUpperCase();
    if (s.key === 'summary') o.summary = (o.summary ? o.summary + ' ' : '') + s.lines.map(l => l.t.replace(/\t+/g, ' ')).join(' ');
    else if (s.key === 'skills' || s.key === 'certs' || s.key === 'languages') o[s.key] = o[s.key].concat(toItems(s.lines, source, doc));
    else o[s.key] = o[s.key].concat(parseEntries(s.lines, s.key, source, doc));
  });
  const cf = o.experience.filter(e => e._companyFirst).length, rf = o.experience.filter(e => e._roleFirst).length;
  o.experience_order = cf > rf ? 'company_first' : 'role_first';
  return o;
}

/* ---- city matching ---- */
const CITY_ALIASES = { mecca: 'Makkah', makka: 'Makkah', medina: 'Madinah', madina: 'Madinah', jiddah: 'Jeddah', jedda: 'Jeddah', jidda: 'Jeddah', hasa: 'Al-Ahsa', ahsa: 'Al-Ahsa', hassa: 'Al-Ahsa', buraidah: 'Buraydah', buraida: 'Buraydah', khamismushayt: 'Khamis Mushait', yenbo: 'Yanbu', tabouk: 'Tabuk', hael: 'Hail', jizan: 'Jazan', gizan: 'Jazan', unayzah: 'Unaizah', onaizah: 'Unaizah', qurayyat: 'Al-Qurayyat', kharj: 'Al-Kharj', sakakah: 'Sakaka', diriyah: 'Ad-Diriyah', alula: 'Al-Ula', ula: 'Al-Ula' };
const cityKey = s => String(s || '').toLowerCase().replace(/^(al|el|ad|ar|as|az|an|ash)[\s-]+/, '').replace(/[\s\-'’]/g, '');
const CITY_INDEX = (() => {
  const m = new Map();
  KSA_CITIES.forEach(([, list]) => list.forEach(([en, ar]) => { m.set(cityKey(en), en); m.set(ar, en); }));
  Object.entries(CITY_ALIASES).forEach(([k, v]) => m.set(k, v));
  return m;
})();
function matchCity(s) {
  const raw = String(s || '').split(/[,،|]/)[0].trim();
  return CITY_INDEX.get(cityKey(raw)) || CITY_INDEX.get(raw) || '';
}
function findCity(text) {
  const parts = String(text).split(/[|,،\t]/).map(x => x.replace(/^.*?:\s*/, '').trim()).filter(Boolean);
  for (const p of parts) { const c = matchCity(p); if (c) return c; }
  return '';
}

/* ---- put the parsed CV into the form ---- */
const str = v => (v == null ? '' : String(v)).trim();
const arr = v => (Array.isArray(v) ? v : (str(v) ? str(v).split(/\n+/) : [])).map(str).filter(Boolean);
function toState(o) {
  const s = blankState();
  s.name = str(o.name).replace(/\t+/g, ' ');
  if (s.name && s.name === s.name.toUpperCase()) s.name = s.name.toLowerCase().replace(/(^|[\s-])([a-z])/g, (m, a, b) => a + b.toUpperCase()); s.headline = str(o.headline); s.email = str(o.email);
  s.phone = CVEngine.localPhone(o.phone).slice(0, 12);
  s.linkedin = str(o.linkedin).replace(/^https?:\/\//i, '').replace(/\/$/, '');
  s.website = str(o.website).replace(/^https?:\/\//i, '').replace(/\/$/, '');
  const c = matchCity(o.city);
  if (c) s.city = c; else if (str(o.city)) { s.city = 'Other'; s.cityOther = str(o.city).replace(/,?\s*(saudi arabia|ksa|السعودية|المملكة العربية السعودية)\s*$/i, '').trim(); }
  s.showNationality = !!str(o.nationality);
  const t = o.titles || {};
  Object.keys(CVEngine.DEFAULT_TITLES).forEach(k => { if (str(t[k])) s.titles[k] = str(t[k]).replace(/[:：]\s*$/, '').toUpperCase(); });
  s.summary = str(o.summary);
  const linkText = (k) => k === 'experience' ? 'See project' : 'See more';
  s.education = (o.education || []).map(e => ({ degree: str(e.degree), school: str(e.school), dates: str(e.dates), link: str(e.link), linkText: str(e.link) ? 'See more' : '', details: arr(e.details).join('\n') })).filter(e => e.degree || e.school);
  s.projects = (o.projects || []).map(p => ({ title: str(p.title), dates: str(p.dates), link: str(p.link), linkText: str(p.link) ? linkText('projects') : '', format: p.format === 'paragraph' ? 'paragraph' : 'bullets', details: arr(p.details).join('\n') })).filter(p => p.title || p.details);
  s.expOrder = o.experience_order === 'company_first' ? 'company' : 'role';
  s.experience = (o.experience || []).map(e => ({ role: str(e.role), company: str(e.company), dates: str(e.dates), link: str(e.link), linkText: str(e.link) ? linkText('experience') : '', format: e.format === 'paragraph' ? 'paragraph' : 'bullets', details: arr(e.details).join('\n'), logo: null })).filter(e => e.role || e.company || e.details);
  ['education', 'projects', 'experience'].forEach(k => { if (!s[k].length) s[k] = [{ education: blankEdu, projects: blankProj, experience: blankExp }[k]()]; });
  s.skills = arr(o.skills).join('\n'); s.certs = arr(o.certs).join('\n'); s.languages = arr(o.languages).join('\n');
  return s;
}

/* ---- flow ---- */
let pendingImport = null;
function showImp(part) {
  ['#imp-idle', '#imp-progress', '#imp-confirm', '#imp-result'].forEach(id => { $(id).hidden = id !== part; });
}
function filledCount(s) {
  let n = 0;
  ['name', 'headline', 'email', 'phone', 'linkedin', 'website', 'summary', 'skills', 'certs', 'languages'].forEach(k => { if (str(s[k])) n++; });
  if (s.city) n++;
  ['education', 'projects', 'experience'].forEach(k => (s[k] || []).forEach(it => { Object.keys(it).forEach(f => { if (f !== 'logo' && f !== 'format' && str(it[f])) n++; }); }));
  return n;
}
function applyImport(next, meta) {
  state = next;
  renderEditor();
  document.querySelectorAll('#editor details.sec').forEach(d => { d.open = true; });
  renderPreview(); save();
  imp.set(100, 'نعبّي الخانات…');
  const miss = (meta.other || []).filter(Boolean);
  $('#imp-result-text').textContent = 'تم سحب بياناتك وتعبئة الخانات. راجعها وعدّل اللي تبي.' +
    (meta.arabic ? ' بعض النصوص عربية، اكتبها بالإنجليزي عشان تظهر في الملف.' : '') +
    (miss.length ? ' هذي الأقسام ما لها خانة وما انضافت: ' + miss.join('، ') + '.' : '');
  const nBad = analyze(CVEngine.layout(cvData(), measure).pages.length).filter(r => !r.ok).length;
  if (nBad && !anHidden) $('#imp-result-text').textContent += ' لقينا ' + notesWord(nBad) + ' تحتاج تعديل.';
  setTimeout(() => { showImp('#imp-result'); }, 650);
  setTimeout(() => { goStep(anHidden ? 3 : 2); }, 1500);
}
async function importCv(file) {
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  if (!/^(pdf|docx)$/.test(ext)) { failImport('الملف لازم يكون PDF أو Word بصيغة .docx. لو ملفك .doc احفظه من Word بصيغة .docx.'); return; }
  if (file.size > 10 * 1024 * 1024) { failImport('حجم الملف أكبر من 10 ميجا. جرّب نسخة أصغر.'); return; }
  showImp('#imp-progress'); imp.reset(); imp.set(5, 'نقرأ الملف…');
  try {
    const buf = await file.arrayBuffer();
    imp.set(14, 'نجهّز القارئ…');
    const doc = ext === 'pdf' ? await readPdf(buf) : await readDocx(buf);
    const chars = doc.lines.reduce((a, l) => a + l.t.length, 0);
    if (chars < 40) { failImport(ext === 'pdf' ? 'ما لقينا نص في الملف. غالبًا الـPDF عبارة عن صورة. جرّب ترفع نسخة Word أو PDF مصدّرة من Word.' : 'الملف فاضي أو ما فيه نص نقدر نقرأه.'); return; }
    imp.set(56, 'نرتّب بياناتك في الخانات…');
    const parsed = heuristicParse(doc);
    imp.set(92);
    const next = toState(parsed);
    if (filledCount(next) < 2) { failImport('ما قدرنا نتعرّف على بيانات السيرة في هذا الملف. جرّب ملف ثاني أو عبّي الخانات يدويًا.'); return; }
    linkedinHint = /linked\s*-?in/i.test(doc.lines.map(l => l.t).join(' ') + ' ' + doc.links.join(' ')) && !next.linkedin;
    const meta = { other: parsed.other_sections || [], arabic: /[؀-ۿ]/.test(JSON.stringify(next)) };
    imp.set(97, 'جاهز');
    if (filledCount(state) > 0) { pendingImport = { next, meta }; showImp('#imp-confirm'); return; }
    applyImport(next, meta);
  } catch (e) {
    failImport('ما قدرنا نقرأ الملف. تأكد من الإنترنت، أو جرّب ملف ثاني.');
  }
}
function failImport(msg) {
  $('#imp-error').textContent = msg; $('#imp-error').hidden = false;
  showImp('#imp-idle'); setStatus(msg);
}
$('#import-cv').addEventListener('change', e => {
  const f = e.target.files && e.target.files[0]; e.target.value = '';
  $('#imp-error').hidden = true;
  if (f) importCv(f);
});
const dz = $('#import');
['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('drag'); }));
['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('drag'); }));
dz.addEventListener('drop', e => { const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; $('#imp-error').hidden = true; if (f) importCv(f); });
dz.addEventListener('click', e => {
  const b = e.target.closest('[data-imp]');
  if (!b) return;
  const a = b.dataset.imp;
  if (a === 'replace' && pendingImport) { const p = pendingImport; pendingImport = null; showImp('#imp-progress'); applyImport(p.next, p.meta); }
  if (a === 'cancel') { pendingImport = null; showImp('#imp-idle'); setStatus('ما تغيّر شي في بياناتك.'); }
  if (a === 'again') { showImp('#imp-idle'); }
});

$('#an-show').hidden = !anHidden;
function renderHeroSheet() {
  const r = CVEngine.layout(SAMPLE_CV, measure);
  $('#hero-sheet').innerHTML = pageSvg(r.pages[0], r);
}
renderHeroSheet();
/* ---------- account: email-link sign in + cloud save (see assets/cloud.js) ---------- */
let cloud = null, user = null, cloudStarted = false, mode = 'loading';
let cloudT = null, cloudBusy = false, cloudAgain = false, resendAt = 0, finishing = false;

function setMode(m, note) {
  mode = m;
  $('#auth').hidden = !(m === 'auth' || m === 'loading');
  $('#app-main').hidden = !(m === 'app' || m === 'local');
  $('#account').hidden = !(m === 'app' || m === 'local');
  if (m === 'loading') showAuth('busy', 'نجهّز الصفحة…');
  if (m === 'app') {
    $('#acc-email').textContent = user ? user.email : '';
    $('#acc-signout').hidden = false;
    $('#acc-note').hidden = true;
  }
  if (m === 'local') {
    $('#acc-email').textContent = '';
    $('#acc-signout').hidden = true;
    $('#acc-note').hidden = false;
    $('#acc-note').textContent = note || 'وضع بدون حساب: بياناتك تنحفظ على هذا الجهاز فقط.';
    setSync('');
  }
  if (m === 'app' || m === 'local') goStep(CVEngine.isEmpty(cvData()) ? 1 : 3, { noScroll: true });
}
function showAuth(part, busyText) {
  ['form', 'sent', 'confirm', 'busy'].forEach(p => { $('#auth-' + p).hidden = p !== part; });
  if (part === 'busy') $('#auth-busy-text').textContent = busyText || 'لحظة…';
  if (part === 'form') { authMsg(''); setTimeout(() => { const f = $('#auth-email'); if (f && !f.value) f.focus({ preventScroll: true }); }, 50); }
}
function authMsg(text, bad, where) {
  const el = $(where || '#auth-msg');
  el.textContent = text || '';
  el.classList.toggle('bad', !!bad);
}
function setSync(s) {
  const el = $('#acc-sync');
  const map = { saving: 'يحفظ…', saved: 'محفوظ في حسابك', error: 'ما انحفظ. تأكد من الإنترنت', loading: 'نجيب بياناتك…' };
  el.textContent = map[s] || '';
  el.dataset.state = s || '';
}
const AUTH_ERRORS = {
  'auth/invalid-email': 'الإيميل مكتوب غلط. تأكد إنه مثل example@gmail.com.',
  'auth/missing-email': 'اكتب إيميلك.',
  'auth/quota-exceeded': 'وصلنا الحد اليومي لرسائل الدخول. جرّب بعد شوي أو بكرة.',
  'auth/too-many-requests': 'محاولات كثيرة. انتظر دقيقة وجرّب مرة ثانية.',
  'auth/invalid-action-code': 'رابط الدخول منتهي أو مستخدم من قبل. اطلب رابط جديد.',
  'auth/expired-action-code': 'رابط الدخول منتهي. اطلب رابط جديد.',
  'auth/unauthorized-continue-uri': 'دومين الموقع مو مضاف في إعدادات Firebase (Authorized domains).',
  'auth/unauthorized-domain': 'دومين الموقع مو مضاف في إعدادات Firebase (Authorized domains).',
  'auth/operation-not-allowed': 'الدخول برابط الإيميل مو مفعّل في Firebase.',
  'auth/network-request-failed': 'ما فيه اتصال بالإنترنت. تأكد من الاتصال وجرّب.'
};
const authError = e => AUTH_ERRORS[e && e.code] || 'صار خطأ. جرّب مرة ثانية.';
const validEmail = v => /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(String(v || '').trim());

function startCloud() {
  const go = c => {
    if (cloudStarted) return;
    cloudStarted = true; cloud = c;
    if (!c || !c.configured) {
      setMode('local', c && c.error ? 'ما قدرنا نتصل بالخادم. بياناتك تنحفظ على هذا الجهاز مؤقتًا.' : 'وضع التجربة: إعدادات Firebase ناقصة، فبياناتك تنحفظ على هذا الجهاز فقط.');
      return;
    }
    if (c.isLinkInUrl()) {
      const em = c.storedEmail();
      finishing = true;
      if (em) finishLink(em); else { setMode('auth'); showAuth('confirm'); }
    }
    c.onUser(u => {
      const was = user;
      user = u;
      if (u) { if (!was || was.uid !== u.uid) loadUserData(u); }
      else if (!finishing) { setMode('auth'); showAuth('form'); }
    });
  };
  if (window.CVCloud) go(window.CVCloud);
  else {
    window.addEventListener('cvcloud', e => go(e.detail), { once: true });
    setTimeout(() => go({ configured: false, error: true }), 15000);
  }
}
async function finishLink(email) {
  finishing = true;
  setMode('auth'); showAuth('busy', 'ندخلك لحسابك…');
  try { await cloud.completeLink(email); finishing = false; }
  catch (e) { finishing = false; cloud.clearLink(); showAuth('form'); authMsg(authError(e), true); }
}
async function loadUserData(u) {
  setMode('auth'); showAuth('busy', 'نجيب بياناتك…');
  try {
    const d = await cloud.load(u.uid);
    if (d && d.state) state = normalize(d.state);
    else if (!CVEngine.isEmpty(cvData())) cloudSaveSoon(0);   // first sign in: keep what was typed before
    try { localStorage.setItem(KEY, JSON.stringify(state)); } catch (e) {}
    renderEditor(); renderPreview();
    setMode('app'); setSync('saved');
  } catch (e) {
    renderEditor(); renderPreview();
    setMode('app'); setSync('error');
  }
}
function cloudSaveSoon(delay) {
  if (!cloud || !cloud.configured || !user) return;
  clearTimeout(cloudT);
  setSync('saving');
  cloudT = setTimeout(cloudSaveNow, delay == null ? 1200 : delay);
}
async function cloudSaveNow() {
  if (!user) return;
  if (cloudBusy) { cloudAgain = true; return; }
  cloudBusy = true;
  try { await cloud.save(user.uid, { state: JSON.parse(JSON.stringify(state)), email: user.email || '' }); setSync('saved'); }
  catch (e) { setSync('error'); }
  cloudBusy = false;
  if (cloudAgain) { cloudAgain = false; cloudSaveNow(); }
}
async function signOutNow() {
  clearTimeout(cloudT);
  if (cloudBusy || $('#acc-sync').dataset.state === 'saving') await cloudSaveNow();
  try { await cloud.signOut(); } catch (e) {}
  user = null;
  state = blankState();
  try { localStorage.removeItem(KEY); } catch (e) {}
  renderEditor(); renderPreview();
  showImp('#imp-idle');
  setMode('auth'); showAuth('form');
}
async function sendLink(email) {
  const btn = $('#auth-send');
  if (!validEmail(email)) { authMsg('اكتب إيميل صحيح، مثل example@gmail.com.', true); return; }
  btn.disabled = true; authMsg('نرسل الرابط…');
  try {
    await cloud.sendLink(email.trim());
    resendAt = Date.now() + 30000;
    $('#auth-sent-to').textContent = email.trim();
    showAuth('sent'); authMsg('', false, '#auth-sent-msg');
  } catch (e) { authMsg(authError(e), true); }
  btn.disabled = false;
}
async function resendLink() {
  const left = Math.ceil((resendAt - Date.now()) / 1000);
  if (left > 0) { authMsg(`تقدر تعيد الإرسال بعد ${left} ثانية.`, false, '#auth-sent-msg'); return; }
  const email = $('#auth-sent-to').textContent;
  authMsg('نرسل الرابط مرة ثانية…', false, '#auth-sent-msg');
  try { await cloud.sendLink(email); resendAt = Date.now() + 30000; authMsg('أرسلنا رابط جديد.', false, '#auth-sent-msg'); }
  catch (e) { authMsg(authError(e), true, '#auth-sent-msg'); }
}
$('#auth-form-el').addEventListener('submit', e => { e.preventDefault(); sendLink($('#auth-email').value); });
$('#auth-confirm-form').addEventListener('submit', e => {
  e.preventDefault();
  const v = $('#auth-confirm-email').value;
  if (!validEmail(v)) { authMsg('اكتب إيميل صحيح.', true, '#auth-confirm-msg'); return; }
  finishLink(v.trim());
});

document.body.dataset.view = 'edit';
renderEditor();
setZoom(0);
renderPreview();
setMode('loading');
startCloud();
if (document.fonts && document.fonts.ready) document.fonts.ready.then(renderPreview);
})();
