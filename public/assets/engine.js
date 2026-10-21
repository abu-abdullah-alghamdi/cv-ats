/* CV engine: state -> positioned ops (preview + PDF) and -> Word document */
(function (root) {
  const PW = 595.28, PH = 841.89, ML = 44, MR = 44, MT = 40, MB = 42, CW = PW - ML - MR;
  const SZ = { name: 13.5, headline: 12, contact: 9.5, h2: 11.5, body: 10 };
  const INK = '#141414', LINK = '#1F5FA8', RULE = '#4a4a4a';
  const DEFAULT_TITLES = {
    summary: 'SUMMARY', education: 'EDUCATION', skills: 'SKILLS', projects: 'PROJECTS',
    experience: 'EXPERIENCE', certs: 'CERTIFICATIONS', languages: 'LANGUAGES'
  };
  const ORDER = ['summary', 'education', 'skills', 'projects', 'experience', 'certs', 'extras', 'languages'];

  function clean(s) {
    return String(s == null ? '' : s).replace(/\t/g, ' ').replace(/[​-‏‪-‮]/g, '');
  }
  const tr = s => clean(s).trim();
  function lines(text) {
    return clean(text).split(/\r?\n/).map(l => l.trim().replace(/^[•●▪\-*]\s*/, '')).filter(Boolean);
  }
  // "Label: text" -> bold label. Only short labels count, so a sentence that happens to contain a colon stays plain.
  function labelRuns(line) {
    const m = line.match(/^([^:]{1,80}):\s*(.*)$/);
    if (m && !/https?$/i.test(m[1].trim())) {
      const label = m[1].trim();
      if (!m[2]) return [{ t: label + ':', s: 'bold' }];
      if (label.length <= 45 && label.split(/\s+/).length <= 5 && !/[,;(]/.test(label)) return [{ t: label + ':', s: 'bold' }, { t: ' ' + m[2] }];
    }
    return [{ t: line }];
  }
  // "Course name | 2024" -> text + date shown on the right like other dated lines
  function splitTailDate(line) {
    const m = String(line).match(/^(.*\S)\s*\|\s*((?:[A-Za-z]{3,9}\.?\s*)?(?:\d{1,2}\/)?(?:19|20)\d{2}(?:\s*[–—-]\s*(?:(?:[A-Za-z]{3,9}\.?\s*)?(?:\d{1,2}\/)?(?:19|20)\d{2}|present|current))?)\.?$/i);
    return m && m[2].length <= 24 ? { text: m[1].replace(/[|,\s]+$/, ''), date: m[2].replace(/\s*[–—-]\s*/, ' – ') } : { text: line, date: '' };
  }
  const nameText = S => S.nameCase === 'asis' ? tr(S.name) : tr(S.name).toUpperCase();
  const extraItems = S => (S.extras || []).filter(x => tr(x.title) || lines(x.content).length);
  function urlOf(v) {
    v = tr(v);
    if (!v) return '';
    return /^[a-z]+:/i.test(v) ? v : 'https://' + v.replace(/^\/+/, '');
  }
  const shown = v => tr(v).replace(/^https?:\/\//i, '').replace(/\/$/, '');
  // Saudi mobile: stored as local digits without 0 (e.g. 501234567); shown as +966 50 123 4567
  function localPhone(v) {
    let d = clean(v).replace(/\D/g, '');
    if (d.startsWith('00')) d = d.slice(2);
    if (d.startsWith('966')) d = d.slice(3);
    return d.replace(/^0+/, '');
  }
  function phoneShown(v) {
    const d = localPhone(v);
    if (!d) return '';
    const g = d.length === 9 ? d.slice(0, 2) + ' ' + d.slice(2, 5) + ' ' + d.slice(5) : d;
    return '+966 ' + g;
  }
  const telOf = v => 'tel:+966' + localPhone(v);
  const titleOf = (S, k) => tr((S.titles || {})[k]) || DEFAULT_TITLES[k];

  function titleRuns(main, sub, link, linkText) {
    const r = [];
    if (tr(main)) r.push({ t: tr(main), s: 'bold' });
    if (tr(sub)) { if (r.length) r.push({ t: ' | ', s: 'bold' }); r.push({ t: tr(sub), s: 'bold' }); }
    if (urlOf(link)) { if (r.length) r.push({ t: ' | ', s: 'bold' }); r.push({ t: tr(linkText) || 'See more', link: urlOf(link) }); }
    return r;
  }
  // experience line order follows the CV: "Role | Company" (default) or "Company | Role"
  const expRuns = (S, e) => S.expOrder === 'company' ? titleRuns(e.company, e.role, e.link, e.linkText) : titleRuns(e.role, e.company, e.link, e.linkText);
  function contactLines(S) {
    const join = parts => parts.filter(Boolean).reduce((acc, p, i) => (i ? acc.concat([{ t: ' | ', s: 'bold' }], p) : acc.concat(p)), []);
    const lbl = (l, v, link) => tr(v) ? [{ t: l + ': ', s: 'bold' }, { t: tr(v), link: link || '' }] : null;
    return [
      join([
        lbl('Address', S.address),
        lbl('Email', S.email, tr(S.email) ? 'mailto:' + tr(S.email) : ''),
        lbl('Phone', phoneShown(S.phone), localPhone(S.phone) ? telOf(S.phone) : '')
      ]),
      join([
        lbl('Nationality', S.nationality),
        tr(S.linkedin) ? lbl('LinkedIn', shown(S.linkedin), urlOf(S.linkedin)) : null,
        tr(S.website) ? [{ t: shown(S.website), link: urlOf(S.website) }] : null
      ])
    ].filter(l => l.length);
  }
  const eduItems = S => (S.education || []).filter(e => tr(e.degree) || tr(e.school));
  const projItems = S => (S.projects || []).filter(p => tr(p.title));
  const expItems = S => (S.experience || []).filter(e => tr(e.role) || tr(e.company));
  function isEmpty(S) {
    return !(tr(S.name) || tr(S.headline) || contactLines(S).length || lines(S.summary).length || eduItems(S).length ||
      lines(S.skills).length || projItems(S).length || expItems(S).length || lines(S.certs).length || lines(S.languages).length || extraItems(S).length);
  }

  /* ====================== layout for preview + PDF ====================== */
  function layout(S, measure) {
    const pages = [[]];
    let page = 0, y = MT;
    const ops = () => pages[page];
    const lh = size => size * 1.32;
    function newPage() { pages.push([]); page++; y = MT; }
    function ensure(h) { if (y + h > PH - MB) newPage(); }

    function breakLines(runs, maxW, size) {
      const units = [];
      let cur = null;
      runs.forEach(r => {
        clean(r.t).split(/(\s+)/).forEach(p => {
          if (!p) return;
          if (/^\s+$/.test(p)) { cur = null; units.push({ space: true, run: r }); }
          else {
            if (!cur) { cur = { parts: [] }; units.push(cur); }
            cur.parts.push({ t: p, s: r.s || 'normal', link: r.link || '' });
          }
        });
      });
      const wOf = part => measure(part.t, size, part.s);
      const out = [];
      let line = [], lw = 0, pendingSpace = null;
      const flush = () => { out.push(line); line = []; lw = 0; pendingSpace = null; };
      units.forEach(u => {
        if (u.space) { if (line.length) pendingSpace = { t: ' ', s: u.run.s || 'normal', link: u.run.link || '' }; return; }
        const ww = u.parts.reduce((a, p) => a + wOf(p), 0);
        const sw = pendingSpace ? wOf(pendingSpace) : 0;
        if (line.length && lw + sw + ww > maxW) flush();
        if (!line.length && ww > maxW) {
          u.parts.forEach(p => {
            let chunk = '';
            for (const ch of p.t) {
              const test = chunk + ch;
              if (lw + measure(test, size, p.s) > maxW && (chunk || line.length)) {
                if (chunk) line.push({ t: chunk, s: p.s, link: p.link });
                flush(); chunk = ch;
              } else chunk = test;
            }
            if (chunk) { line.push({ t: chunk, s: p.s, link: p.link }); lw += measure(chunk, size, p.s); }
          });
          return;
        }
        if (line.length && pendingSpace) { line.push(pendingSpace); lw += sw; }
        pendingSpace = null;
        u.parts.forEach(p => line.push(p));
        lw += ww;
      });
      if (line.length) flush();
      return out.map(l => {
        const segs = [];
        l.forEach(p => {
          const last = segs[segs.length - 1];
          if (last && last.s === p.s && last.link === p.link) last.t += p.t;
          else segs.push({ t: p.t, s: p.s, link: p.link });
        });
        segs.forEach(sg => sg.w = measure(sg.t, size, sg.s));
        return segs;
      });
    }

    function rich(runs, o) {
      const size = o.size || SZ.body, L = o.lh || lh(size);
      const ls = breakLines(runs, o.w, size);
      ls.forEach((segs, i) => {
        ensure(L);
        const base = y + size * 0.98;
        if (i === 0 && o.onFirst) o.onFirst(base);
        const total = segs.reduce((a, s) => a + s.w, 0);
        let x = o.x;
        if (o.align === 'center') x = o.x + (o.w - total) / 2;
        let ws = 0;
        if (o.align === 'justify' && i < ls.length - 1 && segs.length === 1 && !segs[0].link) {
          const spaces = (segs[0].t.match(/ /g) || []).length;
          if (spaces > 0) {
            const extra = (o.w - total) / spaces;
            if (extra > 0 && extra < size * 0.9) ws = extra;
          }
        }
        segs.forEach(sg => {
          const w = sg.w + (ws ? (sg.t.match(/ /g) || []).length * ws : 0);
          ops().push({ k: 'text', x, y: base, t: sg.t, size, s: sg.s, c: sg.link ? LINK : INK, ws });
          if (sg.link) {
            ops().push({ k: 'line', x1: x, y1: base + 1.3, x2: x + sg.w, y2: base + 1.3, w: 0.5, c: LINK });
            ops().push({ k: 'link', x, y: base - size * 0.82, w: sg.w, h: size * 1.05, url: sg.link });
          }
          x += w;
        });
        y += L;
      });
    }
    function rule() {
      y += 0.5;
      ops().push({ k: 'line', x1: ML - 2, y1: y, x2: PW - MR + 2, y2: y, w: 0.6, c: RULE });
      y += 4;
    }
    function heading(title) {
      ensure(lh(SZ.h2) + 8 + lh(SZ.body) * 3);
      y += 7;
      rich([{ t: title.replace(/:?$/, ':'), s: 'bold' }], { x: ML, w: CW, size: SZ.h2 });
      rule();
    }
    function bullets(items, size) {
      const dotX = ML + 8, tx = ML + 22;
      items.forEach(runs => {
        rich(runs, { x: tx, w: PW - MR - tx, size, onFirst: base => ops().push({ k: 'dot', x: dotX, y: base - size * 0.32, r: 1.55 }) });
        y += 1.2;
      });
    }
    function entryHeader(runs, date) {
      const size = SZ.body + 0.8;
      ensure(lh(size) * 3);
      const dw = date ? measure(date, size, 'italic') : 0;
      if (!runs.length) runs = [{ t: '' }];
      rich(runs, {
        x: ML, w: CW - (dw ? dw + 14 : 0), size,
        onFirst: base => { if (date) ops().push({ k: 'text', x: PW - MR - dw, y: base, t: date, size, s: 'italic', c: INK }); }
      });
      y += 1.5;
    }
    function datedLine(line) {
      const { text, date } = splitTailDate(line);
      if (!date) { rich(labelRuns(line), { x: ML, w: CW, size: SZ.body }); return; }
      const dw = measure(date, SZ.body, 'italic');
      rich(labelRuns(text), {
        x: ML, w: CW - dw - 14, size: SZ.body,
        onFirst: base => ops().push({ k: 'text', x: PW - MR - dw, y: base, t: date, size: SZ.body, s: 'italic', c: INK })
      });
    }
    function details(text, format) {
      const ls = lines(text);
      if (!ls.length) return;
      if (format === 'paragraph') rich([{ t: ls.join(' ') }], { x: ML, w: CW, size: SZ.body, align: 'justify' });
      else bullets(ls.map(labelRuns), SZ.body);
    }

    if (tr(S.name)) rich([{ t: nameText(S), s: 'bold' }], { x: ML, w: CW, size: SZ.name, align: 'center' });
    y += 1;
    if (tr(S.headline)) rich([{ t: tr(S.headline), s: 'bold' }], { x: ML, w: CW, size: SZ.headline, align: 'center' });
    y += 1;
    contactLines(S).forEach(l => rich(l, { x: ML, w: CW, size: SZ.contact, align: 'center' }));

    ORDER.forEach(key => {
      if (key === 'summary' && lines(S.summary).length) {
        heading(titleOf(S, key));
        rich([{ t: lines(S.summary).join(' ') }], { x: ML, w: CW, size: SZ.body + 0.5, align: 'justify', lh: lh(SZ.body + 0.5) * 1.04 });
      }
      if (key === 'education' && eduItems(S).length) {
        heading(titleOf(S, key));
        eduItems(S).forEach((e, i) => {
          if (i) y += 5;
          entryHeader(titleRuns(e.degree, e.school, e.link, e.linkText), tr(e.dates));
          details(e.details, 'bullets');
        });
      }
      if (key === 'skills' && lines(S.skills).length) {
        heading(titleOf(S, key));
        bullets(lines(S.skills).map(labelRuns), SZ.body);
      }
      if (key === 'projects' && projItems(S).length) {
        heading(titleOf(S, key));
        projItems(S).forEach((p, i) => {
          if (i) y += 5;
          entryHeader(titleRuns(p.title, '', p.link, p.linkText), tr(p.dates));
          details(p.details, p.format);
        });
      }
      if (key === 'experience' && expItems(S).length) {
        heading(titleOf(S, key));
        expItems(S).forEach((e, i) => {
          if (i) y += 7;
          if (e.logo && e.logo.src) {
            const h = 20, w = Math.min(h * 5, h * ((e.logo.w / e.logo.h) || 2));
            ensure(h + 3 + lh(SZ.body + 0.8) * 3);
            ops().push({ k: 'img', x: ML, y, w, h, src: e.logo.src });
            y += h + 3;
          }
          entryHeader(expRuns(S, e), tr(e.dates));
          details(e.details, e.format);
        });
      }
      if (key === 'certs' && lines(S.certs).length) {
        heading(titleOf(S, key));
        lines(S.certs).forEach(datedLine);
      }
      if (key === 'extras') extraItems(S).forEach(x => {
        heading(tr(x.title) || 'ADDITIONAL INFORMATION');
        const ls = lines(x.content);
        if (x.format === 'paragraph') rich([{ t: ls.join(' ') }], { x: ML, w: CW, size: SZ.body, align: 'justify' });
        else if (x.format === 'bullets') bullets(ls.map(labelRuns), SZ.body);
        else ls.forEach(datedLine);
      });
      if (key === 'languages' && lines(S.languages).length) {
        heading(titleOf(S, key));
        lines(S.languages).forEach(l => rich(labelRuns(l), { x: ML, w: CW, size: SZ.body }));
      }
    });
    return { pages, PW, PH };
  }

  function toPDF(result, jsPDF, meta) {
    const imgAlias = {};
    const doc = new jsPDF({ unit: 'pt', format: 'a4', compress: true });
    doc.setProperties({ title: meta.title || 'CV', subject: 'Curriculum Vitae (CV)', author: meta.author || '', keywords: meta.keywords || '', creator: 'CV Builder' });
    result.pages.forEach((ops, pi) => {
      if (pi) doc.addPage();
      ops.forEach(op => {
        if (op.k === 'text') {
          doc.setFont('helvetica', op.s || 'normal');
          doc.setFontSize(op.size);
          doc.setTextColor(op.c);
          if (op.ws) {
            const nat = doc.getTextWidth(op.t), sp = (op.t.match(/ /g) || []).length;
            doc.text([op.t, ''], op.x, op.y, { align: 'justify', maxWidth: nat + sp * op.ws + 0.001 });
          } else doc.text(op.t, op.x, op.y);
        } else if (op.k === 'line') {
          doc.setDrawColor(op.c); doc.setLineWidth(op.w); doc.line(op.x1, op.y1, op.x2, op.y2);
        } else if (op.k === 'dot') {
          doc.setFillColor('#141414'); doc.circle(op.x, op.y, op.r, 'F');
        } else if (op.k === 'link') {
          doc.link(op.x, op.y, op.w, op.h, { url: op.url });
        } else if (op.k === 'img') {
          const alias = imgAlias[op.src] || (imgAlias[op.src] = 'logo' + Object.keys(imgAlias).length);
          doc.addImage(op.src, 'PNG', op.x, op.y, op.w, op.h, alias, 'FAST');
        }
      });
    });
    return doc;
  }

  /* ====================== Word (.docx) ====================== */
  function toDocx(S, D) {
    const FONT = 'Calibri', PAGE_W = 11906, MARG = 850, CONTENT = PAGE_W - MARG * 2;
    const half = pt => Math.round(pt * 2);
    const runsToChildren = (runs, size, extra) => runs.map(r => {
      const base = { text: r.t, font: FONT, size: half(size), bold: r.s === 'bold', italics: r.s === 'italic', color: '141414' };
      Object.assign(base, extra || {});
      if (r.link) {
        return new D.ExternalHyperlink({ link: r.link, children: [new D.TextRun(Object.assign(base, { color: '1F5FA8', underline: {} }))] });
      }
      return new D.TextRun(base);
    });
    const P = (opts) => new D.Paragraph(Object.assign({ spacing: { after: 20, line: 264 } }, opts));
    const out = [];
    if (tr(S.name)) out.push(P({ alignment: D.AlignmentType.CENTER, spacing: { after: 30 }, children: runsToChildren([{ t: nameText(S), s: 'bold' }], 14) }));
    if (tr(S.headline)) out.push(P({ alignment: D.AlignmentType.CENTER, spacing: { after: 40 }, children: runsToChildren([{ t: tr(S.headline), s: 'bold' }], 12.5) }));
    contactLines(S).forEach(l => out.push(P({ alignment: D.AlignmentType.CENTER, spacing: { after: 10 }, children: runsToChildren(l, 10) })));

    const heading = t => out.push(P({
      spacing: { before: 200, after: 80 }, keepNext: true,
      border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: '4A4A4A', space: 2 } },
      children: runsToChildren([{ t: t.replace(/:?$/, ':'), s: 'bold' }], 12)
    }));
    const bullet = runs => out.push(P({ numbering: { reference: 'cv-bullets', level: 0 }, spacing: { after: 30, line: 259 }, children: runsToChildren(runs, 10.5) }));
    const plain = (runs, opts) => out.push(P(Object.assign({ children: runsToChildren(runs, 10.5) }, opts || {})));
    const entry = (runs, date) => out.push(P({
      keepNext: true, spacing: { before: 60, after: 40 },
      tabStops: [{ type: D.TabStopType.RIGHT, position: CONTENT }],
      children: runsToChildren(runs, 11).concat(date ? [new D.TextRun({ children: [new D.Tab(), date], italics: true, font: FONT, size: half(11), color: '141414' })] : [])
    }));
    const logo = lg => {
      if (!lg || !lg.src) return;
      const hPt = 20, wPt = Math.min(hPt * 5, hPt * ((lg.w / lg.h) || 2));
      const bin = atob(lg.src.split(',')[1]), bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      out.push(P({ keepNext: true, spacing: { before: 120, after: 20 }, children: [new D.ImageRun({ data: bytes, transformation: { width: Math.round(wPt * 4 / 3), height: Math.round(hPt * 4 / 3) } })] }));
    };
    const datedPlain = line => {
      const { text, date } = splitTailDate(line);
      if (!date) { plain(labelRuns(line)); return; }
      out.push(P({
        tabStops: [{ type: D.TabStopType.RIGHT, position: CONTENT }],
        children: runsToChildren(labelRuns(text), 10.5).concat([new D.TextRun({ children: [new D.Tab(), date], italics: true, font: FONT, size: half(10.5), color: '141414' })])
      }));
    };
    const details = (text, format) => {
      const ls = lines(text);
      if (!ls.length) return;
      if (format === 'paragraph') plain([{ t: ls.join(' ') }], { alignment: D.AlignmentType.JUSTIFIED });
      else ls.forEach(l => bullet(labelRuns(l)));
    };

    ORDER.forEach(key => {
      if (key === 'summary' && lines(S.summary).length) { heading(titleOf(S, key)); plain([{ t: lines(S.summary).join(' ') }], { alignment: D.AlignmentType.JUSTIFIED }); }
      if (key === 'education' && eduItems(S).length) { heading(titleOf(S, key)); eduItems(S).forEach(e => { entry(titleRuns(e.degree, e.school, e.link, e.linkText), tr(e.dates)); details(e.details, 'bullets'); }); }
      if (key === 'skills' && lines(S.skills).length) { heading(titleOf(S, key)); lines(S.skills).forEach(l => bullet(labelRuns(l))); }
      if (key === 'projects' && projItems(S).length) { heading(titleOf(S, key)); projItems(S).forEach(p => { entry(titleRuns(p.title, '', p.link, p.linkText), tr(p.dates)); details(p.details, p.format); }); }
      if (key === 'experience' && expItems(S).length) { heading(titleOf(S, key)); expItems(S).forEach(e => { logo(e.logo); entry(expRuns(S, e), tr(e.dates)); details(e.details, e.format); }); }
      if (key === 'certs' && lines(S.certs).length) { heading(titleOf(S, key)); lines(S.certs).forEach(datedPlain); }
      if (key === 'extras') extraItems(S).forEach(x => {
        heading(tr(x.title) || 'ADDITIONAL INFORMATION');
        const ls = lines(x.content);
        if (x.format === 'paragraph') plain([{ t: ls.join(' ') }], { alignment: D.AlignmentType.JUSTIFIED });
        else if (x.format === 'bullets') ls.forEach(l => bullet(labelRuns(l)));
        else ls.forEach(datedPlain);
      });
      if (key === 'languages' && lines(S.languages).length) { heading(titleOf(S, key)); lines(S.languages).forEach(l => plain(labelRuns(l))); }
    });

    return new D.Document({
      title: (tr(S.name) || 'CV') + ' – CV', creator: tr(S.name) || 'CV Builder', description: 'Curriculum Vitae',
      styles: { default: { document: { run: { font: FONT, size: 21 } } } },
      numbering: { config: [{ reference: 'cv-bullets', levels: [{ level: 0, format: D.LevelFormat.BULLET, text: '•', alignment: D.AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 240 } } } }] }] },
      sections: [{ properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 780, bottom: 780, left: MARG, right: MARG } } }, children: out }]
    });
  }

  root.CVEngine = { localPhone, layout, toPDF, toDocx, isEmpty, lines, DEFAULT_TITLES, PW, PH };
})(typeof window !== 'undefined' ? window : globalThis);
