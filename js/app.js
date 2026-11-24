/* Dashboard Pembiayaan Daerah — Kota Kediri (LRA 2023–2025). Tanpa framework; Chart.js untuk grafik. */
(function () {
  'use strict';
  const D = window.DATA, YRS = D.years, A = 0, R = 1, P = 2;
  const NAMES = {}; D.rows.forEach((r) => { NAMES[r[0]] = r; });
  const $ = (s, el) => (el || document).querySelector(s), $$ = (s, el) => Array.from((el || document).querySelectorAll(s));
  const nf = (d) => new Intl.NumberFormat('id-ID', { minimumFractionDigits: d, maximumFractionDigits: d });
  const F0 = nf(0), F1 = nf(1), F2 = nf(2);
  const esc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const v = (y, n, j) => (D.v[y][n] ? D.v[y][n][j] : 0);
  const cap = (y, n) => (v(y, n, A) ? v(y, n, R) / v(y, n, A) : 0);
  const chg = (y, n) => (v(y, n, P) ? (v(y, n, R) - v(y, n, P)) / Math.abs(v(y, n, P)) : null);
  const pct = (f, d) => nf(d == null ? 2 : d).format(f * 100) + '%';
  const sgn = (f) => (f >= 0 ? '+' : '−') + nf(2).format(Math.abs(f) * 100) + '%';
  const rp = (x) => {
    const a = Math.abs(x), s = x < 0 ? '−' : '';
    if (a >= 1e12) return s + 'Rp' + F2.format(a / 1e12) + ' triliun';
    if (a >= 1e9) return s + 'Rp' + F2.format(a / 1e9) + ' miliar';
    if (a >= 1e6) return s + 'Rp' + F2.format(a / 1e6) + ' juta';
    return s + 'Rp' + F0.format(a);
  };
  const rpT = (x) => { const a = Math.abs(x), s = x < 0 ? '−' : ''; return a >= 1e12 ? s + 'Rp' + F2.format(a / 1e12) + ' T' : s + 'Rp' + F2.format(a / 1e9) + ' M'; };
  const mil = (x) => (x === 0 ? '–' : x < 0 ? '(' + F2.format(-x / 1e9) + ')' : F2.format(x / 1e9));
  const up = (f) => (f >= 0 ? 'naik' : 'turun');
  const lcase = (s) => s.charAt(0).toLowerCase() + s.slice(1);

  /* ---------- status ---------- */
  const store = { get(k, d) { try { const x = localStorage.getItem(k); return x === null ? d : x; } catch (e) { return d; } }, set(k, x) { try { localStorage.setItem(k, x); } catch (e) { /* abaikan */ } } };
  const state = { year: +store.get('kdr.year', 2025), theme: store.get('kdr.theme', window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'), sheet: 'Rincian LRA', jin: 'all', jout: 'all' };
  if (!YRS.includes(state.year)) state.year = 2025;
  const VIEWS = [['ringkasan', 'Ringkasan'], ['penerimaan', 'Penerimaan'], ['pengeluaran', 'Pengeluaran'], ['semua', 'Semua Tahun'], ['lembar', 'Lembar Kerja'], ['unduh', 'Unduh Excel'], ['tentang', 'Tentang Data']];
  const YEAR_VIEWS = ['ringkasan', 'penerimaan', 'pengeluaran', 'lembar'];
  const fileOf = (y) => 'downloads/Laporan-Pembiayaan-Kota-Kediri-' + y + '.xlsx';
  const ICON = {
    dl: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12m0 0l-4-4m4 4l4-4M4 19h16"/></svg>',
    sun: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2m0 16v2M2 12h2m16 0h2M5 5l1.5 1.5M17.5 17.5L19 19M5 19l1.5-1.5M17.5 6.5L19 5"/></svg>',
    moon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 13a9 9 0 11-10-10 7 7 0 0010 10z"/></svg>'
  };
  const GROUPS = {
    in: { total: 43, label: 'Penerimaan Pembiayaan', items: [44, 45], sel: 'jin', title: 'Jenis Penerimaan' },
    out: { total: 46, label: 'Pengeluaran Pembiayaan', items: [47, 48], sel: 'jout', title: 'Jenis Pengeluaran' }
  };

  /* ---------- tema & warna grafik ---------- */
  const cssv = (n) => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
  function pal() { return { ink: cssv('--ink'), text: cssv('--text'), muted: cssv('--muted'), line: cssv('--line'), soft: cssv('--soft'), card: cssv('--card'), indigo: cssv('--indigo'), indigo2: cssv('--indigo2'), amber: cssv('--amber'), coral: cssv('--coral'), teal: cssv('--teal'), plum: cssv('--plum'), lilac: cssv('--lilac') }; }
  const alpha = (hex, a) => { const m = /^#?([\da-f]{2})([\da-f]{2})([\da-f]{2})$/i.exec(hex); return m ? 'rgba(' + parseInt(m[1], 16) + ',' + parseInt(m[2], 16) + ',' + parseInt(m[3], 16) + ',' + a + ')' : hex; };
  function applyTheme() {
    document.documentElement.setAttribute('data-theme', state.theme);
    $('#themeBtn').innerHTML = state.theme === 'dark' ? ICON.sun : ICON.moon;
    const m = $('meta[name=theme-color]'); if (m) m.content = state.theme === 'dark' ? '#0A0C24' : '#171A4A';
  }

  /* ---------- narasi (dihitung dari data) ---------- */
  function narasi(y) {
    const a43 = v(y, 43, A), r43 = v(y, 43, R), sil = v(y, 44, R), dc = v(y, 45, R), r46 = v(y, 46, R), a46 = v(y, 46, A), p46 = v(y, 46, P);
    const r47 = v(y, 47, R), r48 = v(y, 48, R), net = v(y, 49, R), sur = v(y, 41, R), sl = v(y, 50, R), g43 = chg(y, 43), g50 = chg(y, 50);
    const out = [];
    out.push({ c: 'var(--indigo2)', ic: '⬇', t: 'Penerimaan pembiayaan', p: 'Penerimaan pembiayaan terealisasi <b>' + rp(r43) + '</b> (<b>' + pct(a43 ? r43 / a43 : 0) + '</b> dari anggaran ' + rp(a43) + '). Sumber utamanya penggunaan SiLPA tahun sebelumnya sebesar <b>' + rp(sil) + '</b> (' + pct(r43 ? sil / r43 : 0) + ' dari penerimaan)' + (dc > 0 ? ', ditambah pencairan dana cadangan ' + rp(dc) : ', tanpa pencairan dana cadangan') + '. Dibanding ' + (y - 1) + ', penerimaan ' + up(g43 || 0) + ' <b>' + pct(Math.abs(g43 || 0)) + '</b>.' });
    out.push({ c: 'var(--amber)', ic: '⬆', t: 'Pengeluaran pembiayaan', p: r46 === 0 ? 'Tidak ada pengeluaran pembiayaan pada TA ' + y + ': penyertaan modal dan pembentukan dana cadangan sama-sama <b>Rp0</b>' + (p46 > 0 ? ', sedangkan tahun ' + (y - 1) + ' ada pengeluaran ' + rp(p46) : '') + '.' : 'Pengeluaran pembiayaan <b>' + rp(r46) + '</b> (' + pct(a46 ? r46 / a46 : 0) + ' dari anggaran ' + rp(a46) + '), terdiri dari pembentukan dana cadangan ' + rp(r48) + ' dan penyertaan modal ' + rp(r47) + '. Pengeluaran ini setara ' + pct(r43 ? r46 / r43 : 0) + ' dari penerimaan pembiayaan.' });
    out.push({ c: 'var(--teal)', ic: '⇅', t: 'Pembiayaan netto', p: 'Pembiayaan netto sebesar <b>' + rp(net) + '</b>, yaitu penerimaan ' + rp(r43) + ' dikurangi pengeluaran ' + rp(r46) + '. Artinya, ' + (net >= 0 ? 'pembiayaan menambah dana untuk menutup defisit dan membentuk SILPA.' : 'pengeluaran pembiayaan lebih besar daripada penerimaannya.') });
    out.push({ c: 'var(--plum)', ic: '∑', t: 'SILPA', p: '<b>SILPA akhir tahun ' + rp(sl) + '</b>: ' + (sur >= 0 ? 'surplus ' : 'defisit ') + rp(Math.abs(sur)) + ' ditambah pembiayaan netto ' + rp(net) + '. SILPA ' + up(g50 || 0) + ' ' + pct(Math.abs(g50 || 0)) + ' dibanding ' + (y - 1) + ' (' + rp(v(y, 50, P)) + ').' });
    return out;
  }
  function insight(y) {
    const sur = v(y, 41, R);
    return 'Pada TA ' + y + ', Kota Kediri menerima pembiayaan ' + rp(v(y, 43, R)) + ' dan mengeluarkan ' + rp(v(y, 46, R)) + ', sehingga pembiayaan netto ' + rp(v(y, 49, R)) + '. Setelah ' + (sur >= 0 ? 'surplus ' : 'defisit ') + rp(Math.abs(sur)) + ', SILPA akhir tahun mencapai ' + rp(v(y, 50, R)) + '.';
  }
  function narasiAll() {
    const f = YRS[0], l = YRS[YRS.length - 1], g = (n) => (v(f, n, R) ? (v(l, n, R) - v(f, n, R)) / Math.abs(v(f, n, R)) : 0), list = (n) => YRS.map((y) => rpT(v(y, n, R))).join(' → ');
    return [
      { c: 'var(--indigo2)', ic: '⬇', t: 'Arah penerimaan', p: 'Penerimaan pembiayaan bergerak <b>' + list(43) + '</b> (' + YRS.join(', ') + '), ' + up(g(43)) + ' ' + pct(Math.abs(g(43))) + ' selama periode ini. Hampir seluruhnya berasal dari penggunaan SiLPA tahun sebelumnya.' },
      { c: 'var(--amber)', ic: '⬆', t: 'Arah pengeluaran', p: 'Pengeluaran pembiayaan: ' + YRS.map((y) => y + ' ' + rpT(v(y, 46, R))).join(', ') + '. Pengeluaran hanya terjadi pada ' + (YRS.filter((y) => v(y, 46, R) > 0).join(', ') || 'tidak ada tahun') + ', yaitu pembentukan dana cadangan.' },
      { c: 'var(--teal)', ic: '⇅', t: 'Pembiayaan netto', p: 'Pembiayaan netto: <b>' + list(49) + '</b>. Nilainya mengikuti besar SiLPA yang digunakan untuk menutup defisit tiap tahun.' },
      { c: 'var(--plum)', ic: '∑', t: 'SILPA', p: 'SILPA akhir tahun: <b>' + list(50) + '</b>. ' + YRS.map((y) => y + ': ' + (v(y, 41, R) >= 0 ? 'surplus ' : 'defisit ') + rpT(Math.abs(v(y, 41, R)))).join('; ') + '.' }
    ];
  }
  const narrHTML = (list) => '<div class="narr' + (list.length === 1 ? ' n1' : '') + '">' + list.map((n, i) => '<article class="nc reveal" style="--c:' + n.c + ';--i:' + i + '"><div class="ic" aria-hidden="true">' + n.ic + '</div><h3>' + n.t + '</h3><p>' + n.p + '</p></article>').join('') + '</div>';

  /* ---------- komponen HTML ---------- */
  const reveal = (i, html, cls) => '<section class="card reveal ' + (cls || '') + '" style="--i:' + i + '">' + html + '</section>';
  const head = (t, s) => '<div class="card-head"><div><h2>' + t + '</h2>' + (s ? '<p class="sub">' + s + '</p>' : '') + '</div></div>';
  const cv = (id, cls, aria) => '<div class="cv ' + (cls || '') + '"><canvas data-ch="' + id + '" role="img" aria-label="' + esc(aria || id) + '"></canvas></div>';
  function fmtK(x, k) { return k === 'pct' ? pct(x) : rpT(x); }
  const stat = (lab, val, kind, em, emCls, c, foot) => '<div class="stat reveal" style="--c:' + c + '"><small>' + lab + '</small><b data-v="' + val + '" data-k="' + kind + '">' + fmtK(val, kind) + '</b><em class="' + (emCls || 'nt') + '">' + em + '</em>' + (foot ? '<span class="foot">' + foot + '</span>' : '') + '</div>';
  const dir = (f) => (f == null ? '' : (f >= 0 ? '▲ ' : '▼ ') + pct(Math.abs(f)) + ' dari tahun lalu');
  const dcls = (f) => (f == null ? 'nt' : f >= 0 ? 'up' : 'dn');
  const dlBtn = (y) => '<a class="btn" href="' + fileOf(y) + '" download>' + ICON.dl + 'Unduh Excel ' + y + '</a>';
  const range = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);
  const sel = (g, cur) => { const G = GROUPS[g]; return '<label class="selw"><span>' + G.title + '</span><select data-sel="' + G.sel + '" aria-label="' + G.title + '"><option value="all"' + (cur === 'all' ? ' selected' : '') + '>Semua ' + lcase(G.label).replace(' pembiayaan', '') + '</option>' + G.items.map((n) => '<option value="' + n + '"' + (+cur === n ? ' selected' : '') + '>' + esc(NAMES[n][1]) + '</option>').join('') + '</select></label>'; };

  function heroHTML(view, y) {
    const titles = {
      ringkasan: ['Pembiayaan Daerah <em>Kota Kediri</em>', insight(y)],
      penerimaan: ['Dari mana <em>penerimaan</em> pembiayaan berasal', 'Penggunaan SiLPA dan pencairan dana cadangan TA ' + y + '. Realisasi ' + rp(v(y, 43, R)) + ' atau ' + pct(cap(y, 43)) + ' dari anggaran.'],
      pengeluaran: ['Ke mana <em>pengeluaran</em> pembiayaan digunakan', 'Penyertaan modal dan pembentukan dana cadangan TA ' + y + '. Realisasi ' + rp(v(y, 46, R)) + (v(y, 46, A) ? ' atau ' + pct(cap(y, 46)) + ' dari anggaran.' : ', tidak ada anggaran pengeluaran.')],
      semua: ['Tiga tahun <em>sekaligus</em>', 'Perbandingan pembiayaan ' + YRS[0] + '–' + YRS[YRS.length - 1] + ': penerimaan, pengeluaran, pembiayaan netto, dan SILPA dalam satu halaman.'],
      lembar: ['Lembar kerja: <em>nilai dan rumus</em>', 'Ketuk sel mana pun untuk melihat rumus yang dipakai, seperti di Excel. Isinya sama persis dengan file Excel TA ' + y + '.'],
      unduh: ['Unduh laporan <em>Excel</em>', 'Satu file untuk setiap tahun. Semua hasil hitung berupa rumus Excel yang bisa diperiksa.'],
      tentang: ['Tentang <em>data</em> ini', 'Sumber, cara membaca, dan catatan penting dari Laporan Realisasi Anggaran Kota Kediri.']
    }[view];
    const showYears = YEAR_VIEWS.includes(view) || view === 'semua';
    const pills = showYears ? '<div class="years" role="group" aria-label="Pilih tahun anggaran">' + YRS.map((yy) => '<button data-year="' + yy + '" aria-pressed="' + (view !== 'semua' && yy === y) + '">' + yy + '</button>').join('') + '<button class="all" data-all="1" aria-pressed="' + (view === 'semua') + '">Semua Tahun</button></div>' : '';
    return '<div class="hero-in"><span class="eyebrow">Kota Kediri · LRA ' + (view === 'semua' ? YRS[0] + '–' + YRS[YRS.length - 1] : 'TA ' + y) + '</span><h1>' + titles[0] + '</h1><p>' + titles[1] + '</p>' + pills + '</div>';
  }

  function statsHTML(view, y) {
    const four = () => stat('Penerimaan Pembiayaan', v(y, 43, R), 'rp', dir(chg(y, 43)), dcls(chg(y, 43)), 'var(--indigo2)', 'Capaian ' + pct(cap(y, 43)) + ' dari anggaran ' + rpT(v(y, 43, A))) +
      stat('Pengeluaran Pembiayaan', v(y, 46, R), 'rp', v(y, 46, P) || v(y, 46, R) ? dir(chg(y, 46)) : 'Tidak ada pengeluaran', 'nt', 'var(--amber)', v(y, 46, A) ? 'Capaian ' + pct(cap(y, 46)) + ' dari anggaran ' + rpT(v(y, 46, A)) : 'Anggaran pengeluaran Rp0') +
      stat('Pembiayaan Netto', v(y, 49, R), 'rp', 'Penerimaan − pengeluaran', 'nt', 'var(--teal)', 'Rasio pengeluaran ' + pct(v(y, 43, R) ? v(y, 46, R) / v(y, 43, R) : 0) + ' dari penerimaan') +
      stat('SILPA', v(y, 50, R), 'rp', dir(chg(y, 50)), dcls(chg(y, 50)), 'var(--plum)', (v(y, 41, R) >= 0 ? 'Surplus ' : 'Defisit ') + rpT(Math.abs(v(y, 41, R))) + ' + netto ' + rpT(v(y, 49, R)));
    if (view === 'ringkasan') return '<div class="stats">' + four() + '</div>';
    if (view === 'penerimaan') return '<div class="stats">' +
      stat('Penerimaan Pembiayaan', v(y, 43, R), 'rp', dir(chg(y, 43)), dcls(chg(y, 43)), 'var(--indigo2)', 'Anggaran ' + rpT(v(y, 43, A))) +
      stat('Penggunaan SiLPA', v(y, 44, R), 'rp', 'Porsi ' + pct(v(y, 43, R) ? v(y, 44, R) / v(y, 43, R) : 0), 'nt', 'var(--teal)', 'Capaian ' + pct(cap(y, 44))) +
      stat('Pencairan Dana Cadangan', v(y, 45, R), 'rp', 'Porsi ' + pct(v(y, 43, R) ? v(y, 45, R) / v(y, 43, R) : 0), 'nt', 'var(--amber)', v(y, 45, A) ? 'Capaian ' + pct(cap(y, 45)) : 'Tidak dianggarkan') +
      stat('Capaian Penerimaan', cap(y, 43), 'pct', 'Realisasi ÷ anggaran', 'nt', 'var(--plum)', 'Sisa/(lebih) anggaran ' + rpT(v(y, 43, A) - v(y, 43, R))) + '</div>';
    if (view === 'pengeluaran') return '<div class="stats">' +
      stat('Pengeluaran Pembiayaan', v(y, 46, R), 'rp', v(y, 46, P) || v(y, 46, R) ? dir(chg(y, 46)) : 'Tidak ada pengeluaran', 'nt', 'var(--amber)', 'Anggaran ' + rpT(v(y, 46, A))) +
      stat('Penyertaan Modal', v(y, 47, R), 'rp', v(y, 47, A) ? 'Capaian ' + pct(cap(y, 47)) : 'Tidak dianggarkan', 'nt', 'var(--indigo2)', 'Realisasi penyertaan modal daerah') +
      stat('Pembentukan Dana Cadangan', v(y, 48, R), 'rp', v(y, 48, A) ? 'Capaian ' + pct(cap(y, 48)) : 'Tidak dianggarkan', 'nt', 'var(--teal)', 'Realisasi pembentukan dana cadangan') +
      stat('Rasio thd Penerimaan', v(y, 43, R) ? v(y, 46, R) / v(y, 43, R) : 0, 'pct', 'Pengeluaran ÷ penerimaan', 'nt', 'var(--plum)', 'Netto ' + rpT(v(y, 49, R))) + '</div>';
    if (view === 'semua') { const f = YRS[0], l = YRS[YRS.length - 1], d = (n) => (v(f, n, R) ? sgn((v(l, n, R) - v(f, n, R)) / Math.abs(v(f, n, R))) : '–');
      return '<div class="stats">' +
        stat('Penerimaan ' + l, v(l, 43, R), 'rp', d(43) + ' vs ' + f, 'nt', 'var(--indigo2)', 'Realisasi, dibanding ' + f) +
        stat('Pengeluaran ' + l, v(l, 46, R), 'rp', 'Tahun ' + f + ': ' + rpT(v(f, 46, R)), 'nt', 'var(--amber)', 'Realisasi pengeluaran pembiayaan') +
        stat('Pembiayaan Netto ' + l, v(l, 49, R), 'rp', d(49) + ' vs ' + f, 'nt', 'var(--teal)', 'Penerimaan − pengeluaran') +
        stat('SILPA ' + l, v(l, 50, R), 'rp', d(50) + ' vs ' + f, 'nt', 'var(--plum)', 'Sisa lebih pembiayaan anggaran') + '</div>'; }
    return '';
  }

  function tableRows(ids, y) {
    return '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Uraian</th><th>Anggaran</th><th>Realisasi</th><th>Capaian</th><th>Realisasi ' + (y - 1) + '</th><th>Perubahan</th><th>Sisa / (Lebih)</th></tr></thead><tbody>' +
      ids.map((n) => {
        const r = NAMES[n], lv = r[2], big = n === 50;
        if (n === 42) return '<tr class="l0"><td colspan="7">' + esc(r[1]) + '</td></tr>';
        const a = v(y, n, A), re = v(y, n, R), p = v(y, n, P), c = a ? re / a : 0, g = p ? (re - p) / Math.abs(p) : null;
        return '<tr class="l' + lv + (big ? ' big' : '') + '"><td>' + esc(r[1]) + '</td><td>' + mil(a) + '</td><td>' + mil(re) + '</td><td class="cp" style="--w:' + Math.min(Math.abs(c), 1.3) / 1.3 * 100 + '%">' + (a ? pct(c) : '–') + '</td><td>' + mil(p) + '</td><td class="' + (g != null && g < 0 ? 'neg' : '') + '">' + (g == null ? '–' : sgn(g)) + '</td><td class="' + (a - re < 0 ? 'neg' : '') + '">' + mil(a - re) + '</td></tr>';
      }).join('') + '</tbody></table></div>';
  }
  const tblCard = (i, t, s, ids, y) => reveal(i, head(t, s + ' · Rp miliar') + tableRows(ids, y) + '<div class="xl-key" style="margin-top:10px"><span>Batang di kolom Capaian = realisasi ÷ anggaran.</span><span>Angka dalam kurung = negatif. Tanda – = nol.</span></div>');

  /* ---------- tampilan ---------- */
  function vRingkasan(y) {
    return '<div class="grid g12">' + reveal(1, head('Penerimaan dan Pengeluaran Pembiayaan', 'Realisasi per tahun, Rp miliar · tahun ' + y + ' disorot') + cv('cmp', 'tall', 'Grafik batang perbandingan penerimaan dan pengeluaran pembiayaan per tahun')) +
      reveal(2, head('Dari Surplus/Defisit ke SILPA ' + y, 'Alur realisasi, Rp miliar') + cv('wf', 'tall', 'Grafik alur surplus defisit, penerimaan, pengeluaran, dan SILPA')) + '</div>' +
      '<div class="grid g2">' + reveal(1, head('Pembiayaan Netto', 'Penerimaan − pengeluaran, Rp miliar · label hijau/merah = perubahan dari tahun lalu') + cv('netto', 'short', 'Grafik area pembiayaan netto')) +
      reveal(2, head('SILPA', 'Sisa lebih pembiayaan anggaran, Rp miliar') + cv('silpa', 'short', 'Grafik area SILPA')) + '</div>' +
      '<h2 class="sec-title">Narasi dan analisis TA ' + y + '</h2><p class="sec-sub">Kalimat di bawah dihitung otomatis dari data LRA tahun yang dipilih.</p>' + narrHTML(narasi(y)) +
      '<div class="grid">' + tblCard(1, 'Tabel Pembiayaan Daerah', 'Anggaran, realisasi, dan capaian TA ' + y, range(41, 50), y) + '</div>';
  }
  function vGroup(g, y) {
    const G = GROUPS[g], cur = state[G.sel], tot = G.total, one = cur === 'all' ? tot : cur, nm = cur === 'all' ? G.label : NAMES[cur][1];
    return '<div class="selbar reveal">' + sel(g, cur) + '<span class="selnote">Pilihan memengaruhi grafik area di bawah: <b>' + esc(nm) + '</b></span></div>' +
      '<div class="grid g2">' + reveal(1, head(esc(nm), 'Realisasi per tahun (area) dan anggaran (garis putus), Rp miliar') + cv(g + 'Area', '', 'Grafik area ' + nm)) +
      reveal(2, head('Anggaran vs Realisasi ' + y, 'Per jenis, Rp miliar') + cv(g + 'Bar', '', 'Grafik batang anggaran dan realisasi ' + G.label)) + '</div>' +
      '<div class="grid g12">' + reveal(1, head('Komposisi per Tahun', 'Realisasi menurut jenis, Rp miliar') + cv(g + 'Stk', 'short', 'Grafik kolom bertumpuk komposisi')) +
      reveal(2, head('Komposisi TA ' + y, 'Porsi tiap jenis terhadap total') + '<div class="gwrap"><div class="cv short"><canvas data-ch="' + g + 'Don" role="img" aria-label="Donat komposisi"></canvas></div><div class="dmid"><b>' + rpT(v(y, tot, R)) + '</b><small>total</small></div></div>') + '</div>' +
      '<h2 class="sec-title">Narasi dan analisis TA ' + y + '</h2>' + narrHTML([narasi(y)[g === 'in' ? 0 : 1]]) +
      '<div class="grid">' + tblCard(1, 'Tabel ' + G.label, 'Anggaran, realisasi, dan capaian TA ' + y, [tot].concat(G.items), y) + '</div>';
  }
  const vPenerimaan = (y) => vGroup('in', y), vPengeluaran = (y) => vGroup('out', y);
  function vSemua() {
    const rowsAll = (ids) => '<div class="tbl-wrap"><table class="tbl" style="min-width:900px"><thead><tr><th>Uraian</th>' + YRS.map((y) => '<th>Realisasi ' + y + '</th>').join('') + YRS.map((y) => '<th>Capaian ' + y + '</th>').join('') + '<th>Δ ' + YRS[2] + ' vs ' + YRS[1] + '</th></tr></thead><tbody>' +
      ids.map((n) => { const r = NAMES[n], big = n === 50; if (n === 42) return '<tr class="l0"><td colspan="8">' + esc(r[1]) + '</td></tr>';
        const g = chg(YRS[2], n);
        return '<tr class="l' + r[2] + (big ? ' big' : '') + '"><td>' + esc(r[1]) + '</td>' + YRS.map((y) => '<td>' + mil(v(y, n, R)) + '</td>').join('') + YRS.map((y) => { const a = v(y, n, A); return '<td class="cp" style="--w:' + Math.min(Math.abs(cap(y, n)), 1.3) / 1.3 * 100 + '%">' + (a ? pct(cap(y, n)) : '–') + '</td>'; }).join('') + '<td class="' + (g != null && g < 0 ? 'neg' : '') + '">' + (g == null ? '–' : sgn(g)) + '</td></tr>'; }).join('') + '</tbody></table></div>';
    return '<div class="grid g12">' + reveal(1, head('Penerimaan, Pengeluaran, Netto, dan SILPA', 'Realisasi tiga tahun, Rp miliar') + cv('trend2', 'tall', 'Grafik garis tiga tahun')) +
      reveal(2, head('Surplus/Defisit, Netto, dan SILPA', 'Perbandingan tiga tahun, Rp miliar') + cv('fin', 'tall', 'Grafik batang surplus defisit netto SILPA')) + '</div>' +
      '<div class="grid g2">' + reveal(1, head('Komposisi Penerimaan per Tahun', 'Penggunaan SiLPA dan pencairan dana cadangan · Rp miliar') + cv('inStk', 'short', 'Grafik kolom komposisi penerimaan')) +
      reveal(2, head('Penerimaan vs Pengeluaran', 'Realisasi per tahun, Rp miliar') + cv('cmp', 'short', 'Grafik batang penerimaan dan pengeluaran')) + '</div>' +
      '<h2 class="sec-title">Narasi dan analisis ' + YRS[0] + '–' + YRS[2] + '</h2>' + narrHTML(narasiAll()) +
      '<div class="grid">' + reveal(1, head('Data Lengkap Tiga Tahun', 'Seluruh baris pembiayaan LRA, realisasi dan capaian · Rp miliar') + rowsAll(range(41, 50)) + '<div class="xl-key" style="margin-top:10px"><span>Geser tabel ke samping di layar kecil.</span></div>') + '</div>';
  }
  function vUnduh() {
    return '<div class="grid g3">' + YRS.map((y, i) => '<section class="card dl reveal" style="--i:' + (i + 1) + '"><div class="yr">' + y + '</div><div class="chips"><span class="chip">Penerimaan <b>' + rpT(v(y, 43, R)) + '</b></span><span class="chip">Pengeluaran <b>' + rpT(v(y, 46, R)) + '</b></span><span class="chip">SILPA <b>' + rpT(v(y, 50, R)) + '</b></span></div><ul><li>Dashboard: 6 angka utama dan 3 grafik</li><li>Rincian LRA: pembiayaan, total dan SILPA berupa rumus</li><li>Rumus &amp; Perhitungan: 15 indikator dan 3 pemeriksaan</li><li>Data Grafik dan Catatan &amp; Sumber</li></ul><a class="btn amb" href="' + fileOf(y) + '" download>' + ICON.dl + 'Unduh Excel TA ' + y + '</a></section>').join('') + '</div>' +
      reveal(1, head('Cara melihat rumus di Excel', '') + '<div class="about"><p>Semua hasil hitung (total penerimaan, total pengeluaran, pembiayaan netto, SILPA, capaian, selisih, perubahan, sisa anggaran, porsi, serta data grafik) adalah <b>rumus Excel</b>. Angka biru adalah salinan dari dokumen sumber, angka hitam adalah hasil rumus.</p><ul><li><b>Laptop:</b> klik sel hasil, rumusnya tampil di bilah rumus (fx) di atas lembar.</li><li><b>Android/iPhone:</b> buka dengan Excel, Google Sheets, atau WPS, ketuk sel. Pratinjau file bawaan HP kadang tidak punya bilah rumus.</li><li><b>Tanpa membuka Excel:</b> pakai menu <a href="#lembar">Lembar Kerja</a>. Ketuk sel dan rumusnya muncul seperti di Excel, juga di HP.</li></ul></div>', 'mt');
  }
  function vTentang() {
    return reveal(1, head('Tentang data dan dashboard', '') + '<div class="about"><img src="assets/logo.png" alt="Logo" style="height:84px;float:right;margin:0 0 8px 16px" onerror="this.remove()"><p><b>Sumber.</b> Laporan Realisasi Anggaran (LRA) Pemerintah Kota Kediri untuk tahun berakhir 31 Desember 2023, 2024, dan 2025.</p><p><b>Cakupan.</b> Pembiayaan daerah: penerimaan pembiayaan (penggunaan SiLPA, pencairan dana cadangan), pengeluaran pembiayaan (penyertaan modal, pembentukan dana cadangan), pembiayaan netto, dan SILPA. Surplus/defisit dari LRA dipakai sebagai angka pembantu untuk menghitung SILPA.</p><p><b>Pemeriksaan data.</b></p><ul><li>Total penerimaan, total pengeluaran, pembiayaan netto, dan SILPA dihitung ulang dengan rumus dan cocok dengan dokumen sumber.</li><li>Kolom pembanding tahun lalu pada LRA 2025 sama dengan realisasi LRA 2024, dan LRA 2024 sama dengan LRA 2023.</li><li>Pada LRA 2023, anggaran Pencairan Dana Cadangan tertulis 14.940.979.980 sedangkan realisasinya 14.940.879.980. Angka disalin sesuai dokumen.</li></ul><p><b>Istilah.</b> Pembiayaan netto = penerimaan − pengeluaran pembiayaan. SILPA = surplus/defisit + pembiayaan netto. Capaian = realisasi ÷ anggaran.</p></div>');
  }

  /* ---------- lembar kerja ---------- */
  const XL_SHEETS = ['Rincian LRA', 'Rumus & Perhitungan', 'Data Grafik'];
  const colL = (n) => { let t = ''; while (n > 0) { const m = (n - 1) % 26; t = String.fromCharCode(65 + m) + t; n = Math.floor((n - 1) / 26); } return t; };
  const colN = (t) => t.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const xlFmt = (x, k) => {
    if (x === null || x === undefined || x === '') return '';
    if (typeof x === 'string') return x;
    switch (k) {
      case 'r': return x === 0 ? '-' : x < 0 ? '(' + F2.format(-x) + ')' : F2.format(x);
      case 'p': return x === 0 ? '-' : x < 0 ? '(' + F2.format(-x * 100) + '%)' : F2.format(x * 100) + '%';
      case 'm': return F2.format(x);
      case 'q': return F2.format(x * 100) + '%';
      case 's': return (x >= 0 ? '+' : '−') + F2.format(Math.abs(x * 100)) + '%';
      case 'o': return F1.format(x * 100) + '%';
      default: return Number.isInteger(x) ? String(x) : F2.format(x);
    }
  };
  function xlRefs(f) {
    const t = f.replace(/'[^']+'!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?/g, '');
    const re = /\$?([A-Z]{1,2})\$?(\d+)(?::\$?([A-Z]{1,2})\$?(\d+))?/g, out = []; let m;
    while ((m = re.exec(t))) out.push(m[3] ? [colN(m[1]), +m[2], colN(m[3]), +m[4]] : [colN(m[1]), +m[2], colN(m[1]), +m[2]]);
    return out;
  }
  function xlExplain(f) {
    const other = /'([^']+)'!/.exec(f);
    if (/ABS\(/.test(f) && /"OK"/.test(f)) return 'Pemeriksaan otomatis: hasilnya OK jika selisihnya kurang dari 1.';
    if (other) return 'Mengambil angka dari sheet ' + other[1] + (/\/10\^9/.test(f) ? ' lalu mengubahnya ke Rp miliar (÷ 1.000.000.000).' : '.');
    if (/^=SUM\(/.test(f)) return 'Penjumlahan sel yang disorot kuning.';
    if (/^=IF\([^=]+=0,0,[A-Z]+\d+\/ABS\([A-Z]+\d+\)\)$/.test(f)) return 'Persentase perubahan = selisih ÷ nilai tahun lalu (0 jika pembagi nol).';
    if (/^=IF\([^=]+=0,0,.+\/.+\)$/.test(f)) return 'Pembagian sel yang disorot kuning (hasil 0 jika pembagi nol).';
    if (/^=[A-Z]+\d+-[A-Z]+\d+$/.test(f)) return 'Selisih: sel pertama dikurangi sel kedua.';
    if (/^=[A-Z]+\d+\+[A-Z]+\d+$/.test(f)) return 'Penjumlahan dua sel yang disorot kuning.';
    if (/^=[A-Z]+\d+(\+[A-Z]+\d+)+$/.test(f)) return 'Penjumlahan sel-sel yang disorot kuning.';
    if (/^=[A-Z]+\d+$/.test(f)) return 'Sama dengan sel yang disorot kuning.';
    return '';
  }
  const xlSheet = () => window.SHEETS[state.year][state.sheet];
  function xlLabels(sh, r, c) {
    const row = sh.rows[r - 1] || []; let rl = '', cl = '';
    for (let i = 0; i < 3; i++) { const x = row[i]; if (x && typeof x[0] === 'string' && !/^[\d.]+$/.test(x[0]) && !(x[3] || '').split(' ').includes('code')) { rl = x[0]; break; } }
    for (let i = r - 1; i >= 1; i--) { const x = (sh.rows[i - 1] || [])[c - 1]; if (x && typeof x[0] === 'string' && (x[3] || '').split(' ').includes('h')) { cl = x[0]; break; } }
    return [rl, cl];
  }
  function xlSelect(addr) {
    const m = /^([A-Z]+)(\d+)$/.exec(addr); if (!m) return;
    const c = colN(m[1]), r = +m[2], sh = xlSheet(), x = (sh.rows[r - 1] || [])[c - 1];
    $$('.xl td.sel,.xl td.pre').forEach((e) => e.classList.remove('sel', 'pre'));
    const td = $('.xl td[data-a="' + addr + '"]'); if (td) td.classList.add('sel');
    const st = x ? (x[3] || '').split(' ') : [];
    let fv = '', note = '', val = '';
    if (x && x[1]) {
      fv = '<code>' + esc(x[1]) + '</code>'; note = xlExplain(x[1]); val = xlFmt(x[0], x[2]);
      xlRefs(x[1]).forEach((q) => { for (let rr = q[1]; rr <= q[3]; rr++) for (let cc = q[0]; cc <= q[2]; cc++) { const t = $('.xl td[data-a="' + colL(cc) + rr + '"]'); if (t && t !== td) t.classList.add('pre'); } });
    } else if (x && typeof x[0] === 'number') {
      fv = '<b>' + esc(xlFmt(x[0], x[2])) + '</b>'; note = st.includes('in') ? 'Angka input: disalin dari dokumen LRA (teks biru). Bukan hasil rumus.' : '';
    } else if (x) { fv = esc(x[0]).replace(/\n/g, ' ⏎ '); note = st.includes('code') ? 'Teks yang menampilkan rumus untuk dibaca (bukan sel rumus).' : ''; }
    const lb = xlLabels(sh, r, c), lab = [lb[0], lb[1]].filter(Boolean).join(' · ');
    $('#xlName').textContent = addr;
    $('#xlFv').innerHTML = fv || '<span class="mut">(sel kosong)</span>';
    $('#xlDesc').innerHTML = (val ? '<div class="hasil">Hasil: <b>' + esc(val) + '</b></div>' : '') + (lab ? '<b>' + esc(lab) + '</b>' : '') + (lab && note ? ' — ' : '') + esc(note);
  }
  function xlTable(sh) {
    const skip = {}, span = {};
    sh.merges.forEach((q) => { span[q[0] + ',' + q[1]] = [q[2] - q[0] + 1, q[3] - q[1] + 1]; for (let r = q[0]; r <= q[2]; r++) for (let c = q[1]; c <= q[3]; c++) if (r !== q[0] || c !== q[1]) skip[r + ',' + c] = 1; });
    const px = sh.w.map((w) => Math.round(w * 7.2 + 10)), total = 44 + px.reduce((a, b) => a + b, 0);
    let h = '<table class="xl" style="width:' + total + 'px"><colgroup><col style="width:44px">' + px.map((w) => '<col style="width:' + w + 'px">').join('') + '</colgroup><thead><tr><th class="corner"></th>' + sh.w.map((_, i) => '<th>' + colL(i + 1) + '</th>').join('') + '</tr></thead><tbody>';
    sh.rows.forEach((row, ri) => {
      const r = ri + 1; h += '<tr><th class="rn">' + r + '</th>';
      for (let c = 1; c <= sh.w.length; c++) {
        if (skip[r + ',' + c]) continue;
        const x = row[c - 1], sp = span[r + ',' + c], cls = [];
        if (x) { (x[3] || '').split(' ').forEach((t) => t && cls.push(t)); if (x[1]) cls.push('fc'); if (typeof x[0] === 'number') cls.push('num'); }
        h += '<td data-a="' + colL(c) + r + '"' + (sp ? ' rowspan="' + sp[0] + '" colspan="' + sp[1] + '"' : '') + (cls.length ? ' class="' + cls.join(' ') + '"' : '') + '>' + (x ? esc(xlFmt(x[0], x[2])) : '') + '</td>';
      }
      h += '</tr>';
    });
    return h + '</tbody></table>';
  }
  function vLembar(y) {
    const tabs = '<div class="seg" role="group" aria-label="Pilih sheet">' + XL_SHEETS.map((n) => '<button data-sheet="' + n + '" aria-pressed="' + (n === state.sheet) + '">' + n + '</button>').join('') + '</div>';
    return '<section class="card reveal" style="--i:1"><div class="card-head"><div><h2>Lembar kerja LRA · ' + state.sheet + '</h2><p class="sub">Ketuk sel untuk melihat alamat dan rumusnya · teks biru = angka input · teks hitam = hasil rumus</p></div>' + dlBtn(y) + '</div>' + tabs +
      '<div class="xl-bar"><div class="fbar"><span class="nb" id="xlName">–</span><span class="fxl">fx</span><span class="fv" id="xlFv"></span></div><div class="fdesc" id="xlDesc"></div></div>' +
      '<div class="xl-wrap" id="xlWrap">' + xlTable(xlSheet()) + '</div>' +
      '<div class="xl-key"><span><i class="k-in"></i>Angka input dari LRA</span><span><i class="k-fx"></i>Hasil rumus</span><span><i class="k-sel"></i>Sel dipilih</span><span><i class="k-pre"></i>Sel yang dipakai rumus</span></div></section>';
  }
  function xlInit() {
    const sh = xlSheet(); let first = null;
    for (let r = 0; r < sh.rows.length && !first; r++) for (let c = 0; c < sh.rows[r].length; c++) { const x = sh.rows[r][c]; if (x && x[1]) { first = colL(c + 1) + (r + 1); break; } }
    xlSelect(first || 'A1');
  }

  /* ---------- grafik ---------- */
  let live = [], observers = [];
  const mil9 = (x) => x / 1e9;
  function baseOpts(T, extra) {
    return Object.assign({ responsive: true, maintainAspectRatio: false, animation: { duration: 1400, easing: 'easeOutQuart', delay: (c) => (c.type === 'data' && c.mode === 'default' ? c.dataIndex * 90 + c.datasetIndex * 140 : 0) }, layout: { padding: { top: 24, right: 8 } },
      plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, pointStyle: 'rectRounded', boxWidth: 9, font: { size: 12 } } }, tooltip: tip(T) } }, extra || {});
  }
  const tip = (T, cb) => ({ backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, boxPadding: 4, callbacks: cb || { label: (c) => ' ' + (c.dataset.label ? c.dataset.label + ': ' : '') + 'Rp' + F2.format(c.parsed.y) + ' miliar' } });
  const axX = (T, o) => Object.assign({ grid: { display: false }, ticks: { color: T.muted, font: { size: 11.5 } }, border: { color: T.line } }, o || {});
  const axY = (T, o) => Object.assign({ grid: { color: T.line }, ticks: { color: T.muted, font: { size: 11.5 }, callback: (t) => F0.format(t) }, border: { display: false } }, o || {});
  function grad(ctx, color, a0, a1) { const ar = ctx.chart.chartArea; if (!ar) return alpha(color, a1); const g = ctx.chart.ctx.createLinearGradient(0, ar.top, 0, ar.bottom); g.addColorStop(0, alpha(color, a0)); g.addColorStop(1, alpha(color, a1)); return g; }
  // label nilai di atas batang
  const valLabels = { id: 'valLabels', afterDatasetsDraw(ch) {
    if (!ch.options.plugins.valLabels || !ch.options.plugins.valLabels.on) return;
    const T = pal(), c = ch.ctx; c.save(); c.font = '600 11.5px DM Sans, sans-serif'; c.textAlign = 'center'; c.fillStyle = T.ink;
    ch.data.datasets.forEach((ds, di) => { if (ch.isDatasetVisible(di)) ch.getDatasetMeta(di).data.forEach((bar, i) => { const x = ds.data[i]; const val = Array.isArray(x) ? (ds.vals ? ds.vals[i] : x[1] - x[0]) : x; c.fillText((val < 0 ? '−' : '') + F1.format(Math.abs(val)), bar.x, Math.min(bar.y, bar.base) - 6); }); });
    c.restore(); } };
  // label nilai + perubahan di titik grafik area
  const ptLabels = { id: 'ptLabels', afterDatasetsDraw(ch) {
    const o = ch.options.plugins.ptLabels; if (!o || !o.on) return; const T = pal(), c = ch.ctx, m = ch.getDatasetMeta(o.ds || 0), ds = ch.data.datasets[o.ds || 0];
    c.save(); c.textAlign = 'center'; m.data.forEach((pt, i) => { const val = ds.data[i], g = ds.chgs ? ds.chgs[i] : null;
      c.font = '700 12px DM Sans, sans-serif'; c.fillStyle = T.ink; c.fillText((val < 0 ? '−' : '') + F1.format(Math.abs(val)), pt.x, pt.y - (g == null ? 12 : 26));
      if (g != null) { c.font = '600 11px DM Sans, sans-serif'; c.fillStyle = g >= 0 ? T.teal : T.coral; c.fillText((g >= 0 ? '▲ ' : '▼ ') + F2.format(Math.abs(g * 100)) + '%', pt.x, pt.y - 11); } });
    c.restore(); } };
  const labelsOf = () => YRS.map(String);
  function areaChart(cx, y, spec) {       // spec: {ids:[no], color, name, withBudget}
    const T = pal(), col = T[spec.color], f = (n, j) => YRS.map((yy) => mil9(v(yy, n, j)));
    const first = (n) => YRS.map((yy) => chg(yy, n));
    const ds = [{ label: 'Realisasi', data: f(spec.no, R), borderColor: col, backgroundColor: (c) => grad(c, col, 0.42, 0.03), fill: true, tension: 0.35, borderWidth: 3, pointRadius: YRS.map((yy) => (yy === y ? 8 : 5)), pointBackgroundColor: T.card, pointBorderColor: col, pointBorderWidth: 3, chgs: first(spec.no), order: 1 }];
    if (spec.budget) ds.push({ label: 'Anggaran', data: f(spec.no, A), borderColor: T.muted, borderDash: [6, 5], borderWidth: 2, pointRadius: 0, fill: false, tension: 0.35, order: 2 });
    return new Chart(cx, { type: 'line', data: { labels: labelsOf(), datasets: ds }, plugins: [ptLabels],
      options: baseOpts(T, { layout: { padding: { top: 34, left: 8, right: 16 } }, interaction: { mode: 'index', intersect: false }, scales: { x: axX(T, { offset: true }), y: axY(T, { beginAtZero: true, grace: '22%' }) }, plugins: { legend: { display: !!spec.budget, position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, ptLabels: { on: true }, tooltip: tip(T) } }) });
  }
  function groupBars(cx, y, ids, names, cols) {    // anggaran vs realisasi tahun y
    const T = pal();
    return new Chart(cx, { type: 'bar', data: { labels: names, datasets: [{ label: 'Anggaran', data: ids.map((n) => mil9(v(y, n, A))), backgroundColor: T.lilac, borderRadius: 8, borderSkipped: false }, { label: 'Realisasi', data: ids.map((n) => mil9(v(y, n, R))), backgroundColor: cols || T.indigo2, borderRadius: 8, borderSkipped: false }] },
      plugins: [valLabels], options: baseOpts(T, { scales: { x: axX(T), y: axY(T, { beginAtZero: true, suggestedMax: 5, grace: '14%' }) }, plugins: { valLabels: { on: true }, legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: tip(T) } }) });
  }
  function stacked(cx, ids, names) {
    const T = pal(), cols = [T.indigo2, T.amber, T.teal, T.plum];
    return new Chart(cx, { type: 'bar', data: { labels: labelsOf(), datasets: ids.map((n, i) => ({ label: names[i], data: YRS.map((yy) => mil9(v(yy, n, R))), backgroundColor: cols[i], borderRadius: 6, borderSkipped: false, borderColor: T.card, borderWidth: 1.5 })) },
      options: baseOpts(T, { scales: { x: axX(T, { stacked: true }), y: axY(T, { stacked: true, beginAtZero: true, suggestedMax: 5 }) }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: tip(T) } }) });
  }
  function donut(cx, y, ids, names) {
    const T = pal(), cols = [T.indigo2, T.amber, T.teal], vals = ids.map((n) => v(y, n, R)), tot = vals.reduce((a, b) => a + b, 0);
    return new Chart(cx, { type: 'doughnut', data: { labels: tot ? names : ['Tidak ada realisasi'], datasets: [{ data: tot ? vals : [1], backgroundColor: tot ? cols : [alpha(T.muted, 0.2)], borderColor: T.card, borderWidth: 3, borderRadius: 8, hoverOffset: 8 }] },
      options: { responsive: true, maintainAspectRatio: false, cutout: '66%', animation: { animateRotate: true, animateScale: true, duration: 1500, easing: 'easeOutQuart' }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: tip(T, { label: (c) => tot ? ' ' + c.label + ': ' + rpT(c.raw) + ' (' + pct(c.raw / tot) + ')' : ' Rp0' }) } } });
  }
  const wfPlugin = { id: 'wfLabels', afterDatasetsDraw(ch) {
    const T = pal(), c = ch.ctx, m = ch.getDatasetMeta(0), ds = ch.data.datasets[0]; c.save(); c.font = '700 12px DM Sans, sans-serif'; c.textAlign = 'center'; c.fillStyle = T.ink;
    m.data.forEach((bar, i) => { const val = ds.vals[i]; c.fillText((val < 0 ? '−' : '') + F1.format(Math.abs(val)), bar.x, Math.min(bar.y, bar.base) - 7); }); c.restore(); } };
  const MAKE = {
    cmp: (cx, y) => {
      const T = pal(), isAll = route() === 'semua', col = (c) => YRS.map((yy) => (isAll || yy === y ? c : alpha(c, 0.45)));
      return new Chart(cx, { type: 'bar', data: { labels: labelsOf(), datasets: [{ label: 'Penerimaan Pembiayaan', data: YRS.map((yy) => mil9(v(yy, 43, R))), backgroundColor: col(T.indigo2), borderRadius: 10, borderSkipped: false }, { label: 'Pengeluaran Pembiayaan', data: YRS.map((yy) => mil9(v(yy, 46, R))), backgroundColor: col(T.amber), borderRadius: 10, borderSkipped: false }] },
        plugins: [valLabels], options: baseOpts(T, { scales: { x: axX(T), y: axY(T, { beginAtZero: true, grace: '14%', title: { display: true, text: 'Rp miliar', color: T.muted } }) }, plugins: { valLabels: { on: true }, legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: tip(T) } }) });
    },
    wf: (cx, y) => {
      const T = pal(), s = mil9(v(y, 41, R)), a = mil9(v(y, 43, R)), b = mil9(v(y, 46, R)), sl = mil9(v(y, 50, R));
      const bars = [[Math.min(0, s), Math.max(0, s)], [s, s + a], [s + a - b, s + a], [Math.min(0, sl), Math.max(0, sl)]];
      return new Chart(cx, { type: 'bar', data: { labels: [s >= 0 ? 'Surplus' : 'Defisit', 'Penerimaan', 'Pengeluaran', 'SILPA'], datasets: [{ data: bars, vals: [s, a, -b, sl], backgroundColor: [s >= 0 ? T.teal : T.coral, T.indigo2, T.amber, T.plum], borderRadius: 10, borderSkipped: false, barPercentage: 0.7 }] },
        plugins: [wfPlugin], options: baseOpts(T, { layout: { padding: { top: 26 } }, scales: { x: axX(T), y: axY(T) }, plugins: { legend: { display: false }, tooltip: tip(T, { label: (c) => ' ' + (c.dataset.vals[c.dataIndex] < 0 ? '−' : '') + 'Rp' + F2.format(Math.abs(c.dataset.vals[c.dataIndex])) + ' miliar' }) } }) });
    },
    netto: (cx, y) => areaChart(cx, y, { no: 49, color: 'teal' }),
    silpa: (cx, y) => areaChart(cx, y, { no: 50, color: 'plum' }),
    inArea: (cx, y) => areaChart(cx, y, { no: state.jin === 'all' ? 43 : state.jin, color: 'indigo2', budget: true }),
    outArea: (cx, y) => areaChart(cx, y, { no: state.jout === 'all' ? 46 : state.jout, color: 'amber', budget: true }),
    inBar: (cx, y) => groupBars(cx, y, [44, 45], ['Penggunaan SiLPA', 'Pencairan Dana Cadangan']),
    outBar: (cx, y) => groupBars(cx, y, [47, 48], ['Penyertaan Modal', 'Pembentukan Dana Cadangan'], pal().amber),
    inStk: (cx) => stacked(cx, [44, 45], ['Penggunaan SiLPA', 'Pencairan Dana Cadangan']),
    outStk: (cx) => stacked(cx, [47, 48], ['Penyertaan Modal', 'Pembentukan Dana Cadangan']),
    inDon: (cx, y) => donut(cx, y, [44, 45], ['Penggunaan SiLPA', 'Pencairan Dana Cadangan']),
    outDon: (cx, y) => donut(cx, y, [47, 48], ['Penyertaan Modal', 'Pembentukan Dana Cadangan']),
    trend2: (cx) => {
      const T = pal(), mk = (lab, n, col, dash) => ({ label: lab, data: YRS.map((yy) => mil9(v(yy, n, R))), borderColor: col, backgroundColor: (c) => grad(c, col, 0.22, 0.01), fill: n === 43, tension: 0.35, borderWidth: 3, borderDash: dash || [], pointRadius: 5, pointBackgroundColor: T.card, pointBorderColor: col, pointBorderWidth: 3 });
      return new Chart(cx, { type: 'line', data: { labels: labelsOf(), datasets: [mk('Penerimaan', 43, T.indigo2), mk('Pengeluaran', 46, T.amber), mk('Pembiayaan Netto', 49, T.teal, [6, 5]), mk('SILPA', 50, T.plum)] },
        options: baseOpts(T, { interaction: { mode: 'index', intersect: false }, scales: { x: axX(T, { offset: true }), y: axY(T, { beginAtZero: true, grace: '10%' }) }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: tip(T) } }) });
    },
    fin: (cx) => {
      const T = pal(), ids = [41, 49, 50], names = ['Surplus/(Defisit)', 'Pembiayaan netto', 'SILPA'], cols = [T.coral, T.teal, T.plum];
      return new Chart(cx, { type: 'bar', data: { labels: labelsOf(), datasets: ids.map((n, i) => ({ label: names[i], data: YRS.map((yy) => mil9(v(yy, n, R))), backgroundColor: cols[i], borderRadius: 8, borderSkipped: false })) },
        plugins: [valLabels], options: baseOpts(T, { scales: { x: axX(T), y: axY(T, { grace: '12%' }) }, plugins: { valLabels: { on: true }, legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: tip(T) } }) });
    }
  };
  function mountCharts(y) {
    if (!window.Chart) return;
    const T = pal(); Chart.defaults.font.family = "'DM Sans','Segoe UI',system-ui,sans-serif"; Chart.defaults.color = T.muted;
    const build = (cn) => { if (cn.dataset.built || !MAKE[cn.dataset.ch]) return; cn.dataset.built = '1'; try { live.push(MAKE[cn.dataset.ch](cn.getContext('2d'), y)); } catch (e) { console.error(e); } };
    const io = 'IntersectionObserver' in window ? new IntersectionObserver((es) => es.forEach((e) => { if (e.isIntersecting) { build(e.target); io.unobserve(e.target); } }), { threshold: 0.2 }) : null;
    observers.push(io); $$('canvas[data-ch]').forEach((cn) => (io ? io.observe(cn) : build(cn)));
  }
  function countUp() {
    $$('.stat b[data-v]').forEach((el) => { const to = +el.dataset.v, k = el.dataset.k, t0 = performance.now(), d = 1100; if (matchMedia('(prefers-reduced-motion: reduce)').matches) return;
      (function tick(t) { const p = Math.min(1, (t - t0) / d), e = 1 - Math.pow(1 - p, 3); el.textContent = fmtK(to * e, k); if (p < 1) requestAnimationFrame(tick); else el.textContent = fmtK(to, k); })(t0); });
  }
  function reveals() {
    const els = $$('.reveal'); if (!('IntersectionObserver' in window)) { els.forEach((e) => e.classList.add('in')); return; }
    const io = new IntersectionObserver((es) => es.forEach((e) => { if (e.isIntersecting) { e.target.classList.add('in'); io.unobserve(e.target); } }), { threshold: 0.08 }); observers.push(io); els.forEach((e) => io.observe(e));
  }

  /* ---------- render ---------- */
  function route() { const h = (location.hash || '').replace('#', ''); return VIEWS.some((x) => x[0] === h) ? h : 'ringkasan'; }
  function render() {
    live.forEach((c) => c.destroy()); live = []; observers.forEach((o) => o && o.disconnect()); observers = [];
    const view = route(), y = state.year;
    $('#nav').innerHTML = VIEWS.map((x) => '<a href="#' + x[0] + '"' + (x[0] === view ? ' aria-current="page"' : '') + '>' + x[1] + '</a>').join('');
    $('#hero').innerHTML = heroHTML(view, y);
    const pre = statsHTML(view, y);
    const body = { ringkasan: vRingkasan, penerimaan: vPenerimaan, pengeluaran: vPengeluaran, semua: vSemua, lembar: vLembar, unduh: vUnduh, tentang: vTentang }[view](y);
    $('#app').innerHTML = (pre || '<div style="height:20px"></div>') + body;
    if (view === 'lembar') xlInit();
    mountCharts(y); reveals(); countUp();
    document.title = 'Dashboard Pembiayaan Daerah · Kota Kediri · ' + (VIEWS.find((x) => x[0] === view) || [])[1];
  }
  document.addEventListener('click', (e) => {
    const yb = e.target.closest('[data-year]'); if (yb) { state.year = +yb.dataset.year; store.set('kdr.year', state.year); if (route() === 'semua') location.hash = '#ringkasan'; else render(); return; }
    if (e.target.closest('[data-all]')) { location.hash = '#semua'; return; }
    const sb = e.target.closest('[data-sheet]'); if (sb) { state.sheet = sb.dataset.sheet; render(); return; }
    const td = e.target.closest('.xl td[data-a]'); if (td) xlSelect(td.dataset.a);
  });
  document.addEventListener('change', (e) => { const s = e.target.closest('select[data-sel]'); if (s) { state[s.dataset.sel] = s.value === 'all' ? 'all' : +s.value; render(); } });
  $('#themeBtn').addEventListener('click', () => { state.theme = state.theme === 'dark' ? 'light' : 'dark'; store.set('kdr.theme', state.theme); applyTheme(); render(); });
  window.addEventListener('hashchange', () => { render(); window.scrollTo({ top: 0 }); });
  applyTheme(); render();
})();
