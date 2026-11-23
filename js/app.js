/* Dashboard Anggaran dan Pendapatan Daerah — Kota Kediri (LRA 2023–2025). Tanpa framework; Chart.js untuk grafik. */
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
  const rp = (x) => {                                   // ringkas, dengan satuan kata
    const a = Math.abs(x), s = x < 0 ? '−' : '';
    if (a >= 1e12) return s + 'Rp' + F2.format(a / 1e12) + ' triliun';
    if (a >= 1e9) return s + 'Rp' + F2.format(a / 1e9) + ' miliar';
    if (a >= 1e6) return s + 'Rp' + F2.format(a / 1e6) + ' juta';
    return s + 'Rp' + F0.format(a);
  };
  const rpT = (x) => { const a = Math.abs(x), s = x < 0 ? '−' : ''; return a >= 1e12 ? s + 'Rp' + F2.format(a / 1e12) + ' T' : s + 'Rp' + F2.format(a / 1e9) + ' M'; };
  const mil = (x) => (x === 0 ? '–' : x < 0 ? '(' + F2.format(-x / 1e9) + ')' : F2.format(x / 1e9));
  const up = (f) => (f >= 0 ? 'naik' : 'turun');

  /* ---------- status ---------- */
  const store = { get(k, d) { try { const x = localStorage.getItem(k); return x === null ? d : x; } catch (e) { return d; } }, set(k, x) { try { localStorage.setItem(k, x); } catch (e) { /* abaikan */ } } };
  const state = { year: +store.get('kdr.year', 2025), theme: store.get('kdr.theme', window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'), sheet: 'Rincian LRA' };
  if (!YRS.includes(state.year)) state.year = 2025;
  const VIEWS = [['ringkasan', 'Ringkasan'], ['pendapatan', 'Pendapatan'], ['belanja', 'Belanja'], ['pembiayaan', 'Pembiayaan'], ['semua', 'Semua Tahun'], ['lembar', 'Lembar Kerja'], ['unduh', 'Unduh Excel'], ['tentang', 'Tentang Data']];
  const YEAR_VIEWS = ['ringkasan', 'pendapatan', 'belanja', 'pembiayaan', 'lembar'];
  const fileOf = (y) => 'downloads/Laporan-APBD-Kota-Kediri-' + y + '.xlsx';
  const ICON = {
    dl: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12m0 0l-4-4m4 4l4-4M4 19h16"/></svg>',
    sun: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2m0 16v2M2 12h2m16 0h2M5 5l1.5 1.5M17.5 17.5L19 19M5 19l1.5-1.5M17.5 6.5L19 5"/></svg>',
    moon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 13a9 9 0 11-10-10 7 7 0 0010 10z"/></svg>'
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
  function topOf(y, ids) { return ids.slice().sort((a, b) => v(y, b, R) - v(y, a, R))[0]; }
  function label(n) { return NAMES[n][1].replace(/ - LRA$/, '').replace(/^Pendapatan /, '').replace(/^Belanja /, 'Belanja '); }
  function narasi(y) {
    const aP = v(y, 1, A), rP = v(y, 1, R), gP = chg(y, 1), rB = v(y, 24, R), aB = v(y, 24, A), gB = chg(y, 24), sur = v(y, 41, R), net = v(y, 49, R), sil = v(y, 50, R);
    const pad = v(y, 2, R), tr = v(y, 7, R), sPad = rP ? pad / rP : 0;
    const padTop = topOf(y, [3, 4, 5, 6]), padBest = [3, 4, 5, 6].filter((n) => v(y, n, A) > 0).sort((a, b) => cap(y, b) - cap(y, a))[0];
    const trTop = topOf(y, [9, 12, 13, 16, 18]);
    const belTop = topOf(y, [26, 27, 28, 29, 32, 33, 34, 35, 37]);
    const kelB = [25, 30, 37, 39].filter((n) => v(y, n, R) > 0);
    const out = [];
    out.push({ c: 'var(--indigo2)', ic: '◎', t: 'Pendapatan daerah', p: 'Pendapatan terealisasi <b>' + rp(rP) + '</b>, yaitu <b>' + pct(aP ? rP / aP : 0) + '</b> dari target ' + rp(aP) + '. Realisasi ' + (rP >= aP ? 'melampaui' : 'berada di bawah') + ' target sebesar <b>' + rp(Math.abs(rP - aP)) + '</b>. Dibanding ' + (y - 1) + ', pendapatan ' + up(gP || 0) + ' <b>' + pct(Math.abs(gP || 0)) + '</b>.' });
    out.push({ c: 'var(--amber)', ic: '▲', t: 'Kemandirian fiskal', p: 'PAD mencapai <b>' + rp(pad) + '</b> (capaian ' + pct(cap(y, 2)) + ') dan menyumbang <b>' + pct(sPad) + '</b> dari pendapatan; sisanya <b>' + pct(1 - sPad) + '</b> berasal dari transfer. Sumber PAD terbesar adalah ' + label(padTop).toLowerCase() + ' (' + rpT(v(y, padTop, R)) + '), dengan capaian tertinggi pada ' + label(padBest).toLowerCase() + ' (' + pct(cap(y, padBest)) + ').' });
    out.push({ c: 'var(--teal)', ic: '⇄', t: 'Pendapatan transfer', p: 'Transfer sebesar <b>' + rp(tr) + '</b> dengan capaian <b>' + pct(cap(y, 7)) + '</b>. Komponen terbesar: ' + label(trTop) + ' (' + rpT(v(y, trTop, R)) + ', ' + pct(v(y, trTop, R) / tr) + ' dari transfer). Dana Alokasi Umum sendiri ' + rpT(v(y, 12, R)) + ' dan Dana Alokasi Khusus ' + rpT(v(y, 13, R)) + '.' });
    out.push({ c: 'var(--coral)', ic: '◧', t: 'Belanja daerah', p: 'Belanja terealisasi <b>' + rp(rB) + '</b> (<b>' + pct(aB ? rB / aB : 0) + '</b> dari anggaran ' + rp(aB) + '), sehingga ' + rp(aB - rB) + ' belum terserap. Belanja ' + up(gB || 0) + ' ' + pct(Math.abs(gB || 0)) + ' dibanding ' + (y - 1) + '. Belanja operasi menyerap ' + pct(v(y, 25, R) / rB) + ', belanja modal ' + pct(v(y, 30, R) / rB) + ' (capaian ' + pct(cap(y, 30)) + '). Pos terbesar: ' + label(belTop).toLowerCase() + ' (' + rpT(v(y, belTop, R)) + ').' });
    out.push({ c: 'var(--plum)', ic: '∑', t: 'Surplus/defisit dan SILPA', p: (sur >= 0 ? 'Pendapatan melebihi belanja sehingga terjadi <b>surplus ' + rp(sur) + '</b>. ' : 'Belanja melebihi pendapatan sehingga terjadi <b>defisit ' + rp(-sur) + '</b>. ') + 'Pembiayaan netto ' + rp(net) + ' (penggunaan SiLPA tahun sebelumnya ' + rp(v(y, 44, R)) + ') membuat <b>SILPA akhir tahun ' + rp(sil) + '</b>.' });
    return out;
  }
  function insight(y) {
    const rP = v(y, 1, R), sur = v(y, 41, R);
    return 'Pada TA ' + y + ', Kota Kediri merealisasikan pendapatan ' + rp(rP) + ' (' + pct(cap(y, 1)) + ' dari target) dan belanja ' + rp(v(y, 24, R)) + ' (' + pct(cap(y, 24)) + ' dari anggaran). ' + (sur >= 0 ? 'Hasilnya surplus ' : 'Hasilnya defisit ') + rp(Math.abs(sur)) + ' dan SILPA ' + rp(v(y, 50, R)) + '.';
  }
  function narasiAll() {
    const f = YRS[0], l = YRS[YRS.length - 1], g = (n) => (v(f, n, R) ? (v(l, n, R) - v(f, n, R)) / v(f, n, R) : 0);
    const share = (y) => v(y, 2, R) / v(y, 1, R);
    const mx = (n) => YRS.slice().sort((a, b) => v(b, n, R) - v(a, n, R))[0];
    return [
      { c: 'var(--indigo2)', ic: '◎', t: 'Arah pendapatan', p: 'Pendapatan ' + up(g(1)) + ' dari <b>' + rp(v(f, 1, R)) + '</b> (' + f + ') menjadi <b>' + rp(v(l, 1, R)) + '</b> (' + l + '), atau ' + sgn(g(1)) + '. Capaian terhadap target: ' + YRS.map((y) => y + ' ' + pct(cap(y, 1))).join(', ') + '.' },
      { c: 'var(--amber)', ic: '▲', t: 'Peran PAD', p: 'Porsi PAD terhadap pendapatan bergerak ' + YRS.map((y) => pct(share(y))).join(' → ') + '. PAD ' + up(g(2)) + ' ' + pct(Math.abs(g(2))) + ' selama periode ini, sedangkan transfer ' + up(g(7)) + ' ' + pct(Math.abs(g(7))) + '.' },
      { c: 'var(--coral)', ic: '◧', t: 'Arah belanja', p: 'Belanja ' + up(g(24)) + ' dari <b>' + rp(v(f, 24, R)) + '</b> menjadi <b>' + rp(v(l, 24, R)) + '</b>. Daya serap anggaran: ' + YRS.map((y) => y + ' ' + pct(cap(y, 24))).join(', ') + '. Belanja modal tertinggi terjadi pada ' + mx(30) + ' (' + rp(v(mx(30), 30, R)) + ').' },
      { c: 'var(--plum)', ic: '∑', t: 'Surplus/defisit & SILPA', p: YRS.map((y) => y + ': ' + (v(y, 41, R) >= 0 ? 'surplus ' : 'defisit ') + rpT(Math.abs(v(y, 41, R)))).join('; ') + '. SILPA akhir tahun: ' + YRS.map((y) => rpT(v(y, 50, R))).join(' → ') + '.' }
    ];
  }
  const narrHTML = (list) => '<div class="narr' + (list.length === 5 ? ' n5' : list.length === 1 ? ' n1' : '') + '">' + list.map((n, i) => '<article class="nc reveal" style="--c:' + n.c + ';--i:' + i + '"><div class="ic" aria-hidden="true">' + n.ic + '</div><h3>' + n.t + '</h3><p>' + n.p + '</p></article>').join('') + '</div>';

  /* ---------- komponen HTML ---------- */
  const reveal = (i, html, cls) => '<section class="card reveal ' + (cls || '') + '" style="--i:' + i + '">' + html + '</section>';
  const head = (t, s) => '<div class="card-head"><div><h2>' + t + '</h2>' + (s ? '<p class="sub">' + s + '</p>' : '') + '</div></div>';
  const cv = (id, cls, aria) => '<div class="cv ' + (cls || '') + '"><canvas data-ch="' + id + '" role="img" aria-label="' + esc(aria || id) + '"></canvas></div>';
  const stat = (lab, val, kind, em, emCls, c, foot) => '<div class="stat reveal" style="--c:' + c + '"><small>' + lab + '</small><b data-v="' + val + '" data-k="' + kind + '">' + fmtK(val, kind) + '</b><em class="' + (emCls || 'nt') + '">' + em + '</em>' + (foot ? '<span class="foot">' + foot + '</span>' : '') + '</div>';
  function fmtK(x, k) { return k === 'pct' ? pct(x) : rpT(x); }
  const dir = (f) => (f == null ? '' : (f >= 0 ? '▲ ' : '▼ ') + pct(Math.abs(f)) + ' dari tahun lalu');
  const dcls = (f) => (f == null ? 'nt' : f >= 0 ? 'up' : 'dn');
  const dlBtn = (y) => '<a class="btn" href="' + fileOf(y) + '" download>' + ICON.dl + 'Unduh Excel ' + y + '</a>';

  function heroHTML(view, y) {
    const titles = {
      ringkasan: ['Anggaran dan <em>Pendapatan Daerah</em> Kota Kediri', insight(y)],
      pendapatan: ['Dari mana <em>pendapatan</em> daerah berasal', 'Rincian target dan realisasi PAD, dana transfer, dan lain-lain pendapatan yang sah TA ' + y + '. Capaian ' + pct(cap(y, 1)) + ' dengan porsi PAD ' + pct(v(y, 2, R) / v(y, 1, R)) + '.'],
      belanja: ['Ke mana <em>belanja</em> daerah mengalir', 'Serapan anggaran belanja operasi, modal, tak terduga, dan transfer TA ' + y + '. Daya serap ' + pct(cap(y, 24)) + ' dari anggaran ' + rp(v(y, 24, A)) + '.'],
      pembiayaan: ['Surplus, defisit, dan <em>SILPA</em>', 'Pendapatan dikurangi belanja, ditambah pembiayaan netto, menghasilkan SILPA ' + rp(v(y, 50, R)) + ' pada TA ' + y + '.'],
      semua: ['Tiga tahun <em>sekaligus</em>', 'Perbandingan lengkap LRA ' + YRS[0] + '–' + YRS[YRS.length - 1] + ': pendapatan, belanja, surplus/defisit, pembiayaan, dan SILPA dalam satu halaman.'],
      lembar: ['Lembar kerja: <em>nilai dan rumus</em>', 'Ketuk sel mana pun untuk melihat rumus yang dipakai, seperti di Excel. Isinya sama persis dengan file Excel TA ' + y + '.'],
      unduh: ['Unduh laporan <em>Excel</em>', 'Satu file untuk setiap tahun. Semua hasil hitung berupa rumus Excel yang bisa diperiksa.'],
      tentang: ['Tentang <em>data</em> ini', 'Sumber, cara membaca, dan catatan penting dari Laporan Realisasi Anggaran Kota Kediri.']
    }[view];
    const showYears = YEAR_VIEWS.includes(view) || view === 'semua';
    const pills = showYears ? '<div class="years" role="group" aria-label="Pilih tahun anggaran">' + YRS.map((yy) => '<button data-year="' + yy + '" aria-pressed="' + (view !== 'semua' && yy === y) + '">' + yy + '</button>').join('') + '<button class="all" data-all="1" aria-pressed="' + (view === 'semua') + '">Semua Tahun</button></div>' : '';
    return '<div class="hero-in"><span class="eyebrow">Kota Kediri · LRA ' + (view === 'semua' ? YRS[0] + '–' + YRS[YRS.length - 1] : 'TA ' + y) + '</span><h1>' + titles[0] + '</h1><p>' + titles[1] + '</p>' + pills + '</div>';
  }

  function statsHTML(view, y) {
    if (view === 'ringkasan') return '<div class="stats">' +
      stat('Realisasi Pendapatan', v(y, 1, R), 'rp', dir(chg(y, 1)), dcls(chg(y, 1)), 'var(--indigo2)', 'Capaian ' + pct(cap(y, 1)) + ' dari target ' + rpT(v(y, 1, A))) +
      stat('Realisasi Belanja', v(y, 24, R), 'rp', dir(chg(y, 24)), dcls(-(chg(y, 24) || 0) + 0), 'var(--coral)', 'Daya serap ' + pct(cap(y, 24)) + ' dari anggaran ' + rpT(v(y, 24, A))) +
      stat(v(y, 41, R) >= 0 ? 'Surplus Daerah' : 'Defisit Daerah', v(y, 41, R), 'rp', 'Pendapatan − belanja', 'nt', 'var(--teal)', 'Ditutup/ditambah pembiayaan netto ' + rpT(v(y, 49, R))) +
      stat('SILPA Akhir Tahun', v(y, 50, R), 'rp', dir(chg(y, 50)), 'nt', 'var(--amber)', 'Surplus/defisit + pembiayaan netto') + '</div>';
    if (view === 'pendapatan') return '<div class="stats">' +
      stat('Realisasi Pendapatan', v(y, 1, R), 'rp', dir(chg(y, 1)), dcls(chg(y, 1)), 'var(--indigo2)', 'Target ' + rpT(v(y, 1, A))) +
      stat('Pendapatan Asli Daerah', v(y, 2, R), 'rp', 'Capaian ' + pct(cap(y, 2)), cap(y, 2) >= 1 ? 'up' : 'dn', 'var(--amber)', 'Porsi ' + pct(v(y, 2, R) / v(y, 1, R)) + ' dari pendapatan') +
      stat('Pendapatan Transfer', v(y, 7, R), 'rp', 'Capaian ' + pct(cap(y, 7)), cap(y, 7) >= 1 ? 'up' : 'dn', 'var(--teal)', 'Porsi ' + pct(v(y, 7, R) / v(y, 1, R)) + ' dari pendapatan') +
      stat('Melampaui / (di bawah) target', v(y, 1, R) - v(y, 1, A), 'rp', 'Realisasi − anggaran', 'nt', 'var(--plum)', 'Capaian total ' + pct(cap(y, 1))) + '</div>';
    if (view === 'belanja') return '<div class="stats">' +
      stat('Realisasi Belanja', v(y, 24, R), 'rp', dir(chg(y, 24)), 'nt', 'var(--coral)', 'Anggaran ' + rpT(v(y, 24, A))) +
      stat('Daya Serap Anggaran', cap(y, 24), 'pct', 'Realisasi ÷ anggaran', 'nt', 'var(--indigo2)', 'Sisa anggaran ' + rpT(v(y, 24, A) - v(y, 24, R))) +
      stat('Belanja Operasi', v(y, 25, R), 'rp', 'Capaian ' + pct(cap(y, 25)), 'nt', 'var(--amber)', 'Porsi ' + pct(v(y, 25, R) / v(y, 24, R)) + ' dari belanja') +
      stat('Belanja Modal', v(y, 30, R), 'rp', 'Capaian ' + pct(cap(y, 30)), 'nt', 'var(--teal)', 'Porsi ' + pct(v(y, 30, R) / v(y, 24, R)) + ' dari belanja') + '</div>';
    if (view === 'pembiayaan') return '<div class="stats">' +
      stat(v(y, 41, R) >= 0 ? 'Surplus' : 'Defisit', v(y, 41, R), 'rp', 'Pendapatan − belanja', 'nt', 'var(--teal)', 'Anggaran ' + rpT(v(y, 41, A))) +
      stat('Penerimaan Pembiayaan', v(y, 43, R), 'rp', 'Capaian ' + pct(cap(y, 43)), 'nt', 'var(--indigo2)', 'Penggunaan SiLPA ' + rpT(v(y, 44, R))) +
      stat('Pembiayaan Netto', v(y, 49, R), 'rp', 'Penerimaan − pengeluaran', 'nt', 'var(--amber)', 'Pengeluaran ' + rpT(v(y, 46, R))) +
      stat('SILPA Akhir Tahun', v(y, 50, R), 'rp', dir(chg(y, 50)), 'nt', 'var(--plum)', 'Surplus/defisit + pembiayaan netto') + '</div>';
    if (view === 'semua') { const f = YRS[0], l = YRS[YRS.length - 1]; return '<div class="stats">' +
      stat('Pendapatan ' + l, v(l, 1, R), 'rp', sgn((v(l, 1, R) - v(f, 1, R)) / v(f, 1, R)) + ' vs ' + f, 'nt', 'var(--indigo2)', 'Realisasi, naik/turun dibanding ' + f) +
      stat('Belanja ' + l, v(l, 24, R), 'rp', sgn((v(l, 24, R) - v(f, 24, R)) / v(f, 24, R)) + ' vs ' + f, 'nt', 'var(--coral)', 'Realisasi, naik/turun dibanding ' + f) +
      stat('Porsi PAD ' + l, v(l, 2, R) / v(l, 1, R), 'pct', 'dari ' + pct(v(f, 2, R) / v(f, 1, R)) + ' pada ' + f, 'nt', 'var(--amber)', 'Pendapatan asli daerah ÷ pendapatan') +
      stat('SILPA ' + l, v(l, 50, R), 'rp', 'dari ' + rpT(v(f, 50, R)) + ' pada ' + f, 'nt', 'var(--plum)', 'Sisa lebih pembiayaan anggaran') + '</div>'; }
    return '';
  }

  function tableRows(ids, y) {
    return '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Uraian</th><th>Anggaran</th><th>Realisasi</th><th>Capaian</th><th>Realisasi ' + (y - 1) + '</th><th>Perubahan</th><th>Sisa / (Lebih)</th></tr></thead><tbody>' +
      ids.map((n) => {
        const r = NAMES[n], lv = r[2], big = n === 1 || n === 24 || n === 50;
        if (n === 42) return '<tr class="l0"><td colspan="7">' + esc(r[1]) + '</td></tr>';
        const a = v(y, n, A), re = v(y, n, R), p = v(y, n, P), c = a ? re / a : 0, g = p ? (re - p) / Math.abs(p) : null;
        return '<tr class="l' + (lv === 0 ? '0' : lv) + (big ? ' big' : '') + '"><td>' + esc(r[1]) + '</td><td>' + mil(a) + '</td><td>' + mil(re) + '</td><td class="cp" style="--w:' + Math.min(Math.abs(c), 1.3) / 1.3 * 100 + '%">' + (a ? pct(c) : '–') + '</td><td>' + mil(p) + '</td><td class="' + (g != null && g < 0 ? 'neg' : '') + '">' + (g == null ? '–' : sgn(g)) + '</td><td class="' + (a - re < 0 ? 'neg' : '') + '">' + mil(a - re) + '</td></tr>';
      }).join('') + '</tbody></table></div>';
  }
  const tblCard = (i, t, s, ids, y) => reveal(i, head(t, s + ' · Rp miliar') + tableRows(ids, y) + '<div class="xl-key" style="margin-top:10px"><span>Batang di kolom Capaian = realisasi ÷ anggaran.</span><span>Angka dalam kurung = negatif.</span></div>');
  const range = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);

  /* ---------- tampilan ---------- */
  function vRingkasan(y) {
    const rl = (c, lab, val, sub, i) => '<li style="--c:' + c + '"><span class="dot"></span><div><small>' + lab + '</small><b>' + val + '</b><em>' + sub + '</em></div></li>';
    return reveal(1, head('Tiga Cincin Kinerja TA ' + y, 'Cincin luar ke dalam: capaian pendapatan, daya serap belanja, porsi PAD · satu putaran penuh = 100%') +
      '<div class="rings"><div class="cv ringcv"><canvas data-ch="rings" role="img" aria-label="Grafik tiga cincin kinerja"></canvas><div class="dmid"><b>' + y + '</b><small>tahun anggaran</small></div></div><ul class="rlist">' +
      rl('var(--indigo2)', 'Capaian pendapatan', pct(cap(y, 1)), (cap(y, 1) >= 1 ? 'melampaui target ' : 'dari target ') + rpT(v(y, 1, A))) +
      rl('var(--coral)', 'Daya serap belanja', pct(cap(y, 24)), 'dari anggaran ' + rpT(v(y, 24, A))) +
      rl('var(--amber)', 'Porsi PAD', pct(v(y, 2, R) / v(y, 1, R)), 'pendapatan berasal dari PAD, sisanya transfer') + '</ul></div>') +
      '<div class="grid g12">' + reveal(1, head('Pendapatan dan Belanja ' + YRS[0] + '–' + YRS[YRS.length - 1], 'Realisasi per tahun, Rp triliun · ketuk titik untuk memilih tahun') + cv('trend', '', 'Grafik garis pendapatan dan belanja')) +
      reveal(2, head('Dari Pendapatan sampai SILPA ' + y, 'Alur realisasi, Rp miliar') + cv('wf', '', 'Grafik alur pendapatan, belanja, surplus, pembiayaan, SILPA')) + '</div>' +
      '<h2 class="sec-title">Narasi dan analisis TA ' + y + '</h2><p class="sec-sub">Kalimat di bawah dihitung otomatis dari data LRA tahun yang dipilih.</p>' + narrHTML(narasi(y));
  }
  function vPendapatan(y) {
    return '<div class="grid g21">' + reveal(1, head('Anggaran vs Realisasi per Sumber', 'Batang lebar = anggaran (target), batang tipis = realisasi · Rp miliar') + cv('bullet', 'tall', 'Grafik bullet anggaran dan realisasi per sumber pendapatan')) +
      reveal(2, head('Komposisi Pendapatan', 'Porsi realisasi') + '<div class="gwrap"><div class="cv"><canvas data-ch="kom" role="img" aria-label="Diagram donat komposisi pendapatan"></canvas></div><div class="dmid"><b>' + rpT(v(y, 1, R)) + '</b><small>total pendapatan</small></div></div><div class="legend"><span><i style="background:var(--indigo2)"></i>PAD</span><span><i style="background:var(--amber)"></i>Transfer</span><span><i style="background:var(--teal)"></i>Lain-lain</span></div>') + '</div>' +
      '<div class="grid g2">' + reveal(1, head('Rincian PAD', 'Realisasi per sumber, Rp miliar') + cv('pad', 'short', 'Grafik area kutub rincian PAD')) +
      reveal(2, head('Komposisi Pendapatan Transfer', 'Porsi tiap komponen terhadap total transfer') + cv('trf', 'short', 'Grafik batang tersegmentasi komposisi transfer')) + '</div>' +
      '<h2 class="sec-title">Narasi dan analisis TA ' + y + '</h2>' + narrHTML(narasi(y).slice(0, 3)) +
      '<div class="grid">' + tblCard(1, 'Tabel Pendapatan Daerah', 'Anggaran, realisasi, dan capaian TA ' + y, range(1, 23), y) + '</div>';
  }
  function vBelanja(y) {
    return '<div class="grid g12">' + reveal(1, head('Pos Belanja Terbesar', 'Realisasi TA ' + y + ' diurutkan dari terbesar, Rp miliar') + cv('rank', 'tall', 'Grafik peringkat pos belanja')) +
      reveal(2, head('Profil Capaian Belanja', 'Capaian tiap jenis belanja, tiga tahun · 100% = anggaran terserap penuh') + cv('radar', 'tall', 'Grafik radar capaian belanja')) + '</div>' +
      '<div class="grid">' + reveal(1, head('Komposisi Belanja per Tahun', 'Realisasi menurut jenis, Rp miliar') + cv('stk', 'short', 'Grafik kolom bertumpuk belanja per jenis')) + '</div>' +
      '<h2 class="sec-title">Narasi dan analisis TA ' + y + '</h2>' + narrHTML([narasi(y)[3]]) +
      '<div class="grid">' + tblCard(1, 'Tabel Belanja Daerah', 'Anggaran, realisasi, dan capaian TA ' + y, range(24, 40), y) + '</div>';
  }
  function vPembiayaan(y) {
    return '<div class="grid g12">' + reveal(1, head('Alur ke SILPA ' + y, 'Pendapatan → belanja → surplus/defisit → pembiayaan → SILPA, Rp miliar') + cv('wf', 'tall', 'Grafik alur SILPA')) +
      reveal(2, head('Surplus/Defisit, Pembiayaan, dan SILPA', 'Perbandingan tiga tahun, Rp miliar') + cv('fin', 'tall', 'Grafik batang kelompok surplus, pembiayaan, SILPA')) + '</div>' +
      '<h2 class="sec-title">Narasi dan analisis TA ' + y + '</h2>' + narrHTML([narasi(y)[4]]) +
      '<div class="grid">' + tblCard(1, 'Tabel Surplus/Defisit dan Pembiayaan', 'Realisasi TA ' + y, range(41, 50), y) + '</div>';
  }
  function vSemua() {
    const rowsAll = (ids) => '<div class="tbl-wrap"><table class="tbl" style="min-width:900px"><thead><tr><th>Uraian</th>' + YRS.map((y) => '<th>Realisasi ' + y + '</th>').join('') + YRS.map((y) => '<th>Capaian ' + y + '</th>').join('') + '<th>Δ ' + YRS[2] + ' vs ' + YRS[1] + '</th></tr></thead><tbody>' +
      ids.map((n) => { const r = NAMES[n], big = n === 1 || n === 24 || n === 50; if (n === 42) return '<tr class="l0"><td colspan="8">' + esc(r[1]) + '</td></tr>';
        const g = chg(YRS[2], n);
        return '<tr class="l' + r[2] + (big ? ' big' : '') + '"><td>' + esc(r[1]) + '</td>' + YRS.map((y) => '<td>' + mil(v(y, n, R)) + '</td>').join('') + YRS.map((y) => { const a = v(y, n, A); return '<td class="cp" style="--w:' + Math.min(Math.abs(cap(y, n)), 1.3) / 1.3 * 100 + '%">' + (a ? pct(cap(y, n)) : '–') + '</td>'; }).join('') + '<td class="' + (g != null && g < 0 ? 'neg' : '') + '">' + (g == null ? '–' : sgn(g)) + '</td></tr>'; }).join('') + '</tbody></table></div>';
    return '<div class="grid g12">' + reveal(1, head('Pendapatan, Belanja, dan SILPA', 'Realisasi tiga tahun, Rp triliun') + cv('trend2', '', 'Grafik garis tiga tahun')) +
      reveal(2, head('Capaian per Indikator', 'Realisasi ÷ anggaran, tiga tahun') + cv('capall', '', 'Grafik kolom capaian per indikator')) + '</div>' +
      '<div class="grid g2">' + reveal(1, head('Komposisi Pendapatan per Tahun', 'PAD, transfer, lain-lain · Rp miliar') + cv('komall', 'short', 'Grafik kolom komposisi pendapatan')) +
      reveal(2, head('Komposisi Belanja per Tahun', 'Operasi, modal, tak terduga, transfer · Rp miliar') + cv('stk', 'short', 'Grafik kolom komposisi belanja')) + '</div>' +
      '<h2 class="sec-title">Narasi dan analisis ' + YRS[0] + '–' + YRS[2] + '</h2>' + narrHTML(narasiAll()) +
      '<div class="grid">' + reveal(1, head('Data Lengkap Tiga Tahun', 'Seluruh baris LRA (nomor 1–50), realisasi dan capaian · Rp miliar') + rowsAll(range(1, 50)) + '<div class="xl-key" style="margin-top:10px"><span>Geser tabel ke samping di layar kecil.</span></div>') + '</div>';
  }
  function vUnduh() {
    return '<div class="grid g3">' + YRS.map((y, i) => '<section class="card dl reveal" style="--i:' + (i + 1) + '"><div class="yr">' + y + '</div><div class="chips"><span class="chip">Pendapatan <b>' + rpT(v(y, 1, R)) + '</b></span><span class="chip">Belanja <b>' + rpT(v(y, 24, R)) + '</b></span><span class="chip">SILPA <b>' + rpT(v(y, 50, R)) + '</b></span></div><ul><li>Dashboard: 6 angka utama dan 3 grafik</li><li>Rincian LRA: 50 baris, subtotal dan total berupa rumus</li><li>Rumus &amp; Perhitungan: 16 indikator dan 3 pemeriksaan</li><li>Data Grafik dan Catatan &amp; Sumber</li></ul><a class="btn amb" href="' + fileOf(y) + '" download>' + ICON.dl + 'Unduh Excel TA ' + y + '</a></section>').join('') + '</div>' +
      reveal(1, head('Cara melihat rumus di Excel', '') + '<div class="about"><p>Semua hasil hitung (subtotal, total, surplus/defisit, SILPA, capaian, selisih, perubahan, sisa anggaran, porsi, serta data grafik) adalah <b>rumus Excel</b>. Angka biru adalah salinan dari dokumen sumber, angka hitam adalah hasil rumus.</p><ul><li><b>Laptop:</b> klik sel hasil, rumusnya tampil di bilah rumus (fx) di atas lembar.</li><li><b>Android/iPhone:</b> buka dengan Excel, Google Sheets, atau WPS, ketuk sel. Pratinjau file bawaan HP kadang tidak punya bilah rumus.</li><li><b>Tanpa membuka Excel:</b> pakai menu <a href="#lembar">Lembar Kerja</a>. Ketuk sel dan rumusnya muncul seperti di Excel, juga di HP.</li></ul></div>', 'mt');
  }
  function vTentang() {
    return reveal(1, head('Tentang data dan dashboard', '') + '<div class="about"><img src="assets/logo.png" alt="Logo" style="height:84px;float:right;margin:0 0 8px 16px" onerror="this.remove()"><p><b>Sumber.</b> Laporan Realisasi Anggaran (LRA) Pemerintah Kota Kediri untuk tahun berakhir 31 Desember 2023, 2024, dan 2025.</p><p><b>Cakupan.</b> Pendapatan (PAD, transfer, lain-lain), belanja (operasi, modal, tak terduga, transfer), surplus/defisit, pembiayaan, dan SILPA. Seluruh 50 baris LRA.</p><p><b>Pemeriksaan data.</b></p><ul><li>Semua subtotal dihitung ulang dan cocok dengan dokumen sumber. Satu pengecualian: LRA 2023 dimuat dalam rupiah penuh sehingga total anggaran pendapatan, anggaran surplus/defisit, dan anggaran SILPA berselisih 1 rupiah karena pembulatan.</li><li>Kolom pembanding tahun lalu pada LRA 2025 sama dengan realisasi LRA 2024, dan LRA 2024 sama dengan LRA 2023 (selisih paling besar 1 rupiah karena pembulatan 2023).</li><li>Pada LRA 2023, anggaran Pencairan Dana Cadangan tertulis 14.940.979.980 sedangkan realisasinya 14.940.879.980. Angka disalin sesuai dokumen.</li></ul><p><b>Istilah.</b> Capaian = realisasi ÷ anggaran. SILPA = surplus/defisit + pembiayaan netto. Sisa/(Lebih) anggaran = anggaran − realisasi.</p></div>');
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
  const lastY = () => YRS[YRS.length - 1];
  function baseOpts(T, extra) {
    return Object.assign({ responsive: true, maintainAspectRatio: false, animation: { duration: 1400, easing: 'easeOutQuart', delay: (c) => (c.type === 'data' && c.mode === 'default' ? c.dataIndex * 90 + c.datasetIndex * 140 : 0) }, layout: { padding: { top: 6, right: 8 } },
      plugins: { legend: { labels: { color: T.text, usePointStyle: true, pointStyle: 'rectRounded', boxWidth: 9, font: { size: 12 } } }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, boxPadding: 4 } } }, extra || {});
  }
  const axX = (T, o) => Object.assign({ grid: { display: false }, ticks: { color: T.muted, font: { size: 11.5 } }, border: { color: T.line } }, o || {});
  const axY = (T, o) => Object.assign({ grid: { color: T.line }, ticks: { color: T.muted, font: { size: 11.5 } }, border: { display: false } }, o || {});
  const tipRp = (ctx) => ' ' + (ctx.dataset.label ? ctx.dataset.label + ': ' : '') + 'Rp' + F2.format(ctx.parsed.y !== undefined && ctx.chart.options.indexAxis !== 'y' ? ctx.parsed.y : ctx.parsed.x) + ' miliar';
  function grad(ctx, color, a0, a1) { const ar = ctx.chart.chartArea; if (!ar) return alpha(color, a1); const g = ctx.chart.ctx.createLinearGradient(0, ar.top, 0, ar.bottom); g.addColorStop(0, alpha(color, a0)); g.addColorStop(1, alpha(color, a1)); return g; }
  const wfPlugin = { id: 'wfLabels', afterDatasetsDraw(ch) {
    const T = pal(), c = ch.ctx, m = ch.getDatasetMeta(0); c.save(); c.font = '600 11.5px DM Sans, sans-serif'; c.textAlign = 'center'; c.fillStyle = T.ink;
    m.data.forEach((bar, i) => { const d = ch.data.datasets[0].data[i], val = d[1] - d[0], top = Math.min(bar.y, bar.base); c.fillText((val < 0 ? '−' : '') + F0.format(Math.abs(val)), bar.x, top - 6); });
    c.restore(); } };
  const MAKE = {
    rings: (cx, y) => {
      const T = pal(), vals = [cap(y, 1), cap(y, 24), v(y, 2, R) / v(y, 1, R)], cols = [T.indigo2, T.coral, T.amber], names = ['Capaian pendapatan', 'Daya serap belanja', 'Porsi PAD'];
      return new Chart(cx, { type: 'doughnut', data: { labels: ['Terisi', 'Sisa'], datasets: vals.map((x, i) => ({ data: [Math.min(x, 1), 1 - Math.min(x, 1)], backgroundColor: [cols[i], alpha(T.muted, 0.16)], borderColor: T.card, borderWidth: 7, borderRadius: [30, 0], spacing: 0, weight: 1 })) },
        options: { responsive: true, maintainAspectRatio: false, cutout: '34%', rotation: 0, circumference: 360, animation: { animateRotate: true, animateScale: true, duration: 1900, easing: 'easeOutQuart', delay: (c) => (c.type === 'data' && c.mode === 'default' ? c.datasetIndex * 260 : 0) },
          plugins: { legend: { display: false }, tooltip: { filter: (t) => t.dataIndex === 0, backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { title: (it) => names[it[0].datasetIndex], label: (c) => ' ' + pct(vals[c.datasetIndex]) } } } } });
    },
    trend: (cx, y) => {
      const T = pal();
      const mk = (lab, n, col, dash) => ({ label: lab, data: YRS.map((yy) => mil9(v(yy, n, R)) / 1000), borderColor: col, backgroundColor: (c) => grad(c, col, 0.35, 0.02), fill: true, tension: 0.4, borderWidth: 3, borderDash: dash || [], pointRadius: YRS.map((yy) => (yy === y ? 8 : 5)), pointBackgroundColor: T.card, pointBorderColor: col, pointBorderWidth: 3, pointHoverRadius: 9 });
      return new Chart(cx, { type: 'line', data: { labels: YRS.map(String), datasets: [mk('Pendapatan', 1, T.indigo2), mk('Belanja', 24, T.coral)] },
        options: baseOpts(T, { interaction: { mode: 'index', intersect: false }, onClick: (e, els, ch) => { const p = ch.getElementsAtEventForMode(e, 'index', { intersect: false }, true)[0]; if (p) { state.year = YRS[p.index]; store.set('kdr.year', state.year); render(); } },
          scales: { x: axX(T), y: axY(T, { ticks: { color: T.muted, callback: (t) => F2.format(t) + ' T' }, suggestedMin: 1.2 }) }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: (c) => ' ' + c.dataset.label + ': Rp' + F2.format(c.parsed.y) + ' triliun' } } } }) });
    },
    trend2: (cx) => {
      const T = pal();
      const mk = (lab, n, col, dash) => ({ label: lab, data: YRS.map((yy) => mil9(v(yy, n, R)) / 1000), borderColor: col, backgroundColor: (c) => grad(c, col, 0.3, 0.01), fill: n !== 50, tension: 0.4, borderWidth: 3, borderDash: dash || [], pointRadius: 5, pointBackgroundColor: T.card, pointBorderColor: col, pointBorderWidth: 3 });
      return new Chart(cx, { type: 'line', data: { labels: YRS.map(String), datasets: [mk('Pendapatan', 1, T.indigo2), mk('Belanja', 24, T.coral), mk('SILPA', 50, T.amber, [6, 5])] },
        options: baseOpts(T, { interaction: { mode: 'index', intersect: false }, scales: { x: axX(T), y: axY(T, { ticks: { color: T.muted, callback: (t) => F2.format(t) + ' T' }, suggestedMin: 0 }) }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: (c) => ' ' + c.dataset.label + ': Rp' + F2.format(c.parsed.y) + ' triliun' } } } }) });
    },
    wf: (cx, y) => {
      const T = pal(), p = mil9(v(y, 1, R)), b = mil9(v(y, 24, R)), s = mil9(v(y, 41, R)), n = mil9(v(y, 49, R)), sl = mil9(v(y, 50, R));
      const bars = [[0, p], [p - b, p], [Math.min(0, s), Math.max(0, s)], [s, s + n], [0, sl]];
      const cols = [T.indigo2, T.coral, s >= 0 ? T.teal : T.coral, T.plum, T.amber];
      return new Chart(cx, { type: 'bar', data: { labels: ['Pendapatan', 'Belanja', s >= 0 ? 'Surplus' : 'Defisit', 'Pembiayaan netto', 'SILPA'], datasets: [{ data: bars, backgroundColor: cols, borderRadius: 10, borderSkipped: false, barPercentage: 0.7 }] },
        options: baseOpts(T, { layout: { padding: { top: 22 } }, scales: { x: axX(T), y: axY(T, { ticks: { color: T.muted, callback: (t) => F0.format(t) } }) }, plugins: { legend: { display: false }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: (c) => { const d = c.raw, val = d[1] - d[0]; return ' ' + (val < 0 ? '−' : '') + 'Rp' + F2.format(Math.abs(val)) + ' miliar'; } } } } }), plugins: [wfPlugin] });
    },
    bullet: (cx, y) => {
      const T = pal(), ids = [3, 4, 5, 6, 9, 12, 13, 16, 18, 21].filter((n) => v(y, n, A) > 0 || v(y, n, R) > 0);
      return new Chart(cx, { type: 'bar', data: { labels: ids.map((n) => label(n).replace('Pendapatan ', '').replace('Dana Alokasi Umum (DAU)', 'DAU').replace('Dana Alokasi Khusus', 'DAK').replace('Transfer Pemerintah Pusat - Lainnya', 'Transfer Pusat Lainnya').replace('Transfer Pemerintah Daerah Lainnya', 'Transfer Pemda Lainnya').replace('Hasil Pengelolaan Kekayaan Daerah yang Dipisahkan', 'Kekayaan Dipisahkan')),
        datasets: [{ label: 'Anggaran', data: ids.map((n) => mil9(v(y, n, A))), backgroundColor: T.lilac, borderRadius: 8, barThickness: 24, grouped: false, order: 2 }, { label: 'Realisasi', data: ids.map((n) => mil9(v(y, n, R))), backgroundColor: ids.map((n) => (v(y, n, R) >= v(y, n, A) ? T.teal : T.indigo)), borderRadius: 6, barThickness: 10, grouped: false, order: 1 }] },
        options: baseOpts(T, { indexAxis: 'y', scales: { x: axY(T, { ticks: { color: T.muted, callback: (t) => F0.format(t) } }), y: axX(T, { ticks: { color: T.text, font: { size: 11.5 } } }) }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: (c) => ' ' + c.dataset.label + ': Rp' + F2.format(c.parsed.x) + ' miliar' + (c.datasetIndex === 1 && v(y, ids[c.dataIndex], A) ? ' (' + pct(cap(y, ids[c.dataIndex])) + ')' : '') } } } }) });
    },
    kom: (cx, y) => {
      const T = pal(), d = [2, 7, 21].map((n) => v(y, n, R));
      return new Chart(cx, { type: 'doughnut', data: { labels: ['PAD', 'Transfer', 'Lain-lain'], datasets: [{ data: d, backgroundColor: [T.indigo2, T.amber, T.teal], borderColor: T.card, borderWidth: 3, borderRadius: 8, hoverOffset: 8 }] },
        options: { responsive: true, maintainAspectRatio: false, cutout: '68%', animation: { animateRotate: true, duration: 1500, easing: 'easeOutQuart' }, plugins: { legend: { display: false }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: (c) => ' ' + c.label + ': ' + rpT(c.raw) + ' (' + pct(c.raw / d.reduce((a, b) => a + b, 0)) + ')' } } } } });
    },
    pad: (cx, y) => {
      const T = pal(), ids = [3, 4, 5, 6], cols = [T.indigo2, T.amber, T.teal, T.plum];
      return new Chart(cx, { type: 'polarArea', data: { labels: ['Pajak Daerah', 'Retribusi', 'Kekayaan Dipisahkan', 'Lain-lain PAD Sah'], datasets: [{ data: ids.map((n) => mil9(v(y, n, R))), backgroundColor: cols.map((c) => alpha(c, 0.72)), borderColor: cols, borderWidth: 2 }] },
        options: baseOpts(T, { scales: { r: { grid: { color: T.line }, ticks: { display: false }, angleLines: { color: T.line } } }, plugins: { legend: { position: 'right', labels: { color: T.text, usePointStyle: true, boxWidth: 9, font: { size: 12 } } }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: (c) => ' Rp' + F2.format(c.raw) + ' miliar' } } }, animation: { duration: 1500, easing: 'easeOutQuart' } }) });
    },
    trf: (cx, y) => {
      const T = pal(), ids = [9, 12, 13, 16, 18], cols = [T.indigo2, T.amber, T.teal, T.plum, T.coral], tot = v(y, 7, R);
      const names = ['Dana Bagi Hasil', 'DAU', 'DAK', 'Transfer Pusat Lainnya', 'Transfer Pemda Lainnya'];
      return new Chart(cx, { type: 'bar', data: { labels: ['Transfer ' + y], datasets: ids.map((n, i) => ({ label: names[i], data: [v(y, n, R) / tot * 100], backgroundColor: cols[i], borderRadius: 8, borderSkipped: false, borderColor: T.card, borderWidth: 2 })) },
        options: baseOpts(T, { indexAxis: 'y', scales: { x: axY(T, { stacked: true, max: 100, ticks: { color: T.muted, callback: (t) => t + '%' } }), y: axX(T, { stacked: true, display: false }) }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9, font: { size: 11.5 } } }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: (c) => ' ' + c.dataset.label + ': ' + F2.format(c.raw) + '% (' + rpT(v(y, ids[c.datasetIndex], R)) + ')' } } } }) });
    },
    rank: (cx, y) => {
      const T = pal(), ids = [26, 27, 28, 29, 32, 33, 34, 35, 36, 37, 39].filter((n) => v(y, n, R) > 0).sort((a, b) => v(y, b, R) - v(y, a, R)), col = (n) => (n <= 29 ? T.indigo2 : n <= 36 ? T.amber : T.plum);
      return new Chart(cx, { type: 'bar', data: { labels: ids.map((n) => (n === 37 ? 'Tak Terduga' : n === 39 ? 'Transfer' : NAMES[n][1].replace('Belanja Modal ', 'Modal ').replace('Belanja ', '').replace('Jalan, Irigasi dan Jaringan', 'Jalan, Irigasi, Jaringan'))), datasets: [{ data: ids.map((n) => mil9(v(y, n, R))), backgroundColor: ids.map(col), borderRadius: 99, borderSkipped: false, barThickness: 18 }] },
        options: baseOpts(T, { indexAxis: 'y', scales: { x: axY(T, { ticks: { color: T.muted, callback: (t) => F0.format(t) } }), y: axX(T, { ticks: { color: T.text, font: { size: 11.5 } } }) }, plugins: { legend: { display: false }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: (c) => ' Rp' + F2.format(c.raw) + ' miliar · capaian ' + pct(cap(y, ids[c.dataIndex])) } } } }) });
    },
    radar: (cx, y) => {
      const T = pal(), ids = [25, 30, 37], cols = [T.muted, T.teal, T.coral];
      return new Chart(cx, { type: 'radar', data: { labels: ['Belanja Operasi', 'Belanja Modal', 'Tak Terduga', 'Total Belanja'], datasets: YRS.map((yy, i) => ({ label: String(yy), data: ids.concat([24]).map((n) => cap(yy, n) * 100), borderColor: yy === y ? T.amber : [T.indigo2, T.plum, T.teal][i], backgroundColor: yy === y ? alpha(T.amber, 0.28) : 'transparent', borderWidth: yy === y ? 3 : 1.5, pointRadius: yy === y ? 4 : 2, pointBackgroundColor: yy === y ? T.amber : [T.indigo2, T.plum, T.teal][i] })) },
        options: baseOpts(T, { scales: { r: { min: 0, max: 100, grid: { color: T.line }, angleLines: { color: T.line }, pointLabels: { color: T.text, font: { size: 12 } }, ticks: { color: T.muted, backdropColor: 'transparent', stepSize: 25, callback: (t) => t + '%' } } }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: (c) => ' ' + c.dataset.label + ': ' + F2.format(c.raw) + '%' } } } }) });
    },
    stk: (cx, y) => {
      const T = pal(), ids = [25, 30, 37, 39], names = ['Operasi', 'Modal', 'Tak Terduga', 'Transfer'], cols = [T.indigo2, T.amber, T.teal, T.plum];
      return new Chart(cx, { type: 'bar', data: { labels: YRS.map(String), datasets: ids.map((n, i) => ({ label: names[i], data: YRS.map((yy) => mil9(v(yy, n, R))), backgroundColor: cols[i], borderRadius: 6, borderSkipped: false, borderColor: T.card, borderWidth: 1.5 })) },
        options: baseOpts(T, { scales: { x: axX(T, { stacked: true }), y: axY(T, { stacked: true, ticks: { color: T.muted, callback: (t) => F0.format(t) } }) }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: tipRp } } } }) });
    },
    komall: (cx) => {
      const T = pal(), ids = [2, 7, 21], names = ['PAD', 'Transfer', 'Lain-lain'], cols = [T.indigo2, T.amber, T.teal];
      return new Chart(cx, { type: 'bar', data: { labels: YRS.map(String), datasets: ids.map((n, i) => ({ label: names[i], data: YRS.map((yy) => mil9(v(yy, n, R))), backgroundColor: cols[i], borderRadius: 6, borderSkipped: false, borderColor: T.card, borderWidth: 1.5 })) },
        options: baseOpts(T, { scales: { x: axX(T, { stacked: true }), y: axY(T, { stacked: true, ticks: { color: T.muted, callback: (t) => F0.format(t) } }) }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: tipRp } } } }) });
    },
    capall: (cx) => {
      const T = pal(), ids = [1, 2, 7, 24, 25, 30], names = ['Pendapatan', 'PAD', 'Transfer', 'Belanja', 'B. Operasi', 'B. Modal'], cols = [T.indigo2, T.amber, T.teal];
      return new Chart(cx, { type: 'bar', data: { labels: names, datasets: YRS.map((yy, i) => ({ label: String(yy), data: ids.map((n) => cap(yy, n) * 100), backgroundColor: cols[i], borderRadius: 6, borderSkipped: false })) },
        options: baseOpts(T, { scales: { x: axX(T), y: axY(T, { max: 130, ticks: { color: T.muted, callback: (t) => t + '%' } }) }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: (c) => ' ' + c.dataset.label + ': ' + F2.format(c.raw) + '%' } } } }) });
    },
    fin: (cx) => {
      const T = pal(), ids = [41, 49, 50], names = ['Surplus/(Defisit)', 'Pembiayaan netto', 'SILPA'], cols = [T.teal, T.plum, T.amber];
      return new Chart(cx, { type: 'bar', data: { labels: YRS.map(String), datasets: ids.map((n, i) => ({ label: names[i], data: YRS.map((yy) => mil9(v(yy, n, R))), backgroundColor: cols[i], borderRadius: 8, borderSkipped: false })) },
        options: baseOpts(T, { scales: { x: axX(T), y: axY(T, { ticks: { color: T.muted, callback: (t) => F0.format(t) } }) }, plugins: { legend: { position: 'bottom', labels: { color: T.text, usePointStyle: true, boxWidth: 9 } }, tooltip: { backgroundColor: T.ink, titleColor: T.card, bodyColor: T.card, padding: 10, cornerRadius: 10, callbacks: { label: tipRp } } } }) });
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
    const pre = view === 'semua' ? statsHTML('semua', y) : statsHTML(view, y);
    const body = { ringkasan: vRingkasan, pendapatan: vPendapatan, belanja: vBelanja, pembiayaan: vPembiayaan, semua: vSemua, lembar: vLembar, unduh: vUnduh, tentang: vTentang }[view](y);
    $('#app').innerHTML = (pre || '<div style="height:20px"></div>') + body;
    if (view === 'lembar') xlInit();
    mountCharts(y); reveals(); countUp();
    document.title = 'Dashboard Anggaran dan Pendapatan Daerah · Kota Kediri · ' + (VIEWS.find((x) => x[0] === view) || [])[1];
  }
  document.addEventListener('click', (e) => {
    const yb = e.target.closest('[data-year]'); if (yb) { state.year = +yb.dataset.year; store.set('kdr.year', state.year); if (route() === 'semua') location.hash = '#ringkasan'; else render(); return; }
    if (e.target.closest('[data-all]')) { location.hash = '#semua'; return; }
    const sb = e.target.closest('[data-sheet]'); if (sb) { state.sheet = sb.dataset.sheet; render(); return; }
    const td = e.target.closest('.xl td[data-a]'); if (td) xlSelect(td.dataset.a);
  });
  $('#themeBtn').addEventListener('click', () => { state.theme = state.theme === 'dark' ? 'light' : 'dark'; store.set('kdr.theme', state.theme); applyTheme(); render(); });
  window.addEventListener('hashchange', () => { render(); window.scrollTo({ top: 0 }); });
  applyTheme(); render();
})();
