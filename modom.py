const pptxgen = require('pptxgenjs');

const NAVY = "1E2761";
const ICE = "CADCFC";
const WHITE = "FFFFFF";
const NAVY_LIGHT = "3A4A96";
const GRAY_TXT = "4A4A4A";
const CARD_BG = "F4F7FE";
const GREEN = "97BC62";
const CODE_BLUE = "6BA6FF";

let pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
const PW = 13.333, PH = 7.5;

const TOTAL_PAGES = 15;

function footer(slide, page, dark) {
  const c1 = dark ? "8090C0" : "8A93B8";
  slide.addText("Setup Environment & Hello World  •  Level 1 – Fundamental", {
    x: 0.5, y: PH - 0.42, w: 8, h: 0.3, fontFace: "Calibri", fontSize: 10,
    color: c1, align: "left", margin: 0
  });
  slide.addText(`${String(page).padStart(2,"0")} / ${TOTAL_PAGES}`, {
    x: PW - 2.5, y: PH - 0.42, w: 2, h: 0.3, fontFace: "Calibri", fontSize: 10,
    color: c1, align: "right", margin: 0
  });
}

function kicker(slide, text, color) {
  slide.addText(text.toUpperCase(), {
    x: 0.8, y: 0.55, w: 10, h: 0.4, fontFace: "Calibri", fontSize: 14,
    color: color || NAVY, bold: true, charSpacing: 2, margin: 0
  });
}

function heading(slide, text, color, size) {
  slide.addText(text, {
    x: 0.8, y: 0.95, w: 11.6, h: 0.8, fontFace: "Cambria", fontSize: size || 30,
    color: color || NAVY, bold: true, margin: 0
  });
}

function codeCard(slide, x, y, w, h, filename, codeLines, opts) {
  opts = opts || {};
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.12, fill: { color: NAVY }, line: { type: "none" },
    shadow: { type: "outer", color: "1E2761", opacity: 0.18, blur: 12, offset: 3, angle: 90 }
  });
  ["E06C75", "E5C07B", "98C379"].forEach((c, i) => {
    slide.addShape(pres.ShapeType.ellipse, {
      x: x + 0.25 + i * 0.3, y: y + 0.2, w: 0.16, h: 0.16, fill: { color: c }, line: { type: "none" }
    });
  });
  if (filename) {
    slide.addText(filename, {
      x: x + w - 3.0, y: y + 0.12, w: 2.7, h: 0.3, fontFace: "Courier New", fontSize: 11,
      color: "8090C0", align: "right", margin: 0
    });
  }
  slide.addText(codeLines, {
    x: x + 0.4, y: y + 0.55, w: w - 0.8, h: h - 0.75, fontFace: "Courier New",
    fontSize: opts.fontSize || 13, color: ICE, margin: 0, valign: "top",
    lineSpacingMultiple: opts.lineSpacing || 1.3
  });
}

// ============ SLIDE 1: TITLE ============
{
  let s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText("</>", {
    x: 8.7, y: 1.0, w: 5.0, h: 5.5, fontFace: "Courier New", fontSize: 260,
    color: NAVY_LIGHT, align: "center", valign: "middle", margin: 0, bold: true
  });
  s.addText("LEVEL 1 — FUNDAMENTAL", {
    x: 0.8, y: 1.4, w: 8, h: 0.5, fontFace: "Calibri", fontSize: 16,
    color: ICE, bold: true, charSpacing: 3, margin: 0
  });
  s.addText("Setup Environment\n& Hello World", {
    x: 0.8, y: 1.9, w: 8.6, h: 2.3, fontFace: "Cambria", fontSize: 44,
    color: WHITE, bold: true, margin: 0, lineSpacingMultiple: 1.05
  });
  s.addText("Panduan lengkap: teori, langkah instalasi, contoh kode, hingga latihan mandiri.", {
    x: 0.8, y: 4.2, w: 7.6, h: 0.9, fontFace: "Calibri", fontSize: 16,
    color: ICE, italic: true, margin: 0
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 5.55, w: 2.3, h: 0.55, rectRadius: 0.28,
    fill: { color: NAVY_LIGHT }, line: { type: "none" }
  });
  s.addText("31 Agustus 2026", {
    x: 0.8, y: 5.55, w: 2.3, h: 0.55, fontFace: "Calibri", fontSize: 13,
    color: WHITE, align: "center", valign: "middle", margin: 0, bold: true
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 3.25, y: 5.55, w: 2.85, h: 0.55, rectRadius: 0.28,
    fill: { type: "none" }, line: { color: ICE, width: 1.25 }
  });
  s.addText("15 Slide Materi", {
    x: 3.25, y: 5.55, w: 2.85, h: 0.55, fontFace: "Calibri", fontSize: 13,
    color: ICE, align: "center", valign: "middle", margin: 0, bold: true
  });
}

// ============ SLIDE 2: TUJUAN PEMBELAJARAN ============
{
  let s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Tujuan Pembelajaran");
  heading(s, "Dari teori ke praktik nyata", NAVY, 32);

  s.addShape(pres.ShapeType.ellipse, {
    x: 0.9, y: 2.3, w: 3.0, h: 3.0, fill: { color: CARD_BG }, line: { color: ICE, width: 1.5 }
  });
  s.addImage({ path: "ic_laptop_navy.png", x: 1.75, y: 3.15, w: 1.3, h: 1.3 });

  s.addShape(pres.ShapeType.roundRect, {
    x: 4.5, y: 2.3, w: 8.0, h: 3.6, rectRadius: 0.12,
    fill: { color: CARD_BG }, line: { type: "none" },
    shadow: { type: "outer", color: "1E2761", opacity: 0.12, blur: 12, offset: 3, angle: 90 }
  });
  s.addText("Mahasiswa mampu menyiapkan environment coding sendiri & menjalankan program dari teori ke praktik nyata.", {
    x: 4.9, y: 2.6, w: 7.2, h: 1.5, fontFace: "Cambria", fontSize: 21,
    color: NAVY, bold: true, margin: 0, valign: "top"
  });
  const subgoals = [
    "Memahami konsep dasar code editor, runtime, struktur file, dan komentar kode",
    "Mampu menginstal VS Code dan Node.js secara mandiri",
    "Mampu membuat, menyimpan, dan menjalankan file .js pertama",
    "Mampu menggunakan console.log() untuk menampilkan output"
  ];
  s.addText(subgoals.map((b, i) => ({
    text: b, options: { bullet: { code: "2022" }, breakLine: i !== subgoals.length - 1, paraSpaceAfter: 8 }
  })), {
    x: 4.95, y: 4.15, w: 7.1, h: 1.6, fontFace: "Calibri", fontSize: 13,
    color: GRAY_TXT, margin: 0, valign: "top"
  });

  footer(s, 2);
}

// ============ SLIDE 3: AGENDA ============
{
  let s = pres.addSlide();
  s.background = { color: NAVY };
  kicker(s, "Agenda", ICE);
  heading(s, "Lima bagian materi hari ini", WHITE, 30);

  const agenda = [
    { n: "01", t: "Materi Teori", d: "Konsep dasar: editor, runtime, struktur file, komentar" },
    { n: "02", t: "Praktik Instalasi", d: "Install VS Code & Node.js langkah demi langkah" },
    { n: "03", t: "Hello World", d: "Membuat & menjalankan file .js pertama" },
    { n: "04", t: "console.log() Detail", d: "Ragam contoh penggunaan output ke konsol" },
    { n: "05", t: "Latihan & Penilaian", d: "Tugas mandiri \u201cPerkenalan Diri\u201d + rubrik" },
  ];
  const y0 = 2.35, rowH = 0.86;
  agenda.forEach((a, i) => {
    const y = y0 + i * rowH;
    s.addText(a.n, {
      x: 0.9, y, w: 1.0, h: rowH - 0.12, fontFace: "Cambria", fontSize: 24, bold: true,
      color: NAVY_LIGHT, valign: "middle", margin: 0
    });
    s.addShape(pres.ShapeType.roundRect, {
      x: 2.1, y, w: 9.6, h: rowH - 0.18, rectRadius: 0.08,
      fill: { color: NAVY_LIGHT }, line: { type: "none" }
    });
    s.addText(a.t, {
      x: 2.4, y, w: 3.2, h: rowH - 0.18, fontFace: "Calibri", fontSize: 15, bold: true,
      color: WHITE, valign: "middle", margin: 0
    });
    s.addText(a.d, {
      x: 5.7, y, w: 5.8, h: rowH - 0.18, fontFace: "Calibri", fontSize: 12.5,
      color: ICE, valign: "middle", margin: 0
    });
  });

  footer(s, 3, true);
}

// ============ SLIDE 4: MATERI TEORI OVERVIEW ============
{
  let s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Materi Teori");
  heading(s, "Empat konsep dasar sebelum praktik", NAVY, 30);

  const items = [
    { icon: "ic_download_navy.png", title: "Instalasi Tools", desc: "Code editor & runtime bahasa pemrograman" },
    { icon: "ic_filecode_navy.png", title: "Struktur File", desc: "Susunan file program & konvensi penamaan" },
    { icon: "ic_cogs_navy.png", title: "Menjalankan Program", desc: "Cara menjalankan / compile program" },
    { icon: "ic_hashtag_navy.png", title: "Komentar Kode", desc: "Menulis komentar untuk dokumentasi kode" },
  ];
  const cardW = 2.85, gap = 0.28, startX = 0.8, y = 2.25, cardH = 3.55;
  items.forEach((it, i) => {
    const x = startX + i * (cardW + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w: cardW, h: cardH, rectRadius: 0.1,
      fill: { color: CARD_BG }, line: { type: "none" },
      shadow: { type: "outer", color: "1E2761", opacity: 0.10, blur: 10, offset: 2, angle: 90 }
    });
    s.addShape(pres.ShapeType.ellipse, {
      x: x + cardW/2 - 0.55, y: y + 0.35, w: 1.1, h: 1.1,
      fill: { color: WHITE }, line: { color: ICE, width: 1.25 }
    });
    s.addImage({ path: it.icon, x: x + cardW/2 - 0.32, y: y + 0.58, w: 0.64, h: 0.64 });
    s.addText(it.title, {
      x: x + 0.15, y: y + 1.65, w: cardW - 0.3, h: 0.55, fontFace: "Calibri", fontSize: 15,
      color: NAVY, bold: true, align: "center", margin: 0, valign: "top"
    });
    s.addText(it.desc, {
      x: x + 0.2, y: y + 2.2, w: cardW - 0.4, h: 1.2, fontFace: "Calibri", fontSize: 11.5,
      color: GRAY_TXT, align: "center", margin: 0, valign: "top"
    });
    s.addText(`0${i+1}`, {
      x: x + 0.15, y: y + 0.12, w: 0.7, h: 0.4, fontFace: "Cambria", fontSize: 13,
      color: ICE, bold: true, margin: 0
    });
  });

  footer(s, 4);
}

// ============ SLIDE 5: DETAIL - EDITOR & RUNTIME ============
{
  let s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Materi Teori 01");
  heading(s, "Code Editor & Runtime — Apa Bedanya?", NAVY, 27);

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 2.15, w: 5.7, h: 4.35, rectRadius: 0.12,
    fill: { color: CARD_BG }, line: { type: "none" },
    shadow: { type: "outer", color: "1E2761", opacity: 0.1, blur: 10, offset: 2, angle: 90 }
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 1.1, y: 2.45, w: 0.85, h: 0.85, fill: { color: WHITE }, line: { color: ICE, width: 1.25 }
  });
  s.addImage({ path: "ic_code_navy.png", x: 1.27, y: 2.62, w: 0.5, h: 0.5 });
  s.addText("Code Editor", {
    x: 2.15, y: 2.5, w: 4.2, h: 0.5, fontFace: "Calibri", fontSize: 18, bold: true,
    color: NAVY, valign: "middle", margin: 0
  });
  s.addText("Aplikasi untuk menulis & mengedit kode program. Dilengkapi fitur seperti syntax highlighting, auto-complete, dan pengelolaan file.", {
    x: 1.1, y: 3.45, w: 5.1, h: 1.1, fontFace: "Calibri", fontSize: 13, color: GRAY_TXT, margin: 0, valign: "top"
  });
  s.addText("Contoh:", { x: 1.1, y: 4.55, w: 2, h: 0.35, fontFace: "Calibri", fontSize: 12, bold: true, color: NAVY, margin: 0 });
  s.addText("Visual Studio Code, Sublime Text, Atom", {
    x: 1.1, y: 4.9, w: 5.1, h: 0.4, fontFace: "Calibri", fontSize: 12.5, italic: true, color: GRAY_TXT, margin: 0
  });
  s.addShape(pres.ShapeType.line, { x: 1.1, y: 5.5, w: 5.1, h: 0, line: { color: ICE, width: 1 } });
  s.addText("Analogi: seperti Microsoft Word — tempat menulis, tapi khusus untuk kode.", {
    x: 1.1, y: 5.65, w: 5.1, h: 0.7, fontFace: "Calibri", fontSize: 11.5, italic: true, color: "6B7AB8", margin: 0
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 6.8, y: 2.15, w: 5.7, h: 4.35, rectRadius: 0.12,
    fill: { color: NAVY }, line: { type: "none" },
    shadow: { type: "outer", color: "1E2761", opacity: 0.15, blur: 10, offset: 2, angle: 90 }
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 7.1, y: 2.45, w: 0.85, h: 0.85, fill: { color: WHITE }, line: { type: "none" }
  });
  s.addImage({ path: "ic_node_navy.png", x: 7.27, y: 2.62, w: 0.5, h: 0.5 });
  s.addText("Runtime", {
    x: 8.15, y: 2.5, w: 4.2, h: 0.5, fontFace: "Calibri", fontSize: 18, bold: true,
    color: WHITE, valign: "middle", margin: 0
  });
  s.addText("Program yang menjalankan / mengeksekusi kode yang sudah ditulis, sehingga instruksi dalam kode benar-benar diproses oleh komputer.", {
    x: 7.1, y: 3.45, w: 5.1, h: 1.1, fontFace: "Calibri", fontSize: 13, color: ICE, margin: 0, valign: "top"
  });
  s.addText("Contoh:", { x: 7.1, y: 4.55, w: 2, h: 0.35, fontFace: "Calibri", fontSize: 12, bold: true, color: WHITE, margin: 0 });
  s.addText("Node.js (JavaScript), Python Interpreter, JVM (Java)", {
    x: 7.1, y: 4.9, w: 5.1, h: 0.4, fontFace: "Calibri", fontSize: 12.5, italic: true, color: ICE, margin: 0
  });
  s.addShape(pres.ShapeType.line, { x: 7.1, y: 5.5, w: 5.1, h: 0, line: { color: NAVY_LIGHT, width: 1 } });
  s.addText("Analogi: seperti mesin pemutar kaset — teks lagu (kode) baru bisa \u201cdidengar\u201d lewat mesinnya (runtime).", {
    x: 7.1, y: 5.65, w: 5.1, h: 0.7, fontFace: "Calibri", fontSize: 11.5, italic: true, color: "AEBBE8", margin: 0
  });

  footer(s, 5);
}

// ============ SLIDE 6: DETAIL - STRUKTUR FILE ============
{
  let s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Materi Teori 02");
  heading(s, "Struktur File Program — Contoh Nyata", NAVY, 27);

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 2.2, w: 5.5, h: 4.3, rectRadius: 0.12,
    fill: { color: NAVY }, line: { type: "none" },
    shadow: { type: "outer", color: "1E2761", opacity: 0.15, blur: 10, offset: 2, angle: 90 }
  });
  s.addText("EXPLORER", {
    x: 1.15, y: 2.45, w: 3, h: 0.3, fontFace: "Calibri", fontSize: 11, bold: true,
    color: "8090C0", charSpacing: 2, margin: 0
  });
  const tree = [
    { icon: "ic_folderopen_navy.png", label: "belajar-coding/", indent: 0, folder: true },
    { icon: "ic_filecode_navy.png", label: "perkenalan.js", indent: 1 },
    { icon: "ic_filecode_navy.png", label: "latihan2.js", indent: 1 },
    { icon: "ic_filecode_navy.png", label: "README.md", indent: 1 },
  ];
  let ty = 2.95;
  tree.forEach((t) => {
    const x = 1.2 + t.indent * 0.55;
    s.addImage({ path: t.icon, x, y: ty, w: 0.32, h: 0.32 });
    s.addText(t.label, {
      x: x + 0.42, y: ty - 0.06, w: 4.0, h: 0.42, fontFace: "Courier New", fontSize: 14,
      color: t.folder ? WHITE : ICE, bold: !!t.folder, valign: "middle", margin: 0
    });
    ty += 0.58;
  });
  s.addShape(pres.ShapeType.line, { x: 1.15, y: ty + 0.15, w: 4.9, h: 0, line: { color: NAVY_LIGHT, width: 1 } });
  s.addText("Satu folder proyek berisi seluruh file yang saling terkait untuk satu program/latihan.", {
    x: 1.15, y: ty + 0.35, w: 4.9, h: 0.9, fontFace: "Calibri", fontSize: 12, italic: true, color: "AEBBE8", margin: 0
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 6.6, y: 2.2, w: 5.9, h: 4.3, rectRadius: 0.12,
    fill: { color: CARD_BG }, line: { type: "none" },
    shadow: { type: "outer", color: "1E2761", opacity: 0.1, blur: 10, offset: 2, angle: 90 }
  });
  s.addText("Aturan Penamaan File", {
    x: 6.95, y: 2.45, w: 5.2, h: 0.45, fontFace: "Calibri", fontSize: 16, bold: true, color: NAVY, margin: 0
  });
  const rules = [
    { t: "Gunakan huruf kecil", d: "perkenalan.js, bukan Perkenalan.js" },
    { t: "Hindari spasi", d: "gunakan tanda hubung/underscore: data-siswa.js" },
    { t: "Sesuaikan ekstensi file", d: ".js untuk JavaScript, .py untuk Python, dst." },
    { t: "Nama harus deskriptif", d: "mencerminkan isi/fungsi file, contoh: hitung-nilai.js" },
  ];
  let ry = 3.05;
  rules.forEach((r) => {
    s.addShape(pres.ShapeType.ellipse, {
      x: 6.95, y: ry + 0.02, w: 0.22, h: 0.22, fill: { color: NAVY }, line: { type: "none" }
    });
    s.addText(r.t, {
      x: 7.3, y: ry - 0.08, w: 5.1, h: 0.35, fontFace: "Calibri", fontSize: 13.5, bold: true, color: NAVY, margin: 0
    });
    s.addText(r.d, {
      x: 7.3, y: ry + 0.24, w: 5.05, h: 0.4, fontFace: "Calibri", fontSize: 12, color: GRAY_TXT, margin: 0
    });
    ry += 0.78;
  });

  footer(s, 6);
}

// ============ SLIDE 7: DETAIL - MENJALANKAN / COMPILE ============
{
  let s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Materi Teori 03");
  heading(s, "Cara Menjalankan Program — Contoh Perintah", NAVY, 27);

  s.addText("JavaScript bersifat interpreted — tidak perlu proses compile terpisah. Kode langsung dibaca & dijalankan baris demi baris oleh runtime Node.js.", {
    x: 0.8, y: 1.9, w: 11.6, h: 0.6, fontFace: "Calibri", fontSize: 14, color: GRAY_TXT, margin: 0
  });

  codeCard(s, 0.8, 2.75, 11.6, 3.7, "Terminal", [
    { text: "$ ", options: { color: GREEN, bold: true } },
    { text: "cd belajar-coding", options: { breakLine: true } },
    { text: "$ ", options: { color: GREEN, bold: true } },
    { text: "node perkenalan.js", options: { breakLine: true, color: WHITE, bold: true } },
    { text: " ", options: { breakLine: true, fontSize: 6 } },
    { text: "Nama : Budi Santoso", options: { breakLine: true, color: GREEN } },
    { text: "NIM  : 2026010123", options: { breakLine: true, color: GREEN } },
    { text: "Hobi : Coding & Membaca", options: { breakLine: false, color: GREEN } },
  ], { fontSize: 16, lineSpacing: 1.5 });

  s.addText([
    { text: "Perintah ", options: {} },
    { text: "node <nama-file>.js", options: { fontFace: "Courier New", bold: true, color: NAVY } },
    { text: "  menjalankan file JavaScript melalui Node.js. Bandingkan dengan bahasa terkompilasi (mis. Java/C++) yang butuh langkah compile lebih dulu sebelum dijalankan.", options: {} }
  ], {
    x: 0.8, y: 6.65, w: 11.6, h: 0.55, fontFace: "Calibri", fontSize: 12.5, color: GRAY_TXT, margin: 0
  });

  footer(s, 7);
}

// ============ SLIDE 8: DETAIL - KOMENTAR KODE ============
{
  let s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Materi Teori 04");
  heading(s, "Komentar Kode — Jenis & Contoh", NAVY, 28);

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 2.2, w: 4.7, h: 4.3, rectRadius: 0.12,
    fill: { color: CARD_BG }, line: { type: "none" },
    shadow: { type: "outer", color: "1E2761", opacity: 0.1, blur: 10, offset: 2, angle: 90 }
  });
  s.addImage({ path: "ic_comment_navy.png", x: 1.1, y: 2.5, w: 0.6, h: 0.6 });
  s.addText("Apa itu Komentar?", {
    x: 1.9, y: 2.5, w: 3.4, h: 0.6, fontFace: "Calibri", fontSize: 16, bold: true, color: NAVY, valign: "middle", margin: 0
  });
  s.addText("Teks dalam kode yang diabaikan oleh runtime — tidak dieksekusi. Digunakan untuk menjelaskan maksud kode atau menonaktifkan kode sementara.", {
    x: 1.1, y: 3.35, w: 4.1, h: 1.1, fontFace: "Calibri", fontSize: 13, color: GRAY_TXT, margin: 0, valign: "top"
  });
  const uses = ["Dokumentasi & penjelasan logika", "Menonaktifkan kode saat debugging", "Menandai TODO / catatan untuk tim"];
  s.addText(uses.map((b,i)=>({text:b, options:{bullet:{code:"2022"}, breakLine: i!==uses.length-1, paraSpaceAfter:8}})), {
    x: 1.15, y: 4.6, w: 4.0, h: 1.5, fontFace: "Calibri", fontSize: 12.5, color: GRAY_TXT, margin: 0, valign: "top"
  });

  codeCard(s, 5.75, 2.2, 6.75, 4.3, "contoh.js", [
    { text: "// Ini komentar satu baris", options: { breakLine: true, color: "6B7AB8" } },
    { text: "console.log(\"Halo!\");", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "/* Ini komentar", options: { breakLine: true, color: "6B7AB8" } },
    { text: "   beberapa baris */", options: { breakLine: true, color: "6B7AB8" } },
    { text: "let nama = \"Budi\"; ", options: { breakLine: false } },
    { text: "// komentar di akhir baris", options: { breakLine: true, color: "6B7AB8" } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "// console.log(\"tidak jalan\");", options: { breakLine: false, color: "6B7AB8", italic: true } },
  ], { fontSize: 14.5, lineSpacing: 1.4 });

  footer(s, 8);
}

// ============ SLIDE 9: PRAKTIK - INSTALL VS CODE ============
{
  let s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Praktik — Langkah 1 dari 3");
  heading(s, "Instalasi Visual Studio Code", NAVY, 28);

  const steps = [
    { t: "Buka situs resmi", d: "Kunjungi code.visualstudio.com melalui browser." },
    { t: "Unduh installer", d: "Klik tombol Download, pilih versi sesuai OS (Windows/macOS/Linux)." },
    { t: "Jalankan installer", d: "Buka file yang terunduh, klik Next hingga selesai, lalu Finish." },
    { t: "Buka VS Code", d: "Jalankan aplikasi, buat folder proyek baru, misal belajar-coding." },
  ];
  const colW = 2.78, gap = 0.24, startX = 0.8, y = 2.35, colH = 3.9;
  steps.forEach((st, i) => {
    const x = startX + i * (colW + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w: colW, h: colH, rectRadius: 0.1, fill: { color: CARD_BG }, line: { type: "none" },
      shadow: { type: "outer", color: "1E2761", opacity: 0.1, blur: 10, offset: 2, angle: 90 }
    });
    s.addShape(pres.ShapeType.ellipse, {
      x: x + 0.25, y: y + 0.25, w: 0.6, h: 0.6, fill: { color: NAVY }, line: { type: "none" }
    });
    s.addText(String(i+1), {
      x: x + 0.25, y: y + 0.25, w: 0.6, h: 0.6, fontFace: "Cambria", fontSize: 20, bold: true,
      color: WHITE, align: "center", valign: "middle", margin: 0
    });
    s.addText(st.t, {
      x: x + 0.22, y: y + 1.05, w: colW - 0.44, h: 0.65, fontFace: "Calibri", fontSize: 14, bold: true,
      color: NAVY, margin: 0, valign: "top"
    });
    s.addText(st.d, {
      x: x + 0.22, y: y + 1.7, w: colW - 0.44, h: 2.0, fontFace: "Calibri", fontSize: 11.5,
      color: GRAY_TXT, margin: 0, valign: "top"
    });
    if (i < steps.length - 1) {
      s.addText("\u2192", { x: x + colW + 0.01, y: y + colH/2 - 0.25, w: 0.22, h: 0.5, fontFace: "Arial",
        fontSize: 20, color: ICE, align: "center", valign: "middle", margin: 0, bold: true });
    }
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 6.4, w: 11.6, h: 0.6, rectRadius: 0.08, fill: { color: "FFF6DD" }, line: { type: "none" }
  });
  s.addImage({ path: "ic_bulb_navy.png", x: 1.0, y: 6.52, w: 0.36, h: 0.36 });
  s.addText("Tips: pilih paket sesuai sistem operasi — Windows: .exe, macOS: .zip, Linux: .deb/.rpm.", {
    x: 1.5, y: 6.4, w: 10.6, h: 0.6, fontFace: "Calibri", fontSize: 12, color: "6B5A00", valign: "middle", margin: 0, italic: true
  });

  footer(s, 9);
}

// ============ SLIDE 10: PRAKTIK - INSTALL NODE.JS ============
{
  let s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Praktik — Langkah 2 dari 3");
  heading(s, "Instalasi Node.js & Verifikasi", NAVY, 28);

  const steps = [
    { t: "Buka nodejs.org", d: "Kunjungi situs resmi Node.js melalui browser." },
    { t: "Unduh versi LTS", d: "Pilih versi LTS (Long Term Support) — lebih stabil untuk pemula." },
    { t: "Install seperti biasa", d: "Jalankan installer, klik Next hingga selesai (Next-Next-Finish)." },
  ];
  const stW = 5.6, y0 = 2.3;
  steps.forEach((st, i) => {
    const y = y0 + i * 1.15;
    s.addShape(pres.ShapeType.ellipse, {
      x: 0.8, y, w: 0.55, h: 0.55, fill: { color: NAVY }, line: { type: "none" }
    });
    s.addText(String(i+1), {
      x: 0.8, y, w: 0.55, h: 0.55, fontFace: "Cambria", fontSize: 18, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(st.t, {
      x: 1.55, y: y - 0.05, w: stW, h: 0.4, fontFace: "Calibri", fontSize: 14.5, bold: true, color: NAVY, margin: 0
    });
    s.addText(st.d, {
      x: 1.55, y: y + 0.34, w: stW, h: 0.65, fontFace: "Calibri", fontSize: 12, color: GRAY_TXT, margin: 0, valign: "top"
    });
  });

  codeCard(s, 7.0, 2.3, 5.5, 4.2, "Terminal", [
    { text: "$ ", options: { color: GREEN, bold: true } },
    { text: "node -v", options: { breakLine: true, bold: true } },
    { text: "v20.14.0", options: { breakLine: true, color: GREEN } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "$ ", options: { color: GREEN, bold: true } },
    { text: "npm -v", options: { breakLine: true, bold: true } },
    { text: "10.7.0", options: { breakLine: false, color: GREEN } },
  ], { fontSize: 15, lineSpacing: 1.45 });
  s.addText("Langkah 4 — Verifikasi Instalasi", {
    x: 7.0, y: 1.95, w: 5.5, h: 0.35, fontFace: "Calibri", fontSize: 13, bold: true, color: NAVY, margin: 0
  });
  s.addText("Jika muncul nomor versi seperti di atas, instalasi berhasil.", {
    x: 7.0, y: 6.6, w: 5.5, h: 0.5, fontFace: "Calibri", fontSize: 12, italic: true, color: GRAY_TXT, margin: 0
  });

  footer(s, 10);
}

// ============ SLIDE 11: PRAKTIK - HELLO WORLD ============
{
  let s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Praktik — Langkah 3 dari 3");
  heading(s, "Membuat & Menjalankan File Pertama", NAVY, 28);

  codeCard(s, 0.8, 2.2, 5.7, 4.35, "hello.js", [
    { text: "// Program Hello World", options: { breakLine: true, color: "6B7AB8" } },
    { text: "console.log(", options: {} },
    { text: "\"Hello, World!\"", options: { color: GREEN } },
    { text: ");", options: { breakLine: true } },
  ], { fontSize: 17, lineSpacing: 1.6 });

  codeCard(s, 6.8, 2.2, 5.7, 4.35, "Terminal", [
    { text: "$ ", options: { color: GREEN, bold: true } },
    { text: "cd belajar-coding", options: { breakLine: true } },
    { text: "$ ", options: { color: GREEN, bold: true } },
    { text: "node hello.js", options: { breakLine: true, bold: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "Hello, World!", options: { breakLine: false, color: GREEN, bold: true } },
  ], { fontSize: 16, lineSpacing: 1.55 });

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 6.75, w: 11.7, h: 0.5, rectRadius: 0.08, fill: { color: CARD_BG }, line: { type: "none" }
  });
  s.addText("Langkah: (1) buat file hello.js di VS Code  →  (2) simpan (Ctrl+S)  →  (3) buka terminal di folder yang sama  →  (4) jalankan node hello.js", {
    x: 1.0, y: 6.75, w: 11.3, h: 0.5, fontFace: "Calibri", fontSize: 11.5, color: NAVY, valign: "middle", margin: 0
  });

  footer(s, 11);
}

// ============ SLIDE 12: CONSOLE.LOG() DETAIL ============
{
  let s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Praktik — console.log()");
  heading(s, "console.log() — Ragam Contoh Penggunaan", NAVY, 27);

  codeCard(s, 0.8, 2.15, 6.5, 4.5, "contoh-log.js", [
    { text: "console.log(", options: {} }, { text: "\"Halo dunia\"", options: { color: GREEN } }, { text: ");", options: { breakLine: true } },
    { text: "console.log(", options: {} }, { text: "2026", options: { color: CODE_BLUE } }, { text: ");", options: { breakLine: true } },
    { text: "let nama = ", options: {} }, { text: "\"Sari\"", options: { color: GREEN } }, { text: ";", options: { breakLine: true } },
    { text: "console.log(nama);", options: { breakLine: true } },
    { text: "console.log(", options: {} }, { text: "\"Nama:\"", options: { color: GREEN } }, { text: ", nama);", options: { breakLine: true } },
    { text: "console.log(`Halo, ${nama}!`);", options: { breakLine: false } },
  ], { fontSize: 14, lineSpacing: 1.45 });

  s.addShape(pres.ShapeType.roundRect, {
    x: 7.6, y: 2.15, w: 4.9, h: 4.5, rectRadius: 0.12, fill: { color: CARD_BG }, line: { type: "none" },
    shadow: { type: "outer", color: "1E2761", opacity: 0.1, blur: 10, offset: 2, angle: 90 }
  });
  s.addText("Output di Terminal", {
    x: 7.9, y: 2.4, w: 4.3, h: 0.4, fontFace: "Calibri", fontSize: 14, bold: true, color: NAVY, margin: 0
  });
  s.addText([
    { text: "Halo dunia", options: { breakLine: true } },
    { text: "2026", options: { breakLine: true } },
    { text: "Sari", options: { breakLine: true } },
    { text: "Nama: Sari", options: { breakLine: true } },
    { text: "Halo, Sari!", options: { breakLine: false } },
  ], {
    x: 7.9, y: 2.95, w: 4.3, h: 2.6, fontFace: "Courier New", fontSize: 14, color: NAVY,
    margin: 0, valign: "top", lineSpacingMultiple: 1.9
  });
  s.addShape(pres.ShapeType.line, { x: 7.9, y: 5.75, w: 4.3, h: 0, line: { color: ICE, width: 1 } });
  s.addText("Bisa mencetak teks, angka, variabel, gabungan teks+variabel, hingga template literal (`...`).", {
    x: 7.9, y: 5.9, w: 4.3, h: 0.7, fontFace: "Calibri", fontSize: 11.5, italic: true, color: GRAY_TXT, margin: 0
  });

  footer(s, 12);
}

// ============ SLIDE 13: TOOLS RECAP ============
{
  let s = pres.addSlide();
  s.background = { color: NAVY };
  kicker(s, "Tools yang Digunakan", ICE);
  heading(s, "Perlengkapan wajib di Level 1", WHITE, 30);

  const tools = [
    { icon: "ic_code_navy.png", label: "VS Code", desc: "Code editor utama untuk menulis program" },
    { icon: "ic_node_navy.png", label: "Node.js", desc: "Runtime untuk menjalankan JavaScript" },
    { icon: "ic_terminal_navy.png", label: "Terminal / CMD", desc: "Antarmuka baris perintah untuk eksekusi program" },
  ];
  const cardW = 3.5, gap = 0.55, startX = (PW - (cardW*3 + gap*2))/2, y = 2.5, cardH = 3.4;
  tools.forEach((t, i) => {
    const x = startX + i * (cardW + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w: cardW, h: cardH, rectRadius: 0.12, fill: { color: NAVY_LIGHT }, line: { type: "none" }
    });
    s.addShape(pres.ShapeType.ellipse, {
      x: x + cardW/2 - 0.6, y: y + 0.4, w: 1.2, h: 1.2, fill: { color: WHITE }, line: { type: "none" }
    });
    s.addImage({ path: t.icon, x: x + cardW/2 - 0.35, y: y + 0.65, w: 0.7, h: 0.7 });
    s.addText(t.label, {
      x: x + 0.2, y: y + 1.8, w: cardW - 0.4, h: 0.5, fontFace: "Calibri", fontSize: 18, bold: true,
      color: WHITE, align: "center", margin: 0
    });
    s.addText(t.desc, {
      x: x + 0.3, y: y + 2.35, w: cardW - 0.6, h: 0.9, fontFace: "Calibri", fontSize: 12.5,
      color: ICE, align: "center", margin: 0, valign: "top"
    });
  });

  footer(s, 13, true);
}

// ============ SLIDE 14: LATIHAN MANDIRI DETAIL ============
{
  let s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Latihan Mandiri");
  heading(s, "Program \u201cPerkenalan Diri\u201d — Contoh Lengkap", NAVY, 26);

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 2.05, w: 5.6, h: 2.0, rectRadius: 0.1, fill: { color: CARD_BG }, line: { type: "none" }
  });
  s.addImage({ path: "ic_clipboard_navy.png", x: 1.05, y: 2.25, w: 0.45, h: 0.45 });
  s.addText("Instruksi", {
    x: 1.65, y: 2.25, w: 4, h: 0.45, fontFace: "Calibri", fontSize: 14, bold: true, color: NAVY, valign: "middle", margin: 0
  });
  const bullets = ["Gunakan console.log()", "Cetak Nama, NIM, dan Hobi", "Jalankan lewat: node <file>.js"];
  s.addText(bullets.map((b,i)=>({text:b, options:{bullet:{code:"2022"}, breakLine:i!==bullets.length-1, paraSpaceAfter:6}})), {
    x: 1.1, y: 2.8, w: 5.1, h: 1.2, fontFace: "Calibri", fontSize: 12, color: GRAY_TXT, margin: 0, valign: "top"
  });

  s.addText("Rubrik Penilaian", {
    x: 0.8, y: 4.2, w: 5, h: 0.4, fontFace: "Calibri", fontSize: 14, bold: true, color: NAVY, margin: 0
  });
  const rubric = [
    ["Kriteria", "Bobot"],
    ["Program berjalan tanpa error", "40%"],
    ["Menampilkan Nama, NIM, Hobi", "40%"],
    ["Kerapian & komentar kode", "20%"],
  ];
  s.addTable(rubric, {
    x: 0.8, y: 4.65, w: 5.6, colW: [3.9, 1.7],
    fontFace: "Calibri", fontSize: 11.5, color: GRAY_TXT, border: { type: "solid", color: "DCE3F5", pt: 1 },
    fill: { color: WHITE }, autoPage: false,
    rowH: 0.42,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.8, y: 4.65, w: 5.6, h: 0.42, fill: { color: NAVY }, line: { type: "none" } });
  s.addText("Kriteria", { x: 0.95, y: 4.65, w: 3.6, h: 0.42, fontFace: "Calibri", fontSize: 11.5, bold: true, color: WHITE, valign: "middle", margin: 0 });
  s.addText("Bobot", { x: 4.7, y: 4.65, w: 1.5, h: 0.42, fontFace: "Calibri", fontSize: 11.5, bold: true, color: WHITE, valign: "middle", margin: 0 });

  codeCard(s, 6.75, 2.05, 5.75, 3.05, "perkenalan.js", [
    { text: "// Program Perkenalan Diri", options: { breakLine: true, color: "6B7AB8" } },
    { text: "console.log(", options: {} }, { text: "\"Nama : Budi Santoso\"", options: { color: GREEN } }, { text: ");", options: { breakLine: true } },
    { text: "console.log(", options: {} }, { text: "\"NIM  : 2026010123\"", options: { color: GREEN } }, { text: ");", options: { breakLine: true } },
    { text: "console.log(", options: {} }, { text: "\"Hobi : Coding & Membaca\"", options: { color: GREEN } }, { text: ");", options: { breakLine: false } },
  ], { fontSize: 13.5, lineSpacing: 1.45 });

  s.addShape(pres.ShapeType.roundRect, {
    x: 6.75, y: 5.3, w: 5.75, h: 1.95, rectRadius: 0.1, fill: { color: CARD_BG }, line: { type: "none" }
  });
  s.addText("Output yang Diharapkan:", {
    x: 7.0, y: 5.48, w: 5.2, h: 0.35, fontFace: "Calibri", fontSize: 12, bold: true, color: NAVY, margin: 0
  });
  s.addText("Nama : Budi Santoso\nNIM  : 2026010123\nHobi : Coding & Membaca", {
    x: 7.0, y: 5.85, w: 5.2, h: 1.3, fontFace: "Courier New", fontSize: 13, color: "3A7D2E", margin: 0, valign: "top", lineSpacingMultiple: 1.35
  });

  footer(s, 14);
}

// ============ SLIDE 15: CAPAIAN / CLOSING ============
{
  let s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.ellipse, {
    x: PW/2 - 0.75, y: 0.85, w: 1.5, h: 1.5, fill: { color: NAVY_LIGHT }, line: { type: "none" }
  });
  s.addImage({ path: "ic_check_white.png", x: PW/2 - 0.45, y: 1.15, w: 0.9, h: 0.9 });

  s.addText("CAPAIAN PEMBELAJARAN", {
    x: 1.5, y: 2.65, w: 10.3, h: 0.4, fontFace: "Calibri", fontSize: 14,
    color: ICE, bold: true, charSpacing: 2, align: "center", margin: 0
  });
  s.addText("Bisa Setup Environment Sendiri\n& Menjalankan Program", {
    x: 1.0, y: 3.1, w: 11.3, h: 1.7, fontFace: "Cambria", fontSize: 32,
    color: WHITE, bold: true, align: "center", margin: 0, lineSpacingMultiple: 1.1
  });
  s.addText("Dari teori instalasi tools hingga praktik menulis & menjalankan kode pertama — mahasiswa siap melanjutkan ke Level berikutnya.", {
    x: 2.15, y: 4.9, w: 9.0, h: 0.8, fontFace: "Calibri", fontSize: 15,
    color: ICE, italic: true, align: "center", margin: 0
  });

  const recap = ["Editor & Runtime", "Struktur File", "Menjalankan Program", "Komentar Kode", "Hello World"];
  const chipW = 2.15, chipGap = 0.18, totalW = recap.length*chipW + (recap.length-1)*chipGap;
  let cx = (PW - totalW)/2;
  recap.forEach((r) => {
    s.addShape(pres.ShapeType.roundRect, {
      x: cx, y: 5.85, w: chipW, h: 0.5, rectRadius: 0.25, fill: { color: NAVY_LIGHT }, line: { type: "none" }
    });
    s.addText(r, {
      x: cx, y: 5.85, w: chipW, h: 0.5, fontFace: "Calibri", fontSize: 10.5, color: WHITE,
      align: "center", valign: "middle", margin: 0
    });
    cx += chipW + chipGap;
  });

  footer(s, 15, true);
}

pres.writeFile({ fileName: "output2.pptx" }).then(() => console.log("done"));
SCRIPT_EOF
echo "written"
