const pptxgen = require("pptxgenjs");

const NAVY = "21295C";
const DEEPBLUE = "065A82";
const TEAL = "1C7293";
const WHITE = "FFFFFF";
const LIGHTBG = "F4F7F9";
const TEXT = "24313F";
const MUTED = "5B6B79";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

const FONT_HEAD = "Cambria";
const FONT_BODY = "Calibri";

function circleNum(slide, num, x, y, d = 0.85) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: DEEPBLUE }, line: { type: "none" } });
  slide.addText(String(num), {
    x, y, w: d, h: d, align: "center", valign: "middle",
    fontFace: FONT_HEAD, fontSize: 26, bold: true, color: WHITE,
  });
}

function contentHeader(slide, num, title, subtitle) {
  slide.background = { color: WHITE };
  circleNum(slide, num, 0.55, 0.5, 0.9);
  slide.addText(title, {
    x: 1.65, y: 0.4, w: 11.0, h: 0.6, fontFace: FONT_HEAD, fontSize: 27, bold: true, color: NAVY,
  });
  slide.addText(subtitle, {
    x: 1.65, y: 0.93, w: 11.0, h: 0.4, fontFace: FONT_BODY, fontSize: 13, italic: true, color: TEAL,
  });
}

function colBlock(slide, x, y, w, h, header, items) {
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: LIGHTBG }, line: { type: "none" },
  });
  slide.addText(header, {
    x: x + 0.25, y: y + 0.18, w: w - 0.5, h: 0.4,
    fontFace: FONT_BODY, fontSize: 15, bold: true, color: DEEPBLUE,
  });
  const bulletItems = items.map((it, i) => ({
    text: it,
    options: {
      bullet: { code: "25AA" },
      breakLine: i !== items.length - 1,
      paraSpaceAfter: 8,
      color: TEXT,
    },
  }));
  slide.addText(bulletItems, {
    x: x + 0.25, y: y + 0.62, w: w - 0.5, h: h - 0.8,
    fontFace: FONT_BODY, fontSize: 12.5, valign: "top", lineSpacingMultiple: 1.05,
  });
}

/* ============ SLIDE 1: TITLE ============ */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape("ellipse", { x: 10.3, y: -1.8, w: 5, h: 5, fill: { color: DEEPBLUE, transparency: 60 }, line: { type: "none" } });
  s.addShape("ellipse", { x: -1.5, y: 5.2, w: 4, h: 4, fill: { color: TEAL, transparency: 65 }, line: { type: "none" } });

  s.addText("SECUREDOOR", {
    x: 0.9, y: 1.9, w: 11.5, h: 1.0, fontFace: FONT_HEAD, fontSize: 46, bold: true, color: WHITE, charSpacing: 2,
  });
  s.addText("Kontribusi Individu dalam Keamanan Perangkat Lunak Proyek IoT", {
    x: 0.9, y: 2.85, w: 11.5, h: 0.6, fontFace: FONT_BODY, fontSize: 17, color: "CADCFC",
  });
  s.addShape("line", { x: 0.95, y: 3.45, w: 2.2, h: 0, line: { color: TEAL, width: 2.5 } });
  s.addText("Presentasi Ujian Akhir Semester (AAS) Praktikum", {
    x: 0.9, y: 3.65, w: 11.0, h: 0.4, fontFace: FONT_BODY, fontSize: 13, italic: true, color: "9FB8DC",
  });

  s.addShape("roundRect", { x: 0.9, y: 4.5, w: 6.4, h: 1.85, rectRadius: 0.1, fill: { color: DEEPBLUE, transparency: 25 }, line: { type: "none" } });
  s.addText("Tomingse Lingga", { x: 1.2, y: 4.68, w: 6, h: 0.5, fontFace: FONT_HEAD, fontSize: 22, bold: true, color: WHITE });
  s.addText("NIM: [isi NIM]   |   Kelas: [isi kelas]", { x: 1.2, y: 5.2, w: 6, h: 0.4, fontFace: FONT_BODY, fontSize: 13, color: "CADCFC" });
  s.addText("Kontributor: Luaran 1, 2, 3, 4, 5, dan 6", { x: 1.2, y: 5.6, w: 6, h: 0.4, fontFace: FONT_BODY, fontSize: 12, italic: true, color: "9FB8DC" });

  s.addNotes("Slide pembuka video individu. Sebutkan nama, NIM, kelas, lalu sampaikan bahwa presentasi ini berfokus pada bagian yang menjadi kontribusi pribadi pada proyek SecureDoor di keenam luaran.");
}

/* ============ SLIDE 2: PETA KONTRIBUSI ============ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("Peta Kontribusi Saya", {
    x: 0.7, y: 0.5, w: 11.5, h: 0.7, fontFace: FONT_HEAD, fontSize: 30, bold: true, color: NAVY,
  });
  s.addText("Ringkasan bagian yang saya kerjakan di setiap luaran", {
    x: 0.7, y: 1.15, w: 11.5, h: 0.4, fontFace: FONT_BODY, fontSize: 14, italic: true, color: TEAL,
  });

  const map = [
    ["1", "Persyaratan Keamanan", "Akuntabilitas, Anti-Piracy, Klasifikasi CIA, Model Serangan"],
    ["2", "Desain Keamanan", "Bahasa Pemrograman & Desain Tipe Data"],
    ["3", "Pengembangan Keamanan", "Code Review & Code Signing"],
    ["4", "Pengujian Keamanan", "Pengujian dengan Data Sintetis (Dummy)"],
    ["5", "Penerapan Keamanan", "Persetujuan Rilis & Kebijakan Risiko"],
    ["6", "Pemeliharaan Keamanan", "Kontrol Operasional, Monitoring & Audit"],
  ];
  const cols = 3, rows = 2;
  const cardW = 3.85, cardH = 2.35, gapX = 0.25, gapY = 0.3;
  const startX = 0.7, startY = 1.9;
  map.forEach((item, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);
    s.addShape("roundRect", { x, y, w: cardW, h: cardH, rectRadius: 0.09, fill: { color: LIGHTBG }, line: { type: "none" } });
    circleNum(s, item[0], x + 0.22, y + 0.22, 0.55);
    s.addText(item[1], { x: x + 0.95, y: y + 0.2, w: cardW - 1.15, h: 0.6, fontFace: FONT_BODY, fontSize: 13.5, bold: true, color: NAVY, valign: "top" });
    s.addText(item[2], { x: x + 0.25, y: y + 1.0, w: cardW - 0.5, h: 1.2, fontFace: FONT_BODY, fontSize: 11, color: MUTED, valign: "top", lineSpacingMultiple: 1.05 });
  });
  s.addNotes("Tampilkan peta ini sebagai gambaran cepat: kontribusi tersebar di semua luaran, bukan hanya satu bagian saja.");
}

/* ============ SLIDE 3: LUARAN 1 & 2 ============ */
{
  const s = pres.addSlide();
  contentHeader(s, "1-2", "Persyaratan & Desain Keamanan", "Kontribusi pada Luaran 1 dan Luaran 2");
  colBlock(s, 0.55, 1.55, 6.0, 5.4, "Luaran 1 — Persyaratan Keamanan", [
    "Persyaratan Akuntabilitas — audit trail seluruh aktivitas akses pintu (berhasil/gagal)",
    "Persyaratan Anti-Piracy — autentikasi Google OAuth + registrasi & validasi perangkat",
    "Pelabelan data sesuai prinsip CIA (Confidentiality, Integrity, Availability)",
    "Model serangan — identifikasi ancaman (akses tidak sah, manipulasi status, brute force PIN) & sumber ancaman",
  ]);
  colBlock(s, 6.75, 1.55, 6.0, 5.4, "Luaran 2 — Desain Keamanan", [
    "Penentuan bahasa pemrograman per komponen: C/C++ (IoT), Node.js (backend), Kotlin (mobile)",
    "Desain jenis & format data: tipe primitif, User-Defined Type (ENUM), check constraint",
    "Penanganan ketidakcocokan tipe data melalui explicit casting & exception handling",
  ]);
  s.addNotes("Jelaskan bagaimana persyaratan akuntabilitas & anti-piracy yang dirumuskan di Luaran 1 diteruskan menjadi keputusan desain teknis di Luaran 2, seperti pemilihan bahasa dan struktur tipe data.");
}

/* ============ SLIDE 4: LUARAN 3 & 4 ============ */
{
  const s = pres.addSlide();
  contentHeader(s, "3-4", "Pengembangan & Pengujian Keamanan", "Kontribusi pada Luaran 3 dan Luaran 4");
  colBlock(s, 0.55, 1.55, 6.0, 5.4, "Luaran 3 — Pengembangan Keamanan", [
    "Code Review — peninjauan kode secara sistematis & manual oleh sesama pengembang",
    "Memastikan kode sesuai standar dari sisi fungsionalitas, struktur, dan keamanan",
    "Code Signing — sudah diterapkan pada aplikasi mobile",
    "Pada perangkat IoT (ESP32) masih perlu dilengkapi secure boot & firmware signing",
  ]);
  colBlock(s, 6.75, 1.55, 6.0, 5.4, "Luaran 4 — Pengujian Keamanan", [
    "Pengujian menggunakan data dummy/sintetis, tidak terhubung data pengguna asli",
    "Contoh data dummy: akun simulasi, log akses palsu, status pintu percobaan",
    "Digunakan untuk uji login, akses IoT–server, notifikasi, dan performa backend",
    "Menjaga keamanan & privasi data asli selama proses pengujian",
  ]);
  s.addNotes("Sampaikan peran code review dan code signing sebagai kontrol kualitas kode, lalu jelaskan penggunaan data dummy pada pengujian sebagai praktik yang aman dan tidak berisiko terhadap data pengguna nyata.");
}

/* ============ SLIDE 5: LUARAN 5 & 6 ============ */
{
  const s = pres.addSlide();
  contentHeader(s, "5-6", "Penerapan & Pemeliharaan Keamanan", "Kontribusi pada Luaran 5 dan Luaran 6");
  colBlock(s, 0.55, 1.55, 6.0, 5.4, "Luaran 5 — Penerapan Keamanan", [
    "Persetujuan rilis dilakukan melalui koordinasi tim setelah modul lolos pengujian",
    "Masih bersifat informal (diskusi tim); dokumentasi resmi direncanakan untuk tahap produksi",
    "Kebijakan risiko: pendekatan pragmatis — risiko kritis (OAuth, JWT, HTTPS) diprioritaskan sebelum rilis",
  ]);
  colBlock(s, 6.75, 1.55, 6.0, 5.4, "Luaran 6 — Pemeliharaan Keamanan", [
    "Menerapkan 5 kontrol operasional: detektif, preventif, pencegahan, korektif, kompensasi",
    "Pemantauan berkelanjutan melalui AnalyticsScreen (total akses, akses ditolak, tren)",
    "Audit pemantauan melalui histori log akses & inferensi AI berdasarkan perangkat & rentang waktu",
  ]);
  s.addNotes("Tekankan bahwa kontribusi di Luaran 6 relatif besar: lima jenis kontrol keamanan operasional, dashboard monitoring, hingga fitur audit histori — ini bagian paling teknis dari kontribusi pribadi.");
}

/* ============ SLIDE 6: PENUTUP ============ */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape("ellipse", { x: -1.6, y: -1.6, w: 4.5, h: 4.5, fill: { color: DEEPBLUE, transparency: 60 }, line: { type: "none" } });
  s.addText("Penutup", {
    x: 0.9, y: 1.0, w: 11.0, h: 0.8, fontFace: FONT_HEAD, fontSize: 34, bold: true, color: WHITE,
  });
  const closing = [
    "Kontribusi saya mencakup keenam luaran: dari akuntabilitas & anti-piracy, desain bahasa/tipe data, code review & signing, pengujian data sintetis, persetujuan rilis & kebijakan risiko, hingga kontrol operasional & audit pemantauan",
    "Beberapa bagian (dokumentasi rilis resmi, code signing IoT) masih menjadi catatan pengembangan lanjutan",
  ].map((t, i, arr) => ({ text: t, options: { bullet: { code: "25AA" }, breakLine: i !== arr.length - 1, paraSpaceAfter: 14, color: "CADCFC" } }));
  s.addText(closing, {
    x: 0.95, y: 2.1, w: 10.8, h: 2.3, fontFace: FONT_BODY, fontSize: 15, valign: "top", lineSpacingMultiple: 1.15,
  });
  s.addShape("line", { x: 0.95, y: 4.9, w: 2.2, h: 0, line: { color: TEAL, width: 2.5 } });
  s.addText("Terima Kasih", {
    x: 0.9, y: 5.1, w: 8, h: 0.6, fontFace: FONT_HEAD, fontSize: 22, bold: true, color: WHITE,
  });
  s.addNotes("Tutup dengan ringkasan kontribusi pribadi dan ucapan terima kasih.");
}

pres.writeFile({ fileName: "./Tomingse_Lingga_Kontribusi_Individu.pptx" }).then(() => {
  console.log("done");
});