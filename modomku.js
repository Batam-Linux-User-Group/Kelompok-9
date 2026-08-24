const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle,
  PageBreak, TableOfContents, VerticalAlign, Footer, PageNumber, NumberFormat,
} = require("docx");

const FONT = "Times New Roman";
const SZ = 24; // 12pt
const BLACK = "000000";
// Margin 3 cm semua sisi. 1 cm = 566.9 twip -> 3 cm = 1701 twip
const M = 1701;
const PAGE = { size: { width: 11906, height: 16838 }, margin: { top: M, bottom: M, left: M, right: M } };
const FIRSTLINE = 720; // 1.27 cm

function footer(format, start) {
  return {
    default: new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20, color: BLACK })],
        }),
      ],
    }),
  };
}

// Paragraf isi: rata kanan-kiri, indentasi awal 1,27 cm, spasi 1,5, hitam
function body(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 360, after: 200 },
    indent: { firstLine: opts.noIndent ? 0 : FIRSTLINE },
    children: [new TextRun({ text, font: FONT, size: SZ, color: BLACK, italics: opts.italics || false, bold: opts.bold || false })],
  });
}

// Instruksi dalam kurung siku (petunjuk pengisian, italic, tidak dihitung sebagai isi jadi)
function instr(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, after: 200 },
    indent: { left: 200 },
    children: [new TextRun({ text: `[${text}]`, font: FONT, size: SZ, color: BLACK, italics: true })],
  });
}

// Poin kerangka / bahan mentah (bullet manual pakai tanda hubung, TANPA indentasi awal paragraf)
function bullet(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, after: 120 },
    indent: { left: 420, hanging: 220 },
    children: [new TextRun({ text: `-  ${text}`, font: FONT, size: SZ, color: BLACK })],
  });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 240 },
    pageBreakBefore: false,
    children: [new TextRun({ text, font: FONT, size: 28, bold: true, color: BLACK, allCaps: true })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    alignment: AlignmentType.LEFT,
    spacing: { before: 240, after: 160 },
    children: [new TextRun({ text, font: FONT, size: SZ, bold: true, color: BLACK })],
  });
}

function refEntry(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, after: 200 },
    indent: { left: 720, hanging: 720 },
    children: [new TextRun({ text, font: FONT, size: SZ, color: BLACK })],
  });
}

function cell(text, opts = {}) {
  const plainBorder = { style: BorderStyle.SINGLE, size: 4, color: BLACK };
  return new TableCell({
    width: { size: opts.width || 2000, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    borders: { top: plainBorder, bottom: plainBorder, left: plainBorder, right: plainBorder },
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { line: 240, after: 0 },
      children: [new TextRun({ text, font: FONT, size: 20, bold: opts.header || false, color: BLACK })],
    })],
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function tableCaption(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 200, after: 100, line: 240 },
    children: [new TextRun({ text, font: FONT, size: 22, bold: true, color: BLACK })],
  });
}

function tableSource(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 60, after: 200, line: 240 },
    children: [new TextRun({ text, font: FONT, size: 20, color: BLACK })],
  });
}

// ---------- Tabel 1: Ringkasan Literatur (bahan mentah hasil penelusuran, sudah diverifikasi bisa ditemukan) ----------
const litRows = [
  ["Kurniati, Komala, & Zulkarnaini (2021)", "Jurnal ilmiah - mekanisme", "Beban pencemar KJA: TN sekitar 84,20% dan TP sekitar 91,83% dari total beban di Danau Maninjau; pupuk pertanian (Urea, SP36, Phonska) menambah beban fosfor.", ""],
  ["Sunaryani (2023)", "Jurnal ilmiah - kasus", "Status mutu air Danau Maninjau 2011-2020 menurun ke kelas cemar berat; status trofik eutrofik-hipereutrofik di lokasi KJA.", ""],
  ["Sulastri, Hartoto, & Yuniarti (2012)", "Jurnal ilmiah - kasus", "Danau Maninjau: danau vulkanik dalam, tekanan ekologis dari perikanan KJA terhadap sumber daya ikan.", ""],
  ["Ibrahim, Sudarso, Imroatushshoolikhah, Toruan, & Sari (2021)", "Jurnal ilmiah - kasus", "Kualitas air sungai inlet Danau Maninjau dinilai lewat makrozoobentos; indikasi gangguan kualitas air.", ""],
  ["Kurniadi, Akbar, & Desmaiani (2025)", "Jurnal ilmiah - kasus/relevansi", "KJA sebagai sumber pencemar nutrien di beberapa danau vulkanik Indonesia (Batur, Maninjau); rekomendasi moratorium KJA baru.", ""],
  ["Irianto & Triweko (2019)", "Buku - mekanisme", "Mekanisme eutrofikasi waduk/danau: dekomposisi bahan organik menurunkan oksigen terlarut; pemodelan dan pengendalian.", ""],
  ["Kementerian Negara LH RI (2009)", "Peraturan resmi", "Permen LH No. 28/2009: metode penetapan daya tampung beban pencemaran air danau/waduk dan klasifikasi status trofik.", ""],
  ["ANTARA News (2025, 20 Desember)", "Berita/laporan resmi instansi - kasus", "DKPP Agam: kematian ikan KJA 1.428,73 ton akibat upwelling dipicu cuaca ekstrem; kerugian sekitar Rp32,86 miliar.", ""],
  ["Radar Sumbar (2026, 2 Februari)", "Berita/laporan resmi instansi - kasus", "Rincian sebaran kematian ikan per nagari sepanjang 2025; kejadian susulan Januari 2026 sekitar 5 ton, mekanisme upwelling sama.", ""],
];
const tableHeaderCells = ["Penulis dan tahun", "Jenis sumber", "Temuan atau penjelasan pokok", "Dipakai pada subbab (isi sendiri)"];
const colWidths = [2200, 1800, 4200, 1600];
const litTable = new Table({
  width: { size: 9800, type: WidthType.DXA },
  columnWidths: colWidths,
  rows: [
    new TableRow({ tableHeader: true, children: tableHeaderCells.map((t, i) => cell(t, { header: true, width: colWidths[i], align: AlignmentType.CENTER })) }),
    ...litRows.map(row => new TableRow({ children: row.map((t, i) => cell(t, { width: colWidths[i] })) })),
  ],
});

// ---------- Daftar Pustaka (bahan mentah, WAJIB diverifikasi sendiri lewat Google Scholar/DOI sebelum final) ----------
const pustaka = [
  "ANTARA News. (2025, 20 Desember). Sekitar 1.428,73 ton ikan mati dampak upwelling air Danau Maninjau. ANTARA News. https://www.antaranews.com/berita/5314777/sekitar-142873-ton-ikan-mati-dampak-upwelling-air-danau-maninjau",
  "Ibrahim, A., Sudarso, J., Imroatushshoolikhah, I., Toruan, R. L., & Sari, L. (2021). Penggunaan makrozoobentos dalam penilaian kualitas perairan sungai inlet Danau Maninjau, Sumatera Barat. Jurnal Ilmu Lingkungan, 19(3), 649-660.",
  "Irianto, E. W., & Triweko, R. W. (2019). Eutrofikasi waduk dan danau: Permasalahan, pemodelan, dan upaya pengendalian. ITB Press.",
  "Kementerian Negara Lingkungan Hidup Republik Indonesia. (2009). Peraturan Menteri Negara Lingkungan Hidup Nomor 28 Tahun 2009 tentang daya tampung beban pencemaran air danau dan/atau waduk.",
  "Kurniadi, Y. A., Akbar, A. A., & Desmaiani, H. (2025). Valuasi kerugian pada danau terdampak eutrofikasi: Studi dampak terhadap perikanan dan kualitas lingkungan. Journal of Environmental Policy and Technology, 2(2), 28-41. https://journal.ummat.ac.id/index.php/jeptec/article/view/35713",
  "Kurniati, R. I., Komala, P. S., & Zulkarnaini. (2021). Analisis beban pencemar total nitrogen dan total fosfat akibat aktivitas antropogenik di Danau Maninjau. Jurnal Ilmu Lingkungan, 19(2), 355-364.",
  "Radar Sumbar. (2026, 2 Februari). 1.428 ton ikan di Danau Maninjau mati sepanjang 2025, kerugian capai Rp35,7 miliar. Radar Sumbar. https://www.radarsumbar.com/baca/197459/1-428-ton-ikan-di-danau-maninjau-mati-sepanjang-2025-kerugian-capai-rp357-miliar/",
  "Sulastri, D., Hartoto, D. I., & Yuniarti, I. (2012). Environmental condition, fish resources and management of Maninjau Lake of West Sumatera. Indonesian Fisheries Research Journal, 18(1), 1-12.",
  "Sunaryani, A. (2023). Penentuan status mutu air dan status trofik di perairan Danau Maninjau. Jurnal Teknologi Lingkungan, 24(1), 21-27. https://doi.org/10.55981/jtl.2023.246",
];

// ================= BUILD DOCUMENT =================
const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: SZ, color: BLACK } } },
  },
  sections: [
    // ===== BAGIAN DEPAN: sampul, daftar isi, daftar tabel (angka romawi) =====
    {
      properties: {
        page: { ...PAGE, pageNumbers: { start: 1, formatType: NumberFormat.LOWER_ROMAN } },
      },
      footers: footer(),
      children: [
        new Paragraph({ spacing: { after: 200 }, children: [] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "MAKALAH", font: FONT, size: 28, bold: true, color: BLACK })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "[JUDUL MAKALAH ANDA \u2014 sebutkan proses kimia/fisikanya, kasusnya, dan sudut pandangnya]", font: FONT, size: 28, bold: true, color: BLACK })] }),
        new Paragraph({ spacing: { after: 600 }, children: [] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: "Diajukan untuk memenuhi Tugas Terstruktur 1", font: FONT, size: SZ, color: BLACK })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: "Mata Kuliah Dasar Kimia dan Fisika Lingkungan", font: FONT, size: SZ, color: BLACK })] }),
        new Paragraph({ spacing: { after: 600 }, children: [] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: "Disusun oleh:", font: FONT, size: SZ, color: BLACK })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "[Nama Lengkap Mahasiswa]", font: FONT, size: SZ, color: BLACK })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "NIM. ........................", font: FONT, size: SZ, color: BLACK })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: "Kelas: ........................", font: FONT, size: SZ, color: BLACK })] }),
        new Paragraph({ spacing: { after: 600 }, children: [] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: "Dosen Pengampu:", font: FONT, size: SZ, color: BLACK })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 800 }, children: [new TextRun({ text: "Roni Saputra, S.Si., M.Si.", font: FONT, size: SZ, color: BLACK })] }),
        new Paragraph({ spacing: { after: 800 }, children: [] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "PROGRAM STUDI S1 KESEHATAN LINGKUNGAN", font: FONT, size: SZ, bold: true, color: BLACK })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "FAKULTAS ILMU KESEHATAN", font: FONT, size: SZ, bold: true, color: BLACK })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "UNIVERSITAS IBNU SINA", font: FONT, size: SZ, bold: true, color: BLACK })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ text: "2026", font: FONT, size: SZ, bold: true, color: BLACK })] }),

        pageBreak(),
        heading1("DAFTAR ISI"),
        new TableOfContents("Daftar Isi", { hyperlink: true, headingStyleRange: "1-2" }),

        pageBreak(),
        heading1("DAFTAR TABEL"),
        new Table({
          width: { size: 9000, type: WidthType.DXA },
          columnWidths: [1200, 6000, 1800],
          rows: [
            new TableRow({ tableHeader: true, children: [cell("Nomor", { header: true, width: 1200, align: AlignmentType.CENTER }), cell("Judul", { header: true, width: 6000, align: AlignmentType.CENTER }), cell("Halaman", { header: true, width: 1800, align: AlignmentType.CENTER })] }),
            new TableRow({ children: [cell("Tabel 1", { width: 1200 }), cell("Ringkasan literatur yang digunakan", { width: 6000 }), cell("", { width: 1800 })] }),
          ],
        }),
        new Paragraph({ spacing: { before: 200 }, children: [new TextRun({ text: "Catatan: setelah dokumen final, klik kanan pada Daftar Isi \u2192 Update Field \u2192 Update entire table, agar nomor halaman terisi otomatis.", font: FONT, size: 20, italics: true, color: BLACK })] }),
      ],
    },

    // ===== BAGIAN ISI: BAB I - Lampiran (angka arab mulai 1) =====
    {
      properties: {
        page: { ...PAGE, pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL } },
      },
      footers: footer(),
      children: [
        heading1("BAB I"),
        heading1("PENDAHULUAN"),

        heading2("1.1 Latar Belakang"),
        instr("Susun 1.1 TERAKHIR, setelah BAB III selesai (lihat Langkah 5). Rangkai poin-poin di bawah menjadi paragraf memakai kalimat Anda sendiri, urutkan dari persoalan umum ke persoalan khusus kasus Anda, dan tutup dengan alasan topik ini penting bagi kesehatan lingkungan."),
        bullet("Definisi singkat eutrofikasi dan mekanisme dasarnya (rujuk Irianto & Triweko, 2019)."),
        bullet("Data beban pencemar KJA di Danau Maninjau: TN \u2248 84,20%, TP \u2248 91,83% dari total beban (Kurniati et al., 2021)."),
        bullet("Status trofik Danau Maninjau eutrofik-hipereutrofik pada 2011-2020, terutama di lokasi KJA (Sunaryani, 2023)."),
        bullet("Kematian ikan massal Desember 2025 (1.428,73 ton, kerugian \u2248 Rp32,86-35,7 miliar) dan Januari 2026 (\u2248 5 ton), keduanya dipicu upwelling akibat cuaca ekstrem (ANTARA News, 2025; Radar Sumbar, 2026)."),
        bullet("Kalimat penutup: mengapa persoalan ini relevan bagi kesehatan lingkungan (ekosistem, ekonomi warga, kualitas sumber air) \u2014 rumuskan sendiri."),

        heading2("1.2 Rumusan Masalah"),
        instr("Contoh rumusan masalah di bawah adalah draf awal, bukan versi final. Baca ulang, sesuaikan kalimatnya dengan gaya Anda sendiri, dan pastikan jumlahnya sama dengan jumlah tujuan (1.3) dan jumlah butir simpulan (4.1)."),
        bullet("Bagaimana proses masuknya nitrogen dan fosfor dari aktivitas keramba jaring apung ke perairan Danau Maninjau hingga menyebabkan eutrofikasi?"),
        bullet("Bagaimana perubahan kondisi fisika dan kimia perairan, khususnya stratifikasi, upwelling, dan penurunan oksigen terlarut, menyebabkan kematian ikan secara massal?"),
        bullet("Mengapa kondisi ini menjadi permasalahan kesehatan lingkungan dan bagaimana upaya penanganannya?"),

        heading2("1.3 Tujuan Penulisan"),
        instr("Nyatakan ulang tiap rumusan masalah di atas menjadi pernyataan tujuan (\"Menjelaskan...\", \"Menganalisis...\", \"Menguraikan...\"). Ketik di sini."),

        pageBreak(),
        heading1("BAB II"),
        heading1("TINJAUAN PUSTAKA"),

        heading2("2.1 Dasar Teori Kimia atau Fisika"),
        instr("Wajib memuat sedikitnya satu persamaan reaksi/matematis dan penjelasan tiap lambangnya, ditulis dengan kalimat Anda sendiri (bukan menyalin Lampiran 2 petunjuk tugas atau tabel di bawah ini mentah-mentah)."),
        body("Bahan mentah \u2014 persamaan dekomposisi aerobik bahan organik berbasis nisbah Redfield (C:N:P = 106:16:1), dari Irianto & Triweko (2019):", { noIndent: true, bold: true }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 100, after: 100 },
          children: [new TextRun({ text: "(CH2O)106(NH3)16(H3PO4) + 138 O2 \u2192 106 CO2 + 122 H2O + 16 HNO3 + H3PO4", font: FONT, size: SZ, bold: true, color: BLACK })],
        }),
        new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 200 }, children: [new TextRun({ text: "(2.1)", font: FONT, size: SZ, color: BLACK })] }),
        bullet("(CH2O)106(NH3)16(H3PO4): bahan organik hasil fotosintesis fitoplankton, mengikuti nisbah Redfield."),
        bullet("O2: oksigen terlarut yang dipakai bakteri untuk menguraikan bahan organik secara aerobik."),
        bullet("CO2, H2O: hasil akhir oksidasi karbon."),
        bullet("HNO3, H3PO4: nitrogen dan fosfor anorganik yang dilepas kembali ke air (mineralisasi)."),
        bullet("Catatan fisika sebagai bahan tambahan: stratifikasi termal (epilimnion-termoklin-hipolimnion) dan mekanisme upwelling saat stratifikasi terganggu oleh cuaca ekstrem \u2014 jelaskan keterkaitannya dengan persamaan (2.1) di atas dengan kalimat Anda sendiri."),
        instr("Cek jumlah atom di kedua ruas persamaan (2.1) sebelum final \u2014 lihat daftar periksa Bagian M."),

        heading2("2.2 Ringkasan Literatur"),
        instr("Tabel 1 di bawah adalah bahan mentah hasil penelusuran. WAJIB Anda periksa keberadaan tiap sumber sendiri lewat Google Scholar/DOI (Bagian 4.4 petunjuk) sebelum dipakai, dan isi kolom terakhir (dipakai pada subbab mana). Setelah tabel, tulis 2-3 paragraf sendiri: di mana sumber-sumber ini sepakat dan di mana berbeda penekanan."),
        tableCaption("Tabel 1. Ringkasan literatur yang digunakan"),
        litTable,
        tableSource("Sumber: olahan penulis"),
        instr("Tulis ulasan 2-3 paragraf Anda di sini."),

        pageBreak(),
        heading1("BAB III"),
        heading1("PEMBAHASAN"),

        heading2("3.1 Kronologi Kasus"),
        instr("Susun 3.1 LEBIH DULU (Langkah 5), sebelum bagian lain. Rangkai poin waktu di bawah menjadi narasi runtut dengan kalimat sendiri, maksimum sekitar satu halaman."),
        bullet("Latar jangka panjang: penurunan mutu air Danau Maninjau 2011-2020 hingga kelas cemar berat, status eutrofik-hipereutrofik di lokasi KJA (Sunaryani, 2023)."),
        bullet("Desember 2025: curah hujan tinggi + angin kencang \u2192 upwelling \u2192 1.428,73 ton ikan nila mati di 5 nagari (Sungai Batang, Tanjung Sani, Duo Koto, Maninjau, Koto Gadang Anam Koto); kerugian \u2248 Rp32,86-35,7 miliar (ANTARA News, 2025; Radar Sumbar, 2026)."),
        bullet("Respons: DKPP Agam menerbitkan surat imbauan, edaran, sosialisasi pasca-kejadian."),
        bullet("Januari 2026: kejadian berulang, \u2248 5 ton ikan mati, mekanisme upwelling serupa (Radar Sumbar, 2026)."),

        heading2("3.2 Rantai Sebab-Akibat Kimia dan Fisika"),
        instr("INI BAGIAN DENGAN BOBOT NILAI TERBESAR. Tulis dulu rantainya dalam SATU BARIS (Langkah 4), lalu kembangkan tiap mata rantai jadi paragraf dengan kalimat sendiri. Setiap kali Anda tulis APA yang terjadi, kalimat berikutnya harus menjelaskan MENGAPA itu terjadi secara kimia/fisika (rujuk kembali ke persamaan 2.1 dan konsep stratifikasi di 2.1)."),
        body("Rantai satu baris (bahan mentah, silakan sesuaikan redaksinya): sisa pakan dan metabolisme ikan KJA \u2192 beban N dan P berlebih \u2192 status eutrofik-hipereutrofik \u2192 dekomposisi aerobik bahan organik di dasar danau menghabiskan oksigen hipolimnion \u2192 cuaca ekstrem merusak stratifikasi termal \u2192 upwelling mengangkat air anoksik ke permukaan \u2192 hipoksia akut pada ikan dalam KJA \u2192 kematian ikan massal.", { noIndent: true, italics: true }),
        instr("Kembangkan tiap tanda panah di atas menjadi satu atau dua paragraf. Mata rantai yang tersedia datanya: beban N/P dari KJA (Kurniati et al., 2021); status trofik (Sunaryani, 2023); mekanisme dekomposisi (persamaan 2.1, Irianto & Triweko, 2019); pemicu cuaca dan upwelling Desember 2025 (ANTARA News, 2025). Jangan meloncati mata rantai mana pun \u2014 lihat Bagian O, kesalahan paling umum ada di sini."),

        heading2("3.3 Dampak Kesehatan dan Respons yang Diambil"),
        instr("Kembangkan poin-poin berikut menjadi paragraf sendiri."),
        bullet("Kerugian ekonomi petani KJA \u2248 Rp32,86-35,7 miliar (ANTARA News, 2025; Radar Sumbar, 2026); kelompok paling berisiko: petani KJA skala kecil."),
        bullet("Potensi dampak tambahan: pembusukan bangkai ikan menambah beban organik dan risiko paparan gas dekomposisi (amonia, H2S) bagi warga sekitar \u2014 cari rujukan tambahan bila memungkinkan."),
        bullet("Respons DKPP Agam: surat imbauan/edaran/sosialisasi, sifatnya reaktif (setelah kejadian)."),
        bullet("Indikasi keterlambatan: kejadian berulang Januari 2026 hanya \u2248 1 bulan setelah kejadian besar \u2014 jelaskan mengapa respons yang ada belum cukup."),

        heading2("3.4 Relevansi bagi Indonesia"),
        instr("Sebutkan kesamaan, perbedaan, dan akibat perbedaan itu bagi pilihan penanganan. Hindari kalimat umum seperti 'Indonesia harus berhati-hati'."),
        bullet("Kesamaan: pola KJA-eutrofikasi juga terjadi di Danau Batur (kontribusi KJA terhadap N \u2248 45,19%) dan danau vulkanik lain (Kurniadi et al., 2025)."),
        bullet("Perbedaan tahap trofik: Danau Maninjau sudah hipereutrofik dengan kematian ikan berulang; danau lain (mis. Batur, Toba) umumnya masih pada tahap trofik lebih rendah."),
        bullet("Implikasi: peluang pencegahan lewat penetapan daya tampung beban pencemaran (Permen LH No. 28/2009) masih terbuka untuk danau yang belum hipereutrofik, sedangkan Danau Maninjau memerlukan penanganan yang lebih mahal dan kompleks \u2014 uraikan dengan kalimat sendiri."),

        pageBreak(),
        heading1("BAB IV"),
        heading1("PENUTUP"),

        heading2("4.1 Simpulan"),
        instr("Jawab tiap rumusan masalah pada 1.2 secara berurutan. Tidak ada informasi baru dan tidak ada sitasi di bagian ini. Jumlah butir harus sama dengan jumlah rumusan masalah."),

        heading2("4.2 Saran"),
        instr("Tujukan saran ke pihak tertentu (mis. DKPP Kabupaten Agam, petani KJA, peneliti berikutnya). Hindari saran umum yang berlaku untuk topik apa pun."),

        pageBreak(),
        heading1("DAFTAR PUSTAKA"),
        instr("Sembilan rujukan di bawah adalah bahan mentah yang sudah saya telusuri. WAJIB Anda periksa ulang keberadaan tiap sumber (judul di Google Scholar dan/atau DOI di doi.org) sebelum final \u2014 lihat Bagian 4.4 petunjuk tugas. Sumber ini memenuhi syarat sedikitnya 4 jurnal ilmiah dan sedikitnya 2 sumber mekanisme + 2 sumber kasus, tetapi Anda boleh dan dianjurkan menambah rujukan lain hasil penelusuran Anda sendiri."),
        ...pustaka.map(refEntry),

        pageBreak(),
        heading1("LAMPIRAN"),
        heading2("Lampiran 1. Bukti Rujukan Jurnal"),
        instr("Tempelkan tangkapan layar halaman judul dan abstrak dari salah satu artikel jurnal yang Anda kutip, misalnya Kurniati et al. (2021) atau Sunaryani (2023)."),
        heading2("Lampiran 2. Lembar Deklarasi Penggunaan Kecerdasan Buatan"),
        instr("Isi dan tanda tangani sesuai template resmi dari petunjuk tugas (Lampiran 1 berkas petunjuk). Salinan kosongnya ada di halaman berikut."),

        pageBreak(),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300, line: 240 }, children: [new TextRun({ text: "LEMBAR DEKLARASI PENGGUNAAN KECERDASAN BUATAN", font: FONT, size: SZ, bold: true, color: BLACK })] }),
        body("Nama lengkap: ........................................................", { noIndent: true }),
        body("NIM: ........................................................", { noIndent: true }),
        body("Nomor topik: 6 (Eutrofikasi dan kematian ikan massal)", { noIndent: true }),
        body("Judul makalah: ........................................................", { noIndent: true }),
        new Paragraph({ spacing: { before: 200, after: 100, line: 240 }, children: [new TextRun({ text: "Perangkat kecerdasan buatan yang digunakan (tulis \"tidak ada\" apabila memang tidak digunakan):", font: FONT, size: SZ, color: BLACK })] }),
        new Table({
          width: { size: 9000, type: WidthType.DXA },
          columnWidths: [2500, 3000, 3500],
          rows: [
            new TableRow({ tableHeader: true, children: [cell("Nama perangkat", { header: true, width: 2500, align: AlignmentType.CENTER }), cell("Bagian yang dibantu", { header: true, width: 3000, align: AlignmentType.CENTER }), cell("Bentuk bantuan", { header: true, width: 3500, align: AlignmentType.CENTER })] }),
            new TableRow({ children: [cell(" ", { width: 2500 }), cell(" ", { width: 3000 }), cell(" ", { width: 3500 })] }),
            new TableRow({ children: [cell(" ", { width: 2500 }), cell(" ", { width: 3000 }), cell(" ", { width: 3500 })] }),
            new TableRow({ children: [cell(" ", { width: 2500 }), cell(" ", { width: 3000 }), cell(" ", { width: 3500 })] }),
          ],
        }),
        new Paragraph({ spacing: { before: 300 }, children: [] }),
        body("Dengan ini saya menyatakan bahwa makalah ini saya susun sendiri, seluruh rujukan yang tercantum telah saya periksa keberadaannya, dan setiap bantuan perangkat kecerdasan buatan telah saya cantumkan pada tabel di atas. Saya bersedia menerima sanksi akademik apabila pernyataan ini terbukti tidak benar.", { noIndent: true }),
        new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Batam, ......................... 2026", font: FONT, size: SZ, color: BLACK })] }),
        new Paragraph({ spacing: { before: 800 }, alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "(...........................................)", font: FONT, size: SZ, color: BLACK })] }),
        new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Nama dan NIM", font: FONT, size: SZ, color: BLACK })] }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  require("fs").writeFileSync("./Template_Kerangka_Makalah_Eutrofikasi_Maninjau.docx", buf);
  console.log("done");
});