const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  Numbering, LevelFormat, convertInchesToTwip, PageBreak
} = require("docx");

const FONT = "Calibri";

const H1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 300, after: 150 },
  children: [new TextRun({ text, bold: true, size: 30, font: FONT, color: "1F4E79" })],
});

const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 220, after: 100 },
  children: [new TextRun({ text, bold: true, size: 24, font: FONT, color: "2E75B6" })],
});

const P = (text, opts = {}) => new Paragraph({
  spacing: { after: 150, line: 300 },
  alignment: opts.align || AlignmentType.JUSTIFIED,
  children: [new TextRun({ text, font: FONT, size: 22, bold: opts.bold || false, italics: opts.italics || false })],
});

const Bullet = (text, level = 0) => new Paragraph({
  numbering: { reference: "bullet-list", level },
  spacing: { after: 80, line: 280 },
  children: [new TextRun({ text, font: FONT, size: 22 })],
});

const Numbered = (text, level = 0) => new Paragraph({
  numbering: { reference: "num-list", level },
  spacing: { after: 80, line: 280 },
  children: [new TextRun({ text, font: FONT, size: 22 })],
});

function cell(text, opts = {}) {
  return new TableCell({
    width: { size: opts.width || 2000, type: WidthType.DXA },
    shading: opts.header ? { type: ShadingType.CLEAR, fill: "2E75B6" } : undefined,
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({
        text,
        font: FONT,
        size: 20,
        bold: !!opts.header,
        color: opts.header ? "FFFFFF" : "000000",
      })],
    })],
  });
}

function row(cells) {
  return new TableRow({ children: cells });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullet-list",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 260 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 840, hanging: 260 } } } },
        ],
      },
      {
        reference: "num-list",
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 260 } } } },
        ],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 },
      },
    },
    children: [

      // ===== HALAMAN JUDUL =====
      new Paragraph({ spacing: { after: 200 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "PROPOSAL SOLUSI", font: FONT, size: 24, color: "808080" })] }),
      new Paragraph({ spacing: { after: 100 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "INFEST HACKATHON XII 2026", font: FONT, size: 24, bold: true, color: "2E75B6" })] }),
      new Paragraph({ spacing: { after: 400 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "\"AI for All: Intelligent Solutions for an Equitable Future\"", font: FONT, size: 22, italics: true })] }),

      new Paragraph({ spacing: { before: 400, after: 100 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "AJAR LURING", font: FONT, size: 44, bold: true, color: "1F4E79" })] }),
      new Paragraph({ spacing: { after: 400 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Platform Pembelajaran Adaptif Berbasis AI\nuntuk Siswa di Daerah 3T (Offline-First)", font: FONT, size: 24, bold: true })] }),

      new Paragraph({ spacing: { before: 500, after: 60 },
        children: [new TextRun({ text: "Nama Tim\t: [Isi nama tim]", font: FONT, size: 22 })] }),
      new Paragraph({ spacing: { after: 60 },
        children: [new TextRun({ text: "Sub-tema\t: Aksesibilitas Pendidikan (AI for Education Equity)", font: FONT, size: 22 })] }),
      new Paragraph({ spacing: { after: 60 },
        children: [new TextRun({ text: "Anggota Tim\t:", font: FONT, size: 22 })] }),
      Bullet("[Nama Anggota 1] — [Institusi] — Ketua Tim"),
      Bullet("[Nama Anggota 2] — [Institusi]"),
      Bullet("[Nama Anggota 3] — [Institusi]"),
      new Paragraph({ spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Kontak Ketua Tim\t: [email] / [No. HP]", font: FONT, size: 22 })] }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 1. LATAR BELAKANG MASALAH =====
      H1("1. Latar Belakang Masalah"),
      P("Pemerataan akses pendidikan masih menjadi tantangan besar di Indonesia, khususnya di wilayah 3T (Tertinggal, Terdepan, Terluar). Ketimpangan ini bukan hanya soal ketersediaan sekolah atau guru, tetapi juga soal akses terhadap teknologi pembelajaran modern. Sebagian besar platform belajar berbasis kecerdasan buatan yang berkembang saat ini — baik untuk personalisasi materi maupun asisten belajar interaktif — dirancang dengan asumsi koneksi internet yang stabil, sehingga tidak dapat diakses secara optimal oleh siswa di wilayah dengan keterbatasan jaringan."),
      P("Kondisi ini menciptakan ironi tersendiri: wilayah yang paling membutuhkan bantuan teknologi untuk pemerataan pendidikan justru menjadi wilayah yang paling sulit menjangkau teknologi tersebut. Siswa di daerah 3T sering kali hanya mengandalkan modul cetak atau materi yang diberikan guru secara terbatas, tanpa mekanisme yang dapat menyesuaikan tingkat kesulitan materi dengan kemampuan belajar masing-masing siswa."),
      P("Berdasarkan kondisi tersebut, dibutuhkan sebuah solusi pembelajaran adaptif berbasis kecerdasan buatan yang dapat berjalan sepenuhnya tanpa koneksi internet (offline-first), sehingga kecerdasan sistem tetap dapat menyesuaikan materi dengan kemampuan siswa meskipun perangkat tidak terhubung ke jaringan sama sekali."),
      H2("Data dan Fakta Pendukung"),
      Bullet("Wilayah 3T di Indonesia tersebar di berbagai provinsi, dengan tantangan utama berupa keterbatasan infrastruktur telekomunikasi dan listrik."),
      Bullet("Kesenjangan capaian belajar antara siswa di wilayah perkotaan dan wilayah 3T tercermin dari perbedaan hasil asesmen nasional antarwilayah."),
      Bullet("Sebagian besar solusi EdTech berbasis AI yang beredar saat ini (baik produk dalam maupun luar negeri) bergantung pada pemrosesan di server cloud, sehingga fitur personalisasinya tidak dapat diakses tanpa koneksi internet yang memadai."),

      // ===== 2. TUJUAN DAN SASARAN =====
      H1("2. Tujuan dan Sasaran"),
      H2("Tujuan Utama"),
      Numbered("Merancang dan mengembangkan prototipe aplikasi pembelajaran adaptif yang dapat berjalan sepenuhnya tanpa koneksi internet (offline-first)."),
      Numbered("Membangun model kecerdasan buatan ringan yang mampu memprediksi tingkat pemahaman siswa dan menyesuaikan materi belajar secara otomatis, dilatih menggunakan data respons latihan soal."),
      Numbered("Menyediakan mekanisme sinkronisasi opsional, sehingga progres belajar siswa dapat terekam di sistem pusat ketika perangkat kebetulan terhubung ke internet."),
      H2("Sasaran"),
      Bullet("Sasaran langsung: siswa jenjang pendidikan dasar dan menengah di wilayah 3T dengan keterbatasan atau ketiadaan akses internet."),
      Bullet("Sasaran tidak langsung: guru dan sekolah di wilayah 3T, yang memperoleh data agregat capaian belajar siswa saat sinkronisasi dilakukan."),
      Bullet("Sasaran jangka panjang: dinas pendidikan daerah, sebagai salah satu alat bantu pemantauan pemerataan capaian belajar antarwilayah."),

      // ===== 3. KONSEP SOLUSI =====
      H1("3. Konsep Solusi"),
      H2("3.1 Deskripsi Solusi"),
      P("Ajar Luring adalah aplikasi pembelajaran adaptif yang seluruh proses intelijensinya berjalan langsung di perangkat pengguna (on-device), bukan melalui pemanggilan API ke server AI di cloud. Aplikasi memadukan materi belajar yang telah diunduh sebelumnya dengan model kecerdasan buatan ringan yang memprediksi tingkat pemahaman siswa berdasarkan pola jawaban, kemudian menyesuaikan tingkat kesulitan dan jenis materi berikutnya secara otomatis — seluruhnya tanpa memerlukan koneksi internet."),
      H2("3.2 Inovasi dan Keunikan Dibandingkan Pendekatan yang Sudah Ada"),
      Bullet("Platform EdTech mapan seperti Ruangguru maupun asisten belajar berbasis AI internasional pada umumnya menyimpan konten secara offline, namun logika adaptif/personalisasinya tetap bergantung pada pemrosesan di server cloud."),
      Bullet("Ajar Luring memindahkan logika adaptif itu sendiri ke perangkat, menggunakan model prediksi kemampuan siswa berskala kecil yang dilatih dan dijalankan secara lokal, sehingga sistem tetap \"cerdas\" walau perangkat tidak pernah terhubung ke internet."),
      Bullet("Pendekatan ini secara khusus menyasar segmen yang selama ini kurang terlayani solusi EdTech arus utama: wilayah dengan konektivitas yang sangat terbatas atau tidak ada sama sekali, bukan sekadar \"koneksi lambat\"."),
      H2("3.3 Mekanisme Penggunaan AI (Alur Sistem, Peran AI, dan Logika Solusi)"),
      P("Komponen kecerdasan buatan pada sistem ini bukan model bahasa besar yang dipanggil secara instan, melainkan model prediksi kemampuan siswa yang dibangun dan dilatih secara khusus selama pengembangan, menggunakan pendekatan yang selaras dengan Item Response Theory (IRT) dan Bayesian Knowledge Tracing (BKT). Proses pengembangannya mencakup:"),
      Numbered("Kurasi dataset: menyusun dataset simulasi maupun data hasil uji coba berupa riwayat respons siswa terhadap butir soal (benar/salah, waktu pengerjaan, topik soal), sebagai bahan pelatihan model."),
      Numbered("Pelatihan model (training/fine-tuning): melatih model probabilistik ringan untuk memperkirakan tingkat penguasaan siswa terhadap suatu topik berdasarkan pola respons, serta memprediksi tingkat kesulitan soal berikutnya yang paling sesuai."),
      Numbered("Kompresi dan konversi model: mengonversi model hasil pelatihan ke format yang ringan dan dapat dieksekusi langsung di perangkat (on-device), tanpa keterlibatan server eksternal saat runtime."),
      Numbered("Integrasi pipeline: menghubungkan modul model prediksi dengan modul konten (bank soal dan materi) serta modul progres lokal, sehingga keputusan \"materi apa yang ditampilkan berikutnya\" sepenuhnya diambil oleh perangkat itu sendiri."),
      P("Peran kecerdasan buatan generatif (misalnya model bahasa) bersifat sekunder dan opsional — hanya digunakan pada saat perangkat kebetulan terhubung ke internet, untuk membantu menyusun penjelasan tambahan dalam bahasa yang lebih komunikatif. Inti pengambilan keputusan adaptif tetap berada pada model prediksi yang dilatih dan berjalan secara lokal."),

      H2("3.4 Business Model Canvas (BMC)"),
      new Table({
        width: { size: 9638, type: WidthType.DXA },
        columnWidths: [3213, 3213, 3212],
        rows: [
          row([cell("Komponen", { header: true, width: 3213 }), cell("Penjelasan", { header: true, width: 3213 }), cell("Keterangan", { header: true, width: 3212 })]),
          row([cell("Value Proposition", { width: 3213 }), cell("Pembelajaran adaptif berbasis AI yang tetap berfungsi tanpa koneksi internet sama sekali", { width: 3213 }), cell("Membedakan dari EdTech arus utama", { width: 3212 })]),
          row([cell("Customer Segments", { width: 3213 }), cell("Siswa jenjang dasar-menengah di wilayah 3T; sekolah dan dinas pendidikan daerah", { width: 3213 }), cell("Fokus B2G & B2B2C", { width: 3212 })]),
          row([cell("Channels", { width: 3213 }), cell("Distribusi melalui sekolah, kerja sama dinas pendidikan, dan unduhan via WiFi sekolah/kota terdekat", { width: 3213 }), cell("Instalasi awal sekali unduh", { width: 3212 })]),
          row([cell("Customer Relationships", { width: 3213 }), cell("Dukungan luring melalui guru sebagai fasilitator; sinkronisasi berkala saat ada akses internet", { width: 3213 }), cell("-", { width: 3212 })]),
          row([cell("Key Activities", { width: 3213 }), cell("Pengembangan model prediksi, kurasi konten, pelatihan guru fasilitator, pemeliharaan aplikasi", { width: 3213 }), cell("-", { width: 3212 })]),
          row([cell("Key Resources", { width: 3213 }), cell("Model AI on-device, bank soal & materi, tim pengembang, kemitraan dengan sekolah/dinas", { width: 3213 }), cell("-", { width: 3212 })]),
          row([cell("Key Partners", { width: 3213 }), cell("Dinas Pendidikan, sekolah di wilayah 3T, lembaga donor/CSR pendidikan", { width: 3213 }), cell("-", { width: 3212 })]),
          row([cell("Cost Structure", { width: 3213 }), cell("Pengembangan aplikasi, pelatihan model, distribusi perangkat/konten, pemeliharaan server sinkronisasi", { width: 3213 }), cell("-", { width: 3212 })]),
          row([cell("Revenue Streams", { width: 3213 }), cell("Kerja sama institusional (dinas pendidikan/CSR), lisensi B2G, model freemium untuk sekolah swasta", { width: 3213 }), cell("Non-profit inti, komersial opsional", { width: 3212 })]),
        ],
      }),
      new Paragraph({ spacing: { after: 200 } }),

      // ===== 4. DESAIN DAN PROTOTIPE SOLUSI =====
      H1("4. Desain dan Prototipe Solusi"),
      H2("4.1 Arsitektur Sistem"),
      P("Arsitektur sistem terbagi menjadi dua bagian utama: perangkat siswa yang beroperasi penuh secara offline, dan server pusat yang berperan opsional untuk sinkronisasi data."),
      new Table({
        width: { size: 9638, type: WidthType.DXA },
        columnWidths: [3213, 6425],
        rows: [
          row([cell("Komponen", { header: true, width: 3213 }), cell("Fungsi", { header: true, width: 6425 })]),
          row([cell("Model prediksi (on-device)", { width: 3213 }), cell("Memprediksi tingkat pemahaman siswa dan menentukan materi/soal berikutnya, berjalan lokal tanpa internet", { width: 6425 })]),
          row([cell("Cache materi", { width: 3213 }), cell("Menyimpan video, modul, dan bank soal yang telah diunduh sebelumnya", { width: 6425 })]),
          row([cell("Progres lokal", { width: 3213 }), cell("Menyimpan riwayat jawaban dan nilai siswa di perangkat (SQLite/Hive)", { width: 6425 })]),
          row([cell("Modul sinkronisasi", { width: 3213 }), cell("Mengirim data progres ke server pusat saat perangkat kebetulan terhubung internet", { width: 6425 })]),
          row([cell("Dashboard guru (server)", { width: 3213 }), cell("Menampilkan rekap capaian belajar siswa dan memperbarui materi baru", { width: 6425 })]),
        ],
      }),
      new Paragraph({ spacing: { after: 200 } }),
      H2("4.2 Alur Proses Penggunaan"),
      Numbered("Siswa membuka aplikasi di perangkat (tanpa memerlukan koneksi internet)."),
      Numbered("Siswa mengerjakan latihan soal pada suatu topik."),
      Numbered("Model prediksi on-device menganalisis pola jawaban dan memperbarui estimasi tingkat pemahaman siswa terhadap topik tersebut."),
      Numbered("Sistem secara otomatis menyesuaikan tingkat kesulitan atau memberikan materi remedial/lanjutan berdasarkan hasil prediksi tersebut."),
      Numbered("Seluruh riwayat disimpan secara lokal di perangkat."),
      Numbered("Ketika perangkat kebetulan terhubung ke internet (misalnya saat guru membawa perangkat ke wilayah berjaringan), data disinkronkan ke server pusat untuk keperluan dashboard guru dan pembaruan materi."),
      H2("4.3 Tech Stack"),
      new Table({
        width: { size: 9638, type: WidthType.DXA },
        columnWidths: [3213, 6425],
        rows: [
          row([cell("Lapisan", { header: true, width: 3213 }), cell("Teknologi", { header: true, width: 6425 })]),
          row([cell("Aplikasi (client)", { width: 3213 }), cell("Flutter — mendukung distribusi Android dengan performa ringan", { width: 6425 })]),
          row([cell("Model AI", { width: 3213 }), cell("Model IRT/BKT ringan, dilatih menggunakan Python, dikonversi ke format TensorFlow Lite untuk dijalankan on-device", { width: 6425 })]),
          row([cell("Penyimpanan lokal", { width: 3213 }), cell("SQLite / Hive", { width: 6425 })]),
          row([cell("Backend sinkronisasi", { width: 3213 }), cell("Node.js + PostgreSQL, atau Firebase sebagai alternatif ringan", { width: 6425 })]),
          row([cell("Dashboard guru", { width: 3213 }), cell("Aplikasi web sederhana berbasis React/Next.js", { width: 6425 })]),
        ],
      }),
      new Paragraph({ spacing: { after: 200 } }),

      // ===== 5. ANALISIS DAMPAK DAN MANFAAT =====
      H1("5. Analisis Dampak dan Manfaat"),
      H2("5.1 Dampak Sosial dan Pendidikan"),
      Bullet("Membuka akses pembelajaran adaptif berbasis AI bagi siswa di wilayah yang selama ini tidak terjangkau solusi EdTech arus utama karena keterbatasan konektivitas."),
      Bullet("Membantu guru di wilayah 3T memperoleh gambaran capaian belajar siswa secara lebih terukur melalui data yang tersinkronisasi secara berkala."),
      Bullet("Mendorong pemerataan kualitas pembelajaran antara wilayah perkotaan dan wilayah 3T."),
      H2("5.2 Manfaat bagi Target Pengguna"),
      Bullet("Siswa: memperoleh materi yang disesuaikan dengan tingkat kemampuan masing-masing, tanpa bergantung pada koneksi internet yang stabil."),
      Bullet("Guru: memperoleh alat bantu pemantauan capaian belajar siswa secara lebih personal dan berbasis data."),
      Bullet("Dinas pendidikan: memperoleh gambaran agregat capaian belajar antarwilayah sebagai dasar kebijakan pemerataan pendidikan."),

      // ===== 6. PENUTUP =====
      H1("6. Penutup"),
      H2("6.1 Kesimpulan"),
      P("Ajar Luring merupakan gagasan aplikasi pembelajaran adaptif berbasis kecerdasan buatan yang dirancang khusus untuk menjawab tantangan pemerataan pendidikan di wilayah 3T, dengan pendekatan offline-first yang memindahkan seluruh proses kecerdasan sistem ke perangkat pengguna. Melalui model prediksi kemampuan siswa yang dilatih dan dijalankan secara lokal, aplikasi ini tetap dapat beradaptasi dengan kebutuhan belajar siswa meskipun tanpa koneksi internet sama sekali."),
      H2("6.2 Harapan Pengembangan"),
      Bullet("Perluasan dukungan multibahasa daerah untuk meningkatkan relevansi konten bagi siswa di berbagai wilayah."),
      Bullet("Pengembangan dashboard admin yang lebih komprehensif untuk pengelolaan konten dan aturan sistem."),
      Bullet("Kerja sama dengan dinas pendidikan daerah untuk uji coba dan distribusi lebih luas."),

      // ===== 7. DAFTAR PUSTAKA =====
      H1("7. Daftar Pustaka"),
      P("[Daftar pustaka disusun mengikuti sumber data, referensi akademik terkait Item Response Theory dan Bayesian Knowledge Tracing, serta referensi kondisi pendidikan wilayah 3T yang digunakan tim dalam penyusunan proposal ini.]", { italics: true }),

      // ===== 8. LAMPIRAN =====
      H1("8. Lampiran (Opsional)"),
      P("[Lampiran dapat memuat wireframe tambahan, contoh dataset simulasi respons siswa, atau dokumen pendukung lain yang relevan.]", { italics: true }),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  require("fs").writeFileSync("./Proposal_Ajar_Luring.docx", buffer);
  console.log("done");
});