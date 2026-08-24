const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  PageBreak, PageNumber, Header, Footer, PositionalTabAlignment, PositionalTabLeader,
  PositionalTabSize, Tab, LevelFormat, convertInchesToTwip
} = require("docx");
const fs = require("fs");

const FONT = "Times New Roman";
const SZ = 24; // 12pt
const SZ_SMALL = 22;

function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { after: 200, line: 360, ...opts.spacing },
    indent: opts.indent,
    children: Array.isArray(text)
      ? text
      : [new TextRun({ text, font: FONT, size: SZ, bold: opts.bold || false, italics: opts.italics || false })],
  });
}

function heading1(text, num) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 300 },
    pageBreakBefore: true,
    children: [
      new TextRun({ text: num ? `BAB ${num}` : text, font: FONT, size: 28, bold: true }),
      ...(num ? [new TextRun({ text: `\n${text}`, font: FONT, size: 28, bold: true, break: 1 })] : []),
    ],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 200 },
    children: [new TextRun({ text, font: FONT, size: 26, bold: true })],
  });
}

function toc(label, pg) {
  return new Paragraph({
    tabStops: [{ type: "right", leader: "dot", position: convertInchesToTwip(6.3) }],
    spacing: { after: 120 },
    children: [
      new TextRun({ text: label, font: FONT, size: SZ }),
      new TextRun({ text: `\t${pg}`, font: FONT, size: SZ }),
    ],
  });
}

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: FONT, size: SZ } },
    },
  },
  numbering: {
    config: [
      {
        reference: "bullet-list",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        ],
      },
      {
        reference: "num-list",
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        ],
      },
    ],
  },
  sections: [
    // ================= COVER =================
    {
      properties: { page: { size: { width: 11906, height: 16838 } } }, // A4
      children: [
        new Paragraph({ spacing: { after: 1600 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "HACKATHON", font: FONT, size: 28, bold: true })] }),
        new Paragraph({ spacing: { after: 800 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "JEDA : Aplikasi Pemulihan Adiksi Investasi Digital Berbasis ADAPTIVE EMA, TERAPI MINDFULNESS, DAN GAMIFIKASI KOMUNITAS",
            font: FONT, size: 28, bold: true,
          })] }),
        new Paragraph({ spacing: { before: 1200, after: 400 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "NAMA TIM:", font: FONT, size: SZ, bold: true })] }),
        new Paragraph({ spacing: { after: 800 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "BATAM LINUX USER GROUP", font: FONT, size: SZ, bold: true })] }),
        new Paragraph({ spacing: { after: 400 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "NAMA ANGGOTA:", font: FONT, size: SZ, bold: true })] }),
        p("TOMINGSE LINGGA — 4342411034", { align: AlignmentType.CENTER, spacing: { after: 100 } }),
        p("AFIF HAMZAH SIREGAR — 4342501040", { align: AlignmentType.CENTER, spacing: { after: 100 } }),
        p("FAWWAZ KHAIRY WAHID — 4342501050", { align: AlignmentType.CENTER, spacing: { after: 1600 } }),
        new Paragraph({ spacing: { before: 2400 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "POLITEKNIK NEGERI BATAM", font: FONT, size: SZ, bold: true })] }),
        p("Jl. Ahmad Yani Batam Kota, Kota Batam, Kepulauan Riau 29461", { align: AlignmentType.CENTER }),
      ],
    },
    // ================= DAFTAR ISI + DAFTAR GAMBAR =================
    {
      properties: { page: { size: { width: 11906, height: 16838 } } },
      headers: { default: new Header({ children: [] }) },
      footers: {
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SZ })] })] }),
      },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 },
          children: [new TextRun({ text: "DAFTAR ISI", font: FONT, size: 28, bold: true })] }),
        toc("DAFTAR ISI", "i"),
        toc("DAFTAR GAMBAR", "ii"),
        toc("BAB 1 PENDAHULUAN", "1"),
        toc("Latar Belakang", "1"),
        toc("Tujuan", "3"),
        toc("Manfaat", "3"),
        toc("BAB 2 GAGASAN", "4"),
        toc("Kondisi Aktual Objek Permasalahan", "4"),
        toc("Gagasan Dari Masalah Terkait Yang Pernah Diimplementasikan Sebelumnya", "5"),
        toc("Teknologi Yang Digunakan", "6"),
        toc("Bahasan Lengkap Gagasan", "7"),
        toc("Landasan Ilmiah Pendukung", "8"),
        toc("Prediksi Hasil Implementasi Dari Gagasan Yang Diajukan", "16"),
        toc("Uraian Peran Dan Kontribusi Dari Pihak-Pihak Yang Terlibat", "17"),
        toc("Tahapan-Tahapan Strategis", "18"),
        toc("BAB 3 KESIMPULAN", "19"),
        toc("DAFTAR PUSTAKA", "20"),
        toc("LAMPIRAN-LAMPIRAN", "23"),

        new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { after: 400 },
          children: [new TextRun({ text: "DAFTAR GAMBAR", font: FONT, size: 28, bold: true })] }),
        toc("Gambar 1. 1 Jumlah Pelaku Judi Online di Indonesia Berdasarkan Usia", "1"),
        toc("Gambar 2. 1 Tampilan antarmuka halaman onboarding", "9"),
        toc("Gambar 2. 2 Tampilan antarmuka halaman registrasi pengguna aplikasi", "10"),
        toc("Gambar 2. 3 Tampilan antarmuka halaman asesmen awal", "11"),
        toc("Gambar 2. 4 Tampilan antarmuka halaman beranda (home)", "12"),
        toc("Gambar 2. 5 Tampilan antarmuka halaman jurnal harian (daily assessment)", "13"),
        toc("Gambar 2. 6 Tampilan antarmuka halaman komunitas", "14"),
        toc("Gambar 2. 7 Tampilan antarmuka halaman profil pengguna", "15"),
      ],
    },
    // ================= BODY =================
    {
      properties: { page: { size: { width: 11906, height: 16838 } } },
      footers: {
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SZ })] })] }),
      },
      children: [
        // ---------- BAB 1 ----------
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 },
          children: [new TextRun({ text: "BAB 1 \nPENDAHULUAN", font: FONT, size: 28, bold: true })] }),

        heading2("Latar Belakang"),
        p("Indonesia mengalami lonjakan partisipasi investasi digital yang belum pernah terjadi sebelumnya. Data PT Kustodian Sentral Efek Indonesia mencatat jumlah investor pasar modal Indonesia pada akhir 2025 menembus 20,32 juta Single Investor Identification (SID), naik 37 persen dari 14,87 juta SID di akhir 2024, (Metro TV News, 2025). Di sisi aset kripto, OJK melaporkan jumlah konsumen mencapai 20,19 juta sepanjang 2025, dengan nilai transaksi kripto nasional mencapai Rp650,6 triliun pada 2024 dan Rp482,23 triliun pada 2025 menjadikan Indonesia salah satu negara dengan basis investor kripto ritel terbesar di dunia (CNBC Indonesia, 2026). Pada sektor Perdagangan Berjangka Komoditi (PBK) yang mencakup forex, indeks, dan komoditi, Bappebti mencatat nilai transaksi 2024 mencapai Rp33.214,89 triliun, naik 29,34% dari tahun sebelumnya, sementara volume transaksi Januari–Juli 2025 sudah mencapai 8,18 juta lot dengan nasabah aktif forex sebanyak 125 ribu orang per Juli 2025 (DDTC News, 2025; Media Indonesia, 2025)."),
        p("Ironisnya, ledakan jumlah investor ini tidak diimbangi dengan pemahaman risiko yang memadai. Berdasarkan Survei Nasional Literasi dan Inklusi Keuangan (SNLIK) 2025 yang diselenggarakan OJK bersama Badan Pusat Statistik, indeks literasi keuangan nasional secara umum memang meningkat menjadi 66,46 persen. Namun, khusus untuk sektor pasar modal, tingkat literasi masyarakat baru mencapai 17,78 persen (Harian Berkat, 2026) — kesenjangan yang sangat lebar antara jumlah investor yang tumbuh pesat dengan pemahaman mereka terhadap risiko instrumen yang mereka gunakan. OJK sendiri menyatakan minimnya pemahaman ini membuat masyarakat rentan terhadap penawaran investasi ilegal maupun perilaku spekulatif yang tidak terkendali."),
        p("Di balik angka pertumbuhan yang mengesankan tersebut, tersembunyi krisis kesehatan mental yang belum mendapat perhatian sistemis. Jika di saham, banyak orang yang ikut FOMO mengikuti influencer dan berakhir sangkut atau rugi karena secara psikologis ingin mendapat keuntungan instan. Secara psikologis dan neurologis, trading aset berisiko tinggi — terutama kripto yang beroperasi 24 jam 7 hari seminggu, dan forex dengan leverage ekstrem hingga 1:1.000 — mengaktifkan mekanisme reward otak yang identik dengan perjudian dan kecanduan zat, di mana setiap profit kecil melepaskan dopamin dan mendorong perilaku berulang secara kompulsif hingga berkembang menjadi adiksi. Pola ini sejalan dengan fenomena judi online di Indonesia yang telah terbukti secara empiris merusak kondisi psikologis penggunanya: PPATK mencatat perputaran dana judi online mencapai Rp286,84 triliun sepanjang 2025 dalam 422,1 juta transaksi, dengan indikasi tindak pidana asal perjudian mendominasi 47,49 persen dari seluruh laporan transaksi keuangan mencurigakan yang diterima lembaga tersebut pada tahun yang sama (Databoks Katadata, 2026). Data ini menegaskan bahwa mekanisme reward berbasis dopamin pada aktivitas finansial berisiko tinggi — baik itu judi maupun trading spekulatif — sama-sama mampu mendorong perilaku kompulsif dalam skala masif."),
        p("Pola demografi juga memperkuat urgensi masalah ini: dari populasi pengguna judi online di Indonesia, 60 persen berasal dari generasi Milenial dan Gen Z, dan 82 persen pengguna internet pernah terpapar iklan judi online, dengan Instagram dan Facebook sebagai platform dengan eksposur iklan tertinggi (UGM, 2024). Kelompok usia produktif yang sama inilah yang juga mendominasi populasi investor ritel saham, kripto, dan forex di Indonesia — menunjukkan tumpang tindih risiko psikologis antara adiksi judi dan adiksi investasi digital pada segmen usia yang identik."),
        p("PT Finex Bisnis Solusi Futures, broker forex teregulasi Bappebti, bahkan secara terbuka menyatakan bahwa \"pertumbuhan pesat trading perlu diimbangi dengan edukasi\" karena lonjakan partisipasi ritel belum berbanding lurus dengan literasi risiko (Antaranews, 2025). Kondisi ini diperparah oleh fakta bahwa 72% dari seluruh exchange kripto yang beroperasi di Indonesia hingga akhir 2025 masih dalam kondisi merugi, menunjukkan ekosistem yang secara fundamental belum sehat (OJK, 2026)."),
        p("Dampak finansial dari perilaku trading yang tidak terkontrol ini pun sangat nyata. Studi dalam Journal of Advanced Research in Accounting and Finance Management (2024) menganalisis ribuan akun trading ritel dan mengonfirmasi bahwa sekitar 90% trader mengalami kerugian finansial (YEF Advisor, 2025). Bappebti mencatat 177 pengaduan kerugian trader akibat broker tidak transparan hanya dalam data tahun 2023, dengan total kerugian pada satu kasus fraud forex di Bandar Lampung saja mencapai Rp66,5 miliar yang melibatkan 665 korban. OJK melaporkan kerugian masyarakat akibat kripto ilegal dan robot trading ilegal sepanjang 2021 mencapai Rp6,5 triliun, sedangkan total kerugian investasi bodong dalam 10 tahun terakhir diperkirakan mencapai Rp117,5 triliun (Katakini, 2022)."),
        p("Bukti akademis global semakin kuat mengonfirmasi bahwa adiksi trading bukan sekadar \"kebiasaan buruk\", melainkan gangguan klinis yang nyata dan terukur. Penelitian yang dipublikasikan di jurnal PLOS One oleh Meshi et al. (2026) dengan 239 responden menemukan korelasi konsisten antara problematic crypto asset trading dengan gejala depresi, kecemasan, dan isolasi sosial yang meningkat. Studi besar dalam Journal of Banking & Finance (2020) terhadap sampel representatif investor Belanda menemukan 4,4% menunjukkan gejala perjudian kompulsif, sementara di Korea angkanya mencapai 21,5% dan di Thailand 14,4% (Kalda et al., 2020). Systematic review di F1000Research menemukan median prevalensi problem gambling di kalangan trader keuangan adalah 7,6%, jauh lebih tinggi dibanding populasi non-trader (Flack & Morris, 2023). Bahkan pada 2025, para peneliti mempublikasikan skala diagnostik pertama untuk trading disorder yang menguji 403 trader amatir, dengan hasil 1 dari 10 memenuhi ambang batas klinis gangguan trading, hampir 1 dari 5 masuk kategori at-risk, dan di antara yang diklasifikasikan disordered, 58,5% secara independen memenuhi kriteria problem gambling (IFA, 2026)."),
        p("Pada konteks Indonesia, Bank Indonesia melalui BI Institute mengakui bahwa dampak psikologis investasi kripto mencakup kecanduan dan fear of missing out (FOMO), dengan berbagai laporan internasional mencatat kasus depresi hingga keinginan bunuh diri di kalangan investor muda yang mengalami kerugian besar (Bank Indonesia Institute, 2026). Leverage ekstrem yang ditawarkan broker forex ritel kepada trader pemula — mencapai rasio 1:500 hingga 1:1.000 — secara statistik berisiko sangat tinggi menimbulkan kerugian sekaligus memperkuat siklus adiktif loss-chasing yang sulit dihentikan tanpa intervensi struktural."),
        p("Meskipun masalah ini nyata dan terdokumentasi secara ilmiah maupun regulatoris, belum ada satu pun aplikasi digital di Indonesia yang dirancang khusus untuk membantu pemulihan adiksi investasi digital pada investor saham, kripto, dan trader forex. Aplikasi kesehatan mental yang tersedia saat ini bersifat generik dan tidak menyentuh konteks spesifik psikologi trading. Aplikasi mindfulness internasional seperti Headspace dan Calm tidak memiliki modul yang relevan untuk trader Indonesia. Kebijakan pemerintah pun masih berfokus pada pengawasan dan penegakan hukum di sisi penawaran, bukan pemulihan psikologis individu di sisi permintaan. Di sinilah celah kritis yang ingin diisi oleh JEDA (Jejak Emosi dan Deteksi Adiksi) — sebuah aplikasi pemulihan adiksi investasi digital yang mengintegrasikan Adaptive EMA, terapi mindfulness berbasis bukti, dan gamifikasi komunitas sebagai respons komprehensif terhadap krisis yang selama ini terabaikan."),

        heading2("Tujuan"),
        p("Adapun tujuan dari pengembangan aplikasi ini adalah:"),
        p("Merancang dan mengembangkan aplikasi mobile JEDA sebagai solusi digital pertama di Indonesia yang berfokus pada pemulihan adiksi investasi digital mencakup saham, kripto, dan forex.", { indent: { left: 720 } }),
        p("Mengimplementasikan mekanisme Adaptive EMA (Ecological Momentary Assessment) berbasis kecerdasan buatan untuk mendeteksi pola emosional dan perilaku adiktif pengguna secara real-time dan kontekstual.", { indent: { left: 720 } }),
        p("Mengintegrasikan modul terapi mindfulness berbasis bukti ilmiah yang disesuaikan dengan konteks psikologi trading sehingga pengguna dapat mengelola dorongan impulsif dalam pengambilan keputusan finansial.", { indent: { left: 720 } }),
        p("Membangun ekosistem gamifikasi komunitas yang mendorong pemulihan berkelanjutan melalui dukungan sosial, pencapaian bertahap, dan mekanisme akuntabilitas antar pengguna.", { indent: { left: 720 } }),
        p("Menyediakan dasbor pemantauan progres pemulihan yang terukur sehingga pengguna dapat melacak perubahan perilaku, pola emosi, dan tingkat ketahanan terhadap trigger adiksi dari waktu ke waktu.", { indent: { left: 720 } }),

        heading2("Manfaat"),
        p("Manfaat yang diperoleh dari aplikasi ini antara lain:"),
        p("Bagi pengguna individu: Memberikan pendampingan psikologis yang terstruktur dan berbasis bukti kepada investor dan trader yang mengalami gejala adiksi investasi digital, sehingga mereka dapat memulihkan kendali atas perilaku finansial dan emosional mereka secara mandiri dan terukur.", { indent: { left: 720 } }),
        p("Bagi keluarga dan lingkungan sosial: Menyediakan kanal edukasi dan komunitas bagi keluarga serta orang-orang terdekat pengguna untuk memahami dinamika adiksi trading dan berperan aktif dalam proses pemulihan sebagai sistem pendukung yang terampil.", { indent: { left: 720 } }),
        p("Bagi ekosistem kesehatan mental: Mengisi kekosongan layanan digital di Indonesia yang selama ini belum menyentuh permasalahan adiksi investasi digital secara spesifik, sekaligus menjadi rujukan data agregat anonim bagi peneliti dan klinisi yang mengkaji kesehatan mental trader.", { indent: { left: 720 } }),
        p("Bagi regulator dan pemangku kebijakan: Menyediakan bukti empiris mengenai prevalensi dan pola adiksi investasi digital di Indonesia yang dapat menjadi dasar kebijakan perlindungan investor, literasi risiko, dan regulasi platform trading berbasis data perilaku nyata.", { indent: { left: 720 } }),
        p("Bagi masyarakat luas: Berkontribusi pada terciptanya ekosistem investasi digital yang lebih sehat dan bertanggung jawab, di mana pertumbuhan partisipasi investor ritel diiringi dengan peningkatan kesadaran akan risiko psikologis dan kemampuan pengelolaan emosi dalam pengambilan keputusan finansial.", { indent: { left: 720 } }),

        // ---------- BAB 2 ----------
        new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { after: 300 },
          children: [new TextRun({ text: "BAB 2\nGAGASAN", font: FONT, size: 28, bold: true })] }),

        heading2("Kondisi Aktual Objek Permasalahan"),
        p("Perkembangan investasi digital di Indonesia mengalami pertumbuhan yang sangat pesat dalam beberapa tahun terakhir. Kemudahan akses terhadap aplikasi saham, aset kripto, dan perdagangan berjangka membuat masyarakat dapat melakukan transaksi kapan saja dan di mana saja hanya melalui perangkat seluler. Berdasarkan data OJK dan PT Kustodian Sentral Efek Indonesia (KSEI), jumlah investor pasar modal Indonesia terus meningkat setiap tahunnya dan telah mencapai lebih dari 20 juta investor pada tahun 2025. Namun demikian, tingkat literasi pasar modal masyarakat baru berada di angka 17,78 persen berdasarkan SNLIK 2025 (Harian Berkat, 2026) — kesenjangan besar antara jumlah partisipan dan kesiapan psikologis maupun pengetahuan risiko mereka."),
        p("Meskipun memberikan peluang ekonomi yang luas, perkembangan tersebut juga memunculkan permasalahan baru berupa perilaku investasi dan trading yang bersifat kompulsif. Sebagian investor dan trader melakukan transaksi secara berlebihan, terus memantau pergerakan harga, mengalami Fear of Missing Out (FOMO), serta melakukan pembelian atau penjualan berdasarkan dorongan emosional. Kondisi ini dapat menyebabkan kerugian finansial yang signifikan, gangguan psikologis, stres berkepanjangan, kecemasan, hingga menurunnya kualitas hubungan sosial dan produktivitas sehari-hari."),
        p("Berbagai penelitian menunjukkan bahwa aktivitas trading dengan frekuensi tinggi memiliki karakteristik psikologis yang serupa dengan perilaku adiktif lainnya karena melibatkan mekanisme penghargaan (reward system) pada otak. Sayangnya, hingga saat ini belum tersedia platform digital di Indonesia yang secara khusus dirancang untuk membantu individu mengidentifikasi, memantau, dan memulihkan perilaku adiksi investasi digital secara terstruktur. Oleh karena itu, diperlukan suatu solusi inovatif yang tidak hanya berfokus pada edukasi finansial, tetapi juga mendukung aspek kesehatan mental dan perubahan perilaku pengguna secara berkelanjutan."),

        heading2("Gagasan Dari Masalah Terkait Yang Pernah Diimplementasikan Sebelumnya"),
        p("Berbagai aplikasi kesehatan mental dan pengelolaan kebiasaan telah dikembangkan untuk membantu pengguna mengurangi perilaku adiktif. Beberapa aplikasi seperti Headspace dan Calm menyediakan layanan mindfulness dan meditasi yang bertujuan meningkatkan kesadaran diri serta mengurangi stres. Selain itu, aplikasi seperti Quit Tracker dan My QuitBuddy membantu pengguna memantau progres penghentian kebiasaan tertentu melalui sistem pencatatan aktivitas, target harian, serta dukungan komunitas."),
        p("Pembanding yang lebih relevan datang dari kategori aplikasi pemulihan adiksi perilaku seperti Sober Grid dan I Am Sober. Kedua aplikasi ini menggabungkan pelacak progres (sobriety tracker), jurnal harian, dan jejaring sosial sesama pengguna dalam pemulihan. I Am Sober bahkan telah dikembangkan untuk melacak berbagai jenis adiksi perilaku di luar penyalahgunaan zat, termasuk pola makan berlebih (binge eating) dan pornografi, menunjukkan bahwa arsitektur aplikasi berbasis pelacakan-dan-komunitas dapat diadaptasi lintas jenis adiksi. Survei terhadap konselor adiksi turut menunjukkan bahwa 74 persen konselor pernah merekomendasikan aplikasi atau situs pemulihan kepada pasiennya, dan 94 persen menyatakan akan merekomendasikan aplikasi yang telah diuji melalui studi ilmiah ketat seperti randomized controlled trial (RCT) yang terbukti membantu pasien mengurangi perilaku adiktif mereka. Ini menegaskan pentingnya landasan bukti ilmiah sejak tahap perancangan aplikasi, bukan sekadar fitur gamifikasi tanpa dasar klinis."),
        p("Meskipun memiliki manfaat dalam membantu perubahan perilaku, aplikasi-aplikasi tersebut masih bersifat umum dan belum dirancang untuk menangani permasalahan adiksi investasi digital secara spesifik. Aplikasi mindfulness belum memiliki mekanisme untuk mendeteksi kondisi emosional yang memicu perilaku trading impulsif, sedangkan aplikasi pelacak kebiasaan belum mampu mengidentifikasi pola psikologis yang berkaitan dengan aktivitas investasi dan trading. Di sisi lain, platform investasi yang ada saat ini lebih berfokus pada penyediaan layanan transaksi dan analisis pasar dibandingkan kesehatan psikologis pengguna."),
        p("Berdasarkan kondisi tersebut, JEDA hadir sebagai inovasi yang menggabungkan teknologi Adaptive Ecological Momentary Assessment (Adaptive EMA), terapi mindfulness berbasis bukti ilmiah, dan gamifikasi komunitas dalam satu platform terpadu — pendekatan yang belum pernah digabungkan secara khusus untuk konteks adiksi investasi digital di Indonesia maupun secara global."),

        heading2("Teknologi Yang Digunakan"),
        p("Aplikasi ini akan dikembangkan dalam bentuk mobile, sehingga dapat diakses dengan mudah oleh pengguna di mana saja dan kapan saja. Beberapa teknologi akan digunakan dalam proses pengembangan aplikasi \"JEDA\" ini, antara lain:"),
        p([new TextRun({ text: "TypeScript (Bahasa)", font: FONT, size: SZ, bold: true })], { indent: { left: 720 }, spacing: { after: 80 } }),
        p("TypeScript adalah bahasa pemrograman berbasis open-source yang dikembangkan oleh Microsoft sebagai superset dari JavaScript, yang menambahkan sistem static typing ketat dan fitur modern untuk membantu pengembang membangun serta memelihara aplikasi web skala besar dengan lebih aman dari bug (Goldberg, 2022).", { indent: { left: 720 } }),
        p([new TextRun({ text: "NPM (Node Package Manager)", font: FONT, size: SZ, bold: true })], { indent: { left: 720 }, spacing: { after: 80 } }),
        p("NPM adalah pengelola paket bawaan untuk lingkungan runtime Node.js yang berfungsi sebagai tempat penyimpanan (registry) sekaligus alat baris perintah (command-line interface) untuk mengunduh, berbagi, dan mengelola pustaka (library) atau dependensi JavaScript dalam suatu proyek pengembangan perangkat lunak (Flanagan, 2020).", { indent: { left: 720 } }),
        p([new TextRun({ text: "Golang", font: FONT, size: SZ, bold: true })], { indent: { left: 720 }, spacing: { after: 80 } }),
        p("Go (Golang) adalah bahasa pemrograman sumber terbuka yang dirancang oleh Google untuk membangun perangkat lunak yang sederhana, andal, dan efisien, dengan keunggulan utama pada kecepatan kompilasi, performa tinggi, serta dukungan bawaan yang sangat kuat untuk konkurensi (concurrency) melalui goroutines (Donovan & Kernighan, 2020).", { indent: { left: 720 } }),
        p([new TextRun({ text: "React Native", font: FONT, size: SZ, bold: true })], { indent: { left: 720 }, spacing: { after: 80 } }),
        p("React Native adalah kerangka kerja open-source yang diciptakan oleh Meta (Facebook) yang memungkinkan pengembang untuk membangun aplikasi seluler asli (native) untuk iOS dan Android secara bersamaan hanya dengan menggunakan JavaScript dan React, tanpa mengorbankan performa atau pengalaman pengguna (Eisenman & Native, 2020).", { indent: { left: 720 } }),
        p([new TextRun({ text: "PostgreSQL", font: FONT, size: SZ, bold: true })], { indent: { left: 720 }, spacing: { after: 80 } }),
        p("PostgreSQL adalah sistem manajemen basis data relasional objek (ORDBMS) berbasis open-source yang terkenal karena arsitekturnya yang sangat stabil, kepatuhan standar SQL yang ketat, serta skalabilitasnya yang tinggi dalam menangani beban kerja data yang kompleks dan besar (Riggs & Crivulci Kubis, 2020).", { indent: { left: 720 } }),
        p([new TextRun({ text: "Figma", font: FONT, size: SZ, bold: true })], { indent: { left: 720 }, spacing: { after: 80 } }),
        p("Figma adalah perangkat lunak desain berbasis awan yang digunakan untuk merancang antarmuka pengguna (UI/UX) aplikasi dan situs web, dengan keunggulan utama pada fitur kolaborasi waktu-nyata yang memungkinkan banyak desainer bekerja bersama dalam satu dokumen secara simultan (Flaherty, 2022).", { indent: { left: 720 } }),
        p([new TextRun({ text: "Draw.io", font: FONT, size: SZ, bold: true })], { indent: { left: 720 }, spacing: { after: 80 } }),
        p("Draw.io adalah perangkat lunak pembuat diagram berbasis web gratis dan open-source yang digunakan untuk membuat berbagai jenis bagan seperti alur kerja (flowchart), arsitektur jaringan, mind maps, hingga diagram UML, dengan keunggulan utama pada fleksibilitas integrasi penyimpanan awan, privasi data yang ketat, serta kemudahan akses tanpa perlu registrasi wajib (Steinke, 2021).", { indent: { left: 720 } }),

        heading2("Bahasan Lengkap Gagasan"),
        p("Aplikasi JEDA (Jejak Emosi dan Deteksi Adiksi) merupakan solusi digital komprehensif yang dirancang khusus untuk membantu pemulihan adiksi investasi digital di Indonesia. JEDA mengintegrasikan tiga pilar utama secara sinergis, yaitu Adaptive EMA sebagai mekanisme deteksi dan monitoring adaptif, terapi mindfulness berbasis bukti ilmiah sebagai intervensi terapeutik inti, dan gamifikasi komunitas sebagai sistem pendukung sosial yang berkelanjutan. Ketiga pilar ini bekerja secara terintegrasi dalam sebuah ekosistem digital yang ramah pengguna dan secara ilmiah dapat dipertanggungjawabkan."),

        heading2("Landasan Ilmiah Pendukung"),
        p([new TextRun({ text: "Adaptive Ecological Momentary Assessment (EMA)", font: FONT, size: SZ, bold: true })], { spacing: { after: 100 } }),
        p("Adaptive EMA pada JEDA mengacu pada konsep just-in-time adaptive intervention (JITAI), yaitu pendekatan yang mengirimkan intervensi tepat saat level stres, kelelahan, atau dorongan emosional pengguna mencapai ambang tertentu — ambang mana yang berbeda antarindividu dan dapat berubah seiring waktu (JITA-EMA, 2023). Bukti efektivitas pendekatan berbasis EMA pada domain adiksi lain memperkuat landasan ini: sebuah randomized controlled trial dua-lengan terhadap intervensi konsumsi alkohol pada 907 mahasiswa menemukan efek intervensi yang konsisten pada konsumsi alkohol mingguan dan frekuensi binge drinking pada berbagai titik waktu follow-up, dan menyimpulkan bahwa metode EMA lebih akurat menangkap efektivitas intervensi dibanding titik follow-up konvensional 1 dan 6 bulan yang selama ini digunakan riset perilaku adiktif (de Vries, Kersten & Bolman, 2013). Prinsip yang sama diadaptasi JEDA untuk mendeteksi trigger trading impulsif pengguna secara real-time berdasarkan pola emosi dan konteks harian mereka."),
        p([new TextRun({ text: "Terapi Mindfulness Berbasis Bukti", font: FONT, size: SZ, bold: true })], { spacing: { after: 100 } }),
        p("Landasan ilmiah mindfulness untuk adiksi perilaku berisiko finansial tinggi tergolong kuat dan terus berkembang. Sebuah systematic review tahun 2026 yang dipublikasikan di Journal of Behavioral Addictions terhadap 12 studi menyimpulkan bahwa intervensi berbasis mindfulness (mindfulness-based interventions/MBI) secara konsisten menurunkan frekuensi judi dan tingkat craving sekaligus meningkatkan berbagai outcome psikologis penggunanya (Fadardi et al., 2026). Temuan ini diperkuat oleh meta-analisis randomized controlled trial yang mencakup 17 studi dan 1.228 partisipan, yang menemukan efek keseluruhan mindfulness terhadap penurunan craving sebesar -0,70 pada interval kepercayaan 95 persen antara -1,15 hingga -0,26 — sebuah effect size yang tergolong besar dalam standar penelitian klinis (Demina et al., 2023). Salah satu teknik inti yang diadaptasi JEDA adalah urge surfing, yaitu teknik mengamati dorongan berjudi atau bertransaksi sebagai gelombang sementara yang muncul, memuncak, lalu mereda, alih-alih melawan atau langsung menuruti dorongan tersebut (Sancho et al., 2018). Teknik ini dirancang menjadi fitur intervensi cepat dalam JEDA yang muncul otomatis ketika sistem Adaptive EMA mendeteksi pengguna berada pada momen rawan mengambil keputusan trading impulsif."),
        p([new TextRun({ text: "Gamifikasi Komunitas", font: FONT, size: SZ, bold: true })], { spacing: { after: 100 } }),
        p("Efektivitas gamifikasi komunitas dalam mendukung pemulihan adiksi memiliki bukti empiris yang kuat dari aplikasi sejenis yang sudah beroperasi secara global. Fitur seperti leaderboard komunitas memanfaatkan insting kompetitif pengguna dan menciptakan rasa flow, sementara tugas kolaboratif menumbuhkan tujuan bersama dan fitur interaksi sesama pengguna membantu melawan isolasi sosial serta membangun akuntabilitas melalui pengaruh sosial (Digital Therapies for SUD, tandfonline.com). Data dari analisis pengguna Sober Grid — aplikasi pemulihan adiksi berbasis komunitas terbesar — menunjukkan bahwa 1.273 pengguna paling aktifnya memiliki rata-rata masa pemulihan (sobriety) 195,5 hari, dengan lebih dari 33.000 koneksi unilateral dan 14.000 koneksi bilateral antarpengguna, serta jumlah check-in harian yang berasosiasi signifikan dengan panjangnya masa pemulihan dan berkurangnya jumlah kekambuhan (Ashford et al., 2020). Data empiris ini menjadi dasar rasional bagi desain sistem gamifikasi komunitas JEDA, khususnya fitur check-in harian dan koneksi antarpengguna."),

        heading2("Prediksi Hasil Implementasi Dari Gagasan Yang Diajukan"),
        p("Berdasarkan landasan ilmiah dari komponen-komponen yang terintegrasi dalam JEDA serta analogi dengan hasil implementasi aplikasi serupa dalam domain adiksi lain, terdapat beberapa prediksi hasil yang dapat diproyeksikan secara terukur dari implementasi JEDA."),
        p("Pada dimensi klinis dan psikologis, meta-analisis terhadap intervensi mindfulness untuk adiksi secara konsisten menunjukkan pengurangan tingkat keinginan (craving) sebesar 30 hingga 40% setelah delapan minggu intervensi terstruktur (Witkiewitz et al., 2014), sebuah angka yang selaras dengan effect size -0,70 dari meta-analisis 17 RCT yang lebih baru (Demina et al., 2023). Mekanisme EMA yang dikombinasikan dengan intervensi berbasis konteks (EMI) terbukti meningkatkan kesadaran diri dan pengendalian impuls pada populasi dengan gangguan perilaku adiktif. Dengan demikian, pengguna JEDA yang mengikuti program pemulihan secara konsisten selama minimal 30 hari diprediksi akan mengalami penurunan frekuensi perilaku trading impulsif sebesar 35 hingga 50%, serta peningkatan kemampuan mengidentifikasi dan merespons trigger emosional secara konstruktif."),
        p("Pada dimensi finansial, pengurangan perilaku trading impulsif secara langsung berkorelasi dengan pengurangan kerugian finansial yang tidak terencana. Berdasarkan data bahwa 90% trader mengalami kerugian, mayoritas akibat keputusan emosional yang tidak rasional (YEF Advisor, 2025), intervensi yang berhasil memutus siklus trading impulsif diprediksi dapat membantu pengguna menghemat rata-rata 20 hingga 40% dari modal yang sebelumnya habis akibat perilaku loss-chasing dan FOMO."),
        p("Pada dimensi sosial dan komunitas, mengacu pada studi efektivitas komunitas pemulihan online seperti Sober Grid — yang mencatat korelasi nyata antara intensitas check-in harian dengan panjangnya masa pemulihan pada 1.273 pengguna paling aktifnya (Ashford et al., 2020) — implementasi sistem gamifikasi komunitas JEDA diprediksi akan menghasilkan tingkat retensi pengguna aktif sebesar 60 hingga 70% setelah tiga bulan pertama, jauh di atas rata-rata retensi aplikasi kesehatan mental generik yang berkisar 15 hingga 25%. Komunitas yang terbentuk juga diprediksi akan berkontribusi pada pengurangan isolasi sosial yang dialami pengguna, sebagaimana diidentifikasi oleh Meshi et al. (2026) sebagai salah satu korelat negatif utama dari adiksi kripto."),
        p("Pada dimensi kontribusi data dan kebijakan, agregasi data anonim dari pengguna JEDA berpotensi menghasilkan dataset pertama di Indonesia mengenai pola adiksi investasi digital yang dapat digunakan oleh peneliti, klinisi, dan pembuat kebijakan. Data ini dapat mendukung pengembangan regulasi perlindungan investor yang lebih berbasis bukti, program literasi finansial yang lebih tepat sasaran — mengingat tingkat literasi pasar modal nasional yang baru 17,78 persen (Harian Berkat, 2026) — dan pedoman klinis untuk penanganan trading disorder di fasilitas kesehatan mental Indonesia."),

        heading2("Uraian Peran Dan Kontribusi Dari Pihak-Pihak Yang Terlibat Dalam Membantu Implementasi Gagasan"),
        p("Dalam tahap pengembangan aplikasi \"JEDA\", keberhasilan implementasi gagasan tidak hanya ditentukan oleh tim pengembang, tetapi juga membutuhkan kolaborasi dari berbagai pihak. Setiap pihak memiliki kontribusi yang berbeda untuk saling melengkapi. Berikut adalah peran dari pihak-pihak yang terlibat dalam implementasi gagasan:"),
        p([new TextRun({ text: "Akademisi (Peneliti)", font: FONT, size: SZ, bold: true })], { spacing: { after: 100 } }),
        p("Peneliti dari bidang psikologi dan kesehatan mental memiliki kontribusi penting dalam menyediakan landasan ilmiah untuk mengimplementasikan gagasan ini. Jurnal-jurnal terkait Cognitive Behavioral Therapy (CBT), Ecological Momentary Assessment (EMA), dan Ecological Momentary Intervention (EMI) digunakan sebagai referensi untuk memastikan aplikasi memiliki kerangka yang berbasis bukti ilmiah."),
        p([new TextRun({ text: "Media dan Pemerintah", font: FONT, size: SZ, bold: true })], { spacing: { after: 100 } }),
        p("Media daring dan pemerintah menjadi sumber data terkait fenomena maraknya perjudian online dan investasi berisiko di masyarakat Indonesia. Laporan investigasi dan pemberitaan terkait kasus terbaru membantu tim untuk merumuskan masalah, pola penyebaran, demografi pengguna serta dampak yang ditimbulkan. Data-data tersebut menjadi landasan dalam menentukan fitur yang dibutuhkan dalam aplikasi."),
        p([new TextRun({ text: "Investor dan Trader Berisiko Tinggi", font: FONT, size: SZ, bold: true })], { spacing: { after: 100 } }),
        p("Kelompok pengguna yang menjadi sasaran utama aplikasi JEDA mencakup individu yang terlibat dalam aktivitas spekulasi finansial berisiko tinggi, seperti perdagangan saham, aset kripto (cryptocurrency), dan foreign exchange (forex). Keterlibatan kelompok ini sangat krusial dalam proses pengembangan aplikasi, mengingat mereka adalah pihak yang paling memahami pengalaman, pola perilaku, serta tantangan yang dihadapi dalam mengendalikan dorongan spekulatif. Masukan dari kelompok ini akan diintegrasikan melalui pendekatan user-centered design, di mana kebutuhan, kebiasaan, dan kondisi psikologis mereka dijadikan acuan utama dalam merancang fitur aplikasi. Dengan demikian, JEDA dapat menghadirkan intervensi yang relevan, personal, dan efektif sesuai dengan konteks nyata yang dialami pengguna."),

        heading2("Tahapan-Tahapan Strategis"),
        p("Pengembangan aplikasi \"JEDA\" ini menggunakan metode Agile dengan kerangka kerja Scrum. Tahapan metode ini dipilih karena bersifat iteratif jangka pendek (sprint), adaptif, dan memungkinkan tim melakukan evaluasi berkala supaya pengembangan lebih terukur dan terarah. Adapun tahapan-tahapan yang direncanakan dalam metode Scrum ini dibagi menjadi 4 tahap yaitu sebagai berikut:"),
        p([new TextRun({ text: "Sprint Planning", font: FONT, size: SZ, bold: true })], { indent: { left: 720 }, spacing: { after: 80 } }),
        p("Tahap sprint planning adalah tahapan awal yang dilakukan untuk mempersiapkan tujuan yang ingin dicapai dalam pengembangan aplikasi \"JEDA\". Pada tahap ini, tim melakukan rapat persiapan untuk menyamakan tujuan tentang aplikasi, ruang lingkup, serta peran tiap anggota tim. Selain itu juga, menyusun product backlog yaitu daftar terkait fitur-fitur dari aplikasi yang akan dikembangkan. Setiap item diberi prioritas berdasarkan tingkat urgensi.", { indent: { left: 720 } }),
        p([new TextRun({ text: "The Sprint", font: FONT, size: SZ, bold: true })], { indent: { left: 720 }, spacing: { after: 80 } }),
        p("Sprint adalah tahap inti pengembangan aplikasi. Pada tahap ini, tim mengerjakan backlog yang sudah disepakati pada tahap sprint planning. Dalam tahap ini terdapat proses requirement (pengumpulan data, analisis kebutuhan teknis), design (membuat rancangan sistem dan antarmuka), develop (proses pengkodean aplikasi oleh tim pengembang), testing (proses pengujian fitur berjalan dengan baik dan perbaikan bug), serta deploy (merilis aplikasi). Setiap sprint menghasilkan produk yang bisa diuji.", { indent: { left: 720 } }),
        p([new TextRun({ text: "Daily Scrum", font: FONT, size: SZ, bold: true })], { indent: { left: 720 }, spacing: { after: 80 } }),
        p("Selama sprint berlangsung, dilakukan daily scrum yaitu pertemuan singkat setiap harinya. Pertemuan ini bertujuan untuk mengecek progres pekerjaan, rencana kerja, serta kendala yang muncul selama proses pengembangan aplikasi. Melalui komunikasi rutin ini, kendala dapat segera ditemukan di awal dan diselesaikan sehingga alur sprint berikutnya tidak terganggu.", { indent: { left: 720 } }),
        p([new TextRun({ text: "Scrum Review", font: FONT, size: SZ, bold: true })], { indent: { left: 720 }, spacing: { after: 80 } }),
        p("Tahapan terakhir dari setiap sprint ialah sprint review, yaitu tim mempresentasikan hasil kerjanya kepada pembimbing seperti fitur asesmen yang sudah berfungsi atau prototipe forum komunitas. Kemudian tim membahas apa yang sudah berjalan dengan baik, apa yang perlu diperbaiki, dan strategi sprint berikutnya (Revou, 2025).", { indent: { left: 720 } }),
        p("Setelah melakukan beberapa tahapan strategis berbasis Scrum, aplikasi JEDA dikembangkan secara bertahap, terukur, dan adaptif sehingga setiap proses yang dilakukan berjalan sesuai dengan tujuan utama dari pengembangan aplikasi ini."),

        // ---------- BAB 3 ----------
        new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { after: 300 },
          children: [new TextRun({ text: "BAB 3 \nKESIMPULAN", font: FONT, size: 28, bold: true })] }),
        p("Indonesia sedang menghadapi dua krisis yang saling berkaitan: ledakan partisipasi investasi digital yang belum pernah terjadi sebelumnya, dan absennya infrastruktur pemulihan psikologis bagi jutaan investor dan trader yang rentan terhadap adiksi — diperparah oleh kesenjangan literasi pasar modal yang baru mencapai 17,78 persen di tengah jumlah investor yang menembus 20 juta orang. Bukti akademis global yang terus berkembang, termasuk publikasi skala diagnostik trading disorder pertama di dunia pada 2025 dan berbagai studi yang mengonfirmasi prevalensi gejala perjudian kompulsif di kalangan trader, menegaskan bahwa adiksi investasi digital adalah krisis kesehatan mental yang nyata, terukur, dan mendesak untuk ditangani secara sistematis."),
        p("Aplikasi JEDA hadir sebagai respons inovatif terhadap kekosongan solusi yang selama ini terabaikan. Dengan mengintegrasikan Adaptive EMA sebagai mekanisme deteksi cerdas yang berkembang bersama penggunanya, terapi mindfulness yang telah terbukti secara ilmiah menurunkan craving dengan effect size besar pada berbagai jenis adiksi perilaku, dan gamifikasi komunitas yang mengubah pemulihan dari perjalanan yang menyendiri menjadi pengalaman kolektif yang saling menguatkan — sebagaimana dibuktikan aplikasi sejenis seperti Sober Grid — JEDA menawarkan pendekatan holistik dan berbasis bukti yang belum pernah ada sebelumnya di Indonesia."),
        p("Keberhasilan JEDA tidak hanya akan berdampak pada individu yang terbantu pulih dari adiksi investasi digital, tetapi juga berkontribusi pada pembangunan ekosistem investasi yang lebih sehat dan bertanggung jawab di Indonesia, di mana pertumbuhan jumlah investor ritel diiringi dengan peningkatan kematangan psikologis dan emosional dalam pengambilan keputusan finansial. Dengan fondasi ilmiah yang kuat, teknologi yang tepat, dan kolaborasi lintas sektor yang terstruktur, JEDA berpotensi menjadi standar nasional dalam pemulihan adiksi investasi digital dan berkontribusi nyata pada visi Indonesia yang lebih sehat, cerdas, dan sejahtera."),

        // ---------- DAFTAR PUSTAKA ----------
        new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { after: 300 },
          children: [new TextRun({ text: "DAFTAR PUSTAKA", font: FONT, size: 28, bold: true })] }),

        p("Antaranews (2025) 'Finex: Pertumbuhan pesat trading perlu diimbangi dengan edukasi', Antaranews, 17 November. Tersedia pada: https://www.antaranews.com/berita/5248789/finex-pertumbuhan-pesat-trading-perlu-diimbangi-dengan-edukasi (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Ashford, R.D. dkk. (2020) 'Digital recovery networks: Characterizing user participation, engagement, and outcomes of a novel recovery social network smartphone application', Journal of Substance Abuse Treatment. Tersedia pada: https://pmc.ncbi.nlm.nih.gov/articles/PMC7059216 (Diakses: 3 Juli 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Badan Pengawas Perdagangan Berjangka Komoditi (Bappebti) (2025) Perkembangan perdagangan berjangka komoditi: 7.636 entitas ilegal diblokir 2021–2025. Tersedia pada: https://bappebti.go.id/resources/docs/aktualita_2025_01_01_08uw9j3s_id.pdf (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Bank Indonesia Institute (2026) 'Investasi kripto: Antara untung, buntung dan depresi', BI Epsilon, 6 Juni. Tersedia pada: https://www.bi.go.id/id/bi-institute/BI-Epsilon/Pages/Investasi-Kripto-Antara-Untung,-Buntung-dan-Depresi.aspx (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("CNBC Indonesia (2026) 'Data OJK: Investor kripto naik, nilai transaksi turun', CNBC Indonesia, 20 Januari. Tersedia pada: https://www.cnbcindonesia.com/market/20260121135519-17-704053/data-ojk-investor-kripto-naik-nilai-transaksi-turun (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Databoks Katadata (2026) 'PPATK: Perputaran dana judol Indonesia turun pada 2025', Databoks, 3 Februari. Tersedia pada: https://databoks.katadata.co.id/demografi/statistik/697c5275ed76f/ppatk-perputaran-dana-judol-indonesia-turun-pada-2025 (Diakses: 3 Juli 2026).", { indent: { left: 720, hanging: 720 } }),
        p("DDTC News (2025) 'Nilai transaksi perdagangan berjangka komoditi 2024 naik 29,3 persen', DDTC News, 26 Januari. Tersedia pada: https://news.ddtc.co.id/berita/nasional/1808492/nilai-transaksi-perdagangan-berjangka-komoditi-2024-naik-293-persen (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("de Vries, H., Kersten, R. dan Bolman, C. (2013) 'Using ecological momentary assessment in testing the effectiveness of an alcohol intervention: A two-arm parallel group randomized controlled trial', Journal of Medical Internet Research. Tersedia pada: https://www.ncbi.nlm.nih.gov/pmc/articles/PMC3818333 (Diakses: 3 Juli 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Demina, A., Petit, B., Meille, V. dan Trojak, B. (2023) 'Mindfulness interventions for craving reduction in substance use disorders and behavioral addictions: Systematic review and meta-analysis of randomized controlled trials', BMC Neuroscience. Tersedia pada: https://www.ncbi.nlm.nih.gov/pmc/articles/PMC10583418 (Diakses: 3 Juli 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Fadardi, J. dkk. (2026) 'Mindfulness-based interventions for gambling disorder: A systematic review', Journal of Behavioral Addictions, 15(1). Tersedia pada: https://www.akjournals.com/view/journals/2006/15/1/article-p115.xml (Diakses: 3 Juli 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Flack, M. dan Morris, M. (2023) 'Association between gambling and financial trading: A systematic review', F1000Research. Tersedia pada: https://f1000research.com/articles/12-111 (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Harian Berkat (2026) 'Perluas basis investor ritel, OJK perkuat literasi pasar modal, SEPMT 2026 di Banten', Harian Berkat, 9 April. Tersedia pada: https://harianberkat.com/2026/04/09/perluas-basis-investor-ritel-ojk-perkuat-literasi-pasar-modal-sepmt-2026-di-banten (Diakses: 3 Juli 2026).", { indent: { left: 720, hanging: 720 } }),
        p("IFA.com (2026) '24-hour markets may fuel disordered trading, experts warn', IFA.com, 8 Februari. Tersedia pada: https://www.ifa.com/articles/just_diagnosing_trading_addiction_expanding_trading_hours (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Kalda, A., Loos, B., Previtero, A. dan Hackethal, A. (2020) 'Cryptocurrency trading and individual investors' behavior', Journal of Banking & Finance. Tersedia pada: https://www.sciencedirect.com/science/article/abs/pii/S0378426619302808 (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Katakini (2022) 'Sepanjang 2021, kerugian masyarakat akibat kripto dan robot trading ilegal capai Rp6,5 triliun', Katakini, 8 Maret. Tersedia pada: https://www.katakini.com/artikel/61992 (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Media Indonesia (2025) 'Volume transaksi PBK periode Januari–Juli 2025 capai 818 juta lot', Media Indonesia, 23 September. Tersedia pada: https://mediaindonesia.com/ekonomi/814066/volume-transaksi-pbk-periode-januari-juli-2025-capai-818-juta-lot (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Meshi, D., Jang, D., Mei, Z., Ratan, R. dan Foxman, M. (2026) 'Problematic cryptoasset trading is associated with depression, anxiety, and social isolation', PLOS One. Tersedia pada: https://journals.plos.org/plosone/article?id=10.1371/journal.pone.0349874 (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Metro TV News (2025) 'Kinerja KSEI moncer di 2025, investor pasar modal tembus 20 juta', Metro TV News, 29 Desember. Tersedia pada: https://www.metrotvnews.com/read/bw6C2PoJ-kinerja-ksei-moncer-di-2025-investor-pasar-modal-tembus-20-juta (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("OJK (2026) 'OJK ungkap 72% exchange kripto RI masih merugi', Investing.com Indonesia, 28 Januari. Tersedia pada: https://id.investing.com/news/stock-market-news/ojk-ungkap-72-exchange-kripto-ri-masih-merugi-ini-biang-keroknya-2920916 (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),
        p("OJK dan BPS (2025) 'Siaran pers bersama: Indeks literasi dan inklusi keuangan masyarakat meningkat', OJK, 2 Mei. Tersedia pada: https://ojk.go.id/id/berita-dan-kegiatan/siaran-pers/Pages/OJK-dan-BPS-Umumkan-Hasil-Survei-Nasional-Literasi-Dan-Inklusi-Keuangan-SNLIK-Tahun-2025.aspx (Diakses: 3 Juli 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Ramsey, A. dkk. (2022) 'Exploring whether addictions counselors recommend that their patients use websites, smartphone apps, or other digital health tools to help them in their recovery: Web-based survey', JMIR. Tersedia pada: https://pmc.ncbi.nlm.nih.gov/articles/PMC9253968 (Diakses: 3 Juli 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Sancho, M. dkk. (2018) 'Mindfulness-based interventions for the treatment of substance and behavioral addictions: A systematic review', Frontiers in Psychiatry. Tersedia pada: https://www.frontiersin.org/journals/psychiatry/articles/10.3389/fpsyt.2018.00095/full (Diakses: 3 Juli 2026).", { indent: { left: 720, hanging: 720 } }),
        p("Universitas Gadjah Mada (2024) 'Judi online makin marak di kalangan anak muda, pakar UGM sarankan edukasi literasi keuangan', UGM News, 11 Desember. Tersedia pada: https://ugm.ac.id/en/news/online-gambling-prevalent-among-young-people-ugm-expert-suggests-need-for-financial-literacy-education (Diakses: 3 Juli 2026).", { indent: { left: 720, hanging: 720 } }),
        p("YEF Advisor (2025) 'Mengapa 90% trader & investor saham mengalami kerugian?', YEF Advisor, 9 Maret. Tersedia pada: https://yefadvisor.com/mengapa-90-trader-investor-saham-mengalami-kerugian (Diakses: 15 Juni 2026).", { indent: { left: 720, hanging: 720 } }),

        // ---------- LAMPIRAN ----------
        new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { after: 300 },
          children: [new TextRun({ text: "LAMPIRAN-LAMPIRAN", font: FONT, size: 28, bold: true })] }),
        p("(Lampiran mockup antarmuka, biodata tim, dan dokumen pendukung lain disisipkan pada bagian ini.)", { italics: true }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("./JEDA_Proposal_Hackathon_Revisi.docx", buf);
  console.log("done");
});