from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import cm

path = "Analisis_IMK_Shopee.pdf"

styles = getSampleStyleSheet()
title = ParagraphStyle(
    "TitleCustom", parent=styles["Title"], fontSize=20, leading=25,
    alignment=TA_CENTER, spaceAfter=18
)
h1 = ParagraphStyle("H1Custom", parent=styles["Heading1"], fontSize=14, leading=18, spaceBefore=8, spaceAfter=8)
h2 = ParagraphStyle("H2Custom", parent=styles["Heading2"], fontSize=11, leading=14, spaceBefore=6, spaceAfter=5)
body = ParagraphStyle("BodyCustom", parent=styles["BodyText"], fontSize=10, leading=15, spaceAfter=7)
small = ParagraphStyle("SmallCustom", parent=body, fontSize=9, leading=13)

doc = SimpleDocTemplate(
    path, pagesize=A4,
    rightMargin=1.8*cm, leftMargin=1.8*cm,
    topMargin=1.7*cm, bottomMargin=1.7*cm
)

story = []

story.append(Spacer(1, 2*cm))
story.append(Paragraph("ANALISIS INTERAKSI MANUSIA DAN KOMPUTER (IMK)", title))
story.append(Paragraph("<b>Analisis Antarmuka Produk Digital: Shopee</b>", ParagraphStyle(
    "sub", parent=body, alignment=TA_CENTER, fontSize=13, leading=18
)))
story.append(Spacer(1, 2*cm))
story.append(Paragraph("Nama: ______________________________", body))
story.append(Paragraph("NIM: _______________________________", body))
story.append(Paragraph("Kelas: ______________________________", body))
story.append(Paragraph("Mata Kuliah: Interaksi Manusia dan Komputer", body))
story.append(Spacer(1, 2*cm))
story.append(Paragraph("<i>Catatan: Screenshot asli Shopee perlu ditambahkan pada bagian yang disediakan.</i>", small))
story.append(PageBreak())

story.append(Paragraph("1. Nama Produk Digital", h1))
story.append(Paragraph(
    "<b>Nama Produk:</b> Shopee<br/>"
    "<b>Platform:</b> Android / iOS / Web<br/>"
    "<b>Jenis Produk:</b> E-commerce / Marketplace",
    body
))
story.append(Paragraph(
    "Shopee merupakan platform e-commerce yang memungkinkan pengguna mencari produk, "
    "melihat informasi produk, memasukkan barang ke keranjang, melakukan checkout, "
    "memilih metode pembayaran, dan memantau status pesanan.", body
))

story.append(Paragraph("2. Target Pengguna", h1))
story.append(Paragraph(
    "<b>Target pengguna:</b> Pengguna umum yang ingin membeli berbagai kebutuhan secara online, "
    "seperti mahasiswa, pekerja, maupun masyarakat umum.", body
))
story.append(Paragraph(
    "<b>Konteks penggunaan:</b> Shopee digunakan ketika pengguna ingin mencari produk, "
    "membandingkan harga dan ulasan, melakukan pembelian, serta memantau proses pengiriman. "
    "Pengguna dapat mengakses Shopee melalui smartphone maupun website.", body
))

story.append(Paragraph("3. Screenshot dan Temuan", h1))
story.append(Paragraph(
    "Pada dokumen final, lampirkan 3–5 screenshot asli Shopee. Tandai bagian antarmuka yang "
    "diamati menggunakan kotak atau lingkaran agar temuan mudah dikenali.", body
))

findings = [
    ("Screenshot 1 — Beranda",
     "Temuan: Banner promo berukuran besar lebih menarik perhatian dibandingkan beberapa informasi produk.",
     "Dampak: Pengguna dapat terdistraksi oleh promo ketika ingin langsung mencari atau melihat produk.",
     "Prinsip IMK: Attention (Perhatian)",
     "Alasan: Perhatian pengguna terbatas sehingga elemen visual yang lebih dominan cenderung lebih cepat menarik perhatian. Banner promo yang menonjol dapat mengarahkan perhatian pengguna sebelum melakukan pencarian produk."),
    ("Screenshot 2 — Hasil Pencarian",
     "Temuan: Foto, harga, rating, dan informasi produk ditampilkan secara langsung.",
     "Dampak: Pengguna lebih mudah mengenali dan membandingkan beberapa produk.",
     "Prinsip IMK: Perception (Persepsi)",
     "Alasan: Perception berkaitan dengan kemampuan pengguna menerima dan memahami informasi melalui indera. Penyajian informasi secara visual membantu pengguna mengenali informasi penting dengan lebih cepat."),
    ("Screenshot 3 — Detail Produk",
     "Temuan: Banyak informasi tersedia sebelum pengguna melakukan pembelian, seperti harga, rating, ulasan, variasi, dan pengiriman.",
     "Dampak: Pengguna dapat mengambil keputusan berdasarkan informasi yang tersedia, tetapi terlalu banyak informasi dapat meningkatkan beban berpikir.",
     "Prinsip IMK: Cognitive Load (Beban Kognitif)",
     "Alasan: Cognitive load berkaitan dengan jumlah informasi dan proses berpikir yang harus dilakukan pengguna. Informasi lengkap membantu keputusan, tetapi terlalu banyak informasi dapat membuat pengguna membutuhkan waktu lebih lama untuk menentukan pilihan."),
    ("Screenshot 4 — Checkout",
     "Temuan: Shopee menampilkan ringkasan pesanan dan total pembayaran sebelum transaksi selesai.",
     "Dampak: Pengguna dapat memeriksa kembali produk, alamat, metode pembayaran, dan total biaya.",
     "Prinsip IMK: Feedback & Human Error",
     "Alasan: Feedback membantu pengguna mengetahui kondisi atau hasil tindakan. Ringkasan checkout juga dapat membantu mencegah kesalahan, seperti salah alamat, produk, atau metode pembayaran.")
]

for heading, temuan, dampak, prinsip, alasan in findings:
    story.append(Paragraph(heading, h2))
    placeholder = Table([[Paragraph("<b>AREA SCREENSHOT</b><br/><br/>Tempel screenshot Shopee di area ini.<br/><br/><br/><br/><br/>", small)]],
                         colWidths=[16.8*cm], rowHeights=[5.0*cm])
    placeholder.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 1, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ]))
    story.append(placeholder)
    story.append(Spacer(1, 0.2*cm))
    for p in [temuan, dampak, prinsip, alasan]:
        story.append(Paragraph(p, body))
    story.append(Spacer(1, 0.25*cm))

story.append(PageBreak())
story.append(Paragraph("4. Tabel Ringkasan Analisis", h1))

data = [
    [Paragraph("<b>Screenshot</b>", small), Paragraph("<b>Temuan</b>", small), Paragraph("<b>Prinsip IMK</b>", small), Paragraph("<b>Dampak</b>", small)],
    ["Beranda", "Banner promo lebih dominan menarik perhatian.", "Attention", "Pengguna dapat terdistraksi."],
    ["Hasil pencarian", "Foto, harga, rating, dan informasi produk mudah dikenali.", "Perception", "Memudahkan perbandingan produk."],
    ["Detail produk", "Banyak informasi tersedia sebelum membeli.", "Cognitive Load", "Membantu keputusan tetapi dapat menambah beban berpikir."],
    ["Checkout", "Ringkasan dan total pembayaran ditampilkan.", "Feedback & Human Error", "Membantu memeriksa pesanan dan mengurangi kesalahan."]
]
table = Table(data, colWidths=[2.5*cm, 5.0*cm, 3.5*cm, 5.8*cm], repeatRows=1)
table.setStyle(TableStyle([
    ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
    ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
    ("VALIGN", (0,0), (-1,-1), "TOP"),
    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ("FONTSIZE", (0,0), (-1,-1), 8.5),
    ("LEADING", (0,0), (-1,-1), 11),
    ("LEFTPADDING", (0,0), (-1,-1), 5),
    ("RIGHTPADDING", (0,0), (-1,-1), 5),
    ("TOPPADDING", (0,0), (-1,-1), 5),
    ("BOTTOMPADDING", (0,0), (-1,-1), 5),
]))
story.append(table)

story.append(Spacer(1, 0.6*cm))
story.append(Paragraph("5. Kesimpulan Singkat", h1))
story.append(Paragraph(
    "Berdasarkan hasil pengamatan, Shopee telah menerapkan beberapa prinsip Interaksi Manusia "
    "dan Komputer (IMK) dalam antarmukanya. Informasi produk seperti harga, foto, rating, dan "
    "ulasan membantu pengguna mengenali dan membandingkan produk. Sistem juga memberikan "
    "feedback melalui ringkasan pesanan dan status transaksi. Namun, banyaknya promo dan "
    "informasi pada beberapa halaman dapat menarik perhatian pengguna dan meningkatkan beban "
    "kognitif. Secara umum, antarmuka Shopee mendukung pengguna dalam menyelesaikan proses "
    "belanja, mulai dari mencari produk hingga checkout, dengan beberapa aspek yang masih dapat "
    "memengaruhi kenyamanan interaksi pengguna.", body
))

story.append(Paragraph("6. Refleksi", h1))
story.append(Paragraph(
    "Jika saya menjadi UI/UX Designer, saya perlu memahami bagaimana manusia melihat, berpikir, "
    "mengingat, mengambil keputusan, dan melakukan kesalahan saat menggunakan sebuah aplikasi. "
    "Saya perlu memperhatikan agar informasi tidak membuat pengguna bingung atau terdistraksi. "
    "Saya juga perlu memahami kebiasaan dan ekspektasi pengguna agar alur pencarian hingga "
    "checkout terasa mudah. Dengan memahami faktor manusia tersebut, saya dapat merancang "
    "antarmuka yang lebih mudah dipahami, nyaman digunakan, dan membantu pengguna menyelesaikan "
    "tugas dengan lebih efektif.", body
))

doc.build(story)

print(f"PDF berhasil dibuat: {path}")
