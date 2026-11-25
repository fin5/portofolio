# -*- coding: utf-8 -*-
"""
Generator deck portofolio Refin Izmi Abiyuna (ID + EN).
Style: dark/tech, senada dengan refinizmiabiyuna.netlify.app (gradient fuchsia -> cyan).
Output: Refin_Izmi_Abiyuna_Portofolio_ID.pptx & Refin_Izmi_Abiyuna_Portfolio_EN.pptx
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE = os.path.dirname(os.path.abspath(__file__))
PHOTO = os.path.normpath(os.path.join(BASE, "..", "assets", "images", "profil", "fotoporto.png"))

# ---------- Palette (match web dark theme) ----------
BG        = RGBColor(0x03, 0x07, 0x12)   # gray-950
CARD      = RGBColor(0x11, 0x18, 0x27)   # gray-900
LINE      = RGBColor(0x1F, 0x29, 0x37)   # gray-800
WHITE     = RGBColor(0xF9, 0xFA, 0xFB)
GRAY      = RGBColor(0x9C, 0xA3, 0xAF)   # gray-400
GRAY2     = RGBColor(0x6B, 0x72, 0x80)   # gray-500
FUCHSIA   = RGBColor(0xE8, 0x79, 0xF9)   # fuchsia-400
FUCHSIA_D = RGBColor(0xC0, 0x26, 0xD3)   # fuchsia-600
CYAN      = RGBColor(0x22, 0xD3, 0xEE)   # cyan-400
CYAN_D    = RGBColor(0x08, 0x91, 0xB2)   # cyan-600
BLUE      = RGBColor(0x3B, 0x82, 0xF6)   # blue-500
GREEN     = RGBColor(0x4A, 0xDE, 0x80)   # green-400

F_SANS = "Segoe UI"
F_MONO = "Consolas"

SW, SH = Inches(13.333), Inches(7.5)
YEAR = 2026

# ---------- Helpers ----------

def _lerp(c1, c2, t):
    return RGBColor(
        round(c1[0] + (c2[0] - c1[0]) * t),
        round(c1[1] + (c2[1] - c1[1]) * t),
        round(c1[2] + (c2[2] - c1[2]) * t),
    )


def box(slide, x, y, w, h, fill=CARD, line=LINE, lw=1.0, radius=0.08):
    sp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    try:
        sp.adjustments[0] = radius
    except Exception:
        pass
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    return sp


def grad_bar(slide, x, y, w, h, c1=FUCHSIA_D, c2=CYAN, n=18):
    seg = w / n
    for i in range(n):
        t = i / (n - 1)
        sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + i * seg, y, seg + 1, h)
        sp.fill.solid()
        sp.fill.fore_color.rgb = _lerp(c1, c2, t)
        sp.line.fill.background()
        sp.shadow.inherit = False


def text(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", PP_ALIGN.LEFT)
        if p.get("ls"):
            para.line_spacing = p["ls"]
        if p.get("sa") is not None:
            para.space_after = Pt(p["sa"])
        if p.get("sb") is not None:
            para.space_before = Pt(p["sb"])
        runs = p["t"] if isinstance(p["t"], list) else [(p["t"], {})]
        for txt, ov in runs:
            r = para.add_run()
            r.text = txt
            f = r.font
            f.size = Pt(ov.get("size", p.get("size", 12)))
            f.bold = ov.get("bold", p.get("bold", False))
            f.name = ov.get("font", p.get("font", F_SANS))
            f.color.rgb = ov.get("color", p.get("color", WHITE))
    return tb


def chip(slide, x, y, w, h, label, color=CYAN, fill=CARD, line=LINE, size=11, font=F_MONO, bold=False):
    box(slide, x, y, w, h, fill=fill, line=line, radius=0.5)
    text(slide, x, y, w, h, [{"t": label, "size": size, "color": color, "font": font,
                              "bold": bold, "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)


def chips_flow(slide, items, x0, y0, maxw, gap_x=0.14, gap_y=0.14, h=0.42, size=11,
               color=CYAN, font=F_MONO, char_w=0.088, pad=0.38):
    """Lay chips left-to-right, wrapping; returns bottom y (inches float)."""
    rows, cur, curw = [], [], 0.0
    for it in items:
        w = pad + char_w * len(it)
        if cur and curw + gap_x + w > maxw:
            rows.append(cur)
            cur, curw = [], 0.0
        cur.append((it, w))
        curw += (gap_x if len(cur) > 1 else 0) + w
    if cur:
        rows.append(cur)
    y = y0
    for row in rows:
        total = sum(w for _, w in row) + gap_x * (len(row) - 1)
        x = x0 + (maxw - total) / 2
        for label, w in row:
            chip(slide, Inches(x), Inches(y), Inches(w), Inches(h), label, color=color, size=size, font=font)
            x += w + gap_x
        y += h + gap_y
    return y


def new_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG
    bg.line.fill.background()
    bg.shadow.inherit = False
    grad_bar(s, 0, 0, SW.inches, 0.055)
    return s


def header(slide, kicker, title):
    text(slide, Inches(0.6), Inches(0.38), Inches(12.1), Inches(0.3),
         [{"t": kicker, "size": 12, "color": CYAN, "font": F_MONO, "bold": True}])
    text(slide, Inches(0.6), Inches(0.68), Inches(12.1), Inches(0.6),
         [{"t": title, "size": 30, "bold": True, "color": WHITE}])
    grad_bar(slide, Inches(0.62), Inches(1.32), Inches(1.7), Inches(0.045))


def footer(slide, n):
    text(slide, Inches(0.6), Inches(7.08), Inches(6), Inches(0.3),
         [{"t": "refinizmiabiyuna.netlify.app", "size": 9, "color": GRAY2, "font": F_MONO}])
    text(slide, Inches(12.3), Inches(7.08), Inches(0.5), Inches(0.3),
         [{"t": f"{n:02d}", "size": 9, "color": GRAY2, "font": F_MONO, "align": PP_ALIGN.RIGHT}])


def project_card(slide, x, y, w, h, title, desc, tag, featured=False, link=None):
    box(slide, Inches(x), Inches(y), Inches(w), Inches(h),
        line=BLUE if featured else LINE, lw=1.4 if featured else 1.0)
    ty = y + 0.22
    text(slide, Inches(x + 0.25), Inches(ty), Inches(w - 0.5), Inches(0.62),
         [{"t": title, "size": 13, "bold": True, "color": WHITE, "ls": 1.0}])
    text(slide, Inches(x + 0.25), Inches(ty + 0.62), Inches(w - 0.5), Inches(h - 1.5),
         [{"t": desc, "size": 10, "color": GRAY, "ls": 1.12}])
    tag_y = y + h - 0.48
    tw = 0.35 + 0.075 * len(tag)
    chip(slide, Inches(x + 0.25), Inches(tag_y), Inches(tw), Inches(0.3), tag,
         color=CYAN, size=9, font=F_MONO)
    if link:
        text(slide, Inches(x + 0.25 + tw + 0.15), Inches(tag_y), Inches(w - tw - 0.7), Inches(0.3),
             [{"t": link, "size": 9, "color": GRAY2, "font": F_MONO}], anchor=MSO_ANCHOR.MIDDLE)


# ---------- Content ----------

ID = {
    "out": "Refin_Izmi_Abiyuna_Portofolio_ID.pptx",
    "cover": {
        "kicker": "PORTOFOLIO PROFESIONAL // " + str(YEAR),
        "name1": "Refin Izmi ", "name2": "Abiyuna",
        "role": "Web Developer & Automation Specialist",
        "desc": "Saya bantu bisnis & UMKM membangun website fast-loading, landing page berkonversi tinggi, "
                "dan sistem internal otomatis yang menghemat waktu & biaya operasional.",
        "status": "Tersedia untuk Proyek Freelance",
        "contacts": ["refin555@gmail.com", "+62 822-4491-2640",
                     "linkedin.com/in/refin-izmi-abiyuna", "refinizmiabiyuna.netlify.app"],
        "loc": "Yogyakarta, Indonesia",
    },
    "about": {
        "kicker": "01 // PROFIL", "title": "Tentang Saya",
        "p1": "Lulusan Sistem Informasi dari Universitas Teknologi Yogyakarta, dengan pemahaman mendalam "
              "tentang integrasi teknologi dan strategi bisnis.",
        "p2": "Fokus pada pengembangan web, analisis sistem, dan otomasi bisnis — dari landing page "
              "berkonversi tinggi sampai sistem internal yang menghemat waktu operasional.",
        "edu_head": "Pendidikan",
        "edu1": "Universitas Teknologi Yogyakarta",
        "edu2": "Sistem Informasi  ·  2021 – 2025",
        "edu3": "IPK 3.44 / 4.00",
        "focus": "Fokus: Web Development  ·  Automasi Bisnis  ·  Sistem Informasi",
    },
    "exp": {
        "kicker": "02 // PENGALAMAN", "title": "Pengalaman Profesional",
        "items": [
            {"role": "Staff Business Development (IT Development)", "comp": "Zipzap Corporation",
             "date": "Mei 2026 – September 2026",
             "bullets": ["Lanjutan peran IT Developer dengan cakupan proyek lebih luas",
                         "Membangun Website Kaliber + Landing Page Zipzap & Kaliber Garment (Laravel, JS, Tailwind)",
                         "Maintenance, bug fixing & input laporan performa iklan harian tim marketing"]},
            {"role": "IT Developer", "comp": "Zipzap Corporation — Maganghub Kemnaker",
             "date": "Nov 2025 – Mei 2026",
             "bullets": ["Sistem absensi berbasis web real-time — hemat biaya ±Rp2.000.000/tahun",
                         "3 internal tools CS (kalkulator harga, sablon, stok kain) — respons < 1 menit",
                         "Pelaporan iklan berbasis Gemini AI — otomasi ekstraksi data screenshot"]},
            {"role": "Praktik Lapangan Tugas Akhir", "comp": "PT. Bubejo Makmur Sejahtera",
             "date": "Agu 2024 – Mei 2025",
             "bullets": ["Perancangan Sistem Informasi Penjualan berbasis web (PHP/Laravel + MySQL)",
                         "Integrasi API RajaOngkir & Midtrans (payment gateway)",
                         "Meningkatkan efisiensi sistem hingga 82%"]},
        ],
    },
    "stats": {
        "kicker": "03 // DAMPAK", "title": "Sorotan Pencapaian",
        "items": [
            {"v": "82%", "c": "Peningkatan efisiensi sistem di proyek tugas akhir"},
            {"v": "<1 mnt", "c": "Waktu respons CS dengan 3 kalkulator internal (dari 5+ menit)"},
            {"v": "±Rp2jt/thn", "c": "Hemat biaya operasional lewat sistem absensi mandiri"},
            {"v": "12+", "c": "Kalkulator harga & biaya produksi terpasang di portal internal"},
        ],
    },
    "services": {
        "kicker": "04 // LAYANAN", "title": "Layanan Saya",
        "sub": "Solusi digital yang disesuaikan dengan kebutuhan bisnis Anda.",
        "items": [
            {"icon": "</>", "title": "Landing Page & Website Profil Bisnis",
             "desc": "Website modern, cepat, dan responsif dari nol untuk menunjang kebutuhan pemasaran bisnis.",
             "points": ["Desain modern, responsive, & fast-loading",
                        "Fokus konversi dengan integrasi WhatsApp / form kontak",
                        "Clean code & SEO-friendly agar mudah ditemukan di Google"]},
            {"icon": "SYS", "title": "Custom Web App & Sistem Informasi",
             "desc": "Aplikasi web kustom dari nol yang dibangun presisi sesuai aturan dan kebutuhan unik bisnis.",
             "points": ["Custom business logic 100% sesuai alur bisnis",
                        "Database terpusat dengan hak akses berjenjang",
                        "Dashboard visual interaktif & laporan real-time"]},
            {"icon": "AUTO", "title": "Automasi Bisnis & Tools Produktivitas",
             "desc": "Tools produktivitas kustom untuk memangkas proses manual dan meningkatkan efisiensi kerja harian.",
             "points": ["Tools akselerasi tim — pangkas waktu operasional & CS",
                        "Portal mandiri pelanggan & form input terintegrasi",
                        "Otomatisasi notifikasi/email & rekap laporan harian"]},
        ],
    },
    "skills": {
        "kicker": "05 // STACK", "title": "Skill & Teknologi",
        "items": ["PHP", "HTML", "CSS", "JavaScript", "TypeScript", "React JS", "Laravel",
                  "Tailwind CSS", "Bootstrap CSS", "Git", "MySQL", "Google Sheets API",
                  "WordPress", "Google Apps Script", "Google Tag Manager",
                  "Google Workspace Mail API", "Gemini AI API", "Figma", "Canva"],
    },
    "proj1": {
        "kicker": "06 // PROYEK", "title": "Proyek Unggulan",
        "items": [
            {"t": "Sistem Informasi Absensi & Payroll", "d": "Sistem HR lengkap: absensi GPS + verifikasi selfie, "
                 "approval izin/lembur, payroll 1-klik, payslip PDF, dashboard eksekutif, 10 laporan Excel, "
                 "portal berbasis role dengan audit trail.", "tag": "Laravel 12 · React 19 · MySQL", "link": None},
            {"t": "Web Tools Customer Service Zipzap", "d": "Portal internal dua brand: katalog produk, size chart, "
                 "8 kalkulator harga produk & 4 kalkulator biaya printing, cek stok kain, design tools, "
                 "dan log pengunjung.", "tag": "Apps Script · Sheets · Drive", "link": "zipzap.co.id/tools"},
            {"t": "Cryon Warranty System", "d": "Registrasi klaim garansi online: intake multi-step, capture KTP "
                 "via kamera live, ID garansi otomatis, data ke Google Sheets + foto ke Drive, "
                 "konfirmasi email otomatis.", "tag": "Sheets · Drive · Email API", "link": None},
        ],
    },
    "proj2": {
        "kicker": "06 // PROYEK", "title": "Proyek Lainnya",
        "items": [
            {"t": "Website Korporat Konveksi Jakarta + Admin Panel", "d": "Company profile, katalog, blog, karier, "
                 "testimoni, FAQ, sitemap dinamis, 4 landing page SEO + admin panel lengkap (Laravel).",
             "tag": "Laravel · MySQL", "link": "konveksijakarta.co.id"},
            {"t": "Automasi Data Iklan dengan AI", "d": "Gemini AI mengekstrak metrik iklan dari screenshot → "
                 "review → Google Sheets, arsip Drive otomatis. Proxy PHP sisi server.",
             "tag": "Gemini AI · Apps Script", "link": None},
            {"t": "Tools Skoring DISC Zipzap", "d": "Tes kepribadian DISC: 24 grup pertanyaan Most/Least, "
                 "scoring engine, hasil chart bar & radar, ekspor PDF/Excel.",
             "tag": "JavaScript · Chart", "link": "zipzap.co.id/test-disc"},
            {"t": "Landing Page Zipzap Garment", "d": "Landing page konversi kaos event massal: showcase klien "
                 "(Adaro, Shopee, Garuda), harga, proses, alur WhatsApp. Static PHP + Tailwind.",
             "tag": "PHP · Tailwind CSS", "link": "konveksijakarta.co.id"},
            {"t": "Landing Page Kaliber Garment", "d": "Hub premium kaos & polo custom: material Cotton Combed "
                 "24s/30s, sablon & bordir, katalog harga, versi CMS Laravel + Filament.",
             "tag": "Laravel · Filament", "link": None},
            {"t": "Lihat Semua Proyek", "d": "Detail lengkap setiap proyek, screenshot, dan studi kasus tersedia "
                 "di website portofolio.", "tag": "refinizmiabiyuna.netlify.app", "link": None},
        ],
    },
    "certs": {
        "kicker": "07 // SERTIFIKASI", "title": "Sertifikasi",
        "items": ["Intro to Data Analytics", "Intro to Software Engineering", "Machine Learning Professional",
                  "Data Engineering Professional", "Intro to Product Management",
                  "Webinar Cybersecurity Vulnerability Assessment", "Exploring the World of Technology Industry"],
    },
    "contact": {
        "kicker": "08 // KONTAK", "title": "Mari Bekerja Sama",
        "desc": "Punya ide website, sistem internal, atau kebutuhan otomasi? Mari diskusikan — konsultasi awal gratis.",
        "cards": [
            {"h": "WhatsApp", "v": "+62 822-4491-2640"},
            {"h": "Email", "v": "refin555@gmail.com"},
            {"h": "LinkedIn", "v": "/refin-izmi-abiyuna"},
            {"h": "Website", "v": "refinizmiabiyuna.netlify.app"},
        ],
        "status": "Tersedia untuk proyek freelance & full-time",
        "loc": "Yogyakarta, Indonesia",
    },
}

EN = {
    "out": "Refin_Izmi_Abiyuna_Portfolio_EN.pptx",
    "cover": {
        "kicker": "PROFESSIONAL PORTFOLIO // " + str(YEAR),
        "name1": "Refin Izmi ", "name2": "Abiyuna",
        "role": "Web Developer & Automation Specialist",
        "desc": "I help businesses & SMEs build fast-loading websites, high-converting landing pages, "
                "and automated internal systems that save operational time & cost.",
        "status": "Available for Freelance Projects",
        "contacts": ["refin555@gmail.com", "+62 822-4491-2640",
                     "linkedin.com/in/refin-izmi-abiyuna", "refinizmiabiyuna.netlify.app"],
        "loc": "Yogyakarta, Indonesia",
    },
    "about": {
        "kicker": "01 // PROFILE", "title": "About Me",
        "p1": "Information Systems graduate from Yogyakarta University of Technology, with a deep understanding "
              "of technology integration and business strategy.",
        "p2": "Focused on web development, systems analysis, and business automation — from high-converting "
              "landing pages to internal systems that cut operational time.",
        "edu_head": "Education",
        "edu1": "Yogyakarta University of Technology",
        "edu2": "Information Systems  ·  2021 – 2025",
        "edu3": "GPA 3.44 / 4.00",
        "focus": "Focus: Web Development  ·  Business Automation  ·  Information Systems",
    },
    "exp": {
        "kicker": "02 // EXPERIENCE", "title": "Professional Experience",
        "items": [
            {"role": "Business Development Staff (IT Development)", "comp": "Zipzap Corporation",
             "date": "May 2026 – September 2026",
             "bullets": ["Continued IT Developer role with an expanded project scope",
                         "Built Kaliber Website + Zipzap & Kaliber Garment Landing Pages (Laravel, JS, Tailwind)",
                         "Maintenance, bug fixing & daily ad performance reporting for marketing"]},
            {"role": "IT Developer", "comp": "Zipzap Corporation — Maganghub Kemnaker",
             "date": "Nov 2025 – May 2026",
             "bullets": ["Real-time web-based attendance system — saving ±IDR 2,000,000/year",
                         "3 internal CS tools (price, printing, fabric stock) — response < 1 minute",
                         "Gemini AI ad reporting — automated data extraction from screenshots"]},
            {"role": "Final Project Field Internship", "comp": "PT. Bubejo Makmur Sejahtera",
             "date": "Aug 2024 – May 2025",
             "bullets": ["Designed a web-based Sales Information System (PHP/Laravel + MySQL)",
                         "Integrated RajaOngkir & Midtrans (payment gateway) APIs",
                         "Improved system efficiency by 82%"]},
        ],
    },
    "stats": {
        "kicker": "03 // IMPACT", "title": "Key Achievements",
        "items": [
            {"v": "82%", "c": "System efficiency improvement in final project"},
            {"v": "<1 min", "c": "CS response time with 3 internal calculators (from 5+ minutes)"},
            {"v": "±IDR 2M/yr", "c": "Operational cost savings via in-house attendance system"},
            {"v": "12+", "c": "Product & production cost calculators deployed in the internal portal"},
        ],
    },
    "services": {
        "kicker": "04 // SERVICES", "title": "My Services",
        "sub": "Digital solutions tailored to your business needs.",
        "items": [
            {"icon": "</>", "title": "Landing Page & Business Profile Website",
             "desc": "Modern, fast, and responsive websites built from scratch to drive your business marketing.",
             "points": ["Modern, responsive, & fast-loading design",
                        "Conversion-focused with WhatsApp / lead-capture integration",
                        "Clean code & SEO-friendly for search visibility"]},
            {"icon": "SYS", "title": "Custom Web App & Information Systems",
             "desc": "Custom web applications built from scratch, precisely aligned with your unique business rules.",
             "points": ["100% custom business logic for your operations",
                        "Centralized database with granular role-based access",
                        "Interactive visual dashboards & real-time reporting"]},
            {"icon": "AUTO", "title": "Business Automation & Productivity Tools",
             "desc": "Custom productivity tools designed to cut down manual tasks and boost daily work efficiency.",
             "points": ["Team acceleration tools — cut ops & CS response time",
                        "Customer self-service portals & integrated forms",
                        "Automated email/notification triggers & daily reports"]},
        ],
    },
    "skills": {
        "kicker": "05 // STACK", "title": "Skills & Technologies",
        "items": ID["skills"]["items"],
    },
    "proj1": {
        "kicker": "06 // PROJECTS", "title": "Featured Projects",
        "items": [
            {"t": "Attendance & Payroll Information System", "d": "Full HR system: GPS + selfie-verified attendance, "
                 "leave/overtime approval workflow, 1-click payroll, PDF payslips, executive dashboards, "
                 "10 Excel reports, role-based portals with audit trail.", "tag": "Laravel 12 · React 19 · MySQL",
             "link": None},
            {"t": "Web Tools Customer Service Zipzap", "d": "Internal portal for two brands: product e-catalogs, "
                 "size charts, 8 product price calculators & 4 printing cost calculators, fabric stock checker, "
                 "design tools, and visitor log.", "tag": "Apps Script · Sheets · Drive", "link": "zipzap.co.id/tools"},
            {"t": "Cryon Warranty System", "d": "Online warranty claim registration: multi-step intake, live camera "
                 "KTP capture, sequential warranty IDs, data to Google Sheets + photos to Drive, "
                 "automated email confirmations.", "tag": "Sheets · Drive · Email API", "link": None},
        ],
    },
    "proj2": {
        "kicker": "06 // PROJECTS", "title": "Other Projects",
        "items": [
            {"t": "Konveksi Jakarta Corporate Website + Admin Panel", "d": "Company profile, catalog, blog, careers, "
                 "testimonials, FAQs, dynamic sitemap, 4 SEO landing pages + full admin panel (Laravel).",
             "tag": "Laravel · MySQL", "link": "konveksijakarta.co.id"},
            {"t": "AI-Powered Ads Data Automation", "d": "Google Gemini AI extracts ad metrics from screenshots → "
                 "review → Google Sheets, auto-archived to Drive. Server-side PHP proxy.",
             "tag": "Gemini AI · Apps Script", "link": None},
            {"t": "DISC Scoring Tool Zipzap", "d": "DISC personality test: 24 Most/Least question groups, scoring "
                 "engine, animated bar & radar chart results, PDF/Excel export.",
             "tag": "JavaScript · Chart", "link": "zipzap.co.id/test-disc"},
            {"t": "Zipzap Garment Landing Page", "d": "Conversion-focused landing page for mass event t-shirts: "
                 "big-name clients (Adaro, Shopee, Garuda), pricing, process, WhatsApp funnel. Static PHP + Tailwind.",
             "tag": "PHP · Tailwind CSS", "link": "konveksijakarta.co.id"},
            {"t": "Kaliber Garment Landing Page", "d": "Premium custom t-shirt & polo hub: Cotton Combed 24s/30s "
                 "materials, printing & embroidery, priced catalog, Laravel + Filament CMS version.",
             "tag": "Laravel · Filament", "link": None},
            {"t": "See All Projects", "d": "Full details, screenshots, and case studies for every project are "
                 "available on the portfolio website.", "tag": "refinizmiabiyuna.netlify.app", "link": None},
        ],
    },
    "certs": {
        "kicker": "07 // CERTIFICATIONS", "title": "Certifications",
        "items": ID["certs"]["items"],
    },
    "contact": {
        "kicker": "08 // CONTACT", "title": "Let's Work Together",
        "desc": "Have an idea for a website, internal system, or automation? Let's discuss — the initial "
                "consultation is free.",
        "cards": [
            {"h": "WhatsApp", "v": "+62 822-4491-2640"},
            {"h": "Email", "v": "refin555@gmail.com"},
            {"h": "LinkedIn", "v": "/refin-izmi-abiyuna"},
            {"h": "Website", "v": "refinizmiabiyuna.netlify.app"},
        ],
        "status": "Available for freelance projects & full-time roles",
        "loc": "Yogyakarta, Indonesia",
    },
}


# ---------- Build ----------

def build(L):
    prs = Presentation()
    prs.slide_width, prs.slide_height = SW, SH
    prs.core_properties.author = "Refin Izmi Abiyuna"
    prs.core_properties.title = L["cover"]["role"] + " — " + (
        "Portofolio" if L["out"].endswith("_ID.pptx") else "Portfolio")

    n = 0

    # ----- 1. Cover -----
    s = new_slide(prs); n += 1
    c = L["cover"]
    # decorative faint frame lines
    for i, col in enumerate((CYAN_D, FUCHSIA_D)):
        ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.55 + 0.07 * i), Inches(0.55 + 0.07 * i),
                                Inches(12.23), Inches(6.4))
        ln.fill.background()
        ln.line.color.rgb = col
        ln.line.width = Pt(0.75)
        ln.shadow.inherit = False
    text(s, Inches(0.95), Inches(1.05), Inches(7.2), Inches(0.35),
         [{"t": c["kicker"], "size": 13, "color": CYAN, "font": F_MONO, "bold": True}])
    text(s, Inches(0.95), Inches(1.5), Inches(7.6), Inches(1.4),
         [{"t": [(c["name1"], {"color": FUCHSIA}), (c["name2"], {"color": CYAN})],
           "size": 45, "bold": True}])
    text(s, Inches(0.95), Inches(2.55), Inches(7.2), Inches(0.45),
         [{"t": c["role"], "size": 19, "bold": True, "color": WHITE}])
    grad_bar(s, Inches(0.98), Inches(3.15), Inches(2.4), Inches(0.05))
    text(s, Inches(0.95), Inches(3.45), Inches(6.9), Inches(1.1),
         [{"t": c["desc"], "size": 12.5, "color": GRAY, "ls": 1.25}])
    text(s, Inches(0.95), Inches(4.6), Inches(6.9), Inches(0.35),
         [{"t": [("●  ", {"color": GREEN}), (c["status"], {"color": GREEN})], "size": 12, "bold": True}])
    text(s, Inches(0.95), Inches(5.15), Inches(7.0), Inches(1.3),
         [{"t": [("→  ", {"color": CYAN}), (x, {"color": GRAY})], "size": 10.5, "font": F_MONO, "sa": 4}
          for x in c["contacts"]])
    text(s, Inches(0.95), Inches(6.55), Inches(6), Inches(0.3),
         [{"t": c["loc"] + f"  ·  {YEAR}", "size": 10, "color": GRAY2, "font": F_MONO}])
    # photo with offset accent cards
    if os.path.exists(PHOTO):
        ph = s.shapes.add_picture(PHOTO, Inches(9.2), Inches(1.75), height=Inches(3.9))
        ph.left = Inches(12.55) - ph.width
        ph.top = Inches(1.85)
        for dx, dy, col in ((0.14, 0.14, CYAN_D), (-0.14, -0.14, FUCHSIA_D)):
            fr = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, ph.left + Inches(dx), ph.top + Inches(dy),
                                    ph.width, ph.height)
            fr.fill.background()
            fr.line.color.rgb = col
            fr.line.width = Pt(1.2)
            fr.shadow.inherit = False

    # ----- 2. About -----
    s = new_slide(prs); n += 1
    a = L["about"]
    header(s, a["kicker"], a["title"])
    text(s, Inches(0.6), Inches(1.85), Inches(7.1), Inches(2.6),
         [{"t": a["p1"], "size": 13.5, "color": GRAY, "ls": 1.3, "sa": 12},
          {"t": a["p2"], "size": 13.5, "color": GRAY, "ls": 1.3}])
    text(s, Inches(0.6), Inches(4.7), Inches(7.1), Inches(0.4),
         [{"t": a["focus"], "size": 11, "color": CYAN, "font": F_MONO}])
    box(s, Inches(8.1), Inches(1.85), Inches(4.6), Inches(3.6), line=LINE)
    text(s, Inches(8.45), Inches(2.2), Inches(3.9), Inches(3.0),
         [{"t": a["edu_head"].upper(), "size": 11, "color": CYAN, "font": F_MONO, "bold": True, "sa": 14},
          {"t": a["edu1"], "size": 15, "bold": True, "color": WHITE, "sa": 4},
          {"t": a["edu2"], "size": 12, "color": GRAY, "sa": 10},
          {"t": a["edu3"], "size": 13, "bold": True, "color": FUCHSIA}])
    grad_bar(s, Inches(8.45), Inches(5.05), Inches(1.5), Inches(0.04))
    footer(s, n)

    # ----- 3. Experience -----
    s = new_slide(prs); n += 1
    e = L["exp"]
    header(s, e["kicker"], e["title"])
    y = 1.75
    for i, it in enumerate(e["items"]):
        box(s, Inches(0.6), Inches(y), Inches(12.13), Inches(1.68), line=LINE)
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(y + 0.15), Inches(0.06), Inches(1.38))
        bar.fill.solid(); bar.fill.fore_color.rgb = [FUCHSIA, CYAN, BLUE][i % 3]
        bar.line.fill.background(); bar.shadow.inherit = False
        text(s, Inches(0.95), Inches(y + 0.18), Inches(8.6), Inches(0.32),
             [{"t": it["role"], "size": 14, "bold": True, "color": WHITE}])
        text(s, Inches(0.95), Inches(y + 0.52), Inches(8.6), Inches(0.28),
             [{"t": it["comp"], "size": 11.5, "color": CYAN}])
        text(s, Inches(0.95), Inches(y + 0.85), Inches(11.5), Inches(0.78),
             [{"t": [("•  ", {"color": FUCHSIA if i == 0 else CYAN}), (b, {})], "size": 10.5, "color": GRAY, "sa": 2}
              for b in it["bullets"]])
        chip(s, Inches(10.45), Inches(y + 0.18), Inches(2.05), Inches(0.34), it["date"],
             color=CYAN, size=9, font=F_MONO)
        y += 1.82
    footer(s, n)

    # ----- 4. Stats -----
    s = new_slide(prs); n += 1
    st = L["stats"]
    header(s, st["kicker"], st["title"])
    x = 0.5
    for i, it in enumerate(st["items"]):
        box(s, Inches(x), Inches(2.3), Inches(2.95), Inches(2.5), line=LINE)
        text(s, Inches(x + 0.25), Inches(2.75), Inches(2.45), Inches(0.8),
             [{"t": it["v"], "size": 30, "bold": True, "color": CYAN if i % 2 == 0 else FUCHSIA}])
        grad_bar(s, Inches(x + 0.27), Inches(3.55), Inches(0.9), Inches(0.04))
        text(s, Inches(x + 0.25), Inches(3.75), Inches(2.45), Inches(0.95),
             [{"t": it["c"], "size": 10.5, "color": GRAY, "ls": 1.15}])
        x += 3.13
    footer(s, n)

    # ----- 5. Services -----
    s = new_slide(prs); n += 1
    sv = L["services"]
    header(s, sv["kicker"], sv["title"])
    text(s, Inches(0.6), Inches(1.5), Inches(12), Inches(0.3),
         [{"t": sv["sub"], "size": 12, "color": GRAY}])
    x = 0.27
    for i, it in enumerate(sv["items"]):
        accent = [BLUE, FUCHSIA, CYAN][i]
        box(s, Inches(x), Inches(2.0), Inches(4.1), Inches(4.5), line=LINE)
        ic = box(s, Inches(x + 0.3), Inches(2.3), Inches(0.75), Inches(0.75), fill=RGBColor(0x0B, 0x12, 0x20),
                 line=accent, radius=0.25)
        text(s, Inches(x + 0.3), Inches(2.3), Inches(0.75), Inches(0.75),
             [{"t": it["icon"], "size": 13, "bold": True, "color": accent, "font": F_MONO,
               "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(x + 0.3), Inches(3.25), Inches(3.5), Inches(0.75),
             [{"t": it["title"], "size": 13.5, "bold": True, "color": WHITE, "ls": 1.05}])
        text(s, Inches(x + 0.3), Inches(4.0), Inches(3.5), Inches(0.8),
             [{"t": it["desc"], "size": 10, "color": GRAY, "ls": 1.15}])
        text(s, Inches(x + 0.3), Inches(4.85), Inches(3.5), Inches(1.5),
             [{"t": [("✓  ", {"color": accent}), (p, {})], "size": 10, "color": GRAY, "ls": 1.1, "sa": 5}
              for p in it["points"]])
        x += 4.35
    footer(s, n)

    # ----- 6. Skills -----
    s = new_slide(prs); n += 1
    sk = L["skills"]
    header(s, sk["kicker"], sk["title"])
    chips_flow(s, sk["items"], 0.57, 2.3, 12.2, gap_x=0.16, gap_y=0.18, h=0.5, size=11.5)
    text(s, Inches(0.6), Inches(6.3), Inches(12.1), Inches(0.3),
         [{"t": f"// {len(sk['items'])} tools & technologies", "size": 10.5, "color": GRAY2, "font": F_MONO,
           "align": PP_ALIGN.CENTER}])
    footer(s, n)

    # ----- 7. Featured projects -----
    s = new_slide(prs); n += 1
    p1 = L["proj1"]
    header(s, p1["kicker"], p1["title"])
    x = 0.27
    for it in p1["items"]:
        project_card(s, x, 1.9, 4.1, 4.35, it["t"], it["d"], it["tag"], featured=True, link=it["link"])
        x += 4.35
    footer(s, n)

    # ----- 8. Other projects -----
    s = new_slide(prs); n += 1
    p2 = L["proj2"]
    header(s, p2["kicker"], p2["title"])
    pos = [(0.27, 1.8), (4.62, 1.8), (8.97, 1.8), (0.27, 4.35), (4.62, 4.35), (8.97, 4.35)]
    for it, (x, y) in zip(p2["items"], pos):
        project_card(s, x, y, 4.1, 2.4, it["t"], it["d"], it["tag"], featured=False, link=it["link"])
    footer(s, n)

    # ----- 9. Certifications -----
    s = new_slide(prs); n += 1
    ct = L["certs"]
    header(s, ct["kicker"], ct["title"])
    bottom = chips_flow(s, ct["items"], 0.57, 2.6, 12.2, gap_x=0.2, gap_y=0.2, h=0.75, size=11.5,
                        color=FUCHSIA, font=F_SANS, char_w=0.11, pad=0.5)
    text(s, Inches(0.6), bottom + 0.35, Inches(12.1), Inches(0.3),
         [{"t": f"// {len(ct['items'])} certifications", "size": 10.5, "color": GRAY2, "font": F_MONO,
           "align": PP_ALIGN.CENTER}])
    footer(s, n)

    # ----- 10. Contact -----
    s = new_slide(prs); n += 1
    co = L["contact"]
    header(s, co["kicker"], co["title"])
    text(s, Inches(0.6), Inches(1.9), Inches(12.1), Inches(0.6),
         [{"t": co["desc"], "size": 14, "color": GRAY, "ls": 1.25}])
    x = 0.5
    for it in co["cards"]:
        box(s, Inches(x), Inches(2.9), Inches(2.95), Inches(1.7), line=LINE)
        text(s, Inches(x + 0.25), Inches(3.15), Inches(2.45), Inches(0.3),
             [{"t": it["h"].upper(), "size": 10, "color": CYAN, "font": F_MONO, "bold": True}])
        text(s, Inches(x + 0.25), Inches(3.5), Inches(2.45), Inches(0.9),
             [{"t": it["v"], "size": 12, "bold": True, "color": WHITE, "ls": 1.1}])
        x += 3.13
    text(s, Inches(0.6), Inches(5.15), Inches(12.1), Inches(0.4),
         [{"t": [("●  ", {"color": GREEN}), (co["status"], {"color": GREEN})], "size": 13, "bold": True,
           "align": PP_ALIGN.CENTER}])
    text(s, Inches(0.6), Inches(5.65), Inches(12.1), Inches(0.3),
         [{"t": co["loc"] + f"  ·  {YEAR}", "size": 11, "color": GRAY2, "font": F_MONO,
           "align": PP_ALIGN.CENTER}])
    footer(s, n)

    out = os.path.join(BASE, L["out"])
    prs.save(out)
    return out, n


if __name__ == "__main__":
    for L in (ID, EN):
        path, count = build(L)
        print(f"OK  {count} slides  ->  {path}")
