# -*- coding: utf-8 -*-
"""
Deck portofolio versi LAMARAN KERJA (ID + EN).
Fokus: profil profesional, pendidikan, pengalaman kerja (bullets lengkap),
pencapaian terukur, skill. Tone formal & to the point untuk recruiter/HR.
Style: sama dengan deck general (dark/tech).

Output: Refin_Izmi_Abiyuna_Portofolio_Lamaran_ID.pptx &
        Refin_Izmi_Abiyuna_Portfolio_Job_Application_EN.pptx

Jalankan: python slides/job_application.py
"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

# Reuse helpers & palette dari deck general
from generate import (  # noqa: E402
    BASE as ROOT, PHOTO, YEAR,
    BG, CARD, LINE, WHITE, GRAY, GRAY2, FUCHSIA, FUCHSIA_D, CYAN, CYAN_D, BLUE, GREEN,
    F_SANS, F_MONO, SW, SH,
    _lerp, box, grad_bar, text, chip, chips_flow, new_slide, header, footer,
)
from pptx import Presentation  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402


# ---------- Content ----------

ID = {
    "out": "Refin_Izmi_Abiyuna_Portofolio_Lamaran_ID.pptx",
    "cover": {
        "kicker": "PORTOFOLIO // LAMARAN KERJA // " + str(YEAR),
        "name1": "Refin Izmi ", "name2": "Abiyuna",
        "role": "Web Developer & Automation Specialist",
        "desc": "Lulusan Sistem Informasi (IPK 3.44) dengan pengalaman mengembangkan sistem internal "
                "berbasis web yang terbukti menghemat biaya dan waktu operasional perusahaan.",
        "status": "Terbuka untuk kesempatan full-time",
        "contacts": ["refin555@gmail.com", "+62 822-4491-2640",
                     "linkedin.com/in/refin-izmi-abiyuna", "refinizmiabiyuna.netlify.app"],
        "loc": "Yogyakarta, Indonesia",
    },
    "about": {
        "kicker": "01 // PROFIL", "title": "Tentang Saya",
        "p1": "Lulusan Sistem Informasi dari Universitas Teknologi Yogyakarta dengan pemahaman mendalam "
              "tentang integrasi teknologi dan strategi bisnis.",
        "p2": "Berpengalaman sebagai IT Developer: membangun website perusahaan, sistem HR internal, "
              "dan otomasi berbasis AI yang memberikan dampak terukur bagi operasional bisnis.",
        "focus": "Fokus: Web Development  ·  Sistem Informasi  ·  Automasi Bisnis",
        "edu_head": "RINGKASAN",
        "facts": [
            ("Pendidikan", "S1 Sistem Informasi — UTY, 2021–2025"),
            ("IPK", "3.44 / 4.00"),
            ("Pengalaman", "±1,5 tahun (magang → staff)"),
            ("Lokasi", "Yogyakarta — siap onsite/hybrid/remote"),
        ],
    },
    "edu": {
        "kicker": "02 // PENDIDIKAN", "title": "Pendidikan & Sertifikasi",
        "edu1": "Universitas Teknologi Yogyakarta",
        "edu2": "S1 Sistem Informasi  ·  2021 – 2025",
        "edu3": "IPK 3.44 / 4.00",
        "cert_head": "SERTIFIKASI",
        "certs": ["Intro to Data Analytics", "Intro to Software Engineering",
                  "Machine Learning Professional", "Data Engineering Professional",
                  "Intro to Product Management", "Webinar Cybersecurity Vulnerability Assessment",
                  "Exploring the World of Technology Industry"],
    },
    "exp": {
        "kicker": "03 // PENGALAMAN", "title": "Pengalaman Profesional",
        "org": "Zipzap Corporation",
        "note": "Magang (Maganghub Kemnaker) Nov 2025 – Mei 2026, dilanjutkan sebagai Staff "
                "Business Development (IT Development) Mei 2026 – September 2026.",
        "blocks": [
            {"role": "Staff Business Development (IT Development)", "date": "Mei 2026 – September 2026",
             "bullets": [
                 "Melanjutkan tanggung jawab pengembangan sistem dari posisi sebelumnya dengan cakupan proyek yang lebih luas.",
                 "Membangun Website Kaliber, Landing Page Zipzap Garment, dan Landing Page Kaliber Garment menggunakan Laravel, JavaScript, dan Tailwind CSS.",
                 "Menangani maintenance, bug fixing, dan input laporan performa iklan harian untuk mendukung tim marketing.",
                 "Menyusun laporan progres mingguan sebagai bentuk koordinasi lintas tim.",
             ]},
            {"role": "IT Developer — Maganghub Kemnaker", "date": "Nov 2025 – Mei 2026",
             "bullets": [
                 "Mengembangkan dan memelihara sistem absensi karyawan berbasis web untuk memantau kehadiran secara real-time, menggantikan layanan pihak ketiga dan menghemat biaya operasional hingga ±Rp2.000.000/tahun.",
                 "Merancang dan membangun 3 internal tools tim Customer Service (kalkulator harga, kalkulator sablon, cek stok kain) yang mempercepat waktu respons ke pelanggan dari 5+ menit menjadi kurang dari 1 menit.",
                 "Membangun sistem pelaporan iklan berbasis Gemini AI yang mengotomasi ekstraksi data dari screenshot laporan harian, menggantikan proses input manual tim marketing.",
                 "Mengelola dan mengoptimalkan landing page bisnis untuk mendukung kampanye pemasaran digital perusahaan.",
                 "Melakukan maintenance, bug fixing, dan optimasi performa sistem secara berkala.",
             ]},
        ],
    },
    "exp2": {
        "kicker": "03 // PENGALAMAN", "title": "Pengalaman Profesional",
        "role": "Praktik Lapangan Tugas Akhir",
        "comp": "PT. Bubejo Makmur Sejahtera",
        "date": "Agu 2024 – Mei 2025",
        "intro": "Merancang dan membangun Sistem Informasi Penjualan berbasis web dari analisis "
                 "kebutuhan hingga implementasi full-stack.",
        "bullets": [
            "Melakukan analisis kebutuhan sistem bisnis untuk merancang solusi Sistem Informasi Penjualan berbasis web.",
            "Mengembangkan full-stack sistem menggunakan PHP/Laravel dan MySQL untuk manajemen produk, pesanan, dan laporan penjualan.",
            "Integrasi Third-Party API RajaOngkir (ongkir) dan Midtrans (payment gateway).",
            "Berhasil meningkatkan efisiensi sistem sebesar 82%.",
        ],
    },
    "stats": {
        "kicker": "04 // DAMPAK", "title": "Pencapaian Terukur",
        "items": [
            {"v": "82%", "c": "Peningkatan efisiensi sistem informasi penjualan di PT. Bubejo Makmur Sejahtera"},
            {"v": "<1 mnt", "c": "Waktu respons CS dengan 3 kalkulator internal (sebelumnya 5+ menit)"},
            {"v": "±Rp2jt/thn", "c": "Hemat biaya operasional lewat sistem absensi mandiri"},
            {"v": "8+", "c": "Sistem & tools yang dibangun: website korporat, HR system, portal internal, dan lainnya"},
        ],
    },
    "skills": {
        "kicker": "05 // STACK", "title": "Skill & Teknologi",
        "items": ["PHP", "HTML", "CSS", "JavaScript", "TypeScript", "React JS", "Laravel",
                  "Tailwind CSS", "Bootstrap CSS", "Git", "MySQL", "Google Sheets API",
                  "WordPress", "Google Apps Script", "Google Tag Manager",
                  "Google Workspace Mail API", "Gemini AI API", "Figma", "Canva"],
    },
    "proj": {
        "kicker": "06 // PROYEK", "title": "Proyek Pilihan",
        "items": [
            {"t": "Sistem Informasi Absensi & Payroll", "d": "Sistem HR lengkap: absensi GPS + selfie, "
                 "workflow approval, payroll 1-klik, payslip PDF, 10 laporan Excel, portal berbasis role "
                 "dengan audit trail.", "tag": "Laravel 12 · React 19 · MySQL"},
            {"t": "Website Korporat Konveksi Jakarta", "d": "Company profile, katalog, blog, karier, 4 landing "
                 "page SEO + admin panel lengkap dengan analitik pengunjung & user roles.",
             "tag": "Laravel · MySQL"},
            {"t": "Web Tools Customer Service Zipzap", "d": "Portal internal dua brand: 8 kalkulator harga, "
                 "4 kalkulator printing, cek stok kain, dan log pengunjung. Ditinjau demi NDA.",
             "tag": "Apps Script · Sheets"},
            {"t": "Cryon Warranty System", "d": "Registrasi klaim garansi online: capture KTP via kamera, "
                 "ID garansi otomatis, data ke Google Sheets + foto ke Drive, konfirmasi email otomatis.",
             "tag": "Sheets · Drive · Email"},
            {"t": "Automasi Data Iklan dengan AI", "d": "Gemini AI mengekstrak metrik iklan dari screenshot "
                 "→ Google Sheets, arsip Drive otomatis, API key aman di sisi server.",
             "tag": "Gemini AI · Apps Script"},
            {"t": "Tools Skoring DISC Zipzap", "d": "Tes kepribadian DISC dengan scoring engine, chart bar & "
                 "radar, ekspor PDF/Excel untuk screening kandidat HR.",
             "tag": "JavaScript · Chart"},
        ],
    },
    "closing": {
        "kicker": "07 // KONTAK", "title": "Terima Kasih",
        "desc": "Saya terbuka untuk diskusi lebih lanjut mengenai kesempatan berkontribusi di tim Anda. "
                "Dokumen CV lengkap tersedia dan dapat dikirimkan sesuai permintaan.",
        "cards": [
            {"h": "WhatsApp", "v": "+62 822-4491-2640"},
            {"h": "Email", "v": "refin555@gmail.com"},
            {"h": "LinkedIn", "v": "/refin-izmi-abiyuna"},
            {"h": "Website", "v": "refinizmiabiyuna.netlify.app"},
        ],
        "loc": "Yogyakarta, Indonesia",
    },
}

EN = {
    "out": "Refin_Izmi_Abiyuna_Portfolio_Job_Application_EN.pptx",
    "cover": {
        "kicker": "PORTFOLIO // JOB APPLICATION // " + str(YEAR),
        "name1": "Refin Izmi ", "name2": "Abiyuna",
        "role": "Web Developer & Automation Specialist",
        "desc": "Information Systems graduate (GPA 3.44) with hands-on experience building internal "
                "web systems that measurably cut company operational time & cost.",
        "status": "Open to full-time opportunities",
        "contacts": ["refin555@gmail.com", "+62 822-4491-2640",
                     "linkedin.com/in/refin-izmi-abiyuna", "refinizmiabiyuna.netlify.app"],
        "loc": "Yogyakarta, Indonesia",
    },
    "about": {
        "kicker": "01 // PROFILE", "title": "About Me",
        "p1": "Information Systems graduate from Yogyakarta University of Technology with a deep "
              "understanding of technology integration and business strategy.",
        "p2": "Experienced IT Developer: built corporate websites, an internal HR system, and AI-powered "
              "automation that delivered measurable business impact.",
        "focus": "Focus: Web Development  ·  Information Systems  ·  Business Automation",
        "edu_head": "SNAPSHOT",
        "facts": [
            ("Education", "BSc Information Systems — UTY, 2021–2025"),
            ("GPA", "3.44 / 4.00"),
            ("Experience", "±1.5 years (intern → staff)"),
            ("Location", "Yogyakarta — onsite/hybrid/remote ready"),
        ],
    },
    "edu": {
        "kicker": "02 // EDUCATION", "title": "Education & Certifications",
        "edu1": "Yogyakarta University of Technology",
        "edu2": "BSc Information Systems  ·  2021 – 2025",
        "edu3": "GPA 3.44 / 4.00",
        "cert_head": "CERTIFICATIONS",
        "certs": ["Intro to Data Analytics", "Intro to Software Engineering",
                  "Machine Learning Professional", "Data Engineering Professional",
                  "Intro to Product Management", "Webinar Cybersecurity Vulnerability Assessment",
                  "Exploring the World of Technology Industry"],
    },
    "exp": {
        "kicker": "03 // EXPERIENCE", "title": "Professional Experience",
        "org": "Zipzap Corporation",
        "note": "Intern (Maganghub Kemnaker) Nov 2025 – May 2026, continued as Business Development "
                "Staff (IT Development) May 2026 – September 2026.",
        "blocks": [
            {"role": "Business Development Staff (IT Development)", "date": "May 2026 – September 2026",
             "bullets": [
                 "Continued software development responsibilities from the previous role with an expanded project scope.",
                 "Built the Kaliber Website, Zipzap Garment Landing Page, and Kaliber Garment Landing Page using Laravel, JavaScript, and Tailwind CSS.",
                 "Managed maintenance, bug fixes, and daily ad performance data entry to support the marketing team.",
                 "Compiled weekly progress reports to maintain cross-team alignment.",
             ]},
            {"role": "IT Developer — Maganghub Kemnaker", "date": "Nov 2025 – May 2026",
             "bullets": [
                 "Developed and maintained a web-based employee attendance system for real-time tracking, replacing third-party services and saving up to ±IDR 2,000,000/year in operational costs.",
                 "Designed and built 3 internal Customer Service tools (price calculator, screen-printing calculator, fabric stock checker) that reduced customer response time from 5+ minutes to under 1 minute.",
                 "Built a Gemini AI-powered ad reporting system to automate data extraction from daily report screenshots, replacing manual marketing data entry.",
                 "Managed and optimized business landing pages to support the company's digital marketing campaigns.",
                 "Performed routine maintenance, bug fixes, and system performance optimizations.",
             ]},
        ],
    },
    "exp2": {
        "kicker": "03 // EXPERIENCE", "title": "Professional Experience",
        "role": "Final Project Field Internship",
        "comp": "PT. Bubejo Makmur Sejahtera",
        "date": "Aug 2024 – May 2025",
        "intro": "Designed and built a web-based Sales Information System from requirements analysis "
                 "through full-stack implementation.",
        "bullets": [
            "Conducted business system requirements analysis to design a web-based Sales Information System solution.",
            "Developed a full-stack system using PHP/Laravel and MySQL for product management, orders, and sales reports.",
            "Integrated third-party APIs: RajaOngkir (shipping costs) and Midtrans (payment gateway).",
            "Successfully improved system efficiency by 82%.",
        ],
    },
    "stats": {
        "kicker": "04 // IMPACT", "title": "Measurable Achievements",
        "items": [
            {"v": "82%", "c": "Efficiency improvement of the sales information system at PT. Bubejo Makmur Sejahtera"},
            {"v": "<1 min", "c": "CS response time with 3 internal calculators (previously 5+ minutes)"},
            {"v": "±IDR 2M/yr", "c": "Operational cost savings via the in-house attendance system"},
            {"v": "8+", "c": "Systems & tools built: corporate websites, HR system, internal portals, and more"},
        ],
    },
    "skills": {
        "kicker": "05 // STACK", "title": "Skills & Technologies",
        "items": ID["skills"]["items"],
    },
    "proj": {
        "kicker": "06 // PROJECTS", "title": "Selected Projects",
        "items": [
            {"t": "Attendance & Payroll Information System", "d": "Full HR system: GPS + selfie attendance, "
                 "approval workflows, 1-click payroll, PDF payslips, 10 Excel reports, role-based portals "
                 "with audit trail.", "tag": "Laravel 12 · React 19 · MySQL"},
            {"t": "Konveksi Jakarta Corporate Website", "d": "Company profile, catalog, blog, careers, 4 SEO "
                 "landing pages + full admin panel with visitor analytics & user roles.",
             "tag": "Laravel · MySQL"},
            {"t": "Web Tools Customer Service Zipzap", "d": "Internal portal for two brands: 8 price calculators, "
                 "4 printing cost calculators, fabric stock checker, and visitor log. Withheld due to NDA.",
             "tag": "Apps Script · Sheets"},
            {"t": "Cryon Warranty System", "d": "Online warranty claim registration: live camera ID capture, "
                 "sequential warranty IDs, data to Google Sheets + photos to Drive, automated email confirmations.",
             "tag": "Sheets · Drive · Email"},
            {"t": "AI-Powered Ads Data Automation", "d": "Gemini AI extracts ad metrics from screenshots → "
                 "Google Sheets, auto-archived to Drive, API key kept server-side.",
             "tag": "Gemini AI · Apps Script"},
            {"t": "DISC Scoring Tool Zipzap", "d": "DISC personality test with a scoring engine, bar & radar "
                 "charts, and PDF/Excel export for HR candidate screening.",
             "tag": "JavaScript · Chart"},
        ],
    },
    "closing": {
        "kicker": "07 // CONTACT", "title": "Thank You",
        "desc": "I would welcome the opportunity to discuss how I can contribute to your team. "
                "A complete CV is available upon request.",
        "cards": [
            {"h": "WhatsApp", "v": "+62 822-4491-2640"},
            {"h": "Email", "v": "refin555@gmail.com"},
            {"h": "LinkedIn", "v": "/refin-izmi-abiyuna"},
            {"h": "Website", "v": "refinizmiabiyuna.netlify.app"},
        ],
        "loc": "Yogyakarta, Indonesia",
    },
}


# ---------- Build ----------

def build(L):
    prs = Presentation()
    prs.slide_width, prs.slide_height = SW, SH
    is_id = L["out"].endswith("_ID.pptx")
    prs.core_properties.author = "Refin Izmi Abiyuna"
    prs.core_properties.title = ("Portofolio Lamaran Kerja" if is_id else "Job Application Portfolio") \
        + " — Refin Izmi Abiyuna"

    n = 0

    # ----- 1. Cover -----
    s = new_slide(prs); n += 1
    c = L["cover"]
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
    if os.path.exists(PHOTO):
        ph = s.shapes.add_picture(PHOTO, Inches(9.2), Inches(1.85), height=Inches(3.9))
        ph.left = Inches(12.55) - ph.width
        ph.top = Inches(1.85)
        for dx, dy, col in ((0.14, 0.14, CYAN_D), (-0.14, -0.14, FUCHSIA_D)):
            fr = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, ph.left + Inches(dx), ph.top + Inches(dy),
                                    ph.width, ph.height)
            fr.fill.background()
            fr.line.color.rgb = col
            fr.line.width = Pt(1.2)
            fr.shadow.inherit = False

    # ----- 2. Profile -----
    s = new_slide(prs); n += 1
    a = L["about"]
    header(s, a["kicker"], a["title"])
    text(s, Inches(0.6), Inches(1.85), Inches(7.1), Inches(2.6),
         [{"t": a["p1"], "size": 13.5, "color": GRAY, "ls": 1.3, "sa": 12},
          {"t": a["p2"], "size": 13.5, "color": GRAY, "ls": 1.3}])
    text(s, Inches(0.6), Inches(4.7), Inches(7.1), Inches(0.4),
         [{"t": a["focus"], "size": 11, "color": CYAN, "font": F_MONO}])
    box(s, Inches(8.1), Inches(1.85), Inches(4.6), Inches(3.6), line=LINE)
    fy = 2.2
    text(s, Inches(8.45), Inches(fy), Inches(3.9), Inches(0.3),
         [{"t": a["edu_head"], "size": 11, "color": CYAN, "font": F_MONO, "bold": True}])
    fy += 0.5
    for label, val in a["facts"]:
        text(s, Inches(8.45), Inches(fy), Inches(3.9), Inches(0.62),
             [{"t": label.upper(), "size": 9, "color": GRAY2, "font": F_MONO, "sa": 1},
              {"t": val, "size": 11.5, "bold": True, "color": WHITE}])
        fy += 0.75
    footer(s, n)

    # ----- 3. Education & Certifications -----
    s = new_slide(prs); n += 1
    e = L["edu"]
    header(s, e["kicker"], e["title"])
    box(s, Inches(0.6), Inches(1.9), Inches(4.6), Inches(4.4), line=LINE)
    text(s, Inches(0.95), Inches(2.25), Inches(3.9), Inches(0.3),
         [{"t": e["kicker"].split("//")[0].strip() + " //", "size": 11, "color": CYAN,
           "font": F_MONO, "bold": True}])
    text(s, Inches(0.95), Inches(2.75), Inches(3.9), Inches(0.9),
         [{"t": e["edu1"], "size": 16, "bold": True, "color": WHITE, "ls": 1.1}])
    text(s, Inches(0.95), Inches(3.65), Inches(3.9), Inches(0.35),
         [{"t": e["edu2"], "size": 12.5, "color": GRAY}])
    text(s, Inches(0.95), Inches(4.15), Inches(3.9), Inches(0.4),
         [{"t": e["edu3"], "size": 14, "bold": True, "color": FUCHSIA}])
    grad_bar(s, Inches(0.97), Inches(4.75), Inches(1.5), Inches(0.04))
    text(s, Inches(0.95), Inches(5.0), Inches(3.9), Inches(1.0),
         [{"t": [("→  ", {"color": CYAN}), ("CV: assets/document/CV_Refin_Izmi_Abiyuna_ID.pdf",
                                           {"color": GRAY})], "size": 10, "font": F_MONO}])
    text(s, Inches(5.7), Inches(1.95), Inches(7.0), Inches(0.3),
         [{"t": e["cert_head"], "size": 11, "color": CYAN, "font": F_MONO, "bold": True}])
    cy = 2.45
    for i, cert in enumerate(e["certs"]):
        box(s, Inches(5.7), Inches(cy), Inches(7.0), Inches(0.52), fill=RGBColor(0x0B, 0x12, 0x20), line=LINE)
        text(s, Inches(5.95), Inches(cy), Inches(0.5), Inches(0.52),
             [{"t": f"{i + 1:02d}", "size": 11, "color": FUCHSIA, "font": F_MONO, "bold": True}],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(6.55), Inches(cy), Inches(6.0), Inches(0.52),
             [{"t": cert, "size": 11.5, "color": WHITE}], anchor=MSO_ANCHOR.MIDDLE)
        cy += 0.62
    footer(s, n)

    # ----- 4. Experience: Zipzap -----
    s = new_slide(prs); n += 1
    x1 = L["exp"]
    header(s, x1["kicker"], x1["title"])
    text(s, Inches(0.6), Inches(1.55), Inches(12.1), Inches(0.35),
         [{"t": [("▣  ", {"color": FUCHSIA}), (x1["org"], {"color": WHITE, "bold": True, "size": 15})],
           "size": 15}])
    text(s, Inches(0.6), Inches(1.95), Inches(12.1), Inches(0.35),
         [{"t": x1["note"], "size": 10.5, "color": GRAY, "ls": 1.15}])
    y = 2.5
    for bi, blk in enumerate(x1["blocks"]):
        bh = 0.55 + 0.34 * len(blk["bullets"])
        box(s, Inches(0.6), Inches(y), Inches(12.13), Inches(bh), line=LINE)
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(y + 0.12), Inches(0.06),
                                 Inches(bh - 0.24))
        bar.fill.solid(); bar.fill.fore_color.rgb = FUCHSIA if bi == 0 else CYAN
        bar.line.fill.background(); bar.shadow.inherit = False
        text(s, Inches(0.95), Inches(y + 0.14), Inches(9.6), Inches(0.3),
             [{"t": blk["role"], "size": 13.5, "bold": True, "color": WHITE}])
        chip(s, Inches(10.55), Inches(y + 0.14), Inches(2.0), Inches(0.32), blk["date"],
             color=CYAN, size=9, font=F_MONO)
        text(s, Inches(0.95), Inches(y + 0.52), Inches(11.5), Inches(bh - 0.6),
             [{"t": [("•  ", {"color": FUCHSIA if bi == 0 else CYAN}), (b, {})],
               "size": 10.5, "color": GRAY, "ls": 1.12, "sa": 3} for b in blk["bullets"]])
        y += bh + 0.22
    footer(s, n)

    # ----- 5. Experience: Bubejo -----
    s = new_slide(prs); n += 1
    x2 = L["exp2"]
    header(s, x2["kicker"], x2["title"])
    box(s, Inches(0.6), Inches(1.85), Inches(12.13), Inches(4.55), line=LINE)
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(2.0), Inches(0.06), Inches(4.25))
    bar.fill.solid(); bar.fill.fore_color.rgb = BLUE
    bar.line.fill.background(); bar.shadow.inherit = False
    text(s, Inches(0.98), Inches(2.1), Inches(9.8), Inches(0.4),
         [{"t": x2["role"], "size": 16, "bold": True, "color": WHITE}])
    text(s, Inches(0.98), Inches(2.52), Inches(9.8), Inches(0.32),
         [{"t": x2["comp"], "size": 12.5, "color": CYAN}])
    chip(s, Inches(10.95), Inches(2.12), Inches(1.55), Inches(0.34), x2["date"],
         color=CYAN, size=9, font=F_MONO)
    text(s, Inches(0.98), Inches(3.0), Inches(11.3), Inches(0.5),
         [{"t": x2["intro"], "size": 11.5, "color": GRAY, "ls": 1.2}])
    text(s, Inches(0.98), Inches(3.65), Inches(11.3), Inches(2.5),
         [{"t": [("•  ", {"color": CYAN}), (b, {})], "size": 11.5, "color": GRAY, "ls": 1.2, "sa": 6}
          for b in x2["bullets"]])
    footer(s, n)

    # ----- 6. Achievements -----
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

    # ----- 7. Skills -----
    s = new_slide(prs); n += 1
    sk = L["skills"]
    header(s, sk["kicker"], sk["title"])
    chips_flow(s, sk["items"], 0.57, 2.3, 12.2, gap_x=0.16, gap_y=0.18, h=0.5, size=11.5)
    text(s, Inches(0.6), Inches(6.3), Inches(12.1), Inches(0.3),
         [{"t": f"// {len(sk['items'])} tools & technologies", "size": 10.5, "color": GRAY2,
           "font": F_MONO, "align": PP_ALIGN.CENTER}])
    footer(s, n)

    # ----- 8. Selected projects -----
    s = new_slide(prs); n += 1
    p = L["proj"]
    header(s, p["kicker"], p["title"])
    pos = [(0.27, 1.8), (4.62, 1.8), (8.97, 1.8), (0.27, 4.35), (4.62, 4.35), (8.97, 4.35)]
    for it, (px, py) in zip(p["items"], pos):
        box(s, Inches(px), Inches(py), Inches(4.1), Inches(2.4), line=LINE)
        text(s, Inches(px + 0.25), Inches(py + 0.2), Inches(3.6), Inches(0.55),
             [{"t": it["t"], "size": 12, "bold": True, "color": WHITE, "ls": 1.0}])
        text(s, Inches(px + 0.25), Inches(py + 0.78), Inches(3.6), Inches(1.1),
             [{"t": it["d"], "size": 9.5, "color": GRAY, "ls": 1.1}])
        tw = 0.35 + 0.068 * len(it["tag"])
        chip(s, Inches(px + 0.25), Inches(py + 1.95), Inches(tw), Inches(0.28), it["tag"],
             color=CYAN, size=8.5, font=F_MONO)
    footer(s, n)

    # ----- 9. Closing / Contact -----
    s = new_slide(prs); n += 1
    cl = L["closing"]
    header(s, cl["kicker"], cl["title"])
    text(s, Inches(0.6), Inches(1.9), Inches(12.1), Inches(0.6),
         [{"t": cl["desc"], "size": 14, "color": GRAY, "ls": 1.25}])
    x = 0.5
    for it in cl["cards"]:
        box(s, Inches(x), Inches(2.9), Inches(2.95), Inches(1.7), line=LINE)
        text(s, Inches(x + 0.25), Inches(3.15), Inches(2.45), Inches(0.3),
             [{"t": it["h"].upper(), "size": 10, "color": CYAN, "font": F_MONO, "bold": True}])
        text(s, Inches(x + 0.25), Inches(3.5), Inches(2.45), Inches(0.9),
             [{"t": it["v"], "size": 12, "bold": True, "color": WHITE, "ls": 1.1}])
        x += 3.13
    text(s, Inches(0.6), Inches(5.4), Inches(12.1), Inches(0.3),
         [{"t": cl["loc"] + f"  ·  {YEAR}", "size": 11, "color": GRAY2, "font": F_MONO,
           "align": PP_ALIGN.CENTER}])
    footer(s, n)

    out = os.path.join(BASE, L["out"])
    prs.save(out)
    return out, n


if __name__ == "__main__":
    for L in (ID, EN):
        path, count = build(L)
        print(f"OK  {count} slides  ->  {path}")
