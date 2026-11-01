"""
Generate Ravindra Chowdary Jonnagadla's Student & Fellowship Resume in Word (.docx)
Strict 1-Page Calibrated Architecture matching resume_student.tex:
- Education at TOP (SRM IST, CGPA 8.01, Coursework)
- Research Publications & Technical Projects (AMERS, SmartCap, Smart QR)
- Entrepreneurial Leadership & Technical Initiatives (Waani & LynkZen - Founder & Lead Engineer)
- Technical Skills & Certifications
- Native Right Tab Stops at 7.7 inches (Zero tables)
"""

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE

def add_hyperlink(paragraph, url, text, color="111827", underline=False, bold=False, italic=False, size_pt=9.4):
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')
    rPr.append(rFonts)

    c = OxmlElement('w:color')
    c.set(qn('w:val'), color)
    rPr.append(c)

    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(size_pt * 2)))
    rPr.append(sz)

    if bold:
        rPr.append(OxmlElement('w:b'))
    if italic:
        rPr.append(OxmlElement('w:i'))
    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def generate_student_docx():
    doc = docx.Document()

    COLOR_CARBON = RGBColor(17, 24, 39)
    COLOR_SLATE = RGBColor(75, 85, 99)
    HEX_CARBON = "111827"
    HEX_SLATE = "4B5563"

    for section in doc.sections:
        section.top_margin = Inches(0.32)
        section.bottom_margin = Inches(0.30)
        section.left_margin = Inches(0.40)
        section.right_margin = Inches(0.40)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)

    RIGHT_TAB_POS = Inches(7.70)

    def format_para(p, before=0.0, after=0.0, line_spacing=1.05):
        pf = p.paragraph_format
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.line_spacing = line_spacing

    def add_section_header(title):
        p = doc.add_paragraph()
        format_para(p, before=6.5, after=3.0)
        run = p.add_run(title)
        run.font.name = "Arial"
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_CARBON

        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '4')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), 'D1D5DB')
        pBdr.append(bottom)
        pPr.append(pBdr)

    def add_tab_heading(left_runs, right_runs, before=0.0, after=0.0, size_pt=9.6):
        p = doc.add_paragraph()
        format_para(p, before=before, after=after, line_spacing=1.04)
        p.paragraph_format.tab_stops.add_tab_stop(RIGHT_TAB_POS, WD_TAB_ALIGNMENT.RIGHT)

        for item in left_runs:
            if isinstance(item, tuple) and item[0] == "link":
                _, url, text, bold, italic = item
                add_hyperlink(p, url, text, color=HEX_CARBON, bold=bold, italic=italic, size_pt=size_pt)
            else:
                text, bold, italic, color = item
                r = p.add_run(text)
                r.font.name = "Arial"
                r.font.size = Pt(size_pt)
                r.font.bold = bold
                r.font.italic = italic
                r.font.color.rgb = color

        p.add_run("\t")

        for item in right_runs:
            if isinstance(item, tuple) and item[0] == "link":
                _, url, text, bold, italic = item
                add_hyperlink(p, url, text, color=HEX_CARBON, bold=bold, italic=italic, size_pt=size_pt)
            else:
                text, bold, italic, color = item
                r = p.add_run(text)
                r.font.name = "Arial"
                r.font.size = Pt(size_pt)
                r.font.bold = bold
                r.font.italic = italic
                r.font.color.rgb = color

    def add_bullet_item(run_specs, before=0.0, after=2.5):
        p = doc.add_paragraph()
        format_para(p, before=before, after=after, line_spacing=1.05)
        p.paragraph_format.left_indent = Inches(0.14)
        p.paragraph_format.first_line_indent = Inches(-0.10)

        bullet_run = p.add_run("•  ")
        bullet_run.font.name = "Arial"
        bullet_run.font.size = Pt(8.5)
        bullet_run.font.color.rgb = COLOR_SLATE

        for item in run_specs:
            if isinstance(item, tuple) and item[0] == "link":
                _, url, text, bold = item
                add_hyperlink(p, url, text, color=HEX_CARBON, bold=bold, size_pt=9.3)
            else:
                text, bold = item
                r = p.add_run(text)
                r.font.name = "Arial"
                r.font.size = Pt(9.3)
                r.font.bold = bold
                r.font.color.rgb = COLOR_CARBON

    # ==================== HEADER ====================
    p_name = doc.add_paragraph()
    format_para(p_name, before=0.0, after=2.0)
    run_name = p_name.add_run("Ravindra Chowdary Jonnagadla")
    run_name.font.name = "Arial"
    run_name.font.size = Pt(20.5)
    run_name.font.bold = True
    run_name.font.color.rgb = COLOR_CARBON

    p_sub = doc.add_paragraph()
    format_para(p_sub, before=0.0, after=3.0)
    run_sub = p_sub.add_run("Computer Science Undergraduate  •  AI Systems Researcher")
    run_sub.font.name = "Arial"
    run_sub.font.size = Pt(10.0)
    run_sub.font.color.rgb = COLOR_SLATE

    p_c1 = doc.add_paragraph()
    format_para(p_c1, before=0.0, after=2.0)
    r = p_c1.add_run("Chennai, India  •  +91 9398151678  •  ")
    r.font.name = "Arial"
    r.font.size = Pt(8.8)
    r.font.color.rgb = COLOR_SLATE
    add_hyperlink(p_c1, "mailto:ravindrachowdary8008@gmail.com", "ravindrachowdary8008@gmail.com", color=HEX_SLATE, size_pt=8.8)

    p_c2 = doc.add_paragraph()
    format_para(p_c2, before=0.0, after=3.5)
    add_hyperlink(p_c2, "https://linkedin.com/in/ravindrachowdary", "linkedin.com/in/ravindrachowdary", color=HEX_SLATE, size_pt=8.8)
    r = p_c2.add_run("  •  ")
    r.font.name = "Arial"
    r.font.size = Pt(8.8)
    r.font.color.rgb = COLOR_SLATE
    add_hyperlink(p_c2, "https://github.com/RAVINDRA8008", "github.com/RAVINDRA8008", color=HEX_SLATE, size_pt=8.8)
    r = p_c2.add_run("  •  ")
    r.font.name = "Arial"
    r.font.size = Pt(8.8)
    r.font.color.rgb = COLOR_SLATE
    add_hyperlink(p_c2, "https://huggingface.co/r5avindra", "huggingface.co/r5avindra", color=HEX_SLATE, size_pt=8.8)
    r = p_c2.add_run("  •  ")
    r.font.name = "Arial"
    r.font.size = Pt(8.8)
    r.font.color.rgb = COLOR_SLATE
    add_hyperlink(p_c2, "https://lynkzen.in", "lynkzen.in", color=HEX_SLATE, size_pt=8.8)

    # ==================== EDUCATION (TOP) ====================
    add_section_header("Education")
    add_tab_heading(
        left_runs=[("SRM Institute of Science and Technology", True, False, COLOR_CARBON)],
        right_runs=[("Chennai, India", False, False, COLOR_SLATE)],
        before=2.5, after=0.5
    )
    add_tab_heading(
        left_runs=[("B.Tech in Computer Science and Engineering – Specialization in Cloud Computing", False, True, COLOR_CARBON)],
        right_runs=[("Expected May 2026", False, False, COLOR_SLATE)],
        before=0.0, after=1.5
    )
    add_bullet_item([
        ("Academic Standing: ", True),
        ("Cumulative Grade Point Average (", False),
        ("CGPA", True),
        ("): ", False),
        ("8.01 / 10.0.", True)
    ], after=1.8)
    add_bullet_item([
        ("Relevant Coursework: ", True),
        ("Distributed Systems, Deep Learning, Cloud Computing Architecture, Database Management Systems, Data Structures & Algorithms, Operating Systems, Computer Networks, Object-Oriented Programming (C++).", False)
    ], after=2.5)

    # ==================== RESEARCH PUBLICATIONS & PROJECTS ====================
    add_section_header("Research Publications & Technical Projects")

    # AMERS
    add_tab_heading(
        left_runs=[
            ("link", "https://github.com/RAVINDRA8008/Adaptive-GAN-Augmented-EEG-Speech-Emotion-Recognition-via-Deep-Reinforcement-Learning", "AMERS: Multimodal Emotion Recognition System", True, False),
            ("  •  ", False, False, COLOR_SLATE),
            ("PyTorch, Signal Processing", False, True, COLOR_CARBON)
        ],
        right_runs=[
            ("link", "https://huggingface.co/r5avindra/amers-checkpoints", "Hugging Face", True, False),
            ("  •  ", False, False, COLOR_SLATE),
            ("IEEE ICETEA 2026", True, False, COLOR_CARBON)
        ],
        before=2.8, after=1.2
    )
    add_bullet_item([
        ("Developed a multimodal emotion recognition architecture synchronizing multi-channel EEG (128Hz) and speech spectrograms via Cross-Modal Mutual Attention (CMMA) to capture inter-modal temporal-spatial dynamics.", False)
    ])
    add_bullet_item([
        ("Integrated an Emotion-Aware Gating (EAG) WGAN-GP generator with a PPO reinforcement learning controller, dynamically scheduling synthetic-to-real training batches against discriminator reward stability.", False)
    ])
    add_bullet_item([
        ("Evaluated on DEAP/SEED benchmarks via LOSO cross-validation across 4 emotion classes, achieving ", False),
        ("82.55% accuracy", True),
        (" (+7.4% over baselines), ", False),
        ("0.83 macro F1-score", True),
        (", and ", False),
        ("0.77 Cohen's kappa", True),
        (", open-sourcing checkpoints on ", False),
        ("link", "https://huggingface.co/r5avindra/amers-checkpoints", "Hugging Face", True),
        (".", False)
    ])

    # SmartCap
    add_tab_heading(
        left_runs=[
            ("link", "https://github.com/RAVINDRA8008/SmartCap-Real-Time-Synchronized-Subtitles-for-Videos", "SmartCap: Automated Subtitle & Transcoding Pipeline", True, False),
            ("  •  ", False, False, COLOR_SLATE),
            ("Python, Whisper, FFmpeg, ONNX", False, True, COLOR_CARBON)
        ],
        right_runs=[
            ("link", "https://ieeexplore.ieee.org/document/11085491", "IEEE ICICV 2025", True, False)
        ],
        before=3.0, after=1.2
    )
    add_bullet_item([
        ("Built an automated transcription pipeline combining Whisper with timestamp alignment for frame-accurate subtitles (.srt/.vtt), quantizing to INT8 via ONNX Runtime to achieve a ", False),
        ("0.35x Real-Time Factor (RTF)", True),
        (", published as Lead Author (Doc ID: 11085491).", False)
    ])
    add_bullet_item([
        ("Architected an asynchronous worker pipeline orchestrating FFmpeg stream filters for dynamic audio normalization and hardcoded subtitle burning, isolating CPU-bound transcoding to eliminate event-loop starvation.", False)
    ])

    # Smart QR
    add_tab_heading(
        left_runs=[
            ("link", "https://github.com/RAVINDRA8008/Smart-QR-Based-Parking-System-main", "Smart QR-Based University Parking System", True, False),
            ("  •  ", False, False, COLOR_SLATE),
            ("Flask, Python, PostgreSQL, Redis", False, True, COLOR_CARBON)
        ],
        right_runs=[
            ("link", "https://ieeexplore.ieee.org/document/11166870", "IEEE CONIT 2025", True, False)
        ],
        before=3.0, after=1.2
    )
    add_bullet_item([
        ("Designed a campus parking management application with HMAC-SHA256 signed dynamic QR codes, timestamped nonces cached in Redis to prevent replay attacks, and Twilio OTP verification, published as Co-Author (Doc ID: 11166870).", False)
    ])
    add_bullet_item([
        ("Engineered slot reservation concurrency using optimistic row-level locking with versioning in PostgreSQL, maintaining ", False),
        ("zero double-booking anomalies and a 99.8% transaction success rate", True),
        (" under simulated 1,000-user rush-hour contention.", False)
    ])

    # ==================== ENTREPRENEURSHIP & LEADERSHIP ====================
    add_section_header("Entrepreneurial Leadership & Technical Initiatives")

    # Waani
    add_tab_heading(
        left_runs=[("Waani – Conversational Voice AI Platform", True, False, COLOR_CARBON)],
        right_runs=[("Chennai, India", False, False, COLOR_SLATE)],
        before=2.8, after=0.5
    )
    add_tab_heading(
        left_runs=[("Founder & Lead Engineer", False, True, COLOR_CARBON)],
        right_runs=[("July 2026 – Present", False, False, COLOR_SLATE)],
        before=0.0, after=1.2
    )
    add_bullet_item([
        ("Architected an ultra-low-latency duplex voice pipeline orchestrating telephony protocols, streaming STT (Sarvam AI), and LLM token inference, achieving ", False),
        ("<800ms p95 voice-to-voice turn-around time", True),
        (" across 5,000+ concurrent voice sessions.", False)
    ])
    add_bullet_item([
        ("Engineered conversational state machines and dynamic intent branching for automated lead qualification, instrumenting Prometheus metrics and Grafana dashboards to track streaming audio packet jitter and queue depth.", False)
    ])
    add_bullet_item([
        ("Built an asynchronous dual-channel audio ring buffer with edge Voice Activity Detection (VAD) over WebSockets, implementing atomic cancellation tokens to flush in-flight TTS buffers on instant user barge-in.", False)
    ])

    # LynkZen
    add_tab_heading(
        left_runs=[("link", "https://lynkzen.in", "LynkZen – Instagram DM Automation SaaS", True, False)],
        right_runs=[("Chennai, India", False, False, COLOR_SLATE)],
        before=3.0, after=0.5
    )
    add_tab_heading(
        left_runs=[("Founder & Lead Engineer", False, True, COLOR_CARBON)],
        right_runs=[("April 2026 – June 2026", False, False, COLOR_SLATE)],
        before=0.0, after=1.2
    )
    add_bullet_item([
        ("Built and launched an event-driven Instagram automation engine using Next.js, TypeScript, and Tailwind CSS, decoupling webhook ingest via BullMQ asynchronous worker queues to absorb bursts of 10k+ concurrent webhook events.", False)
    ])
    add_bullet_item([
        ("Scaled platform to ", False),
        ("3,000 active beta signups for closed testing", True),
        (", implementing token-bucket rate limiting via atomic Redis Lua scripts to eliminate race conditions and guarantee zero 429 quota violations against Meta Graph API limits.", False)
    ])

    # ==================== TECHNICAL SKILLS & CERTIFICATIONS ====================
    add_section_header("Technical Skills & Certifications")

    skill_rows = [
        ("Languages: ", "Python, C++, SQL (PostgreSQL, MySQL), JavaScript, TypeScript, Bash, Shell Scripting"),
        ("AI/ML & Research: ", "PyTorch, TensorFlow, Scikit-learn, Hugging Face, Multimodal Learning, WGAN-GP, Reinforcement Learning, PPO, NLP, EEG Signal Processing, Vector Search, ChromaDB"),
        ("Generative AI & Speech: ", "Whisper, Faster-Whisper, ONNX Runtime, Model Quantization, Sentence Transformers, RAG, FFmpeg, Audio/Speech Processing (STT/TTS)"),
        ("Backend, Cloud & DevOps: ", "FastAPI, Flask, Django, Node.js, Next.js, WebSockets, Redis, AWS (EC2, S3, Lambda), Docker, Kubernetes, Prometheus, Grafana")
    ]

    for label, items in skill_rows:
        p = doc.add_paragraph()
        format_para(p, before=0.0, after=1.8, line_spacing=1.04)
        r_lbl = p.add_run(label)
        r_lbl.font.name = "Arial"
        r_lbl.font.size = Pt(9.1)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = COLOR_CARBON

        r_val = p.add_run(items)
        r_val.font.name = "Arial"
        r_val.font.size = Pt(9.1)
        r_val.font.bold = False
        r_val.font.color.rgb = COLOR_CARBON

    p_cert = doc.add_paragraph()
    format_para(p_cert, before=0.0, after=1.0, line_spacing=1.04)
    r_c_lbl = p_cert.add_run("Certifications: ")
    r_c_lbl.font.name = "Arial"
    r_c_lbl.font.size = Pt(9.1)
    r_c_lbl.font.bold = True
    r_c_lbl.font.color.rgb = COLOR_CARBON

    add_hyperlink(p_cert, "https://www.credly.com/badges/3dc45737-a240-4ebd-865c-ac932ab9d71c/public_url", "Oracle Certified Professional: MySQL 8.0 Database Developer (1Z0-909)", color=HEX_CARBON, bold=False, size_pt=9.1)
    r_sep = p_cert.add_run("  •  ")
    r_sep.font.name = "Arial"
    r_sep.font.size = Pt(9.1)
    r_sep.font.color.rgb = COLOR_SLATE

    add_hyperlink(p_cert, "https://www.credly.com/badges/3dc45737-a240-4ebd-865c-ac932ab9d71c/public_url", "AWS Certified Cloud Practitioner", color=HEX_CARBON, bold=False, size_pt=9.1)

    output_path = r"C:\Users\ravin\Downloads\RESUM\Ravindra_Chowdary_Student_Resume.docx"
    doc.save(output_path)
    print(f"Student DOCX generated at: {output_path}")

if __name__ == "__main__":
    generate_student_docx()
