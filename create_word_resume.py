import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.opc.constants import RELATIONSHIP_TYPE

def create_resume():
    doc = docx.Document()

    # 1. Page Margins & Setup (US Letter, tight margins matching resume12.tex)
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        section.top_margin = Inches(0.32)
        section.bottom_margin = Inches(0.32)
        section.left_margin = Inches(0.40)
        section.right_margin = Inches(0.40)

    # 2. Executive Color Palette (Stripe/Apple 3-tier ink system)
    COLOR_CARBON = RGBColor(17, 24, 39)    # #111827 Primary Text
    COLOR_SLATE = RGBColor(75, 85, 99)     # #4B5563 Secondary Text / Dates
    COLOR_BORDER_HEX = "D1D5DB"            # #D1D5DB Hairline rule
    HEX_CARBON = "111827"
    HEX_SLATE = "4B5563"
    HEX_LINK = "111827"

    # Helper: Set paragraph spacing
    def format_para(p, before=0, after=0, line_spacing=1.05, align=WD_ALIGN_PARAGRAPH.LEFT):
        p.alignment = align
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.line_spacing = line_spacing

    # Helper: Add hyperlink with styling
    def add_hyperlink(paragraph, url, text, color=HEX_LINK, bold=False, italic=False, size_pt=9.0):
        part = paragraph.part
        r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w")} r:id="{r_id}" {nsdecls("r")}/>')
        
        new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
        rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
        
        font_elem = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>')
        rPr.append(font_elem)
        
        sz_val = str(int(size_pt * 2))
        sz_elem = parse_xml(f'<w:sz {nsdecls("w")} w:val="{sz_val}"/>')
        rPr.append(sz_elem)
        
        if color:
            c_elem = parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>')
            rPr.append(c_elem)
        if bold:
            rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
        if italic:
            rPr.append(parse_xml(f'<w:i {nsdecls("w")}/>'))
            
        new_run.append(rPr)
        t_elem = OxmlElement('w:t')
        t_elem.text = text
        new_run.append(t_elem)
        hyperlink.append(new_run)
        paragraph._p.append(hyperlink)

    # Helper: Section Header with Hairline Bottom Border
    def add_section_header(title):
        p = doc.add_paragraph()
        format_para(p, before=4.0, after=2.5)
        run = p.add_run(title)
        run.font.name = "Arial"
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_CARBON
        
        # Add 0.5pt bottom border in #D1D5DB
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="2" w:color="{COLOR_BORDER_HEX}"/></w:pBdr>')
        pPr.append(pBdr)

    # Helper: 2-Column Table for Subheadings / Project Headings
    def add_two_column_heading(left_runs, right_runs):
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.autofit = False
        
        # 7.7 inches total printable width (8.5 - 0.8)
        table.columns[0].width = Inches(5.6)
        table.columns[1].width = Inches(2.1)
        
        cell_left = table.cell(0, 0)
        cell_right = table.cell(0, 1)
        
        # Remove padding
        for cell in (cell_left, cell_right):
            tcPr = cell._tc.get_or_add_tcPr()
            tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>')
            tcPr.append(tcMar)
            
        p_left = cell_left.paragraphs[0]
        format_para(p_left, before=2.0, after=0.5)
        for item in left_runs:
            if isinstance(item, tuple) and len(item) == 5 and item[0] == "link":
                _, url, text, bold, italic = item
                add_hyperlink(p_left, url, text, color=HEX_CARBON, bold=bold, italic=italic, size_pt=9.5)
            else:
                text, bold, italic, color = item
                r = p_left.add_run(text)
                r.font.name = "Arial"
                r.font.size = Pt(9.5)
                r.font.bold = bold
                r.font.italic = italic
                r.font.color.rgb = color
                
        p_right = cell_right.paragraphs[0]
        format_para(p_right, before=2.0, after=0.5, align=WD_ALIGN_PARAGRAPH.RIGHT)
        for item in right_runs:
            if isinstance(item, tuple) and len(item) == 5 and item[0] == "link":
                _, url, text, bold, italic = item
                add_hyperlink(p_right, url, text, color=HEX_CARBON, bold=bold, italic=italic, size_pt=9.5)
            else:
                text, bold, italic, color = item
                r = p_right.add_run(text)
                r.font.name = "Arial"
                r.font.size = Pt(9.5)
                r.font.bold = bold
                r.font.italic = italic
                r.font.color.rgb = color

    # Helper: Bullet point with bold text parsing
    def add_bullet_item(run_specs):
        p = doc.add_paragraph()
        format_para(p, before=0.5, after=1.5, line_spacing=1.05)
        p.paragraph_format.left_indent = Inches(0.18)
        p.paragraph_format.first_line_indent = Inches(-0.12)
        
        # Bullet glyph
        bullet_run = p.add_run("•  ")
        bullet_run.font.name = "Arial"
        bullet_run.font.size = Pt(8.5)
        bullet_run.font.color.rgb = COLOR_SLATE
        
        for item in run_specs:
            if isinstance(item, tuple) and len(item) == 4 and item[0] == "link":
                _, url, text, bold = item
                add_hyperlink(p, url, text, color=HEX_CARBON, bold=bold, size_pt=9.5)
            else:
                text, bold = item
                r = p.add_run(text)
                r.font.name = "Arial"
                r.font.size = Pt(9.5)
                r.font.bold = bold
                r.font.color.rgb = COLOR_CARBON

    # ==================== HEADER ====================
    # Name
    p_name = doc.add_paragraph()
    format_para(p_name, before=0, after=1.5)
    run_name = p_name.add_run("Ravindra Chowdary Jonnagadla")
    run_name.font.name = "Arial"
    run_name.font.size = Pt(20)
    run_name.font.bold = True
    run_name.font.color.rgb = COLOR_CARBON

    # Subtitle
    p_sub = doc.add_paragraph()
    format_para(p_sub, before=0, after=2.5)
    run_sub = p_sub.add_run("AI Systems Engineer  •  Machine Learning Researcher")
    run_sub.font.name = "Arial"
    run_sub.font.size = Pt(10)
    run_sub.font.bold = False
    run_sub.font.color.rgb = COLOR_SLATE

    # Contact Line 1
    p_c1 = doc.add_paragraph()
    format_para(p_c1, before=0, after=1.5)
    r = p_c1.add_run("Chennai, India  •  +91 9398151678  •  ")
    r.font.name = "Arial"
    r.font.size = Pt(8.5)
    r.font.color.rgb = COLOR_SLATE
    add_hyperlink(p_c1, "mailto:ravindrachowdary8008@gmail.com", "ravindrachowdary8008@gmail.com", color=HEX_SLATE, size_pt=8.5)

    # Contact Line 2 (Interactive Profiles)
    p_c2 = doc.add_paragraph()
    format_para(p_c2, before=0, after=3.0)
    add_hyperlink(p_c2, "https://linkedin.com/in/ravindrachowdary", "linkedin.com/in/ravindrachowdary", color=HEX_SLATE, size_pt=8.5)
    r = p_c2.add_run("  •  ")
    r.font.name = "Arial"
    r.font.size = Pt(8.5)
    r.font.color.rgb = COLOR_SLATE
    add_hyperlink(p_c2, "https://github.com/RAVINDRA8008", "github.com/RAVINDRA8008", color=HEX_SLATE, size_pt=8.5)
    r = p_c2.add_run("  •  ")
    r.font.name = "Arial"
    r.font.size = Pt(8.5)
    r.font.color.rgb = COLOR_SLATE
    add_hyperlink(p_c2, "https://huggingface.co/r5avindra", "huggingface.co/r5avindra", color=HEX_SLATE, size_pt=8.5)
    r = p_c2.add_run("  •  ")
    r.font.name = "Arial"
    r.font.size = Pt(8.5)
    r.font.color.rgb = COLOR_SLATE
    add_hyperlink(p_c2, "https://lynkzen.in", "lynkzen.in", color=HEX_SLATE, size_pt=8.5)


    # ==================== PROFESSIONAL EXPERIENCE ====================
    add_section_header("Professional Experience")

    # Entry 1: Waani
    add_two_column_heading(
        left_runs=[("Waani – Conversational Voice AI Platform", True, False, COLOR_CARBON)],
        right_runs=[("Chennai, India", False, False, COLOR_SLATE)]
    )
    add_two_column_heading(
        left_runs=[("Founder & Lead Engineer", False, True, COLOR_CARBON)],
        right_runs=[("July 2026 – Present", False, False, COLOR_SLATE)]
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

    # Entry 2: LynkZen
    add_two_column_heading(
        left_runs=[("link", "https://lynkzen.in", "LynkZen – Instagram DM Automation SaaS", True, False)],
        right_runs=[("Chennai, India", False, False, COLOR_SLATE)]
    )
    add_two_column_heading(
        left_runs=[("Founding Engineer", False, True, COLOR_CARBON)],
        right_runs=[("April 2026 – June 2026", False, False, COLOR_SLATE)]
    )
    add_bullet_item([
        ("Built and launched an event-driven Instagram automation engine using Next.js, TypeScript, and Tailwind CSS, decoupling webhook ingest via BullMQ asynchronous worker queues to absorb bursts of 10k+ concurrent webhook events.", False)
    ])
    add_bullet_item([
        ("Developed an interactive node-based visual workflow builder allowing creators to configure multi-step conditional messaging funnels and automated response paths without writing code.", False)
    ])
    add_bullet_item([
        ("Scaled platform to ", False),
        ("3,000 active beta signups for closed testing", True),
        (", implementing token-bucket rate limiting via atomic Redis Lua scripts to eliminate race conditions and guarantee zero 429 quota violations against Meta Graph API limits.", False)
    ])


    # ==================== TECHNICAL PROJECTS & RESEARCH ====================
    add_section_header("Technical Projects & Research")

    # Project 1: AMERS
    add_two_column_heading(
        left_runs=[
            ("link", "https://github.com/RAVINDRA8008/Adaptive-GAN-Augmented-EEG-Speech-Emotion-Recognition-via-Deep-Reinforcement-Learning", "AMERS: Multimodal Emotion Recognition System", True, False),
            ("  •  ", False, False, COLOR_SLATE),
            ("PyTorch, Signal Processing", False, True, COLOR_CARBON)
        ],
        right_runs=[
            ("link", "https://huggingface.co/r5avindra/amers-checkpoints", "Hugging Face", True, False),
            ("  •  ", False, False, COLOR_SLATE),
            ("IEEE ICETEA 2026", True, False, COLOR_CARBON)
        ]
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
        ("0.77 Cohen’s kappa", True),
        (", open-sourcing checkpoints on ", False),
        ("link", "https://huggingface.co/r5avindra/amers-checkpoints", "Hugging Face", True),
        (".", False)
    ])

    # Project 2: SmartCap
    add_two_column_heading(
        left_runs=[
            ("link", "https://github.com/RAVINDRA8008/SmartCap-Real-Time-Synchronized-Subtitles-for-Videos", "SmartCap: Automated Subtitle & Transcoding Pipeline", True, False),
            ("  •  ", False, False, COLOR_SLATE),
            ("Python, Whisper, FFmpeg, ONNX", False, True, COLOR_CARBON)
        ],
        right_runs=[
            ("link", "https://ieeexplore.ieee.org/document/11085491", "IEEE ICICV 2025", True, False)
        ]
    )
    add_bullet_item([
        ("Built an automated transcription pipeline combining Whisper with timestamp alignment for frame-accurate subtitles (.srt/.vtt), quantizing to INT8 via ONNX Runtime to achieve a ", False),
        ("0.35x Real-Time Factor (RTF)", True),
        (", published as Lead Author (Doc ID: 11085491).", False)
    ])
    add_bullet_item([
        ("Architected an asynchronous worker pipeline orchestrating FFmpeg stream filters for dynamic audio normalization and hardcoded subtitle burning, isolating CPU-bound transcoding to eliminate event-loop starvation.", False)
    ])

    # Project 3: Smart QR Parking
    add_two_column_heading(
        left_runs=[
            ("link", "https://github.com/RAVINDRA8008/Smart-QR-Based-Parking-System-main", "Smart QR-Based University Parking System", True, False),
            ("  •  ", False, False, COLOR_SLATE),
            ("Flask, Python, PostgreSQL, Redis", False, True, COLOR_CARBON)
        ],
        right_runs=[
            ("link", "https://ieeexplore.ieee.org/document/11166870", "IEEE CONIT 2025", True, False)
        ]
    )
    add_bullet_item([
        ("Designed a campus parking management application with HMAC-SHA256 signed dynamic QR codes, timestamped nonces cached in Redis to prevent replay attacks, and Twilio OTP verification, published as Co-Author (Doc ID: 11166870).", False)
    ])
    add_bullet_item([
        ("Engineered slot reservation concurrency using optimistic row-level locking with versioning in PostgreSQL, maintaining ", False),
        ("zero double-booking anomalies and a 99.8% transaction success rate", True),
        (" under simulated 1,000-user rush-hour contention.", False)
    ])


    # ==================== TECHNICAL SKILLS ====================
    add_section_header("Technical Skills")

    skills_data = [
        ("Languages: ", "Python, C++, SQL (PostgreSQL, MySQL), JavaScript, TypeScript, Bash, Shell Scripting"),
        ("AI/ML & Research: ", "PyTorch, TensorFlow, Scikit-learn, Hugging Face, Multimodal Learning, WGAN-GP, Reinforcement Learning, PPO, NLP, EEG Signal Processing, Vector Search, ChromaDB"),
        ("Generative AI & Speech: ", "Whisper, Faster-Whisper, ONNX Runtime, Model Quantization, Sentence Transformers, RAG, FFmpeg, Audio/Speech Processing (STT/TTS)"),
        ("Backend & Systems: ", "FastAPI, Flask, Django, Node.js, Next.js, React, RESTful APIs, WebSockets, Redis, BullMQ, Celery, Microservices"),
        ("Cloud, DevOps & Observability: ", "AWS (EC2, S3, Lambda, IAM), Cloudflare R2, Docker, Kubernetes, Terraform, GitHub Actions, Linux, Git, Prometheus, Grafana")
    ]

    for label, text in skills_data:
        p = doc.add_paragraph()
        format_para(p, before=0.5, after=1.5, line_spacing=1.05)
        p.paragraph_format.left_indent = Inches(0.0)
        
        r_lbl = p.add_run(label)
        r_lbl.font.name = "Arial"
        r_lbl.font.size = Pt(9.5)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = COLOR_CARBON
        
        r_txt = p.add_run(text)
        r_txt.font.name = "Arial"
        r_txt.font.size = Pt(9.5)
        r_txt.font.color.rgb = COLOR_CARBON

    # Certifications Row with Hyperlinks
    p_cert = doc.add_paragraph()
    format_para(p_cert, before=0.5, after=1.5, line_spacing=1.05)
    r_cert = p_cert.add_run("Certifications: ")
    r_cert.font.name = "Arial"
    r_cert.font.size = Pt(9.5)
    r_cert.font.bold = True
    r_cert.font.color.rgb = COLOR_CARBON
    
    add_hyperlink(p_cert, "https://catalog-education.oracle.com/pls/certview/sharebadge?id=612702728476686D7B2D24830AC8D7769D0C9345E82B11C94343869FEEB8DF92", 
                  "Oracle Certified Professional: MySQL 8.0 Database Developer (1Z0-909)", color=HEX_CARBON, size_pt=9.5)
    
    r_sep = p_cert.add_run("  •  ")
    r_sep.font.name = "Arial"
    r_sep.font.size = Pt(9.5)
    r_sep.font.color.rgb = COLOR_SLATE
    
    add_hyperlink(p_cert, "https://www.credly.com/badges/3dc45737-a240-4ebd-865c-ac932ab9d71c/public_url",
                  "AWS Certified Cloud Practitioner", color=HEX_CARBON, size_pt=9.5)


    # ==================== EDUCATION ====================
    add_section_header("Education")

    add_two_column_heading(
        left_runs=[("SRM Institute of Science and Technology", True, False, COLOR_CARBON)],
        right_runs=[("Chennai, India", False, False, COLOR_SLATE)]
    )
    add_two_column_heading(
        left_runs=[
            ("B.Tech in Computer Science and Engineering – Cloud Computing ", False, True, COLOR_CARBON),
            ("(CGPA: 8.01/10)", False, False, COLOR_SLATE)
        ],
        right_runs=[("May 2026", False, False, COLOR_SLATE)]
    )

    # 3. Save to DOCX
    out_path_full = os.path.abspath("Ravindra_Chowdary_Jonnagadla_Resume.docx")
    out_path_resume12 = os.path.abspath("resume12.docx")
    
    doc.save(out_path_full)
    doc.save(out_path_resume12)
    print(f"Generated: {out_path_full}")
    print(f"Generated: {out_path_resume12}")

if __name__ == "__main__":
    create_resume()
