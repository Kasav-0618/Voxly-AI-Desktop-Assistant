"""
Generate a professional, high-impact PowerPoint presentation (.pptx)
showcasing VOXLY AI Desktop Assistant and its architecture & features.
"""

import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_deck(output_path="C:\\voicebridge\\VOXLY_AI_Desktop_Assistant_Presentation.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Color Palette - Modern Executive Cyber / Slate Dark Theme
    COLOR_BG = RGBColor(15, 23, 42)          # Slate-900 (#0F172A)
    COLOR_CARD = RGBColor(30, 41, 59)        # Slate-800 (#1E293B)
    COLOR_CARD_BORDER = RGBColor(51, 65, 85) # Slate-700 (#334155)
    COLOR_CYAN = RGBColor(6, 182, 212)       # Cyan-500 (#06B6D4)
    COLOR_PURPLE = RGBColor(139, 92, 246)    # Violet-500 (#8B5CF6)
    COLOR_EMERALD = RGBColor(16, 185, 129)   # Emerald-500 (#10B981)
    COLOR_WHITE = RGBColor(248, 250, 252)    # Slate-50 (#F8FAFC)
    COLOR_MUTED = RGBColor(148, 163, 184)    # Slate-400 (#94A3B8)
    COLOR_ACCENT = RGBColor(99, 102, 241)    # Indigo-500 (#6366F1)

    blank_layout = prs.slide_layouts[6]

    def set_slide_background(slide):
        bg_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5)
        )
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = COLOR_BG
        bg_shape.line.fill.background()

    def add_header(slide, title_text, category_text="VOXLY AI DESKTOP ASSISTANT"):
        # Category / Kicker
        kicker_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.35))
        tf_k = kicker_box.text_frame
        tf_k.word_wrap = True
        p_k = tf_k.paragraphs[0]
        p_k.text = category_text.upper()
        p_k.font.size = Pt(11)
        p_k.font.bold = True
        p_k.font.color.rgb = COLOR_CYAN
        p_k.font.name = "Segoe UI"

        # Main Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.7), Inches(0.7))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.size = Pt(26)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_WHITE
        p_t.font.name = "Segoe UI"

        # Subtle divider
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.45), Inches(11.733), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_CARD_BORDER
        line.line.fill.background()

    def add_card(slide, left, top, width, height, title, subtitle=None, border_color=COLOR_CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = border_color
        card.line.width = Pt(1.5)
        
        # Header inside card
        tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), Inches(0.8))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = COLOR_CYAN
        p.font.name = "Segoe UI"

        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.font.size = Pt(11)
            p2.font.color.rgb = COLOR_MUTED
            p2.font.name = "Segoe UI"

        return card

    # ==========================================================
    # SLIDE 1: Title Slide
    # ==========================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)

    # Decorative background glows / pill badge
    badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.3), Inches(3.2), Inches(0.42))
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_CARD
    badge.line.color.rgb = COLOR_CYAN
    badge.line.width = Pt(1)
    b_tf = badge.text_frame
    b_p = b_tf.paragraphs[0]
    b_p.text = "★ NEXT-GEN DESKTOP AI ASSISTANT"
    b_p.font.size = Pt(10)
    b_p.font.bold = True
    b_p.font.color.rgb = COLOR_CYAN
    b_p.font.name = "Segoe UI"
    b_p.alignment = PP_ALIGN.CENTER

    # Main Title Box
    title_box = s1.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.3), Inches(2.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "VOXLY AI"
    p.font.size = Pt(64)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    p.font.name = "Segoe UI"

    p_sub = tf.add_paragraph()
    p_sub.text = "Intelligent Voice Control, Deep Windows Automation & ChatGPT-Level Brain"
    p_sub.font.size = Pt(22)
    p_sub.font.color.rgb = COLOR_CYAN
    p_sub.font.name = "Segoe UI"

    # Pitch description
    desc_box = s1.shapes.add_textbox(Inches(1.0), Inches(4.3), Inches(10.5), Inches(1.2))
    d_tf = desc_box.text_frame
    d_tf.word_wrap = True
    dp = d_tf.paragraphs[0]
    dp.text = "A complete, production-grade native Windows desktop application that listens to natural voice instructions, reasons with LLM intelligence, and executes real system, application, and file operations on your PC without manual effort."
    dp.font.size = Pt(14)
    dp.font.color.rgb = COLOR_MUTED
    dp.font.name = "Segoe UI"

    # Meta Tags row
    tags = ["Voice Recognition", "Zero-Hardcode App Discovery", "Deep Folder Navigation", "Playwright Web", "Python + CustomTkinter"]
    tx = Inches(1.0)
    for tag in tags:
        tw = Inches(len(tag) * 0.13 + 0.5)
        tg = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tx, Inches(5.8), tw, Inches(0.45))
        tg.fill.solid()
        tg.fill.fore_color.rgb = COLOR_CARD
        tg.line.color.rgb = COLOR_CARD_BORDER
        tp = tg.text_frame.paragraphs[0]
        tp.text = tag
        tp.font.size = Pt(11)
        tp.font.color.rgb = COLOR_WHITE
        tp.font.name = "Segoe UI"
        tp.alignment = PP_ALIGN.CENTER
        tx += tw + Inches(0.2)

    # ==========================================================
    # SLIDE 2: Executive Summary & The Problem Solved
    # ==========================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_header(s2, "Executive Overview: Reimagining Human-PC Interaction", "Problem & Vision")

    cards_data_s2 = [
        ("The Traditional Limitation", "Traditional voice assistants (Siri, Cortana) are constrained to closed ecosystems, canned web searches, and surface-level commands. They cannot dynamically discover your unique local software or deeply navigate multi-nested drives and study folders.", COLOR_PURPLE),
        ("The VOXLY Breakthrough", "VOXLY bridges ChatGPT's conversational intelligence directly with native Windows operating system APIs. Speak naturally in Indian or Global English, and VOXLY translates spoken intent into concrete, validated desktop actions.", COLOR_CYAN),
        ("Zero-Hardcode Philosophy", "Unlike toy prototypes that only know 'Chrome' and 'Notepad', VOXLY inspects your Windows registry, Start Menu shortcuts, and PATH directories dynamically to launch any installed application on your system.", COLOR_EMERALD),
    ]

    cx = Inches(0.8)
    cw = Inches(3.64)
    for title, desc, col in cards_data_s2:
        add_card(s2, cx, Inches(1.8), cw, Inches(4.8), title, border_color=col)
        tb = s2.shapes.add_textbox(cx + Inches(0.25), Inches(2.7), cw - Inches(0.5), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = desc
        p.font.size = Pt(13)
        p.font.color.rgb = COLOR_MUTED
        p.font.name = "Segoe UI"
        cx += cw + Inches(0.4)

    # ==========================================================
    # SLIDE 3: Architecture & Technology Stack
    # ==========================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_header(s3, "System Architecture & High-Performance Technology Stack", "Technology Foundation")

    tech_items = [
        ("Core & Runtime", "Python 3.10+", "Reliable multi-threading, asynchronous I/O, subprocess orchestration and WinAPI integration."),
        ("Desktop UI", "CustomTkinter 6.0", "Hardware-accelerated dark theme GUI with animated status orbs, responsive chat, and quick controls."),
        ("Voice Input (STT)", "SpeechRecognition / Whisper", "High-accuracy speech-to-text with Indian and global English acoustic models and ambient noise calibration."),
        ("Voice Output (TTS)", "pyttsx3 (SAPI5)", "Zero-latency offline text-to-speech engine with adjustable cadence and natural pitch."),
        ("Cognitive Brain", "OpenAI GPT-4o / Local", "Structured tool calling and natural intent routing with dual-engine offline rule-based fallback."),
        ("OS & Browser Automation", "Win32 + Playwright", "Direct Windows Shell integration, winreg registry crawler, PyAutoGUI, and Chromium web automation.")
    ]

    col_w = Inches(3.64)
    row_h = Inches(2.25)
    for idx, (title, sub, body) in enumerate(tech_items):
        r = idx // 3
        c = idx % 3
        x = Inches(0.8) + c * (col_w + Inches(0.4))
        y = Inches(1.8) + r * (row_h + Inches(0.35))
        
        card = add_card(s3, x, y, col_w, row_h, title, sub, border_color=COLOR_ACCENT)
        tb = s3.shapes.add_textbox(x + Inches(0.25), y + Inches(1.05), col_w - Inches(0.5), Inches(1.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = body
        p.font.size = Pt(12)
        p.font.color.rgb = COLOR_MUTED
        p.font.name = "Segoe UI"

    # ==========================================================
    # SLIDE 4: Voice & Speech Intelligence
    # ==========================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_header(s4, "Voice Recognition & Conversational Interaction", "Core Feature 1")

    # Left Column: Key Highlights
    add_card(s4, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Speech Input & Audio Pipeline", "Designed for real-world acoustics", COLOR_CYAN)
    tb_left = s4.shapes.add_textbox(Inches(1.05), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_l = tb_left.text_frame
    tf_l.word_wrap = True
    
    bullets_s4 = [
        ("Multi-Accent Robustness:", "Optimized for Indian English ('en-IN') and standard conversational accents with phonetic noise filtering."),
        ("Wake-Word Activation:", "Hands-free background detection for 'Hey Voxly' / 'Voxly' with low CPU consumption."),
        ("Push-to-Talk & Auto-Listen:", "Click-to-talk mic button or hands-free continuous conversational turn-taking."),
        ("Real-Time Ambient Calibration:", "Dynamically adjusts energy thresholds to filter out fan hum, background AC, and laptop typing."),
        ("Visual Orb Indicator:", "Visual canvas state machine displaying Idle, Listening (green wave), Thinking (purple pulse), and Speaking (blue equalizer).")
    ]
    for i, (b_title, b_desc) in enumerate(bullets_s4):
        p = tf_l.paragraphs[0] if i == 0 else tf_l.add_paragraph()
        run1 = p.add_run()
        run1.text = "• " + b_title + " "
        run1.font.bold = True
        run1.font.color.rgb = COLOR_WHITE
        run1.font.size = Pt(12)
        run2 = p.add_run()
        run2.text = b_desc + "\n"
        run2.font.color.rgb = COLOR_MUTED
        run2.font.size = Pt(12)

    # Right Column: Voice Output & State Machine
    add_card(s4, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), "Speech Synthesis & Audio States", "Low-latency vocal feedback", COLOR_PURPLE)
    tb_right = s4.shapes.add_textbox(Inches(7.15), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_r = tb_right.text_frame
    tf_r.word_wrap = True

    states_s4 = [
        ("Offline Voice Generation:", "Powered by Windows SAPI5 (pyttsx3); works instantly without requiring an internet connection."),
        ("Thread-Safe Speech Queue:", "Dedicated worker thread prevents speech operations from freezing the interactive CustomTkinter UI."),
        ("Adjustable Speech Rates:", "Customizable words-per-minute (130-220 WPM) and volume balancing directly from user preferences."),
        ("Instant Interruption Support:", "Stop and pause voice output immediately when new user voice triggers arrive."),
        ("Multimodal Synchrony:", "Spoken feedback is delivered synchronously with rich action cards in the desktop chat feed.")
    ]
    for i, (b_title, b_desc) in enumerate(states_s4):
        p = tf_r.paragraphs[0] if i == 0 else tf_r.add_paragraph()
        run1 = p.add_run()
        run1.text = "✔ " + b_title + " "
        run1.font.bold = True
        run1.font.color.rgb = COLOR_WHITE
        run1.font.size = Pt(12)
        run2 = p.add_run()
        run2.text = b_desc + "\n"
        run2.font.color.rgb = COLOR_MUTED
        run2.font.size = Pt(12)

    # ==========================================================
    # SLIDE 5: Dynamic Windows Application Launcher
    # ==========================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_header(s5, "Dynamic Windows Application Discovery & Launching", "Core Feature 2")

    col_data_s5 = [
        ("Multi-Source Registry Crawler", "Discovers Apps Everywhere", [
            "Scans Windows Start Menu .lnk shortcuts (User + All Users).",
            "Queries Windows Registry App Paths (HKLM & HKCU).",
            "Enumerates Installed Program Uninstall keys.",
            "Inspects system environment PATH for CLI/GUI binaries."
        ], COLOR_EMERALD),
        ("Fuzzy Matching & Alias Resolution", "Understands Natural Variations", [
            "Resolves colloquial names (e.g. 'VS Code' -> 'Visual Studio Code').",
            "Maps aliases: 'word' -> MS Word, 'excel' -> Excel, 'cmd' -> Terminal.",
            "Levenshtein distance matching for typo tolerance.",
            "Asks for clarification when multiple matches are found."
        ], COLOR_CYAN),
        ("Intelligent Process Execution", "Verified & Failure-Resistant", [
            "Uses subprocess.Popen with detached process flags.",
            "Verifies actual process launch via psutil inspection.",
            "Never falsely reports success if an app fails to start.",
            "Brings already running applications to the foreground."
        ], COLOR_PURPLE)
    ]

    cx = Inches(0.8)
    cw = Inches(3.64)
    for title, subtitle, bullets, col in col_data_s5:
        add_card(s5, cx, Inches(1.8), cw, Inches(4.8), title, subtitle, border_color=col)
        tb = s5.shapes.add_textbox(cx + Inches(0.25), Inches(2.7), cw - Inches(0.5), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + b + "\n"
            p.font.size = Pt(12)
            p.font.color.rgb = COLOR_MUTED
            p.font.name = "Segoe UI"
        cx += cw + Inches(0.4)

    # ==========================================================
    # SLIDE 6: Deep File Explorer & Nested Folder Navigation
    # ==========================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_header(s6, "Deep File & Folder Traversal ('Local Disk D -> Study Folder')", "Core Feature 3")

    # Left: The User Scenario
    add_card(s6, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Real-World Complex Path Traversal", "Handles deep multi-layered navigation", COLOR_CYAN)
    tb_scen = s6.shapes.add_textbox(Inches(1.05), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_scen = tb_scen.text_frame
    tf_scen.word_wrap = True

    scenario_text = [
        ("The Exact User Request Handled:", ""),
        ("'Open file explorer, then local disk D, then III Year, then CN, then PDF, then UNIT-4'", "COLOR_CYAN"),
        ("How VOXLY Solves It:", "VOXLY decomposes the spoken chain into hierarchical drive & directory milestones."),
        ("Fuzzy Segment Matching:", "Resolves abbreviation 'CN' to 'Computer Networks', matches Roman numerals ('III Year' vs '3rd Year')."),
        ("Native Windows Shell Execution:", "Opens files directly in default system readers (e.g., Adobe Acrobat, Edge PDF, VLC) via os.startfile.")
    ]
    for i, (heading, body) in enumerate(scenario_text):
        p = tf_scen.paragraphs[0] if i == 0 else tf_scen.add_paragraph()
        if body == "COLOR_CYAN":
            p.text = f"\n\"{heading}\"\n"
            p.font.bold = True
            p.font.color.rgb = COLOR_CYAN
            p.font.size = Pt(13)
        else:
            run1 = p.add_run()
            run1.text = "▸ " + heading + " "
            run1.font.bold = True
            run1.font.color.rgb = COLOR_WHITE
            run1.font.size = Pt(12)
            run2 = p.add_run()
            run2.text = body + "\n"
            run2.font.color.rgb = COLOR_MUTED
            run2.font.size = Pt(12)

    # Right: Technical File Capabilities
    add_card(s6, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), "File Operations & Safety Engine", "Complete filesystem management", COLOR_EMERALD)
    tb_fc = s6.shapes.add_textbox(Inches(7.15), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_fc = tb_fc.text_frame
    tf_fc.word_wrap = True

    fc_features = [
        ("Drive Discovery:", "Automatically scans available drive letters (C:\\, D:\\, E:\\) and mounts."),
        ("User Standard Folders:", "Instant access to Downloads, Documents, Desktop, Pictures, Music, and Videos."),
        ("Fast Recursive File Search:", "Finds documents, PDFs, or scripts like 'assignment.py' anywhere in target drives."),
        ("Safe File/Folder Creation:", "Creates directories and writes text files with path sanitization."),
        ("Destructive Action Protection:", "Dangerous operations (e.g. bulk delete, format) are blocked or require explicit user confirmation.")
    ]
    for i, (b_title, b_desc) in enumerate(fc_features):
        p = tf_fc.paragraphs[0] if i == 0 else tf_fc.add_paragraph()
        run1 = p.add_run()
        run1.text = "• " + b_title + " "
        run1.font.bold = True
        run1.font.color.rgb = COLOR_WHITE
        run1.font.size = Pt(12)
        run2 = p.add_run()
        run2.text = b_desc + "\n"
        run2.font.color.rgb = COLOR_MUTED
        run2.font.size = Pt(12)

    # ==========================================================
    # SLIDE 7: System Operations & Desktop Automation
    # ==========================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_header(s7, "Native OS Controls & Desktop Hardware Automation", "Core Feature 4")

    sys_ops = [
        ("Volume & Audio Control", "Master Volume Adjustment", "Turn volume up/down, mute, unmute, or set specific percentage levels using Windows multimedia APIs."),
        ("Instant Screen Capture", "Pictures / Screenshots", "Captures full-screen screenshots timestamped and saved automatically to your Pictures\\Screenshots folder."),
        ("Workstation Security", "Instant Lock & Privacy", "Immediately locks your Windows laptop screen ('Lock my PC') via user32.dll LockWorkStation."),
        ("Task Manager & Utilities", "Diagnostic Tools", "Launches Task Manager, Command Prompt, PowerShell, or Settings instantly via hands-free voice."),
        ("Window Management", "Desktop Organization", "Minimize all active windows, show desktop, or cycle applications via PyAutoGUI desktop automation."),
        ("Battery & System Diagnostics", "Resource Awareness", "Queries real-time CPU, RAM, and battery charging status via psutil system metrics.")
    ]

    for idx, (title, sub, body) in enumerate(sys_ops):
        r = idx // 3
        c = idx % 3
        x = Inches(0.8) + c * (col_w + Inches(0.4))
        y = Inches(1.8) + r * (row_h + Inches(0.35))
        
        card = add_card(s7, x, y, col_w, row_h, title, sub, border_color=COLOR_CYAN)
        tb = s7.shapes.add_textbox(x + Inches(0.25), y + Inches(1.05), col_w - Inches(0.5), Inches(1.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = body
        p.font.size = Pt(12)
        p.font.color.rgb = COLOR_MUTED
        p.font.name = "Segoe UI"

    # ==========================================================
    # SLIDE 8: Cognitive Brain & Dual-Engine Intent Architecture
    # ==========================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_header(s8, "Dual-Engine AI Brain: Offline Speed + LLM Reasoning", "Core Feature 5")

    # Left Box: Tier 1 Fast Engine
    add_card(s8, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Tier 1: Offline Rule & Regex Engine", "0ms Latency • Works without API Key", COLOR_EMERALD)
    tb_t1 = s8.shapes.add_textbox(Inches(1.05), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_t1 = tb_t1.text_frame
    tf_t1.word_wrap = True

    t1_points = [
        ("Instant Local Routing:", "Directly catches 'open chrome', 'volume up', 'take screenshot', and folder paths in under 5 milliseconds."),
        ("Zero Network Dependency:", "Core laptop control functions flawlessly even when disconnected from the internet."),
        ("7 Intent Categories:", "Classifies commands into App Launch, System Ops, Browser, Files, Education, Automation, and General Chat."),
        ("Colloquial Phrasing Support:", "Recognizes Indian English idioms like 'can you open my...', 'please put volume', 'show my files'.")
    ]
    for i, (b_title, b_desc) in enumerate(t1_points):
        p = tf_t1.paragraphs[0] if i == 0 else tf_t1.add_paragraph()
        run1 = p.add_run()
        run1.text = "⚡ " + b_title + " "
        run1.font.bold = True
        run1.font.color.rgb = COLOR_WHITE
        run1.font.size = Pt(12)
        run2 = p.add_run()
        run2.text = b_desc + "\n"
        run2.font.color.rgb = COLOR_MUTED
        run2.font.size = Pt(12)

    # Right Box: Tier 2 OpenAI Tool Calling
    add_card(s8, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), "Tier 2: OpenAI Function Calling Brain", "Deep Reasoning • Complex Multi-Step Tasks", COLOR_PURPLE)
    tb_t2 = s8.shapes.add_textbox(Inches(7.15), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_t2 = tb_t2.text_frame
    tf_t2.word_wrap = True

    t2_points = [
        ("Structured Tool Invocations:", "Uses OpenAI Function Calling JSON schemas for strict parameter validation."),
        ("ChatGPT-Grade Intelligence:", "Answers educational questions, writes code (e.g. binary search), explains operating system deadlocks."),
        ("Conversation Memory:", "Maintains multi-turn context (last 10 turns) stored persistently in SQLite."),
        ("Autonomous Fallback:", "If OpenAI API is unavailable or rate-limited, smoothly degrades to Tier 1 without crashing.")
    ]
    for i, (b_title, b_desc) in enumerate(t2_points):
        p = tf_t2.paragraphs[0] if i == 0 else tf_t2.add_paragraph()
        run1 = p.add_run()
        run1.text = "🧠 " + b_title + " "
        run1.font.bold = True
        run1.font.color.rgb = COLOR_WHITE
        run1.font.size = Pt(12)
        run2 = p.add_run()
        run2.text = b_desc + "\n"
        run2.font.color.rgb = COLOR_MUTED
        run2.font.size = Pt(12)

    # ==========================================================
    # SLIDE 9: Web & Browser Automation (Playwright)
    # ==========================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_header(s9, "Intelligent Web Search & Playwright Browser Automation", "Core Feature 6")

    cards_data_s9 = [
        ("Automated Web Search", "Google & Knowledge Lookups", [
            "Voice: 'Search for machine learning tutorials on Google'.",
            "Launches browser, navigates to Google search, and inputs query automatically.",
            "Supports headless and full desktop browser interaction."
        ], COLOR_CYAN),
        ("YouTube Media Streaming", "Direct Video Playback", [
            "Voice: 'Play relaxing lo-fi beats on YouTube'.",
            "Opens YouTube directly with the queried keywords.",
            "Bypasses repetitive manual clicks and typing."
        ], COLOR_PURPLE),
        ("Resilient Browser Strategy", "Dual Engine Execution", [
            "Primary: Playwright headless/headed Chromium driver.",
            "Secondary Fallback: Windows default browser dispatch (webbrowser API).",
            "Guarantees seamless navigation regardless of environment."
        ], COLOR_EMERALD)
    ]

    cx = Inches(0.8)
    cw = Inches(3.64)
    for title, subtitle, bullets, col in cards_data_s9:
        add_card(s9, cx, Inches(1.8), cw, Inches(4.8), title, subtitle, border_color=col)
        tb = s9.shapes.add_textbox(cx + Inches(0.25), Inches(2.7), cw - Inches(0.5), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + b + "\n"
            p.font.size = Pt(12)
            p.font.color.rgb = COLOR_MUTED
            p.font.name = "Segoe UI"
        cx += cw + Inches(0.4)

    # ==========================================================
    # SLIDE 10: Modern User Interface & Visual Experience
    # ==========================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_header(s10, "Modern UI/UX: CustomTkinter Desktop & Web Companion", "User Experience")

    ui_highlights = [
        ("Hardware-Accelerated Dark Theme", "CustomTkinter", "Sleek slate-and-violet aesthetic tailored for developer and power-user workflows with zero lag."),
        ("Interactive Voice Orb", "Real-Time Canvas", "Smooth animated visual feedback: Breathing (Idle), Waves (Listening), Vortex (Thinking), Equalizer (Speaking)."),
        ("Dual Interaction Modes", "Voice & Keyboard", "Switch effortlessly between hands-free voice commands and high-speed keyboard text queries."),
        ("Rich System Action Cards", "Visual Confirmations", "Displays styled status cards for launched applications, captured screenshot thumbnails, and opened files."),
        ("SQLite History & Sessions", "Persistent Knowledge", "Stores chats, launched app frequency, and user settings across restarts in local voxly.db."),
        ("Companion Web Interface", "React + Vite Frontend", "Preserved web companion with live audio wave bars, hands-free auto-listening, and quick actions.")
    ]

    for idx, (title, sub, body) in enumerate(ui_highlights):
        r = idx // 3
        c = idx % 3
        x = Inches(0.8) + c * (col_w + Inches(0.4))
        y = Inches(1.8) + r * (row_h + Inches(0.35))
        
        card = add_card(s10, x, y, col_w, row_h, title, sub, border_color=COLOR_ACCENT)
        tb = s10.shapes.add_textbox(x + Inches(0.25), y + Inches(1.05), col_w - Inches(0.5), Inches(1.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = body
        p.font.size = Pt(12)
        p.font.color.rgb = COLOR_MUTED
        p.font.name = "Segoe UI"

    # ==========================================================
    # SLIDE 11: Security, Safety & Enterprise Privacy
    # ==========================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11)
    add_header(s11, "Security Architecture: Enterprise-Grade Privacy & Guardrails", "Security & Reliability")

    sec_cards = [
        ("Strict Sandbox & Path Safety", "Prevent Unauthorized Traversal", [
            "Path sanitization guards against directory traversal attacks (../).",
            "Critical system directories (System32, Windows) are protected from accidental modifications.",
            "File read/write boundaries restricted to safe user-approved spaces."
        ], COLOR_CYAN),
        ("Destructive Action Confirmation", "Human-in-the-Loop Safeguards", [
            "System shutdown, reboot, or bulk deletion strictly requires explicit verbal or UI confirmation.",
            "No silent or unprompted destruction of user assets.",
            "Comprehensive action logging in logs/voxly.log."
        ], COLOR_EMERALD),
        ("Local Privacy & Key Protection", "Your Data Stays on Your Device", [
            "API keys stored exclusively in local .env files; never written to SQLite.",
            "All speech processing audio buffers cleared immediately after inference.",
            "Zero third-party telemetry or hidden background data uploads."
        ], COLOR_PURPLE)
    ]

    cx = Inches(0.8)
    cw = Inches(3.64)
    for title, subtitle, bullets, col in sec_cards:
        add_card(s11, cx, Inches(1.8), cw, Inches(4.8), title, subtitle, border_color=col)
        tb = s11.shapes.add_textbox(cx + Inches(0.25), Inches(2.7), cw - Inches(0.5), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + b + "\n"
            p.font.size = Pt(12)
            p.font.color.rgb = COLOR_MUTED
            p.font.name = "Segoe UI"
        cx += cw + Inches(0.4)

    # ==========================================================
    # SLIDE 12: Summary & Future Roadmap
    # ==========================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12)
    add_header(s12, "Project Summary & Future Innovation Roadmap", "Conclusion & Next Steps")

    # Left: What is Delivered Today
    add_card(s12, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Delivered Capabilities (Production Ready)", "Fully operational desktop assistant", COLOR_CYAN)
    tb_del = s12.shapes.add_textbox(Inches(1.05), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_del = tb_del.text_frame
    tf_del.word_wrap = True

    del_items = [
        ("Autonomous App Launcher:", "Dynamically discovers and opens any installed Windows software."),
        ("Multi-Level Path Traversal:", "Opens deep drives (D:\\III Year\\CN\\PDF\\UNIT-4) via natural speech."),
        ("Voice Synthesis & STT:", "Indian English acoustic adaptation with real-time audio visualization."),
        ("Full OS Desktop Automation:", "Volume, screenshots, screen lock, task manager, and window control."),
        ("ChatGPT Reasoning Brain:", "Multi-turn context with structured tool calling and offline fallback.")
    ]
    for i, (b_title, b_desc) in enumerate(del_items):
        p = tf_del.paragraphs[0] if i == 0 else tf_del.add_paragraph()
        run1 = p.add_run()
        run1.text = "✔ " + b_title + " "
        run1.font.bold = True
        run1.font.color.rgb = COLOR_WHITE
        run1.font.size = Pt(12)
        run2 = p.add_run()
        run2.text = b_desc + "\n"
        run2.font.color.rgb = COLOR_MUTED
        run2.font.size = Pt(12)

    # Right: Future Evolution
    add_card(s12, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), "Strategic Roadmap", "Upcoming enhancements", COLOR_PURPLE)
    tb_fut = s12.shapes.add_textbox(Inches(7.15), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_fut = tb_fut.text_frame
    tf_fut.word_wrap = True

    future_items = [
        ("Local Offline LLMs:", "Integrating Ollama / Llama-3.2 for 100% offline private neural reasoning."),
        ("Computer Vision Multimodal Input:", "Real-time screen understanding (analyze diagrams or code on screen via voice)."),
        ("Custom Plugin Ecosystem:", "User-defined Python plugins for Photoshop, Blender, and specialized IDE tools."),
        ("Contextual Proactivity:", "Intelligent calendar meeting reminders and automated workspace preparation."),
        ("PyInstaller 1-Click Binary:", "Single .exe deployment installer for friction-free setup on any Windows machine.")
    ]
    for i, (b_title, b_desc) in enumerate(future_items):
        p = tf_fut.paragraphs[0] if i == 0 else tf_fut.add_paragraph()
        run1 = p.add_run()
        run1.text = "🚀 " + b_title + " "
        run1.font.bold = True
        run1.font.color.rgb = COLOR_WHITE
        run1.font.size = Pt(12)
        run2 = p.add_run()
        run2.text = b_desc + "\n"
        run2.font.color.rgb = COLOR_MUTED
        run2.font.size = Pt(12)

    # Save
    prs.save(output_path)
    print(f"Presentation saved successfully at: {output_path}")

if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "C:\\voicebridge\\VOXLY_AI_Desktop_Assistant_Presentation.pptx"
    create_deck(out)
