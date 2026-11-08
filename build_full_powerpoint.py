"""
VOXLY AI - Comprehensive Feature Presentation Generator
Generates a 14-slide executive presentation covering every feature built in the AI app:
- Voice input / speech recognition (English only)
- Live Audio Waveform Visualizer
- Continuous hands-free auto-listening mode
- Real Gemini & OpenAI AI intelligence
- Dynamic Windows Application Launcher
- Deep nested drive & folder navigation (Drive D -> III Year -> CN -> PDF -> UNIT-4)
- Persistent LocalStorage and SQLite caching
- System controls (Volume, Screenshot, Lock, Task Manager)
- Dual Desktop (CustomTkinter) + Web Companion (React) architecture
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def build_presentation(output_path="C:\\voicebridge\\VOXLY_Features_Complete_Presentation.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Corporate Dark / Cyber Tech Colors
    BG_DARK = RGBColor(11, 15, 25)           # #0B0F19 - Main Background
    CARD_BG = RGBColor(22, 30, 49)           # #161E31 - Card Background
    CARD_BORDER = RGBColor(40, 53, 84)       # Border slate
    ACCENT_CYAN = RGBColor(6, 182, 212)      # #06B6D4 - Cyan accent
    ACCENT_VIOLET = RGBColor(139, 92, 246)   # #8B5CF6 - Violet accent
    ACCENT_GREEN = RGBColor(16, 185, 129)    # #10B981 - Emerald accent
    ACCENT_AMBER = RGBColor(245, 158, 11)    # #F59E0B - Amber accent
    TEXT_WHITE = RGBColor(248, 250, 252)     # #F8FAFC - Headings
    TEXT_MUTED = RGBColor(148, 163, 184)     # #94A3B8 - Descriptions
    TEXT_BODY = RGBColor(203, 213, 225)      # #CBD5E1 - List text

    blank_layout = prs.slide_layouts[6]

    def set_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_DARK
        bg.line.fill.background()

    def add_header(slide, title, category="VOXLY AI ASSISTANT • FEATURE PRESENTATION"):
        # Category Kicker
        kbox = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.32))
        ktf = kbox.text_frame
        ktf.word_wrap = True
        kp = ktf.paragraphs[0]
        kp.text = category.upper()
        kp.font.size = Pt(11)
        kp.font.bold = True
        kp.font.color.rgb = ACCENT_CYAN
        kp.font.name = "Segoe UI"

        # Main Title
        tbox = slide.shapes.add_textbox(Inches(0.8), Inches(0.68), Inches(11.7), Inches(0.65))
        ttf = tbox.text_frame
        ttf.word_wrap = True
        tp = ttf.paragraphs[0]
        tp.text = title
        tp.font.size = Pt(26)
        tp.font.bold = True
        tp.font.color.rgb = TEXT_WHITE
        tp.font.name = "Segoe UI"

        # Accent Rule
        divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.38), Inches(11.733), Inches(0.025))
        divider.fill.solid()
        divider.fill.fore_color.rgb = CARD_BORDER
        divider.line.fill.background()

    def add_card(slide, left, top, width, height, title, subtitle=None, border_color=CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = border_color
        card.line.width = Pt(1.5)

        tb = slide.shapes.add_textbox(left + Inches(0.28), top + Inches(0.22), width - Inches(0.56), Inches(0.75))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = border_color if border_color != CARD_BORDER else ACCENT_CYAN
        p.font.name = "Segoe UI"

        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.font.size = Pt(11)
            p2.font.color.rgb = TEXT_MUTED
            p2.font.name = "Segoe UI"

        return card

    # ==========================================================
    # SLIDE 1: Title Slide
    # ==========================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_bg(s1)

    # Feature Badge Pill
    badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.2), Inches(4.5), Inches(0.45))
    badge.fill.solid()
    badge.fill.fore_color.rgb = CARD_BG
    badge.line.color.rgb = ACCENT_CYAN
    badge.line.width = Pt(1.2)
    bp = badge.text_frame.paragraphs[0]
    bp.text = "✦ COMPREHENSIVE AI FEATURE SHOWCASE"
    bp.font.size = Pt(11)
    bp.font.bold = True
    bp.font.color.rgb = ACCENT_CYAN
    bp.font.name = "Segoe UI"
    bp.alignment = PP_ALIGN.CENTER

    # Main Title Box
    title_box = s1.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.3), Inches(2.2))
    tf1 = title_box.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "VOXLY AI ASSISTANT"
    p1.font.size = Pt(56)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE
    p1.font.name = "Segoe UI"

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "Next-Generation Intelligent Voice Control & Deep Windows Automation"
    p1_sub.font.size = Pt(22)
    p1_sub.font.bold = True
    p1_sub.font.color.rgb = ACCENT_CYAN
    p1_sub.font.name = "Segoe UI"

    # Description
    desc_box = s1.shapes.add_textbox(Inches(1.0), Inches(4.3), Inches(11.3), Inches(1.2))
    dtf = desc_box.text_frame
    dtf.word_wrap = True
    dp = dtf.paragraphs[0]
    dp.text = "VOXLY is a fully functional desktop & web AI assistant that listens to voice commands, reasons through conversational AI, dynamically launches installed Windows applications, navigates deep file directories, and automates OS tasks effortlessly without manual intervention."
    dp.font.size = Pt(14)
    dp.font.color.rgb = TEXT_MUTED
    dp.font.name = "Segoe UI"

    # Key Features Strip
    key_tags = ["Live Waveform Visualizer", "Continuous Hands-Free Mode", "Zero-Hardcode App Launcher", "Deep Folder Navigation", "Real Gemini/OpenAI Brain", "Native OS Automation"]
    tx = Inches(1.0)
    for tag in key_tags:
        tw = Inches(len(tag) * 0.115 + 0.5)
        pill = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tx, Inches(5.8), tw, Inches(0.45))
        pill.fill.solid()
        pill.fill.fore_color.rgb = CARD_BG
        pill.line.color.rgb = CARD_BORDER
        pp = pill.text_frame.paragraphs[0]
        pp.text = tag
        pp.font.size = Pt(10)
        pp.font.bold = True
        pp.font.color.rgb = TEXT_WHITE
        pp.font.name = "Segoe UI"
        pp.alignment = PP_ALIGN.CENTER
        tx += tw + Inches(0.18)

    # ==========================================================
    # SLIDE 2: Table of Features
    # ==========================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_bg(s2)
    add_header(s2, "Complete Capabilities & Feature Spectrum", "Feature Matrix")

    feats = [
        ("1. Pure English Voice Engine", "Clear natural voice recognition and English-only speech synthesis with no language mixing.", ACCENT_CYAN),
        ("2. Live Waveform Visualizer", "Animated audio wave bars that react in real-time while you speak or when VOXLY speaks.", ACCENT_VIOLET),
        ("3. Continuous Hands-Free Mode", "Auto-listens after each reply, enabling a seamless back-and-forth dialogue without mouse clicks.", ACCENT_GREEN),
        ("4. Real Gemini & OpenAI Brain", "Connects to live LLM APIs to answer questions, write code, explain deadlocks, and reason.", ACCENT_AMBER),
        ("5. Dynamic Windows App Launcher", "Scans Windows Registry & Start Menu to open Chrome, VS Code, Spotify, Word, WhatsApp, etc.", ACCENT_CYAN),
        ("6. Deep Drive & Folder Traversal", "Speaks a nested path ('Drive D -> III Year -> CN -> PDF -> UNIT-4') and it opens directly.", ACCENT_VIOLET),
        ("7. Native Hardware & OS Controls", "System volume adjustment, instant full-screen screenshots, lock screen, and Task Manager.", ACCENT_GREEN),
        ("8. Persistent Storage & History", "Remembers conversation sessions, recent action cards, and preferences across reloads.", ACCENT_AMBER)
    ]

    col_w = Inches(5.6)
    row_h = Inches(1.15)
    for i, (title, desc, color) in enumerate(feats):
        r = i % 4
        c = i // 4
        x = Inches(0.8) + c * (col_w + Inches(0.53))
        y = Inches(1.7) + r * (row_h + Inches(0.2))

        box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, col_w, row_h)
        box.fill.solid()
        box.fill.fore_color.rgb = CARD_BG
        box.line.color.rgb = color
        box.line.width = Pt(1.5)

        tb = s2.shapes.add_textbox(x + Inches(0.2), y + Inches(0.15), col_w - Inches(0.4), row_h - Inches(0.3))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = color
        p.font.name = "Segoe UI"

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(11)
        p2.font.color.rgb = TEXT_BODY
        p2.font.name = "Segoe UI"

    # ==========================================================
    # SLIDE 3: Pure English Voice & Accent Optimization
    # ==========================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_bg(s3)
    add_header(s3, "Pure English Voice Recognition & Synthesis", "Feature Spotlight 1")

    add_card(s3, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Speech Recognition (STT)", "Fast, accurate and accent-adaptive", ACCENT_CYAN)
    tb = s3.shapes.add_textbox(Inches(1.05), Inches(2.7), Inches(5.1), Inches(3.6))
    tf = tb.text_frame
    tf.word_wrap = True
    stt_items = [
        ("Clean English-Only Acoustic Model:", "Strict English language processing ('en-IN' and 'en-US') with no unwanted mixed dialect confusion."),
        ("Acoustic Noise Filtering:", "Dynamic calibration adjusts for background ambient noise such as laptop cooling fans, room echoes, or typing."),
        ("Microphone Level Auto-Gain:", "Adapts speech threshold based on distance from the laptop microphone."),
        ("Immediate Speech Transcription:", "Instantaneous speech-to-text conversion displayed directly in the chat feed.")
    ]
    for i, (head, text) in enumerate(stt_items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r1 = p.add_run()
        r1.text = "✔ " + head + " "
        r1.font.bold = True
        r1.font.color.rgb = TEXT_WHITE
        r1.font.size = Pt(12)
        r2 = p.add_run()
        r2.text = text + "\n"
        r2.font.color.rgb = TEXT_MUTED
        r2.font.size = Pt(12)

    add_card(s3, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), "Natural Speech Synthesis (TTS)", "Zero-latency vocal replies", ACCENT_VIOLET)
    tb2 = s3.shapes.add_textbox(Inches(7.15), Inches(2.7), Inches(5.1), Inches(3.6))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    tts_items = [
        ("Pure English Pronunciation:", "Voice output is configured strictly to English male/female voices with crisp diction."),
        ("Offline Windows SAPI5 Integration:", "Speaks instantly via pyttsx3 without requiring network bandwidth or cloud API latency."),
        ("Synchronous Action Confirmation:", "Speaks clear confirmation (e.g., 'Launching Google Chrome for you now') alongside visual action cards."),
        ("Adjustable Speaking Pace:", "Customizable speech speed (140 to 200 WPM) and volume balancing for comfortable listening.")
    ]
    for i, (head, text) in enumerate(tts_items):
        p = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
        r1 = p.add_run()
        r1.text = "✦ " + head + " "
        r1.font.bold = True
        r1.font.color.rgb = TEXT_WHITE
        r1.font.size = Pt(12)
        r2 = p.add_run()
        r2.text = text + "\n"
        r2.font.color.rgb = TEXT_MUTED
        r2.font.size = Pt(12)

    # ==========================================================
    # SLIDE 4: Live Audio Waveform Visualizer
    # ==========================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_bg(s4)
    add_header(s4, "Live Audio Waveform Visualizer", "Feature Spotlight 2")

    cards_s4 = [
        ("Dynamic Audio Reactive Bars", "Real-Time Frequency Analysis", [
            "Displays multi-bar equalizer animation reacting live to user voice amplitude.",
            "Visual feedback lets you know immediately that your microphone is capturing audio cleanly.",
            "Zero lag visual synchronization rendered through hardware-accelerated canvas/CSS."
        ], ACCENT_CYAN),
        ("Multi-State Status Indication", "Visual State Machine", [
            "IDLE STATE: Gentle breathing glow indicating ready standby.",
            "LISTENING STATE: Vibrant reactive soundwave bars undulating with user speech.",
            "THINKING STATE: Concentric orbital vortex during AI intent processing.",
            "SPEAKING STATE: Symmetrical audio pulse bars matching assistant voice output."
        ], ACCENT_VIOLET),
        ("Enhanced Interaction Feedback", "No More Guesswork", [
            "Instantly reveals if room acoustics are too quiet or speech has ended.",
            "Provides an intuitive, sci-fi desktop assistant experience akin to modern AI interfaces.",
            "Works across both the CustomTkinter desktop window and the web companion."
        ], ACCENT_GREEN)
    ]

    cx = Inches(0.8)
    cw = Inches(3.64)
    for title, subtitle, bullets, col in cards_s4:
        add_card(s4, cx, Inches(1.8), cw, Inches(4.8), title, subtitle, border_color=col)
        tb = s4.shapes.add_textbox(cx + Inches(0.25), Inches(2.7), cw - Inches(0.5), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + b + "\n"
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_MUTED
            p.font.name = "Segoe UI"
        cx += cw + Inches(0.4)

    # ==========================================================
    # SLIDE 5: Continuous Hands-Free Conversation Mode
    # ==========================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_bg(s5)
    add_header(s5, "Continuous Hands-Free Conversation Mode", "Feature Spotlight 3")

    add_card(s5, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Hands-Free Auto-Listening Loop", "True conversational back-and-forth", ACCENT_GREEN)
    tb_hf = s5.shapes.add_textbox(Inches(1.05), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_hf = tb_hf.text_frame
    tf_hf.word_wrap = True

    hf_bullets = [
        ("No Repetitive Button Clicks:", "Once activated, VOXLY automatically re-opens the microphone after speaking its answer so you can reply naturally."),
        ("Fluid Multi-Turn Dialogue:", "Ask a follow-up question or issue a secondary command without touching your mouse or keyboard."),
        ("Smart Silence Detection:", "Recognizes pauses (0.8s threshold) to conclude voice queries without cutting off midway through a sentence."),
        ("Toggle Anytime:", "Enable hands-free mode for full voice productivity, or switch back to push-to-talk mic button at your convenience.")
    ]
    for i, (head, text) in enumerate(hf_bullets):
        p = tf_hf.paragraphs[0] if i == 0 else tf_hf.add_paragraph()
        r1 = p.add_run()
        r1.text = "▸ " + head + " "
        r1.font.bold = True
        r1.font.color.rgb = TEXT_WHITE
        r1.font.size = Pt(12)
        r2 = p.add_run()
        r2.text = text + "\n"
        r2.font.color.rgb = TEXT_MUTED
        r2.font.size = Pt(12)

    add_card(s5, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), "Real-World Hands-Free Scenarios", "Effortless multitasking", ACCENT_AMBER)
    tb_sc = s5.shapes.add_textbox(Inches(7.15), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_sc = tb_sc.text_frame
    tf_sc.word_wrap = True

    scenarios = [
        ("Studying & Researching:", "User: 'Explain operating system deadlocks.' -> VOXLY explains -> User: 'Give me a real-world example.' -> VOXLY continues seamlessly."),
        ("Coding & Development:", "User: 'Open VS Code.' -> VOXLY launches VS Code -> User: 'Now search Google for Python binary search.' -> VOXLY opens the search."),
        ("Media & Relaxation:", "User: 'Play lo-fi beats on YouTube.' -> VOXLY launches video -> User: 'Increase volume.' -> VOXLY raises volume by 10%.")
    ]
    for i, (title, dialog) in enumerate(scenarios):
        p = tf_sc.paragraphs[0] if i == 0 else tf_sc.add_paragraph()
        r1 = p.add_run()
        r1.text = "💬 " + title + "\n"
        r1.font.bold = True
        r1.font.color.rgb = ACCENT_AMBER
        r1.font.size = Pt(12)
        r2 = p.add_run()
        r2.text = dialog + "\n\n"
        r2.font.color.rgb = TEXT_BODY
        r2.font.size = Pt(11)

    # ==========================================================
    # SLIDE 6: Real Gemini & OpenAI Brain
    # ==========================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_bg(s6)
    add_header(s6, "Intelligent AI Brain: Real Gemini & OpenAI Integration", "Feature Spotlight 4")

    cards_s6 = [
        ("Real AI API Connection", "Answers Anything in the World", [
            "Direct integration with OpenAI GPT-4o / Google Gemini APIs.",
            "Answers complex educational topics, writes clean algorithms, explains technical concepts.",
            "Full knowledge retention across conversation sessions."
        ], ACCENT_AMBER),
        ("Structured Tool & Function Calling", "Translates Thoughts to Actions", [
            "Validates user instructions into structured JSON schemas.",
            "Distinguishes between general chat, system commands, and application launches.",
            "Never hallucinates fake operations; executes verified Python methods only."
        ], ACCENT_CYAN),
        ("Offline Fallback Resilience", "Always Available On-Device", [
            "If internet connectivity drops or API key is absent, Tier 1 local rule engine steps in.",
            "App launching, file opening, and volume control work 100% offline without network dependency.",
            "Smooth transition ensures zero crashes or interruptions."
        ], ACCENT_VIOLET)
    ]

    cx = Inches(0.8)
    cw = Inches(3.64)
    for title, subtitle, bullets, col in cards_s6:
        add_card(s6, cx, Inches(1.8), cw, Inches(4.8), title, subtitle, border_color=col)
        tb = s6.shapes.add_textbox(cx + Inches(0.25), Inches(2.7), cw - Inches(0.5), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + b + "\n"
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_MUTED
            p.font.name = "Segoe UI"
        cx += cw + Inches(0.4)

    # ==========================================================
    # SLIDE 7: Dynamic Windows Application Launcher
    # ==========================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_bg(s7)
    add_header(s7, "Dynamic Windows Application Discovery & Launching", "Feature Spotlight 5")

    add_card(s7, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "No Hardcoding — Discovers Any App", "Deep Windows Ecosystem Crawler", ACCENT_CYAN)
    tb_app = s7.shapes.add_textbox(Inches(1.05), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_app = tb_app.text_frame
    tf_app.word_wrap = True

    app_points = [
        ("Start Menu Shortcut Crawler:", "Scans both User and All-Users Start Menu shortcut repositories (.lnk files) to discover installed apps."),
        ("Windows Registry Inspection:", "Queries HKLM and HKCU 'App Paths' and software registrations for direct executable binaries."),
        ("System PATH Resolution:", "Automatically locates CLI and GUI tools registered in the Windows environment PATH."),
        ("Alias & Colloquial Mapping:", "Understands 'VS Code', 'Chrome', 'WhatsApp', 'Spotify', 'Word', 'Excel', 'Calculator', 'IDLE', and custom names.")
    ]
    for i, (head, text) in enumerate(app_points):
        p = tf_app.paragraphs[0] if i == 0 else tf_app.add_paragraph()
        r1 = p.add_run()
        r1.text = "⚡ " + head + " "
        r1.font.bold = True
        r1.font.color.rgb = TEXT_WHITE
        r1.font.size = Pt(12)
        r2 = p.add_run()
        r2.text = text + "\n"
        r2.font.color.rgb = TEXT_MUTED
        r2.font.size = Pt(12)

    add_card(s7, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), "Reliable Process Verification", "Never pretends an app launched", ACCENT_GREEN)
    tb_app2 = s7.shapes.add_textbox(Inches(7.15), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_app2 = tb_app2.text_frame
    tf_app2.word_wrap = True

    app_points2 = [
        ("Independent Detached Execution:", "Launches applications via subprocess with creation flags that decouple them from VOXLY's lifecycle."),
        ("Real Process Verification:", "Inspects Windows running processes using psutil; confirms when the application's PID is live."),
        ("Already-Running Handling:", "If the requested application is already active, VOXLY informs the user and brings it to focus."),
        ("Ambiguity Clarification:", "If multiple programs match (e.g. 'Python 3.10' vs 'Python IDLE'), VOXLY prompts the user for clarification.")
    ]
    for i, (head, text) in enumerate(app_points2):
        p = tf_app2.paragraphs[0] if i == 0 else tf_app2.add_paragraph()
        r1 = p.add_run()
        r1.text = "✔ " + head + " "
        r1.font.bold = True
        r1.font.color.rgb = TEXT_WHITE
        r1.font.size = Pt(12)
        r2 = p.add_run()
        r2.text = text + "\n"
        r2.font.color.rgb = TEXT_MUTED
        r2.font.size = Pt(12)

    # ==========================================================
    # SLIDE 8: Deep Nested Folder & Drive Traversal
    # ==========================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_bg(s8)
    add_header(s8, "Deep Nested Drive & Folder Navigation", "Feature Spotlight 6")

    # Top Hero Card with the exact user prompt
    hero = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.75), Inches(11.733), Inches(1.4))
    hero.fill.solid()
    hero.fill.fore_color.rgb = CARD_BG
    hero.line.color.rgb = ACCENT_CYAN
    hero.line.width = Pt(1.5)
    
    htb = s8.shapes.add_textbox(Inches(1.1), Inches(1.85), Inches(11.1), Inches(1.1))
    htf = htb.text_frame
    htf.word_wrap = True
    hp1 = htf.paragraphs[0]
    hp1.text = "ACCOMPLISHING COMPLEX SPOKEN PATH NAVIGATION:"
    hp1.font.size = Pt(11)
    hp1.font.bold = True
    hp1.font.color.rgb = ACCENT_CYAN
    hp1.font.name = "Segoe UI"

    hp2 = htf.add_paragraph()
    hp2.text = "\"Open file explorer, then local disk D, then III Year, then CN, then PDF, then UNIT-4\""
    hp2.font.size = Pt(17)
    hp2.font.bold = True
    hp2.font.color.rgb = TEXT_WHITE
    hp2.font.name = "Segoe UI"

    # 3 Breakdown Cards
    b_cards = [
        ("1. Multi-Segment Parser", "Decodes Spoken Chains", [
            "Decomposes long verbal instructions into hierarchical directory targets.",
            "Detects drive letters ('Local Disk D' -> D:\\) across all mounted drives.",
            "Sequential step execution matches human mental folders."
        ], ACCENT_CYAN),
        ("2. Fuzzy Acronym Matching", "Handles Study Acronyms", [
            "Resolves 'CN' to 'Computer Networks' or 'Computer Network'.",
            "Translates Roman numerals ('III Year' <-> '3rd Year' <-> 'Third Year').",
            "Tolerant of minor acoustic pronunciation differences."
        ], ACCENT_VIOLET),
        ("3. Native File Launch", "Opens in Default Reader", [
            "Opens folder in File Explorer or opens target PDF in Adobe / Edge.",
            "Uses Windows Shell os.startfile for guaranteed compatibility.",
            "Displays visual badge confirmation card with file path in UI."
        ], ACCENT_GREEN)
    ]

    cx = Inches(0.8)
    cw = Inches(3.64)
    for title, subtitle, bullets, col in b_cards:
        add_card(s8, cx, Inches(3.4), cw, Inches(3.3), title, subtitle, border_color=col)
        tb = s8.shapes.add_textbox(cx + Inches(0.2), Inches(4.2), cw - Inches(0.4), Inches(2.3))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "▸ " + b + "\n"
            p.font.size = Pt(11)
            p.font.color.rgb = TEXT_MUTED
            p.font.name = "Segoe UI"
        cx += cw + Inches(0.4)

    # ==========================================================
    # SLIDE 9: System Operations & Hardware Automation
    # ==========================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_bg(s9)
    add_header(s9, "Native Hardware & Desktop System Controls", "Feature Spotlight 7")

    sys_cards = [
        ("Master Volume Controls", "Windows Audio API", [
            "Say 'Volume up', 'Volume down', or 'Mute audio'.",
            "Precision volume stepping via Windows multimedia keyboard events.",
            "Auditory confirmation spoken back after volume change."
        ], ACCENT_CYAN),
        ("Instant Screen Capture", "Full-Resolution Screenshots", [
            "Voice: 'Take a screenshot'.",
            "Captures entire monitor screen and saves directly to Pictures\\Screenshots.",
            "Displays timestamped confirmation thumbnail card in chat feed."
        ], ACCENT_VIOLET),
        ("Workstation Security", "Instant PC Lock", [
            "Say 'Lock my PC' or 'Lock screen'.",
            "Triggers user32.dll LockWorkStation immediately for privacy.",
            "Requires no touching of the physical keyboard."
        ], ACCENT_GREEN),
        ("Utility Launchers", "Core Diagnostics", [
            "Launch Windows Task Manager, Command Prompt, PowerShell, or Settings.",
            "Quick access to system monitoring and terminal workflows.",
            "Handles minimized and background window states."
        ], ACCENT_AMBER),
        ("Window Management", "Desktop Organization", [
            "Minimize all windows, show desktop, or bring focus to an application.",
            "Automated keyboard shortcut dispatch via PyAutoGUI.",
            "Reduces window clutter during busy multi-app sessions."
        ], ACCENT_CYAN),
        ("System Resource Monitoring", "Hardware Awareness", [
            "Query CPU usage, available RAM, and battery charging status.",
            "Alerts on low battery or excessive background processes.",
            "Powered by psutil low-level system diagnostic hooks."
        ], ACCENT_VIOLET)
    ]

    for idx, (title, sub, bullets, col) in enumerate(sys_cards):
        r = idx // 3
        c = idx % 3
        x = Inches(0.8) + c * (col_w + Inches(0.4))
        y = Inches(1.75) + r * (row_h + Inches(1.3))

        add_card(s9, x, y, col_w, Inches(2.35), title, sub, border_color=col)
        tb = s9.shapes.add_textbox(x + Inches(0.2), y + Inches(0.95), col_w - Inches(0.4), Inches(1.3))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + b + "\n"
            p.font.size = Pt(11)
            p.font.color.rgb = TEXT_MUTED
            p.font.name = "Segoe UI"

    # ==========================================================
    # SLIDE 10: Persistent LocalStorage & SQLite Storage
    # ==========================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_bg(s10)
    add_header(s10, "Permanent LocalStorage & SQLite Persistence", "Feature Spotlight 8")

    cards_s10 = [
        ("Permanent LocalStorage", "Web Companion State", [
            "Saves all recent commands, action history, and reminders across page reloads.",
            "Maintains user preferences (voice selection, speed, hands-free toggle).",
            "Zero cloud lock-in: data stays 100% inside your browser environment."
        ], ACCENT_CYAN),
        ("SQLite Database Engine", "Desktop Session History", [
            "Stores chat messages, multi-turn contexts, and session IDs in voxly.db.",
            "Thread-safe WAL (Write-Ahead Logging) prevents database locks.",
            "Allows recalling previous chats and reviewing earlier instructions."
        ], ACCENT_VIOLET),
        ("Smart Application Cache", "Instant Startup Matching", [
            "Caches discovered Windows apps and records frequency of launches.",
            "Frequently used apps launch in under 5 milliseconds.",
            "Periodically refreshes cache when new software is installed."
        ], ACCENT_GREEN)
    ]

    cx = Inches(0.8)
    cw = Inches(3.64)
    for title, subtitle, bullets, col in cards_s10:
        add_card(s10, cx, Inches(1.8), cw, Inches(4.8), title, subtitle, border_color=col)
        tb = s10.shapes.add_textbox(cx + Inches(0.25), Inches(2.7), cw - Inches(0.5), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + b + "\n"
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_MUTED
            p.font.name = "Segoe UI"
        cx += cw + Inches(0.4)

    # ==========================================================
    # SLIDE 11: Web & Browser Automation (Playwright)
    # ==========================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_bg(s11)
    add_header(s11, "Browser Automation: Google Search & YouTube Media", "Feature Spotlight 9")

    cards_s11 = [
        ("Automated Google Search", "Playwright & Browser Engine", [
            "Voice: 'Search for Python machine learning tutorials'.",
            "Automatically opens default browser, executes query, and surfaces top results.",
            "Hands-free web research while working in code or documents."
        ], ACCENT_CYAN),
        ("Direct YouTube Media Streaming", "Instant Video & Music Playback", [
            "Voice: 'Play lo-fi study beats on YouTube'.",
            "Navigates straight to YouTube search results and begins streaming.",
            "Eliminates repetitive clicking, tab opening, and manual keyboard typing."
        ], ACCENT_AMBER),
        ("Universal URL Navigation", "Any Web Destination", [
            "Voice: 'Open GitHub' or 'Go to Wikipedia'.",
            "Resolves popular web domains and launches them in default browser.",
            "Gracefully handles headless browser mode and standard browser dispatch."
        ], ACCENT_GREEN)
    ]

    cx = Inches(0.8)
    cw = Inches(3.64)
    for title, subtitle, bullets, col in cards_s11:
        add_card(s11, cx, Inches(1.8), cw, Inches(4.8), title, subtitle, border_color=col)
        tb = s11.shapes.add_textbox(cx + Inches(0.25), Inches(2.7), cw - Inches(0.5), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + b + "\n"
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_MUTED
            p.font.name = "Segoe UI"
        cx += cw + Inches(0.4)

    # ==========================================================
    # SLIDE 12: Dual Desktop & Web Companion Architecture
    # ==========================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_bg(s12)
    add_header(s12, "Dual Interface System: Native Desktop + Web Companion", "Architecture")

    add_card(s12, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Native CustomTkinter Desktop App", "High-speed Windows desktop executable", ACCENT_CYAN)
    tb_dt = s12.shapes.add_textbox(Inches(1.05), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_dt = tb_dt.text_frame
    tf_dt.word_wrap = True
    dt_items = [
        ("Native Windows Performance:", "Built with Python 3.10 and CustomTkinter 6.0; zero browser memory overhead."),
        ("Offline Voice & System APIs:", "Direct access to winreg, user32.dll, SAPI5, and psutil for sub-millisecond hardware operations."),
        ("Integrated Animated Canvas Orb:", "Hardware-accelerated visual state transitions (Idle, Listening, Thinking, Speaking)."),
        ("Single-Click Launch:", "Starts via Start-Voxly.bat or desktop icon.")
    ]
    for i, (head, text) in enumerate(dt_items):
        p = tf_dt.paragraphs[0] if i == 0 else tf_dt.add_paragraph()
        r1 = p.add_run()
        r1.text = "🖥 " + head + " "
        r1.font.bold = True
        r1.font.color.rgb = TEXT_WHITE
        r1.font.size = Pt(12)
        r2 = p.add_run()
        r2.text = text + "\n"
        r2.font.color.rgb = TEXT_MUTED
        r2.font.size = Pt(12)

    add_card(s12, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), "Interactive React + Vite Web Companion", "Modern web browser dashboard", ACCENT_VIOLET)
    tb_wb = s12.shapes.add_textbox(Inches(7.15), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_wb = tb_wb.text_frame
    tf_wb.word_wrap = True
    wb_items = [
        ("Real-Time Waveform Visualizer:", "Dynamic SVG audio wave equalizer displaying active microphone frequency bars."),
        ("Permanent LocalStorage Saving:", "Maintains chat recents, reminders, and user settings across reloads."),
        ("Seamless Node.js System Bridge:", "Vite development plugin bridge triggers actual Windows commands from web UI."),
        ("Fully Preserved & Operational:", "Both desktop and web apps operate synergistically on the same laptop.")
    ]
    for i, (head, text) in enumerate(wb_items):
        p = tf_wb.paragraphs[0] if i == 0 else tf_wb.add_paragraph()
        r1 = p.add_run()
        r1.text = "🌐 " + head + " "
        r1.font.bold = True
        r1.font.color.rgb = TEXT_WHITE
        r1.font.size = Pt(12)
        r2 = p.add_run()
        r2.text = text + "\n"
        r2.font.color.rgb = TEXT_MUTED
        r2.font.size = Pt(12)

    # ==========================================================
    # SLIDE 13: Enterprise Safety & Privacy Architecture
    # ==========================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_bg(s13)
    add_header(s13, "Enterprise Safety, Privacy & Guardrails", "Security Architecture")

    cards_s13 = [
        ("Path Sanitization & Isolation", "Zero Directory Traversal", [
            "Validates every file and folder target before opening or writing.",
            "Rejects parent directory escapes (e.g., ../../) and unauthorized disk scans.",
            "Protects critical Windows system folders (C:\\Windows, System32)."
        ], ACCENT_CYAN),
        ("Destructive Action Confirmation", "Explicit User Consent", [
            "Irreversible commands (system shutdown, format, bulk file deletion) are strictly blocked or require explicit confirmation.",
            "Never performs destructive actions silently.",
            "All actions logged in logs\\voxly.log for complete auditability."
        ], ACCENT_GREEN),
        ("Local Privacy & Secure Config", "Your Data Stays With You", [
            "API keys are stored exclusively in local .env files; never stored in databases.",
            "Microphone audio buffers are discarded immediately after recognition.",
            "Zero background telemetry or hidden cloud profiling."
        ], ACCENT_VIOLET)
    ]

    cx = Inches(0.8)
    cw = Inches(3.64)
    for title, subtitle, bullets, col in cards_s13:
        add_card(s13, cx, Inches(1.8), cw, Inches(4.8), title, subtitle, border_color=col)
        tb = s13.shapes.add_textbox(cx + Inches(0.25), Inches(2.7), cw - Inches(0.5), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + b + "\n"
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_MUTED
            p.font.name = "Segoe UI"
        cx += cw + Inches(0.4)

    # ==========================================================
    # SLIDE 14: Conclusion & Value Summary
    # ==========================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_bg(s14)
    add_header(s14, "Conclusion & Project Value Summary", "Executive Summary")

    add_card(s14, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Delivered Innovation & Impact", "True AI desktop copilot", ACCENT_CYAN)
    tb_c1 = s14.shapes.add_textbox(Inches(1.05), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_c1 = tb_c1.text_frame
    tf_c1.word_wrap = True
    c1_items = [
        ("Hands-Free Efficiency:", "Eliminates repetitive manual clicks; launch any software, play music, or open files via speech."),
        ("No Hardcoded Limits:", "Autonomous Windows crawler finds all current and future installed desktop applications."),
        ("Complex Path Traversal:", "Handles academic and personal deep directory structures ('D:\\III Year\\CN\\PDF\\UNIT-4')."),
        ("Dual Intelligence Layer:", "Combines zero-latency offline controls with ChatGPT/Gemini conversational depth.")
    ]
    for i, (head, text) in enumerate(c1_items):
        p = tf_c1.paragraphs[0] if i == 0 else tf_c1.add_paragraph()
        r1 = p.add_run()
        r1.text = "✦ " + head + " "
        r1.font.bold = True
        r1.font.color.rgb = TEXT_WHITE
        r1.font.size = Pt(12)
        r2 = p.add_run()
        r2.text = text + "\n"
        r2.font.color.rgb = TEXT_MUTED
        r2.font.size = Pt(12)

    add_card(s14, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), "Future Expansion Horizon", "Upcoming enhancements", ACCENT_AMBER)
    tb_c2 = s14.shapes.add_textbox(Inches(7.15), Inches(2.7), Inches(5.1), Inches(3.6))
    tf_c2 = tb_c2.text_frame
    tf_c2.word_wrap = True
    c2_items = [
        ("Local Edge LLM Inference:", "Integration with Ollama (Llama-3.2 / Mistral) for 100% offline conversational reasoning."),
        ("Screen Vision Multimodal AI:", "Voice-driven screen inspection (e.g., 'Look at this diagram on my screen and explain it')."),
        ("Custom Macro Automation:", "Record voice routines like 'Start study session' to open VS Code, PDF notes, and lo-fi music in one go."),
        ("Single-Click Executable (.exe):", "Standalone PyInstaller distribution package for seamless installation on any Windows device.")
    ]
    for i, (head, text) in enumerate(c2_items):
        p = tf_c2.paragraphs[0] if i == 0 else tf_c2.add_paragraph()
        r1 = p.add_run()
        r1.text = "🚀 " + head + " "
        r1.font.bold = True
        r1.font.color.rgb = TEXT_WHITE
        r1.font.size = Pt(12)
        r2 = p.add_run()
        r2.text = text + "\n"
        r2.font.color.rgb = TEXT_MUTED
        r2.font.size = Pt(12)

    prs.save(output_path)
    print(f"Presentation successfully created at: {output_path}")

if __name__ == "__main__":
    build_presentation()
