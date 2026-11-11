import json
import queue
import re
import subprocess
import sys
import threading
import uuid
from datetime import datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from markdown_it import MarkdownIt


ROOT = Path(__file__).resolve().parent
STAGES = {
    "search": ("Search", "Finds current sources"),
    "reader": ("Reader", "Extracts source evidence"),
    "writer": ("Writer", "Builds the report"),
    "critic": ("Critic", "Evaluates the result"),
}
PAGES = ("Home", "New Research", "Research Workspace", "Research History", "About")

# Run the existing pipeline in an isolated process so its print events can be
# observed without patching shared backend state in the Streamlit process.
WORKER = r"""
import json, sys, traceback

def emit(kind, **values):
    print("__RESEARCH_EVENT__" + json.dumps({"kind": kind, **values}, default=str), flush=True)

active = None
try:
    import pipeline
    payload = json.loads(sys.stdin.readline())
    starts = {
        "Running search agent...": "search",
        "Running scraper agent for extracting resources...": "reader",
        "Running writer chain for generating content...": "writer",
        "Running critic chain for evaluating the report...": "critic",
    }
    outputs = {
        "Search Results:": "search",
        "Scrape Results:": "reader",
        "Final Report Generated:": "writer",
        "Critique of the Report:": "critic",
    }
    def observe(*args, **kwargs):
        global active
        text = str(args[0]).strip() if args else ""
        if text in starts:
            active = starts[text]
            emit("stage", stage=active, status="active")
        elif text in outputs:
            emit("stage", stage=outputs[text], status="completed",
                 content=str(args[1]) if len(args) > 1 else "")
            active = None
    pipeline.print = observe
    result = pipeline.run_research_pipeline(payload["topic"])
    emit("result", result=result)
except Exception as exc:
    emit("error", stage=active, message=f"{type(exc).__name__}: {exc}")
    traceback.print_exc(file=sys.stderr)
    sys.exit(1)
"""

ACTIVITIES = {
    "Brain Teasers": [
        ("What has keys but cannot open locks?", ["A piano", "A map", "A clock"], 0,
         "A piano has keys, but they make music rather than open locks."),
        ("What comes once in a minute, twice in a moment, and never in a thousand years?",
         ["The letter M", "A second", "A pause"], 0, "The answer is the letter M."),
    ],
    "Quick Quizzes": [
        ("Which planet is the largest in our solar system?", ["Mars", "Jupiter", "Venus"], 1,
         "Jupiter is the largest planet in our solar system."),
        ("Which method tests a hypothesis by comparing outcomes?",
         ["An experiment", "A metaphor", "A guess"], 0, "Experiments compare observations under defined conditions."),
        ("What does a primary source provide?", ["Firsthand evidence", "Only opinions", "A search ranking"], 0,
         "Primary sources provide firsthand or original evidence."),
    ],
    "Guess the Concept": [
        ("I let computers learn patterns from examples rather than explicit rules. What am I?",
         ["Machine learning", "A spreadsheet", "A firewall"], 0, "Machine learning models learn patterns from data."),
        ("I am a network of interlinked documents accessible over the internet. What am I?",
         ["The World Wide Web", "A database row", "A processor"], 0, "The web is a system of interlinked resources."),
    ],
}
FACTS = {
    "Knowledge Bites": [
        "Light from the Sun takes about eight minutes to reach Earth.",
        "Octopuses have three hearts.",
        "A day on Venus is longer than its year.",
        "The human brain consumes a substantial share of the body's energy.",
    ],
    "Research Insights": [
        "A publication date and an event date can be different. Check both when assessing news.",
        "Independent sources can strengthen a finding; several articles repeating one source do not.",
        "A missing date is a limitation, not evidence that a source is current.",
        "A good research question specifies what evidence would change the answer.",
    ],
}


def initialize():
    defaults = {
        "page": "Home", "theme": "Dark", "history": [], "job": None,
        "activity": "Brain Teasers", "activity_index": {}, "answers": {}, "score": 0,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    if "nav_pending" in st.session_state:
        st.session_state.page = st.session_state.pop("nav_pending")


def style():
    dark = st.session_state.theme == "Dark"
    bg, panel, text, muted, border, surface, accent = (
        ("#070a13", "#0e1423", "#f0f3ff", "#a4acc2", "#273044", "#131b2b", "#9b8cff") if dark else
        ("#f4f6fb", "#ffffff", "#1a2032", "#58647a", "#dbe0ed", "#edf0f8", "#5744c5")
    )
    light_overrides = "" if dark else """
    [data-testid="stBaseButton-secondary"], button[kind="secondary"] {
      background:#ffffff; color:#1a2032; border-color:#cbd3e1;
    }
    [data-testid="stBaseButton-secondary"]:disabled, button[kind="secondary"]:disabled {
      background:#edf0f8; color:#687488;
    }
    [data-testid="stAlertContainer"] {background:#e8f1fb; color:#1a2032;}
    [data-testid="stAlertContainer"]:has([data-testid="stAlertContentSuccess"]) {background:#e5f5ee;}
    [data-testid="stAlertContainer"]:has([data-testid="stAlertContentWarning"]) {background:#fff3dc;}
    [data-testid="stAlertContainer"]:has([data-testid="stAlertContentError"]) {background:#fdebed;}
    [data-testid="stAlertContainer"] p {color:#1a2032;}
    [data-testid="stExpander"] summary,
    [data-testid="stExpander"] summary:hover {background:#f7f8fc; color:#1a2032;}
    [data-testid="stExpander"] details {border-color:#cbd3e1;}
    [data-testid="stProgressBarTrack"] {background:#dce2ed;}
    [data-testid="stProgressBarTrack"] > div {background:linear-gradient(90deg,#7462e9,#36b5bc);}
    [data-testid="stSelectbox"] [role="group"] {background:#edf0f8; border-color:#cbd3e1;}
    [data-testid="stSelectbox"] [role="combobox"],
    [data-testid="stSelectbox"] [role="group"] button {color:#1a2032; background:transparent;}
    [data-testid="stSelectboxVirtualDropdown"], [data-baseweb="popover"] {
      background:#ffffff; color:#1a2032; border:1px solid #cbd3e1;
    }
    [data-testid="stSelectboxVirtualDropdown"] [role="option"] {color:#1a2032;}
    [data-testid="stSelectboxVirtualDropdown"] [role="option"][data-focused],
    [data-testid="stSelectboxVirtualDropdown"] [role="option"]:hover {background:#edeafb;}
    [data-testid="stMain"] [data-testid="stRadioOption"] > div > div:first-child {
      background:#ffffff; border:1px solid #8794ab;
    }
    [data-testid="stMain"] [data-testid="stRadioOption"]:has(input:checked) > div > div:first-child {
      background:#6756cf; border-color:#6756cf;
    }
    [data-testid="stCheckbox"] label > span + div {background:#ffffff; border-color:#8794ab;}
    [data-testid="stCheckbox"] label:has(input:checked) > span + div {background:#6756cf; border-color:#6756cf;}
    [data-testid="stMarkdownContainer"] a {color:#5143b3;}
    [data-testid="stMarkdownContainer"] table {color:#1a2032;}
    [data-testid="stMarkdownContainer"] th,
    [data-testid="stMarkdownContainer"] td {background:#ffffff; border-color:#cbd3e1; color:#1a2032;}
    [data-testid="stMarkdownContainer"] code {background:#edf0f8; color:#1a2032;}
    [data-testid="stMarkdownContainer"] pre {background:#edf0f8; color:#1a2032;}
    [data-testid="stTab"][aria-selected="true"] p {color:#5744c5;}
    [data-testid="stTab"] [data-testid="stMarkdownContainer"] p {opacity:1;}
    [data-testid="stCaptionContainer"] {opacity:1;}
    [data-testid="stCaptionContainer"] p {color:#58647a;}
    .status.active {color:#5744c5;} .status.completed {color:#16715e;} .status.failed {color:#ae3243;}
    """
    st.markdown(f"""
    <style>
    .stApp {{
      --ink:{text}; --muted:{muted}; --panel:{panel}; --surface:{surface};
      --line:{border}; --accent:{accent};
      background:radial-gradient(ellipse at 85% -15%, #7563ee24, transparent 42%),
                 radial-gradient(ellipse at 0% 40%, #22b7bc0d, transparent 35%), {bg};
      color:var(--ink); font-family:"Segoe UI Variable","Segoe UI",sans-serif;
    }}
    [data-testid="stHeader"] {{background:{bg};}}
    [data-testid="stSidebar"] {{background:{panel}; border-right:1px solid var(--line);}}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3,
    .stApp h4, .stApp li {{color:var(--ink);}}
    .block-container {{max-width:1250px; padding-top:4.5rem; padding-bottom:3rem;}}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {{gap:.8rem;}}
    [data-testid="stSidebar"] [role="radiogroup"] label {{
      border:1px solid transparent; border-radius:12px; padding:.55rem .7rem;
      transition:background .2s,border-color .2s;
    }}
    [data-testid="stSidebar"] [role="radiogroup"] label:hover {{background:var(--surface);}}
    [data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {{
      background:var(--surface); border-color:var(--line);
    }}
    [data-testid="stSidebar"] [role="radiogroup"] label > div > div:first-child {{display:none;}}
    [data-testid="stSidebar"] hr {{border-color:var(--line);}}
    .brand {{display:flex; align-items:center; gap:12px; margin:4px 0 20px;}}
    .brand-icon {{display:grid; place-items:center; width:38px; height:38px; border-radius:13px;
      color:#fff; font-size:21px; background:linear-gradient(135deg,#8274ff,#26bcca);
      box-shadow:0 8px 22px #7064ff45;}}
    .logo-glyph {{position:relative; font-family:"Segoe UI",sans-serif; font-size:1em; font-weight:800;
      line-height:1; letter-spacing:-.08em; color:#fff;}}
    .logo-glyph:after {{content:""; position:absolute; width:.2em; height:.2em; border-radius:50%;
      background:#b9ffef; right:-.18em; top:.08em;}}
    .brand-name {{font-weight:750; font-size:1.3rem; letter-spacing:-.05em; color:var(--ink);}}
    .brand-sub {{color:var(--muted); font-size:.66rem; letter-spacing:.14em; text-transform:uppercase;}}
    .eyebrow {{color:var(--accent); letter-spacing:.15em; font-weight:700; font-size:.68rem; text-transform:uppercase;}}
    .hero {{position:relative; overflow:hidden; padding:3.3rem; border:1px solid var(--line);
      border-radius:28px; background:linear-gradient(125deg,var(--panel),{surface});
      box-shadow:0 24px 70px #00000012;}}
    .hero:before {{content:""; position:absolute; width:460px; height:460px; right:-130px; top:-190px;
      background:radial-gradient(circle,#8170ff2c,transparent 68%); pointer-events:none;}}
    .hero-grid {{display:grid; grid-template-columns:1.2fr .8fr; align-items:center; gap:1rem;}}
    .hero h1 {{font-size:clamp(2.6rem,5vw,4.5rem); line-height:1.04; letter-spacing:-.06em;
      margin:1rem 0 1.2rem; font-weight:730;}}
    .hero h1 [data-heading-text] > span {{color:var(--accent);}}
    .hero h1 {{padding:0;}}
    .hero p {{max-width:520px; font-size:1.03rem; line-height:1.7; color:var(--muted)!important;}}
    .hero-meta {{display:flex; gap:1.3rem; flex-wrap:wrap; margin-top:2rem;
      font-size:.7rem; letter-spacing:.09em; text-transform:uppercase; color:var(--muted);}}
    .hero-meta b {{color:var(--ink);}}
    .orbit {{width:min(100%,290px); aspect-ratio:1; position:relative; margin:auto;}}
    .orbit-ring {{position:absolute; inset:12%; border:1px solid #8e80e74d; border-radius:50%;}}
    .orbit-ring.second {{inset:23%; border-color:#48cfd54d;}}
    .orbit-ring.third {{inset:2%; border-style:dashed; animation:spin 35s linear infinite;}}
    .orbit-core {{position:absolute; inset:37%; border-radius:50%; display:grid; place-items:center;
      background:linear-gradient(135deg,#8170ef,#29b9c9); color:white; font-size:2rem;
      box-shadow:0 0 50px #8476ff70;}}
    .orbit-dot {{position:absolute; width:13px; height:13px; border-radius:50%;
      background:#65dce0; box-shadow:0 0 18px #65dce0; top:14%; left:24%;}}
    .orbit-dot.two {{top:68%; left:83%; background:#aa9bff; box-shadow:0 0 18px #aa9bff;}}
    .orbit-dot.three {{top:81%; left:20%; width:8px; height:8px;}}
    .muted {{color:var(--muted)!important;}}
    .section-label {{display:flex; align-items:center; gap:.75rem; margin:2rem 0 .4rem;}}
    .section-label:after {{content:""; height:1px; background:var(--line); flex:1;}}
    .tile {{position:relative; padding:1.35rem; min-height:150px; border:1px solid var(--line); border-radius:18px;
      background:var(--panel); transition:transform .2s ease,border-color .2s ease,box-shadow .2s ease;}}
    .tile:hover {{transform:translateY(-4px); border-color:var(--accent); box-shadow:0 12px 30px #7465e41c;}}
    .tile strong {{font-size:1.05rem; color:var(--ink);}}
    .tile p {{margin:.7rem 0 .35rem;}}
    .tile-icon {{float:right; color:var(--accent); font-size:1.2rem;}}
    .agent-flow {{display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:1rem 0 1.4rem;}}
    .agent-node {{position:relative; padding:1.35rem; min-height:160px; border:1px solid var(--line);
      border-radius:19px; background:var(--panel); transition:border-color .3s,transform .3s,box-shadow .3s;}}
    .agent-node:hover {{transform:translateY(-3px);}}
    .agent-node.is-active {{border-color:#8f82ff; box-shadow:0 0 0 1px #8f82ff55,0 0 32px #8f82ff28;}}
    .agent-node.is-completed {{border-color:#41bca456;}}
    .agent-node.is-failed {{border-color:#f17d87;}}
    .agent-node:not(:last-child):after {{content:""; position:absolute; right:-13px; top:49%; width:12px; height:1px;
      background:linear-gradient(90deg,#8b7bfd,#44cbd1); z-index:1;}}
    .agent-symbol {{display:grid; place-items:center; width:40px; height:40px; border-radius:13px;
      background:linear-gradient(145deg,#8c7bf32c,#46c8d31b); color:var(--accent); font-size:1.2rem;}}
    .agent-name {{display:block; color:var(--ink); font-weight:700; margin:.9rem 0 .3rem;}}
    .agent-desc {{color:var(--muted); font-size:.79rem; line-height:1.4;}}
    .status {{font-size:.67rem; font-weight:750; text-transform:uppercase; letter-spacing:.1em;
      margin-top:1rem; display:flex; align-items:center; gap:.45rem;}}
    .status:before {{content:""; display:block; width:7px; height:7px; border-radius:50%; background:currentColor;}}
    .status.active {{color:#8676ed;}} .status.active:before {{animation:pulse 1.4s infinite;}}
    .status.completed {{color:#32a68e;}} .status.failed {{color:#dc6473;}} .status.waiting {{color:var(--muted);}}
    .footer {{text-align:center; color:var(--muted); font-size:.72rem; letter-spacing:.02em;
      margin-top:4rem; padding-top:1.3rem; border-top:1px solid var(--line);}}
    div.stButton > button, div.stDownloadButton > button, div.stFormSubmitButton > button {{
      border-radius:12px; font-weight:650; border-color:var(--line); transition:transform .18s,box-shadow .18s;
    }}
    div.stButton > button:hover, div.stDownloadButton > button:hover, div.stFormSubmitButton > button:hover {{
      transform:translateY(-2px); box-shadow:0 10px 24px #7869ec2e; border-color:var(--accent);
    }}
    button[kind="primary"], button[kind="primaryFormSubmit"] {{
      background:linear-gradient(110deg,#7462e9,#586de5); border:0; color:white;
    }}
    button[kind="primary"] p, button[kind="primaryFormSubmit"] p {{color:white;}}
    [data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea {{
      color:var(--ink); background:var(--surface); border:1px solid var(--line); border-radius:10px;
    }}
    [data-testid="stTextInput"] input::placeholder, [data-testid="stTextArea"] textarea::placeholder {{
      color:var(--muted); opacity:1;
    }}
    [data-testid="stTextInput"] [data-baseweb="input"], [data-testid="stTextArea"] [data-baseweb="textarea"] {{
      background:var(--surface); border-color:var(--line); border-radius:10px;
    }}
    [data-baseweb="select"] > div, [data-baseweb="popover"] {{color:var(--ink); background:var(--surface);}}
    [data-testid="stForm"] {{border:1px solid var(--line); border-radius:20px; padding:1.5rem; background:var(--panel);}}
    [data-testid="stExpander"] {{border:1px solid var(--line); border-radius:14px; background:var(--panel);}}
    [data-testid="stCaptionContainer"] p {{color:var(--muted);}}
    [data-testid="stProgressBar"] > div > div > div {{background:linear-gradient(90deg,#8170ef,#38c6cb);}}
    {light_overrides}
    @keyframes pulse {{50% {{opacity:.45; transform:scale(1.45);}}}}
    @keyframes spin {{to {{transform:rotate(360deg);}}}}
    @media(max-width:1100px) {{
      .hero-grid {{grid-template-columns:1fr;}}
      .hero-grid > div:first-child {{position:relative; z-index:1;}}
      .orbit {{position:absolute; width:180px; right:1rem; top:3rem; opacity:.2; pointer-events:none;}}
    }}
    @media(max-width:800px) {{
      .hero {{padding:2rem;}} .hero-grid {{grid-template-columns:1fr;}}
      .agent-flow {{grid-template-columns:repeat(2,minmax(0,1fr));}}
      .agent-node:nth-child(2):after {{display:none;}}
    }}
    @media(max-width:480px) {{
      .block-container {{padding-top:4rem;}} .hero {{padding:1.5rem;}}
      .agent-flow {{grid-template-columns:1fr;}} .agent-node:not(:last-child):after {{display:none;}}
    }}
    @media (prefers-reduced-motion: reduce) {{*, *:before, *:after {{animation-duration:.01ms!important; transition:none!important;}}}}
    </style>
    """, unsafe_allow_html=True)


def run_worker(job):
    try:
        process = subprocess.Popen(
            [sys.executable, "-u", "-c", WORKER], cwd=str(ROOT),
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", bufsize=1,
        )
        assert process.stdin is not None and process.stdout is not None
        process.stdin.write(json.dumps({"topic": job["input"]}) + "\n")
        process.stdin.close()
        got_terminal_event = False
        output = []
        for line in process.stdout:
            if line.startswith("__RESEARCH_EVENT__"):
                event = json.loads(line[len("__RESEARCH_EVENT__"):])
                got_terminal_event |= event["kind"] in ("result", "error")
                job["events"].put(event)
            else:
                output.append(line.rstrip())
        code = process.wait()
        if not got_terminal_event:
            detail = "\n".join(output[-8:]) or f"Worker exited with code {code}."
            job["events"].put({"kind": "error", "message": detail})
    except Exception as exc:
        job["events"].put({"kind": "error", "message": f"{type(exc).__name__}: {exc}"})


def start_research(topic, context):
    job = {
        "topic": topic, "input": topic if not context else f"{topic}\n\nAdditional research instructions: {context}",
        "started": datetime.now().astimezone().isoformat(timespec="seconds"),
        "status": "running", "stages": {name: "waiting" for name in STAGES},
        "outputs": {}, "timeline": [], "events": queue.Queue(), "saved": False,
    }
    st.session_state.job = job
    threading.Thread(target=run_worker, args=(job,), daemon=True).start()


def receive_events():
    job = st.session_state.job
    if job is None:
        return
    while True:
        try:
            event = job["events"].get_nowait()
        except queue.Empty:
            break
        kind = event["kind"]
        if kind == "stage":
            name, status = event["stage"], event["status"]
            job["stages"][name] = status
            if status == "completed":
                job["outputs"][name] = event.get("content", "")
            job["timeline"].append((datetime.now().strftime("%H:%M:%S"), name, status))
        elif kind == "error":
            job["status"] = "failed"
            job["error"] = event["message"]
            name = event.get("stage")
            if name in job["stages"]:
                job["stages"][name] = "failed"
                job["timeline"].append((datetime.now().strftime("%H:%M:%S"), name, "failed"))
        elif kind == "result":
            job["status"] = "completed"
            job["result"] = event["result"]
            job["completed"] = datetime.now().astimezone().isoformat(timespec="seconds")
            if not job["saved"]:
                st.session_state.history.insert(0, {
                    "id": uuid.uuid4().hex, "topic": job["topic"],
                    "date": job["completed"], "result": job["result"],
                })
                job["saved"] = True


def go(page):
    st.session_state.nav_pending = page
    st.rerun()


def hyperlink(paragraph, text, url):
    if not url.startswith(("http://", "https://")):
        paragraph.add_run(text)
        return
    relation = paragraph.part.relate_to(
        url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), relation)
    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "167C9D")
    props.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    props.append(underline)
    run.append(props)
    node = OxmlElement("w:t")
    node.text = text
    run.append(node)
    link.append(run)
    paragraph._p.append(link)


def inline(paragraph, children):
    bold = italic = False
    link = None
    for token in children or []:
        if token.type == "strong_open":
            bold = True
        elif token.type == "strong_close":
            bold = False
        elif token.type == "em_open":
            italic = True
        elif token.type == "em_close":
            italic = False
        elif token.type == "link_open":
            link = token.attrGet("href")
        elif token.type == "link_close":
            link = None
        elif token.type in ("text", "code_inline"):
            if link:
                hyperlink(paragraph, token.content, link)
            else:
                parts = re.split(r"(https?://[^\s<>]+)", token.content)
                for part in parts:
                    if part.startswith(("http://", "https://")):
                        url = part.rstrip(".,;)")
                        hyperlink(paragraph, url, url)
                        paragraph.add_run(part[len(url):])
                    else:
                        run = paragraph.add_run(part)
                        run.bold, run.italic = bold, italic
        elif token.type in ("softbreak", "hardbreak"):
            paragraph.add_run("\n")


def add_markdown(doc, text):
    tokens = MarkdownIt("commonmark").enable("table").parse(text or "")
    paragraph = table = None
    row_index = cell_index = -1
    list_types = []
    for token in tokens:
        kind = token.type
        if kind == "bullet_list_open":
            list_types.append("List Bullet")
        elif kind == "ordered_list_open":
            list_types.append("List Number")
        elif kind in ("bullet_list_close", "ordered_list_close"):
            list_types.pop()
        elif kind == "table_open":
            table = doc.add_table(rows=0, cols=0)
            table.style = "Table Grid"
        elif kind == "tr_open":
            row_index += 1
            cell_index = -1
            if table is not None and len(table.columns):
                table.add_row()
        elif kind in ("th_open", "td_open"):
            cell_index += 1
            if table is not None:
                if row_index == 0:
                    table.add_column(Inches(1.5))
                    if len(table.rows) == 0:
                        table.add_row()
                paragraph = table.rows[row_index].cells[cell_index].paragraphs[0]
        elif kind == "table_close":
            table = None
            row_index = -1
        elif kind == "heading_open":
            paragraph = doc.add_heading(level=min(int(token.tag[1]), 3))
        elif kind == "paragraph_open" and table is None:
            paragraph = doc.add_paragraph(style=list_types[-1] if list_types else None)
        elif kind == "inline" and paragraph is not None:
            inline(paragraph, token.children)
        elif kind in ("fence", "code_block"):
            doc.add_paragraph(token.content, style="No Spacing")


def make_docx(record):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Aptos"
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor.from_string("26364A")
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.15
    for level in range(1, 4):
        heading = doc.styles[f"Heading {level}"]
        heading.font.name = "Aptos"
        heading.font.color.rgb = RGBColor.from_string("5143A1")
        heading.paragraph_format.space_before = Pt(14)
    section = doc.sections[0]
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.header.paragraphs[0].text = "LUMEN  /  RESEARCH"
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("LUMEN  |  ")
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    footer._p.append(field)
    doc.add_paragraph("\n\n")
    doc.add_paragraph("LUMEN  /  RESEARCH REPORT", style="Subtitle")
    doc.add_heading(record["topic"], 0)
    doc.add_paragraph(f"Generated {record['date']}")
    doc.add_paragraph("Research brief  •  Sources, findings and independent critique")
    doc.add_page_break()
    add_markdown(doc, record["result"].get("report", ""))
    doc.add_page_break()
    doc.add_heading("Critic evaluation", 1)
    add_markdown(doc, record["result"].get("feedback", ""))
    result = BytesIO()
    doc.save(result)
    return result.getvalue()


def show_report(record, key):
    st.markdown('<div class="eyebrow">Research complete</div>', unsafe_allow_html=True)
    title_col, download_col = st.columns([3, 1])
    with title_col:
        st.title(record["topic"])
    with download_col:
        try:
            document = make_docx(record)
        except Exception as exc:
            st.error(f"Could not create the Word document: {exc}")
        else:
            filename = re.sub(r"[^\w-]+", "-", record["topic"], flags=re.UNICODE).strip("-")[:60]
            st.download_button(
                "↓  Download report", document, file_name=f"{filename or 'research-report'}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=key, type="primary", use_container_width=True,
            )
    st.caption(f"Generated {record['date']} · Report and evaluation from your research pipeline")
    report = record["result"].get("report", "")
    feedback = record["result"].get("feedback", "")
    if not re.search(r"(?im)^#{0,6}\s*(?:\*\*)?executive summary", report):
        st.caption("The generated report did not include a separate executive summary.")
    report_tab, critique_tab = st.tabs(["Research report", "Critic evaluation"])
    with report_tab:
        if report:
            st.markdown(report)
        else:
            st.warning("No report was returned by the pipeline.")
    with critique_tab:
        if feedback:
            st.markdown(feedback)
        else:
            st.warning("No critic feedback was returned by the pipeline.")
def lounge():
    st.markdown("### Research Lounge")
    st.caption("A little curiosity while your research runs. Your task continues independently.")
    activity = st.selectbox("Choose an activity", list(ACTIVITIES) + list(FACTS), key="activity")
    items = ACTIVITIES.get(activity) or FACTS[activity]
    index = st.session_state.activity_index.get(activity, 0) % len(items)
    if activity in ACTIVITIES:
        question, choices, correct, explanation = items[index]
        st.markdown(f"**{question}**")
        answer = st.radio("Your answer", choices, key=f"choice-{activity}-{index}")
        answer_key = f"{activity}-{index}"
        if st.button("Check answer", key=f"check-{answer_key}"):
            if answer_key not in st.session_state.answers:
                is_correct = choices.index(answer) == correct
                st.session_state.answers[answer_key] = is_correct
                if activity == "Quick Quizzes" and is_correct:
                    st.session_state.score += 1
        if answer_key in st.session_state.answers:
            if st.session_state.answers[answer_key]:
                st.success(f"Correct. {explanation}")
            else:
                st.info(f"Not quite. {explanation}")
        if activity == "Quick Quizzes":
            st.caption(f"Quiz score: {st.session_state.score} / {len(ACTIVITIES['Quick Quizzes'])}")
    else:
        st.info(items[index])
    if st.button("Next / skip", key=f"next-{activity}"):
        st.session_state.activity_index[activity] = index + 1
        st.rerun()


def home():
    st.markdown("""
    <div class="hero">
      <div class="hero-grid">
        <div>
          <div class="eyebrow">Your research, illuminated</div>
          <h1>Go deeper.<br><span>See clearer.</span></h1>
          <p>From a question to a considered report. Meet your team of AI research agents, working together to find, read, write and review.</p>
          <div class="hero-meta"><span><b>04</b> specialized agents</span><span><b>01</b> connected workflow</span></div>
        </div>
        <div class="orbit" aria-hidden="true">
          <div class="orbit-ring third"></div><div class="orbit-ring"></div><div class="orbit-ring second"></div>
          <div class="orbit-core"><span class="logo-glyph">L</span></div><div class="orbit-dot"></div><div class="orbit-dot two"></div><div class="orbit-dot three"></div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.write("")
    if st.button("Start a new research brief  →", type="primary"):
        go("New Research")
    st.markdown('<div class="section-label"><span class="eyebrow">Your research team</span></div>', unsafe_allow_html=True)
    st.markdown("### Four perspectives. One clear picture.")
    columns = st.columns(4)
    for index, (column, (name, description)) in enumerate(zip(columns, STAGES.values()), 1):
        icon = ("⌕", "◈", "✎", "✧")[index - 1]
        column.markdown(
            f'<div class="tile"><span class="eyebrow">0{index} / AGENT</span><span class="tile-icon">{icon}</span>'
            f'<p><strong>{name}</strong></p><span class="muted">{description}</span></div>',
            unsafe_allow_html=True,
        )
    st.caption("Powered by a search, reading, writing and critique workflow.")


def new_research():
    st.markdown('<div class="eyebrow">Start with a question</div>', unsafe_allow_html=True)
    st.title("New research")
    st.write("Give your research team a clear direction. The more specific your question, the more focused the brief.")
    running = st.session_state.job is not None and st.session_state.job["status"] == "running"
    if running:
        st.info("A research task is already running. Open the workspace to follow it.")
        if st.button("Open research workspace"):
            go("Research Workspace")
    with st.form("research-form"):
        topic = st.text_input("Research topic or question", placeholder="What would you like to understand?")
        context = st.text_area(
            "Additional context or instructions (optional)",
            placeholder="Add a region, audience, timeframe, or angle...",
        )
        st.caption("Research depth is determined by the existing pipeline; no separate depth control is available.")
        submitted = st.form_submit_button("Launch research  →", type="primary", disabled=running)
    if submitted:
        if not topic.strip():
            st.error("Enter a research topic before starting.")
        else:
            start_research(topic.strip(), context.strip())
            go("Research Workspace")


def render_workspace():
    previous_status = st.session_state.job["status"] if st.session_state.job else None
    receive_events()
    job = st.session_state.job
    if job is not None and previous_status == "running" and job["status"] != "running":
        st.rerun()
    if job is None:
        st.markdown('<div class="eyebrow">Your research team</div>', unsafe_allow_html=True)
        st.title("Research workspace")
        st.info("No research has been started in this session.")
        if st.button("Start research"):
            go("New Research")
        return
    st.markdown('<div class="eyebrow">Live research workspace</div>', unsafe_allow_html=True)
    st.title(job["topic"])
    st.caption(f"Started {job['started']} · Status: {job['status'].capitalize()}")
    completed = sum(value == "completed" for value in job["stages"].values())
    st.progress(completed / len(STAGES), text=f"{completed} of {len(STAGES)} stages completed")
    nodes = []
    for index, (name, (title, description)) in enumerate(STAGES.items()):
        status = job["stages"][name]
        symbol = ("⌕", "◈", "✎", "✧")[index]
        nodes.append(
            f'<div class="agent-node is-{status}"><div class="agent-symbol">{symbol}</div>'
            f'<span class="agent-name">{title}</span><div class="agent-desc">{description}</div>'
            f'<div class="status {status}">{status}</div></div>'
        )
    st.markdown('<div class="agent-flow">' + "".join(nodes) + "</div>", unsafe_allow_html=True)
    with st.expander("Execution timeline", expanded=job["status"] == "running"):
        if job["timeline"]:
            for when, name, status in reversed(job["timeline"]):
                st.write(f"{when} · {STAGES[name][0]} · {status.capitalize()}")
        else:
            st.caption("Waiting for the pipeline to start.")
    for name in ("search", "reader"):
        if name in job["outputs"]:
            with st.expander(f"{STAGES[name][0]} results"):
                st.markdown(job["outputs"][name])
    if job["status"] == "running":
        lounge()
    elif job["status"] == "failed":
        st.error(f"Research stopped: {job.get('error', 'Unknown error')}")
        if "writer" in job["outputs"]:
            with st.expander("View the report generated before the failure"):
                st.markdown(job["outputs"]["writer"])
        if st.button("Start another research"):
            go("New Research")
    elif job["status"] == "completed":
        st.success("Research complete.")
        show_report({"topic": job["topic"], "date": job["completed"], "result": job["result"]}, "current-download")


def workspace():
    job = st.session_state.job
    interval = "2s" if job is not None and job["status"] == "running" else None
    st.fragment(run_every=interval)(render_workspace)()


def history():
    st.markdown('<div class="eyebrow">Your library</div>', unsafe_allow_html=True)
    st.title("Research history")
    st.caption("Reports are stored in this browser session only and disappear when the session ends.")
    records = st.session_state.history
    if not records:
        st.info("No completed reports yet.")
        return
    query = st.text_input("Search reports", placeholder="Search by topic")
    filtered = [record for record in records if query.casefold() in record["topic"].casefold()]
    if not filtered:
        st.info("No reports match your search.")
        return
    selection = st.selectbox(
        "Open a report", [record["id"] for record in filtered],
        format_func=lambda record_id: next(
            f"{record['topic']} · {record['date']}" for record in filtered if record["id"] == record_id
        ),
    )
    selected_record = next(record for record in records if record["id"] == selection)
    show_report(selected_record, f"history-download-{selection}")
    if st.checkbox("Confirm deletion of this report", key=f"confirm-delete-{selection}"):
        if st.button("Delete report", key=f"delete-{selection}"):
            records.remove(selected_record)
            st.rerun()


def about():
    st.markdown('<div class="eyebrow">The platform</div>', unsafe_allow_html=True)
    st.title("About Lumen")
    st.write(
        "Lumen is a research workspace built on multi-agent assistants. "
        "It connects your question to a sequential research workflow and makes its real outputs easier to explore."
    )
    st.markdown("""
    **Search** -> finds sources with Tavily API.  
    **Reader** -> extracts source content using the BeautifulSoup scraping tool.  
    **Writer** -> assembles the report from gathered research.  
    **Critic** -> evaluates the report against the evidence.

    Stage indicators reflect the pipeline's actual messages, not estimated progress.
    """)


st.set_page_config(page_title="Lumen | Research workspace", page_icon="🔎", layout="wide")
initialize()
receive_events()
style()
with st.sidebar:
    st.markdown(
        '<div class="brand"><div class="brand-icon"><span class="logo-glyph">L</span></div><div>'
        '<div class="brand-name">lumen</div><div class="brand-sub">Research studio</div></div></div>',
        unsafe_allow_html=True,
    )
    st.radio("Navigate", PAGES, key="page", label_visibility="collapsed")
    st.divider()
    st.radio("Appearance", ["Dark", "Light"], key="theme", horizontal=True)
    if st.session_state.job is not None:
        st.caption(f"Current task: {st.session_state.job['status']}")

page = st.session_state.page
if page == "Home":
    home()
elif page == "New Research":
    new_research()
elif page == "Research Workspace":
    workspace()
elif page == "Research History":
    history()
else:
    about()
st.markdown('<div class="footer">Made by Mohd Umar | 2026</div>', unsafe_allow_html=True)

