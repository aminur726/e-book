"""
═══════════════════════════════════════════════════════════════
 অদম্য প্রেস — Ebook Creator v2.0
 
 v2.0 Feature List:
   ✅ Password-protected app
   ✅ Output: PDF + DOCX (both or either)
   ✅ Custom Header & Footer text
   ✅ Logo Watermark — center or side, opacity slider
   ✅ Text Watermark — center or side, opacity slider
   ✅ 50-page batch pause → review → edit/confirm → next
   ✅ Exact preview — PDF rendered as inline images with true colors
   ✅ Progress bar with page-by-page tracking
   ✅ Credit/cost estimation + live running total
═══════════════════════════════════════════════════════════════
"""
import streamlit as st
import re, os, io, time, json, base64, math
from datetime import datetime

# ─── Must be first Streamlit call ───
st.set_page_config(page_title="অদম্য প্রেস — Ebook Creator", page_icon="📚", layout="wide")

# ═══════════════════════════════════════════════════════════════
# 1. PASSWORD GATE
# ═══════════════════════════════════════════════════════════════
APP_PASSWORD = "odommo2025"  # ← Change this or load from env

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.markdown("""
    <div style="text-align:center;padding:80px 0 30px 0;">
        <div style="font-size:3rem;">📚</div>
        <h1 style="color:#C9A84C;font-size:2rem;margin:8px 0 4px 0;">অদম্য প্রেস</h1>
        <p style="color:#888;font-size:0.95rem;">Ebook Creator v2.0 — Password Protected</p>
    </div>
    """, unsafe_allow_html=True)
    _, c, _ = st.columns([1.2, 1, 1.2])
    with c:
        pwd = st.text_input("🔒 Password", type="password", placeholder="Enter app password...")
        if st.button("Login", use_container_width=True, type="primary"):
            if pwd == APP_PASSWORD:
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("❌ Wrong password. Try again.")
    st.stop()

# ═══════════════════════════════════════════════════════════════
# CONSTANTS & HELPERS
# ═══════════════════════════════════════════════════════════════
BN = ["১","২","৩","৪","৫","৬","৭","৮","৯","১০"]

def to_bn(n):
    """Integer → Bangla numeral string: 123 → ১২৩"""
    d = {"0":"০","1":"১","2":"২","3":"৩","4":"৪","5":"৫","6":"৬","7":"৭","8":"৮","9":"৯"}
    return "".join(d.get(c, c) for c in str(n))

def esc(t):
    """HTML-escape text for safe embedding"""
    if not t: return ""
    return t.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

# ═══════════════════════════════════════════════════════════════
# DARK THEME CSS
# ═══════════════════════════════════════════════════════════════
st.markdown("""
<style>
    .stApp { background:#0a0e0d; }
    section[data-testid="stSidebar"] { background:#0f1613; }
    .hero { text-align:center; padding:6px 0 2px 0; }
    .hero h1 { color:#C9A84C; font-size:1.8rem; margin:0; }
    .hero p  { color:#777; font-size:0.85rem; margin:2px 0 0 0; }
    .gold { color:#C9A84C; }
    .cost-card {
        background:#12201a; border:1px solid #253830; border-radius:8px;
        padding:10px 14px; display:inline-block; margin:4px 8px 4px 0;
    }
    .cost-card .lbl { color:#777; font-size:0.72rem; }
    .cost-card .val { color:#C9A84C; font-size:1.15rem; font-weight:700; }
    .page-thumb {
        border:1px solid #333; border-radius:3px;
        box-shadow:0 3px 12px rgba(0,0,0,0.5);
    }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# DOCX PARSER  —  Translated Bangla DOCX → structured pages
# ═══════════════════════════════════════════════════════════════
def parse_translated_docx(uploaded_file):
    """
    Parses a Bangla-translated DOCX that uses 'পৃষ্ঠা ৩' style page markers.
    Each page is expected to contain: title, optional English subtitle, quote,
    intro paragraph, 10 numbered tips (১. through ১০.), and a closing line.
    Returns a list of page dicts.
    """
    import subprocess
    tmp = "/tmp/uploaded_book.docx"
    with open(tmp, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # pandoc gives the cleanest markdown; fall back to python-docx
    md = "/tmp/_parsed.md"
    r = subprocess.run(["pandoc", tmp, "-t", "markdown", "-o", md],
                       capture_output=True, text=True)
    if r.returncode == 0 and os.path.exists(md):
        with open(md, encoding="utf-8") as f:
            raw = f.read()
    else:
        from docx import Document
        raw = "\n".join(p.text for p in Document(tmp).paragraphs)

    # Split on page markers like  পৃষ্ঠা ৩  or  **পৃষ্ঠা ৩**
    parts = re.split(r'\n?\*?\*?পৃষ্ঠা\s*([\u09E6-\u09EF]+)\*?\*?\n', raw)
    pages = []
    for i in range(1, len(parts), 2):
        bn_num = parts[i].strip()
        content = parts[i+1].strip() if i+1 < len(parts) else ""
        pg = dict(page_num=bn_num, title_bn="", title_en="",
                  quote="", quote_author="", intro="", tips=[], closing="")
        lines = [l.strip() for l in content.split('\n') if l.strip()]
        state, tip_buf, intro_l, close_l = "title", "", [], []

        for line in lines:
            c = line.replace('**','').replace('*','').strip()
            if state == "title":
                if c.startswith('(') or (re.match(r'^[A-Z]', c) and len(c) < 80):
                    pg["title_en"] = c.strip('()')
                    state = "quote"
                elif c.startswith('>') or c.startswith('"') or c.startswith('\u201c'):
                    state = "quote"; _qp(pg, c)
                else:
                    pg["title_bn"] = c
            elif state == "quote":
                if c.startswith('>') or c.startswith('"') or c.startswith('\u201c'):
                    _qp(pg, c)
                elif re.match(r'^[\u09E6-\u09EF]+[।.]', c):
                    state = "tips"
                    tip_buf = re.sub(r'^[\u09E6-\u09EF]+[।.]\s*', '', c)
                else:
                    state = "intro"; intro_l.append(c)
            elif state == "intro":
                if re.match(r'^>?\s*\*?\*?[\u09E6-\u09EF]+[।.]', c):
                    state = "tips"
                    tip_buf = re.sub(r'^>?\s*\*?\*?[\u09E6-\u09EF]+[।.]\s*\*?\*?\s*', '', c.lstrip('> '))
                else:
                    intro_l.append(c)
            elif state == "tips":
                tm = re.match(r'^>?\s*\*?\*?([\u09E6-\u09EF]+)[।.]\s*\*?\*?\s*(.*)', c)
                if tm:
                    if tip_buf: pg["tips"].append(tip_buf.strip())
                    tip_buf = tm.group(2)
                elif c.startswith('>'):
                    tip_buf += " " + c.lstrip('> ').strip()
                else:
                    if len(pg["tips"]) >= 9 and tip_buf:
                        pg["tips"].append(tip_buf.strip()); tip_buf = ""
                        state = "closing"; close_l.append(c)
                    else:
                        tip_buf += " " + c
            elif state == "closing":
                close_l.append(c)

        if tip_buf and state == "tips":
            pg["tips"].append(tip_buf.strip())
        pg["intro"]   = " ".join(intro_l).strip()
        pg["closing"] = " ".join(close_l).strip()
        if pg["title_bn"] or pg["tips"]:
            pages.append(pg)
    return pages

def _qp(pg, line):
    """Helper: parse a quote line, splitting author on — or ---"""
    qt = line.lstrip('> ').strip().strip('""\u201c\u201d\\"')
    for sep in ['---', '—', '\u2014']:
        if sep in qt:
            a, b = qt.rsplit(sep, 1)
            pg["quote"] = a.strip().strip('""\u201c\u201d')
            pg["quote_author"] = b.strip()
            return
    pg["quote"] = (pg["quote"] + " " + qt).strip() if pg["quote"] else qt


# ═══════════════════════════════════════════════════════════════
# WATERMARK + LOGO CSS/HTML BUILDER
# ═══════════════════════════════════════════════════════════════
def build_watermark(wm_text="", wm_pos="center", wm_opa=0.06,
                    logo_b64=None, logo_pos="center", logo_opa=0.06):
    """
    Returns (extra_css, html_snippet) to overlay on each chapter page.
    The html_snippet goes inside each .pg div, sitting behind z-index:1 content.
    """
    css, html = "", ""

    # ── Text watermark ──
    if wm_text:
        if wm_pos == "center":
            css += f"""
.wm-txt {{ position:absolute; top:50%; left:50%;
  transform:translate(-50%,-50%) rotate(-35deg);
  font-size:28pt; font-weight:900; white-space:nowrap;
  color:rgba(0,0,0,{wm_opa}); pointer-events:none; z-index:0;
  font-family:'Noto Sans Bengali',sans-serif; }}
"""
        else:  # side — vertical text on right edge
            css += f"""
.wm-txt {{ position:absolute; top:50%; right:-6mm;
  transform:translateY(-50%) rotate(90deg);
  font-size:10pt; font-weight:600; white-space:nowrap; letter-spacing:2px;
  color:rgba(0,0,0,{wm_opa}); pointer-events:none; z-index:0;
  font-family:'Noto Sans Bengali',sans-serif; }}
"""
        html += f'<div class="wm-txt">{esc(wm_text)}</div>'

    # ── Logo watermark ──
    if logo_b64:
        if logo_pos == "center":
            css += f"""
.wm-logo {{ position:absolute; top:50%; left:50%;
  transform:translate(-50%,-50%); opacity:{logo_opa};
  pointer-events:none; z-index:0; }}
.wm-logo img {{ width:40mm; height:auto; }}
"""
        else:
            css += f"""
.wm-logo {{ position:absolute; bottom:14mm; right:4mm;
  opacity:{logo_opa}; pointer-events:none; z-index:0; }}
.wm-logo img {{ width:16mm; height:auto; }}
"""
        html += f'<div class="wm-logo"><img src="data:image/png;base64,{logo_b64}"></div>'

    return css, html


# ═══════════════════════════════════════════════════════════════
# DESIGN CSS  —  V2 Navy-Gold  |  V4 B&W Clean
# ═══════════════════════════════════════════════════════════════

# Header / footer are injected through the in-page HTML elements
# (WeasyPrint's @page margin boxes have limited Bengali support)

V2_CSS = """
@page { size:148mm 210mm; margin:10mm 12mm 12mm 12mm; background:#fff; }
@page cover { margin:0; background:#0B1D3A; }
@page endpage { margin:0; background:#0B1D3A; }
body { font-family:'Noto Sans Bengali',sans-serif; color:#111; margin:0; padding:0; }
.cover{page:cover;page-break-after:always;width:148mm;height:210mm;text-align:center;color:#fff;position:relative}
.cover-border{position:absolute;top:6mm;left:6mm;right:6mm;bottom:6mm;border:1.5px solid #C9A84C}
.cover-inner{position:absolute;top:8mm;left:8mm;right:8mm;bottom:8mm;border:.5px solid rgba(201,168,76,.5)}
.cover .c-tbn{font-size:22pt;font-weight:900;color:#C9A84C;margin-top:55mm}
.cover .c-ten{font-family:'Noto Serif Bengali',serif;font-size:10pt;font-style:italic;color:#DFC06E;margin-top:3mm}
.cover .c-rl{width:40mm;height:.5px;background:#C9A84C;margin:6mm auto}
.cover .c-tag{font-size:8.5pt;font-weight:300;color:rgba(255,255,255,.85);margin-top:4mm}
.cover .c-tage{font-size:7pt;color:rgba(255,255,255,.55);margin-top:2mm}
.cover .c-sub{font-size:7pt;color:rgba(255,255,255,.5);margin-top:8mm}
.cover .c-bot{position:absolute;bottom:18mm;text-align:center;width:100%}
.cover .c-auth{font-size:7.5pt;font-weight:600;color:#fff;margin-bottom:6mm}
.cover .c-pbn{font-size:9pt;font-weight:600;color:#C9A84C}
.cover .c-pen{font-size:5.5pt;color:rgba(255,255,255,.45);margin-top:1mm}

.pg{page-break-after:always;page-break-inside:avoid;box-sizing:border-box;display:flex;flex-direction:column;height:186mm;position:relative;overflow:hidden}
.pg .hdr{display:flex;justify-content:space-between;font-size:6pt;color:#aaa;margin-bottom:0;width:calc(100% + 24mm);margin-left:-12mm;padding:0 12mm}
.pg .topbar{width:calc(100% + 24mm);height:3mm;background:#0B1D3A;margin:-10mm -12mm 0 -12mm;border-bottom:.5px solid #C9A84C}
.pg .pnum{text-align:right;font-size:7pt;color:#999;margin:3mm 0;position:relative;z-index:1}
.pg .t-bn{font-size:13pt;font-weight:700;color:#0B1D3A;text-align:center;line-height:1.4;margin-bottom:1.5mm;position:relative;z-index:1}
.pg .t-en{font-family:'Noto Serif Bengali',serif;font-size:7.5pt;color:#888;text-align:center;font-style:italic;margin-bottom:2.5mm;position:relative;z-index:1}
.pg .g-rl{border:none;border-top:.6px solid #C9A84C;margin:0 15mm 1mm 15mm}
.pg .g-dia{text-align:center;color:#C9A84C;font-size:5pt;margin-bottom:2.5mm;line-height:1}
.pg .qbox{background:rgba(253,248,232,.5);border-left:2px solid #C9A84C;border-radius:2px;padding:3mm 4mm 2.5mm 5mm;margin:1mm 3mm 2.5mm 3mm;position:relative;z-index:1}
.pg .qtxt{font-family:'Noto Serif Bengali',serif;font-size:8pt;font-style:italic;color:#0B1D3A;text-align:center;line-height:1.55}
.pg .qau{font-size:6.5pt;color:#888;text-align:center;font-style:italic;margin-top:1mm}
.pg .intro{font-size:8pt;color:#1a1a2e;line-height:1.6;text-align:justify;margin:2mm 0;position:relative;z-index:1}
.pg .sep{border:none;border-top:.3px solid #DFC06E;margin:1.5mm 20mm}
.pg .tips{flex:1;display:flex;flex-direction:column;justify-content:space-between;margin:1mm 0;position:relative;z-index:1}
.pg .tip{font-size:7.8pt;color:#3D3D5C;line-height:1.55;padding-left:5.5mm;text-indent:-5.5mm;text-align:justify}
.pg .tn{font-weight:700;color:#0B1D3A;font-size:8pt}
.pg .clsw{margin-top:auto;position:relative;z-index:1}
.pg .cls{padding-top:2mm;border-top:.3px solid #DFC06E;font-family:'Noto Serif Bengali',serif;font-size:7.5pt;color:#6B6B8A;text-align:center;line-height:1.55}
.pg .foot{text-align:center;font-size:5pt;color:#bbb;margin-top:1.5mm;position:relative;z-index:1}

.endpg{page:endpage;width:148mm;height:210mm;text-align:center;color:#fff;position:relative}
.endpg .border{position:absolute;top:8mm;left:8mm;right:8mm;bottom:8mm;border:1px solid #C9A84C}
.endpg .cnt{padding-top:70mm}
.endpg .et{font-size:16pt;font-weight:900;color:#C9A84C;margin-bottom:8mm}
.endpg .ep{font-size:10pt;font-weight:600;color:#DFC06E}
.endpg .epe{font-size:6pt;color:rgba(255,255,255,.4);margin-top:2mm}
"""

V4_CSS = """
@page{size:148mm 210mm;margin:12mm 13mm 11mm 13mm;background:#fff}
@page cover{margin:0;background:#111}
@page endpage{margin:0;background:#111}
body{font-family:'Noto Sans Bengali',sans-serif;color:#111;margin:0;padding:0}
.cover{page:cover;page-break-after:always;width:148mm;height:210mm;text-align:center;color:#fff;position:relative}
.cover-border{position:absolute;top:6mm;left:6mm;right:6mm;bottom:6mm;border:1.5px solid #555}
.cover-inner{position:absolute;top:8mm;left:8mm;right:8mm;bottom:8mm;border:.5px solid #444}
.cover .c-tbn{font-size:22pt;font-weight:900;color:#fff;margin-top:55mm}
.cover .c-ten{font-family:'Noto Serif Bengali',serif;font-size:10pt;font-style:italic;color:#aaa;margin-top:3mm}
.cover .c-rl{width:40mm;height:.5px;background:#666;margin:6mm auto}
.cover .c-tag{font-size:8.5pt;font-weight:300;color:rgba(255,255,255,.8);margin-top:4mm}
.cover .c-tage{font-size:7pt;color:rgba(255,255,255,.5);margin-top:2mm}
.cover .c-sub{font-size:7pt;color:rgba(255,255,255,.4);margin-top:8mm}
.cover .c-bot{position:absolute;bottom:18mm;text-align:center;width:100%}
.cover .c-auth{font-size:7.5pt;font-weight:600;color:#fff;margin-bottom:6mm}
.cover .c-pbn{font-size:9pt;font-weight:600;color:#ccc}
.cover .c-pen{font-size:5.5pt;color:rgba(255,255,255,.4);margin-top:1mm}

.pg{page-break-after:always;page-break-inside:avoid;box-sizing:border-box;display:flex;flex-direction:column;height:187mm;position:relative;overflow:hidden}
.pg .hdr{display:flex;justify-content:space-between;font-size:6pt;color:#999;margin-bottom:2mm}
.pg .pnum{text-align:right;font-size:7.5pt;color:#333;margin-bottom:4mm;font-weight:500;position:relative;z-index:1}
.pg .t-bn{font-size:14pt;font-weight:700;color:#000;text-align:center;line-height:1.4;margin-bottom:1.5mm;position:relative;z-index:1}
.pg .t-en{font-family:'Noto Serif Bengali',serif;font-size:8.5pt;color:#555;text-align:center;font-style:italic;margin-bottom:3mm;position:relative;z-index:1}
.pg .sep-b{border:none;border-top:1.2px solid #000;margin:0 15mm 3mm 15mm}
.pg .sep-l{border:none;border-top:.4px solid #888;margin:2mm 18mm}
.pg .qblk{margin:2mm 5mm 3mm 5mm;text-align:center;position:relative;z-index:1}
.pg .qtxt{font-family:'Noto Serif Bengali',serif;font-size:8.5pt;color:#111;font-style:italic;line-height:1.6}
.pg .qau{font-size:7.5pt;color:#666;margin-top:1mm}
.pg .intro{font-size:8.5pt;color:#111;line-height:1.65;text-align:justify;margin:2.5mm 0;position:relative;z-index:1}
.pg .tips{flex:1;display:flex;flex-direction:column;justify-content:space-between;margin:1.5mm 0;position:relative;z-index:1}
.pg .tip{font-size:8pt;color:#1a1a1a;line-height:1.6;padding-left:6mm;text-indent:-6mm;text-align:justify}
.pg .tn{font-weight:700;color:#000;font-size:8.5pt}
.pg .clsw{margin-top:auto;position:relative;z-index:1}
.pg .cls{padding-top:2mm;border-top:.4px solid #999;font-family:'Noto Serif Bengali',serif;font-size:8pt;color:#444;text-align:center;line-height:1.6}
.pg .foot{text-align:center;font-size:5.5pt;color:#bbb;margin-top:2mm;position:relative;z-index:1}

.endpg{page:endpage;width:148mm;height:210mm;text-align:center;color:#fff;position:relative}
.endpg .border{position:absolute;top:8mm;left:8mm;right:8mm;bottom:8mm;border:1px solid #555}
.endpg .cnt{padding-top:70mm}
.endpg .et{font-size:16pt;font-weight:900;color:#fff;margin-bottom:8mm}
.endpg .ep{font-size:10pt;font-weight:600;color:#ccc}
.endpg .epe{font-size:6pt;color:rgba(255,255,255,.4);margin-top:2mm}
"""


# ═══════════════════════════════════════════════════════════════
# HTML BUILDERS — Cover, End, Chapter (V2 + V4)
# ═══════════════════════════════════════════════════════════════
def html_cover(info):
    return f"""<div class="cover">
  <div class="cover-border"></div><div class="cover-inner"></div>
  <div class="c-tbn">{esc(info['title_bn'])}</div>
  <div class="c-ten">{esc(info['title_en'])}</div>
  <div class="c-rl"></div>
  <div class="c-tag">{esc(info.get('tagline_bn',''))}</div>
  <div class="c-tage">{esc(info.get('tagline_en',''))}</div>
  <div class="c-sub">{esc(info.get('subtitle_bn',''))}</div>
  <div class="c-bot">
    <div class="c-auth">{esc(info.get('author',''))}</div>
    <div class="c-pbn">অদম্য প্রেস</div>
    <div class="c-pen">Odommo Press</div>
  </div>
</div>"""

def html_end():
    return """<div class="endpg">
  <div class="border"></div>
  <div class="cnt">
    <div class="et">— সমাপ্ত —</div>
    <div class="ep">অদম্য প্রেস</div>
    <div class="epe">Odommo Press</div>
  </div>
</div>"""

def _tips_html(tips):
    return "".join(
        f'<div class="tip"><span class="tn">{BN[i]}.</span> {esc(t)}</div>\n'
        for i, t in enumerate(tips[:10]) if i < 10
    )

def html_v2_ch(p, foot_txt, hdr_l, hdr_r, wm):
    """Single chapter page — V2 Navy Gold design"""
    au = f'<div class="qau">— {esc(p["quote_author"])}</div>' if p.get("quote_author") else ""
    cl = f'<div class="clsw"><div class="cls">{esc(p["closing"])}</div></div>' if p.get("closing") else ""
    hdr = f'<div class="hdr"><span>{esc(hdr_l)}</span><span>{esc(hdr_r)}</span></div>' if (hdr_l or hdr_r) else ""
    return f"""<div class="pg">
  {wm}
  <div class="topbar"></div>
  {hdr}
  <div class="pnum">পৃষ্ঠা {esc(p['page_num'])}</div>
  <div class="t-bn">{esc(p['title_bn'])}</div>
  <div class="t-en">{esc(p['title_en'])}</div>
  <hr class="g-rl"><div class="g-dia">◆</div>
  <div class="qbox"><div class="qtxt">"{esc(p['quote'])}"</div>{au}</div>
  <div class="intro">{esc(p['intro'])}</div><hr class="sep">
  <div class="tips">{_tips_html(p['tips'])}</div>
  {cl}
  <div class="foot">{esc(foot_txt)}</div>
</div>"""

def html_v4_ch(p, foot_txt, hdr_l, hdr_r, wm):
    """Single chapter page — V4 B&W Clean design"""
    au = f'<div class="qau">— {esc(p["quote_author"])}</div>' if p.get("quote_author") else ""
    cl = f'<div class="clsw"><div class="cls">{esc(p["closing"])}</div></div>' if p.get("closing") else ""
    hdr = f'<div class="hdr"><span>{esc(hdr_l)}</span><span>{esc(hdr_r)}</span></div>' if (hdr_l or hdr_r) else ""
    return f"""<div class="pg">
  {wm}
  {hdr}
  <div class="pnum">পৃষ্ঠা {esc(p['page_num'])}</div>
  <div class="t-bn">{esc(p['title_bn'])}</div>
  <div class="t-en">{esc(p['title_en'])}</div>
  <hr class="sep-b">
  <div class="qblk"><div class="qtxt">"{esc(p['quote'])}"</div>{au}</div>
  <hr class="sep-l">
  <div class="intro">{esc(p['intro'])}</div>
  <div class="tips">{_tips_html(p['tips'])}</div>
  {cl}
  <div class="foot">{esc(foot_txt)}</div>
</div>"""


# ═══════════════════════════════════════════════════════════════
# PDF GENERATION
# ═══════════════════════════════════════════════════════════════
def make_pdf(pages, design, info, wm_css, wm_html,
             hdr_l, hdr_r, foot_txt, progress_cb=None):
    """Build a complete ebook PDF. Returns bytes."""
    import weasyprint
    css = (V2_CSS if design == "v2" else V4_CSS) + "\n" + wm_css
    ch_fn = html_v2_ch if design == "v2" else html_v4_ch
    body = html_cover(info)
    for i, pg in enumerate(pages):
        body += ch_fn(pg, foot_txt, hdr_l, hdr_r, wm_html)
        if progress_cb:
            progress_cb(i + 1, len(pages))
    body += html_end()
    html = f'<!DOCTYPE html><html lang="bn"><head><meta charset="UTF-8"><style>{css}</style></head><body>{body}</body></html>'
    return weasyprint.HTML(string=html).write_pdf()


def make_preview_pdf(pages, design, info, wm_css, wm_html,
                     hdr_l, hdr_r, foot_txt):
    """Preview: cover + first page + last page."""
    import weasyprint
    css = (V2_CSS if design == "v2" else V4_CSS) + "\n" + wm_css
    ch_fn = html_v2_ch if design == "v2" else html_v4_ch
    picks = []
    if pages: picks.append(pages[0])
    if len(pages) > 1: picks.append(pages[-1])
    body = html_cover(info)
    for pg in picks:
        body += ch_fn(pg, foot_txt, hdr_l, hdr_r, wm_html)
    html = f'<!DOCTYPE html><html lang="bn"><head><meta charset="UTF-8"><style>{css}</style></head><body>{body}</body></html>'
    return weasyprint.HTML(string=html).write_pdf()


def pdf_to_images(pdf_bytes, dpi=160):
    """Convert PDF bytes → list of PNG bytes (one per page) for inline preview."""
    import fitz
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    imgs = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
        imgs.append(pix.tobytes("png"))
    doc.close()
    return imgs


# ═══════════════════════════════════════════════════════════════
# DOCX GENERATION  —  Professional output via docx-js (Node.js)
#
# Uses generate_docx.js for full design-matched Word output:
#   • Page borders on cover & end pages
#   • Shaded quote blocks with gold/gray left border
#   • Decorative horizontal rules as thin bordered tables
#   • Color-matched text (navy/gold for V2, B&W for V4)
#   • Tab-stop aligned header (left + right)
#   • Centered footer
#   • Proper hanging-indent numbered tips
#   • Diamond separator (V2 only)
#   • 3-section document: cover → chapters → end page
# ═══════════════════════════════════════════════════════════════

# Resolve the path to generate_docx.js at module load time
# so it works whether run via `streamlit run` or directly.
_DOCX_JS_SCRIPT = None
for _candidate in [
    os.path.join(os.path.dirname(os.path.abspath(__file__ if '__file__' in dir() else '')), "generate_docx.js"),
    os.path.join(os.getcwd(), "generate_docx.js"),
    os.path.join(os.path.dirname(os.path.abspath(os.sys.argv[0])) if os.sys.argv else ".", "generate_docx.js"),
    "/home/claude/ebook-creator/generate_docx.js",
]:
    if os.path.exists(_candidate):
        _DOCX_JS_SCRIPT = _candidate
        break


def make_docx(pages, design, info, hdr_l, hdr_r, foot_txt):
    """Generate a professionally styled Word document.

    Calls generate_docx.js (docx-js / Node.js) which produces a
    design-matched DOCX with borders, shading, colored text,
    tab-stop headers, and decorative elements.

    Returns: DOCX file bytes
    """
    import subprocess, tempfile

    js_script = _DOCX_JS_SCRIPT
    if not js_script or not os.path.exists(js_script):
        raise FileNotFoundError(
            "generate_docx.js not found. Place it next to app.py."
        )

    input_data = {
        "info": info,
        "pages": pages,
        "header_left": hdr_l or "",
        "header_right": hdr_r or "",
        "footer_text": foot_txt or "",
    }

    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".json", delete=False
    ) as f:
        json.dump(input_data, f, ensure_ascii=False)
        input_path = f.name

    output_path = input_path.replace(".json", ".docx")

    try:
        result = subprocess.run(
            ["node", js_script, input_path, output_path, design],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"DOCX generation failed:\n{result.stderr}"
            )
        with open(output_path, "rb") as f:
            return f.read()
    finally:
        for p in [input_path, output_path]:
            try:
                os.unlink(p)
            except OSError:
                pass


# ═══════════════════════════════════════════════════════════════
# COST ESTIMATOR
# ═══════════════════════════════════════════════════════════════
PRICES = {
    # model → (input_per_M, output_per_M) in USD
    "claude-sonnet-4-5-20250929": (3.0, 15.0),
    "claude-haiku-4-5-20251001":  (0.8, 4.0),
    "gpt-4o":      (2.5, 10.0),
    "gpt-4o-mini": (0.15, 0.6),
    "gemini-2.0-flash": (0.1, 0.4),
    "gemini-1.5-pro":   (1.25, 5.0),
}

def estimate_cost(n_pages, model_name):
    """Pre-translation cost estimation. Returns dict with usd, bdt, batches."""
    avg_in, avg_out = 800, 1200  # tokens per page
    n_batches = math.ceil(n_pages / 5)
    total_in = n_pages * avg_in + n_batches * 500  # +system prompt
    total_out = n_pages * avg_out
    ip, op = PRICES.get(model_name, (3.0, 15.0))
    usd = (total_in * ip + total_out * op) / 1e6
    return dict(usd=usd, bdt=usd * 120, batches=n_batches,
                tok_in=total_in, tok_out=total_out)


# ═══════════════════════════════════════════════════════════════
# TRANSLATION PROMPT + PARSER
# ═══════════════════════════════════════════════════════════════
TRANS_PROMPT = """You are a professional English-to-Bangla book translator for অদম্য প্রেস.

BALANCED STYLE (70% Bangla / 30% English):
- English ONLY for: hard Bangla words OR daily-use English (Goal, Habit, Energy, Focus,
  Routine, Environment, Support, Mindset, Confidence, Discipline, Balance, Pattern,
  Trigger, Comfort Zone, Stress, Emotional Flexibility, Limiting Beliefs, Version, Purpose)
- Bangla for: পরিচয়, মূল্যবোধ, স্বপ্ন, সাহস, সচেতনতা, দোষবোধ, পদক্ষেপ, গল্প, সম্পর্ক, ভয়

FORMAT per page:
PAGE [num]
TITLE_BN: ... | TITLE_EN: ... | QUOTE: ... | QUOTE_AUTHOR: ...
INTRO: ... | TIP_1: ... through TIP_10: ... | CLOSING: ...
---
Use আপনি (formal). Bangla numerals."""

def parse_trans(raw):
    pages = []
    for blk in re.split(r'---+', raw):
        blk = blk.strip()
        if not blk: continue
        pg = dict(page_num="",title_bn="",title_en="",quote="",
                  quote_author="",intro="",tips=[],closing="")
        for ln in blk.split('\n'):
            ln = ln.strip()
            if ln.startswith('PAGE '): pg["page_num"] = to_bn(re.sub(r'\D','',ln))
            elif ln.startswith('TITLE_BN:'): pg["title_bn"] = ln[9:].strip()
            elif ln.startswith('TITLE_EN:'): pg["title_en"] = ln[9:].strip()
            elif ln.startswith('QUOTE:'): pg["quote"] = ln[6:].strip()
            elif ln.startswith('QUOTE_AUTHOR:'): pg["quote_author"] = ln[13:].strip()
            elif ln.startswith('INTRO:'): pg["intro"] = ln[6:].strip()
            elif re.match(r'TIP_\d+:', ln): pg["tips"].append(re.sub(r'TIP_\d+:\s*','',ln))
            elif ln.startswith('CLOSING:'): pg["closing"] = ln[8:].strip()
        if pg["title_bn"] or pg["tips"]: pages.append(pg)
    return pages


# ═══════════════════════════════════════════════════════════════
# ███          M A I N   A P P   U I                         ███
# ═══════════════════════════════════════════════════════════════

st.markdown('<div class="hero"><h1>📚 অদম্য প্রেস — Ebook Creator</h1>'
            '<p>v2.0 • PDF + DOCX • Watermark • Batch Review • Cost Tracker</p></div>',
            unsafe_allow_html=True)
st.divider()

mode = st.radio("**Mode**",
    ["📖 Formatting (FREE)", "🌐 Translation + Formatting (API Cost)"],
    horizontal=True)

# ═══════════════════════════════════════════════════════════════
# SIDEBAR — All settings in one place
# ═══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📕 Book Info")
    S = dict(
        title_bn = st.text_input("শিরোনাম (BN)", "আত্মনিয়ন্ত্রণের বই"),
        title_en = st.text_input("Title (EN)", "The Book of Self-Control"),
        author   = st.text_input("Author", "Harvard-Fiction KH"),
        tagline_bn = st.text_input("Tagline BN", "৩৬৫ দিনের অনুপ্রেরণা"),
        tagline_en = st.text_input("Tagline EN", "365 Days of Inspiration"),
        subtitle_bn = st.text_input("Subtitle", "ব্যক্তিগত উন্নয়ন ও পেশাদার সাফল্যের নীতিমালা"),
    )
    st.divider()

    st.markdown("### 🎨 Design + Output")
    design_opt = st.selectbox("Design", ["V2 — Navy Gold Elegant", "V4 — B&W Clean Book"])
    DK = "v2" if "V2" in design_opt else "v4"
    want_pdf  = st.checkbox("Output PDF",  True)
    want_docx = st.checkbox("Output DOCX", True)
    st.divider()

    st.markdown("### 📝 Header & Footer")
    HDR_L = st.text_input("Header Left", "আত্মনিয়ন্ত্রণের বই")
    HDR_R = st.text_input("Header Right", "অদম্য প্রেস")
    FOOT  = st.text_input("Footer", "আত্মনিয়ন্ত্রণের বই  •  অদম্য প্রেস")
    st.divider()

    st.markdown("### 💧 Text Watermark")
    wm_on = st.checkbox("Enable", False, key="wm_on")
    WM_T, WM_P, WM_O = "", "center", 0.06
    if wm_on:
        WM_T = st.text_input("Text", "অদম্য প্রেস", key="wmt")
        WM_P = st.selectbox("Position", ["center","side"], key="wmp")
        WM_O = st.slider("Opacity", 0.02, 0.25, 0.06, 0.01, key="wmo")

    st.markdown("### 🖼️ Logo Watermark")
    lg_on = st.checkbox("Enable", False, key="lg_on")
    LG_B64, LG_P, LG_O = None, "center", 0.06
    if lg_on:
        lg_file = st.file_uploader("Logo (PNG/JPG)", ["png","jpg","jpeg"], key="lgf")
        if lg_file:
            LG_B64 = base64.b64encode(lg_file.read()).decode()
        LG_P = st.selectbox("Position", ["center","side"], key="lgp")
        LG_O = st.slider("Opacity", 0.02, 0.25, 0.06, 0.01, key="lgo")

    # Translation-mode extras
    if "Translation" in mode:
        st.divider()
        st.markdown("### 🔑 API")
        api_prov = st.selectbox("Provider", ["Anthropic","OpenAI","Gemini"])
        api_key  = st.text_input("API Key", type="password")
        if api_prov == "Anthropic":
            MDL = st.selectbox("Model", ["claude-sonnet-4-5-20250929","claude-haiku-4-5-20251001"])
        elif api_prov == "OpenAI":
            MDL = st.selectbox("Model", ["gpt-4o","gpt-4o-mini"])
        else:
            MDL = st.selectbox("Model", ["gemini-2.0-flash","gemini-1.5-pro"])
        pg_s = st.number_input("Start Page", 1, 999, 1)
        pg_e = st.number_input("End Page",   1, 999, 50)

    st.divider()
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.auth = False
        st.rerun()

# ── Derived values ──
wm_css, wm_html = build_watermark(WM_T, WM_P, WM_O, LG_B64, LG_P, LG_O)


# ═══════════════════════════════════════════════════════════════
# Helper: show inline preview images from PDF bytes
# ═══════════════════════════════════════════════════════════════
def show_preview(pdf_bytes, labels=None):
    """Render each page of a PDF as an inline image with shadow."""
    imgs = pdf_to_images(pdf_bytes, dpi=170)
    for idx, img in enumerate(imgs):
        lbl = labels[idx] if labels and idx < len(labels) else f"Page {idx+1}"
        st.caption(lbl)
        b64 = base64.b64encode(img).decode()
        st.markdown(
            f'<img src="data:image/png;base64,{b64}" class="page-thumb" '
            f'style="width:100%;">',
            unsafe_allow_html=True
        )


# ═══════════════════════════════════════════════════════════════
# Helper: generate + offer both outputs (PDF + DOCX)
# ═══════════════════════════════════════════════════════════════
def generate_and_download(pages, label_prefix=""):
    """Full generation with progress bar → download buttons for PDF and/or DOCX."""
    progress = st.progress(0)
    status = st.empty()
    t0 = time.time()

    if want_pdf:
        def upd(c, t):
            progress.progress(c / t)
            status.text(f"Formatting PDF page {c}/{t}...")
        pdf = make_pdf(pages, DK, S, wm_css, wm_html, HDR_L, HDR_R, FOOT, upd)
        dt = time.time() - t0
        st.success(f"📚 PDF ready — **{len(pages)} pages** | {len(pdf)/1024:.0f} KB | {dt:.1f}s")
        fn = f"{S['title_en'].replace(' ','_')}_{DK.upper()}.pdf"
        st.download_button(f"⬇️ {label_prefix}Download PDF", pdf, fn,
                          "application/pdf", type="primary", use_container_width=True)

    if want_docx:
        status.text("Building DOCX...")
        dxb = make_docx(pages, DK, S, HDR_L, HDR_R, FOOT)
        st.success(f"📄 DOCX ready — **{len(pages)} pages** | {len(dxb)/1024:.0f} KB")
        fn = f"{S['title_en'].replace(' ','_')}.docx"
        st.download_button(f"⬇️ {label_prefix}Download DOCX", dxb, fn,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)

    progress.progress(1.0)
    status.text("✅ All outputs ready!")


# ═══════════════════════════════════════════════════════════════
# FORMATTING MODE  (FREE)
# ═══════════════════════════════════════════════════════════════
if "Formatting" in mode:
    st.markdown("### 📤 Upload Translated Bangla DOCX")
    st.info("Upload → exact preview with colours/watermark → review every 50 pages → download PDF + DOCX. **Free.**")

    uploaded = st.file_uploader("Upload DOCX", type=["docx"])
    if not uploaded:
        st.stop()

    with st.spinner("📄 Parsing DOCX..."):
        DATA = parse_translated_docx(uploaded)
    if not DATA:
        st.error("❌ No pages found. Ensure 'পৃষ্ঠা ৩' markers exist.")
        st.stop()

    st.success(f"✅ **{len(DATA)} pages** parsed")
    out_label = "PDF + DOCX" if (want_pdf and want_docx) else ("PDF" if want_pdf else "DOCX")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📄 Pages", len(DATA))
    c2.metric("🎨 Design", "V2 Navy" if DK=="v2" else "V4 B&W")
    c3.metric("📤 Output", out_label)
    c4.metric("💰 Cost", "FREE ✨")
    st.divider()

    # ─── EXACT PREVIEW ───────────────────────────────────────
    st.markdown("### 👁️ Exact Preview")
    st.caption("Shows real PDF rendering — exact colours, watermark, header/footer.")

    col1, col2 = st.columns(2)
    for col, dk, nm in [(col1,"v2","🔷 V2 Navy Gold"),(col2,"v4","⬛ V4 B&W Clean")]:
        with col:
            st.markdown(f"**{nm}**")
            if st.button(f"Render {nm}", key=f"pr_{dk}", use_container_width=True):
                with st.spinner("Rendering..."):
                    pv = make_preview_pdf(DATA, dk, S, wm_css, wm_html, HDR_L, HDR_R, FOOT)
                show_preview(pv, ["Cover", "First Page", "Last Page"])
                st.download_button(f"⬇️ Preview PDF", pv, f"preview_{dk}.pdf",
                    "application/pdf", use_container_width=True, key=f"dl_{dk}")
    st.divider()

    # ─── 50-PAGE BATCH REVIEW + GENERATE ─────────────────────
    st.markdown("### 🚀 Generate Full Ebook")
    BATCH = 50
    n_batches = math.ceil(len(DATA) / BATCH)

    # Session state for batch flow
    if "fmt_batch" not in st.session_state:
        st.session_state.fmt_batch = dict(idx=0, confirmed=[], started=False)
    fb = st.session_state.fmt_batch

    if not fb["started"]:
        st.markdown(f"**{len(DATA)} pages** will be reviewed in batches of {BATCH}.")
        if st.button(f"📚 Start Generation ({len(DATA)} pages)", type="primary",
                     use_container_width=True):
            st.session_state.fmt_batch = dict(idx=0, confirmed=[], started=True)
            st.rerun()
        st.stop()

    # Already started — batch loop
    ci = fb["idx"]

    if ci >= n_batches:
        # ── ALL CONFIRMED → final output ──
        st.markdown("### ✅ All batches confirmed!")
        st.progress(1.0)
        generate_and_download(fb["confirmed"])

        if st.button("🔄 New Job", use_container_width=True):
            st.session_state.fmt_batch = dict(idx=0, confirmed=[], started=False)
            st.rerun()
    else:
        # ── SHOW CURRENT BATCH FOR REVIEW ──
        bs = ci * BATCH
        be = min(bs + BATCH, len(DATA))
        batch = DATA[bs:be]

        st.markdown(f"### 📋 Batch {ci+1}/{n_batches} — Pages {bs+1}–{be}")
        st.progress(ci / n_batches)
        st.caption(f"{len(fb['confirmed'])} pages confirmed so far")

        # Exact preview of batch (first + last page)
        with st.spinner("Rendering batch preview..."):
            bpv = make_preview_pdf(batch, DK, S, wm_css, wm_html, HDR_L, HDR_R, FOOT)
        pvlabels = ["First page of batch"]
        if len(batch) > 1: pvlabels.append("Last page of batch")
        pc1, pc2 = st.columns(2)
        imgs = pdf_to_images(bpv, dpi=150)
        for idx2, img in enumerate(imgs):
            target = pc1 if idx2 == 0 else pc2
            with target:
                st.caption(pvlabels[min(idx2, len(pvlabels)-1)])
                b64 = base64.b64encode(img).decode()
                st.markdown(f'<img src="data:image/png;base64,{b64}" class="page-thumb" '
                           f'style="width:100%;">', unsafe_allow_html=True)

        with st.expander(f"📖 View all {len(batch)} titles"):
            for p in batch:
                st.text(f"পৃষ্ঠা {p['page_num']}: {p['title_bn']}")

        st.markdown("---")
        a1, a2, a3 = st.columns(3)
        with a1:
            if st.button("✅ Confirm & Next", type="primary", use_container_width=True, key="fc"):
                fb["confirmed"].extend(batch); fb["idx"] += 1; st.rerun()
        with a2:
            if st.button("⏭️ Skip Batch", use_container_width=True, key="fs"):
                fb["idx"] += 1; st.rerun()
        with a3:
            if st.button("🛑 Stop & Generate Now", use_container_width=True, key="fstop"):
                fb["confirmed"].extend(batch)
                fb["idx"] = n_batches; st.rerun()


# ═══════════════════════════════════════════════════════════════
# TRANSLATION MODE  (API COST)
# ═══════════════════════════════════════════════════════════════
else:
    st.markdown("### 📤 Upload English PDF")
    st.info("Upload → estimate cost → preview 2 pages → translate in 50-page batches with pause → PDF + DOCX.")
    uploaded = st.file_uploader("Upload PDF", ["pdf"])

    if not uploaded:
        st.info("📤 Upload an English PDF to begin.")
        st.stop()

    if not api_key:
        st.warning("⚠️ Enter API key in sidebar.")
        st.stop()

    import fitz
    tmp = "/tmp/_eng.pdf"
    with open(tmp,"wb") as f: f.write(uploaded.getbuffer())
    pdfdoc = fitz.open(tmp)
    total_pg = pdfdoc.page_count

    st.success(f"✅ PDF loaded: **{total_pg} pages**")
    n_translate = min(pg_e, total_pg) - pg_s + 1
    est = estimate_cost(n_translate, MDL)

    # ─── COST ESTIMATION DASHBOARD ────────────────────────────
    st.markdown("#### 💰 Cost Estimation")
    cc = st.columns(5)
    cc[0].metric("📄 Pages", n_translate)
    cc[1].metric("📦 API Batches", est['batches'])
    cc[2].metric("💵 Est. USD", f"${est['usd']:.3f}")
    cc[3].metric("🇧🇩 Est. BDT", f"৳{est['bdt']:.0f}")
    cc[4].metric("🎨 Design", "V2" if DK=="v2" else "V4")
    st.divider()

    # ─── PREVIEW (2 pages) ────────────────────────────────────
    st.markdown("### 👁️ Translation Preview (2 pages)")

    if st.button("🔍 Preview Translation", use_container_width=True):
        sample = ""
        for i in range(pg_s - 1, min(pg_s + 1, total_pg)):
            sample += f"\n=== PAGE {i+1} ===\n{pdfdoc[i].get_text()}\n"
        with st.spinner("Translating 2 sample pages..."):
            try:
                if api_prov == "Anthropic":
                    import anthropic
                    cl = anthropic.Anthropic(api_key=api_key)
                    rsp = cl.messages.create(model=MDL, max_tokens=8000,
                        system=TRANS_PROMPT,
                        messages=[{"role":"user","content":f"Translate:\n\n{sample}"}])
                    raw = rsp.content[0].text
                    pcost = (rsp.usage.input_tokens * PRICES[MDL][0] +
                             rsp.usage.output_tokens * PRICES[MDL][1]) / 1e6
                elif api_prov == "OpenAI":
                    from openai import OpenAI
                    cl = OpenAI(api_key=api_key)
                    rsp = cl.chat.completions.create(model=MDL, max_tokens=8000,
                        messages=[{"role":"system","content":TRANS_PROMPT},
                                  {"role":"user","content":f"Translate:\n\n{sample}"}])
                    raw = rsp.choices[0].message.content
                    pcost = (rsp.usage.prompt_tokens * PRICES[MDL][0] +
                             rsp.usage.completion_tokens * PRICES[MDL][1]) / 1e6
                else:
                    import google.generativeai as genai
                    genai.configure(api_key=api_key)
                    gm = genai.GenerativeModel(MDL, system_instruction=TRANS_PROMPT)
                    raw = gm.generate_content(f"Translate:\n\n{sample}").text
                    pcost = 0

                pp = parse_trans(raw)
                st.session_state['tprev'] = pp
                st.success(f"✅ {len(pp)} pages previewed | Cost: ${pcost:.4f}")
            except Exception as e:
                st.error(f"❌ {e}")

    if st.session_state.get('tprev'):
        pp = st.session_state['tprev']
        col1, col2 = st.columns(2)
        for col, dk, nm in [(col1,"v2","🔷 V2"),(col2,"v4","⬛ V4")]:
            with col:
                st.markdown(f"**{nm}**")
                pv = make_preview_pdf(pp, dk, S, wm_css, wm_html, HDR_L, HDR_R, FOOT)
                show_preview(pv, ["Cover","Page 1","Page 2"])
                st.download_button(f"⬇️ {nm} Preview", pv, f"tprev_{dk}.pdf",
                    "application/pdf", use_container_width=True, key=f"tpdl_{dk}")
    st.divider()

    # ─── FULL TRANSLATION WITH 50-PAGE PAUSE ──────────────────
    st.markdown("### 🚀 Translate & Generate")
    REVIEW_BATCH = 50
    API_BATCH = 5

    if "ts" not in st.session_state:
        st.session_state.ts = dict(on=False, ri=0, conf=[], cost=0.0)
    ts = st.session_state.ts

    n_review = math.ceil(n_translate / REVIEW_BATCH)

    if not ts["on"]:
        if st.button(f"📚 Start Translation (p{pg_s}–{pg_e}, {n_translate} pages)",
                     type="primary", use_container_width=True):
            st.session_state.ts = dict(on=True, ri=0, conf=[], cost=0.0)
            st.rerun()
        st.stop()

    # ── Batch processing ──
    ri = ts["ri"]

    if ri >= n_review:
        # Done → generate output
        final = ts["conf"]
        st.markdown("### ✅ Translation Complete!")
        st.markdown(f"""
        <div class="cost-card"><span class="lbl">Pages </span><span class="val">{len(final)}</span></div>
        <div class="cost-card"><span class="lbl">Total $ </span><span class="val">${ts['cost']:.4f}</span></div>
        <div class="cost-card"><span class="lbl">BDT </span><span class="val">৳{ts['cost']*120:.0f}</span></div>
        """, unsafe_allow_html=True)
        st.markdown("")
        generate_and_download(final, "Final ")
        if st.button("🔄 New Translation"):
            st.session_state.ts = dict(on=False, ri=0, conf=[], cost=0.0)
            st.rerun()
    else:
        rb_s = pg_s + ri * REVIEW_BATCH
        rb_e = min(rb_s + REVIEW_BATCH - 1, pg_e)
        n_this = rb_e - rb_s + 1

        st.markdown(f"### 📋 Batch {ri+1}/{n_review} — Pages {rb_s}–{rb_e}")
        st.progress(ri / n_review)

        # Live cost tracker
        st.markdown(f"""
        <div class="cost-card"><span class="lbl">Confirmed </span><span class="val">{len(ts['conf'])} pg</span></div>
        <div class="cost-card"><span class="lbl">Cost </span><span class="val">${ts['cost']:.4f}</span></div>
        <div class="cost-card"><span class="lbl">BDT </span><span class="val">৳{ts['cost']*120:.0f}</span></div>
        """, unsafe_allow_html=True)
        st.markdown("")

        bkey = f"_tb_{ri}"
        if bkey not in st.session_state:
            # Translate this review-batch via API in sub-batches of 5
            translated = []
            bcost = 0.0
            n_api = math.ceil(n_this / API_BATCH)
            prog = st.progress(0)
            stat = st.empty()
            clive = st.empty()

            for ab in range(n_api):
                a_s = rb_s - 1 + ab * API_BATCH
                a_e = min(a_s + API_BATCH, rb_e)
                txt = ""
                for i in range(a_s, a_e):
                    if i < total_pg:
                        txt += f"\n=== PAGE {i+1} ===\n{pdfdoc[i].get_text()}\n"
                stat.text(f"Translating pages {a_s+1}–{a_e} ({ab+1}/{n_api})...")

                try:
                    if api_prov == "Anthropic":
                        import anthropic
                        cl = anthropic.Anthropic(api_key=api_key)
                        rsp = cl.messages.create(model=MDL, max_tokens=8000,
                            system=TRANS_PROMPT,
                            messages=[{"role":"user","content":f"Translate:\n\n{txt}"}])
                        raw = rsp.content[0].text
                        bcost += (rsp.usage.input_tokens * PRICES[MDL][0] +
                                  rsp.usage.output_tokens * PRICES[MDL][1]) / 1e6
                    elif api_prov == "OpenAI":
                        from openai import OpenAI
                        cl = OpenAI(api_key=api_key)
                        rsp = cl.chat.completions.create(model=MDL, max_tokens=8000,
                            messages=[{"role":"system","content":TRANS_PROMPT},
                                      {"role":"user","content":f"Translate:\n\n{txt}"}])
                        raw = rsp.choices[0].message.content
                        bcost += (rsp.usage.prompt_tokens * PRICES[MDL][0] +
                                  rsp.usage.completion_tokens * PRICES[MDL][1]) / 1e6
                    else:
                        import google.generativeai as genai
                        genai.configure(api_key=api_key)
                        gm = genai.GenerativeModel(MDL, system_instruction=TRANS_PROMPT)
                        raw = gm.generate_content(f"Translate:\n\n{txt}").text

                    translated.extend(parse_trans(raw))
                except Exception as e:
                    st.error(f"Batch error: {e}")

                prog.progress((ab+1)/n_api)
                clive.text(f"💰 Batch: ${bcost:.4f} | Total: ${ts['cost']+bcost:.4f} "
                          f"(৳{(ts['cost']+bcost)*120:.0f})")

            stat.text(f"✅ {len(translated)} pages translated (${bcost:.4f})")
            st.session_state[bkey] = dict(pages=translated, cost=bcost)

        bd = st.session_state[bkey]
        bpages = bd["pages"]

        st.success(f"✅ {len(bpages)} pages ready for review")

        # Exact preview
        if bpages:
            with st.spinner("Rendering preview..."):
                bpv = make_preview_pdf(bpages, DK, S, wm_css, wm_html, HDR_L, HDR_R, FOOT)
            pc1, pc2 = st.columns(2)
            imgs = pdf_to_images(bpv, dpi=150)
            for idx2, img in enumerate(imgs):
                with (pc1 if idx2 == 0 else pc2):
                    b64 = base64.b64encode(img).decode()
                    st.markdown(f'<img src="data:image/png;base64,{b64}" class="page-thumb" '
                               f'style="width:100%;">', unsafe_allow_html=True)

        with st.expander(f"📖 {len(bpages)} translated titles"):
            for p in bpages:
                st.text(f"পৃষ্ঠা {p['page_num']}: {p['title_bn']}")

        st.markdown("---")
        a1, a2, a3 = st.columns(3)
        with a1:
            if st.button("✅ Confirm & Next", type="primary", use_container_width=True, key="tc"):
                ts["conf"].extend(bpages); ts["cost"] += bd["cost"]
                ts["ri"] += 1; st.rerun()
        with a2:
            if st.button("⏭️ Skip", use_container_width=True, key="tsk"):
                ts["ri"] += 1; st.rerun()
        with a3:
            if st.button("🛑 Stop & Generate", use_container_width=True, key="tstp"):
                ts["conf"].extend(bpages); ts["cost"] += bd["cost"]
                ts["ri"] = n_review; st.rerun()


# ═══════════════════════════════════════════════════════════════
# FOOTER
# ═══════════════════════════════════════════════════════════════
st.divider()
st.markdown('<p style="text-align:center;color:#444;font-size:0.7rem;">'
            'অদম্য প্রেস Ebook Creator v2.0 • Online Tech Academy</p>',
            unsafe_allow_html=True)
