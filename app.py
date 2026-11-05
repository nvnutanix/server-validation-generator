import streamlit as st
from io import BytesIO
from pathlib import Path
import openpyxl, re, html
from pptx import Presentation
from pptx.util import Inches, Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from PIL import Image
import fitz

st.set_page_config(page_title="Server Hardware Validation Report Generator", page_icon="🖥️", layout="wide")
st.title("Server Hardware Validation Report Generator")
st.caption("Upload a validated workbook and chassis visuals to generate a PowerPoint report and Outlook-ready email.")

FONT_NAME = "Montserrat"
HEADLINE_COLOR = "351292"
ASSETS_DIR = Path(__file__).parent / "assets"
DEFAULT_BG_PATH = ASSETS_DIR / "title-background.jpg"
if not DEFAULT_BG_PATH.exists():
    DEFAULT_BG_PATH = Path(__file__).parent / "title-background.jpg"

DEFAULT_DISCLAIMER = "HW BOM validation is performed based on the original service tag/BOM shared by the Nutanix Systems Engineer or customer. Any modification to the original configuration after initial purchase needs to be flagged for accurate results."
DEFAULT_NOTE = 'Note: "N/A" (Not Applicable) is listed for components that are not included in the hardware compatibility lists on the support portal and therefore do not require validation.'

def clean(v):
    return re.sub(r'\s+', ' ', str(v or '')).strip()

def hex_to_rgb(hex_str):
    h = hex_str.lstrip('#')
    return RGBColor(*(int(h[i:i+2], 16) for i in (0, 2, 4)))

def parse_workbook(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    toc = wb[wb.sheetnames[0]]
    all_rows = list(toc.iter_rows(values_only=True))

    # Step 1: Specifically locate the header row of the GROUPS table
    # Ignore any status / legend / Yes-No-Refer comment tables
    header_idx = None
    for idx, r in enumerate(all_rows):
        if not r:
            continue
        c0 = clean(r[0]).lower() if len(r) > 0 else ""
        row_str = " ".join([clean(c).lower() for c in r if c is not None])
        if any(term in row_str for term in ("yes/no", "supported/unsupported", "legend", "status count", "validation status")):
            continue
        if c0 in ("group", "groups", "group name", "group #", "group id"):
            header_idx = idx
            break
        elif "group" in c0 and any(m in row_str for m in ("model", "server", "platform", "tag", "qty")):
            header_idx = idx
            break

    # Fallback if no explicit "group" header was found: first row with 2+ populated non-status columns
    if header_idx is None:
        for idx, r in enumerate(all_rows):
            if not r:
                continue
            row_str = " ".join([clean(c).lower() for c in r if c is not None])
            if any(term in row_str for term in ("yes", "refer comment", "unsupported", "supported")):
                continue
            if sum(1 for c in r if clean(c)) >= 2:
                header_idx = idx
                break

    if header_idx is None:
        header_idx = 0
        raw_header_row = ["Group", "Server Model"]
    else:
        raw_header_row = all_rows[header_idx]

    # Step 2: Detect groups strictly BELOW the header row of the groups table
    group_rows = []
    stop_indicators = (
        "yes", "no", "refer comment", "supported", "not supported",
        "unsupported", "total", "legend", "status", "note", "disclaimer"
    )

    for idx in range(header_idx + 1, len(all_rows)):
        r = all_rows[idx]
        if not r:
            if group_rows:
                break
            continue
        c0 = clean(r[0])
        c0_lower = c0.lower()

        if not c0:
            if group_rows:
                break
            continue

        if any(c0_lower == term or c0_lower.startswith(term) for term in stop_indicators):
            break

        is_sheet = c0 in wb.sheetnames
        is_group_txt = c0_lower.startswith("group")
        has_server_info = len(r) > 1 and bool(clean(r[1]))

        if is_sheet or is_group_txt or has_server_info:
            group_rows.append(r)
        elif group_rows:
            break

    # Step 3: Determine columns for the groups summary table
    max_cols = max([len(raw_header_row)] + [len(r) for r in group_rows], default=2)
    active_cols = [0, 1]
    for col_idx in range(2, max_cols):
        h_val = clean(raw_header_row[col_idx]) if col_idx < len(raw_header_row) else ""
        has_data = any(clean(r[col_idx]) for r in group_rows if col_idx < len(r))
        if h_val or has_data:
            active_cols.append(col_idx)

    active_cols.sort()

    default_names = {0: "Group", 1: "Server Model", 2: "Service Tags", 3: "Quantity"}
    headers = []
    for col_idx in active_cols:
        h_val = clean(raw_header_row[col_idx]) if col_idx < len(raw_header_row) else ""
        if not h_val:
            h_val = default_names.get(col_idx, f"Column {col_idx+1}")
        headers.append(h_val)

    # Step 4: Extract groups and parse corresponding component sheets
    groups = []
    summary_rows = []

    for r in group_rows:
        row_cells = [clean(r[c]) if c < len(r) else "" for c in active_cols]
        summary_rows.append(row_cells)

        sheetname = clean(r[0])
        model = clean(r[1]) if len(r) > 1 else ""
        tags = clean(r[2]) if len(r) > 2 else ""
        qty = clean(r[3]) if len(r) > 3 else ""

        comps = []
        if sheetname in wb.sheetnames:
            ws = wb[sheetname]
            for cr in ws.iter_rows(min_row=2, values_only=True):
                comp, pn, desc, quantity, status, comment = (list(cr[:6]) + [None] * 6)[:6]
                if not any(x is not None for x in (comp, pn, desc, quantity, status, comment)):
                    continue
                if not clean(comp) and not clean(status) and not clean(comment):
                    continue
                comps.append(dict(
                    component=clean(comp),
                    part=clean(pn),
                    description=clean(desc),
                    quantity=clean(quantity),
                    status=clean(status),
                    comment=clean(comment)
                ))

        groups.append(dict(name=sheetname, model=model, tags=tags, qty=qty, components=comps))

    toc_summary = {
        "headers": headers,
        "rows": summary_rows
    }

    return groups, toc_summary

def status_bucket(s):
    s = clean(s).lower()
    if s in ("yes", "supported"):
        return "Supported"
    if "refer" in s or "comment" in s or "info" in s:
        return "Refer comment"
    if s in ("no", "unsupported", "not supported"):
        return "Unsupported"
    if s in ("n/a", "na", "not applicable"):
        return "N/A"
    return "Unspecified"

def reportable_components(group):
    """Only include explicitly validated Yes/No/Refer comment rows."""
    return [c for c in group["components"] if status_bucket(c["status"]) in ("Supported", "Refer comment", "Unsupported")]

def component_line(c):
    label = c["component"] or c["part"] or c["description"]
    qty = f" ({c['quantity']})" if c["quantity"] else ""
    return f"{label}{qty}"

def email_html(customer, use_case, groups, disclaimer, note):
    esc = html.escape
    parts = [
        '<div style="font-family:Aptos,Calibri,sans-serif;font-size:11pt;line-height:1.5;color:#222;">',
        f'<p>Hi Team,</p><p>Please see the hardware validation results for <b>{esc(use_case)}</b> use case below.</p>',
        f'<p><b>Disclaimer</b> - <i>{esc(disclaimer.strip())}</i></p>'
    ]
    for g in groups:
        header_line = f"<b>{esc(g['name'])} – {esc(g['model'])}</b>"
        extra = []
        if g['tags']: extra.append(f"Service Tags: {esc(g['tags'])}")
        if g['qty']: extra.append(f"Quantity: {esc(g['qty'])}")
        if extra:
            header_line += f"<br>{' | '.join(extra)}"
        parts.append(f'<div style="border-top:1px solid rgba(100,100,100,0.18);padding-top:12px;margin-top:20px;"><p style="margin:0 0 10px 0;">{header_line}</p>')
        for bucket, color in (("Supported", "#16803c"), ("Refer comment", "#b77900"), ("Unsupported", "#c62828")):
            items = [c for c in reportable_components(g) if status_bucket(c["status"]) == bucket]
            if not items:
                continue
            parts.append(f'<p style="margin-12px 0 5px 0;"><b><span style="color:{color};">{esc(bucket)}:</span></b></p><ul style="margin-top:0;margin-bottom:12px;padding-left:24px;">')
            for c in items:
                parts.append(f'<li style="margin-bottom:7px;">{esc(component_line(c))}')
                if c["comment"]:
                    parts.append(f'<br><span style="color:#8ab4f8;">{esc(c["comment"])}</span>')
                parts.append('</li>')
            parts.append('</ul>')
        parts.append('</div>')
    parts.append(f'<p style="margin-top:18px;">{esc(note.strip())}</p>')
    parts.append('<p>Please see the updated validation document in the attachment.</p></div>')
    return ''.join(parts)

def ppt_text(slide, x, y, w, h, text, size=16, color=(30, 45, 70), bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    p.font.name = FONT_NAME
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = RGBColor(*color) if isinstance(color, tuple) else color
    return box

def set_cell_border(cell, color="000000", width_pt=0.75):
    tcPr = cell._tc.get_or_add_tcPr()
    width_emu = int(width_pt * 12700)
    for border_name in ('lnL', 'lnR', 'lnT', 'lnB'):
        existing = tcPr.find(f'{{http://schemas.openxmlformats.org/drawingml/2006/main}}{border_name}')
        if existing is not None:
            tcPr.remove(existing)
        line_xml = f'<a:{border_name} {nsdecls("a")} w="{width_emu}" cmpd="s"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:{border_name}>'
        tcPr.append(parse_xml(line_xml))

def style_cell(cell, text, bg_rgb, fg_rgb, bold=False, font_size=7, border_color="000000"):
    cell.text = str(text or "")
    cell.fill.solid()
    cell.fill.fore_color.rgb = bg_rgb
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.08)
    cell.margin_right = Inches(0.08)
    cell.margin_top = Inches(0.04)
    cell.margin_bottom = Inches(0.04)
    if border_color:
        set_cell_border(cell, color=border_color, width_pt=0.75)
    for p in cell.text_frame.paragraphs:
        p.font.name = FONT_NAME
        p.font.size = Pt(font_size)
        p.font.bold = bold
        p.font.color.rgb = fg_rgb

def set_table_row_heights(table, height_cm=1.0):
    for row in table.rows:
        row.height = Cm(height_cm)

def compute_column_widths(num_cols, total_width_inches=12.4):
    if num_cols == 2:
        return [Inches(4.4), Inches(8.0)]
    elif num_cols == 3:
        return [Inches(3.0), Inches(5.0), Inches(4.4)]
    elif num_cols == 4:
        return [Inches(2.2), Inches(3.4), Inches(5.0), Inches(1.8)]
    else:
        each = total_width_inches / num_cols
        return [Inches(each)] * num_cols

def build_ppt(customer, groups, toc_summary, images):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    headline_rgb = hex_to_rgb(HEADLINE_COLOR)

    # --- Slide 1: Cover Slide ---
    s1 = prs.slides.add_slide(blank)
    bg_bytes = None
    if DEFAULT_BG_PATH.exists():
        try:
            bg_bytes = DEFAULT_BG_PATH.read_bytes()
        except Exception:
            bg_bytes = None

    if bg_bytes:
        s1.shapes.add_picture(BytesIO(bg_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
    else:
        s1.background.fill.solid()
        s1.background.fill.fore_color.rgb = RGBColor(18, 10, 31)

    if customer:
        ppt_text(s1, 1.0, 2.2, 11.333, 0.8, str(customer), size=32, color=(255, 255, 255), bold=True, align=PP_ALIGN.CENTER)
    ppt_text(s1, 1.0, 3.1, 11.333, 1.6, "Hardware Validation Report", size=32, color=(255, 255, 255), bold=True, align=PP_ALIGN.CENTER)

    # --- Slide 2: Validation Summary ---
    summary_headers = toc_summary.get("headers", ["Group", "Server Model"])
    summary_rows = toc_summary.get("rows", [])
    if not summary_rows:
        summary_rows = [[g["name"], g["model"]] for g in groups]

    summary_chunk = 12
    purple_header = hex_to_rgb("744EF4")
    white_text = RGBColor(255, 255, 255)
    body_text = RGBColor(40, 40, 40)
    col_widths = compute_column_widths(len(summary_headers), 12.4)

    for s_idx in range(0, max(1, len(summary_rows)), summary_chunk):
        s2 = prs.slides.add_slide(blank)
        slide_title = "Validation Summary" if s_idx == 0 else "Validation Summary (continued)"
        ppt_text(s2, 0.45, 0.35, 12.0, 0.55, slide_title, size=18, color=headline_rgb, bold=True)

        batch = summary_rows[s_idx:s_idx + summary_chunk]
        table_data = [summary_headers] + batch

        t_shape = s2.shapes.add_table(len(table_data), len(summary_headers), Inches(0.45), Inches(1.1), Inches(12.4), Cm(len(table_data)))
        table = t_shape.table
        set_table_row_heights(table, 1.0)

        for j, col_w in enumerate(col_widths):
            table.columns[j].width = col_w

        for i, row in enumerate(table_data):
            for j, val in enumerate(row):
                cell = table.cell(i, j)
                if i == 0:
                    style_cell(cell, val, purple_header, white_text, bold=True, font_size=7, border_color="000000")
                else:
                    bg = RGBColor(248, 249, 252) if i % 2 == 1 else RGBColor(255, 255, 255)
                    style_cell(cell, val, bg, body_text, bold=False, font_size=7, border_color="000000")

    # --- Group Slides ---
    diagram_note = (
        "Note: This diagram is for visual representation only and not to be considered as actual image; "
        "for detailed validation summary refer to the SE Support email.  |  Disclaimer: Refer to slide 2 for the validation summary."
    )

    header_fill = hex_to_rgb("D0D0D0")
    supported_fill = hex_to_rgb("BFE5A9")
    refer_fill = hex_to_rgb("FFFD78")
    unsupported_fill = hex_to_rgb("FA8072")
    default_row_fill = RGBColor(245, 247, 250)

    for g in groups:
        img = images.get(g["name"])
        if img:
            s_img = prs.slides.add_slide(blank)
            ppt_text(s_img, 0.45, 0.25, 12, 0.55, f"{g['name']} — {g['model']}", size=20, color=headline_rgb, bold=True)

            with Image.open(BytesIO(img)) as source:
                source.load()
                iw, ih = source.size
            max_w, max_h = Cm(25), Cm(14.0)
            scale = min(max_w / iw, max_h / ih)
            w, h = int(iw * scale), int(ih * scale)
            left = int((prs.slide_width - w) / 2)
            top = int(Cm(2.0) + (Cm(14.0) - h) / 2)
            s_img.shapes.add_picture(BytesIO(img), left, top, width=w, height=h)

            note_box = s_img.shapes.add_textbox(Inches(0.5), Inches(6.9), Inches(12.33), Inches(0.45))
            ntf = note_box.text_frame
            ntf.word_wrap = True
            np_elem = ntf.paragraphs[0]
            np_elem.text = diagram_note
            np_elem.alignment = PP_ALIGN.CENTER
            np_elem.font.name = FONT_NAME
            np_elem.font.size = Pt(7.2)
            np_elem.font.color.rgb = RGBColor(127, 127, 127)

        # Server Component Validation Tables
        items = [c for c in g["components"] if status_bucket(c["status"]) in ("Supported", "Refer comment", "Unsupported")]
        data = [["Component", "Qty", "Supported", "Comment"]]
        data += [[c["component"], c["quantity"], c["status"], c["comment"]] for c in items]
        chunk = 12

        for start in range(1, len(data), chunk):
            s_tbl = prs.slides.add_slide(blank)
            title = f"{g['name']} — Component Validation Results"
            if start > 1:
                title += " (continued)"
            ppt_text(s_tbl, 0.45, 0.25, 12, 0.55, title, size=20, color=headline_rgb, bold=True)
            sub = [data[0]] + data[start:start+chunk]

            t_shape = s_tbl.shapes.add_table(len(sub), 4, Inches(0.35), Inches(1.0), Inches(12.6), Cm(len(sub)))
            t = t_shape.table
            set_table_row_heights(t, 1.0)

            for j, col_w in enumerate([5.0, 1.0, 2.0, 4.6]):
                t.columns[j].width = Inches(col_w)

            for i, row in enumerate(sub):
                for j, val in enumerate(row):
                    cell = t.cell(i, j)
                    if i == 0:
                        style_cell(cell, val, header_fill, RGBColor(30, 30, 30), bold=True, font_size=7, border_color="000000")
                    else:
                        bucket = status_bucket(row[2])
                        if bucket == "Supported":
                            cell_bg = supported_fill
                        elif bucket == "Refer comment":
                            cell_bg = refer_fill
                        elif bucket == "Unsupported":
                            cell_bg = unsupported_fill
                        else:
                            cell_bg = default_row_fill
                        style_cell(cell, val, cell_bg, RGBColor(30, 45, 70), bold=False, font_size=7, border_color="000000")

    bio = BytesIO()
    prs.save(bio)
    return bio.getvalue()

# --- Sidebar ---
st.sidebar.header("Report settings")
customer = st.sidebar.text_input("Customer name", placeholder="Customer / organization")
use_case = st.sidebar.selectbox("Use case", ["HCI", "Compute Only"])

# --- Main App ---
excel = st.file_uploader("Validated Excel workbook", type=["xlsx"])
visuals = st.file_uploader("Chassis visual: image or PDF (PDF pages follow group order)", type=["png", "jpg", "jpeg", "pdf"])

if excel:
    try:
        groups, toc_summary = parse_workbook(excel)
        st.success(f"Workbook loaded: {len(groups)} group(s) detected.")
        st.dataframe([{"Group": g["name"], "Server Model": g["model"], "Service Tags": g["tags"], "Qty": g["qty"], "Components": len(g["components"])} for g in groups], use_container_width=True)
        images = {}
        if visuals:
            raw = visuals.getvalue()
            if visuals.name.lower().endswith(".pdf"):
                doc = fitz.open(stream=raw, filetype="pdf")
                for i, g in enumerate(groups):
                    if i < len(doc):
                        pix = doc[i].get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
                        images[g["name"]] = pix.tobytes("png")
            else:
                for g in groups:
                    images[g["name"]] = raw

        st.subheader("Generate outputs")

        clean_customer = re.sub(r'[\\/*?:"<>|]', "", customer.strip()) if customer else "Customer"
        clean_use_case = re.sub(r'[\\/*?:"<>|]', "", use_case.strip()) if use_case else "Validation"
        ppt_download_name = f"{clean_customer} - {clean_use_case} validation.pptx"

        if visuals:
            col_email, col_ppt = st.columns([1, 1])
            with col_email:
                if st.button("Generate email", type="primary", use_container_width=True):
                    st.session_state["validation_email_html"] = email_html(customer, use_case, groups, DEFAULT_DISCLAIMER, DEFAULT_NOTE)
            with col_ppt:
                ppt_data = build_ppt(customer or "[Customer Name]", groups, toc_summary, images)
                st.download_button(
                    label="Generate PowerPoint",
                    data=ppt_data,
                    file_name=ppt_download_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True
                )
        else:
            if st.button("Generate email", type="primary"):
                st.session_state["validation_email_html"] = email_html(customer, use_case, groups, DEFAULT_DISCLAIMER, DEFAULT_NOTE)
            st.info("Upload a chassis image or PDF to enable PowerPoint generation. Email generation is available with Excel alone.")

        if "validation_email_html" in st.session_state:
            email = st.session_state["validation_email_html"]
            st.caption("Use Copy email to copy the rendered HTML with its formatting into Outlook.")
            copy_html = """<!doctype html><html><body style="font-family:Aptos,Calibri,sans-serif;font-size:11pt;margin:0;padding:8px;color:#222">
            <button id="copy" style="margin-bottom:12px;padding:8px 14px;cursor:pointer">Copy email</button>
            <span id="status" style="margin-left:8px;font-size:13px"></span>
            <div id="email">""" + email + """</div>
            <script>
            const btn=document.getElementById('copy'), status=document.getElementById('status');
            btn.addEventListener('click', async () => {
              const html=document.getElementById('email').innerHTML;
              const plain=document.getElementById('email').innerText;
              try {
                if (!navigator.clipboard || !window.ClipboardItem) throw new Error('Rich clipboard unavailable');
                await navigator.clipboard.write([new ClipboardItem({'text/html':new Blob([html],{type:'text/html'}),'text/plain':new Blob([plain],{type:'text/plain'})})]);
                status.textContent='Email copied. Paste it into Outlook.';
              } catch(e) {
                status.textContent='Copy was blocked by browser permissions.';
              }
            });
            </script></body></html>"""
            st.components.v1.html(copy_html, height=700, scrolling=True)

    except Exception as e:
        st.error(f"Could not read workbook: {e}")
else:
    st.info("Upload the validated Excel workbook to begin. The first worksheet is read as the table of contents; remaining group sheets are parsed using the Component, Part Number, Description, Quantity, Status and Comment columns.")