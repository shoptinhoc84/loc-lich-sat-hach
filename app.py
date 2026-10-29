import re
from io import BytesIO
from datetime import datetime
import pandas as pd
import streamlit as st
import pdfplumber
import fitz
from docx import Document
from docx.shared import Inches, Pt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from xml.sax.saxutils import escape

st.set_page_config(page_title="Lọc lịch sát hạch Nguyễn Trình", layout="wide")
st.title("Lọc lịch thi / sát hạch tại Trung tâm Nguyễn Trình")
st.caption("Tải PDF lên → nhận diện cột → kiểm tra kết quả → xuất Word, PDF, Excel.")

VENUE_PATTERNS = [
    r"trung tâm giáo dục nghề nghiệp và sát hạch lái xe nguyễn trình",
    r"trung tâm gdnn và shlx nguyễn trình",
    r"trung tâm sát hạch lái xe nguyễn trình",
    r"trung tâm nguyễn trình",
    r"trung tam giao duc nghe nghiep va sat hach lai xe nguyen trinh",
    r"trung tam gdnn va shlx nguyen trinh",
    r"trung tam sat hach lai xe nguyen trinh",
    r"trung tam nguyen trinh",
]
DATE_RE = re.compile(r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](20\d{2})(?!\d)")

def norm(s):
    return re.sub(r"\s+", " ", str(s or "").replace("\u00a0", " ")).strip()

def date_norm(text):
    m = DATE_RE.search(norm(text))
    if not m:
        return ""
    try:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))).strftime("%d/%m/%Y")
    except ValueError:
        return ""

def fold(text):
    # Accent-insensitive matching for Vietnamese terms.
    import unicodedata
    s = unicodedata.normalize("NFD", str(text or "").casefold())
    return "".join(ch for ch in s if unicodedata.category(ch) != "Mn").replace("đ", "d")

def is_venue(text):
    t = fold(norm(text))
    if not t:
        return False
    # Require both Nguyễn Trình and an explicit centre/sat-hach context.
    has_name = "nguyen trinh" in t
    has_context = ("trung tam" in t) or ("sat hach" in t)
    return has_name and has_context

def col_role(header):
    h = fold(header)
    if any(x in h for x in ["thoi gian sat hach", "ngay thi", "ngay sat hach", "thoi gian thi"]):
        return "date"
    if any(x in h for x in ["co so dao tao", "don vi dao tao", "truong/trung tam dao tao", "trung tam dao tao"]):
        return "provider"
    if any(x in h for x in ["dia diem", "noi to chuc sat hach", "dia diem to chuc sat hach", "trung tam sat hach"]):
        return "venue"
    return ""

def table_rows(pdf_bytes):
    rows = []
    table_diagnostics = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for pageno, page in enumerate(pdf.pages, start=1):
            try:
                tables = page.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_tolerance": 5,
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                    "edge_min_length": 12,
                    "min_words_vertical": 1,
                    "min_words_horizontal": 1,
                }) or []
                if not tables:
                    tables = page.extract_tables() or []
            except Exception:
                tables = page.extract_tables() or []
            for table in tables:
                if not table:
                    continue
                # Header may appear on page 1 only; infer column roles from first rows.
                header_idx = None
                roles = {}
                for ridx, row in enumerate(table[:5]):
                    cells = [norm(c) for c in (row or [])]
                    joined = " ".join(cells)
                    if any(k in fold(joined) for k in ["co so dao tao", "thoi gian sat hach", "dia diem to chuc sat hach"]):
                        header_idx = ridx
                        for ci, cell in enumerate(cells):
                            role = col_role(cell)
                            if role:
                                roles[role] = ci
                        break

                for ridx, row in enumerate(table):
                    cells = [norm(c) for c in (row or [])]
                    if not any(cells):
                        continue
                    joined = " | ".join(c for c in cells if c)
                    if header_idx is not None and ridx <= header_idx:
                        continue

                    # Find date by looking first in detected date column, then across cells.
                    date = ""
                    date_col = roles.get("date")
                    if date_col is not None and date_col < len(cells):
                        date = date_norm(cells[date_col])
                    if not date:
                        date = date_norm(joined)
                    if not date:
                        # Continuation rows can contain venue text without a date.
                        if rows and rows[-1]["Trang"] == pageno and joined:
                            # Append only when likely to be a continuation of the previous row.
                            rows[-1]["Raw PDF row"] += " | " + joined
                            if not rows[-1]["Địa điểm sát hạch"]:
                                rows[-1]["Địa điểm sát hạch"] = joined
                                rows[-1]["Ghi chú kiểm tra"] = "Cần xác minh: có dòng nối tiếp"
                        continue

                    provider = cells[roles["provider"]] if "provider" in roles and roles["provider"] < len(cells) else ""
                    venue = cells[roles["venue"]] if "venue" in roles and roles["venue"] < len(cells) else ""

                    # Fallback for documents with no repeated column header: use table shape.
                    if not venue and len(cells) >= 4:
                        venue = cells[-1]
                    if not provider and len(cells) >= 2:
                        provider = cells[1]

                    rows.append({
                        "Ngày thi": date,
                        "Cơ sở đào tạo": provider,
                        "Địa điểm sát hạch": venue,
                        "Raw PDF row": joined,
                        "Trang": pageno,
                        "Ghi chú kiểm tra": "Đã nhận diện cột" if (provider and venue) else "Cần kiểm tra: cột chưa nhận diện chắc chắn",
                    })
                table_diagnostics.append({"Trang": pageno, "Số bảng": len(tables), "Cột": str(roles)})
    return rows, table_diagnostics

def fallback_text_rows(pdf_bytes):
    """Text fallback is deliberately marked unconfirmed; it prevents silent omission."""
    out = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for pageno, page in enumerate(doc, start=1):
        text = page.get_text("text")
        lines = [norm(x) for x in text.splitlines() if norm(x)]
        for i, line in enumerate(lines):
            dt = date_norm(line)
            if not dt:
                continue
            chunk = " | ".join(lines[max(0, i-8):min(len(lines), i+10)])
            out.append({
                "Ngày thi": dt,
                "Cơ sở đào tạo": "",
                "Địa điểm sát hạch": "",
                "Raw PDF row": chunk,
                "Trang": pageno,
                "Ghi chú kiểm tra": "CHƯA XÁC NHẬN: trích xuất dự phòng theo văn bản, kiểm tra trang gốc",
            })
    seen, unique = set(), []
    for r in out:
        key = (r["Ngày thi"], r["Raw PDF row"], r["Trang"])
        if key not in seen:
            seen.add(key); unique.append(r)
    return unique

def extract_all(pdf_bytes):
    try:
        rows, diagnostics = table_rows(pdf_bytes)
    except Exception:
        rows, diagnostics = [], []
    # If table extraction is suspiciously sparse, merge in fallback rows as review-only candidates.
    if len(rows) < 2:
        fallback = fallback_text_rows(pdf_bytes)
        existing = {(r["Ngày thi"], r["Trang"]) for r in rows}
        for r in fallback:
            if (r["Ngày thi"], r["Trang"]) not in existing:
                rows.append(r)
    return rows, diagnostics

def filter_venue(df):
    if df.empty:
        return df.copy()
    confirmed = df["Địa điểm sát hạch"].fillna("").astype(str).apply(is_venue)
    return df[confirmed].copy()

def make_excel(df, all_df=None):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dia diem Nguyen Trinh")
        if all_df is not None:
            all_df.to_excel(writer, index=False, sheet_name="Tat ca de kiem tra")
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                width = min(max(max(len(str(c.value or "")) for c in col) + 2 for c in [col[0]]), 55)
                ws.column_dimensions[col[0].column_letter].width = width
    bio.seek(0)
    return bio.getvalue()

def make_docx(df):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(.5); sec.bottom_margin = Inches(.5)
    doc.add_heading("LỊCH SÁT HẠCH TẠI TRUNG TÂM NGUYỄN TRÌNH", 0)
    p = doc.add_paragraph("Danh sách trích xuất tự động. Vui lòng đối chiếu PDF gốc trước khi thông báo thí sinh.")
    p.runs[0].font.size = Pt(9)
    cols = ["Ngày thi", "Cơ sở đào tạo", "Địa điểm sát hạch", "Trang", "Ghi chú kiểm tra"]
    table = doc.add_table(rows=1, cols=len(cols)); table.style = "Table Grid"
    for i, c in enumerate(cols): table.rows[0].cells[i].text = c
    for _, row in df.iterrows():
        cells = table.add_row().cells
        for i, c in enumerate(cols): cells[i].text = str(row.get(c, "") or "")
    bio = BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()

def rtf_escape(s):
    out = []
    for ch in str(s):
        if ch == "\\": out.append(r"\\")
        elif ch == "{": out.append(r"\{")
        elif ch == "}": out.append(r"\}")
        elif ch == "\n": out.append(r"\line ")
        elif ord(ch) < 128: out.append(ch)
        else: out.append(r"\u" + str(ord(ch) if ord(ch) < 32768 else ord(ch)-65536) + "?")
    return "".join(out)

def make_doc(df):
    cols = ["Ngày thi", "Cơ sở đào tạo", "Địa điểm sát hạch", "Trang", "Ghi chú kiểm tra"]
    parts = [r"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\f0\fs18 ",
             r"\b " + rtf_escape("LỊCH SÁT HẠCH TẠI TRUNG TÂM NGUYỄN TRÌNH") + r"\b0\par\par ",
             rtf_escape("Danh sách trích xuất tự động. Vui lòng đối chiếu PDF gốc trước khi thông báo thí sinh.") + r"\par\par "]
    parts.append(r"\b " + r" | ".join(rtf_escape(c) for c in cols) + r"\b0\par ")
    for _, row in df.iterrows():
        parts.append(" | ".join(rtf_escape(row.get(c, "") or "") for c in cols) + r"\par ")
    parts.append("}")
    return "".join(parts).encode("ascii", errors="ignore")

def make_pdf(df):
    bio = BytesIO()
    styles = getSampleStyleSheet()
    title = ParagraphStyle("Title2", parent=styles["Title"], fontSize=13, leading=16, alignment=TA_CENTER)
    body = ParagraphStyle("Body2", parent=styles["BodyText"], fontSize=7, leading=9, wordWrap="CJK")
    doc = SimpleDocTemplate(bio, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=22, bottomMargin=22)
    story = [Paragraph("LỊCH SÁT HẠCH TẠI TRUNG TÂM NGUYỄN TRÌNH", title), Spacer(1, 8),
             Paragraph("Danh sách trích xuất tự động. Vui lòng đối chiếu PDF gốc trước khi thông báo thí sinh.", body), Spacer(1, 8)]
    cols = ["Ngày thi", "Cơ sở đào tạo", "Địa điểm sát hạch", "Trang", "Ghi chú kiểm tra"]
    data = [[Paragraph(escape(c), body) for c in cols]]
    for _, row in df.iterrows():
        data.append([Paragraph(escape(str(row.get(c, "") or "")), body) for c in cols])
    tbl = Table(data, colWidths=[65, 180, 180, 30, 260], repeatRows=1)
    tbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("GRID",(0,0),(-1,-1),.4,colors.grey),
                             ("VALIGN",(0,0),(-1,-1),"TOP"),("FONTSIZE",(0,0),(-1,-1),7)]))
    story.append(tbl); doc.build(story); bio.seek(0); return bio.getvalue()

uploaded = st.file_uploader("Tải thông báo lịch sát hạch PDF", type=["pdf"])
if uploaded:
    pdf_bytes = uploaded.getvalue()
    with st.spinner("Đang phân tích bảng PDF..."):
        rows, diagnostics = extract_all(pdf_bytes)
    if not rows:
        st.error("Không trích xuất được văn bản. Nếu đây là PDF scan, cần thêm OCR để đọc ảnh.")
        st.stop()
    df = pd.DataFrame(rows)
    for c in ["Ngày thi", "Cơ sở đào tạo", "Địa điểm sát hạch", "Raw PDF row", "Trang", "Ghi chú kiểm tra"]:
        if c not in df: df[c] = ""
    df["_sort"] = pd.to_datetime(df["Ngày thi"], format="%d/%m/%Y", errors="coerce")
    df = df.sort_values(["_sort", "Trang"], kind="stable").drop(columns=["_sort"])
    filtered = filter_venue(df)
    review = df[df["Ghi chú kiểm tra"].astype(str).str.contains("CHƯA XÁC NHẬN|Cần kiểm tra", case=False, na=False)].copy()

    st.success(f"Trích xuất được {len(df)} dòng ứng viên; nhận diện trực tiếp {len(filtered)} dòng có địa điểm Nguyễn Trình.")
    st.warning("Để hạn chế bỏ sót, hãy xem cả bảng 'Cần rà soát thủ công'. Nếu cột địa điểm không được nhận diện từ PDF, chương trình không tự coi đó là kết quả đã xác nhận.")
    st.subheader("A. Kết quả xác nhận theo cột địa điểm sát hạch")
    edited = st.data_editor(filtered, num_rows="dynamic", use_container_width=True, key="confirmed")
    st.subheader("B. Cần rà soát thủ công để tránh bỏ sót")
    if len(review):
        st.dataframe(review, use_container_width=True)
    else:
        st.success("Không có dòng nào bị đánh dấu cần rà soát theo các quy tắc hiện tại.")
    st.subheader("C. Toàn bộ dữ liệu trích xuất")
    with st.expander("Mở toàn bộ bảng"):
        st.dataframe(df, use_container_width=True)
    with st.expander("Chi tiết nhận diện bảng theo trang"):
        st.json(diagnostics)
    if len(edited):
        col1, col2, col3, col4 = st.columns(4)
        col1.download_button("Tải Word .doc", make_doc(edited), "lich_sat_hach_nguyen_trinh.doc", "application/msword")
        col2.download_button("Tải Word .docx", make_docx(edited), "lich_sat_hach_nguyen_trinh.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        col3.download_button("Tải PDF", make_pdf(edited), "lich_sat_hach_nguyen_trinh.pdf", "application/pdf")
        col4.download_button("Tải Excel", make_excel(edited, df), "lich_sat_hach_nguyen_trinh.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif len(review):
        st.info("Chưa có dòng nào được xác nhận tự động. Hãy kiểm tra bảng rà soát và PDF gốc; có thể sửa bảng dữ liệu trước khi xuất.")
else:
    st.markdown("### Cách dùng")
    st.markdown("1. Tải PDF thông báo lên.\\n2. Xem kết quả theo cột địa điểm sát hạch.\\n3. Kiểm tra các dòng cần rà soát thủ công.\\n4. Xuất Word, PDF hoặc Excel sau khi đối chiếu.")
