import io
import re
import pandas as pd
import streamlit as st

# ---------- Page setup ----------
st.set_page_config(
    page_title="Orbico Invoice → Excel",
    page_icon="📊",
    layout="wide"
)

# ---- Optional logo (put a file named 'logo.png' in the repo root) ----
logo_col, title_col = st.columns([1, 6], gap="small")
with logo_col:
    try:
        st.image("logo.png", width=72)  # replace with your company logo file
    except Exception:
        st.write("")  # no logo, ignore

with title_col:
    st.markdown(
        "<h1 style='margin-bottom:0.25rem;'>Orbico Invoice → Excel</h1>"
        "<p style='color:#5b6b7a;margin-top:0;'>Upload a PDF invoice to extract line items and compute "
        "<strong>Stvarna količina (kom)</strong> using L/KG rules.</p>",
        unsafe_allow_html=True,
    )

# ---------- Light CSS polish ----------
st.markdown(
    """
    <style>
      /* tighter default spacing & smoother cards */
      .block-container {padding-top: 2rem; padding-bottom: 2.5rem; max-width: 1100px;}
      .stDownloadButton button, .stButton button { border-radius: 10px; padding: 0.6rem 1rem; }
      .stAlert { border-radius: 10px; }
      .app-card {
        border: 1px solid #e8edf3;
        border-radius: 14px;
        padding: 1rem 1.25rem;
        background: #fff;
        box-shadow: 0 1px 2px rgba(16,24,40,.04);
      }
      .muted { color:#6b7a8c; font-size:0.925rem; }
      .rule { color:#0f172a; background:#f1f5f9; padding:.25rem .5rem; border-radius:.5rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- Sidebar: How it works ----------
with st.sidebar:
    st.markdown("### How it works")
    st.markdown(
        "- **AxB(Unit)** → we use **B** (e.g., `4X4L` → 4 L per bottle; `12X0,4KG` → 0.4 KG)\n"
        "- Otherwise we use the **last** standalone number+unit (e.g., `… 55L` → 55 L)\n"
        "- Supports **L** and **KG**\n"
        "- Avoids false matches like `R4 L`",
    )
    st.markdown("---")
    round_opt = st.checkbox("Round 'Stvarna količina (kom)' to whole numbers", value=False)
    use_alt = st.checkbox("Try alternate parsing strategies", value=True)

# ---------- File uploader card ----------
st.markdown("<div class='app-card'>", unsafe_allow_html=True)
uploaded = st.file_uploader("Drop your PDF here", type=["pdf"])
st.caption("Tip: you can also click **Browse files** and choose a PDF from your computer.")
st.markdown("</div>", unsafe_allow_html=True)

# ---------- Helpers (same logic as before) ----------
def to_float(s: str):
    if s is None:
        return None
    s = s.replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(",", "")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

def extract_text_from_pdf(file_bytes: bytes) -> str:
    text = ""
    # Prefer pdfplumber for better layout extraction
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text() or ""
                text += t + "\n"
        if text.strip():
            return text
    except Exception:
        pass
    # Fallback to PyPDF2
    try:
        import PyPDF2
        reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages:
            t = page.extract_text() or ""
            text += t + "\n"
    except Exception:
        pass
    return text

def parse_lines_simple(text: str):
    """Rows that start with an index number."""
    rows = []
    for ln in text.splitlines():
        s = ln.strip()
        if not re.match(r"^\d+\s", s):
            continue
        parts = s.split()
        if len(parts) < 10:
            continue
        try:
            Ukupno = parts[-1]
            PDV = parts[-2]
            Neto = parts[-3]
            Neto_cijena = parts[-4]
            Kol_za_mjeru = parts[-5]
            Kol = parts[-6]
            Br = parts[0]; Sifra = parts[1]; EAN = parts[2]
            Naziv = " ".join(parts[3:-6])
            rows.append({
                "Br": int(Br),
                "Šifra": Sifra,
                "EAN": EAN,
                "Naziv proizvoda": Naziv,
                "Kol": to_float(Kol),
                "Kol za mjeru": to_float(Kol_za_mjeru),
                "Neto cijena (EUR)": to_float(Neto_cijena),
                "Neto (EUR)": to_float(Neto),
                "PDV (%)": to_float(PDV),
                "Ukupno (EUR)": to_float(Ukupno),
                "Raw line": s
            })
        except Exception:
            continue
    return rows

def parse_lines_between_markers(text: str):
    """Block between Total: and UKUPNA KOLIČINA (if present)."""
    start_idx = text.find("Total:")
    end_idx = text.find("UKUPNA KOLIČINA")
    block = text[start_idx:end_idx] if start_idx != -1 and end_idx != -1 else text
    return parse_lines_simple(block)

def parse_with_tables(file_bytes: bytes):
    """Try pdfplumber table extraction (works even if rows don't start with numbers)."""
    rows = []
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for tbl in tables or []:
                    tbl_clean = [[(c or "").strip() for c in row] for row in tbl if any((c or "").strip() for c in row)]
                    for r in tbl_clean:
                        if len(r) < 8:
                            continue
                        cells = [" ".join(x.split()) for x in r]
                        def is_numlike(x):
                            x = (x or "").replace(" ", "")
                            return bool(re.match(r"^[\d\.,-]+$", x))
                        tail = []
                        for c in reversed(cells):
                            if is_numlike(c):
                                tail.append(c)
                            else:
                                break
                        if len(tail) >= 6:
                            Ukupno, PDV, Neto, Neto_cijena, Kol_za_mjeru, Kol = tail[:6]
                            head = cells[:len(cells)-len(tail)]
                            try:
                                Br = int(re.findall(r"^\d+", head[0])[0]) if head and re.findall(r"^\d+", head[0]) else None
                            except Exception:
                                Br = None
                            Sifra = head[1] if len(head) > 2 else ""
                            EAN = head[2] if len(head) > 3 else ""
                            name_start = 3 if len(head) > 3 else 1
                            Naziv = " ".join(head[name_start:]).strip()
                            if Naziv and to_float(Kol) is not None and to_float(Kol_za_mjeru) is not None:
                                rows.append({
                                    "Br": Br,
                                    "Šifra": Sifra,
                                    "EAN": EAN,
                                    "Naziv proizvoda": Naziv,
                                    "Kol": to_float(Kol),
                                    "Kol za mjeru": to_float(Kol_za_mjeru),
                                    "Neto cijena (EUR)": to_float(Neto_cijena),
                                    "Neto (EUR)": to_float(Neto),
                                    "PDV (%)": to_float(PDV),
                                    "Ukupno (EUR)": to_float(Ukupno),
                                    "Raw line": " | ".join(cells)
                                })
    except Exception:
        return []
    return rows

def extract_denominator(name: str):
    if not name:
        return None, None
    n = name.replace(",", ".")
    # Prefer AxB(Unit) → use B
    m = re.search(r"(?<![A-Za-z])(\d+)\s*[xX]\s*(\d+(?:\.\d+)?)\s*(l|L|kg|KG|Kg|kG)\b", n)
    if m:
        val = float(m.group(2))
        unit = "KG" if "G" in m.group(3).upper() else "L"
        return val, unit
    # Else last standalone number+unit
    tokens = list(re.finditer(r"(?<![A-Za-z])(\d+(?:\.\d+)?)\s*(l|L|kg|KG|Kg|kG)\b", n))
    if tokens:
        last = tokens[-1]
        val = float(last.group(1))
        unit = "KG" if "G" in last.group(2).upper() else "L"
        return val, unit
    return None, None

def compute_real_qty(df: pd.DataFrame, round_qty: bool):
    df[["Denominator value", "Denominator unit"]] = df["Naziv proizvoda"].apply(
        lambda s: pd.Series(extract_denominator(s))
    )
    def real_qty(row):
        val = row["Denominator value"]; kzm = row["Kol za mjeru"]
        if val and kzm is not None and val != 0:
            q = kzm / val
            return round(q) if round_qty else q
        return None
    df["Stvarna količina (kom)"] = df.apply(real_qty, axis=1)
    return df

# ---------- Main flow ----------
if uploaded is not None:
    file_bytes = uploaded.read()
    with st.spinner("Reading PDF…"):
        text = extract_text_from_pdf(file_bytes)

    rows = parse_lines_simple(text)
    if not rows and use_alt:
        rows = parse_lines_between_markers(text)
    if not rows and use_alt:
        rows = parse_with_tables(file_bytes)

    if not rows:
        st.warning("No line items found. Try keeping **Try alternate parsing strategies** enabled. If it still fails, share the Debug preview with us.")
        with st.expander("Debug: raw text preview (first 1500 chars)"):
            st.code(text[:1500] if text else "(no text extracted)")
    else:
        df = pd.DataFrame(rows)
        df = df.sort_values("Br").reset_index(drop=True) if "Br" in df.columns else df.reset_index(drop=True)

        st.markdown("### Preview")
        df = compute_real_qty(df, round_qty=round_opt)
        st.dataframe(df.head(60), use_container_width=True)

        # Excel export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            cols = [
                "Br", "Šifra", "EAN", "Naziv proizvoda",
                "Kol za mjeru", "Denominator value", "Denominator unit", "Stvarna količina (kom)",
                "Kol", "Neto cijena (EUR)", "Neto (EUR)", "PDV (%)", "Ukupno (EUR)"
            ]
            export_cols = [c for c in cols if c in df.columns] + [c for c in df.columns if c not in cols]
            df[export_cols].to_excel(writer, index=False, sheet_name="Stavke")
        output.seek(0)

        st.download_button(
            "⬇️ Download Excel",
            data=output,
            file_name="invoice_extracted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# Footer note
st.markdown("<p class='muted' style='margin-top:2rem;'>Need tweaks (extra columns, rounding rules, categories)? Ping me and I’ll update the app.</p>", unsafe_allow_html=True)
UI update: header, logo support, theme
