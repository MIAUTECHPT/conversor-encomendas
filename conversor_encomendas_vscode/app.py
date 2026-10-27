import io
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber
import qrcode
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

# --- SISTEMA DE LOGIN / PASSWORD ---
def check_password():
    """Retorna True se a palavra-passe estiver correta."""
    def password_entered():
        if st.session_state["password"] == "moldes2026":  # Podes alterar aqui a palavra-passe que quiseres!
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Remove a password da memória por segurança
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # Primeiro acesso: pede a password
        st.title("🔒 Acesso Restrito")
        st.text_input("Introduz a palavra-passe para aceder à aplicação:", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        # Password incorreta
        st.title("🔒 Acesso Restrito")
        st.text_input("Introduz a palavra-passe para aceder à aplicação:", type="password", on_change=password_entered, key="password")
        st.error("😕 Palavra-passe incorreta. Tenta novamente.")
        return False
    else:
        # Autenticado com sucesso
        return True

if not check_password():
    st.stop()  # Impede que o resto da aplicação seja executado se não houver login válido

# --- CONFIGURAÇÃO E VARIÁVEIS DA APLICAÇÃO ---
APP_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

EQUIVALENCES_FILE = APP_DIR / "EQUIVALENCIAS PEDIDOS COM PREÇOS.xlsx"

COLUMNS = [
    "Fornecedor",
    "Nº encomenda",
    "Data documento",
    "Art./Serv.",
    "Descrição",
    "Qtd. Doc.",
    "Und.",
    "Molde e componente",
    "Dt. Prev. Entrega",
    "PDF origem",
]

ITEM_RE = re.compile(
    r"^\s*(\d+)\s+(\d{6,})\s+(.+?)\s+"
    r"(\d+(?:[.,]\d+)?)\s+([A-Za-zÀ-ÿ]{1,8})\s+"
    r".*?\s+(\d{4}-\d{2}-\d{2})\s*$"
)
ORDER_RE = re.compile(r"\bENF\s+([A-Z]?\s*/?\s*[A-Z0-9-]+)\b", re.IGNORECASE)
DATE_RE = re.compile(r"^\s*Data\s+(\d{4}-\d{2}-\d{2})\s*$", re.MULTILINE)


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", str(value)).strip()


def format_qty(val) -> str:
    try:
        if pd.isna(val) or val == "":
            return ""
        s = str(val).replace(",", ".")
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return s.rstrip("0").rstrip(".")
    except Exception:
        return str(val)


@st.cache_data
def load_equivalences() -> pd.DataFrame:
    """Lê a tabela de equivalências: Coluna B (Descrição p/ busca), Coluna C (Correspondente) e Coluna D (Preço)."""
    if not EQUIVALENCES_FILE.exists():
        return pd.DataFrame(columns=["Descricao_Key", "Correspondente_Excel", "P.U."])
    
    try:
        df_eq = pd.read_excel(EQUIVALENCES_FILE, header=1)
        if df_eq.shape[1] >= 4:
            sub_df = df_eq.iloc[:, [1, 2, 3]].copy()
            sub_df.columns = ["Descricao_Excel", "Correspondente_Excel", "P.U."]
        else:
            return pd.DataFrame(columns=["Descricao_Key", "Correspondente_Excel", "P.U."])
        
        sub_df = sub_df.dropna(subset=["Descricao_Excel"])
        
        sub_df["Descricao_Key"] = sub_df["Descricao_Excel"].astype(str).str.lower().str.strip()
        sub_df["Correspondente_Excel"] = sub_df["Correspondente_Excel"].fillna("").astype(str).str.strip()
        sub_df["P.U."] = pd.to_numeric(
            sub_df["P.U."].astype(str).str.replace(",", "."), errors="coerce"
        ).fillna(0.0)
        
        return sub_df[["Descricao_Key", "Correspondente_Excel", "P.U."]]
    except Exception as e:
        st.warning(f"Aviso ao carregar equivalências: {e}")
        return pd.DataFrame(columns=["Descricao_Key", "Correspondente_Excel", "P.U."])


def parse_pdf(file_bytes: bytes, filename: str) -> tuple[list[dict], list[str]]:
    rows = []
    warnings = []
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            pages_text = [(page.extract_text() or "") for page in pdf.pages]
    except Exception as exc:
        return [], [f"{filename}: não foi possível abrir o PDF ({exc})."]

    full_text = "\n".join(pages_text)
    order_match = ORDER_RE.search(full_text)
    order_number = f"ENF {normalize_space(order_match.group(1))}" if order_match else ""
    date_match = DATE_RE.search(full_text)
    document_date = date_match.group(1) if date_match else ""

    # Extração direta do nome da empresa a partir do nome do ficheiro (ex: ENF-IMA-101139.pdf -> IMA)
    supplier = "Empresa"
    clean_filename = filename.upper()
    
    # Procura pelo padrão ENF-[EMPRESA]-[NUMERO] ou similar no nome do ficheiro
    enf_match = re.search(r"ENF[-_]?([A-Z]+)[-_]?\d+", clean_filename)
    if enf_match:
        supplier = enf_match.group(1).strip()
    else:
        # Alternativa caso o nome do ficheiro tenha outra estrutura (ex: IMA_101139.pdf)
        parts = re.split(r"[-_\s]+", Path(filename).stem)
        for part in parts:
            if part.isalpha() and len(part) <= 5:  # Identifica siglas como IMA, IGM, etc.
                supplier = part.upper()
                break

    current_item = None
    for page_text in pages_text:
        for raw_line in page_text.splitlines():
            line = normalize_space(raw_line)
            if not line:
                continue

            item_match = ITEM_RE.match(raw_line)
            if item_match:
                if current_item:
                    rows.append(current_item)
                item_no, code, description, qty, unit, due_date = item_match.groups()
                current_item = {
                    "Fornecedor": supplier,
                    "Nº encomenda": order_number,
                    "Data documento": document_date,
                    "Art./Serv.": code.strip(),
                    "Descrição": normalize_space(description),
                    "Qtd. Doc.": qty.replace(",", "."),
                    "Und.": unit,
                    "Molde e componente": "",
                    "Dt. Prev. Entrega": due_date,
                    "PDF origem": filename,
                    "_Item PDF": item_no,
                }
                continue

            if current_item:
                mould_match = re.search(
                    r"Molde\s*:\s*([^,]+)\s*,\s*Compon\.?\s*:\s*(\S+)",
                    line,
                    re.IGNORECASE,
                )
                if mould_match:
                    current_item["Molde e componente"] = (
                        f"Molde: {mould_match.group(1).strip()}, "
                        f"Compon.: {mould_match.group(2).strip()}"
                    )
                elif (
                    not re.match(r"^\s*\d+\s+\d{6,}\s+", raw_line)
                    and ":" not in line
                    and len(line) <= 45
                    and not any(
                        token in line.lower()
                        for token in (
                            "nota de encomenda", "processado por computador", "página",
                            "total ", "mda -", "conservatória", "rua ", "zona industrial",
                            "tel.:", "e-mail:", "original", "data ", "fornecedor", "contribuinte",
                            "pagamento", "moeda", "item código", "atenção", "tratar com",
                            "local de entrega", "compra sustentável", "3720-502", "santiago da riba",
                            "valor sujeito", "taxa valor", "normal", "desc. doc", "efeitos imediatos"
                        )
                    )
                ):
                    current_item["Descrição"] = normalize_space(
                        current_item["Descrição"] + " " + line
                    )

    if current_item:
        rows.append(current_item)

    if not order_number:
        warnings.append(f"{filename}: não consegui identificar o número da encomenda.")
    if not rows:
        warnings.append(f"{filename}: não foram encontrados artigos no formato esperado.")

    return rows, warnings


def make_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    export_df = df.drop(columns=["_Item PDF"], errors="ignore").copy()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Artigos")
        ws = writer.sheets["Artigos"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for column_cells in ws.columns:
            max_len = max((len(str(c.value or "")) for c in column_cells), default=10)
            ws.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 12), 55)
    return output.getvalue()


def make_pivot_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    
    temp_df = df.copy()
    temp_df["Qtd_Num"] = pd.to_numeric(
        temp_df["Qtd. Doc."].astype(str).str.replace(",", "."), errors="coerce"
    ).fillna(0)
    
    pivot_df = temp_df.groupby(["Art./Serv.", "Descrição", "Und."], as_index=False)["Qtd_Num"].sum()
    pivot_df = pivot_df.rename(columns={"Qtd_Num": "Quantidade Total"})
    
    df_eq = load_equivalences()
    
    if not df_eq.empty:
        pivot_df["Desc_Key"] = pivot_df["Descrição"].astype(str).str.lower().str.strip()
        
        pivot_df = pd.merge(
            pivot_df,
            df_eq,
            left_on="Desc_Key",
            right_on="Descricao_Key",
            how="left"
        )
        
        pivot_df["Correspondente_Fornecedor"] = pivot_df["Correspondente_Excel"].apply(
            lambda x: x if pd.notna(x) and str(x).strip() != "" else None
        )
        pivot_df["Correspondente_Fornecedor"] = pivot_df["Correspondente_Fornecedor"].fillna(pivot_df["Descrição"])
        
        pivot_df["P.U."] = pd.to_numeric(pivot_df["P.U."], errors="coerce").fillna(0.0)
        pivot_df = pivot_df.drop(columns=["Desc_Key", "Descricao_Key", "Correspondente_Excel"], errors="ignore")
    else:
        pivot_df["Correspondente_Fornecedor"] = pivot_df["Descrição"]
        pivot_df["P.U."] = 0.0

    pivot_df["Preço Total (€)"] = pivot_df["Quantidade Total"] * pivot_df["P.U."]
    
    pivot_df = pivot_df.rename(columns={
        "Art./Serv.": "Código Artigo",
        "Descrição": "Descrição (PDF)",
        "Correspondente_Fornecedor": "Correspondente (Fornecedor)",
        "P.U.": "P.U. (€)"
    })

    cols_order = [
        "Código Artigo",
        "Descrição (PDF)",
        "Correspondente (Fornecedor)",
        "Quantidade Total",
        "Und.",
        "P.U. (€)",
        "Preço Total (€)"
    ]
    pivot_df = pivot_df[[c for c in cols_order if c in pivot_df.columns]]
    pivot_df = pivot_df.sort_values(by="Quantidade Total", ascending=False).reset_index(drop=True)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pivot_df.to_excel(writer, index=False, sheet_name="Tabela Dinâmica")
        ws = writer.sheets["Tabela Dinâmica"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for column_cells in ws.columns:
            max_len = max((len(str(c.value or "")) for c in column_cells), default=10)
            ws.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 12), 65)
    return output.getvalue()


def make_labels_pdf(df: pd.DataFrame, layout_choice: str) -> bytes:
    output = io.BytesIO()
    page_w, page_h = A4

    if layout_choice == "2x5 (10 por página)":
        columns_per_page, rows_per_page = 2, 5
        margin, gap_x, gap_y = 8 * mm, 4 * mm, 4 * mm
        qr_max_size = 32 * mm
    elif layout_choice == "2x7 (14 por página)":
        columns_per_page, rows_per_page = 2, 7
        margin, gap_x, gap_y = 7 * mm, 4 * mm, 3 * mm
        qr_max_size = 28 * mm
    else:  # 2x8 (16 por página)
        columns_per_page, rows_per_page = 2, 8
        margin, gap_x, gap_y = 6 * mm, 4 * mm, 2.5 * mm
        qr_max_size = 32 * mm  

    temp_labels_df = df.copy()
    temp_labels_df["Qtd_Num"] = pd.to_numeric(
        temp_labels_df["Qtd. Doc."].astype(str).str.replace(",", "."), errors="coerce"
    ).fillna(0)

    grouped_labels = []
    for _, group in temp_labels_df.groupby(["Fornecedor", "Nº encomenda", "Art./Serv.", "Descrição", "Und."]):
        def get_base_mould(val):
            if pd.isna(val) or str(val) == "Molde: N/D":
                return "N/D"
            m = re.search(r"Molde\s*:\s*([^,]+)", str(val), re.IGNORECASE)
            return m.group(1).strip() if m else str(val)

        group["Base_Molde"] = group["Molde e componente"].apply(get_base_mould)
        
        for base_m, sub_group in group.groupby("Base_Molde"):
            total_qty = sub_group["Qtd_Num"].sum()
            componentes = []
            for val in sub_group["Molde e componente"]:
                if pd.notna(val) and str(val).strip() and str(val) != "Molde: N/D":
                    m_comp = re.search(r"Compon\.?\s*:\s*(.+)", str(val), re.IGNORECASE)
                    if m_comp:
                        componentes.append(m_comp.group(1).strip())

            unique_comps = sorted(list(set(componentes)))
            first_row = sub_group.iloc[0]
            
            grouped_labels.append({
                "Fornecedor": first_row["Fornecedor"],
                "Nº encomenda": first_row["Nº encomenda"],
                "Art./Serv.": first_row["Art./Serv."],
                "Descrição": first_row["Descrição"],
                "Und.": first_row["Und."],
                "Qtd. Doc.": total_qty,
                "Molde": base_m,
                "Componentes": " / ".join(unique_comps) if unique_comps else ""
            })

    grouped_labels = pd.DataFrame(grouped_labels)

    label_w = (page_w - 2 * margin - gap_x * (columns_per_page - 1)) / columns_per_page
    label_h = (page_h - 2 * margin - gap_y * (rows_per_page - 1)) / rows_per_page
    c = canvas.Canvas(output, pagesize=A4)

    for idx, (_, row) in enumerate(grouped_labels.iterrows()):
        position = idx % (columns_per_page * rows_per_page)
        col_idx = position % columns_per_page
        row_idx = position // columns_per_page
        x = margin + col_idx * (label_w + gap_x)
        y = page_h - margin - (row_idx + 1) * label_h - row_idx * gap_y

        c.setStrokeColor(colors.HexColor("#777777"))
        c.setLineWidth(0.6)
        c.roundRect(x, y, label_w, label_h, 2 * mm, stroke=1, fill=0)

        supplier_name = str(row.get("Fornecedor", "") or "")
        order_no = str(row.get("Nº encomenda", "") or "")
        clean_qty = format_qty(row.get('Qtd. Doc.'))
        qtd_val = f"{clean_qty} {row.get('Und.', '')}".strip()
        molde_val = str(row.get("Molde", "") or "")
        comps_val = str(row.get("Componentes", "") or "")

        # Monta a linha de cabeçalho ex: "IGM — Encomenda: ENF A/101139"
        if supplier_name and supplier_name != "Fornecedor Desconhecido":
            header_text = f"{supplier_name} — Encomenda: {order_no}" if order_no else supplier_name
        else:
            header_text = f"Encomenda: {order_no}" if order_no else ""

        payload = {
            "Contribuinte Fornecedor": "PT-515690155"
        }
        if supplier_name:
            payload["Fornecedor"] = supplier_name
        if order_no:
            payload["Nº encomenda"] = order_no
        if pd.notna(row.get("Art./Serv.")):
            payload["Art./Serv."] = str(row.get("Art./Serv."))
        if pd.notna(row.get("Descrição")):
            payload["Descrição"] = str(row.get("Descrição"))
        if clean_qty:
            payload["Qtd. Doc."] = qtd_val
        if molde_val and molde_val != "N/D":
            m_payload = f"Molde: {molde_val}"
            if comps_val:
                m_payload += f", Compon.: {comps_val}"
            payload["Molde e componente"] = m_payload

        qr_size = min(qr_max_size, label_w * 0.48, label_h * 0.72)
        qr_x = x + label_w - qr_size - 1.5 * mm
        qr_y = y + label_h - qr_size - 1.5 * mm
        
        qr_img = qrcode.make(json.dumps(payload, ensure_ascii=False, separators=(",", ":")))
        qr_buffer = io.BytesIO()
        qr_img.save(qr_buffer, format="PNG")
        qr_buffer.seek(0)
        from reportlab.lib.utils import ImageReader
        c.drawImage(ImageReader(qr_buffer), qr_x, qr_y, width=qr_size, height=qr_size, preserveAspectRatio=True, mask="auto")

        text_x = x + 2.5 * mm
        text_y = y + label_h - 4.5 * mm
        max_text_width = qr_x - text_x - 1.5 * mm

        desc_val = str(row.get("Descrição", "") or "")

        # 1. Cabeçalho com Empresa e Encomenda (com corte automático se for comprido)
        if header_text:
            c.setFont("Helvetica-Bold", 8.5)
            avail_h = max(10, max_text_width)
            shown_h = header_text
            while shown_h and c.stringWidth(shown_h, "Helvetica-Bold", 8.5) > avail_h:
                shown_h = shown_h[:-1]
            c.drawString(text_x, text_y, shown_h)
            text_y -= 4.5 * mm

        # 2. Descrição (até 2 linhas)
        if desc_val:
            lbl = "Descrição:"
            c.setFont("Helvetica-Bold", 9)
            c.drawString(text_x, text_y, lbl)
            c.setFont("Helvetica", 9)
            p_w = c.stringWidth(lbl + " ", "Helvetica-Bold", 9)
            avail_l1 = max(10, max_text_width - p_w)
            avail_l2 = max(10, label_w - 5 * mm)

            words = desc_val.split()
            l1, l2 = "", ""
            for w in words:
                test = (l1 + " " + w).strip()
                if c.stringWidth(test, "Helvetica", 9) <= avail_l1:
                    l1 = test
                else:
                    test2 = (l2 + " " + w).strip()
                    if c.stringWidth(test2, "Helvetica", 9) <= avail_l2:
                        l2 = test2
                    else:
                        l2 += "..."
                        break

            c.drawString(text_x + p_w, text_y, l1)
            text_y -= 4 * mm
            if l2:
                c.drawString(text_x, text_y, l2)
                text_y -= 4 * mm
            else:
                text_y -= 1 * mm

        # 3. Quantidade
        if qtd_val:
            lbl = "Quantidade:"
            c.setFont("Helvetica-Bold", 9)
            c.drawString(text_x, text_y, lbl)
            c.setFont("Helvetica", 9)
            p_w = c.stringWidth(lbl + " ", "Helvetica-Bold", 9)
            c.drawString(text_x + p_w, text_y, qtd_val)
            text_y -= 4.5 * mm

        # 4. Molde (Linha dedicada)
        if molde_val and molde_val != "N/D":
            lbl_m = "Molde:"
            c.setFont("Helvetica-Bold", 9)
            c.drawString(text_x, text_y, lbl_m)
            c.setFont("Helvetica", 9)
            p_wm = c.stringWidth(lbl_m + " ", "Helvetica-Bold", 9)
            
            avail_m = max(10, max_text_width - p_wm)
            shown_m = molde_val
            while shown_m and c.stringWidth(shown_m, "Helvetica", 9) > avail_m:
                shown_m = shown_m[:-1]
            c.drawString(text_x + p_wm, text_y, shown_m)
            text_y -= 4.2 * mm

        # 5. Componentes (Linha seguinte dedicada, se existirem)
        if comps_val:
            lbl_c = "Comp.:"
            c.setFont("Helvetica-Bold", 9)
            c.drawString(text_x, text_y, lbl_c)
            c.setFont("Helvetica", 9)
            p_wc = c.stringWidth(lbl_c + " ", "Helvetica-Bold", 9)
            
            avail_c = max(10, max_text_width - p_wc)
            shown_c = comps_val
            while shown_c and c.stringWidth(shown_c, "Helvetica", 9) > avail_c:
                shown_c = shown_c[:-1]
            if shown_c != comps_val and len(shown_c) > 3:
                shown_c = shown_c[:-3] + "..."
            c.drawString(text_x + p_wc, text_y, shown_c)

        if position == columns_per_page * rows_per_page - 1 or idx == len(grouped_labels) - 1:
            c.showPage()
    c.save()
    return output.getvalue()


# --- INTERFACE PRINCIPAL ---
st.set_page_config(page_title="Conversor de Encomendas", page_icon="🏷️", layout="wide")
st.title("Conversor de notas de encomenda")
st.caption("PDF → Excel consolidado → etiquetas A4 com QR Code.")

st.sidebar.header("Configuração de Impressão")
layout_option = st.sidebar.radio(
    "Escolha o formato das etiquetas:",
    (
        "2x5 (10 por página)",
        "2x7 (14 por página)",
        "2x8 (16 por página)"
    ),
    index=2
)

uploads = st.file_uploader("Seleciona uma ou várias notas de encomenda em PDF", type=["pdf"], accept_multiple_files=True)

if "data" not in st.session_state:
    st.session_state.data = None

if uploads:
    all_rows, all_warnings = [], []
    for uploaded in uploads:
        rows, warnings = parse_pdf(uploaded.getvalue(), uploaded.name)
        all_rows.extend(rows)
        all_warnings.extend(warnings)

    if all_rows:
        df = pd.DataFrame(all_rows)
        df = df.sort_values(
            by=["Fornecedor", "Nº encomenda", "_Item PDF", "Art./Serv."],
            kind="stable",
            na_position="last",
        ).reset_index(drop=True)
        st.session_state.data = df
    else:
        st.session_state.data = pd.DataFrame(columns=COLUMNS + ["_Item PDF"])

    for warning in all_warnings:
        st.warning(warning)

df = st.session_state.data
if df is not None and len(df):
    st.subheader("1. Confirmar, Corrigir e Ordenar os Dados (Personalizado)")
    st.write(f"Artigos encontrados: **{len(df)}**")
    
    with st.expander("⚙️ Opções de Ordenação Avançada (Até 3 Níveis)", expanded=False):
        available_cols = [col for col in COLUMNS if col in df.columns]
        
        c1, c2 = st.columns(2)
        with c1:
            sort_1 = st.selectbox("1º Critério de Ordenação", options=["(Nenhum)"] + available_cols, index=2 if "Nº encomenda" in available_cols else 0)
            asc_1 = st.radio("Direção (1º)", options=["Ascendente (A-Z)", "Descendente (Z-A)"], horizontal=True, key="asc_1") == "Ascendente (A-Z)"
        with c2:
            sort_2 = st.selectbox("2º Critério de Ordenação", options=["(Nenhum)"] + available_cols, index=4 if "Molde e componente" in available_cols else 0)
            asc_2 = st.radio("Direção (2º)", options=["Ascendente (A-Z)", "Descendente (Z-A)"], horizontal=True, key="asc_2") == "Ascendente (A-Z)"

        c3, c4 = st.columns(2)
        with c3:
            sort_3 = st.selectbox("3º Critério de Ordenação", options=["(Nenhum)"] + available_cols, index=0)
            asc_3 = st.radio("Direção (3º)", options=["Ascendente (A-Z)", "Descendente (Z-A)"], horizontal=True, key="asc_3") == "Ascendente (A-Z)"

    sort_cols_list = []
    sort_asc_list = []
    
    if sort_1 != "(Nenhum)":
        sort_cols_list.append(sort_1)
        sort_asc_list.append(asc_1)
    if sort_2 != "(Nenhum)" and sort_2 not in sort_cols_list:
        sort_cols_list.append(sort_2)
        sort_asc_list.append(asc_2)
    if sort_3 != "(Nenhum)" and sort_3 not in sort_cols_list:
        sort_cols_list.append(sort_3)
        sort_asc_list.append(asc_3)

    if not sort_cols_list:
        sort_cols_list = ["Fornecedor", "Nº encomenda", "_Item PDF"]
        sort_asc_list = [True, True, True]

    df_sorted = df.sort_values(by=sort_cols_list, ascending=sort_asc_list, kind="stable").reset_index(drop=True)

    edited = st.data_editor(
        df_sorted,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "_Item PDF": st.column_config.NumberColumn("Item PDF", disabled=True),
            "Qtd. Doc.": st.column_config.TextColumn("Qtd. Doc."),
        },
        key="orders_editor",
    )
    st.session_state.data = edited

    st.subheader("2. Exportar")
    col_ex1, col_ex2, col_ex3 = st.columns(3)
    with col_ex1:
        excel_bytes = make_excel(edited)
        st.download_button(
            "Excel Consolidado",
            data=excel_bytes,
            file_name="encomendas_consolidadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_ex2:
        pivot_bytes = make_pivot_excel(edited)
        st.download_button(
            "Tabela Dinâmica (Resumo)",
            data=pivot_bytes,
            file_name="tabela_dinamica_artigos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_ex3:
        labels_bytes = make_labels_pdf(edited, layout_option)
        st.download_button(
            "Etiquetas A4 (PDF)",
            data=labels_bytes,
            file_name="etiquetas_encomendas.pdf",
            mime="application/pdf",
            use_container_width=True,
        )