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
        if st.session_state["password"] == "moldes2026":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 Acesso Restrito")
        st.text_input("Introduz a palavra-passe para aceder à aplicação:", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔒 Acesso Restrito")
        st.text_input("Introduz a palavra-passe para aceder à aplicação:", type="password", on_change=password_entered, key="password")
        st.error("😕 Palavra-passe incorreta. Tenta novamente.")
        return False
    else:
        return True

if not check_password():
    st.stop()

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

    supplier = "Empresa"
    clean_filename = filename.upper()
    
    enf_match = re.search(r"ENF[-_]?([A-Z]+)[-_]?\d+", clean_filename)
    if enf_match:
        supplier = enf_match.group(1).strip()
    else:
        parts = re.split(r"[-_\s]+", Path(filename).stem)
        for part in parts:
            if part.isalpha() and len(part) <= 5:
                supplier = part.upper()
                break

    def clean_description(desc: str) -> str:
        """Corta de vez qualquer texto de morada que venha agarrado."""
        if not desc:
            return ""
        patterns = [
            r"Ouriçosa.*",
            r"3720-\d{3}.*",
            r"Santiago\s+da\s+Riba.*",
            r"Oliveira\s+de\s+Azeméis.*",
            r"\bUL\b\s*$"
        ]
        cleaned = desc
        for pat in patterns:
            cleaned = re.sub(pat, "", cleaned, flags=re.IGNORECASE)
        return normalize_space(cleaned)

    for page_idx, page_text in enumerate(pages_text):
        page_items = []
        current_item = None
        
        for raw_line in page_text.splitlines():
            line = normalize_space(raw_line)
            if not line:
                continue

            line_low = line.lower()
            if any(term in line_low for term in [
                "ouriçosa", "3720-592", "3720-502", "ulmolde", 
                "capital social", "conservatória", "geral@ulmolde"
            ]):
                continue

            item_match = ITEM_RE.match(raw_line)
            if item_match:
                if current_item:
                    current_item["Descrição"] = clean_description(current_item["Descrição"])
                    page_items.append(current_item)

                item_no, code, description, qty, unit, due_date = item_match.groups()
                current_item = {
                    "Fornecedor": supplier,
                    "Nº encomenda": order_number,
                    "Data documento": document_date,
                    "Art./Serv.": code.strip(),
                    "Descrição": clean_description(description),
                    "Qtd. Doc.": qty.replace(",", "."),
                    "Und.": unit,
                    "Molde e componente": "",
                    "Dt. Prev. Entrega": due_date,
                    "PDF origem": filename,
                    "_Item PDF": item_no,
                    "Pág.": page_idx + 1,
                    "⚠️ Verificar": "OK"
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
                else:
                    station_match = re.search(r"Estação\s+Bancada\s*:\s*(.+)", line, re.IGNORECASE)
                    if station_match:
                        station_val = station_match.group(1).strip()
                        existing_mould = current_item.get("Molde e componente", "")
                        if existing_mould:
                            current_item["Molde e componente"] = f"{existing_mould} | Estação: {station_val}"
                        else:
                            current_item["Molde e componente"] = f"Estação: {station_val}"
                    elif (
                        not re.match(r"^\s*\d+\s+\d{6,}\s+", raw_line)
                        and ":" not in line
                        and len(line) <= 45
                    ):
                        current_item["Descrição"] = clean_description(
                            current_item["Descrição"] + " " + line
                        )

        if current_item:
            current_item["Descrição"] = clean_description(current_item["Descrição"])
            page_items.append(current_item)

        if page_items:
            page_items[-1]["⚠️ Verificar"] = "⚠️ VERIFICAR ÚLTIMO"
            rows.extend(page_items)

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
    
    # Garantir que a quantidade é numérica
    temp_df["Qtd_Num"] = pd.to_numeric(
        temp_df["Qtd. Doc."].astype(str).str.replace(",", "."), errors="coerce"
    ).fillna(0)
    
    # Garantir que a coluna Dt. Prev. Entrega existe e não tem valores estranhos
    if "Dt. Prev. Entrega" not in temp_df.columns:
        temp_df["Dt. Prev. Entrega"] = ""
    else:
        temp_df["Dt. Prev. Entrega"] = temp_df["Dt. Prev. Entrega"].astype(str).str.strip()
    
    # Agrupamento para a Tabela Dinâmica / Resumo
    # Agrupa por Artigo, Descrição e Unidade, calculando a Soma da Quantidade e a Menor Data (mais antiga)
    pivot_df = temp_df.groupby(
        ["Art./Serv.", "Descrição", "Und."], 
        as_index=False
    ).agg({
        "Qtd_Num": "sum",
        "Dt. Prev. Entrega": lambda x: sorted([d for d in x if d and d.lower() != "nan"])[0] if any(d for d in x if d and d.lower() != "nan") else ""
    })
    
    pivot_df = pivot_df.rename(columns={"Qtd_Num": "Quantidade Total"})
    
    # Cruzamento com a tabela de equivalências
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
    
    # Mapeamento e renomeação das colunas finais
    pivot_df = pivot_df.rename(columns={
        "Art./Serv.": "Código Artigo",
        "Descrição": "Descrição (PDF)",
        "Correspondente_Fornecedor": "Correspondente (Fornecedor)",
        "P.U.": "P.U. (€)",
        "Dt. Prev. Entrega": "Dt. Prev. Entrega"
    })

    # Ordem exata das colunas: A Dt. Prev. Entrega vai para a Coluna H
    cols_order = [
        "Código Artigo",
        "Descrição (PDF)",
        "Correspondente (Fornecedor)",
        "Quantidade Total",
        "Und.",
        "P.U. (€)",
        "Preço Total (€)",
        "Dt. Prev. Entrega"  # Coluna H
    ]
    
    # Reordenar mantendo as colunas existentes
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
# --- INTERFACE PRINCIPAL DO STREAMLIT ---
st.title("🏷️ Conversor de Notas de Encomenda")
st.write("PDF -> Excel consolidado -> Tabela Dinâmica (Resumo)")

uploaded_files = st.file_uploader(
    "Seleciona uma ou várias notas de encomenda em PDF",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    all_rows = []
    all_warnings = []

    for pdf_file in uploaded_files:
        file_bytes = pdf_file.read()
        rows, warnings = parse_pdf(file_bytes, pdf_file.name)
        all_rows.extend(rows)
        all_warnings.extend(warnings)

    if all_warnings:
        for w in all_warnings:
            st.warning(w)

    if all_rows:
        df = pd.DataFrame(all_rows)

        st.subheader("1. Confirmar, Corrigir e Ordenar os Dados")
        st.write(f"Artigos encontrados: {len(df)}")

        # Tabela editável diretamente na interface
        edited_df = st.data_editor(df, use_container_width=True, height=400)

        st.subheader("2. Exportar")
        col1, col2 = st.columns(2)

        with col1:
            excel_bytes = make_excel(edited_df)
            st.download_button(
                label="📥 Descarregar Excel Consolidado",
                data=excel_bytes,
                file_name="encomendas_consolidadas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with col2:
            pivot_excel_bytes = make_pivot_excel(edited_df)
            st.download_button(
                label="📊 Descarregar Tabela Dinâmica (Resumo)",
                data=pivot_excel_bytes,
                file_name="resumo_tabela_dinamica.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )