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


print("COLUNAS DISPONÍVEIS:", df.columns.tolist())

def make_pivot_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    
    temp_df = df.copy()
    
    # 1. Normalizar nomes das colunas para evitar conflitos de espaços
    temp_df.columns = [c.strip() for c in temp_df.columns]
    
    # Garantir que a quantidade é numérica
    qty_col = "Qtd. Doc." if "Qtd. Doc." in temp_df.columns else temp_df.columns[5]
    temp_df["Qtd_Num"] = pd.to_numeric(
        temp_df[qty_col].astype(str).str.replace(",", "."), errors="coerce"
    ).fillna(0)
    
    # 2. Identificar dinamicamente a coluna de data prevista de entrega
    date_col = next((c for c in temp_df.columns if "prev" in c.lower() or "entrega" in c.lower() or "data" in c.lower()), None)
    
    if date_col:
        temp_df["Data_Entrega_Clean"] = temp_df[date_col].astype(str).str.strip()
    else:
        temp_df["Data_Entrega_Clean"] = ""
    
    # 3. Agrupamento para a Tabela Dinâmica / Resumo
    pivot_df = temp_df.groupby(
        ["Art./Serv.", "Descrição", "Und."], 
        as_index=False
    ).agg({
        "Qtd_Num": "sum",
        "Data_Entrega_Clean": lambda x: sorted([d for d in x if d and d.lower() != "nan"])[0] if any(d for d in x if d and d.lower() != "nan") else ""
    })
    
    pivot_df = pivot_df.rename(columns={
        "Qtd_Num": "Quantidade Total",
        "Data_Entrega_Clean": "Dt. Prev. Entrega"
    })
    
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
        pivot_df["Correspondente_Fornecedor"] = pivot_df["Correspondente_Excel"].fillna(pivot_df["Descrição"])
        
        pivot_df["P.U."] = pd.to_numeric(pivot_df["P.U."], errors="coerce").fillna(0.0)
        pivot_df = pivot_df.drop(columns=["Desc_Key", "Descricao_Key", "Correspondente_Excel"], errors="ignore")
    else:
        pivot_df["Correspondente_Fornecedor"] = pivot_df["Descrição"]
        pivot_df["P.U."] = 0.0

    pivot_df["Preço Total (€)"] = pivot_df["Quantidade Total"] * pivot_df["P.U."]
    
    # Mapeamento final para alinhar na Coluna H
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
        "Preço Total (€)",
        "Dt. Prev. Entrega"  # Fica garantidamente na Coluna H
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