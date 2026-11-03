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

                # 1. Procura por "Molde: ..., Compon.: ..."

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

                    # 2. Procura por "Estação Bancada: ..."

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