import io
import re
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st

# Configuração da página no Streamlit
st.set_page_config(
    page_title="Conversor de Encomendas para Etiquetas",
    page_icon="🏷️",
    layout="wide"
)

# Expressões Regulares globais
ORDER_RE = re.compile(r"ENF\s*([A-Z0-9/\-_]+)", re.IGNORECASE)
DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")
ITEM_RE = re.compile(
    r"^\s*(\d+)\s+([A-Z0-9._\-]+)\s+(.*?)\s+(\d+(?:[.,]\d+)?)\s*([A-Za-z]+)\s+(\d{4}-\d{2}-\d{2})\s*$"
)


def normalize_space(text: str) -> str:
    """Remove espaços duplos e limpa quebras de linha."""
    return re.sub(r"\s+", " ", text or "").strip()


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

    def clean_final_desc(desc: str) -> str:
        """Limpa moradas e ruídos residuais no fim da descrição."""
        desc = re.split(
            r"\b(Ouriçosa|3720-592|3720-502|Santiago da Riba|Oliveira de Azeméis)\b",
            desc,
            flags=re.IGNORECASE,
        )[0]
        desc = re.sub(r"\bUL\b\s*$", "", desc, flags=re.IGNORECASE)
        return normalize_space(desc)

    # Processamento por página para identificar o último item de cada página
    for page_idx, page_text in enumerate(pages_text):
        page_items = []
        current_item = None

        for raw_line in page_text.splitlines():
            line = normalize_space(raw_line)
            if not line:
                continue

            item_match = ITEM_RE.match(raw_line)
            if item_match:
                if current_item:
                    current_item["Descrição"] = clean_final_desc(current_item["Descrição"])
                    page_items.append(current_item)

                item_no, code, description, qty, unit, due_date = item_match.groups()
                current_item = {
                    "Fornecedor": supplier,
                    "Nº encomenda": order_number,
                    "Data documento": document_date,