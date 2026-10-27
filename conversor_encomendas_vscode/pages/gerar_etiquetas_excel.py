from __future__ import annotations
import io
import json
import pandas as pd
import qrcode
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

st.set_page_config(page_title="Gerador de Etiquetas por Excel", page_icon="🏷️", layout="wide")
st.title("Gerador de Etiquetas via Excel com QR Code Dinâmico")
st.write("Carrega o ficheiro Excel consolidado e escolhe abaixo quais os campos que pretendes incluir no QR Code.")

uploaded_excel = st.file_uploader("1. Seleciona o ficheiro Excel (.xlsx)", type=["xlsx"])

df = None
if uploaded_excel is not None:
    try:
        df = pd.read_excel(uploaded_excel)
        if "Cliente" in df.columns and "Fornecedor" not in df.columns:
            df = df.rename(columns={"Cliente": "Fornecedor"})
        st.success(f"Excel carregado com sucesso! {len(df)} artigos encontrados.")
    except Exception as e:
        st.error(f"Erro ao ler o ficheiro Excel: {e}")

if df is not None and not df.empty:
    st.subheader("2. Seleciona os campos a incluir no QR Code")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        show_fornecedor = st.checkbox("Fornecedor", value=True)
        show_encomenda = st.checkbox("Nº Encomenda", value=True)
    with col2:
        show_artigo = st.checkbox("Código do Artigo", value=True)
        show_descricao = st.checkbox("Descrição", value=True)
    with col3:
        show_qtd = st.checkbox("Quantidade e Unidade", value=True)
        show_molde = st.checkbox("Molde e Componente", value=True)
    with col4:
        show_entrega = st.checkbox("Data de Entrega", value=True)
        show_origem = st.checkbox("PDF de Origem", value=False)

    st.markdown("---")
    st.subheader("3. Pré-visualização e Exportação")
    st.dataframe(df.head(5), use_container_width=True)

    print_options = {
        "fornecedor": show_fornecedor,
        "encomenda": show_encomenda,
        "artigo": show_artigo,
        "descricao": show_descricao,
        "qtd": show_qtd,
        "molde": show_molde,
        "entrega": show_entrega,
        "origem": show_origem
    }

    def make_labels_pdf(df: pd.DataFrame, options: dict) -> bytes:
        output = io.BytesIO()
        page_w, page_h = A4
        margin = 8 * mm
        gap_x = 4 * mm
        gap_y = 4 * mm
        columns_per_page, rows_per_page = 2, 5
        
        label_w = (page_w - 2 * margin - gap_x * (columns_per_page - 1)) / columns_per_page
        label_h = (page_h - 2 * margin - gap_y * (rows_per_page - 1)) / rows_per_page
        c = canvas.Canvas(output, pagesize=A4)

        for idx, (_, row) in enumerate(df.iterrows()):
            position = idx % (columns_per_page * rows_per_page)
            col_idx = position % columns_per_page
            row_idx = position // columns_per_page
            x = margin + col_idx * (label_w + gap_x)
            y = page_h - margin - (row_idx + 1) * label_h - row_idx * gap_y

            c.setStrokeColor(colors.HexColor("#777777"))
            c.setLineWidth(0.6)
            c.roundRect(x, y, label_w, label_h, 2 * mm, stroke=1, fill=0)

            order_no = str(row.get("Nº encomenda", "") or "")
            forn_val = str(row.get("Fornecedor", row.get("Cliente", "")) or "")

            payload = {}
            if options["fornecedor"] and forn_val: 
                payload["Fornecedor"] = forn_val
            if options["encomenda"] and order_no: 
                payload["Nº encomenda"] = order_no
            if options["artigo"] and pd.notna(row.get("Art./Serv.")): 
                payload["Art./Serv."] = str(row.get("Art./Serv."))
            if options["descricao"] and pd.notna(row.get("Descrição")): 
                payload["Descrição"] = str(row.get("Descrição"))
            if options["qtd"] and pd.notna(row.get("Qtd. Doc.")): 
                payload["Qtd. Doc."] = f"{row.get('Qtd. Doc.')} {row.get('Und.', '')}".strip()
            if options["molde"] and pd.notna(row.get("Molde e componente")): 
                payload["Molde e componente"] = str(row.get("Molde e componente"))
            if options["entrega"] and pd.notna(row.get("Dt. Prev. Entrega")): 
                payload["Dt. Prev. Entrega"] = str(row.get("Dt. Prev. Entrega"))
            if options["origem"] and pd.notna(row.get("PDF origem")): 
                payload["PDF origem"] = str(row.get("PDF origem"))

            # QR Code posicionado no canto superior direito
            qr_size = min(24 * mm, label_w * 0.32, label_h * 0.42)
            qr_x = x + label_w - qr_size - 3 * mm
            qr_y = y + label_h - qr_size - 4 * mm
            
            qr_img = qrcode.make(json.dumps(payload, ensure_ascii=False, separators=(",", ":")))
            qr_buffer = io.BytesIO()
            qr_img.save(qr_buffer, format="PNG")
            qr_buffer.seek(0)
            from reportlab.lib.utils import ImageReader
            c.drawImage(ImageReader(qr_buffer), qr_x, qr_y, width=qr_size, height=qr_size, preserveAspectRatio=True, mask="auto")

            text_x = x + 3 * mm
            text_y = y + label_h - 6 * mm
            max_text_width = qr_x - text_x - 2 * mm

            desc_val = str(row.get("Descrição", "") or "")
            qtd_val = f"{row.get('Qtd. Doc.', '')} {row.get('Und.', '')}".strip()
            molde_val = str(row.get("Molde e componente", "") or "")

            if order_no:
                lbl = "Encomenda:"
                c.setFont("Helvetica-Bold", 7)
                c.drawString(text_x, text_y, lbl)
                c.setFont("Helvetica-Bold", 7)
                p_w = c.stringWidth(lbl + " ", "Helvetica-Bold", 7)
                c.drawString(text_x + p_w, text_y, order_no)
                text_y -= 5 * mm

            if desc_val:
                lbl = "Descrição:"
                c.setFont("Helvetica-Bold", 6.5)
                c.drawString(text_x, text_y, lbl)
                c.setFont("Helvetica", 6.5)
                p_w = c.stringWidth(lbl + " ", "Helvetica-Bold", 6.5)
                avail_l1 = max(10, max_text_width - p_w)
                avail_l2 = max(10, label_w - 6 * mm)

                words = desc_val.split()
                l1, l2 = "", ""
                for w in words:
                    test = (l1 + " " + w).strip()
                    if c.stringWidth(test, "Helvetica", 6.5) <= avail_l1:
                        l1 = test
                    else:
                        test2 = (l2 + " " + w).strip()
                        if c.stringWidth(test2, "Helvetica", 6.5) <= avail_l2:
                            l2 = test2
                        else:
                            l2 += "..."
                            break

                c.drawString(text_x + p_w, text_y, l1)
                text_y -= 3.5 * mm
                if l2:
                    c.drawString(text_x, text_y, l2)
                    text_y -= 4.5 * mm
                else:
                    text_y -= 1 * mm

            if qtd_val:
                lbl = "Quantidade:"
                c.setFont("Helvetica-Bold", 6.5)
                c.drawString(text_x, text_y, lbl)
                c.setFont("Helvetica", 6.5)
                p_w = c.stringWidth(lbl + " ", "Helvetica-Bold", 6.5)
                c.drawString(text_x + p_w, text_y, qtd_val)
                text_y -= 5 * mm

            if molde_val and molde_val != "Molde: N/D":
                lbl = "Molde/Comp:"
                c.setFont("Helvetica-Bold", 6.5)
                c.drawString(text_x, text_y, lbl)
                c.setFont("Helvetica", 6.5)
                p_w = c.stringWidth(lbl + " ", "Helvetica-Bold", 6.5)
                avail = max(10, label_w - 6 * mm - p_w)
                shown = molde_val
                while shown and c.stringWidth(shown, "Helvetica", 6.5) > avail:
                    shown = shown[:-1]
                if shown != molde_val and len(shown) > 3:
                    shown = shown[:-3] + "..."
                c.drawString(text_x + p_w, text_y, shown)

            if position == columns_per_page * rows_per_page - 1 or idx == len(df) - 1:
                c.showPage()
                
        c.save()
        return output.getvalue()

    if st.button("Gerar e Descarregar PDF de Etiquetas A4", type="primary"):
        pdf_bytes = make_labels_pdf(df, print_options)
        st.download_button(
            "Descarregar PDF de Etiquetas",
            data=pdf_bytes,
            file_name="etiquetas_excel_personalizadas.pdf",
            mime="application/pdf"
        )