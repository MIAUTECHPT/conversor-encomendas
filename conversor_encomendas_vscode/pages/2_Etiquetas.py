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

st.set_page_config(page_title="Gerador de Etiquetas", page_icon="🏷️", layout="wide")
st.title("Gerador de Etiquetas A4 com QR Code")

# --- Configuração da Barra Lateral ---
st.sidebar.header("Configuração de Impressão")
st.sidebar.write("Seleciona os campos para a etiqueta:")

show_artigo = st.sidebar.checkbox("Código do Artigo", value=True)
show_descricao = st.sidebar.checkbox("Descrição", value=True)
show_qtd = st.sidebar.checkbox("Quantidade e Unidade", value=True)
show_molde = st.sidebar.checkbox("Molde e Componente", value=True)
show_entrega = st.sidebar.checkbox("Data de Entrega", value=True)
show_origem = st.sidebar.checkbox("PDF de Origem", value=False)

# Verifica se existem dados da página anterior ou permite carregar um ficheiro
df = st.session_state.get("data", None)

if df is None or df.empty:
    uploaded_file = st.file_uploader("Ou carrega o teu CSV/Excel de artigos", type=["csv", "xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith(".xlsx") else pd.read_csv(uploaded_file)
        st.session_state.data = df

def make_labels_pdf(df: pd.DataFrame, options: dict) -> bytes:
    output = io.BytesIO()
    page_w, page_h = A4
    margin = 8 * mm
    gap_x = 4 * mm
    gap_y = 4 * mm
    columns_per_page, rows_per_page = 2, 4
    label_w = (page_w - 2 * margin - gap_x * (columns_per_page - 1)) / columns_per_page
    label_h = (page_h - 2 * margin - gap_y * (rows_per_page - 1)) / rows_per_page
    c = canvas.Canvas(output, pagesize=A4)

    for idx, (_, row) in enumerate(df.iterrows()):
        position = idx % (columns_per_page * rows_per_page)
        col_idx = position % columns_per_page
        row_idx = position // columns_per_page
        x = margin + col_idx * (label_w + gap_x)
        y = page_h - margin - (row_idx + 1) * label_h - row_idx * gap_y

        # Desenhar retângulo da etiqueta
        c.setStrokeColor(colors.HexColor("#777777"))
        c.setLineWidth(0.6)
        c.roundRect(x, y, label_w, label_h, 2 * mm, stroke=1, fill=0)

        order_no = str(row.get("Nº encomenda", "") or "")
        client = str(row.get("Cliente", "") or "")
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x + 3 * mm, y + label_h - 6 * mm, "ETIQUETA DE ARTIGO")
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x + 3 * mm, y + label_h - 11 * mm, f"Encomenda: {order_no[:30]}")
        c.setFont("Helvetica", 7)
        c.drawString(x + 3 * mm, y + label_h - 15 * mm, f"Cliente: {client[:35]}")

        # Construção dinâmica do Payload do QR Code
        payload = {}
        if client: payload["Cliente"] = client
        if order_no: payload["Nº encomenda"] = order_no
        if options["artigo"] and row.get("Art./Serv."): payload["Art./Serv."] = str(row.get("Art./Serv."))
        if options["descricao"] and row.get("Descrição"): payload["Descrição"] = str(row.get("Descrição"))
        if options["qtd"] and row.get("Qtd. Doc."): payload["Qtd. Doc."] = f"{row.get('Qtd. Doc.')} {row.get('Und.', '')}".strip()
        if options["molde"] and row.get("Molde e componente"): payload["Molde e componente"] = str(row.get("Molde e componente"))
        if options["entrega"] and row.get("Dt. Prev. Entrega"): payload["Dt. Prev. Entrega"] = str(row.get("Dt. Prev. Entrega"))
        if options["origem"] and row.get("PDF origem"): payload["PDF origem"] = str(row.get("PDF origem"))

        # Gerar e desenhar QR Code
        qr_size = min(30 * mm, label_w * 0.32, label_h * 0.46)
        qr_x = x + label_w - qr_size - 3 * mm
        qr_y = y + 4 * mm
        
        qr_img = qrcode.make(json.dumps(payload, ensure_ascii=False, separators=(",", ":")))
        qr_buffer = io.BytesIO()
        qr_img.save(qr_buffer, format="PNG")
        qr_buffer.seek(0)
        from reportlab.lib.utils import ImageReader
        c.drawImage(ImageReader(qr_buffer), qr_x, qr_y, width=qr_size, height=qr_size, preserveAspectRatio=True, mask="auto")

        # Escrever linhas de texto na etiqueta
        text_x = x + 3 * mm
        text_y = y + label_h - 21 * mm
        max_text_width = qr_x - text_x - 2 * mm

        text_lines = []
        if options["artigo"] and row.get("Art./Serv."):
            text_lines.append(("Artigo:", str(row.get("Art./Serv."))))
        if options["descricao"] and row.get("Descrição"):
            text_lines.append(("Descrição:", str(row.get("Descrição"))))
        if options["qtd"] and row.get("Qtd. Doc."):
            q_val = f"{row.get('Qtd. Doc.')} {row.get('Und.', '')}".strip()
            text_lines.append(("Quantidade:", q_val))
        if options["molde"] and row.get("Molde e componente") and str(row.get("Molde e componente")) != "Molde: N/D":
            text_lines.append(("Molde/Comp:", str(row.get("Molde e componente"))))
        if options["entrega"] and row.get("Dt. Prev. Entrega"):
            text_lines.append(("Entrega:", str(row.get("Dt. Prev. Entrega"))))
        if options["origem"] and row.get("PDF origem"):
            text_lines.append(("Origem:", str(row.get("PDF origem"))))

        for label, value in text_lines:
            c.setFont("Helvetica-Bold", 6.7)
            c.drawString(text_x, text_y, label)
            c.setFont("Helvetica", 6.7)
            prefix_width = c.stringWidth(label + " ", "Helvetica-Bold", 6.7)
            available = max(10, max_text_width - prefix_width)
            shown = value
            while shown and c.stringWidth(shown, "Helvetica", 6.7) > available:
                shown = shown[:-1]
            if shown != value and len(shown) > 3:
                shown = shown[:-3] + "..."
            c.drawString(text_x + prefix_width, text_y, shown)
            text_y -= 4.4 * mm

        if position == columns_per_page * rows_per_page - 1 or idx == len(df) - 1:
            c.showPage()
    c.save()
    return output.getvalue()

if df is not None and not df.empty:
    st.success(f"{len(df)} artigos prontos para gerar etiquetas.")
    
    print_options = {
        "artigo": show_artigo,
        "descricao": show_descricao,
        "qtd": show_qtd,
        "molde": show_molde,
        "entrega": show_entrega,
        "origem": show_origem
    }
    
    if st.button("Gerar e Descarregar PDF de Etiquetas A4", type="primary"):
        pdf_bytes = make_labels_pdf(df, print_options)
        st.download_button("Descarregar PDF de Etiquetas", data=pdf_bytes, file_name="etiquetas_personalizadas.pdf", mime="application/pdf")
else:
    st.info("Carrega um PDF na primeira página para começarmos a gerar as etiquetas.")