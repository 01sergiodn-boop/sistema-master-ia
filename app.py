import streamlit as st
from PIL import Image

from core.preprocessing import preprocessar_imagem
from core.ocr import executar_ocr
from core.layout import detectar_layout
from core.segmentation import segmentar_elementos
from core.gemini_engine import analisar_com_gemini
from core.validation import validar_json
from core.corrections import corrigir_desvios
from core.powerpoint import gerar_pptx

st.set_page_config(
    page_title="Sistema Master Enterprise",
    layout="wide"
)

st.title("🚀 Sistema Master Enterprise")

arquivo = st.file_uploader(
    "Envie uma imagem",
    type=["png", "jpg", "jpeg"]
)

if arquivo:

    imagem = Image.open(arquivo).convert("RGB")

    st.image(imagem, use_container_width=True)

    if st.button("🚀 Gerar PPTX"):

        with st.status("Processando", expanded=True):

            st.write("1️⃣ Pré-processamento")
            imagem_processada = preprocessar_imagem(imagem)

            st.write("2️⃣ OCR")
            dados_ocr = executar_ocr(imagem_processada)

            st.write("3️⃣ Layout")
            layout = detectar_layout(imagem_processada)

            st.write("4️⃣ Segmentação")
            segmentos = segmentar_elementos(imagem_processada)

            st.write("5️⃣ IA")
            dados_json = analisar_com_gemini(
                imagem_processada,
                dados_ocr,
                layout,
                segmentos
            )

            st.write("6️⃣ Validação")
            dados_json = validar_json(dados_json)

            st.write("7️⃣ Correção de desvios")
            dados_json = corrigir_desvios(dados_json)

            st.write("8️⃣ Gerando PPTX")
            pptx = gerar_pptx(
                dados_json,
                imagem_processada
            )

        st.download_button(
            label="⬇️ Download PPTX",
            data=pptx,
            file_name="master_enterprise.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )