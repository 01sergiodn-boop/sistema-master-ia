import streamlit as st
import json
import io
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image
from rembg import remove

st.set_page_config(page_title="Sistema Master | IA para PowerPoint", page_icon="✨", layout="centered")

st.title("✨ SISTEMA MASTER")
st.subheader("Engenharia Reversa de Imagens para PowerPoint Editável")

with st.sidebar:
    st.header("⚙️ Configurações")
    chave_api = st.text_input("Sua Chave API do Google Gemini:", type="password")
    st.markdown("---")
    st.markdown("**Como funciona?**\n1. Envie uma imagem.\n2. A IA identifica e recorta.\n3. Baixe o PPTX 100% editável.")

def configurar_ia(chave):
    genai.configure(api_key=chave)
    system_instruction = """
    Você é a IA do SISTEMA MASTER.
    Analise a imagem e devolva a estrutura exata.
    REGRA CRÍTICA: Responda APENAS com um arquivo JSON válido (lista de dicionários).
    
    Chaves do JSON:
    - "tipo": "texto", "forma" ou "imagem"
    - "conteudo": Texto (se for texto)
    - "cor_hex": Cor em Hexadecimal (ex: "FFFFFF")
    - "tamanho_fonte": Estimativa em pontos (ex: 24)
    - "x_percent": Posição X (0.0 a 1.0)
    - "y_percent": Posição Y (0.0 a 1.0)
    - "largura_percent": Largura (0.0 a 1.0)
    - "altura_percent": Altura (0.0 a 1.0)
    """
    return genai.GenerativeModel('gemini-1.5-pro', system_instruction=system_instruction)

def hex_para_rgb(hex_color):
    hex_color = hex_color.replace('#', '')
    if len(hex_color) != 6: return RGBColor(0, 0, 0)
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

def gerar_powerpoint_em_memoria(dados_json, imagem_original, largura_in, altura_in):
    prs = Presentation()
    prs.slide_width = Inches(largura_in)
    prs.slide_height = Inches(altura_in)
    slide = prs.slides.add_slide(prs.slide_layouts[6]) 
    
    largura_px, altura_px = imagem_original.size

    for elemento in dados_json:
        left = Inches(elemento.get("x_percent", 0) * largura_in)
        top = Inches(elemento.get("y_percent", 0) * altura_in)
        width = Inches(elemento.get("largura_percent", 0.1) * largura_in)
        height = Inches(elemento.get("altura_percent", 0.1) * altura_in)
        
        if elemento["tipo"] == "texto":
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.text = elemento.get("conteudo", "")
            p = tf.paragraphs[0]
            p.font.size = Pt(elemento.get("tamanho_fonte", 14))
            p.font.color.rgb = hex_para_rgb(elemento.get("cor_hex", "000000"))
            
        elif elemento["tipo"] == "forma":
            shape = slide.shapes.add_shape(1, left, top, width, height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = hex_para_rgb(elemento.get("cor_hex", "CCCCCC"))
            shape.line.fill.background()
            
        elif elemento["tipo"] == "imagem":
            l_px = int(elemento.get("x_percent", 0) * largura_px)
            t_px = int(elemento.get("y_percent", 0) * altura_px)
            r_px = l_px + int(elemento.get("largura_percent", 0.1) * largura_px)
            b_px = t_px + int(elemento.get("altura_percent", 0.1) * altura_px)
            
            recorte = imagem_original.crop((l_px, t_px, r_px, b_px))
            recorte_limpo = remove(recorte)
            
            img_bytes = io.BytesIO()
            recorte_limpo.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            
            slide.shapes.add_picture(img_bytes, left, top, width, height)

    pptx_bytes = io.BytesIO()
    prs.save(pptx_bytes)
    pptx_bytes.seek(0)
    return pptx_bytes

arquivo_imagem = st.file_uploader("Arraste uma imagem aqui (PNG ou JPG)", type=["png", "jpg", "jpeg"])

if arquivo_imagem:
    st.image(arquivo_imagem, caption="Imagem Original", use_column_width=True)
    st.markdown("---")
    st.subheader("📐 Formato do Slide")
    
    formato_escolhido = st.radio(
        "Escolha o tamanho do PowerPoint de saída:",
        options=["Widescreen (16:9)", "Quadrado (1:1)"],
        horizontal=True
    )
    
    if "Widescreen" in formato_escolhido:
        largura_slide, altura_slide = 10.0, 5.625
    else:
        largura_slide, altura_slide = 10.0, 10.0
    
    if st.button("🚀 Gerar PowerPoint Editável", type="primary"):
        if not chave_api:
            st.error("⚠️ Insira sua Chave da API do Gemini na barra lateral.")
        else:
            try:
                imagem_pil = Image.open(arquivo_imagem)
                with st.spinner('Analisando layout e tipografia...'):
                    modelo = configurar_ia(chave_api)
                    resposta = modelo.generate_content([imagem_pil, "Gere o JSON."])
                    texto_limpo = resposta.text.replace('```json', '').replace('
```', '').strip()
                    dados_json = json.loads(texto_limpo)
                
                with st.spinner('Construindo o arquivo .PPTX...'):
                    arquivo_pptx_pronto = gerar_powerpoint_em_memoria(dados_json, imagem_pil, largura_slide, altura_slide)
                
                st.success("✅ Arquivo gerado com sucesso!")
                st.download_button(
                    label="⬇️ Baixar Apresentação (.pptx)",
                    data=arquivo_pptx_pronto,
                    file_name="reconstrucao_master.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")
