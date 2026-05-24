import streamlit as st
import json
import io
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image

# SUA CHAVE AQUI DIRETO NO CÓDIGO
CHAVE_API = "AIzaSyDkz_EVM8rzs1012mXdHS1J0GH53Iyu3GY"

st.set_page_config(page_title="Sistema Master | IA")

st.title("✨ Sistema Master")
st.write("Transforme imagens em apresentações editáveis.")

def configurar_ia(chave):
    genai.configure(api_key=chave)
    instrucao = (
        'Analise a imagem. Responda APENAS com um JSON '
        'com as chaves: "tipo", "conteudo", "cor_hex", '
        '"tamanho_fonte", "x_percent", "y_percent", '
        '"largura_percent", "altura_percent".'
    )
    return genai.GenerativeModel(
        'gemini-1.5-flash', 
        system_instruction=instrucao
    )

def hex_para_rgb(hex_color):
    hex_color = hex_color.replace('#', '')
    if len(hex_color) != 6: 
        return RGBColor(0, 0, 0)
    
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return RGBColor(r, g, b)

def gerar_powerpoint(dados_json, imagem, larg_in, alt_in):
    prs = Presentation()
    prs.slide_width = Inches(larg_in)
    prs.slide_height = Inches(alt_in)
    slide = prs.slides.add_slide(prs.slide_layouts[6]) 
    larg_px, alt_px = imagem.size

    for el in dados_json:
        left = Inches(el.get("x_percent", 0) * larg_in)
        top = Inches(el.get("y_percent", 0) * alt_in)
        width = Inches(el.get("largura_percent", 0.1) * larg_in)
        height = Inches(el.get("altura_percent", 0.1) * alt_in)
        tipo = el.get("tipo", "")
        
        if tipo == "texto":
            txBox = slide.shapes.add_textbox(
                left, top, width, height
            )
            p = txBox.text_frame.paragraphs[0]
            p.text = el.get("conteudo", "")
            p.font.size = Pt(el.get("tamanho_fonte", 14))
            p.font.color.rgb = hex_para_rgb(el.get("cor_hex", "00"))
            
        elif tipo == "forma":
            shape = slide.shapes.add_shape(
                1, left, top, width, height
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = hex_para_rgb(
                el.get("cor_hex", "CC")
            )
            
        elif tipo == "imagem":
            l = int(el.get("x_percent", 0) * larg_px)
            t = int(el.get("y_percent", 0) * alt_px)
            w = int(el.get("largura_percent", 0.1) * larg_px)
            h = int(el.get("altura_percent", 0.1) * alt_px)
            
            recorte = imagem.crop((l, t, l+w, t+h))
            img_bytes = io.BytesIO()
            recorte.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            slide.shapes.add_picture(
                img_bytes, left, top, width, height
            )

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

img_file = st.file_uploader("Arraste a imagem", type=["png", "jpg"])

if img_file:
    st.image(img_file, use_container_width=True)
    formato = st.selectbox("Formato:", ["Widescreen", "Quadrado"])
    
    larg_slide = 10.0
    alt_slide = 10.0
    
    if formato == "Widescreen":
        alt_slide = 5.625
    
    if st.button("🚀 Processar", type="primary"):
        try:
            img_pil = Image.open(img_file)
            with st.status("Gerando...", expanded=True) as status:
                st.write("Lendo imagem com IA...")
                modelo = configurar_ia(CHAVE_API)
                resposta = modelo.generate_content(
                    [img_pil, "Gere o JSON."]
                )
                
                st.write("Montando slides...")
                texto = resposta.text
                texto = texto.replace('```json', '')
                texto = texto.replace('```', '').strip()
                
                dados_json = json.loads(texto)
                pptx_pronto = gerar_powerpoint(
                    dados_json, img_pil, larg_slide, alt_slide
                )
                status.update(label="Pronto!", state="complete")
            
            st.download_button(
                label="⬇️ Baixar .PPTX", 
                data=pptx_pronto, 
                file_name="sistema.pptx"
            )
        except Exception as e:
            st.error(f"Erro: {e}")
