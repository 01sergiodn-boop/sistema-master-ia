import streamlit as st
import json
import io
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Sistema Master | IA", page_icon="✨", layout="centered")

st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    [data-testid="stFileUploadDropzone"] { border-radius: 20px; border: 2px dashed #4285F4; }
    [data-testid="baseButton-primary"] { border-radius: 30px; background-color: #1A73E8; border: none; }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center;'>✨ Sistema Master</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #888;'>Transforme imagens em apresentações editáveis instantaneamente.</p><br>", unsafe_allow_html=True)

# --- PUXANDO A CHAVE SECRETA AUTOMATICAMENTE ---
try:
    chave_api = st.secrets["GEMINI_API_KEY"]
except Exception:
    chave_api = ""

def configurar_ia(chave):
    genai.configure(api_key=chave)
    instrucao = 'Você é a IA do SISTEMA MASTER. Analise a imagem e devolva a estrutura exata. Responda APENAS com um JSON válido (lista de dicionários) com as chaves: "tipo" ("texto", "forma", "imagem"), "conteudo", "cor_hex", "tamanho_fonte", "x_percent", "y_percent", "largura_percent", "altura_percent".'
    return genai.GenerativeModel('gemini-1.5-pro', system_instruction=instrucao)

def hex_para_rgb(hex_color):
    hex_color = hex_color.replace('#', '')
    if len(hex_color) != 6: return RGBColor(0, 0, 0)
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

def gerar_powerpoint(dados_json, imagem, larg_in, alt_in):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(larg_in), Inches(alt_in)
    slide = prs.slides.add_slide(prs.slide_layouts[6]) 
    larg_px, alt_px = imagem.size

    for el in dados_json:
        left, top = Inches(el.get("x_percent", 0) * larg_in), Inches(el.get("y_percent", 0) * alt_in)
        width, height = Inches(el.get("largura_percent", 0.1) * larg_in), Inches(el.get("altura_percent", 0.1) * alt_in)
        
        if el["tipo"] == "texto":
            txBox = slide.shapes.add_textbox(left, top, width, height)
            p = txBox.text_frame.paragraphs[0]
            p.text = el.get("conteudo", "")
            p.font.size = Pt(el.get("tamanho_fonte", 14))
            p.font.color.rgb = hex_para_rgb(el.get("cor_hex", "000000"))
            
        elif el["tipo"] == "forma":
            shape = slide.shapes.add_shape(1, left, top, width, height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = hex_para_rgb(el.get("cor_hex", "CCCCCC"))
            shape.line.fill.background()
            
        elif el["tipo"] == "imagem":
            l = int(el.get("x_percent", 0) * larg_px)
            t = int(el.get("y_percent", 0) * alt_px)
            r = l + int(el.get("largura_percent", 0.1) * larg_px)
            b = t + int(el.get("altura_percent", 0.1) * alt_px)
            
            recorte = imagem.crop((l, t, r, b))
            img_bytes = io.BytesIO()
            recorte.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            slide.shapes.add_picture(img_bytes, left, top, width, height)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

arquivo_imagem = st.file_uploader("Arraste seu design ou screenshot aqui", type=["png", "jpg", "jpeg"])

if arquivo_imagem:
    st.image(arquivo_imagem, use_container_width=True)
    formato = st.selectbox("📏 Formato:", ["Widescreen (16:9)", "Quadrado (1:1)"])
    larg_slide, alt_slide = (10.0, 5.625) if "Widescreen" in formato else (10.0, 10.0)
    
    if st.button("🚀 Processar com IA", type="primary", use_container_width=True):
        if not chave_api:
            st.error("⚠️ A Chave API não foi encontrada nos Segredos do Streamlit!")
        else:
            try:
                img_pil = Image.open(arquivo_imagem)
                with st.status("Processando...", expanded=True) as status:
                    st.write("🧠 Cérebro IA analisando geometria...")
                    modelo = configurar_ia(chave_api)
                    resposta = modelo.generate_content([img_pil, "Gere o JSON."])
                    
                    st.write("⚡ Montando slides na velocidade da luz...")
                    dados_json = json.loads(resposta.text.replace('```json', '').replace('```', '').strip())
                    pptx_pronto = gerar_powerpoint(dados_json, img_pil, larg_slide, alt_slide)
                    
                    status.update(label="Pronto!", state="complete", expanded=False)
                
                st.download_button("⬇️ Baixar .PPTX", pptx_pronto, "sistema.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
            except Exception as e:
                st.error(f"Erro: {e}")
