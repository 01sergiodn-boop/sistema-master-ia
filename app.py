import streamlit as st
import json
import io
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image

# Configuração
st.set_page_config(
    page_title="Sistema Master | IA",
    page_icon="✨",
    layout="centered"
)

# Estilo Limpo
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

st.title("✨ Sistema Master")
st.write("Transforme imagens em apresentações editáveis.")

# Cofre da Chave API
try:
    chave_api = st.secrets["GEMINI_API_KEY"]
except Exception:
    chave_api = ""

def configurar_ia(chave):
    genai.configure(api_key=chave)
    instrucao = (
        'Você é a IA do SISTEMA MASTER. '
        'Responda APENAS com um JSON '
        'com as chaves: "tipo", "conteudo", "cor_hex", '
        '"tamanho_fonte", "x_percent", "y_percent", '
        '"largura_percent", "altura_percent".'
    )
    # Usando o modelo super rápido
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
            txBox = slide.shapes.add_textbox(left, top, width, height)
            p = txBox.text_frame.paragraphs[0]
            p.text = el.get("conteudo", "")
            p.font.size = Pt(el.get("tamanho_fonte", 14))
            p.font.color.rgb = hex_para_rgb(el.get("cor_hex", "000000"))
            
        elif tipo == "forma":
            shape = slide.shapes.add_shape(1, left, top, width, height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = hex_para_rgb(el.get("cor_hex", "CCCCCC"))
            shape.line.fill.background()
            
        elif tipo == "imagem":
            l = int(el.get("x_percent", 0) * larg_px)
            t = int(el.get("y_percent", 0) * alt_px)
            w = int(el.get("largura_percent", 0.1) * larg_px)
            h = int(el.get("altura_percent", 0.1) * alt_px)
            
            r = l + w
            b = t + h
            
            recorte = imagem.crop((l, t, r, b))
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

# Interface Principal
arquivo_imagem = st.file_uploader(
    "Arraste a imagem aqui", 
    type=["png", "jpg", "jpeg"]
)

if arquivo_imagem:
    st.image(arquivo_imagem, use_container_width=True)
    
    opcoes = ["Widescreen (16:9)", "Quadrado (1:1)"]
    formato = st.selectbox("Formato:", opcoes)
    
    # Aqui era onde o celular cortava a linha!
    if "Widescreen" in formato:
        larg_slide = 10.0
        alt_slide = 5.625
    else:
        larg_slide = 10.0
        alt_slide = 10.0
    
    btn = st.button("🚀 Processar", type="primary")
    
    if btn:
        if not chave_api:
            st.error("⚠️ Chave API não encontrada!")
        else:
            try:
                img_pil = Image.open(arquivo_imagem)
                with st.status("Gerando...", expanded=True) as status:
                    
                    st.write("🧠 Lendo imagem...")
                    modelo = configurar_ia(chave_api)
                    resposta = modelo.generate_content(
                        [img_pil, "Gere o JSON."]
                    )
                    
                    st.write("⚡ Montando PPTX...")
                    texto = resposta.text
                    texto = texto.replace('```json', '')
                    texto = texto.replace('```', '')
                    texto = texto.strip()
                    
                    dados_json = json.loads(texto)
                    pptx_pronto = gerar_powerpoint(
                        dados_json, img_pil, larg_slide, alt_slide
                    )
                    
                    status.update(
                        label="Pronto!", 
                        state="complete", 
                        expanded=False
                    )
                
                st.download_button(
                    label="⬇️ Baixar .PPTX", 
                    data=pptx_pronto, 
                    file_name="sistema.pptx", 
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Erro: {e}")
