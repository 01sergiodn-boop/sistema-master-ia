import streamlit as st
import json
import io
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image

# SUA CHAVE DIRETO NO CÓDIGO
CHAVE_API = "AIzaSyDkz_EVM8rzs1012mXdHS1J0GH53Iyu3GY"
genai.configure(api_key=CHAVE_API)

st.set_page_config(page_title="Sistema Master | IA")
st.title("✨ Sistema Master")

# CASCATA COM 10 MODELOS DIFERENTES
MODELOS = [
    "gemini-1.5-flash",
    "gemini-1.5-flash-latest",
    "gemini-1.5-pro",
    "gemini-1.5-pro-latest",
    "gemini-1.5-flash-001",
    "gemini-1.5-pro-001",
    "gemini-pro-vision",
    "gemini-1.0-pro-vision-latest",
    "models/gemini-1.5-flash",
    "models/gemini-1.5-pro"
]

def hex_para_rgb(hex_color):
    c = hex_color.replace('#', '')
    if len(c) != 6: 
        return RGBColor(0,0,0)
    return RGBColor(int(c[0:2],16), int(c[2:4],16), int(c[4:6],16))

def gerar_ppt(dados, img, larg, alt):
    prs = Presentation()
    prs.slide_width = Inches(larg)
    prs.slide_height = Inches(alt)
    slide = prs.slides.add_slide(prs.slide_layouts[6]) 
    larg_px, alt_px = img.size

    for el in dados:
        tipo = el.get("tipo", "")
        left = Inches(el.get("x_percent", 0) * larg)
        top = Inches(el.get("y_percent", 0) * alt)
        width = Inches(el.get("largura_percent", 0.1) * larg)
        height = Inches(el.get("altura_percent", 0.1) * alt)
        
        if tipo == "texto":
            tx = slide.shapes.add_textbox(left, top, width, height)
            p = tx.text_frame.paragraphs[0]
            p.text = el.get("conteudo", "")
            p.font.size = Pt(el.get("tamanho_fonte", 14))
            p.font.color.rgb = hex_para_rgb(el.get("cor_hex", "00"))
        elif tipo == "forma":
            sh = slide.shapes.add_shape(1, left, top, width, height)
            sh.fill.solid()
            sh.fill.fore_color.rgb = hex_para_rgb(el.get("cor_hex","CC"))
        elif tipo == "imagem":
            l = int(el.get("x_percent", 0) * larg_px)
            t = int(el.get("y_percent", 0) * alt_px)
            w = int(el.get("largura_percent", 0.1) * larg_px)
            h = int(el.get("altura_percent", 0.1) * alt_px)
            rec = img.crop((l, t, l+w, t+h))
            b = io.BytesIO()
            rec.save(b, format='PNG')
            b.seek(0)
            slide.shapes.add_picture(b, left, top, width, height)
            
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

img_file = st.file_uploader("Imagem", type=["png","jpg"])

if img_file:
    st.image(img_file)
    fmt = st.selectbox("Formato", ["Widescreen", "Quadrado"])
    larg_s = 10.0
    alt_s = 5.625 if fmt == "Widescreen" else 10.0
    
    prompt = (
        "Analise a imagem. Retorne APENAS um JSON válido. "
        "Formato: lista de dicionarios com: tipo, conteudo, "
        "cor_hex, tamanho_fonte, x_percent, y_percent, "
        "largura_percent, altura_percent."
    )

    if st.button("🚀 Processar"):
        img_pil = Image.open(img_file)
        sucesso = False
        
        with st.status("Testando modelos da cascata...", expanded=True) as stt:
            for mod in MODELOS:
                try:
                    st.write(f"🔄 Tentando conectar no modelo {mod}...")
                    ia = genai.GenerativeModel(mod)
                    resp = ia.generate_content([img_pil, prompt])
                    
                    st.write("✅ Deu certo! Lendo dados visuais...")
                    txt = resp.text.replace('```json', '')
                    txt = txt.replace('```', '').strip()
                    dados = json.loads(txt)
                    
                    st.write("⚡ Montando slides...")
                    pptx = gerar_ppt(dados, img_pil, larg_s, alt_s)
                    
                    stt.update(label="Finalizado com sucesso!", state="complete")
                    sucesso = True
                    
                    st.download_button(
                        "⬇️ Baixar .PPTX", pptx, "apresentacao.pptx"
                    )
                    break # Sai do loop quando o arquivo é gerado
                    
                except Exception as e:
                    st.write(f"❌ {mod} falhou. Indo para o próximo...")
                    continue # Pula para o próximo modelo da lista
            
            if not sucesso:
                stt.update(label="Erro!", state="error")
                st.error("Todos os 10 modelos falharam. Tente uma nova Chave API.")
