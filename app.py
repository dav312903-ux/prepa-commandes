import streamlit as st
import google.generativeai as genai
import io, json
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.cell import coordinate_to_tuple
from PIL import Image as PILImage
import pandas as pd
import xlsxwriter
from streamlit_paste_button import paste_image_button as pbutton

st.set_page_config(page_title="Préparation de commande", layout="centered")
st.title("Préparation de commande")

# --- Clé API Gemini ---
api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    api_key = st.sidebar.text_input("Clé API Gemini (Google)", type="password")

# --- Uploads ---
st.subheader("1. Fichiers")
template = st.file_uploader("Tableau Excel vierge (.xlsx)", type=["xlsx"])

# Boutons de collage pour les images
st.write("**Photo de la commande** (faites une capture d'écran, puis cliquez sur le bouton)")
order_photo_result = pbutton("📋 Coller la photo de la commande", key="order_paste")

st.write("**Photo du produit** (faites une capture d'écran, puis cliquez sur le bouton)")
product_photo_result = pbutton("📋 Coller la photo du produit", key="product_paste")

st.subheader("2. Informations")
link = st.text_input("Lien Alibaba")
client_name = st.text_input("Nom du client", placeholder="Ex. Christina")

def extract_data_from_photo(image_bytes, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash')
    img = PILImage.open(io.BytesIO(image_bytes))
    
    prompt = """
    Analyse cette image de commande. Extrait les informations et renvoie UNIQUEMENT un JSON avec ces clés (chaîne vide "" si absent) :
    {
      "product_name": "Nom complet du produit",
      "quantity": "Quantité (nombre entier)",
      "color": "Couleur",
      "shipping_name": "Nom du destinataire",
      "shipping_street": "Numéro et rue",
      "shipping_address1": "Complément (Bâtiment, Casa...)",
      "shipping_city": "Ville",
      "shipping_zip": "Code postal",
      "shipping_province": "Province (ex: TI)",
      "shipping_country": "Pays"
    }
    """
    response = model.generate_content([prompt, img])
    text = response.text.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(text)

if st.button("Générer le fichier", type="primary"):
    # Vérification des entrées
    if not all([template, link, client_name]):
        st.error("Veuillez remplir le tableau, le lien et le nom du client.")
        st.stop()
    
    if order_photo_result.image_data is None or product_photo_result.image_data is None:
        st.error("Veuillez coller les deux photos (commande et produit) avec les boutons.")
        st.stop()

    # Conversion des images PIL en bytes pour la suite du traitement
    order_img_bytes = io.BytesIO()
    order_photo_result.image_data.save(order_img_bytes, format="PNG")
    order_img_bytes = order_img_bytes.getvalue()

    product_img_bytes = io.BytesIO()
    product_photo_result.image_data.save(product_img_bytes, format="PNG")
    product_img_bytes = product_img_bytes.getvalue()

    with st.spinner("L'IA lit la commande..."):
        data = extract_data_from_photo(order_img_bytes, api_key)

    with st.spinner("Remplissage du tableau..."):
        # 1. On charge le VRAI modèle pour garder les couleurs
        wb = load_workbook(io.BytesIO(template.getvalue()))
        ws = wb.active
        
        # 2. On écrit les données dans les colonnes B à N
        row = 2
        values = {
            2: data.get("product_name", ""), 3: link, 4: int(data.get("quantity", 1)),
            5: data.get("color", ""), 6: data.get("shipping_name", ""), 
            7: data.get("shipping_street", ""), 8: data.get("shipping_address1", ""),
            9: "", 10: data.get("shipping_city", ""), 11: data.get("shipping_zip", ""),
            12: data.get("shipping_province", ""), 13: data.get("shipping_country", ""), 14: ""
        }
        for col, val in values.items():
            if val not in ("", None): ws.cell(row=row, column=col, value=val)

        # 3. L'astuce pour l'image : on utilise XlsxWriter pour forcer l'affichage
        img = PILImage.open(io.BytesIO(product_img_bytes)).convert("RGB")
        img.thumbnail((100, 70)) 
        img_path = "temp_prod_img.png"
        img.save(img_path, format="PNG")

        # Sauvegarde temporaire d'openpyxl
        temp_out = io.BytesIO()
        wb.save(temp_out)
        
        # Réécriture avec XlsxWriter pour l'image
        final_out = io.BytesIO()
        with pd.ExcelWriter(final_out, engine='xlsxwriter') as writer:
            df = pd.read_excel(io.BytesIO(temp_out.getvalue()), header=None)
            df.to_excel(writer, index=False, header=False)
            worksheet = writer.sheets['Sheet1']
            worksheet.insert_image(1, 15, img_path, {'x_scale': 0.8, 'y_scale': 0.8})
            worksheet.set_row(1, 60)

        final_out.seek(0)

    st.success("Fichier généré !")
    st.download_button(
        "⬇️ Télécharger le fichier Excel", data=final_out,
        file_name=f"Cmde {client_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
