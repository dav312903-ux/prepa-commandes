import streamlit as st
from google import genai
import io, json
from openpyxl import load_workbook
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

st.write("**Photo de la commande** (capture d'écran complète)")
order_photo_result = pbutton("📋 Coller la photo de la commande", key="order_paste", errors="ignore")

st.subheader("2. Lien")
link = st.text_input("Lien Alibaba")

def extract_data_from_photo(image_bytes, api_key):
    # Nouvelle syntaxe avec google-genai
    client = genai.Client(api_key=api_key)
    img = PILImage.open(io.BytesIO(image_bytes))
    
    prompt = """
    Analyse cette image de commande. Extrait l'adresse de livraison globale, le nom du client, le nom du produit, et la LISTE des variantes (couleurs) avec leurs quantités.
    Renvoie UNIQUEMENT un JSON avec cette structure exacte :
    {
      "client_name": "Nom du client pour le nom du fichier",
      "shipping": {
        "name": "Nom du destinataire",
        "street": "Numéro et rue",
        "address1": "Complément (Bâtiment, Casa...)",
        "city": "Ville",
        "zip": "Code postal",
        "province": "Province ou État",
        "country": "Pays"
      },
      "product_name": "Nom complet du produit (sans la couleur ni la quantité)",
      "variants": [
        {
          "color": "Couleur (ex: blanc, or, noir)",
          "quantity": 1
        }
      ]
    }
    """
    response = client.models.generate_content(
        model='gemini-1.5-flash',
        contents=[prompt, img]
    )
    text = response.text.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(text)

if st.button("Générer le fichier", type="primary"):
    if not template or not link:
        st.error("Veuillez fournir le tableau Excel et le lien Alibaba.")
        st.stop()
    
    if order_photo_result.image_data is None:
        st.error("Veuillez coller la photo de la commande.")
        st.stop()

    # Conversion de l'image de commande
    order_img_bytes = io.BytesIO()
    order_photo_result.image_data.save(order_img_bytes, format="PNG")
    order_img_bytes = order_img_bytes.getvalue()

    with st.spinner("L'IA analyse la commande..."):
        data = extract_data_from_photo(order_img_bytes, api_key)
        shipping = data.get("shipping", {})
        product_name = data.get("product_name", "")
        variants = data.get("variants", [])
        client_name = data.get("client_name", "Client")

        if not variants:
            st.error("Aucune variante ou quantité détectée sur la photo.")
            st.stop()

    with st.spinner("Remplissage du tableau..."):
        wb = load_workbook(io.BytesIO(template.getvalue()))
        ws = wb.active
        start_row = 2
        
        for i, var in enumerate(variants):
            current_row = start_row + i
            ws.cell(row=current_row, column=4, value=int(var.get("quantity", 1)))
            ws.cell(row=current_row, column=5, value=var.get("color", ""))

            if i == 0:
                values = {
                    2: product_name, 3: link, 6: shipping.get("name", ""), 
                    7: shipping.get("street", ""), 8: shipping.get("address1", ""),
                    10: shipping.get("city", ""), 11: shipping.get("zip", ""),
                    12: shipping.get("province", ""), 13: shipping.get("country", ""),
                }
                for col, val in values.items():
                    if val not in ("", None): ws.cell(row=current_row, column=col, value=val)

        # Sauvegarde temporaire
        temp_out = io.BytesIO()
        wb.save(temp_out)
        
        # L'astuce pour l'image (découpe de la vignette)
        full_img = PILImage.open(io.BytesIO(order_img_bytes))
        width, height = full_img.size
        crop_box = (0, 0, int(width * 0.15), int(height * 0.15)) 
        product_thumb = full_img.crop(crop_box)
        product_thumb.thumbnail((100, 70))
        img_path = "temp_prod_img.png"
        product_thumb.save(img_path, format="PNG")

        # Finalisation
        final_out = io.BytesIO()
        with pd.ExcelWriter(final_out, engine='xlsxwriter') as writer:
            df = pd.read_excel(io.BytesIO(temp_out.getvalue()), header=None)
            df.to_excel(writer, index=False, header=False)
            worksheet = writer.sheets['Sheet1']
            
            worksheet.insert_image(1, 15, img_path, {'x_scale': 0.8, 'y_scale': 0.8})
            
            for row_idx in range(1, 1 + len(variants)):
                worksheet.set_row(row_idx, 60)

        final_out.seek(0)

    st.success(f"Fichier généré pour {client_name} !")
    st.download_button(
        "⬇️ Télécharger le fichier Excel", data=final_out,
        file_name=f"Cmde {client_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
