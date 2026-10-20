import streamlit as st
import requests
import io, json, base64
from openpyxl import load_workbook
from PIL import Image as PILImage
import pandas as pd
import xlsxwriter

st.set_page_config(page_title="Préparation de commande", layout="centered")
st.title("Préparation de commande")

# --- Clés API Cloudflare ---
account_id = st.secrets.get("CF_ACCOUNT_ID", "")
api_token = st.secrets.get("CF_API_TOKEN", "")
if not account_id or not api_token:
    st.sidebar.warning("Clés Cloudflare manquantes dans les Secrets.")

# --- Uploads ---
st.subheader("1. Fichiers")
template = st.file_uploader("Tableau Excel vierge (.xlsx)", type=["xlsx"])
order_photo = st.file_uploader("Photo de la commande (capture d'écran)", type=["jpg", "jpeg", "png"])

st.subheader("2. Lien")
link = st.text_input("Lien Alibaba")

def extract_data_from_photo(image_bytes, account_id, api_token):
    # Conversion de l'image en base64
    img = PILImage.open(io.BytesIO(image_bytes))
    buffered = io.BytesIO()
    img.save(buffered, format="PNG")
    img_b64 = base64.b64encode(buffered.getvalue()).decode('utf-8')

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

    # Appel direct à l'API Cloudflare Workers AI
    url = f"https://api.cloudflare.com/client/v4/accounts/{account_id}/ai/run/@cf/meta/llama-3.2-11b-vision-instruct"
    headers = {"Authorization": f"Bearer {api_token}"}
    payload = {
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_b64}"}}
                ]
            }
        ]
    }
    
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status() # Lève une erreur si l'API échoue
    result = response.json()
    
    # On récupère la réponse texte de l'IA
    text_response = result["result"]["response"]
    # Nettoyage pour ne garder que le JSON
    text_response = text_response.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(text_response)

if st.button("Générer le fichier", type="primary"):
    if not template or not link or not order_photo:
        st.error("Veuillez fournir le tableau Excel, la photo et le lien Alibaba.")
        st.stop()

    order_img_bytes = order_photo.getvalue()

    with st.spinner("L'IA analyse la commande..."):
        try:
            data = extract_data_from_photo(order_img_bytes, account_id, api_token)
            shipping = data.get("shipping", {})
            product_name = data.get("product_name", "")
            variants = data.get("variants", [])
            client_name = data.get("client_name", "Client")
        except Exception as e:
            st.error(f"Erreur lors de l'analyse par l'IA : {e}")
            st.stop()

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

        temp_out = io.BytesIO()
        wb.save(temp_out)
        
        # Découpe de la vignette (on garde l'astuce)
        full_img = PILImage.open(io.BytesIO(order_img_bytes))
        width, height = full_img.size
        crop_box = (0, 0, int(width * 0.15), int(height * 0.15)) 
        product_thumb = full_img.crop(crop_box)
        product_thumb.thumbnail((100, 70))
        img_path = "temp_prod_img.png"
        product_thumb.save(img_path, format="PNG")

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
