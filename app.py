import os
import streamlit as st
import requests
import re
import pandas as pd

# 1. Main Page and Configuration Layout
st.set_page_config(page_title="Tasuqilt Hybrid Enterprise", page_icon="🇩🇿", layout="wide")
ADMIN_PASSWORD = "APS_Tasuqilt_2026"

# 2. Base Prompts without any hardcoded alerts
if "custom_system_instruction" not in st.session_state:
    st.session_state["custom_system_instruction"] = "You are the official Chief Editor and Translator for the Algeria Press Service (APS). Your absolute mission is to handle translations accurately between languages based on user request. When translating to Tamazight, use standard corporate Tamazight with Latin characters. Strictly adhere to the provided official vocabulary and memory examples. Never invent new terms. Always translate 'ALGER' as 'DZAYER TAMANEƔT' and 'Président de la République' as 'Aselway n Tegduda'. Do not output any unwanted side notes or annotations. Give only the pure translation text."

st.title("Tasuqilt DZ 🇩🇿")
st.markdown("<p style='font-size:1.1rem; color:gray;'>Système de traduction officiel de l'Algérie Press Service (APS)</p>", unsafe_allow_html=True)

# 3. Dynamic Memory Matrix: Reads the updated database tm.xlsx directly from GitHub
@st.cache_data(max_entries=1)
def load_translation_memory():
    file_name = "tm.xlsx"
    if not os.path.exists(file_name):
        return []
    try:
        df = pd.read_excel(file_name)
        memory_pairs = []
        for index, row in df.iterrows():
            if len(row) >= 2:
                french_text = str(row.iloc[0]).strip()
                tamazight_text = str(row.iloc[1]).strip()
                if french_text and tamazight_text and french_text != "nan" and tamazight_text != "nan":
                    memory_pairs.append({"tamazight": tamazight_text, "foreign": french_text})
        return memory_pairs
    except Exception:
        return []

tm_data = load_translation_memory()

# 4. Sidebar Controls and Selector Interface
st.sidebar.header("⚙️ Configuration")
engine_choice = st.sidebar.selectbox("Moteur d'IA / عقل الذكاء الاصطناعي :", ["ChatGPT (GPT-4o)", "Claude 3.5 Sonnet", "DeepSeek V3", "Grok (X-AI)"])
direction = st.sidebar.selectbox("Direction / اتجاه الترجمة :", ["Auto-Detect [Français/Arabe] ➔ Tamazight", "Tamazight ➔ Auto-Detect [Français/Arabe]"])

st.sidebar.markdown("---")
st.sidebar.subheader("🔐 Administration")
admin_input = st.sidebar.text_input("Code d'accès admin :", type="password")

if admin_input == ADMIN_PASSWORD:
    st.sidebar.success("🔓 Mode Éditeur activé !")
    updated_prompt = st.sidebar.text_area("System Instruction (English):", value=st.session_state["custom_system_instruction"], height=120)
    if st.sidebar.button("💾 Enregistrer le Prompt"):
        st.session_state["custom_system_instruction"] = updated_prompt
        st.sidebar.success("✅ Prompt mis à jour !")

# 5. UI Layout split into parallel twin columns (Google Translate style)
col1, col2 = st.columns(2)

with col1:
    src_label = "Texte Source" if "Auto-Detect" in direction else "Texte Source (Tamazight)"
    text_to_translate = st.text_area(src_label + " :", height=250, placeholder="Saisissez ou collez votre paragraphe ici...")
    submit_button = st.button("Traduire la dépêche 🚀", type="primary")

output_text = ""
server_error = False

if submit_button and text_to_translate.strip():
    with st.spinner(f"Analyse contextuelle via {engine_choice}..."):
        search_words = [w.strip().lower() for w in text_to_translate.split() if len(w.strip()) > 4]
        active_dict_context = "\n⚠️ APS CRITICAL LEXICON RULES :\n"
        active_dict_context += "- président de la république = Aselway n Tegduda\n- رئيس الجمهورية = Aselway n Tegduda\n- le président = Aselway\n- ALGER = DZAYER TAMANEƔT\n- Alger = DZAYER TAMANEƔT\n- Conseil des ministres = Aseqqamu n Yineɣlaf\n- réunion = Timlilt\n- gouvernement = Anabaḍ\n"

        matched_pairs = []
        if tm_data:
            for pair in tm_data:
                match_source = pair["foreign"] if "Auto-Detect" in direction else pair["tamazight"]
                if any(word in match_source.lower() for word in search_words):
                    matched_pairs.append(pair)
                if len(matched_pairs) >= 6:
                    break
        if not matched_pairs and tm_data:
            matched_pairs = tm_data[:3]

        tm_context = ""
        for i, pair in enumerate(matched_pairs, 1):
            tm_context += f"Reference Framework {i}:\nOriginal: {pair['foreign']}\nTamazight (APS): {pair['tamazight']}\n---\n"

        direction_note = "Task: Automatically detect the source language and translate it into clear Latin Tamazight." if "Auto-Detect" in direction else "Task: Translate the Latin Tamazight text into professional French or Arabe."
        full_instruction = f"{st.session_state['custom_system_instruction']}\n\n{direction_note}"

        model_map = {"ChatGPT (GPT-4o)": "openai", "Claude 3.5 Sonnet": "claude", "DeepSeek V3": "deepseek", "Grok (X-AI)": "p1"}
        chosen_model = model_map.get(engine_choice, "openai")

        try:
            url = "https://pollinations.ai"
            payload = {
                "messages": [
                    {"role": "system", "content": f"{full_instruction}\n\nRules:\n{active_dict_context}\n\nReference data:\n{tm_context}"},
                    {"role": "user", "content": f"Translate perfectly: {text_to_translate}"}
                ],
                "model": chosen_model,
                "temperature": 0.0
            }
            response = requests.post(url, json=payload, timeout=25)
            if response.status_code == 200 and response.text.strip():
                raw_output = response.text.strip()
                if "Auto-Detect" in direction:
                    raw_output = re.sub(r'\banmazul\b', 'Aselway', raw_output, flags=re.IGNORECASE)
                    raw_output = re.sub(r'\banemhal\b', 'Aselway', raw_output, flags=re.IGNORECASE)
                    raw_output = re.sub(r'yettu[eε]zlen', 'i yettwaheggan', raw_output)
                    raw_output = re.sub(r'[\u0600-\u06FF]+', '', raw_output)
                output_text = raw_output.replace("  ", " ").strip()
            else:
                server_error = True
        except Exception:
            server_error = True

with col2:
    dst_label = "Texte Traduit (Tamazight)" if "Auto-Detect" in direction else "Texte Traduit (Français / Arabe)"
    st.markdown("**" + dst_label + " :**")
    if output_text:
        st.success("Traduction finalisée avec succès.")
        st.text_area("Résultat final :", value=output_text, height=250, key="result_box")
        st.info("💡 Vous pouvez copier le texte du résultat ci-dessus directement.")
    elif server_error:
        st.error("Le moteur sélectionné est surchargé, veuillez basculer sur un autre moteur dans le menu.")
    else:
        st.text_area("Résultat final :", value="", height=250, key="result_box_empty", disabled=True)
