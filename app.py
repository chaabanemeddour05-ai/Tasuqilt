import os
import streamlit as st
import requests
import re
from docx import Document

# 1. Main Page Configuration
st.set_page_config(page_title="Tasuqilt Hybrid Enterprise", page_icon="🇩🇿", layout="wide")
ADMIN_PASSWORD = "APS_Tasuqilt_2026"

# 2. Initialize Core Session States
if "custom_system_instruction" not in st.session_state:
    st.session_state["custom_system_instruction"] = "You are the official Chief Editor and Translator for the Algeria Press Service (APS). Your absolute mission is to handle translations accurately between languages based on user request. When translating to Tamazight, use standard corporate Tamazight with Latin characters. Strictly adhere to the provided official vocabulary and memory examples. Never invent new terms. Always translate 'ALGER' as 'DZAYER TAMANEƔT' and 'Président de la République' as 'Aselway n Tegduda'. Do not output any unwanted side notes. Give only the pure translation."
if "live_corrections" not in st.session_state:
    st.session_state["live_corrections"] = ""
if "lexicon_data" not in st.session_state:
    st.session_state["lexicon_data"] = {}

st.title("Tasuqilt DZ 🇩🇿")
st.markdown("<p style='font-size:1.1rem; color:gray;'>Système de traduction intelligent et multi-moteurs (ChatGPT / Claude / DeepSeek / Grok)</p>", unsafe_allow_html=True)

# 3. Securely Load Translation Memory Document (tm.docx)
@st.cache_data(max_entries=1)
def load_translation_memory():
    file_name = "tm.docx"
    if not os.path.exists(file_name):
        return []
    try:
        doc = Document(file_name)
        memory_pairs = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if "@" in text and len(text) > 10:
                parts = text.split("@")
                if len(parts) >= 2:
                    memory_pairs.append({"tamazight": parts[0].strip(), "foreign": parts[1].strip()})
        return memory_pairs
    except Exception:
        return []

tm_data = load_translation_memory()

# 4. Sidebar Controls & Administration Gate
st.sidebar.header("⚙️ Configuration")
engine_choice = st.sidebar.selectbox("Moteur d'IA / عقل الذكاء الاصطناعي :", ["ChatGPT (GPT-4o)", "Claude 3.5 Sonnet", "DeepSeek V3", "Grok (X-AI)"])
direction = st.sidebar.selectbox("Direction de traduction / اتجاه الترجمة :", ["Auto-Detect [Français/Arabe] ➔ Tamazight", "Tamazight ➔ Auto-Detect [Français/Arabe]"])

st.sidebar.markdown("---")
st.sidebar.subheader("🔐 Espace Administration")
admin_input = st.sidebar.text_input("Code d'accès admin :", type="password")

if admin_input == ADMIN_PASSWORD:
    st.sidebar.success("🔓 Mode Éditeur activé !")
    updated_prompt = st.sidebar.text_area("System Instruction (English):", value=st.session_state["custom_system_instruction"], height=120)
    if st.sidebar.button("💾 Enregistrer le Prompt"):
        st.session_state["custom_system_instruction"] = updated_prompt
        st.sidebar.success("✅ Prompt mis à jour !")
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("📚 Dictionnaire (Copier-Coller)")
    raw_lexicon_text = st.sidebar.text_area("Collez la liste des mots ici :", height=120)
    if st.sidebar.button("⚙️ Injecter le lexique"):
        if raw_lexicon_text.strip():
            count = 0
            for line in raw_lexicon_text.strip().split("\n"):
                amazigh_match = re.match(r"^([a-zA-ZɛɣɣƐƔƔ]+)", line.strip())
                if amazigh_match:
                    amazigh_word = amazigh_match.group(1).strip()
                    french_parts = re.findall(r"([^,\(\)\-\–]+)\s*\(fr\)", line.lower())
                    for fr in french_parts:
                        for sub_fr in fr.split(","):
                            st.session_state["lexicon_data"][sub_fr.strip().lower()] = amazigh_word
                    count += 1
            st.sidebar.success(f"✅ {count} termes intégrés !")

# 5. SECTION 1: INSTANT DICTIONARY (EXPRESS LEXICON)
st.markdown("### 📖 Dictionnaire Express / القاموس الفوري السريع")
dict_col1, dict_col2 = st.columns(2)

with dict_col1:
    word_to_find = st.text_input("Entrez un mot ou terme à chercher :", placeholder="Ex: réunion , président...")
with dict_col2:
    st.markdown("**Résultat du dictionnaire :**")
    if word_to_find.strip():
        clean_query = word_to_find.strip().lower()
        if clean_query in st.session_state["lexicon_data"]:
            st.code(st.session_state["lexicon_data"][clean_query], language="text")
        elif "réunion" in clean_query or "اجتماع" in clean_query:
            st.code("Timlilt", language="text")
        elif "président" in clean_query or "رئيس" in clean_query:
            st.code("Aselway", language="text")
        else:
            st.info("Terme non trouvé dans le dictionnaire local.")

st.markdown("---")

# 6. SECTION 2: PARALLEL TRANSLATOR CLASSIC LAYOUT
st.markdown("### 📰 Traducteur de Dépêches / مترجم البرقيات الإعلامية")
col1, col2 = st.columns(2)

with col1:
    src_label = "Source Text" if "Auto-Detect" in direction else "Texte Source (Tamazight)"
    text_to_translate = st.text_area(src_label + " :", height=250, placeholder="Saisissez ou collez votre paragraphe ici...")
    submit_button = st.button("Traduire s'il vous plaît 🚀", type="primary")

# Flat Global Safe Variables Initialization
output_text = ""
server_error = False

if submit_button and text_to_translate.strip():
    search_words = [w.strip().lower() for w in text_to_translate.split() if len(w.strip()) > 4]
    active_dict_context = "\n⚠️ APS CRITICAL LEXICON RULES :\n"
    
    if "Auto-Detect" in direction:
        active_dict_context += "- président de la république = Aselway n Tegduda\n- رئيس الجمهورية = Aselway n Tegduda\n- le président = Aselway\n- alger = DZAYER TAMANEƔT\n- Conseil des ministres = Aseqqamu n Yineɣlaf\n- réunion = Timlilt\n- gouvernement = Anabaḍ\n"
        if st.session_state["lexicon_data"]:
            for k, v in st.session_state["lexicon_data"].items():
                if k in text_to_translate.lower():
                    active_dict_context += f"- {k} = {v}\n"
    else:
        active_dict_context += "- Aselway n Tegduda = Président de la République\n- Aselway = Président\n- DZAYER TAMANEƔT = ALGER\n- Aseqqamu n Yineɣlaf = Conseil des ministres\n- Timlilt = Réunion\n- Anabaḍ = Gouvernement\n"

    matched_pairs = []
    if tm_data:
        for pair in tm_data:
            match_source = pair["foreign"] if "Auto-Detect" in direction else pair["tamazight"]
            if any(word in match_source.lower() for word in search_words):
                matched_pairs.append(pair)
            if len(matched_pairs) >= 6:
                break
    if not matched_pairs:
        matched_pairs = tm_data[:3] if tm_data else []

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
    except Exception as error:
        st.error(f"Technical Error: {error}")

with col2:
    dst_label = "Texte Traduit (Tamazight)" if "Auto-Detect" in direction else "Texte Traduit (Français / Arabe)"
    st.markdown("**" + dst_label + " :**")
    if output_text:
        st.success("Traduction finalisée avec succès.")
        st.text_area("Résultat final :", value=output_text, height=250, key="result_box")
        st.info("💡 Vous pouvez copier le texte du résultat ci-dessus directement.")
    elif server_error:
        st.error("Le moteur sélectionné est surchargé, veuillez basculer sur un autre moteur.")
    else:
        st.text_area("Résultat final :", value="", height=250, key="result_box_empty", disabled=True)
