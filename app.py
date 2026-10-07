import os
import streamlit as st
import re
from google import genai
from google.genai import types
from docx import Document

st.set_page_config(
    page_title="Tasuqilt Enterprise",
    page_icon="🇩🇿",
    layout="wide"
)

ADMIN_PASSWORD = "APS_Tasuqilt_2026"

# Init default strict English instructions for Gemini
if "custom_system_instruction" not in st.session_state:
    st.session_state["custom_system_instruction"] = """You are the official Chief Editor and Translator for the Algeria Press Service (APS). 
Your absolute mission is to handle translations accurately between languages based on user request.
When translating to Tamazight, use standard corporate Tamazight with Latin characters.
Strictly adhere to the provided official vocabulary and memory examples. Never invent new terms.
Always translate 'ALGER' as 'DZAYER TAMANEƔT' and 'Président de la République' as 'Aselway n Tegduda'.
Do not output any unwanted side notes. Give only the pure translation."""

if "live_corrections" not in st.session_state:
    st.session_state["live_corrections"] = ""
if "lexicon_data" not in st.session_state:
    st.session_state["lexicon_data"] = {}

st.title("Tasuqilt DZ 🇩🇿")
st.markdown("<p style='font-size:1.1rem; color:gray;'>Système de traduction intelligent et gestion lexicale (APS)</p>", unsafe_allow_html=True)

# Secure API Configuration
api_key = os.environ.get("GEMINI_API_KEY")
if not api_key and "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]

if not api_key:
    st.error("Error: GEMINI_API_KEY introuvable dans les paramètres secrets.")
    st.stop()

api_key = api_key.replace('"', '').replace("'", "").strip()
client = genai.Client(api_key=api_key)

# Load translation memory tm.docx safely
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
                    memory_pairs.append({
                        "tamazight": parts.strip(),
                        "foreign": parts.strip()
                    })
        return memory_pairs
    except Exception:
        return []

tm_data = load_translation_memory()

# Sidebar Setup & Controls
st.sidebar.header("⚙️ Configuration")
direction = st.sidebar.selectbox(
    "Direction de traduction / اتجاه الترجمة :",
    ["Auto-Detect [Français/Arabe] ➔ Tamazight", "Tamazight ➔ Auto-Detect [Français/Arabe]"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("🔐 Espace Administration")
admin_input = st.sidebar.text_input("Code d'accès admin :", type="password")

is_admin = (admin_input == ADMIN_PASSWORD)

if is_admin:
    st.sidebar.success("🔓 Mode Éditeur activé !")
    
    # Prompt Editor
    st.sidebar.subheader("📝 Modifier le Prompt (English)")
    updated_prompt = st.sidebar.text_area("System Instruction :", value=st.session_state["custom_system_instruction"], height=120)
    if st.sidebar.button("💾 Enregistrer le Prompt"):
        st.session_state["custom_system_instruction"] = updated_prompt
        st.sidebar.success("✅ Prompt mis à jour !")
        
    # Bulk Lexicon Injector
    st.sidebar.markdown("---")
    st.sidebar.subheader("📚 Dictionnaire (Copier-Coller)")
    raw_lexicon_text = st.sidebar.text_area("Collez la liste des mots ici :", height=120)
    if st.sidebar.button("⚙️ Injecter le lexique"):
        if raw_lexicon_text.strip():
            count = 0
            lines = raw_lexicon_text.strip().split("\n")
            for line in lines:
                amazigh_match = re.match(r"^([a-zA-ZɛɣɣƐƔƔ]+)", line.strip())
                if amazigh_match:
                    amazigh_word = amazigh_match.group(1).strip()
                    french_parts = re.findall(r"([^,\(\)\-\–]+)\s*\(fr\)", line.lower())
                    for fr in french_parts:
                        for sub_fr in fr.split(","):
                            st.session_state["lexicon_data"][sub_fr.strip().lower()] = amazigh_word
                    count += 1
            st.sidebar.success(f"✅ {count} termes intégrés !")

# 📖 SECTION 1: INSTANT DICTIONARY (صندوق القاموس المصغر للمصطلحات والكلمات المفردة)
st.markdown("### 📖 Dictionnaire Express / القاموس الفوري السريع")
dict_col1, dict_col2 = st.columns([2, 2])

with dict_col1:
    word_to_find = st.text_input("Entrez un mot ou terme à chercher / أدخل كلمة أو مصطلح مفرد :", placeholder="Ex: réunion , président, تلميذ...")
with dict_col2:
    st.markdown("**Résultat du dictionnaire :**")
    if word_to_find.strip():
        clean_query = word_to_find.strip().lower()
        # Check in active dynamic lexicon data first
        if clean_query in st.session_state["lexicon_data"]:
            st.code(st.session_state["lexicon_data"][clean_query], language="text")
        elif "réunion" in clean_query or "اجتماع" in clean_query:
            st.code("Timlilt", language="text")
        elif "président" in clean_query or "رئيس" in clean_query:
            st.code("Aselway", language="text")
        else:
            st.info("Terme non trouvé dans le dictionnaire local. Utilisez la zone de texte globale ci-dessous.")

st.markdown("---")

# 📝 SECTION 2: GLOBAL TRANSLATOR (نظام ترجمة البرقيات الطويلة التبادلي المتقابل)
st.markdown("### 📰 Traducteur de Dépêches / مترجم البرقيات الإعلامية")
col1, col2 = st.columns(2)

with col1:
    src_label = "Source Text (Auto-Detect Language)" if "Auto-Detect" in direction else "Texte Source (Tamazight - Lettres Latines)"
    text_to_translate = st.text_area(
        f"{src_label} :",
        height=250,
        placeholder="Saisissez ou collez votre paragraphe ici..."
    )
    submit_button = st.button("Traduire la dépêche 🚀", type="primary")

with col2:
    dst_label = "Texte Traduit (Tamazight - Lettres Latines)" if "Auto-Detect" in direction else "Texte Traduit (Français / Arabe)"
    st.markdown(f"**{dst_label} :**")
    output_placeholder = st.empty()

if submit_button:
    if not text_to_translate.strip():
        st.warning("Veuillez entrer un paragraphe à traduire.")
    else:
        with st.spinner("Analyse contextuelle et traduction en cours..."):
            try:
                search_words = [w.strip().lower() for w in text_to_translate.split() if len(w.strip()) > 4]
                
                # Active morphological injection block
                active_dict_context = "\n⚠️ APS CRITICAL LEXICON RULES :\n"
                if "Auto-Detect" in direction:
                    active_dict_context += "- président de la république = Aselway n Tegduda\n- رئيس الجمهورية = Aselway n Tegduda\n- le président = Aselway\n- alger = DZAYER TAMANEƔT\n- Conseil des ministres = Aseqqamu n Yineɣlaf\n- réunion = Timlilt\n- gouvernement = Anabaḍ\n"
                    if st.session_state["lexicon_data"]:
                        for k, v in st.session_state["lexicon_data"].items():
                            if k in text_to_translate.lower():
                                active_dict_context += f"- {k} = {v}\n"
                else:
                    active_dict_context += "- Aselway n Tegduda = Président de la République\n- Aselway = Président\n- DZAYER TAMANEƔT = ALGER\n- Aseqqamu n Yineɣlaf = Conseil des ministres\n- Timlilt = Réunion\n- Anabaḍ = Gouvernement\n"

                # Extract relevant pairs from tm.docx based on vocabulary
                matched_pairs = []
                if tm_data:
                    for pair in tm_data:
                        match_source = pair["foreign"] if "Auto-Detect" in direction else pair["tamazight"]
                        if any(word in match_source.lower() for word in search_words):
                            matched_pairs.append(pair)
                        if len(matched_pairs) >= 10:
                            break
                if not matched_pairs:
                    matched_pairs = tm_data[:5] if tm_data else []

                tm_context = ""
                for i, pair in enumerate(matched_pairs, 1):
                    tm_context += f"Reference Ejemplo {i}:\nOriginal: {pair['foreign']}\nTamazight (APS): {pair['tamazight']}\n---\n"

                # Define final dynamic configuration instructions
                direction_note = "Task: Automatically detect the source language and translate it into clear Latin Tamazight." if "Auto-Detect" in direction else "Task: Translate the Latin Tamazight text into professional French or Arabe based on structure."
                full_instruction = f"{st.session_state['custom_system_instruction']}\n\n{direction_note}"

                config = types.GenerateContentConfig(
                    temperature=0.0,
                    system_instruction=full_instruction
                )

                prompt = f"""{active_dict_context}
                
                APS Translation Reference Framework:
                {tm_context}

                Text to translate now according to the selected direction:
                {text_to_translate}

                Output Translation:"""

                # Core Request to Gemini 2.5 Flash Free Tier
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt,
                    config=config
                )

                if response.text:
                    output = response.text.strip()
                    
                    # Post-processing clean filter
                    if "Auto-Detect" in direction:
                        output = re.sub(r'\banmazul\b', 'Aselway', output, flags=re.IGNORECASE)
                        output = re.sub(r'\banemhal\b', 'Aselway', output, flags=re.IGNORECASE)
                        output = re.sub(r'yettu[εe]zlen', 'i yettwaheggan', output)
