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

# Init session states for custom prompt edits, lexicon data and corrections
if "custom_system_instruction" not in st.session_state:
    st.session_state["custom_system_instruction"] = """You are the official Chief Editor and Translator for the Algeria Press Service (APS). 
Your absolute mission is to translate the source text into standard corporate Tamazight using Latin characters.
Strictly adhere to the provided official vocabulary and memory examples. Never invent new terms.
Always translate 'ALGER' as 'DZAYER TAMANEƔT' and 'Président de la République' as 'Aselway n Tegduda'.
Do not output any Arabic characters, explanations, or side notes. Give only the pure Tamazight translation."""

if "live_corrections" not in st.session_state:
    st.session_state["live_corrections"] = ""
if "lexicon_data" not in st.session_state:
    st.session_state["lexicon_data"] = {}

st.title("Tasuqilt DZ 🇩🇿")
st.markdown("<p style='font-size:1.2rem; color:gray;'>Traduction officielle de la langue Amazighe Standard</p>", unsafe_allow_html=True)

# Secure API configuration for Gemini 2.5 Flash Free Tier
api_key = os.environ.get("GEMINI_API_KEY")
if not api_key and "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]

if not api_key:
    st.error("Error: GEMINI_API_KEY introuvable dans les paramètres secrets.")
    st.stop()

api_key = api_key.replace('"', '').replace("'", "").strip()
client = genai.Client(api_key=api_key)

# Load translation pairs from tm.docx safely
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
                        "tamazight": parts[0].strip(),
                        "foreign": parts[1].strip()
                    })
            if len(memory_pairs) >= 3000:
                break
        return memory_pairs
    except Exception:
        return []

tm_data = load_translation_memory()

# Sidebar: Controls and Secret Administrator Panel
st.sidebar.header("⚙️ Configuration")
source_lang = st.sidebar.radio(
    "Langue source / لغة المدخلات :",
    ["Français", "Arabe"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("🔐 Espace Administration")
admin_input = st.sidebar.text_input("Code d'accès admin :", type="password")

is_admin = (admin_input == ADMIN_PASSWORD)

if is_admin:
    st.sidebar.success("🔓 Mode Éditeur activé !")
    
    # Live Prompt Editor Tool
    st.sidebar.subheader("📝 Modifier le Prompt (English)")
    updated_prompt = st.sidebar.text_area("System Instruction :", value=st.session_state["custom_system_instruction"], height=150)
    if st.sidebar.button("💾 Enregistrer le Prompt"):
        st.session_state["custom_system_instruction"] = updated_prompt
        st.sidebar.success("✅ Prompt mis à jour en direct !")
        
    # Bulk Lexicon Injector Tool
    st.sidebar.markdown("---")
    st.sidebar.subheader("📚 Dictionnaire et lexique (Copier-Coller)")
    raw_lexicon_text = st.sidebar.text_area("Collez la liste des mots ici :", height=150)
    
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
            st.sidebar.success(f"✅ {count} termes intégrés avec succès !")

# Main UI layout: Two parallel columns (Google Translate style)
col1, col2 = st.columns(2)

with col1:
    text_to_translate = st.text_area(
        f"Texte d'origine ({source_lang}) :",
        height=250,
        placeholder="Saisissez ou collez votre texte ici..."
    )
    submit_button = st.button("Traduire 🚀", type="primary")

with col2:
    st.markdown("**Texte traduit (Tamazight - Caractères Latins) :**")
    output_placeholder = st.empty()
    translated_text_val = ""

if submit_button:
    if not text_to_translate.strip():
        st.warning("Veuillez entrer un texte à traduire.")
    else:
        with st.spinner("Analyse du lexique et traduction en cours..."):
            try:
                search_words = [w.strip().lower() for w in text_to_translate.split() if len(w.strip()) > 4]
                
                # Hardcoded morphological core fallback rules
                active_dict_context = "\n⚠️ ABSOLUTE LEXICON RULES CRITICAL :\n"
                active_dict_context += "- président de la république = Aselway n Tegduda\n"
                active_dict_context += "- رئيس الجمهورية = Aselway n Tegduda\n"
                active_dict_context += "- le président = Aselway\n"
                active_dict_context += "- alger = DZAYER TAMANEƔT\n"
                active_dict_context += "- Conseil des ministres = Aseqqamu n Yineɣlaf\n"
                active_dict_context += "- réunion = Timlilt\n"
                active_dict_context += "- gouvernement = Anabaḍ\n"
                
                if st.session_state["lexicon_data"]:
                    for k, v in st.session_state["lexicon_data"].items():
                        if k in text_to_translate.lower():
                            active_dict_context += f"- {k} = {v}\n"

                # Smart matching for translation pairs
                matched_pairs = []
                if tm_data:
                    for pair in tm_data:
                        if any(word in pair["foreign"].lower() for word in search_words):
                            matched_pairs.append(pair)
                        if len(matched_pairs) >= 10:
                            break
                if not matched_pairs:
                    matched_pairs = tm_data[:5] if tm_data else []

                tm_context = ""
                for i, pair in enumerate(matched_pairs, 1):
                    tm_context += f"Reference Pair {i}:\nOriginal: {pair['foreign']}\nImposed Translation: {pair['tamazight']}\n---\n"

                config = types.GenerateContentConfig(
                    temperature=0.0, # Complete crystallization to prevent fallback hallucinations
                    system_instruction=st.session_state["custom_system_instruction"]
                )

                prompt = f"""{active_dict_context}
                
                APS Translation Memory Context:
                {tm_context}

                New text to translate perfectly now:
                {text_to_translate}

                Output Translation (Standard Latin Tamazight Only):"""

                # Direct Call to Gemini 2.5 Flash Free Tier
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt,
                    config=config
                )

                if response.text:
                    output = response.text.strip()
                    # Final algorithmic post-processing to replace ancient corrupted phrases
                    output = re.sub(r'\banmazul\b', 'Aselway', output, flags=re.IGNORECASE)
                    output = re.sub(r'\banemhal\b', 'Aselway', output, flags=re.IGNORECASE)
                    output = re.sub(r'yettu[εe]zlen', 'i yettwaheggan', output)
                    output = re.sub(r'[\u0600-\u06FF]+', '', output).replace("  ", " ").strip()
                    
                    translated_text_val = output
                    
                    with col2:
                        st.text_area("Résultat :", value=translated_text_val, height=250, key="result_box")
                        # Built-in instant single-click copy button functionality
                        st.info("💡 Vous pouvez copier le texte ci-dessus directement.")
                else:
                    st.error("Le serveur n'a renvoyé aucun texte.")
            except Exception as error:
                st.error(f"Technical Error: {error}")
