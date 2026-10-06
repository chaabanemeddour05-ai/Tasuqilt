import os
import streamlit as st
import pandas as pd
import re
from google import genai
from google.genai import types

st.set_page_config(
    page_title="Tasuqilt Enterprise",
    page_icon="🇩🇿",
    layout="wide"
)

ADMIN_PASSWORD = "APS_Tasuqilt_2026"

if "live_corrections" not in st.session_state:
    st.session_state["live_corrections"] = ""
if "lexicon_data" not in st.session_state:
    st.session_state["lexicon_data"] = {}

st.title("Tasuqilt 🇩🇿")
st.subheader("منصة إدارة الترجمة الإعلامية والأمازيغية المعيارية الصارمة")

api_key = os.environ.get("GEMINI_API_KEY")
if not api_key and "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]

if not api_key:
    st.error("لم يتم العثور على مفتاح الأمان السري GEMINI_API_KEY.")
    st.stop()

api_key = api_key.replace('"', '').replace("'", "").strip()
client = genai.Client(api_key=api_key)

# 📊 قراءة قاعدة البيانات من ملف إكسيل الفائق الدقة والسرعة (tm.xlsx)
@st.cache_data
def load_translation_memory():
    file_name = "tm.xlsx"
    if not os.path.exists(file_name):
        return []
    try:
        df = pd.read_excel(file_name)
        memory_pairs = []
        for index, row in df.iterrows():
            french_text = str(row.iloc[0]).strip()
            tamazight_text = str(row.iloc[1]).strip()
            if french_text and tamazight_text and french_text != "nan" and tamazight_text != "nan":
                memory_pairs.append({
                    "tamazight": tamazight_text,
                    "foreign": french_text
                })
        return memory_pairs
    except Exception as e:
        return []

tm_data = load_translation_memory()

# 🛡️ البوابة الجانبية السرية للمشرفين
st.sidebar.header("🎛️ خيارات التحكم")
mode = st.sidebar.radio(
    "اختر وضع الترجمة الصحفية المعتمد:",
    ["إخباري رسمي وصارم", "أدبي / ثقافي"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("🔐 بوابة الإدارة السرية")
admin_input = st.sidebar.text_input("أدخل رمز الدخول لتمرين وتحديث المنصة:", type="password")

is_admin = (admin_input == ADMIN_PASSWORD)

if is_admin:
    st.sidebar.success("🔓 تم فتح صلاحيات الإدارة والتمرين المباشر!")
    raw_lexicon_text = st.sidebar.text_area("حقن المعجم الشامل (نسخ ولصق):", height=150)
    
    if st.sidebar.button("⚙️ تفكيك وحقن القاموس في الذاكرة الحية"):
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
            st.sidebar.success(f"✅ تم حقن {count} مصطلحاً!")

text_to_translate = st.text_area(
    "أدخل البرقية الصحفية المراد ترجمتها (بالفرنسية أو العربية):",
    height=220,
    placeholder="ضع نص البرقية هنا..."
)

if st.button("بدء الترجمة الاحترافية الموحدة", type="primary"):
    if not text_to_translate.strip():
        st.warning("يرجى إدخال نص أولًا للبدء.")
        st.stop()

    st.info("جاري فحص الذاكرة الجدولية وقفل القواميس الرسمية...")

    try:
        search_words = [w.strip().lower() for w in text_to_translate.split() if len(w.strip()) > 4]
        
        active_dict_context = "\n⚠️ RÈGLES DE TRADUCTION ABSOLUES ET OBLIGATOIRES :\n"
        active_dict_context += "- président = Aselway (Interdiction absolue d'utiliser Anmazul ou Anemhal)\n"
        active_dict_context += "- Conseil des ministres = Aseqqamu n Yineɣlaf\n"
        active_dict_context += "- réunion = Timlilt\n"
        active_dict_context += "- gouvernement = Anabaḍ\n"
        
        if st.session_state["lexicon_data"]:
            for k, v in st.session_state["lexicon_data"].items():
                if k in text_to_translate.lower():
                    active_dict_context += f"- {k} = {v}\n"

        matched_pairs = []
        if tm_data:
            for pair in tm_data:
                if any(word in pair["foreign"].lower() for word in search_words):
                    matched_pairs.append(pair)
                if len(matched_pairs) >= 12:
                    break
        if not matched_pairs:
            matched_pairs = tm_data[:6] if tm_data else []

        tm_context = ""
        for i, pair in enumerate(matched_pairs, 1):
            tm_context += f"نموذج مرجعي رسمي {i}:\nالنص الفرنسي الأصلي: {pair['foreign']}\nالترجمة الأمازيغية المعتمدة: {pair['tamazight']}\n---\n"

        instruction = """أنت رئيس تحرير ومترجم رسمي صارم للغة الأمازيغية المعيارية لصالح وكالة الأنباء (APS).
        مهمتك الحتمية والمقدسة: صياغة وترجمة البرقية المدخلة إلى الأمازيغية المعيارية الصارمة بالحرف اللاتيني.
        التزم التزاماً عسكرياً صارماً بالمعجم والنماذج المرفقة المستخرجة من الإكسيل. يُحظر تماماً التخمين أو الابتكار.
        يُمنع منعاً باتاً استخدام أي حروف عربية في النص الأمازيغي النهائي."""

        config = types.GenerateContentConfig(
            temperature=0.0,
            system_instruction=instruction
        )

        prompt = f"""{active_dict_context}
        
        المرجع الجدولى المعتمد لوكالة الأنباء:
        {tm_context}
        {st.session_state["live_corrections"]}

        النص الجديد المراد صياغته وترجمته الآن بدقة بالغة وبناءً على المعجم الحتمي:
        {text_to_translate}

        الترجمة الأمازيغية الرسمية الصارمة والنهائية (حرف لاتيني فقط):"""

        response = client.models.generate_content(
            model="gemini-3.5-flash-lite",
            contents=prompt,
            config=config
        )

        result = response.text

        if result and result.strip():
            output = result.strip()
            # 🛡️ استبدال قسري مباشر بالكود لمنع تسلل الكلمات الخاطئة نهائياً قبل العرض
            output = re.sub(r'\banmazul\b', 'Aselway', output, flags=re.IGNORECASE)
            output = re.sub(r'\banemhal\b', 'Aselway', output, flags=re.IGNORECASE)
            output = re.sub(r'yettu[εe]zlen', 'i yettwaheggan', output)
            output = re.sub(r'[\u0600-\u06FF]+', '', output).replace("  ", " ").strip()

            st.success("تمت الترجمة بنجاح واكتملت صياغة الخبر بناءً على الذاكرة الجدولية الفورية.")
            st.text_area("الترجمة الأمازيغية المعيارية النهائية (أسلوب APS رسمي ونظيف 100%):", value=output, height=280)
        else:
            st.error("لم ينجح النظام في معالجة النص، يرجى إعادة المحاولة.")

    except Exception as error:
        st.error("حدث خطأ تقني أثناء الاتصال بالذكاء الاصطناعي.")
        st.code(str(error))
