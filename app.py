import os
import streamlit as st
import requests
import re
from docx import Document

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
st.subheader("منصة الترجمة الإعلامية الصارمة (بمحرك الوورد والمستودع المفتوح)")

# قراءة المرجع المقدس للفقرات المترجمة من ملف الوورد tm.docx المخفف والموفر للطاقة
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
            if len(memory_pairs) >= 3000: # حد أمان لضمان سرعة السيرفر الخارقة
                break
        return memory_pairs
    except Exception:
        return []

tm_data = load_translation_memory()

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

    st.info("جاري فحص مستند الوورد وتفعيل المحرك المستقر فوراً...")

    try:
        search_words = [w.strip().lower() for w in text_to_translate.split() if len(w.strip()) > 4]
        
        # قفل المصطلحات الحيوية بقوانين برمجية صارمة في الكود لمنع أي هلوسة
        active_dict_context = "\n⚠️ RÈGLES DE TRADUCTION ABSOLUES ET OBLIGATOIRES :\n"
        active_dict_context += "- رئيس الجمهورية = Aselway n Tegduda\n"
        active_dict_context += "- Le président = Aselway\n"
        active_dict_context += "- ALGER = DZAYER TAMANEƔT\n"
        active_dict_context += "- Alger = DZAYER TAMANEƔT\n"
        active_dict_context += "- Conseil des ministres = Aseqqamu n Yineɣlaf\n"
        active_dict_context += "- réunion = Timlilt\n"
        active_dict_context += "- gouvernement = Anabaḍ\n"
        
        if st.session_state["lexicon_data"]:
            for k, v in st.session_state["lexicon_data"].items():
                if k in text_to_translate.lower():
                    active_dict_context += f"- {k} = {v}\n"

        tm_context = ""
        if tm_data:
            matched_pairs = []
            for pair in tm_data:
                if any(word in pair["foreign"].lower() for word in search_words):
                    matched_pairs.append(pair)
                if len(matched_pairs) >= 10:
                    break
            if not matched_pairs:
                matched_pairs = tm_data[:5]
                
            for i, pair in enumerate(matched_pairs, 1):
                tm_context += f"نموذج مرجعي رسمي {i}:\nالنص الفرنسي الأصلي: {pair['foreign']}\nالترجمة الأمازيغية المعتمدة: {pair['tamazight']}\n---\n"

        instruction = """أنت رئيس تحرير ومترجم رسمي صارم للغة الأمازيغية المعيارية لصالح وكالة الأنباء (APS).
        مهمتك الحتمية والمقدسة: صياغة وترجمة البرقية المدخلة إلى الأمازيغية المعيارية الصارمة بالحرف اللاتيني.
        التزم التزاماً عسكرياً صارماً بالمعجم والنماذج المرفقة المستخرجة من الوورد. يُحظر تماماً التخمين أو الابتكار.
        يُمنع منعاً باتاً استخدام أي حروف عربية في النص الأمازيغي النهائي. استخدم دائماً DZAYER TAMANEƔT لـ ALGER و Aselway n Tegduda لـ رئيس الجمهورية."""

        system_prompt = f"{instruction}\n\nالقواعد المورفولوجية المعتمدة للوكالة:\n{active_dict_context}\n\nالذاكرة المرجعية المعتمدة:\n{tm_context}\n{st.session_state['live_corrections']}"

        # الاتصال الفوري والمستقر بالمحرك المفتوح اللامحدود لتفادي حصار غوغل
        url = "https://pollinations.ai"
        payload = {
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": text_to_translate}
            ],
            "model": "openai-large",
            "temperature": 0.0
        }
        
        response = requests.post(url, json=payload, timeout=30)

        if response.status_code == 200 and response.text.strip():
            output = response.text.strip()
            # المصفي الآلي لطرد أي كلمات مشوهة قديمة من الشاشة نهائياً
            output = re.sub(r'\banmazul\b', 'Aselway', output, flags=re.IGNORECASE)
            output = re.sub(r'\banemhal\b', 'Aselway', output, flags=re.IGNORECASE)
            output = re.sub(r'yettu[εe]zlen', 'i yettwaheggan', output)
            output = re.sub(r'[\u0600-\u06FF]+', '', output).replace("  ", " ").strip()

            st.success("تمت الترجمة بنجاح واكتملت صياغة الخبر بناءً على قواعد الوورد والمحرك الحر الجديد.")
            st.text_area("الترجمة الأمازيغية المعيارية النهائية (أسلوب APS رسمي ونظيف 100%):", value=output, height=280)
        else:
            st.error("الخادم مستغرق في معالجة المستند، يرجى تكرار المحاولة الآن سريعا.")

    except Exception as error:
        st.error(f"حدث خطأ تقني أثناء الاتصال بالمحرك الحر: {error}")
