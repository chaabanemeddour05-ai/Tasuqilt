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

# تهيئة الذاكرة التفاعلية للمعجم الضخم والتصحيحات الحية
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
    st.error("لم يتم العثور على مفتاح الأمان السري GEMINI_API_KEY في الإعدادات الخلفية.")
    st.stop()

api_key = api_key.replace('"', '').replace("'", "").strip()
client = genai.Client(api_key=api_key)

# قراءة المرجع المقدس للفقرات المترجمة tm.docx
@st.cache_data
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
        return memory_pairs
    except Exception:
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
    
    # لوحة حقن المعجم الضخم بالنسخ واللصق المباشر كما هو!
    st.sidebar.subheader("📚 حقن المعجم الشامل (نسخ ولصق)")
    raw_lexicon_text = st.sidebar.text_area("أصق قائمة الكلمات كما هي هنا (سيتم تفكيكها آلياً):", height=200, placeholder="مثال:\nAbabat (ibabaten), père (fr) – father (en) - أب")
    
    if st.sidebar.button("⚙️ تفكيك وحقن القاموس في عقل المنصة"):
        if raw_lexicon_text.strip():
            count = 0
            lines = raw_lexicon_text.strip().split("\n")
            for line in lines:
                if not line.strip():
                    continue
                # استخراج الكلمة الأمازيغية (أول كلمة قبل القوس أو الفاصلة)
                amazigh_match = re.match(r"^([a-zA-ZɛɣɣƐƔƔ]+)", line.strip())
                if amazigh_match:
                    amazigh_word = amazigh_match.group(1).strip()
                    
                    # استخراج الكلمات الفرنسية والعربية المرافقة بالبحث عن الأنماط
                    french_parts = re.findall(r"([^,\(\)\-\–]+)\s*\(fr\)", line.lower())
                    arabic_parts = line.split("-")[-1].split("—")[-1].strip()
                    
                    # ربط الكلمات بالفرنسية
                    for fr in french_parts:
                        for sub_fr in fr.split(","):
                            clean_fr = sub_fr.strip().lower()
                            if clean_fr:
                                st.session_state["lexicon_data"][clean_fr] = amazigh_word
                    
                    # ربط الكلمات بالعربية
                    for ar in arabic_parts.split(","):
                        clean_ar = ar.strip()
                        if clean_ar and not re.match(r'^[a-zA-Z]', clean_ar):
                            st.session_state["lexicon_data"][clean_ar] = amazigh_word
                    count += 1
            st.sidebar.success(f"✅ تم بنجاح تفكيك وحقن {count} مصطلحاً معجمياً في الذاكرة الحية للموقع!")

    st.sidebar.markdown("---")
    st.sidebar.subheader("✍️ لوحة تصحيح الفقرات")
    bad_french = st.sidebar.text_area("النص الأصلي (الفرنسي/العربي) لتصحيحه:")
    good_tamazight = st.sidebar.text_area("الصياغة الأمازيغية المثالية المعتمدة:")
    
    if st.sidebar.button("⚙️ تثبيت التصحيح اللحظي"):
        if bad_french.strip() and good_tamazight.strip():
            new_correction = f"\nنموذج مصحح يدوياً:\nالأصل: {bad_french.strip()}\nالترجمة المفروضة: {good_tamazight.strip()}\n---\n"
            st.session_state["live_corrections"] += new_correction
            st.sidebar.success("✅ تم التثبيت!")

if mode == "إخباري رسمي وصارم":
    chosen_temp = 0.0
    instruction = """أنت رئيس تحرير ومترجم رسمي صارم للغة الأمازيغية المعيارية لصالح وكالة الأنباء (APS).
مهمتك الحتمية والمقدسة: صياغة وترجمة البرقية المدخلة إلى الأمازيغية المعيارية الصارمة بالحرف اللاتيني.

لقد تم تزويدك بـ "معجم الكلمات الإعلامية المعياري الحاسم" و "المرجع المقدس لنماذج الترجمة".
يجب أن تستقي المصطلحات والكلمات بدقة وبشكل إجباري وحصري من المعجم المرفق أولاً ثم النماذج ثانياً.
إذا وجدت أي كلمة متطابقة في المعجم المرفق، استخدم المقابل الأمازيغي لها فوراً ولا تغيره إطلاقاً. النص النهائي يجب أن يكون نقياً 100% بالحرف اللاتيني."""
else:
    chosen_temp = 0.7
    instruction = """أنت مترجم محترف نصوص ثقافية وأدبية. ترجم بدقة وسلاسة بالحرف اللاتيني فقط."""

text_to_translate = st.text_area(
    "أدخل البرقية الصحفية المراد ترجمتها (بالفرنسية أو العربية):",
    height=220,
    placeholder="ضع نص البرقية هنا..."
)

if st.button("بدء الترجمة الاحترافية الموحدة", type="primary"):
    if not text_to_translate.strip():
        st.warning("يرجى إدخال نص أولًا للبدء.")
        st.stop()

    st.info("جاري فحص المعجم التفاعلي والذاكرة الموحدة...")

    try:
        search_words = [w.strip().lower() for w in text_to_translate.split() if len(w.strip()) > 4]
        
        # استخراج الكلمات المطابقة للمعجم المفرغ حياً
        active_dict_context = ""
        if st.session_state["lexicon_data"]:
            active_dict_context = "\n⚠️ معجم الكلمات الإعلامية الحتمي والإلزامي للترجمة الحالية:\n"
            # فحص الكلمات الفرنسية أو العربية المدخلة في النص
            clean_text_words = [w.strip(",.()\"'-:;!?").lower() for w in text_to_translate.split()]
            for clean_w in clean_text_words:
                if clean_w in st.session_state["lexicon_data"]:
                    active_dict_context += f"- الكلمة الأصلية: {clean_w} = المقابل الأمازيغي الإجباري: {st.session_state['lexicon_data'][clean_w]}\n"

        # استرجاع الفقرات المشابهة للبرقية من ملف tm.docx
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
            tm_context += f"نموذج ترجمة معتمد ومقدس {i}:\nالنص الأصلي: {pair['foreign']}\nالترجمة الأمازيغية الرسمية المفروضة: {pair['tamazight']}\n---\n"

        config = types.GenerateContentConfig(
            temperature=chosen_temp,
            system_instruction=instruction
        )

        prompt = f"""{active_dict_context}
        
        المرجع المقدس الحصري والوحيد (نماذج الترجمة المعتمدة لوكالة الأنباء):
        {tm_context}
        {st.session_state["live_corrections"]}

        النص الجديد المراد صياغته وترجمته الآن بدقة بالغة وبناءً على المعجم الحتمي المرفق:
        {text_to_translate}

        الترجمة الأمازيغية الرسمية الصارمة والنهائية (حرف لاتيني فقط):"""

        response = client.models.generate_content(
            model="gemini-3.5-flash-lite",
            contents=prompt,
            config=config
        )

        result = response.text

        if result and result.strip():
            st.success("تمت الترجمة بنجاح واكتملت صياغة الخبر بناءً على المعجم الذكي المفرغ والذاكرة الموحدة.")
            st.text_area("الترجمة الأمازيغية المعيارية المعتمدة (نقية ومحدثة 100%):", value=result.strip(), height=280)
        else:
            st.error("لم ينجح النظام في معالجة النص، يرجى إعادة المحاولة.")

    except Exception as error:
        st.error("حدث خطأ تقني أثناء الاتصال بالذكاء الاصطناعي.")
        st.code(str(error))
