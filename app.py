import os
import streamlit as st
from google import genai
from google.genai import types
from docx import Document

st.set_page_config(
    page_title="Tasuqilt",
    page_icon="🇩🇿",
    layout="wide"
)

st.title("Tasuqilt 🇩🇿")
st.subheader("مترجم الأمازيغية المعيارية الصارم (بمحرك الذاكرة المزدوجة المتكاملة)")

# جلب مفتاح الـ API تلقائياً وبشكل آمن من إعدادات المنصة السرية
api_key = os.environ.get("GEMINI_API_KEY")
if not api_key and "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]

if not api_key:
    st.error("لم يتم العثور على مفتاح الأمان السري في الإعدادات.")
    st.stop()

# تنظيف المفتاح لضمان سلامة الاتصال الفوري بخوادم غوغل
api_key = api_key.replace('"', '').replace("'", "").strip()
client = genai.Client(api_key=api_key)

# 1. المدرسة اللغوية: قراءة المدونة أحادية اللغة لتعلم أسلوب الصرف والنحو والتركيب الصحفي
@st.cache_data
def load_monolingual_corpus():
    file_name = "database.docx"
    if not os.path.exists(file_name):
        return []
    try:
        doc = Document(file_name)
        text_lines = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text and len(text) > 10 and "@" not in text:
                text_lines.append(text)
        return text_lines
    except Exception:
        return []

# 2. المرجع المقدس: قراءة الذاكرة الترجمية الثنائية (المفصولة بـ @) لربط المصطلحات بدقة خارقة
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

mono_corpus = load_monolingual_corpus()
tm_data = load_translation_memory()

mode = st.sidebar.radio(
    "اختر وضع الترجمة الصحفية المعتمد:",
    ["إخباري رسمي وصارم", "أدبي / ثقافي"]
)

if mode == "إخباري رسمي وصارم":
    chosen_temp = 0.0
    instruction = """أنت رئيس تحرير ومترجم رسمي صارم للغة الأمازيغية المعيارية لصالح وكالة الأنباء (APS).
مهمتك: صياغة وترجمة البرقية المدخلة إلى الأمازيغية المعيارية الصارمة بالحرف اللاتيني.

لقد تم تزويدك بمصدرين أساسيين مقدسين يجب أن تبني عليهما ترجمتك بالكامل:
المصدر الأول والأهم (المرجع المقدس للترجمة): هو "نماذج الترجمة المقابلة المعتمدة" المرفقة في الأسفل، التزم بحقن الكلمات والمصطلحات المقابلة منها حرفياً.
المصدر الثاني (المدرسة النحوية): هي "عينات الصياغة والصرف والتركيب النحوي"، وظيفتها أن تتعلم منها النحو، وصرف الأفعال، وطريقة صياغة العبارات لتبدو صحفية رسمية.

التزم حرفياً بالمعلومات، لا تزد ولا تنقص، حافظ على الأرقام والتواريخ، وأعطني النص الأمازيغي اللاتيني فقط دون أي شرح أو تعليق."""
else:
    chosen_temp = 0.7
    instruction = """أنت مترجم محترف متخصص في الأمازيغية المعيارية للنصوص الثقافية والأدبية.
ترجم النص إلى الأمازيغية بأسلوب طبيعي، سلس، وجذاب بناءً على العينات الترجمية والأسلوبية المرفقة أدناه.
أعطني النص الأمازيغي فقط دون أي إضافات تفسيرية."""

text_to_translate = st.text_area(
    "أدخل البرقية الصحفية المراد ترجمتها (بالفرنسية أو العربية):",
    height=200,
    placeholder="ضع النص هنا..."
)

if st.button("بدء الترجمة الاحترافية المدمجة", type="primary"):
    if not text_to_translate.strip():
        st.warning("يرجى إدخال نص أولًا للبدء.")
        st.stop()

    st.info("جاري استرجاع السياق من المرجع المقدس والمدرسة اللغوية...")

    try:
        # استخلاص الكلمات المفتاحية للبحث الذكي السريع
        search_words = [w.strip().lower() for w in text_to_translate.split() if len(w.strip()) > 4]
        
        # أ) البحث الموضعي الفوري داخل "المرجع المقدس" (tm.docx) واستخراج المتطابقات
        matched_pairs = []
        if tm_data:
            for pair in tm_data:
                if any(word in pair["foreign"].lower() for word in search_words):
                    matched_pairs.append(pair)
                if len(matched_pairs) >= 5:  # استدعاء أفضل 5 أمثلة مترجمة للسرعة الفائقة
                    break
        if not matched_pairs:
            matched_pairs = tm_data[:3] if tm_data else []

        # ب) البحث الموضعي الفوري داخل "المدرسة النحوية" (database.docx) لتعلم النحو والتركيب المتطابق
        matched_mono = []
        if mono_corpus:
            sample_tamazight_words = []
            for pair in matched_pairs:
                sample_tamazight_words.extend([w.lower() for w in pair["tamazight"].split() if len(w) > 4])
            
            for line in mono_corpus:
                if any(word in line.lower() for word in sample_tamazight_words) or any(word in line.lower() for word in search_words):
                    matched_mono.append(line)
                if len(matched_mono) >= 6:  # استدعاء أفضل 6 عينات صرف ونحو
                    break
        if not matched_mono:
            matched_mono = mono_corpus[:4] if mono_corpus else []

        # بناء السياق المشفر التوأم بدقة متناهية لـ Gemini
        tm_context = ""
        for i, pair in enumerate(matched_pairs, 1):
            tm_context += f"نموذج ترجمة معتمد {i}:\nالنص الأصلي: {pair['foreign']}\nالترجمة الأمازيغية المقدسة: {pair['tamazight']}\n---\n"
            
        mono_context = "\n".join(matched_mono)

        # صياغة الـ Prompt والـ System Instruction بشكل متوافق تماماً مع مكتبة غوغل الحديثة لعام 2026
        config = types.GenerateContentConfig(
            temperature=chosen_temp,
            system_instruction=instruction
        )

        prompt = f"""المصدر الأول (المرجع المقدس لنماذج الترجمة المقابلة المعتمدة):
        {tm_context}

        المصدر الثاني (عينات الصياغة والصرف والتركيب النحوي الأحادية):
        {mono_context}

        النص الجديد المراد صياغته وترجمته الآن بدقة بالغة:
        {text_to_translate}

        الترجمة الأمازيغية الرسمية الصارمة والنهائية (حرف لاتيني):"""

        # استدعاء المحرك الخفيف الفوري والمستقر
        response = client.models.generate_content(
            model="gemini-3.5-flash-lite",
            contents=prompt,
            config=config
        )

        result = response.text

        if result and result.strip():
            st.success("تمت الترجمة بنجاح واكتملت صياغة الخبر بناءً على المرجعية المزدوجة.")
            st.text_area("الترجمة الأمازيغية المعيارية النهائية:", value=result.strip(), height=250)
        else:
            st.error("لم ينجح النظام في معالجة النص، يرجى إعادة المحاولة.")

    except Exception as error:
        st.error("حدث خطأ تقني أثناء الاتصال بالذكاء الاصطناعي.")
        st.code(str(error))
