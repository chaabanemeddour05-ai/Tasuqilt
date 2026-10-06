import os
import streamlit as st
from google import genai
from docx import Document

st.set_page_config(
    page_title="Tasuqilt",
    page_icon="🇩🇿",
    layout="wide"
)

st.title("Tasuqilt 🇩🇿")
st.subheader("مترجم الأمازيغية المعيارية الفوري والصارم")

# جلب مفتاح الـ API تلقائياً وبشكل آمن من إعدادات المنصة
api_key = os.environ.get("GEMINI_API_KEY")

if not api_key:
    # محاولة جلبها من إعدادات سريمليت السرية الاحتياطية إذا لم تكن في البيئة
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    else:
        st.error("لم يتم العثور على GEMINI_API_KEY في Streamlit Secrets.")
        st.stop()

# تنظيف المفتاح من أي علامات تنصيص زائدة أثناء اللصق لضمان الاتصال السليم
api_key = api_key.replace('"', '').replace("'", "").strip()
client = genai.Client(api_key=api_key)

# قراءة قاعدة البيانات الأسلوبية من ملف الوورد المرفق في المستودع
@st.cache_data
def load_tamazight_corpus():
    file_name = "database.docx"
    if not os.path.exists(file_name):
        return []
    try:
        doc = Document(file_name)
        full_text = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text and len(text) > 5:
                full_text.append(text)
        return full_text
    except Exception:
        return []

corpus_data = load_tamazight_corpus()

mode = st.sidebar.radio(
    "اختر وضع الترجمة:",
    ["إخباري رسمي وصارم", "أدبي / ثقافي"]
)

if mode == "إخباري رسمي وصارم":
    instruction = """أنت رئيس تحرير ومترجم رسمي متخصص في الأمازيغية المعيارية لصالح وكالة الأنباء.
ترجم النص إلى الأمازيغية المعيارية بالحرف اللاتيني.
التزم حرفياً بأسلوب وكالة الأنباء الرسمية (APS) والنحو المرفق في العينات أدناه، ولا تبتكر صياغات خارجة عنها.
لا تضف أو تحذف أي معلومة، حافظ على الأسماء والأرقام والتواريخ، ولا تشرح الترجمة؛ أعطني النص الأمازيغي فقط."""
else:
    instruction = """أنت مترجم محترف متخصص في الأمازيغية المعيارية للنصوص الثقافية.
ترجم النص إلى الأمازيغية المعيارية بالحرف اللاتيني بأسلوب طبيعي وسلس ومناسب للنص الأدبي بناءً على العينات المرفقة أدناه.
حافظ على المعنى الكامل والأسماء والأرقام، ولا تشرح الترجمة؛ أعطني النص الأمازيغي فقط."""

text_to_translate = st.text_area(
    "أدخل النص بالعربية أو الفرنسية:",
    height=220,
    placeholder="اكتب النص هنا..."
)

if st.button("بدء الترجمة", type="primary"):
    if not text_to_translate.strip():
        st.warning("يرجى إدخال نص أولًا.")
        st.stop()

    st.info("جاري فحص الذاكرة اللغوية والترجمة بواسطة Gemini 3.5 Flash-Lite...")

    try:
        # الفلترة الذكية الموضعية لملف الوورد لاستخلاص المصطلحات المطابقة للبرقية فقط لضمان السرعة الفائقة
        search_words = [w.strip().lower() for w in text_to_translate.split() if len(w.strip()) > 3]
        matched_style = []
        for paragraph in corpus_data:
            if any(word in paragraph.lower() for word in search_words):
                matched_style.append(paragraph)
            if len(matched_style) >= 12: # استدعاء أفضل 12 عينة متطابقة لتقليص حجم البيانات
                break
        
        if not matched_style:
            matched_style = corpus_data[:10]
            
        style_context = "\n".join(matched_style)

        # دمج الـ Prompt مع سياق ملف الوورد المفلتر بدقة
        prompt = f"""{instruction}

        العينات الأسلوبية المعتمدة المسترجعة من ملفك المرجعي:
        ---
        {style_context}
        ---

        النص الأصلي المراد ترجمته الآن:
        {text_to_translate}

        الترجمة الأمازيغية المعتمدة:"""

        response = client.models.generate_content(
            model="gemini-3.5-flash-lite",
            contents=prompt
        )

        result = response.text

        if result and result.strip():
            st.success("تمت الترجمة بنجاح.")
            st.text_area(
                "الترجمة الأمازيغية:",
                value=result.strip(),
                height=280
            )
        else:
            st.error("لم يُرجع Gemini نصًا.")

    except Exception as error:
        st.error("حدث خطأ أثناء الاتصال بـ Gemini.")
        st.code(str(error))
