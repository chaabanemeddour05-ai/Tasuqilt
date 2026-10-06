import os
import streamlit as st
from google import genai
from google.genai import types
from docx import Document

st.set_page_config(
    page_title="Tasuqilt Enterprise",
    page_icon="🇩🇿",
    layout="wide"
)

# 🔐 إعدادات أمان لوحة التحكم (يمكنك أنت وصديقك تغيير هذا الرمز السري من الكود في أي وقت)
ADMIN_PASSWORD = "APS_Tasuqilt_2026"

# تهيئة الذاكرة التفاعلية اللحظية لتخزين التصحيحات والمصطلحات الجديدة أثناء تمرين الموقع من المتصفح
if "live_corrections" not in st.session_state:
    st.session_state["live_corrections"] = ""
if "live_dictionary" not in st.session_state:
    st.session_state["live_dictionary"] = {}

st.title("Tasuqilt 🇩🇿")
st.subheader("منصة إدارة الترجمة الإعلامية والأمازيغية المعيارية الصارمة")

# جلب مفتاح الـ API تلقائياً وبشكل آمن من إعدادات المنصة السرية
api_key = os.environ.get("GEMINI_API_KEY")
if not api_key and "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]

if not api_key:
    st.error("لم يتم العثور على مفتاح الأمان السري GEMINI_API_KEY في الإعدادات الخلفية.")
    st.stop()

api_key = api_key.replace('"', '').replace("'", "").strip()
client = genai.Client(api_key=api_key)

# 📚 المرجع المقدس: قراءة القاموس المزدوج المفصول بـ @ بخفة واستقرار كامل
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

# 🛡️ البوابة الجانبية: خيارات الترجمة وبوابة الإدارة السرية
st.sidebar.header("🎛️ خيارات التحكم")
mode = st.sidebar.radio(
    "اختر وضع الترجمة الصحفية المعتمد:",
    ["إخباري رسمي وصارم", "أدبي / ثقافي"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("🔐 بوابة الإدارة السرية (خاصة بالمشرفين)")
admin_input = st.sidebar.text_input("أدخل رمز الدخول لتمرين وتحديث المنصة:", type="password")

is_admin = (admin_input == ADMIN_PASSWORD)

if is_admin:
    st.sidebar.success("🔓 تم فتح صلاحيات الإدارة والتمرين المباشر!")
    
    # 📌 القسم الأول: لوحة تصحيح وتمرين الفقرات المشوشة مباشرة من الموقع
    st.sidebar.subheader("✍️ لوحة تصحيح الفقرات وتمرين الموقع")
    bad_french = st.sidebar.text_area("النص الأصلي (الفرنسي/العربي) الذي تريد تصحيحه:")
    good_tamazight = st.sidebar.text_area("الصياغة الأمازيغية المثالية المعيارية المعتمدة:")
    
    if st.sidebar.button("⚙️ حقن وتثبيت التصحيح في عقل المنصة"):
        if bad_french.strip() and good_tamazight.strip():
            new_correction = f"\nنموذج مصحح يدوياً من رئيس التحرير:\nالأصل: {bad_french.strip()}\nالترجمة المقدسة المفروضة: {good_tamazight.strip()}\n---\n"
            st.session_state["live_corrections"] += new_correction
            st.sidebar.success("✅ تم حفظ وتثبيت التصحيح بنجاح! لن يكرر الخطأ القديم.")
        else:
            st.sidebar.warning("يرجى ملء الخانتين أولاً.")
            
    # 📌 القسم الثاني: لوحة المعجم وقاموس المصطلحات الفردية للإعلام
    st.sidebar.markdown("---")
    st.sidebar.subheader("📖 معجم المصطلحات الإعلامية الفورية")
    dict_word = st.sidebar.text_input("المصطلح الأصلي (مثال: réunion أو اجتماع):")
    dict_translation = st.sidebar.text_input("المقابل الأمازيغي المعياري الحصري (مثال: timlilt):")
    
    if st.sidebar.button("📌 إدراج المصطلح في القاموس المقدس"):
        if dict_word.strip() and dict_translation.strip():
            st.session_state["live_dictionary"][dict_word.strip().lower()] = dict_translation.strip()
            st.sidebar.success(f"✅ تم قفل المصطلح: '{dict_word}' = '{dict_translation}'")
else:
    if admin_input:
        st.sidebar.error("❌ رمز الدخول غير صحيح. الصلاحيات مغلقة.")

# 📝 إعداد تعليمات التحرير الصارمة لـ Gemini
if mode == "إخباري رسمي وصارم":
    chosen_temp = 0.0
    instruction = """أنت رئيس تحرير ومترجم رسمي صارم للغة الأمازيغية المعيارية لصالح وكالة الأنباء (APS).
مهمتك الحتمية والمقدسة: صياغة وترجمة البرقية المدخلة إلى الأمازيغية المعيارية الصارمة بالحرف اللاتيني.

لقد تم تزويدك بـ "المرجع المقدس لنماذج الترجمة المقابلة المعتمدة" بالأسفل. 
يجب أن تستقي طريقة بناء الجمل، النحو، الصرف، والمصطلحات الحديثة بالكامل وحصرياً من هذا المرجع فقط.
إذا زودك رئيس التحرير بـ "تعديلات وتصحيحات مباشرة" أو "قاموس مصطلحات حاسم"، التزم بها كأولوية قصوى فوق أي شيء آخر.
يُمنع منعاً باتاً استخدام مصطلحات بدائية قديمة. النص النهائي يجب أن يكون أمازيغياً لاتينياً إعلامياً ناصعاً ونقياً 100% من الحروف العربية.
لا تضف أو تحذف أي معلومة، حافظ على الأرقام والتواريخ، وأعطني النص الأمازيغي اللاتيني فقط دون أي شرح."""
else:
    chosen_temp = 0.7
    instruction = """أنت مترجم محترف متخصص في الأمازيغية المعيارية للنصوص الثقافية والأدبية.
ترجم النص إلى الأمازيغية بأسلوب طبيعي وسلس بناءً على العينات المرفقة. أعطني النص الأمازيغي فقط بالحرف اللاتيني."""

# 🎚️ واجهة الترجمة الرئيسية لجميع المستخدمين
text_to_translate = st.text_area(
    "أدخل البرقية الصحفية المراد ترجمتها (بالفرنسية أو العربية):",
    height=220,
    placeholder="ضع نص البرقية هنا..."
)

if st.button("بدء الترجمة الاحترافية الموحدة", type="primary"):
    if not text_to_translate.strip():
        st.warning("يرجى إدخال نص أولًا للبدء.")
        st.stop()

    st.info("جاري استدعاء محرك الذاكرة الموحدة المحدث وتطبيق معجم التدريب الحصري...")

    try:
        search_words = [w.strip().lower() for w in text_to_translate.split() if len(w.strip()) > 4]
        
        # البحث الموضعي الفوري الفائق السرعة داخل "المرجع المقدس" (tm.docx)
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

        # دمج قاموس المصطلحات الفورية التي تم إدخالها من لوحة التحكم السرية
        custom_dict_context = ""
        if st.session_state["live_dictionary"]:
            custom_dict_context = "\n⚠️ قاموس المصطلحات الحاسم والإلزامي الصادر من الإدارة:\n"
            for k, v in st.session_state["live_dictionary"].items():
                custom_dict_context += f"- المصطلح: {k} = المقابل الإجباري: {v}\n"

        config = types.GenerateContentConfig(
            temperature=chosen_temp,
            system_instruction=instruction
        )

        # دمج سياق الذاكرة الثابتة مع التصحيحات الحية المباشرة والمصطلحات المحدثة
        prompt = f"""المرجع المقدس الحصري والوحيد (نماذج الترجمة المعتمدة لوكالة الأنباء):
        {tm_context}
        {st.session_state["live_corrections"]}
        {custom_dict_context}

        النص الجديد المراد صياغته وترجمته الآن بدقة بالغة وبنفس الأسلوب والمصطلحات الحديثة:
        {text_to_translate}

        الترجمة الأمازيغية الرسمية الصارمة والنهائية (حرف لاتيني فقط):"""

        response = client.models.generate_content(
            model="gemini-3.5-flash-lite",
            contents=prompt,
            config=config
        )

        result = response.text

        if result and result.strip():
            st.success("تمت الترجمة بنجاح واكتملت صياغة الخبر بناءً على المرجعية الموحدة الجديدة الممرنة.")
            st.text_area("الترجمة الأمازيغية المعيارية المعتمدة (نقية ومحدثة 100%):", value=result.strip(), height=280)
        else:
            st.error("لم ينجح النظام في معالجة النص، يرجى إعادة المحاولة.")

    except Exception as error:
        st.error("حدث خطأ تقني أثناء الاتصال بالذكاء الاصطناعي.")
        st.code(str(error))
