import streamlit as st
import google.generativeai as genai
from docx import Document
import os

st.set_page_config(page_title="Tasuqqilt", layout="wide")

st.title("Tasuqqilt 🇩🇿")
st.subheader("مترجم الأمازيغية المعيارية الفوري والصارم")

# جلب مفتاح الـ API تلقائياً وبشكل سري وآمن من الخزنة الخلفية للمنصة
if "GEMINI_API_KEY" in st.secrets:
    raw_key = st.secrets["GEMINI_API_KEY"]
    api_key = raw_key.replace('"', '').replace("'", "").strip()
else:
    api_key = st.sidebar.text_input("أدخل مفتاح Google API الخاص بك:", type="password")

# اختيار وضع الترجمة لتغيير درجة الحرارة ديناميكياً
mode = st.sidebar.radio(
    "اختر وضع الترجمة المناسب للبرقية:",
    ("إخباري رسمي وصارم (الحرارة 0.0)", "أدبي/ثقافي مرن (الحرارة 0.7)")
)

if mode == "إخباري رسمي وصارم (الحرارة 0.0)":
    chosen_temp = 0.0
    mode_instruction = "التزم حرفياً بأسلوب وكالة الأنباء (APS)، والنحو المسترجع، ولا تبتكر شيئاً إطلاقاً."
else:
    chosen_temp = 0.7
    mode_instruction = "يمكنك صياغة الأمازيغية بمرونة وجمالية لتلائم الأسلوب الأدبي والثقافي المفتوح."

# قراءة قاعدة البيانات من ملف الوورد المرفق
@st.cache_data
def load_tamazight_context():
    file_name = "database.docx"
    if not os.path.exists(file_name):
        return []
    
    doc = Document(file_name)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text.strip())
    return full_text

context_paragraphs = load_tamazight_context()

# مربع النص للمدخلات الصحفية
text_to_translate = st.text_area("أدخل البرقية المراد ترجمتها (بالعربية أو الفرنسية):", height=150)

if st.button("بدء الترجمة الاحترافية"):
    if not api_key:
        st.error("⚠️ يرجى التأكد من حقن مفتاح Google API في إعدادات Secrets لتفعيل المنصة.")
    elif not context_paragraphs and not os.path.exists("database.docx"):
        st.error("⚠️ لم يتم العثور على ملف 'database.docx' كمدونة أسلوبية في المجلد.")
    elif not text_to_translate.strip():
        st.warning("⚠️ يرجى كتابة أو لصق نص البرقية أولاً.")
    else:
        try:
            with st.spinner("⏳ جاري استدعاء الذاكرة الترجمية وصياغة الخبر فوراً..."):
                # الفلترة الذكية للذاكرة الترجمية: استخلاص الكلمات المفتاحية
                words = [w.strip() for w in text_to_translate.lower().split() if len(w.strip()) > 3]
                relevant_lines = []
                
                for p in context_paragraphs:
                    if any(word in p.lower() for word in words):
                        relevant_lines.append(p)
                    if len(relevant_lines) >= 20: # استدعاء أفضل 20 سياقاً متطابقاً لضمان السرعة الفائقة
                        break
                
                if not relevant_lines:
                    relevant_lines = context_paragraphs[:15]
                
                refined_context = "\n".join(relevant_lines)
                
                # بناء التوجيه الاحترافي الصارم المدمج
                system_prompt = f"""أنت رئيس تحرير ومترجم رسمي معتمد للغة الأمازيغية المعيارية الصرفة بالحرف اللاتيني لصالح وكالة الأنباء.
                لديك مدونة أسلوبية ونحوية مسترجعة من الأرشيف الرسمي وهي:
                ---
                {refined_context}
                ---
                تعليمات التحرير الحالية: {mode_instruction}
                مهمتك: صياغة البرقية المدخلة بلغة أمازيغية إعلامية رصينة ومطابقة تماماً للقواعد المرفقة أعلاه دون زيادة أو ابتكار خارجي."""
                
                # الاتصال بمحرك غوغل المستقر والمعتمد لعام 2026
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(
                    model_name="gemini-3.8-flash",
                    generation_config={"temperature": chosen_temp},
                    system_instruction=system_prompt
                )
                
                response = model.generate_content(text_to_translate)
                
                st.success("✅ تمت الترجمة الفورية بنجاح:")
                st.code(response.text, language="text")
                
        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة الاتصال بالذكاء الاصطناعي: {e}")
