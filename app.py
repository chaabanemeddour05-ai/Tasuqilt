import streamlit as st
import google.generativeai as genai
from docx import Document
import os

st.set_page_config(page_title="Tasuqqilt", layout="wide")

st.title("Tasuqqilt 🇩🇿")
st.subheader("مترجم الأمازيغية المعيارية الصارم")

# إدخال مفتاح الـ API بشكل آمن في الواجهة
api_key = st.sidebar.text_input("أدخل مفتاح Google API الخاص بك:", type="password")

# اختيار وضع الترجمة لتغيير درجة الحرارة ديناميكياً
mode = st.sidebar.radio(
    "اختر وضع الترجمة المناسب للبرقية:",
    ("إخباري رسمي وصارم (الحرارة 0.0)", "أدبي/ثقافي مرن (الحرارة 0.7)")
)

# تحديد الإعدادات بناءً على الاختيار
if mode == "إخباري رسمي وصارم (الحرارة 0.0)":
    chosen_temp = 0.0
    mode_instruction = "التزم حرفياً بالبيانات المسترجعة والنحو المرفق ولا تبتكر شيئاً إطلاقاً."
else:
    chosen_temp = 0.7
    mode_instruction = "يمكنك صياغة الأمازيغية بمرونة وجمالية لتلائم الأسلوب الأدبي والثقافي."

# قراءة قواميس الأمازيغية من ملف الوورد الصغير المرفق في المجلد
@st.cache_data
def load_tamazight_context():
    file_name = "database.docx"
    if not os.path.exists(file_name):
        return ""
    
    doc = Document(file_name)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text.strip())
    return "\n".join(full_text)

context_data = load_tamazight_context()

# مربع النص للمدخلات
text_to_translate = st.text_area("أدخل البرقية المراد ترجمتها (بالعربية أو الفرنسية):", height=150)

if st.button("بدء الترجمة الاحترافية"):
    if not api_key:
        st.error("⚠️ يرجى إدخال مفتاح Google API في القائمة الجانبية لتفعيل المترجم.")
    elif not context_data and not os.path.exists("database.docx"):
        st.error("⚠️ لم يتم العثور على ملف 'database.docx' في المجلد. يرجى إضافته.")
    else:
        try:
            genai.configure(api_key=api_key)
            
            system_prompt = f"""أنت مترجم رسمي لوكالة الأنباء الجزائرية. 
            لديك ملف مرجعي يحتوي على نصوص وقواعد باللغة الأمازيغية المعيارية الصرفة بالحرف اللاتيني.
            السياق المرجعي المتاح لك هو:
            ---
            {context_data[:15000]}
            ---
            تعليمات الصياغة الحالية: {mode_instruction}
            التزم تماماً بالقواعد والأسلوب المتبع في المرجع المرفق، ولا تبتكر تعبيرات خارجة عن هذا النطاق."""
            
            model = genai.GenerativeModel(
                model_name="gemini-1.5-flash",
                generation_config={"temperature": chosen_temp},
                system_instruction=system_prompt
            )
            
            with st.spinner("⏳ جاري صياغة الترجمة بدقة..."):
                response = model.generate_content(text_to_translate)
                
            st.success("✅ تمت الترجمة بنجاح:")
            st.code(response.text, language="text")
            
        except Exception as e:
            st.error(f"حدث خطأ أثناء الاتصال بالذكاء الاصطناعي: {e}")
