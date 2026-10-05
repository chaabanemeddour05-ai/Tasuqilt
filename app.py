import streamlit as st
from docx import Document
import os
import requests

st.set_page_config(page_title="Tasuqqilt", layout="wide")

st.title("Tasuqqilt 🇩🇿")
st.subheader("مترجم الأمازيغية المعيارية الفوري والصارم")

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

# قراءة قواميس الأمازيغية من ملف الوورد المرفق في المجلد
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
    if not context_data and not os.path.exists("database.docx"):
        st.error("⚠️ لم يتم العثور على ملف 'database.docx' في المجلد. يرجى إضافته.")
    elif not text_to_translate.strip():
        st.warning("⚠️ يرجى إدخال نص للترجمة.")
    else:
        try:
            with st.spinner("⏳ جاري صياغة الترجمة الفورية بدقة..."):
                system_prompt = f"أنت مترجم رسمي للغة الأمازيغية المعيارية الصرفة. السياق المرجعي المتاح لك من ملفك هو:\n{context_data[:8000]}\nتعليمات الصياغة الحالية: {mode_instruction} التزم تماماً بالقواعد والأسلوب المتبع في المرجع المرفق ولا تخرج عنه."
                
                # استخدام خادم إطلاق فوري متطور وخفيف يتجاوز قيود غوغل المعقدة
                url = "https://pollinations.ai"
                payload = {
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": text_to_translate}
                    ],
                    "model": "openai",
                    "temperature": chosen_temp
                }
                
                response = requests.post(url, json=payload, timeout=30)
                
                if response.status_code == 200:
                    st.success("✅ تمت الترجمة الفورية بنجاح:")
                    st.code(response.text, language="text")
                else:
                    st.error(f"عذراً، الخادم مشغول حالياً، يرجى المحاولة مرة أخرى.")
                    
        except Exception as e:
            st.error(f"حدث خطأ أثناء الاتصال الفوري: {e}")
