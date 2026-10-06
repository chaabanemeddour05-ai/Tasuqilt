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
    mode_instruction = "التزم حرفياً بأسلوب وكالة الأنباء (APS)، والنحو المسترجع، ولا تبتكر شيئاً إطلاقاً."
else:
    chosen_temp = 0.7
    mode_instruction = "يمكنك صياغة الأمازيغية بمرونة وجمالية لتلائم الأسلوب الأدبي والثقافي المفتوح."

# قراءة قواميس الأمازيغية من ملف الوورد المرفق
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
    if not context_paragraphs and not os.path.exists("database.docx"):
        st.error("⚠️ لم يتم العثور على ملف 'database.docx' كمدونة أسلوبية في المجلد.")
    elif not text_to_translate.strip():
        st.warning("⚠️ يرجى كتابة أو لصق نص البرقية أولاً.")
    else:
        try:
            with st.spinner("⏳ جاري استدعاء الذاكرة الترجمية وصياغة الخبر فوراً..."):
                # الفلترة الذكية للذاكرة الترجمية الموضعية لتقليص حجم البيانات وضمان السرعة
                words = [w.strip() for w in text_to_translate.lower().split() if len(w.strip()) > 3]
                relevant_lines = []
                for p in context_paragraphs:
                    if any(word in p.lower() for word in words):
                        relevant_lines.append(p)
                    if len(relevant_lines) >= 15:
                        break
                
                if not relevant_lines:
                    relevant_lines = context_paragraphs[:10]
                
                refined_context = "\n".join(relevant_lines)
                
                system_prompt = f"""أنت رئيس تحرير ومترجم رسمي معتمد للغة الأمازيغية المعيارية الصرفة بالحرف اللاتيني لصالح وكالة الأنباء.
                لديك مدونة أسلوبية ونحوية مسترجعة من الأرشيف الرسمي وهي:
                ---
                {refined_context}
                ---
                تعليمات التحرير الحالية: {mode_instruction}
                مهمتك: صياغة البرقية المدخلة بلغة أمازيغية إعلامية رصينة ومطابقة تماماً للقواعد المرفقة أعلاه دون زيادة أو ابتكار خارجي."""
                
                # استخدام محرك تشغيل فوري، مفتوح وخفيف يتجاوز قيود غوغل اليومية المعقدة تماماً
                url = "https://pollinations.ai"
                payload = {
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": text_to_translate}
                    ],
                    "model": "openai-large",
                    "temperature": chosen_temp
                }
                
                response = requests.post(url, json=payload, timeout=30)
                
                if response.status_code == 200 and response.text.strip():
                    st.success("✅ تمت الترجمة الفورية بنجاح:")
                    st.code(response.text.strip(), language="text")
                else:
                    st.error("الخادم مشغول حالياً، يرجى المحاولة مرة أخرى سريعا.")
                    
        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة الاتصال الفوري: {e}")
