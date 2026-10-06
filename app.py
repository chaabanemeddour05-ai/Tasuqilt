import streamlit as st
from docx import Document
import os
import requests

st.set_page_config(page_title="Tasuqqilt", layout="wide")

st.title("Tasuqqilt 🇩🇿")
st.subheader("مترجم الأمازيغية المعيارية الذكي (ذاكرة ترجمية موضعية)")

# قراءة قاعدة البيانات الأسلوبية من ملف الوورد المرفق
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
            if text and len(text) > 5:  # الاحتفاظ بالجمل والفقرات الفعلية المفيدة
                full_text.append(text)
        return full_text
    except Exception:
        return []

corpus_data = load_tamazight_corpus()

# مربع النص للمدخلات (الفرنسية أو العربية) كما خططنا
text_to_translate = st.text_area("أدخل البرقية الصحفية المراد ترجمتها (الفرنسية أو العربية):", height=150)

if st.button("بدء الترجمة الاحترافية المعتمدة"):
    if not corpus_data and not os.path.exists("database.docx"):
        st.error("⚠️ لم يتم العثور على قاعدة البيانات الأسلوبية 'database.docx' في المستودع.")
    elif not text_to_translate.strip():
        st.warning("⚠️ يرجى كتابة أو لصق نص البرقية أولاً للبدء.")
    else:
        try:
            with st.spinner("⏳ جاري فحص الذاكرة الترجمية واستخلاص الأسلوب الصحفي فوراً..."):
                # 1. استخلاص الكلمات المفتاحية الأساسية من البرقية المدخلة للبحث عنها
                search_words = [w.strip().lower() for w in text_to_translate.split() if len(w.strip()) > 3]
                
                # 2. بناء "الذاكرة الموضعية الخفيفة": استرجاع الفقرات المتطابقة والمشابهة للأسلوب فقط
                matched_style = []
                for paragraph in corpus_data:
                    # فحص ما إذا كانت الفقرة تحتوي على مصطلحات مشابهة للبرقية الصحفية
                    if any(word in paragraph.lower() for word in search_words):
                        matched_style.append(paragraph)
                    if len(matched_style) >= 8:  # نكتفي بأقوى 8 فقرات متطابقة تماماً كعينة أسلوبية مكثفة لضمان الاستجابة اللحظية
                        break
                
                # إذا لم يجد تطابقاً مباشراً، يزوده بأول 8 أسطر قياسية كمدونة نحوية
                if not matched_style:
                    matched_style = corpus_data[:8]
                
                style_context = "\n".join(matched_style)
                
                # 3. صياغة التوجيه الذكي والمصغر لضمان القبول الفوري دون تخطي الحصص
                system_prompt = f"""أنت مساعد ترجمة محترف وصارم للغة الأمازيغية المعيارية بالحرف اللاتيني لوكالة الأنباء الرسمية.
                لديك عينة أسلوبية مسترجعة موضعياً من مدونتك اللغوية لتعليمك الصياغة ونظام النحو وهي:
                ---
                {style_context}
                ---
                مهمتك: ترجمة النص المدخل للغة الأمازيغية المعيارية الصرفة بالحرف اللاتيني. 
                التزم تماماً بالصياغة الرسمية، المصطلحات، وبناء الخبر الصحفي المتبع في العينة المرفقة أعلاه، ولا تبتكر صياغات خارجة عنها."""
                
                # 4. إرسال الطلب عبر محرك معالجة فوري وحر ومستقر بالكامل
                url = "https://pollinations.ai"
                payload = {
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": text_to_translate}
                    ],
                    "model": "openai",
                    "temperature": 0.0  # تجميد الحرارة تماماً لضمان ترجمة إخبارية صارمة وغير مبتكرة
                }
                
                response = requests.post(url, json=payload, timeout=20)
                
                if response.status_code == 200 and response.text.strip():
                    st.success("✅ تمت الترجمة الاحترافية بنجاح فوري:")
                    st.code(response.text.strip(), language="text")
                else:
                    st.error("عذراً، الخادم يمر بضغط مؤقت، يرجى إعادة الضغط على الزر مجدداً الآن.")
                    
        except Exception as e:
            st.error(f"حدث خطأ أثناء الاتصال بالذاكرة الذكية: {e}")
