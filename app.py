import os
import streamlit as st
from google import genai
from google.genai import types

st.set_page_config(
    page_title="Tasuqilt",
    layout="wide"
)

st.title("Tasuqilt 🇩🇿")
st.subheader("مترجم الأمازيغية المعيارية")

# ---------------------------------------
# إعداد Gemini
# ---------------------------------------

api_key = os.environ.get("GEMINI_API_KEY")

if not api_key:
    st.error(
        "⚠️ لم يتم العثور على GEMINI_API_KEY. "
        "يرجى التأكد من وجوده في Streamlit Secrets."
    )
    st.stop()

client = genai.Client(api_key=api_key)

# ---------------------------------------
# اختيار نمط الترجمة
# ---------------------------------------

mode = st.sidebar.radio(
    "اختر وضع الترجمة:",
    (
        "إخباري رسمي وصارم",
        "أدبي / ثقافي"
    )
)

if mode == "إخباري رسمي وصارم":
    temperature = 0.0
    style_instruction = """
اكتب ترجمة أمازيغية معيارية رسمية وصارمة.
لا تضف أي معلومة غير موجودة في النص الأصلي.
حافظ على أسماء الأشخاص والأماكن والأرقام والتواريخ.
لا تشرح الترجمة.
أعطني الترجمة فقط.
"""
else:
    temperature = 0.7
    style_instruction = """
اكتب ترجمة أمازيغية معيارية سليمة وطبيعية.
حافظ بدقة على معنى النص الأصلي.
يمكنك تحسين الصياغة بما يناسب الأسلوب الثقافي والأدبي.
لا تضف معلومات غير موجودة في النص.
أعطني الترجمة فقط.
"""

# ---------------------------------------
# إدخال النص
# ---------------------------------------

text_to_translate = st.text_area(
    "أدخل النص المراد ترجمته بالعربية أو الفرنسية:",
    height=220,
    placeholder="اكتب النص هنا..."
)

# ---------------------------------------
# الترجمة
# ---------------------------------------

if st.button("بدء الترجمة", type="primary"):

    if not text_to_translate.strip():
        st.warning("⚠️ يرجى إدخال نص أولًا.")
        st.stop()

    prompt = f"""
أنت مترجم محترف متخصص في الأمازيغية المعيارية.

مهمتك هي ترجمة النص التالي إلى الأمازيغية
المعيارية بالحرف اللاتيني.

تعليمات الأسلوب:
{style_instruction}

النص الأصلي:
{text_to_translate}

الترجمة:
"""

    try:
        with st.spinner("⏳ Gemini يعالج النص..."):

            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=temperature
                )
            )

        if response.text:
            st.success("✅ تمت الترجمة بنجاح")
            st.text_area(
                "الترجمة الأمازيغية:",
                value=response.text.strip(),
                height=250
            )
        else:
            st.error("⚠️ لم يُرجع Gemini نصًا.")

    except Exception as e:
        st.error(
            "❌ حدث خطأ أثناء الاتصال بـ Gemini."
        )
        st.code(str(e))
