import os
import streamlit as st
from google import genai

st.set_page_config(
page_title="Tasuqilt",
page_icon="🇩🇿",
layout="wide"
)

st.title("Tasuqilt 🇩🇿")
st.subheader("مترجم الأمازيغية المعيارية")

api_key = os.environ.get("GEMINI_API_KEY")

if not api_key:
st.error("لم يتم العثور على GEMINI_API_KEY في Streamlit Secrets.")
st.stop()

try:
client = genai.Client(api_key=api_key)
except Exception as error:
st.error("تعذر إنشاء اتصال Gemini.")
st.code(str(error))
st.stop()

mode = st.sidebar.radio(
"اختر وضع الترجمة:",
[
"إخباري رسمي وصارم",
"أدبي / ثقافي"
]
)

if mode == "إخباري رسمي وصارم":
instructions = """
أنت مترجم محترف متخصص في الأمازيغية المعيارية.

ترجم النص إلى الأمازيغية المعيارية بالحرف اللاتيني.

القواعد:

لا تضف معلومات غير موجودة في النص.

لا تحذف معلومات من النص.

حافظ على أسماء الأشخاص والأماكن والمؤسسات.

حافظ على الأرقام والتواريخ.

حافظ على المعنى الكامل للنص.

استخدم أسلوبًا صحفيًا رسميًا ودقيقًا.

لا تقدم أي شرح.

أخرج الترجمة الأمازيغية فقط.
"""
else:
instructions = """
أنت مترجم محترف متخصص في الأمازيغية المعيارية.

ترجم النص إلى الأمازيغية المعيارية بالحرف اللاتيني.

القواعد:

حافظ على المعنى الكامل للنص.

لا تضف معلومات غير موجودة في النص.

لا تحذف معلومات مهمة.

استخدم لغة أمازيغية سليمة وطبيعية.

اجعل الأسلوب مناسبًا للنص الأدبي أو الثقافي.

حافظ على الأسماء والأماكن والأرقام والتواريخ.

لا تقدم أي شرح.

أخرج الترجمة الأمازيغية فقط.
"""

text_to_translate = st.text_area(
"أدخل النص بالعربية أو الفرنسية:",
height=220,
placeholder="اكتب النص هنا..."
)

if st.button("بدء الترجمة", type="primary"):

if not text_to_translate.strip():
    st.warning("يرجى إدخال نص أولًا.")
    st.stop()

prompt = f"""


{instructions}

النص الأصلي:

{text_to_translate}

الترجمة الأمازيغية:
"""

try:
    with st.spinner("جاري الترجمة بواسطة Gemini..."):

        interaction = client.interactions.create(
            model="gemini-3.8-flash",
            input=prompt
        )

    result = interaction.output_text

    if result and result.strip():

        st.success("تمت الترجمة بنجاح.")

        st.text_area(
            "الترجمة الأمازيغية:",
            value=result.strip(),
            height=280
        )

    else:
        st.error("Gemini لم يرجع نصًا.")

except Exception as error:

    st.error("حدث خطأ أثناء الاتصال بـ Gemini.")

    st.code(
        str(error),
        language="text"
    )
