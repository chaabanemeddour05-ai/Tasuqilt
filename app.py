import os
import streamlit as st
from google import genai

st.set_page_config(page_title="Tasuqilt", page_icon="🇩🇿", layout="wide")

st.title("Tasuqilt 🇩🇿")
st.subheader("مترجم الأمازيغية المعيارية")

api_key = os.environ.get("GEMINI_API_KEY")

if not api_key: st.error("لم يتم العثور على GEMINI_API_KEY في Streamlit Secrets."); st.stop()

client = genai.Client(api_key=api_key)

mode = st.sidebar.radio("اختر وضع الترجمة:", ["إخباري رسمي وصارم", "أدبي / ثقافي"])

instructions = {
"إخباري رسمي وصارم": "أنت مترجم محترف متخصص في الأمازيغية المعيارية. ترجم النص إلى الأمازيغية المعيارية بالحرف اللاتيني. لا تضف أي معلومة غير موجودة في النص. لا تحذف أي معلومة موجودة في النص. حافظ على أسماء الأشخاص والأماكن والمؤسسات والأرقام والتواريخ. حافظ على المعنى الكامل للنص. استخدم أسلوبًا صحفيًا رسميًا ودقيقًا. أخرج الترجمة الأمازيغية فقط.",
"أدبي / ثقافي": "أنت مترجم محترف متخصص في الأمازيغية المعيارية. ترجم النص إلى الأمازيغية المعيارية بالحرف اللاتيني. حافظ على المعنى الكامل للنص. لا تضف معلومات غير موجودة في النص. لا تحذف معلومات مهمة. استخدم لغة أمازيغية سليمة وطبيعية. اجعل الأسلوب مناسبًا للنص الأدبي أو الثقافي. حافظ على الأسماء والأماكن والأرقام والتواريخ. أخرج الترجمة الأمازيغية فقط."
}

text_to_translate = st.text_area("أدخل النص بالعربية أو الفرنسية:", height=220, placeholder="اكتب النص هنا...")

start = st.button("بدء الترجمة", type="primary")

if not start: st.stop()

if not text_to_translate.strip(): st.warning("يرجى إدخال نص أولًا."); st.stop()

prompt = f"{instructions[mode]}\n\nالنص الأصلي:\n\n{text_to_translate}\n\nالترجمة الأمازيغية:"

st.info("جاري الترجمة بواسطة Gemini...")

interaction = client.interactions.create(model="gemini-3.8-flash", input=prompt)

result = interaction.output_text

if result and result.strip(): st.success("تمت الترجمة بنجاح."); st.text_area("الترجمة الأمازيغية:", value=result.strip(), height=280)
else: st.error("Gemini لم يرجع نصًا.")
