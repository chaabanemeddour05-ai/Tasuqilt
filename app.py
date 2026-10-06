try:

    with st.spinner("⏳ جاري الترجمة بواسطة Gemini..."):

        interaction = client.interactions.create(
            model="gemini-3.8-flash",
            input=prompt
        )

    # استخراج النص الناتج
    result = getattr(interaction, "output_text", None)

    if not result:
        result = str(interaction)

    if result.strip():

        st.success("✅ تمت الترجمة بنجاح")

        st.text_area(
            "الترجمة الأمازيغية:",
            value=result.strip(),
            height=280
        )

    else:

        st.error(
            "⚠️ Gemini لم يُرجع نصًا."
        )

except Exception as e:

    st.error(
        "❌ حدث خطأ أثناء الاتصال بـ Gemini."
    )

    st.code(
        str(e),
        language="text"
    )
