import streamlit as st

st.set_page_config(page_title="Dummy App 2", layout="centered")

st.title("Dummy App 2")
st.write("Text eingeben und als Vorschau anzeigen.")

text = st.text_area("Dein Text", height=120, placeholder="Schreibe etwas...")

st.subheader("Vorschau")
st.info(text if text.strip() else "(noch kein Text)")
