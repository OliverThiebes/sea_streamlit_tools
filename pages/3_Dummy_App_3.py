import random
import streamlit as st

st.set_page_config(page_title="Dummy App 3", layout="centered")

st.title("Dummy App 3")
st.write("Zufallszahlen-Generator.")

min_val, max_val = st.slider("Bereich", 0, 100, (10, 90))
if st.button("Zahl erzeugen"):
    st.success(f"Zufallszahl: {random.randint(min_val, max_val)}")
