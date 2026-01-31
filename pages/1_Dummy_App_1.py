import streamlit as st

st.set_page_config(page_title="Dummy App 1", layout="centered")

st.title("Dummy App 1")
st.write("Ein einfacher Zähler.")

if "count" not in st.session_state:
    st.session_state.count = 0

col1, col2 = st.columns(2)
if col1.button("+1"):
    st.session_state.count += 1
if col2.button("Reset"):
    st.session_state.count = 0

st.metric("Zähler", st.session_state.count)
