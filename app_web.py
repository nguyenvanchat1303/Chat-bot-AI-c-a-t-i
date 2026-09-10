import streamlit as st
import google.generativeai as genai
from dotenv import load_dotenv
import os

# Đọc key từ .env
load_dotenv()
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

# Cấu hình trang (phải để ở đầu file)
st.set_page_config(
    page_title="Chatbot AI của tôi",
    page_icon="🤖",
    layout="centered"
)

# Tiêu đề và mô tả
st.title("🤖 Chatbot AI của tôi")
st.caption("Hỏi bất cứ điều gì, tôi sẽ cố gắng trả lời tốt nhất!")

# Thanh bên (sidebar) - thêm thông tin/nút reset
with st.sidebar:
    st.header("⚙️ Cài đặt")
    if st.button("🗑️ Xóa lịch sử trò chuyện"):
        st.session_state.messages = []
        model = genai.GenerativeModel("gemini-3.6-flash")
        st.session_state.chat = model.start_chat(history=[])
        st.rerun()
    st.markdown("---")
    st.markdown("Được xây dựng bằng **Streamlit** + **Google Gemini**")

# Khởi tạo model và lịch sử chat
if "chat" not in st.session_state:
    model = genai.GenerativeModel("gemini-3.6-flash")
    st.session_state.chat = model.start_chat(history=[])
    st.session_state.messages = []

# Hiển thị lịch sử tin nhắn cũ, dùng avatar riêng
for msg in st.session_state.messages:
    avatar = "🧑" if msg["role"] == "user" else "🤖"
    with st.chat_message(msg["role"], avatar=avatar):
        st.write(msg["content"])

# Ô nhập câu hỏi mới
user_input = st.chat_input("Nhập câu hỏi của bạn...")

if user_input:
    st.session_state.messages.append({"role": "user", "content": user_input})
    with st.chat_message("user", avatar="🧑"):
        st.write(user_input)

    with st.chat_message("assistant", avatar="🤖"):
        with st.spinner("Đang suy nghĩ..."):
            response = st.session_state.chat.send_message(user_input)
            st.write(response.text)
    st.session_state.messages.append({"role": "assistant", "content": response.text})