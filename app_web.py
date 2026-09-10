import streamlit as st
import google.generativeai as genai
from dotenv import load_dotenv
import fitz
import os

load_dotenv()
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

st.set_page_config(
    page_title="Trợ lý Pháp lý AI",
    page_icon="⚖️",
    layout="centered"
)

st.title("⚖️ Trợ lý Pháp lý AI")
st.caption("Hỗ trợ tra cứu, soạn thảo Cáo trạng và bài phát biểu")

SYSTEM_PROMPT = """Bạn là một trợ lý pháp lý chuyên nghiệp, hỗ trợ Kiểm sát viên tại Việt Nam.
Khi soạn thảo Cáo trạng hoặc bài phát biểu, bạn PHẢI:
- Tuân thủ đúng thể thức văn bản tố tụng hình sự Việt Nam (Quốc hiệu, tiêu ngữ, phần nhận thấy, phần nhận định, phần quyết định...)
- Chỉ dựa trên thông tin, tình tiết được cung cấp, KHÔNG tự bịa thêm tình tiết hay chứng cứ không có
- Nếu thiếu thông tin cần thiết, ghi rõ [CẦN BỔ SUNG: ...] tại chỗ đó thay vì tự suy diễn
- Văn phong trang trọng, chính xác, đúng thuật ngữ pháp lý
- Cuối mỗi văn bản, LUÔN ghi rõ: "Đây là bản dự thảo do AI hỗ trợ soạn, cần được Kiểm sát viên có thẩm quyền kiểm tra, đối chiếu hồ sơ và phê duyệt trước khi ban hành chính thức." """

# ====== SIDEBAR ======
with st.sidebar:
    st.header("📄 Tải tài liệu")
    uploaded_file = st.file_uploader("Tải lên file PDF (hồ sơ, văn bản liên quan...)", type=["pdf"], key="pdf_uploader")

    if uploaded_file is not None:
        pdf_bytes = uploaded_file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        pdf_text = ""
        for page in doc:
            pdf_text += page.get_text() + "\n"
        doc.close()
        st.session_state.pdf_content = pdf_text
        st.success(f"Đã đọc xong file: {uploaded_file.name} ({len(pdf_text)} ký tự)")

    st.markdown("---")
    st.header("🎯 Chế độ làm việc")
    che_do = st.radio(
        "Chọn việc bạn muốn làm:",
        ["💬 Trò chuyện tự do", "📋 Soạn dự thảo Cáo trạng", "🎤 Soạn bài phát biểu"]
    )

    st.markdown("---")
    if st.button("🗑️ Xóa lịch sử trò chuyện"):
        st.session_state.messages = []
        if "pdf_content" in st.session_state:
            del st.session_state.pdf_content
        model = genai.GenerativeModel("gemini-3.6-flash", system_instruction=SYSTEM_PROMPT)
        st.session_state.chat = model.start_chat(history=[])
        st.rerun()

# ====== KHỞI TẠO MODEL ======
if "chat" not in st.session_state:
    model = genai.GenerativeModel("gemini-3.6-flash", system_instruction=SYSTEM_PROMPT)
    st.session_state.chat = model.start_chat(history=[])
    st.session_state.messages = []

pdf_note = ""
if "pdf_content" in st.session_state:
    pdf_note = f"\n\n[Nội dung tài liệu PDF đã tải lên]:\n{st.session_state.pdf_content}"

# ====== CHẾ ĐỘ: SOẠN DỰ THẢO CÁO TRẠNG ======
if che_do == "📋 Soạn dự thảo Cáo trạng":
    st.subheader("📋 Thông tin để soạn dự thảo Cáo trạng")

    with st.form("form_cao_trang"):
        ten_bi_can = st.text_input("Họ tên bị can")
        ngay_sinh = st.text_input("Ngày sinh / Năm sinh")
        toi_danh = st.text_input("Tội danh bị đề nghị truy tố")
        dieu_luat = st.text_input("Điều, khoản BLHS áp dụng (nếu biết)")
        tinh_tiet = st.text_area("Tóm tắt hành vi, tình tiết vụ án", height=150)
        chung_cu = st.text_area("Các chứng cứ chính (nếu có)", height=100)
        submitted = st.form_submit_button("🚀 Soạn dự thảo Cáo trạng")

    if submitted:
        prompt = f"""Hãy soạn một bản dự thảo Cáo trạng với thông tin sau:
- Họ tên bị can: {ten_bi_can}
- Ngày sinh: {ngay_sinh}
- Tội danh đề nghị truy tố: {toi_danh}
- Điều, khoản BLHS: {dieu_luat}
- Tóm tắt hành vi, tình tiết vụ án: {tinh_tiet}
- Chứng cứ chính: {chung_cu}
{pdf_note}

Soạn theo đúng bố cục: Quốc hiệu tiêu ngữ, số hiệu, phần nhận thấy (tóm tắt hành vi phạm tội), phần nhận định (căn cứ pháp lý), phần quyết định (đề nghị truy tố)."""

        with st.spinner("Đang soạn dự thảo..."):
            response = st.session_state.chat.send_message(prompt)
        st.markdown("### 📄 Kết quả dự thảo")
        st.write(response.text)
        st.download_button("💾 Tải về file văn bản", response.text, file_name="du_thao_cao_trang.txt")

# ====== CHẾ ĐỘ: SOẠN BÀI PHÁT BIỂU ======
elif che_do == "🎤 Soạn bài phát biểu":
    st.subheader("🎤 Thông tin để soạn bài phát biểu")

    with st.form("form_phat_bieu"):
        loai_phat_bieu = st.selectbox("Loại bài phát biểu", ["Luận tội tại phiên tòa", "Phát biểu tranh luận", "Phát biểu khác"])
        vu_viec = st.text_area("Tóm tắt vụ việc / nội dung cần phát biểu", height=150)
        diem_nhan_manh = st.text_area("Các điểm cần nhấn mạnh (nếu có)", height=100)
        do_dai = st.select_slider("Độ dài mong muốn", options=["Ngắn gọn", "Vừa phải", "Chi tiết"], value="Vừa phải")
        submitted2 = st.form_submit_button("🚀 Soạn bài phát biểu")

    if submitted2:
        prompt = f"""Hãy soạn một bài phát biểu loại "{loai_phat_bieu}" với thông tin sau:
- Tóm tắt vụ việc: {vu_viec}
- Các điểm cần nhấn mạnh: {diem_nhan_manh}
- Độ dài mong muốn: {do_dai}
{pdf_note}

Soạn với văn phong trang trọng, mạch lạc, phù hợp để trình bày tại phiên tòa hoặc cuộc họp chính thức."""

        with st.spinner("Đang soạn bài phát biểu..."):
            response = st.session_state.chat.send_message(prompt)
        st.markdown("### 📄 Kết quả bài phát biểu")
        st.write(response.text)
        st.download_button("💾 Tải về file văn bản", response.text, file_name="bai_phat_bieu.txt")

# ====== CHẾ ĐỘ: TRÒ CHUYỆN TỰ DO ======
else:
    for msg in st.session_state.messages:
        avatar = "🧑" if msg["role"] == "user" else "⚖️"
        with st.chat_message(msg["role"], avatar=avatar):
            st.write(msg["content"])

    user_input = st.chat_input("Đặt câu hỏi pháp lý, hoặc hỏi về file PDF vừa tải lên...", key="main_chat_input")

    if user_input:
        final_input = user_input + pdf_note if "pdf_content" in st.session_state else user_input

        st.session_state.messages.append({"role": "user", "content": user_input})
        with st.chat_message("user", avatar="🧑"):
            st.write(user_input)

        with st.chat_message("assistant", avatar="⚖️"):
            with st.spinner("Đang tra cứu..."):
                response = st.session_state.chat.send_message(final_input)
                st.write(response.text)
        st.session_state.messages.append({"role": "assistant", "content": response.text})