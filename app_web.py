import streamlit as st
import google.generativeai as genai
from dotenv import load_dotenv
import fitz
import os
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

load_dotenv()
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

st.set_page_config(
    page_title="VKSND Khu vực 4 - Nghệ An AI",
    page_icon="⚖️",
    layout="centered"
)

st.title("⚖️ Viện Kiểm sát Nhân dân Khu vực 4 - Nghệ An AI")
st.caption("Hỗ trợ tra cứu, soạn thảo Cáo trạng và bài phát biểu")

SYSTEM_PROMPT = """Bạn là một trợ lý pháp lý chuyên nghiệp, hỗ trợ Kiểm sát viên tại Việt Nam.
Khi soạn thảo Cáo trạng hoặc bài phát biểu, bạn PHẢI:
- Tuân thủ đúng thể thức văn bản tố tụng hình sự Việt Nam (Quốc hiệu, tiêu ngữ, phần nhận thấy, phần nhận định, phần quyết định...)
- Chỉ dựa trên thông tin, tình tiết được cung cấp, KHÔNG tự bịa thêm tình tiết hay chứng cứ không có
- Nếu thiếu thông tin cần thiết, ghi rõ [CẦN BỔ SUNG: ...] tại chỗ đó thay vì tự suy diễn
- Văn phong trang trọng, chính xác, đúng thuật ngữ pháp lý
- Cuối mỗi văn bản, LUÔN ghi rõ: "Đây là bản dự thảo do AI hỗ trợ soạn, cần được Kiểm sát viên có thẩm quyền kiểm tra, đối chiếu hồ sơ và phê duyệt trước khi ban hành chính thức." """

# ====== SIDEBAR ======
def tao_file_word(noi_dung_van_ban, tieu_de, so_hieu="...../CT-VKS", dia_danh="Nghệ An"):
    doc = Document()

    # Lề trang theo Nghị định 30
    section = doc.sections[0]
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin = Mm(30)
    section.right_margin = Mm(15)

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(14)

    # ====== BẢNG 2 CỘT: Cơ quan ban hành (trái) | Quốc hiệu tiêu ngữ (phải) ======
    table = doc.add_table(rows=1, cols=2)
    table.autofit = True

    # Cột trái: Cơ quan ban hành
    cell_trai = table.cell(0, 0)
    p_capttren = cell_trai.paragraphs[0]
    p_capttren.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_capttren = p_capttren.add_run("VIỆN KIỂM SÁT NHÂN DÂN\nTỈNH NGHỆ AN")
    run_capttren.font.size = Pt(13)
    run_capttren.font.name = 'Times New Roman'

    p_capduoi = cell_trai.add_paragraph()
    p_capduoi.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_capduoi = p_capduoi.add_run("VIỆN KIỂM SÁT NHÂN DÂN\nKHU VỰC 4")
    run_capduoi.bold = True
    run_capduoi.font.size = Pt(13)
    run_capduoi.font.name = 'Times New Roman'

    p_gach = cell_trai.add_paragraph()
    p_gach.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_gach = p_gach.add_run("________")
    run_gach.font.size = Pt(13)

    p_sohieu = cell_trai.add_paragraph()
    p_sohieu.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sohieu = p_sohieu.add_run(f"Số: {so_hieu}")
    run_sohieu.font.size = Pt(13)
    run_sohieu.font.name = 'Times New Roman'

    # Cột phải: Quốc hiệu tiêu ngữ
    cell_phai = table.cell(0, 1)
    p_qh = cell_phai.paragraphs[0]
    p_qh.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_qh = p_qh.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    run_qh.bold = True
    run_qh.font.size = Pt(13)
    run_qh.font.name = 'Times New Roman'

    p_tn = cell_phai.add_paragraph()
    p_tn.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_tn = p_tn.add_run("Độc lập - Tự do - Hạnh phúc")
    run_tn.bold = True
    run_tn.font.size = Pt(14)
    run_tn.font.name = 'Times New Roman'

    p_gach2 = cell_phai.add_paragraph()
    p_gach2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_gach2 = p_gach2.add_run("---------------")
    run_gach2.font.size = Pt(13)

    from datetime import datetime
    p_ngay = cell_phai.add_paragraph()
    p_ngay.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hom_nay = datetime.now()
    run_ngay = p_ngay.add_run(f"{dia_danh}, ngày {hom_nay.day} tháng {hom_nay.month} năm {hom_nay.year}")
    run_ngay.italic = True
    run_ngay.font.size = Pt(13)
    run_ngay.font.name = 'Times New Roman'

    # Xóa viền bảng (giữ bố cục nhưng không hiện đường kẻ)
    for cell in table._cells:
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = tcPr.makeelement('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders', {})
        for edge in ('top', 'left', 'bottom', 'right'):
            tag = f'{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{edge}'
            el = tcPr.makeelement(tag, {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val': 'nil'})
            tcBorders.append(el)
        tcPr.append(tcBorders)

    doc.add_paragraph()  # dòng trống

    # ====== TÊN VĂN BẢN ======
    p_ten = doc.add_paragraph()
    p_ten.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_ten = p_ten.add_run(tieu_de)
    run_ten.bold = True
    run_ten.font.size = Pt(16)
    run_ten.font.name = 'Times New Roman'

    doc.add_paragraph()  # dòng trống

    # ====== NỘI DUNG CHÍNH (do AI soạn) ======
    for doan in noi_dung_van_ban.split("\n"):
        if doan.strip():
            p = doc.add_paragraph(doan)
            p.paragraph_format.first_line_indent = Mm(10)
            p.paragraph_format.line_spacing = 1.2
            for run in p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(14)

    doc.add_paragraph()

    # ====== NƠI NHẬN + CHỮ KÝ ======
    table2 = doc.add_table(rows=1, cols=2)

    cell_noinhap = table2.cell(0, 0)
    p_nn = cell_noinhap.paragraphs[0]
    run_nn = p_nn.add_run("Nơi nhận:\n- Như trên;\n- Lưu: VT, HSKS.")
    run_nn.italic = True
    run_nn.font.size = Pt(12)
    run_nn.font.name = 'Times New Roman'

    cell_kyten = table2.cell(0, 1)
    p_chucvu = cell_kyten.paragraphs[0]
    p_chucvu.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_chucvu = p_chucvu.add_run("VIỆN TRƯỞNG")
    run_chucvu.bold = True
    run_chucvu.font.size = Pt(14)
    run_chucvu.font.name = 'Times New Roman'

    for _ in range(4):
        cell_kyten.add_paragraph()

    p_hoten = cell_kyten.add_paragraph()
    p_hoten.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_hoten = p_hoten.add_run("[Họ tên]")
    run_hoten.bold = True
    run_hoten.font.size = Pt(14)
    run_hoten.font.name = 'Times New Roman'

    for cell in table2._cells:
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = tcPr.makeelement('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders', {})
        for edge in ('top', 'left', 'bottom', 'right'):
            tag = f'{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{edge}'
            el = tcPr.makeelement(tag, {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val': 'nil'})
            tcBorders.append(el)
        tcPr.append(tcBorders)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


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
        prompt = f"""Hãy soạn một bản dự thảo Cáo trạng dựa trên thông tin sau, tuân thủ CHÍNH XÁC bố cục và trình tự dưới đây:

THÔNG TIN VỤ ÁN:
- Họ tên bị can: {ten_bi_can}
- Ngày sinh: {ngay_sinh}
- Tội danh đề nghị truy tố: {toi_danh}
- Điều, khoản BLHS: {dieu_luat}
- Tóm tắt hành vi, tình tiết vụ án: {tinh_tiet}
- Chứng cứ chính: {chung_cu}
{pdf_note}

BỐ CỤC BẮT BUỘC CỦA CÁO TRẠNG:

1. PHẦN CĂN CỨ (mở đầu):
- Căn cứ các Điều 41, 236, 239 và 243 của Bộ luật Tố tụng hình sự;
- Căn cứ Quyết định khởi tố vụ án hình sự số...ngày...tháng...năm...của...[ghi [CẦN BỔ SUNG] nếu không có thông tin];
- Căn cứ Quyết định khởi tố bị can số...ngày...tháng...năm...của...[ghi [CẦN BỔ SUNG] nếu không có thông tin];
- Căn cứ bản Kết luận điều tra vụ án hình sự đề nghị truy tố số...ngày...tháng...năm...của...[ghi [CẦN BỔ SUNG] nếu không có thông tin].

2. Câu chuyển: "Trên cơ sở kết quả điều tra đã xác định được như sau:"

3. PHẦN NỘI DUNG (dựa trên thông tin vụ án được cung cấp):
- Diễn biến hành vi phạm tội (trình bày chi tiết, theo trình tự thời gian, dựa trên tình tiết vụ án đã cung cấp)
- Phân tích, đánh giá tình tiết tăng nặng, giảm nhẹ và tình tiết khác có ý nghĩa đối với vụ án
- Việc thu giữ, tạm giữ đồ vật, tài liệu, xử lý vật chứng (nếu có thông tin, nếu không ghi [CẦN BỔ SUNG])
- Phần dân sự (nếu có, nếu không ghi "Không có phần dân sự trong vụ án" hoặc [CẦN BỔ SUNG])

4. Câu chuyển: "Căn cứ vào các tình tiết và chứng cứ nêu trên,"

5. KẾT LUẬN:
- Tổng hợp hành vi phạm tội: nêu ngắn gọn hành vi phạm tội của bị can, tính chất, mức độ, hậu quả, vai trò của bị can
- Câu: "Như vậy có đủ căn cứ để xác định bị can có lý lịch dưới đây đã phạm tội [tên tội danh] như sau:"
- Thông tin bị can: Họ và tên, giới tính, ngày sinh, tiền án tiền sự [ghi "không có tiền án, tiền sự" nếu không có thông tin], biện pháp ngăn chặn đang áp dụng đối với bị can
- Khẳng định: "Bị can [tên] đã phạm tội [tên tội danh], quy định tại [trích dẫn chính xác điều, khoản của BLHS]"
- Tình tiết tăng nặng, giảm nhẹ trách nhiệm hình sự: nêu rõ áp dụng tại điểm, khoản nào của Bộ luật Hình sự (nếu chưa có thông tin cụ thể, ghi [CẦN BỔ SUNG])
- Các nội dung kết luận khác (nếu có)

6. Câu chuyển: "Bởi các lẽ trên,"

7. PHẦN "QUYẾT ĐỊNH" (viết in hoa):
1. Truy tố ra trước Tòa án nhân dân khu vực...[CẦN BỔ SUNG tên tòa án cụ thể] để xét xử bị can [tên] về tội [tên tội danh] theo quy định tại điểm...khoản...Điều...của Bộ luật Hình sự.
2. Kèm theo Cáo trạng có: Hồ sơ vụ án gồm...tập,...tờ, đánh số thứ tự từ 01 đến...[CẦN BỔ SUNG số liệu cụ thể]. Bản kê vật chứng (nếu có). Danh sách những người Viện kiểm sát đề nghị Tòa án triệu tập đến phiên tòa [CẦN BỔ SUNG danh sách nếu có thông tin].

8. PHẦN CUỐI: để trống 2 mục "Nơi nhận" và "Người có thẩm quyền ký" (sẽ được điền riêng, không cần AI tạo phần này).

LƯU Ý: Với bất kỳ thông tin cụ thể nào (số hiệu quyết định, ngày tháng, tên cơ quan, tên tòa án...) mà không có trong dữ liệu được cung cấp, PHẢI ghi rõ [CẦN BỔ SUNG: mô tả thông tin còn thiếu] tại đúng vị trí đó, TUYỆT ĐỐI không tự bịa ra số liệu, ngày tháng hay tên cơ quan."""

        with st.spinner("Đang soạn dự thảo..."):
            response = st.session_state.chat.send_message(prompt)
        st.markdown("### 📄 Kết quả dự thảo")
        st.write(response.text)
        file_word = tao_file_word(response.text, "CÁO TRẠNG")
        st.download_button(
            "💾 Tải về file Word (đúng thể thức Nghị định 30)",
            file_word,
            file_name="du_thao_cao_trang.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

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
        st.markdown("### 📄 Kết quả dự thảo")
        st.write(response.text)
        file_word = tao_file_word(response.text, "Bài phát biểu")
        st.download_button(
            "💾 Tải về file Word (đúng thể thức Nghị định 30)",
            file_word,
            file_name="du_thao_bai_phat_bieu.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

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