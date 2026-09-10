import google.generativeai as genai
from dotenv import load_dotenv
import os

# Đọc key từ .env
load_dotenv()
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

# Khởi tạo model và bắt đầu phiên trò chuyện (nhớ ngữ cảnh)
model = genai.GenerativeModel("gemini-3.6-flash")
chat = model.start_chat(history=[])

print("=== Chatbot AI của bạn ===")
print("Gõ 'thoat' để dừng chương trình.\n")

# Vòng lặp hỏi-đáp liên tục
while True:
    user_input = input("Bạn: ")
    
    if user_input.lower() == "thoat":
        print("Tạm biệt!")
        break
    
    response = chat.send_message(user_input)
    print("AI:", response.text, "\n")