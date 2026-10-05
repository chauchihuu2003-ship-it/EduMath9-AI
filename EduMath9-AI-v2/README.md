# EduMath 9 AI v2

Ứng dụng web AI cho giáo viên Toán 9 – Kết nối tri thức.

## Có gì trong v2?
- OpenAI Responses API + Structured Outputs để AI trả dữ liệu đúng schema.
- Danh mục 32 bài học làm metadata nền.
- Kho nguồn giáo viên: upload PDF/DOCX/TXT/MD, gắn nguồn cho từng bài hoặc dùng chung.
- AI ưu tiên nguồn đã nạp khi tạo kế hoạch/slide.
- Chỉnh sửa sản phẩm trực tiếp trên web rồi lưu lại.
- Lưu lịch sử sản phẩm trong `data/projects.json`.
- Xuất Word `.docx` và PowerPoint `.pptx` thật.
- API key chỉ nằm ở backend `.env`.

## Cài đặt
1. Cài Node.js 20+.
2. Mở thư mục bằng VS Code.
3. Chạy `npm install`.
4. Copy `.env.example` thành `.env`.
5. Điền `OPENAI_API_KEY`.
6. Chạy `npm run dev`.
7. Mở http://localhost:3000

## Dữ liệu SGK
App không tự tuyên bố có toàn văn SGK. Để bám sát một bài cụ thể, giáo viên nên nạp tài liệu mình có quyền sử dụng vào Kho nguồn. App sẽ trích xuất văn bản và lưu theo bài.

## Bảo mật
Không đưa API key vào `public/index.html`. Khi triển khai public, cần thêm đăng nhập, rate limit, HTTPS và database thật thay cho JSON file.
