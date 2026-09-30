# Prompt mẫu sử dụng V3

## Prompt tổng quát
```text
Hãy dùng workflow /write-academic-product-v3.

Tôi cần thực hiện sản phẩm học thuật sau: [bài báo/chuyên đề tổng quan/chuyên đề nghiên cứu/chương luận án/luận án/đề tài cấp cơ sở].

Tôi đã cung cấp tài liệu mẫu/quy định trình bày/tài liệu tham khảo/dữ liệu. Hãy tự nhận diện chuẩn trình bày, chuẩn trích dẫn, cấu trúc sản phẩm, yêu cầu Danh mục tài liệu tham khảo và các checklist liên quan.

Yêu cầu bắt buộc:
1. Không hỏi lại dài dòng nếu thông tin đã có trong tài liệu mẫu.
2. Trước khi viết, tạo FORMAT_REQUIREMENTS_PROFILE, EVIDENCE_MATRIX và DECISION_LOG.
3. Nếu tài liệu mẫu dùng [n], hãy dùng [n] và kiểm tra [n] khớp Danh mục tài liệu tham khảo.
4. Nếu tài liệu mẫu dùng APA 7th hoặc chuẩn khác, hãy tự chuyển sang chuẩn đó.
5. Nếu có yếu tố pháp luật/chính sách/quản lý nhà nước, kích hoạt diễn ngôn pháp lý.
6. Nếu có khảo sát/phỏng vấn, thiết kế bảng hỏi, ma trận biến quan sát và bộ câu hỏi phỏng vấn.
7. Nếu có phân tích định lượng, đề xuất pipeline xử lý dữ liệu bằng Python.
8. Viết bản nháp theo CEW-C, có phản biện, có dấu ấn tác giả, không công thức hóa.
9. Cuối cùng xuất checklist: định dạng, trích dẫn, nguồn, chữ viết tắt, dữ liệu, nhất quán, AI disclosure.
```

## Prompt cho trường hợp có template của cơ sở đào tạo
```text
Hãy đọc template/quy định của cơ sở đào tạo đã nạp, tạo FORMAT_REQUIREMENTS_PROFILE và áp dụng đúng cho sản phẩm này. Nếu template yêu cầu đề cương 4 chương, hãy dùng cấu trúc 4 chương; nếu template yêu cầu chuyên đề tổng quan theo 2 phần, hãy dùng đúng cấu trúc đó. Không khóa cứng nếu template mới có bố cục khác.
```

## Prompt cho kiểm tra bản thảo cuối
```text
Hãy audit bản thảo theo V3: kiểm tra format, tiêu đề, trích dẫn, Danh mục tài liệu tham khảo, nguồn web có ngày truy cập, chữ viết tắt, bảng/hình, tính nhất quán giữa chuyên đề và luận án, tính hợp lệ của luận điểm, dữ liệu và tuyên bố AI. Xuất bảng lỗi – mức độ – vị trí – cách sửa.
```
