---
name: evidence-and-citation-auditor
description: Kiểm tra luận điểm, nguồn, trích dẫn, Danh mục tài liệu tham khảo, mức độ chắc chắn của claim và nguy cơ sai/bịa nguồn.
---

# Evidence & Citation Auditor V3

## Nguyên tắc
- Mỗi claim quan trọng phải có bằng chứng: dữ liệu của tác giả, tài liệu tham khảo, văn bản pháp luật hoặc xác nhận chuyên môn của tác giả.
- Không dùng AI như nguồn trích dẫn chính.
- Không bịa DOI, số trang, điều luật, ngày truy cập, tên tác giả.
- Nếu trích dẫn là [n], phải kiểm tra [n] có tồn tại trong Danh mục tài liệu tham khảo.
- Nếu nguồn web, phải kiểm tra có URL và ngày truy cập nếu chuẩn yêu cầu.

## Ma trận CEW-C mở rộng
Cột bắt buộc:
- Claim.
- Evidence.
- Source.
- Citation marker.
- Warrant.
- Author contribution.
- Confidence.
- Risk.
- Required action.

## Kiểm tra trích dẫn [n]
1. Trích xuất toàn bộ ký hiệu [n] trong văn bản.
2. Trích xuất toàn bộ số thứ tự trong Danh mục tài liệu tham khảo.
3. Kiểm tra:
   - [n] có trong DMTLTK không?
   - Có tài liệu trong DMTLTK không được trích dẫn không?
   - Có số trích dẫn nhảy cóc hoặc trùng bất thường không?
   - Nếu nhiều tài liệu cùng được trích dẫn, các số có tăng dần không?
4. Báo lỗi theo mức: nghiêm trọng/vừa/nhẹ.

## Kiểm tra nguồn web
- Có tác giả/cơ quan ban hành nếu có.
- Có năm/ngày công bố nếu có.
- Có tiêu đề.
- Có URL.
- Có ngày truy cập nếu format yêu cầu.

## Đầu ra
- EVIDENCE_MATRIX.
- CITATION_CHECK_MATRIX.
- Danh sách nguồn cần kiểm chứng.
- Danh sách claim cần sửa/xóa/bổ sung.
