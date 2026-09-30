# Dựng Chương 2 có biểu đồ Excel gốc

Bộ công cụ này tạo lại hai tệp ở thư mục gốc:

- `Chuong_2_Thuc_trang_chuan_hoa_bieu_do.docx`: Chương 2 đã chuẩn hóa số liệu, có 15 biểu đồ Excel nhúng. Đây là biểu đồ thật, không phải ảnh: nhấp chuột phải vào biểu đồ trong Word và chọn Chỉnh sửa dữ liệu để mở bảng số liệu.
- `Du_lieu_bieu_do_Chuong_2.xlsx`: mỗi hình một sheet gồm bảng dữ liệu và biểu đồ; kèm sheet `Nhat_ky_chuan_hoa` ghi các chỉnh sửa số liệu và căn cứ đối chiếu.

## Cách chạy

```bash
pip install python-docx xlsxwriter lxml
python3 scripts/chuong2/build.py
```

Đầu vào là `Chuong_2_Thuc_trang_hoan_chinh mới.docx`. Tệp này không bị sửa.

## Cấu trúc

- `du_lieu.py`: số liệu đã đối chiếu với tài liệu gốc; nguồn ghi ngay cạnh từng khối.
- `bieu_do.py`: định nghĩa từng hình, gồm bảng, loại biểu đồ, vị trí chèn và lời bình. Các con số trong lời bình được tính từ `du_lieu.py`.
- `excel_chart.py`: vẽ biểu đồ Excel bằng XlsxWriter.
- `build.py`: sửa số liệu trong văn bản, chèn biểu đồ vào Word, ghi workbook.

Muốn sửa một số liệu: sửa trong `du_lieu.py` rồi chạy lại `build.py`. Bảng, biểu đồ và lời bình sẽ cập nhật cùng lúc.
