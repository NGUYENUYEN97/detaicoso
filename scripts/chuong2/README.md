# Dựng Chương 2 có biểu đồ Excel gốc

Bộ công cụ này tạo lại hai tệp ở thư mục gốc:

- `Chuong_2_Thuc_trang_rut_gon.docx`: Chương 2 bản rút gọn, 8 bảng, 9 biểu đồ Excel nhúng. Đây là biểu đồ thật, không phải ảnh: nhấp chuột phải vào biểu đồ trong Word và chọn Chỉnh sửa dữ liệu để mở bảng số liệu.
- `Du_lieu_bieu_do_Chuong_2.xlsx`: mỗi hình một sheet gồm bảng dữ liệu và biểu đồ; kèm sheet `Nhat_ky_chuan_hoa` ghi các chỉnh sửa số liệu và căn cứ đối chiếu.

## Cách chạy

```bash
pip install python-docx xlsxwriter lxml
python3 scripts/chuong2/build_gon.py
```

`Chuong_2_Thuc_trang_hoan_chinh mới.docx` chỉ được dùng làm khuôn định dạng trang và kiểu chữ; tệp này không bị sửa.

## Cấu trúc

- `du_lieu.py`: số liệu đã đối chiếu với tài liệu gốc; nguồn ghi ngay cạnh từng khối.
- `bieu_do.py`: định nghĩa các hình, gồm bảng dữ liệu và cấu hình biểu đồ.
- `excel_chart.py`: vẽ biểu đồ Excel bằng XlsxWriter; bố cục cố định cho chú giải và vùng vẽ, màu chữ nhãn theo độ tương phản.
- `build_gon.py`: toàn văn Chương 2 bản rút gọn, bảng, chèn biểu đồ, ghi workbook. Các con số trích từ `du_lieu.py` được kiểm tra bằng `assert`.
- `build.py`: bản 43 trang trước đây (sửa trên bản gốc); `build_gon.py` dùng lại các hàm nhúng biểu đồ và dựng workbook của tệp này.

Muốn sửa một số liệu: sửa trong `du_lieu.py` rồi chạy lại `build_gon.py`.
