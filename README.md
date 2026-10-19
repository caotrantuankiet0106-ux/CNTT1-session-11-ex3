# Phần 1: Phân tích tham số VLOOKUP
Trong bài toán này, 4 tham số của hàm VLOOKUP được xác định gồm: Lookup_value là ô chứa mã linh kiện tại sheet Don_Hang (ví dụ A2); Table_array là vùng bảng dữ liệu bên sheet Danh_Muc (bắt buộc phải dùng dấu $ để cố định như $A$2:$C$100 nhằm tránh bị lệch dòng khi kéo công thức); Col_index_num là số thứ tự cột cần lấy dữ liệu (điền 2 cho Tên linh kiện và 3 cho Đơn giá); và Range_lookup bắt buộc dùng FALSE (hoặc 0) để tìm kiếm chính xác tuyệt đối.

# Phần 2: Công thức hoàn chỉnh (đã kèm IFERROR và TRIM)
Để xử lý triệt để lỗi khoảng trắng thừa và lỗi #N/A khi tra cứu, công thức hoàn chỉnh áp dụng cho ô tương ứng tại sheet Don_Hang như sau:

Lấy Tên linh kiện:

# Excel
=IFERROR(VLOOKUP(TRIM(A2), Danh_Muc!$A$2:$C$100, 2, FALSE), "Không tìm thấy")
Lấy Đơn giá:

# Excel
=IFERROR(VLOOKUP(TRIM(A2), Danh_Muc!$A$2:$C$100, 3, FALSE), "Không tìm thấy")
