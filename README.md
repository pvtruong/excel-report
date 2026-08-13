# Excel Template Engine (EJS + ExcelJS)

Một thư viện mạnh mẽ và linh hoạt giúp tạo báo cáo Excel từ template (`.xlsx`) bằng cách kết hợp sức mạnh định dạng của **ExcelJS** và cú pháp template engine của **EJS**.

## 🌟 Tính năng nổi bật

* **Giữ nguyên định dạng:** Bảo toàn 100% style, màu sắc, viền (border), font chữ và các ô gộp (merge cells) từ file mẫu.
* **Cú pháp EJS chuẩn:** Sử dụng `<%= %>` để in dữ liệu, hỗ trợ đầy đủ logic Javascript ngay trong file Excel.
* **Nhân bản dòng thông minh:** Tự động nhân bản dòng dựa trên mảng dữ liệu với các tag đơn giản `[[FOR:...]]`.
* **Hỗ trợ Master-Detail:** Dễ dàng tạo báo cáo phân cấp (Mẹ - Con) với tag `[[MASTER]]` và `[[DETAIL]]`.
* **Tự động nhận diện kiểu số:** Tự động ép kiểu dữ liệu về Number khi xuất ra Excel để các hàm tính toán (`SUM`, `AVERAGE`) hoạt động chính xác.
* **Tài liệu tự động (Self-documenting):** Tự động sinh Sheet `HD_DU_LIEU_DEBUG` chứa danh sách biến và cấu trúc JSON hiện tại để người làm template đối chiếu.
* **An toàn & Chống Crash:** Cơ chế xử lý lỗi an toàn khi người dùng nhập sai tên biến hoặc định dạng ô không hợp lệ (RichText, Date, Formula).

## 🚀 Cài đặt

```bash
npm install excel-report
```

## 📖 Hướng dẫn sử dụng

### 1. Code xử lý (Backend)

```javascript
const renderExcelTemplate = require('excel-template-engine-js');

const data = {
    khach_hang: "Nguyễn Văn A",
    ngay_xuat: new Date(),
    hoa_don: [
        { ten_sp: "Laptop Dell", sl: 1, gia: 15000000 },
        { ten_sp: "Chuột không dây", sl: 2, gia: 350000 }
    ]
};

// Truyền vào Buffer của file template và nhận lại Buffer file kết quả
const resultBuffer = await renderExcelTemplate(templateBuffer, data);
```

### 2. Cú pháp trong Excel Template

| Mục đích | Cú pháp trong ô Excel | Ghi chú |
| :--- | :--- | :--- |
| **In dữ liệu** | `<%= khach_hang %>` | In chuỗi ký tự bình thường |
| **Vòng lặp dòng** | `[[FOR:hoa_don]]` | Đặt ở cột đầu tiên của dòng cần lặp |
| **In STT dòng** | `<%= hoa_don_index + 1 %>` | Tự động tạo biến `_index` bắt đầu từ 0 |



### 3. Báo cáo Master-Detail (Mẹ - Con)

Thiết kế trên 2 dòng liên tiếp trong Excel:
1.  **Dòng Mẹ:** Đặt `[[MASTER:danh_sach_don]]` ở cột A. Truy cập dữ liệu qua biến `don_hang`.
2.  **Dòng Con:** Đặt `[[DETAIL:chi_tiet]]` ở cột A (ngay dưới dòng mẹ). Truy cập dữ liệu qua biến `chi_tiet`.

## 🛠 Tính năng Debug thông minh

Thư viện tự động tạo một sheet tên là `HD_DU_LIEU_DEBUG`. Sheet này giúp người làm mẫu (Template Creator) biết chính xác:
* Danh sách tất cả các biến có thể sử dụng.
* Kiểu dữ liệu của từng biến (String, Number, Array, Function).
* Chữ ký của các hàm helper (ví dụ: `moment(time, format)`).
* Toàn bộ cấu trúc JSON thực tế đang được đổ vào template.

## ⚖️ License

MIT. Tự do sử dụng và đóng góp cho cộng đồng.
