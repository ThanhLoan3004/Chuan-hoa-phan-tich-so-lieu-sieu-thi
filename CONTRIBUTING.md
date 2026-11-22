# Hướng dẫn Đóng góp (Contributing Guidelines)

Cảm ơn bạn đã quan tâm và muốn đóng góp cho dự án **Chuẩn hóa, Xử lý và Phân tích Dữ liệu Kinh doanh Siêu thị (Supermarket Sales)**! Mọi ý kiến đóng góp nhằm cải thiện chất lượng mã nguồn, mở rộng tính năng phân tích hoặc hoàn thiện tài liệu đều được hoan nghênh.

---

## 1. Các hình thức đóng góp

Bạn có thể tham gia đóng góp cho dự án thông qua các hình thức sau:
* **Báo cáo lỗi (Bug Reports):** Phát hiện lỗi xử lý dữ liệu, sai lệch công thức tính toán, lỗi hiển thị biểu đồ hoặc lỗi tương thích môi trường Python.
* **Đề xuất tính năng (Feature Requests):** Gợi ý bổ sung các chiều phân tích mới (ví dụ: phân tích theo khung giờ, phương thức thanh toán, loại khách hàng `Member`/`Normal`), thêm chỉ số thống kê hoặc định dạng xuất báo cáo mới.
* **Cải thiện mã nguồn & Hiệu năng:** Tối ưu hóa tốc độ xử lý trên `pandas`/`numpy`, bổ sung kiểm thử đơn vị (`unit tests`) hoặc nâng cao khả năng kiểm tra dữ liệu đầu vào.
* **Hoàn thiện tài liệu:** Bổ sung giải thích thuật toán, sửa lỗi chính tả hoặc cập nhật hướng dẫn sử dụng trong `README.md`.

---

## 2. Quy trình đóng góp qua Git / GitHub

1. **Fork repository:** Nhấn nút **Fork** ở góc trên bên phải trang kho lưu trữ để tạo bản sao về tài khoản GitHub của bạn.
2. **Clone về máy cục bộ:**
   ```bash
   git clone https://github.com/<ten-tai-khoan-cua-ban>/<ten-repo>.git
   cd <ten-repo>
   ```
3. **Tạo nhánh làm việc mới (Branch):** Không chỉnh sửa trực tiếp trên nhánh `main`. Hãy tạo nhánh theo quy ước đặt tên rõ ràng:
   ```bash
   git checkout -b feature/ten-tinh-nang-moi
   # hoặc nếu sửa lỗi:
   git checkout -b fix/mo-ta-loi-ngan-gon
   ```
4. **Cài đặt môi trường và kiểm thử thay đổi:**
   ```bash
   py -m pip install pandas numpy matplotlib openpyxl
   py phan_tich_supermarket.py --input supermarket_Sales.csv
   ```
5. **Commit thay đổi với thông điệp rõ ràng:**
   ```bash
   git add .
   git commit -m "feat: them phan tich doanh thu theo khung gio"
   ```
6. **Đẩy nhánh lên GitHub và tạo Pull Request (PR):**
   ```bash
   git push origin feature/ten-tinh-nang-moi
   ```
   Sau đó mở **Pull Request** vào nhánh `main` của kho lưu trữ gốc, mô tả rõ mục đích, cách kiểm thử và kết quả đạt được.

---

## 3. Quy tắc chuẩn mực khi đóng góp mã nguồn

* **Giữ nguyên tính toàn vẹn dữ liệu gốc:** Tuyệt đối không tự ý điền khuyết (impute) hoặc ghi đè giá trị trong cột `Total` hay các trường nguồn; mọi bản ghi bị loại phải được lưu vết kèm lý do rõ ràng.
* **Tách biệt chỉ tiêu tài chính:** Luôn phân biệt rõ *Tổng tiền có thuế (`Total`)* với *Doanh thu thuần* hoặc *Lợi nhuận thực tế*.
* **Phong cách mã nguồn:** Tuân thủ chuẩn **PEP 8** của Python, đặt tên biến/hàm rõ nghĩa và viết chú thích (`docstring`/comment) bằng tiếng Việt có dấu thống nhất với `phan_tich_supermarket.py`.
* **Không commit thư mục kết quả tạm:** Không đưa các thư mục sinh tự động trong `ket_qua/` vào commit trừ khi có yêu cầu minh họa cụ thể.

---

## 4. Giấy phép

Khi gửi đóng góp (Pull Request) cho dự án này, bạn đồng ý rằng đóng góp của bạn sẽ được cấp phép theo **[Giấy phép MIT (MIT License)](./LICENSE)** giống như toàn bộ dự án.
