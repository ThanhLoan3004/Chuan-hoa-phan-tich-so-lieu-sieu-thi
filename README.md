# Bài tập 3: Chuẩn hóa, Xử lý và Phân tích Dữ liệu Kinh doanh Siêu thị (Supermarket Sales)

Dự án thực hiện quy trình kiểm tra chất lượng, chuẩn hóa số liệu giao dịch siêu thị, phân tích thống kê mô tả bằng Python và đối chiếu công thức Excel để phát hiện các mẫu hình kinh doanh theo thời gian, nhóm hàng và chi nhánh.

* **Tác giả / Mã học viên:** Trương Thị Thanh Loan (`m2526019`)
* **Môn học:** Công nghệ số (Bài thực hành số 2 – Bài tập 3)
* **Giấy phép:** Dự án được phát hành theo **[Giấy phép MIT (MIT License)](./LICENSE)**.

---

## 1. Giới thiệu Đề tài

Trong hoạt động bán lẻ hiện đại, dữ liệu giao dịch tại điểm bán (POS) chứa nhiều thông tin quan trọng về hành vi mua sắm, cơ cấu mặt hàng và hiệu suất từng chi nhánh. Tuy nhiên, dữ liệu thô thường tiềm ẩn nguy cơ sai định dạng ngày tháng, trùng lặp bản ghi, sai lệch công thức tính thuế hoặc nhầm lẫn giữa các chỉ tiêu tài chính (như tổng tiền có thuế và lợi nhuận ròng).

Đề tài xây dựng một quy trình tự động hóa bằng Python (`phan_tich_supermarket.py`) kết hợp tệp bảng tính đối chiếu (`Supermarket_Sales_Phan_tich_cong_thuc.xlsx`) trên bộ dữ liệu **Supermarket Sales** gồm **1.000 giao dịch** (từ ngày `01/01/2019` đến `30/03/2019` tại 3 chi nhánh A, B, C) nhằm biến đổi dữ liệu thô thành các bảng tổng hợp, biểu đồ trực quan và báo cáo phân tích có tính giải thích cao.

* **Nguồn dữ liệu gốc (Kaggle):** [Supermarket Sales Dataset](https://www.kaggle.com/datasets/aungpyaeap/supermarket-sales)
* **Bản dữ liệu CSV đối chiếu:** [Plotly Supermarket Sales CSV](https://raw.githubusercontent.com/plotly/datasets/master/supermarket_Sales.csv)

---

## 2. Mục tiêu

1. **Kiểm tra và chuẩn hóa dữ liệu nghiêm ngặt:** Đảm bảo tính toàn vẹn của 7 trường dữ liệu bắt buộc (`Invoice ID`, `Branch`, `Product line`, `Unit price`, `Quantity`, `Total`, `Date`), phát hiện và tách riêng các dòng lỗi hoặc trùng lặp mà không tự ý suy đoán hay ghi đè số liệu gốc.
2. **Xác minh tính nhất quán kế toán – thuế:** Đối chiếu cột `Total` gốc với công thức chuẩn $\text{Unit price} \times \text{Quantity} \times 1{,}05$ (thuế 5%), đồng thời phân tách rõ ràng giữa *Tổng tiền có thuế* và *Lợi nhuận thực tế*.
3. **Phát hiện mẫu hình (Pattern Discovery) và giải thích dữ liệu:** Phân tích biến động doanh số theo 3 chiều:
   * **Theo thời gian (Tháng):** Đánh giá tổng tiền, số giao dịch, bình quân/ngày có dữ liệu và bình quân/giao dịch.
   * **Theo nhóm hàng (Product line):** Xác định tỷ trọng đóng góp và mức độ chênh lệch giữa 6 nhóm ngành hàng.
   * **Theo chi nhánh (Branch):** So sánh quy mô giao dịch và giá trị giỏ hàng trung bình giữa các chi nhánh A, B, C.
4. **Tự động hóa xuất báo cáo đa định dạng:** Sinh đồng thời báo cáo giải thích chi tiết (`.txt`), các bảng dữ liệu sạch (`.csv`), 4 biểu đồ trực quan độ phân giải cao (`.png`) và lưu vết mã băm `SHA-256` của tệp nguồn theo từng phiên chạy.

---

## 3. Phương pháp Thực hiện

Quy trình xử lý được thiết kế theo kiến trúc pipeline 4 bước khép kín trong tệp `phan_tich_supermarket.py`:

### Bước 1: Đọc và xác thực cấu trúc đầu vào (`read_source`)
* Hỗ trợ linh hoạt định dạng `.csv` (tự động nhận diện dấu phân cách `,` hoặc `;`, mã hóa `utf-8-sig`) và `.xlsx` (sheet `Du_lieu`).
* Chuẩn hóa tên cột, kiểm tra trùng lặp tên cột và xác minh sự hiện diện của toàn bộ các cột bắt buộc.

### Bước 2: Làm sạch, chuẩn hóa và kiểm soát chất lượng (`prepare`)
* **Chuẩn hóa kiểu dữ liệu:** Cắt khoảng trắng dư thừa ở các trường định danh/phân loại; ép kiểu số thực/nguyên cho đơn giá, số lượng, tổng tiền; chuẩn hóa ngày tháng theo định dạng chuẩn Mỹ (`%m/%d/%Y`) hoặc ISO (`%Y-%m-%d`), tuyệt đối không suy đoán ngày `dd/mm/yyyy` để tránh đảo ngược ngày – tháng.
* **Lọc và lưu vết dữ liệu lỗi:** Loại bỏ các bản ghi thiếu giá trị, số lượng không phải số nguyên dương, hoặc bản ghi trùng lặp hoàn toàn trên 17 trường nguồn. Toàn bộ bản ghi bị loại được xuất ra tệp `Dong_bi_loai.csv` kèm cột `Ly_do_loai`.
* **Kiểm tra công thức thuế:** Tính toán `Tien_truoc_thue_tinh`, `Tong_tinh_thue_5pct` và `Chenh_lech_Total` để cảnh báo nếu sai lệch vượt quá ngưỡng `0,01`.

### Bước 3: Tổng hợp và đối chiếu chéo (`analyse`)
* Nhóm dữ liệu theo `Thang`, `Product line` và `Branch` để tính tổng số giao dịch (`count`), tổng tiền có thuế (`sum`), giá trị bình quân/giao dịch (`mean`), tổng số lượng sản phẩm và tỷ trọng (`Ty_trong`).
* Tính tốc độ tăng trưởng theo tháng (`pct_change`) chỉ trên các tháng liền kề nhau.
* Chạy kiểm tra đối chiếu (`np.isclose`) đảm bảo tổng doanh số và tổng số giao dịch ở mọi bảng phân tích khớp tuyệt đối với tập dữ liệu đã làm sạch.

### Bước 4: Trực quan hóa và xuất báo cáo (`charts` & `report`)
* Khởi tạo thư mục đầu ra riêng biệt theo dấu thời gian (`ket_qua/YYYYMMDD_HHMMSS_<uuid>/`) cho mỗi lần chạy để tránh ghi đè kết quả cũ.
* Vẽ 4 biểu đồ bằng `matplotlib` (`180 DPI`) và xuất tệp `Bao_cao_phan_tich.txt` kèm mã kiểm tra toàn vẹn `SHA-256`.

---

## 4. Kết quả Đạt được

Khi chạy chương trình trên tệp `supermarket_Sales.csv`, hệ thống ghi nhận các kết quả chính sau:

### 4.1. Chất lượng dữ liệu
* **Tổng số dòng đọc vào:** `1.000` giao dịch (`01/01/2019` – `30/03/2019`).
* **Số dòng không hợp lệ / trùng lặp bị loại:** `0` dòng (100% dữ liệu đạt chuẩn).
* **Số dòng lệch công thức thuế (> 0,01):** `0` dòng.
* **Tổng tiền bán hàng có thuế toàn kỳ:** **`322.966,75`** ($ theo đơn vị nguồn).

### 4.2. Các mẫu hình kinh doanh nổi bật
1. **Biến động theo tháng:**
   * **Tháng 01/2019:** `352` giao dịch | Tổng tiền: `116.291,87` | Bình quân/giao dịch: `330,37` | Bình quân/ngày (31 ngày): `3.751,35`.
   * **Tháng 02/2019:** `303` giao dịch | Tổng tiền: `97.219,37` (-16,40% so với T1) | Bình quân/giao dịch: `320,86` | Bình quân/ngày (28 ngày): `3.472,12`.
   * **Tháng 03/2019:** `345` giao dịch | Tổng tiền: `109.455,51` (+12,59% so với T2) | Bình quân/giao dịch: `317,26` | Bình quân/ngày (30 ngày): `3.648,52`.
   * *Nhận xét:* Giá trị bình quân trên mỗi giao dịch có xu hướng giảm nhẹ liên tục qua cả 3 tháng (`330,37` $\rightarrow$ `320,86` $\rightarrow$ `317,26`).
2. **Cơ cấu nhóm hàng:**
   * **Dẫn đầu:** *Thực phẩm và đồ uống (Food and beverages)* đạt `56.144,84` (chiếm `17,38%`, `174` giao dịch).
   * **Thấp nhất:** *Sức khỏe và làm đẹp (Health and beauty)* đạt `49.193,74` (chiếm `15,23%`, `152` giao dịch).
   * Nhóm cao nhất vượt nhóm thấp nhất **`14,13%`** (chênh lệch `2,15` điểm phần trăm tỷ trọng).
3. **So sánh chi nhánh:**
   * **Chi nhánh C** dẫn đầu về tổng tiền với **`110.568,71`** (`328` giao dịch, bình quân **`337,10`** / giao dịch).
   * **Chi nhánh A** có số lượng giao dịch nhiều nhất (**`340`** giao dịch) nhưng tổng tiền đạt `106.200,37` do giá trị bình quân/giao dịch thấp hơn (`312,35`).
   * **Chi nhánh B** đạt `106.197,67` (`332` giao dịch, bình quân `319,87` / giao dịch).

---

## 5. Cấu trúc Thư mục và Hướng dẫn Sử dụng

### Cấu trúc tệp dự án
```text
BaiTap3/
├── phan_tich_supermarket.py                  # Mã nguồn chính xử lý, phân tích và xuất báo cáo
├── supermarket_Sales.csv                     # Dữ liệu giao dịch siêu thị đầu vào (1.000 dòng)
├── Supermarket_Sales_Phan_tich_cong_thuc.xlsx # Bảng tính Excel đối chiếu công thức
├── README.md                                 # Tài liệu giới thiệu dự án
├── LICENSE                                   # Tệp giấy phép mã nguồn mở MIT
└── ket_qua/                                  # Thư mục chứa kết quả theo từng phiên chạy
    └── <YYYYMMDD_HHMMSS_id>/
        ├── Bao_cao_phan_tich.txt             # Báo cáo phân tích & giải thích chi tiết
        ├── Du_lieu_da_kiem_tra.csv           # Dữ liệu sạch kèm các cột kiểm tra thuế
        ├── Dong_bi_loai.csv                  # Danh sách bản ghi lỗi/trùng (nếu có)
        ├── Kiem_tra_du_lieu.csv              # Bảng thống kê chất lượng dữ liệu
        ├── Phan_tich_thang.csv               # Bảng tổng hợp theo tháng
        ├── Phan_tich_nhom_hang.csv           # Bảng tổng hợp theo nhóm hàng
        ├── Phan_tich_chi_nhanh.csv           # Bảng tổng hợp theo chi nhánh
        ├── 01_tong_tien_theo_thang.png       # Biểu đồ tổng tiền theo tháng
        ├── 02_binh_quan_theo_thang.png       # Biểu đồ bình quân ngày & giao dịch
        ├── 03_nhom_hang.png                  # Biểu đồ doanh số theo nhóm hàng
        └── 04_so_sanh_chi_nhanh.png          # Biểu đồ so sánh 3 chi nhánh
```

### Cài đặt và Chạy chương trình
1. **Cài đặt các thư viện Python cần thiết:**
   ```bash
   py -m pip install pandas numpy matplotlib openpyxl
   ```
2. **Chạy phân tích với tệp CSV:**
   ```bash
   py phan_tich_supermarket.py --input supermarket_Sales.csv
   ```
3. **Hoặc chạy phân tích với tệp Excel đối chiếu:**
   ```bash
   py phan_tich_supermarket.py --input Supermarket_Sales_Phan_tich_cong_thuc.xlsx
   ```

---

## 6. Giấy phép (License)

Dự án này được cấp phép theo **Giấy phép MIT (MIT License)** — cho phép tự do sử dụng, sao chép, chỉnh sửa, hợp nhất, xuất bản và phân phối lại mã nguồn với điều kiện giữ nguyên thông báo bản quyền và điều khoản miễn trừ trách nhiệm. Xem chi tiết đầy đủ tại tệp [`LICENSE`](./LICENSE).
