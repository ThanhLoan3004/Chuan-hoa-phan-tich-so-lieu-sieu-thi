#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Bài tập 3: Phát hiện mẫu và giải thích dữ liệu kinh doanh.

Cài thư viện:
    py -m pip install pandas numpy matplotlib openpyxl
Chạy với CSV:
    py phan_tich_supermarket.py --input supermarket_sales.csv
Chạy với tệp Excel đã tạo:
    py phan_tich_supermarket.py --input Supermarket_Sales_Phan_tich_cong_thuc.xlsx

Nguồn Kaggle: https://www.kaggle.com/datasets/aungpyaeap/supermarket-sales
Bản CSV đối chiếu: https://raw.githubusercontent.com/plotly/datasets/master/supermarket_Sales.csv

Chương trình dùng cột Total gốc làm chỉ tiêu tổng tiền có thuế.
Không dùng Gross income để kết luận lợi nhuận thực tế.
Mỗi lần chạy tạo một thư mục kết quả mới để tránh nhầm với kết quả cũ.
"""

from __future__ import annotations

import argparse
from datetime import datetime
import hashlib
from pathlib import Path
import sys
from uuid import uuid4

try:
    import numpy as np
    import pandas as pd
    import matplotlib
    matplotlib.use("Agg")  # Lưu hình, không cần mở cửa sổ đồ họa.
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
except ImportError as exc:
    raise SystemExit(
        "Thiếu thư viện. Chạy: py -m pip install pandas numpy matplotlib openpyxl"
    ) from exc


REQUIRED = ["Invoice ID", "Branch", "Product line", "Unit price", "Quantity", "Total", "Date"]
PRODUCT_VI = {
    "Food and beverages": "Thực phẩm và đồ uống",
    "Sports and travel": "Thể thao và du lịch",
    "Electronic accessories": "Phụ kiện điện tử",
    "Fashion accessories": "Phụ kiện thời trang",
    "Home and lifestyle": "Nhà cửa và đời sống",
    "Health and beauty": "Sức khỏe và làm đẹp",
}


def number(value: float, digits: int = 2) -> str:
    """Định dạng tiếng Việt chỉ khi trình bày; không đổi kiểu số để tính."""
    return f"{value:,.{digits}f}".translate(str.maketrans({",": ".", ".": ","}))


def read_source(path: Path, sheet: str) -> pd.DataFrame:
    if not path.is_file():
        raise ValueError(f"Không tìm thấy tệp: {path}\nHãy truyền đúng đường dẫn sau --input.")
    if path.suffix.lower() == ".csv":
        # sep=None nhận dạng dấu phẩy/chấm phẩy. Giữ ID là chuỗi.
        data = pd.read_csv(path, sep=None, engine="python", encoding="utf-8-sig", dtype=str)
    elif path.suffix.lower() == ".xlsx":
        data = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    else:
        raise ValueError("Chỉ hỗ trợ CSV hoặc XLSX.")
    data.columns = data.columns.astype(str).str.strip()
    if data.columns.duplicated().any():
        raise ValueError("Tên cột bị trùng. Cần sửa trước khi phân tích.")
    missing = sorted(set(REQUIRED) - set(data.columns))
    if missing:
        raise ValueError("Thiếu các cột bắt buộc: " + ", ".join(missing))
    return data


def prepare(data: pd.DataFrame):
    """Kiểm tra; loại dòng không hợp lệ có lưu vết, không tự điền dữ liệu."""
    original = data.copy()
    df = data.copy()
    quality = {"Số dòng đọc từ tệp": len(df)}
    # Excel tạo trước có cột công thức bổ sung và thông tin nguồn ở bên phải.
    # Chỉ kiểm tra tính đầy đủ trên các cột thực sự dùng trong bài phân tích.
    quality["Ô thiếu trong các cột bắt buộc (trước chuẩn hóa)"] = int(df[REQUIRED].isna().sum().sum())
    for col in ["Invoice ID", "Branch", "Product line"]:
        df[col] = df[col].astype("string").str.strip().replace("", pd.NA)
    for col in ["Unit price", "Quantity", "Total"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # CSV gốc dùng tháng/ngày/năm. Chấp nhận thêm ISO yyyy-mm-dd.
    # Không suy đoán ngày dd/mm/yyyy vì dễ đảo tháng và ngày.
    if pd.api.types.is_datetime64_any_dtype(df["Date"]):
        dates = pd.to_datetime(df["Date"], errors="coerce")
    else:
        text = df["Date"].astype("string").str.strip()
        dates = pd.to_datetime(text, format="%m/%d/%Y", errors="coerce")
        iso = pd.to_datetime(text, format="%Y-%m-%d", errors="coerce")
        dates = dates.fillna(iso)
    df["Date"] = dates.dt.normalize()

    reasons = pd.Series("", index=df.index)
    def flag(mask, message):
        mask = mask.fillna(True)
        reasons.loc[mask] = reasons.loc[mask] + message + "; "

    flag(df[REQUIRED].isna().any(axis=1), "Thiếu dữ liệu hoặc sai kiểu số/ngày")
    for col in ["Unit price", "Quantity", "Total"]:
        flag(~np.isfinite(df[col]) | (df[col] <= 0), f"{col} phải là số dương hữu hạn")
    flag(df["Quantity"].notna() & (df["Quantity"] % 1 != 0), "Quantity phải là số nguyên")

    bad = reasons.ne("")
    rejected = original.loc[bad].copy()
    rejected["Ly_do_loai"] = reasons.loc[bad]
    quality["Dòng không hợp lệ bị loại"] = int(bad.sum())
    df = df.loc[~bad].copy()
    # So sánh các trường nguồn, bỏ qua cột công thức và ghi chú của workbook.
    source_cols = [c for c in original.columns[:17] if c in df.columns]
    duplicate = df.duplicated(subset=source_cols, keep="first")
    duplicates = original.loc[df.index[duplicate]].copy()
    duplicates["Ly_do_loai"] = "Dòng nguồn trùng hoàn toàn (giữ lần đầu)"
    rejected = pd.concat([rejected, duplicates], ignore_index=True)
    quality["Dòng trùng hoàn toàn bị loại"] = int(duplicate.sum())
    df = df.loc[~duplicate].copy()
    if df.empty:
        raise ValueError("Không còn dòng hợp lệ. Kiểm tra tên cột, kiểu số và ngày tháng.")
    if df["Invoice ID"].duplicated().any():
        raise ValueError("Có mã Invoice ID trùng nhưng nội dung khác nhau. Cần kiểm tra trước khi đếm giao dịch.")

    df["Tien_truoc_thue_tinh"] = df["Unit price"] * df["Quantity"]
    df["Tong_tinh_thue_5pct"] = df["Tien_truoc_thue_tinh"] * 1.05
    df["Chenh_lech_Total"] = df["Total"] - df["Tong_tinh_thue_5pct"]
    quality["Dòng lệch công thức giá × lượng × 1,05 trên 0,01"] = int((df["Chenh_lech_Total"].abs() > 0.01).sum())
    quality["Số dòng sử dụng phân tích"] = len(df)
    df["Thang"] = df["Date"].dt.to_period("M").astype(str)
    return df, rejected, quality


def aggregate(df: pd.DataFrame, key: str) -> pd.DataFrame:
    result = df.groupby(key, sort=True).agg(
        So_giao_dich=("Invoice ID", "count"),
        Tong_tien_co_thue=("Total", "sum"),
        Binh_quan_giao_dich=("Total", "mean"),
        So_luong_san_pham=("Quantity", "sum"),
    )
    result["Ty_trong"] = result["Tong_tien_co_thue"] / df["Total"].sum()
    return result


def analyse(df: pd.DataFrame):
    monthly = aggregate(df, "Thang")
    monthly["So_ngay_co_du_lieu"] = df.groupby("Thang")["Date"].nunique()
    monthly["Binh_quan_ngay"] = monthly["Tong_tien_co_thue"] / monthly["So_ngay_co_du_lieu"]
    # Không tính tăng trưởng giữa hai tháng không liền nhau.
    periods = pd.PeriodIndex(monthly.index, freq="M")
    adjacent = pd.Series([False] + [periods[i].ordinal - periods[i-1].ordinal == 1 for i in range(1, len(periods))], index=monthly.index)
    monthly["Thay_doi_tong_tien"] = monthly["Tong_tien_co_thue"].pct_change(fill_method=None).where(adjacent)
    monthly["Thay_doi_binh_quan_ngay"] = monthly["Binh_quan_ngay"].pct_change(fill_method=None).where(adjacent)
    products = aggregate(df, "Product line").sort_values("Tong_tien_co_thue", ascending=False)
    branches = aggregate(df, "Branch").sort_values("Tong_tien_co_thue", ascending=False)
    # Đối chiếu để phát hiện bỏ sót hoặc cộng trùng trong bảng tổng hợp.
    for table in [monthly, products, branches]:
        if not np.isclose(table["Tong_tien_co_thue"].sum(), df["Total"].sum()):
            raise ValueError("Tổng hợp không khớp tổng nguồn.")
        if table["So_giao_dich"].sum() != len(df):
            raise ValueError("Số giao dịch tổng hợp không khớp.")
    return monthly, products, branches


def charts(monthly, products, branches, out: Path):
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 11, "axes.spines.top": False, "axes.spines.right": False})
    fmt = FuncFormatter(lambda x, pos: number(x, 0))
    def finish(fig, filename):
        fig.tight_layout()
        fig.savefig(out / filename, dpi=180, bbox_inches="tight")
        plt.close(fig)

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(monthly.index, monthly["Tong_tien_co_thue"], marker="o", linewidth=2, color="#235789")
    ax.set(title="Tổng tiền bán hàng có thuế theo tháng", xlabel="Tháng", ylabel="Tổng tiền ($ theo nguồn)")
    ax.yaxis.set_major_formatter(fmt); ax.grid(axis="y", alpha=.2)
    for x, y in zip(monthly.index, monthly["Tong_tien_co_thue"]):
        ax.annotate(number(y), (x, y), xytext=(0, 10), textcoords="offset points", ha="center", fontsize=9)
    ax.margins(y=.22)
    finish(fig, "01_tong_tien_theo_thang.png")

    fig, axes = plt.subplots(1, 2, figsize=(12, 4.5))
    for ax, col, title in zip(axes, ["Binh_quan_ngay", "Binh_quan_giao_dich"], ["Bình quân/ngày có dữ liệu", "Bình quân/giao dịch"]):
        ax.plot(monthly.index, monthly[col], marker="o", color="#267D75", linewidth=2)
        ax.set(title=title, xlabel="Tháng", ylabel="$ theo nguồn")
        ax.yaxis.set_major_formatter(fmt); ax.grid(axis="y", alpha=.2)
        ax.tick_params(axis="x", rotation=25)
    finish(fig, "02_binh_quan_theo_thang.png")

    fig, ax = plt.subplots(figsize=(10, 5.5))
    labels = [PRODUCT_VI.get(x, x) for x in products.index]
    bars = ax.barh(labels, products["Tong_tien_co_thue"], color="#235789")
    ax.invert_yaxis(); ax.set(title="Tổng tiền bán hàng theo nhóm hàng", xlabel="Tổng tiền có thuế ($ theo nguồn)")
    ax.xaxis.set_major_formatter(fmt); ax.set_xlim(0, products["Tong_tien_co_thue"].max()*1.24)
    ax.bar_label(bars, labels=[number(x) for x in products["Tong_tien_co_thue"]], padding=5, fontsize=9)
    finish(fig, "03_nhom_hang.png")

    fig, axes = plt.subplots(1, 2, figsize=(11, 5))
    for ax, col, title in zip(axes, ["So_giao_dich", "Binh_quan_giao_dich"], ["Số giao dịch", "Giá trị bình quân/giao dịch ($)"]):
        table = branches.sort_index()
        bars = ax.bar(table.index.astype(str), table[col], color="#267D75")
        ax.set(title=title, xlabel="Chi nhánh")
        ax.set_ylim(0, table[col].max()*1.18)
        ax.bar_label(bars, labels=[number(x, 0 if col == "So_giao_dich" else 2) for x in table[col]], padding=4)
    finish(fig, "04_so_sanh_chi_nhanh.png")


def report(df, monthly, products, branches, quality, source: Path, out: Path):
    lines = [
        "BÀI TẬP 3 – PHÁT HIỆN MẪU VÀ GIẢI THÍCH", "",
        f"Tệp đầu vào: {source}",
        f"SHA-256 của tệp: {hashlib.sha256(source.read_bytes()).hexdigest()}",
        f"Thời điểm chạy trên máy: {datetime.now():%d/%m/%Y %H:%M:%S}",
        "Nguồn: https://www.kaggle.com/datasets/aungpyaeap/supermarket-sales",
        "Bản CSV đối chiếu: https://raw.githubusercontent.com/plotly/datasets/master/supermarket_Sales.csv", "",
        "1. MỤC TIÊU VÀ PHƯƠNG PHÁP",
        "Phát hiện các mẫu theo thời gian, nhóm hàng và chi nhánh bằng thống kê mô tả.",
        "Tổng tiền dùng cột Total gốc, đã gồm thuế. Không đồng nhất chỉ tiêu này với doanh thu thuần hoặc lợi nhuận.",
        "Số giao dịch là số mã Invoice ID; không phải số khách hàng duy nhất.",
        "Bình quân/giao dịch = tổng tiền / số giao dịch; tỷ trọng = tổng tiền nhóm / tổng toàn bộ.",
        "Bình quân/ngày = tổng tiền / số ngày phân biệt có dữ liệu; không phải bình quân trên toàn bộ ngày mở cửa.", "",
        "2. KIỂM TRA DỮ LIỆU",
        *[f"- {key}: {val}" for key, val in quality.items()],
        f"- Khoảng ngày: {df.Date.min():%d/%m/%Y} đến {df.Date.max():%d/%m/%Y}.",
        f"- Tổng tiền có thuế: {number(df.Total.sum())} ($ theo nguồn).",
        "Dòng bị loại được lưu riêng. Không tự điền giá trị thiếu. Sai lệch tiền chỉ được cảnh báo, không ghi đè Total.", "",
        "3. MẪU 1 – BIẾN ĐỘNG THEO THÁNG",
    ]
    for month, row in monthly.iterrows():
        lines.append(f"- {month}: {int(row.So_giao_dich)} giao dịch; tổng {number(row.Tong_tien_co_thue)}; bình quân/giao dịch {number(row.Binh_quan_giao_dich)}; {int(row.So_ngay_co_du_lieu)} ngày có dữ liệu; bình quân/ngày {number(row.Binh_quan_ngay)}.")
        if pd.notna(row.Thay_doi_tong_tien):
            lines.append(f"  So với tháng liền trước: tổng tiền thay đổi {number(row.Thay_doi_tong_tien*100)}%; bình quân/ngày thay đổi {number(row.Thay_doi_binh_quan_ngay*100)}%.")
    if len(monthly) < 2:
        lines.append("Chỉ có một tháng: chưa thể kết luận biến động theo tháng.")
    elif monthly.Binh_quan_giao_dich.diff().dropna().lt(0).all():
        lines.append("Quan sát: giá trị giao dịch bình quân giảm qua tất cả các kỳ được ghi nhận.")
    lines += [
        "Giải thích số học: tổng tiền bằng số giao dịch nhân giá trị giao dịch bình quân. Số ngày ghi nhận khác nhau cũng ảnh hưởng tổng tháng.",
        "Giả thuyết kinh doanh: thay đổi khuyến mãi, cơ cấu sản phẩm hoặc nhu cầu mua sắm. Tệp chưa đủ thông tin để xác định nguyên nhân.",
        "Ứng dụng: theo dõi đồng thời lượng giao dịch và giá trị bình quân; thử nghiệm gói mua kèm, đo kết quả trước khi mở rộng.", "",
        "4. MẪU 2 – CƠ CẤU NHÓM HÀNG",
    ]
    for group, row in products.iterrows():
        lines.append(f"- {PRODUCT_VI.get(group, group)}: {number(row.Tong_tien_co_thue)}; tỷ trọng {number(row.Ty_trong*100)}%; {int(row.So_giao_dich)} giao dịch.")
    top, low = products.iloc[0], products.iloc[-1]
    lines += [
        f"Nhóm có tổng tiền cao nhất: {PRODUCT_VI.get(products.index[0], products.index[0])}; thấp nhất: {PRODUCT_VI.get(products.index[-1], products.index[-1])}.",
        f"Nhóm cao nhất hơn nhóm thấp nhất {number((top.Tong_tien_co_thue/low.Tong_tien_co_thue-1)*100)}%. Chênh lệch tỷ trọng: {number((top.Ty_trong-low.Ty_trong)*100)} điểm phần trăm.",
        "Giải thích: tổng tiền nhóm phụ thuộc số giao dịch, đơn giá và số lượng. Nhu cầu thường xuyên là giả thuyết cần dữ liệu mua lặp lại để kiểm chứng.",
        "Ứng dụng: tham khảo cơ cấu bán hàng khi nhập hàng; kết hợp tồn kho, vòng quay và chi phí trước khi tăng hoặc giảm một nhóm.", "",
        "5. MẪU 3 – SO SÁNH CHI NHÁNH",
    ]
    for branch, row in branches.iterrows():
        lines.append(f"- {branch}: {int(row.So_giao_dich)} giao dịch; tổng {number(row.Tong_tien_co_thue)}; bình quân/giao dịch {number(row.Binh_quan_giao_dich)}.")
    leader = branches.index[0]
    most_orders = branches.So_giao_dich.idxmax()
    lines.append(f"Chi nhánh có tổng tiền cao nhất: {leader}. Chi nhánh có nhiều giao dịch nhất: {most_orders}.")
    if leader != most_orders:
        lines.append("Quan sát: chi nhánh dẫn đầu tổng tiền không phải chi nhánh có nhiều giao dịch nhất; giá trị bình quân/giao dịch góp phần giải thích khác biệt.")
    lines += [
        "Giả thuyết: cơ cấu sản phẩm, mức giá hoặc số lượng mỗi lần mua khác nhau. Cần đối chiếu các biến này trước khi kết luận về nguyên nhân.",
        "Ứng dụng: nghiên cứu cơ cấu hàng và giỏ mua tại chi nhánh có giá trị bình quân cao; thử nghiệm có đo lường ở chi nhánh khác.", "",
        "6. GIỚI HẠN VÀ KẾT LUẬN",
        "Các kết quả chỉ mô tả tệp được phân tích, không chứng minh quan hệ nhân quả hoặc khác biệt có ý nghĩa thống kê.",
        "Ba tháng dữ liệu chưa đủ để khẳng định mùa vụ hằng năm. Ngày không xuất hiện có thể là thiếu dữ liệu hoặc không có giao dịch; không tự coi là doanh số bằng 0.",
        "Không có mã khách hàng duy nhất, thông tin khuyến mãi, tồn kho, diện tích hoặc chi phí vận hành để kết luận hành vi mua lặp lại hay hiệu quả lợi nhuận.",
        "Trong bản dữ liệu tham chiếu, Gross income bằng Tax 5%; không dùng cột này để suy ra lợi nhuận thực tế.",
        "Các bảng và biểu đồ cung cấp bằng chứng cho ít nhất hai mẫu khi dữ liệu có đủ nhóm/kỳ. Giải thích kinh doanh là giả thuyết cần kiểm chứng bổ sung.", "",
        "CÁCH ĐỌC KẾT QUẢ",
        "Các cột Ty_trong và Thay_doi_* trong CSV lưu dưới dạng tỷ lệ: 0,10 tương ứng 10%.",
        "CSV chứa kết quả tính từ Python, không chứa công thức Excel. Xem mã .py để kiểm tra cách tính.",
    ]
    (out / "Bao_cao_phan_tich.txt").write_text("\n".join(lines), encoding="utf-8-sig")


def main():
    parser = argparse.ArgumentParser(description="Phân tích Supermarket Sales cho Bài tập 3.")
    parser.add_argument("--input", type=Path, default=Path(__file__).resolve().parent / "supermarket_sales.csv", help="Tệp CSV hoặc XLSX đầu vào")
    parser.add_argument("--sheet", default="Du_lieu", help="Tên sheet khi đọc XLSX (mặc định Du_lieu)")
    parser.add_argument("--output", type=Path, default=Path(__file__).resolve().parent / "ket_qua", help="Thư mục cha lưu kết quả")
    args = parser.parse_args()
    source = args.input.expanduser().resolve()
    df, rejected, quality = prepare(read_source(source, args.sheet))
    monthly, products, branches = analyse(df)
    out = args.output.expanduser().resolve() / (datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid4().hex[:6])
    out.mkdir(parents=True, exist_ok=False)
    # Chỉ xuất trường nguồn có ý nghĩa và các biến tính thêm, bỏ ghi chú bên lề Excel.
    cols = list(dict.fromkeys(REQUIRED + ["City", "Customer type", "Gender", "Tax 5%", "Time", "Payment", "Thang", "Tien_truoc_thue_tinh", "Tong_tinh_thue_5pct", "Chenh_lech_Total"]))
    df[[c for c in cols if c in df.columns]].to_csv(out / "Du_lieu_da_kiem_tra.csv", index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    rejected.to_csv(out / "Dong_bi_loai.csv", index=False, encoding="utf-8-sig")
    pd.Series(quality, name="Gia_tri").rename_axis("Noi_dung").to_csv(out / "Kiem_tra_du_lieu.csv", encoding="utf-8-sig")
    for table, filename in [(monthly, "Phan_tich_thang.csv"), (products, "Phan_tich_nhom_hang.csv"), (branches, "Phan_tich_chi_nhanh.csv")]:
        table.to_csv(out / filename, encoding="utf-8-sig", float_format="%.8f")
    charts(monthly, products, branches, out)
    report(df, monthly, products, branches, quality, source, out)
    print(f"Đã phân tích {len(df):,} giao dịch từ: {source}")
    print(f"Tổng tiền có thuế: {number(df.Total.sum())}")
    print(f"Kết quả MỚI được lưu tại: {out}")
    print("Mở Bao_cao_phan_tich.txt, các bảng CSV và 4 biểu đồ PNG trong thư mục trên.")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, OSError, ImportError) as error:
        print(f"LỖI: {error}", file=sys.stderr)
        sys.exit(1)
