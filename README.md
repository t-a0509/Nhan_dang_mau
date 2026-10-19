# BÀI TẬP 1: XÂY DỰNG PIPELINE HỒI QUY VỚI TIỀN XỬ LÝ VÀ GIẢM CHIỀU PCA

---

## 1. Giới thiệu & Mục tiêu bài toán

Trong học máy và nhận dạng mẫu, bài toán hồi quy nhằm dự báo giá trị liên tục từ các đặc trưng đầu vào. Trong thực tế, các bộ dữ liệu thường gặp phải vấn đề:
- Đa cộng tuyến giữa các biến độc lập.
- Chênh lệch thứ nguyên lớn giữa các thuộc tính.
- Hiện tượng bùng nổ số chiều làm tăng chi phí tính toán và nguy cơ quá khớp.

**Dữ liệu:** California Housing (20.640 mẫu, 8 đặc trưng). Biến mục tiêu `MedHouseVal` — giá nhà trung vị (đơn vị 100.000 USD), là biến liên tục.

### Mục tiêu chính:
1. **Xây dựng pipeline hồi quy khép kín:** Đóng gói toàn bộ các bước từ tiền xử lý (Imputation, Standard Scaling) đến giảm chiều (PCA) và mô hình dự báo bằng `sklearn.pipeline.Pipeline`, tránh rò rỉ dữ liệu (*Data Leakage*).
2. **So sánh với mô hình Baseline:** Dùng `DummyRegressor` (dự báo giá trị trung bình) làm mốc cơ sở đánh giá tính hữu ích của mô hình học máy.
3. **Thử nghiệm đa dạng 8 mô hình hồi quy:** Linear Regression, Ridge, Lasso, K-Nearest Neighbors, Decision Tree, Random Forest, HistGradientBoosting và Support Vector Regression (SVR).
4. **Đánh giá và lựa chọn mô hình:** Chia Train (80%) / Test hold-out (20%). Trên tập Train dùng 5-Fold Cross-Validation tính $MAE$, $RMSE$, $R^2$ kèm độ lệch chuẩn $\pm \text{std}$; mô hình được chọn tự động theo hạng trung bình của MAE, RMSE, R² và độ ổn định (std của R²), rồi đánh giá **một lần** trên tập Test.
5. **Kiểm tra tác động của PCA:** So sánh từng mô hình khi dùng PCA $k = 5$ (số thành phần nhỏ nhất giữ ≥ 90% phương sai) với không dùng PCA, và quét số thành phần $k \in [1, 8]$.

---

## 2. Cấu trúc thư mục

```text
├── regression_pipeline.py       # Pipeline huấn luyện, thực nghiệm và xuất biểu đồ
├── create_word_report.py        # Script tự động tạo báo cáo Word (đọc số liệu từ results/)
├── requirements.txt             # Danh sách thư viện phụ thuộc
├── .gitignore                   # Cấu hình bỏ qua file tạm, tài liệu docx
├── README.md                    # Tài liệu hướng dẫn sử dụng
└── results/                     # Thư mục biểu đồ và kết quả chi tiết
    ├── eda_correlation_distribution.png  # Phân tích tương quan & phân phối dữ liệu
    ├── pca_variance.png / .csv           # Scree Plot và bảng phương sai tích lũy PCA (tập Train)
    ├── model_comparison.png              # So sánh R², RMSE, MAE giữa 9 mô hình (CV)
    ├── pca_impact.png                    # Đánh giá tác động của PCA lên từng mô hình (CV)
    ├── pca_components_sweep.png          # Khảo sát biến thiên số thành phần PCA k=1..8
    ├── pca_sweep_results.csv             # Số liệu khảo sát k=1..8
    ├── results_summary.csv               # Kết quả 5-Fold CV của mọi cấu hình
    ├── model_ranking.csv                 # Bảng xếp hạng dùng để chọn mô hình
    ├── test_results.csv                  # Kết quả cuối cùng trên tập Test hold-out
    └── actual_vs_predicted.png           # Thực tế vs dự báo trên tập Test
```

---

## 3. Tóm tắt kết quả thực nghiệm

### 3.1. So sánh các mô hình (5-Fold CV trên tập Train, 16.512 mẫu)

| Mô hình | R² (Không PCA) | RMSE (Không PCA) | MAE (Không PCA) | R² (Có PCA 5D) | Tác động PCA ($\Delta R^2$) |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Dummy Baseline** | −0,0002 ± 0,0002 | 1,1562 | 0,9139 | −0,0002 ± 0,0002 | 0,0000 |
| **Linear Regression** | 0,6115 ± 0,0124 | 0,7205 | 0,5291 | 0,4557 ± 0,0161 | −0,1557 |
| **Ridge Regression** | 0,6115 ± 0,0124 | 0,7205 | 0,5291 | 0,4557 ± 0,0161 | −0,1557 |
| **Lasso Regression** | 0,6078 ± 0,0117 | 0,7239 | 0,5323 | 0,4558 ± 0,0161 | −0,1520 |
| **K-Nearest Neighbors** | 0,6880 ± 0,0087 | 0,6456 | 0,4426 | 0,5811 ± 0,0101 | −0,1069 |
| **Decision Tree** | 0,6811 ± 0,0118 | 0,6527 | 0,4333 | 0,4792 ± 0,0077 | −0,2019 |
| **Random Forest** | 0,7946 ± 0,0059 | 0,5240 | 0,3490 | 0,5915 ± 0,0125 | −0,2031 |
| **HistGradientBoosting** | **0,8322 ± 0,0055** | **0,4736** | **0,3165** | 0,6022 ± 0,0125 | −0,2300 |
| **Support Vector Regression** | 0,7365 ± 0,0103 | 0,5932 | 0,3951 | 0,5667 ± 0,0166 | −0,1698 |

### 3.2. Đánh giá cuối cùng trên tập Test hold-out (4.128 mẫu)

| Vai trò | Mô hình | Cấu hình | MAE | RMSE | R² |
| :--- | :--- | :---: | :---: | :---: | :---: |
| Baseline | Dummy Baseline | Không PCA | 0,9061 | 1,1449 | −0,0002 |
| **Mô hình được chọn** | **HistGradientBoosting** | **Không PCA** | **0,3111** | **0,4654** | **0,8347** |
| Đối chứng PCA | HistGradientBoosting | PCA 5D | 0,5271 | 0,7411 | 0,5809 |

### 3.3. Kết luận then chốt:
1. **Mô hình được chọn:** **HistGradientBoosting không dùng PCA** đứng hạng 1 ở cả 4 tiêu chí (MAE, RMSE, R², độ ổn định) trên CV: $R^2 = 0,8322 \pm 0,0055$, $RMSE = 0,4736$, $MAE = 0,3165$. Trên tập Test đạt $R^2 = 0,8347$ — khớp với CV, chứng tỏ mô hình tổng quát hóa tốt và vượt xa Dummy Baseline ($R^2 \approx 0$).
2. **Tác động của PCA:**
   - PCA 5D (giữ 90,16% phương sai của X) làm **giảm rõ rệt** hiệu năng của mọi mô hình: R² giảm tương đối 15,5% (KNN) đến 29,6% (Decision Tree); RMSE tăng 15,9% đến 53,9%.
   - Nguyên nhân: PCA không giám sát nên có thể loại bỏ các thành phần ít phương sai nhưng có ích cho dự báo; đồng thời phép quay PCA phá vỡ cấu trúc đặc trưng gốc mà mô hình cây khai thác.
   - Bằng chứng: tại $k = 8$ (không mất thông tin), Linear Regression đạt lại đúng $R^2 = 0,6115$, nhưng Random Forest chỉ đạt 0,7155 và HistGradientBoosting 0,7464 — thấp hơn hẳn so với không dùng PCA.

---

## 4. Hướng dẫn cài đặt và sử dụng

### 4.1. Yêu cầu môi trường
- Python 3.10 trở lên.
- Các thư viện liệt kê trong `requirements.txt`.

### 4.2. Cài đặt các gói phụ thuộc
```bash
pip install -r requirements.txt
```

### 4.3. Chạy toàn bộ thực nghiệm
Lệnh dưới đây sẽ tự động tải dữ liệu California Housing, thực hiện EDA, tách tập Test hold-out, đánh giá 9 mô hình qua 5-Fold CV (có và không có PCA), chọn mô hình, khảo sát PCA từ 1 đến 8 thành phần, đánh giá trên tập Test và lưu toàn bộ biểu đồ/CSV vào thư mục `results/`:
```bash
python regression_pipeline.py
```

### 4.4. Tạo báo cáo Word học thuật
Lệnh dưới đây đọc bìa gốc `BÌA TIỂU LUẬN.docx` (giữ nguyên logo PTIT, định dạng và khung viền; tiêu đề và tên môn được thay thành "Báo cáo Bài tập 1 — Nhận dạng mẫu"), lấy toàn bộ số liệu từ các file CSV trong `results/`, chèn biểu đồ, bảng số liệu và lưu ra file `Bao_cao_Bai_tap_1.docx`. Cần chạy `regression_pipeline.py` trước:
```bash
python create_word_report.py
```

---

## 5. Giấy phép & Tác quyền
Dự án được xây dựng phục vụ nghiên cứu và học tập bài toán hồi quy kết hợp giảm chiều dữ liệu PCA.
