# BÀI TẬP 1: XÂY DỰNG PIPELINE HỒI QUY VỚI TIỀN XỬ LÝ VÀ GIẢM CHIỀU PCA

**Học phần:** Nhận dạng mẫu (*Pattern Recognition*)  
**Chương trình:** Cao học / Thạc sĩ Công nghệ Thông tin  
**Học viên:** Phạm Tuấn Anh (Mã HV: B26CHKH005 - Lớp: M26CQKH02-B)  
**Giảng viên hướng dẫn:** TS. Nguyễn Xuân Đức  
**Đơn vị:** Học viện Công nghệ Bưu chính Viễn thông  

---

## 1. Giới thiệu & Mục tiêu bài toán

Trong học máy và nhận dạng mẫu, bài toán hồi quy nhằm dự báo giá trị liên tục từ các đặc trưng đầu vào. Trong thực tế, các bộ dữ liệu thường gặp phải vấn đề:
- Đa cộng tuyến giữa các biến độc lập.
- Chênh lệch thứ nguyên lớn giữa các thuộc tính.
- Hiện tượng bùng nổ số chiều làm tăng chi phí tính toán và nguy cơ quá khớp.

### Mục tiêu chính:
1. **Xây dựng pipeline hồi quy khép kín:** Đóng gói toàn bộ các bước từ tiền xử lý (Imputation, Standard Scaling) đến giảm chiều (PCA) và mô hình dự báo bằng `sklearn.pipeline.Pipeline`, loại bỏ hoàn toàn nguy cơ rò rỉ dữ liệu (*Data Leakage*).
2. **So sánh với mô hình Baseline:** Dùng `DummyRegressor` (dự báo giá trị trung bình) làm mốc cơ sở đánh giá tính hữu ích của mô hình học máy.
3. **Thử nghiệm đa dạng 8 mô hình hồi quy:** Linear Regression, Ridge, Lasso, K-Nearest Neighbors, Decision Tree, Random Forest, HistGradientBoosting và Support Vector Regression (SVR).
4. **Đánh giá độ ổn định:** Sử dụng kiểm định chéo 5-Fold Cross-Validation, tính toán bộ ba chỉ số $MAE$, $RMSE$, $R^2$ kèm độ lệch chuẩn $\pm \text{std}$.
5. **Kiểm tra tác động của PCA:** Phân tích sự thay đổi hiệu năng khi áp dụng PCA giữ 95% phương sai và quét biến thiên số thành phần $k \in [1, 8]$.

---

## 2. Cấu trúc thư mục

```text
├── regression_pipeline.py       # Pipeline huấn luyện, thực nghiệm và xuất biểu đồ
├── create_word_report.py        # Script tự động tạo báo cáo Word chuẩn học thuật
├── requirements.txt             # Danh sách thư viện phụ thuộc
├── .gitignore                   # Cấu hình bỏ qua file tạm, tài liệu docx
├── README.md                    # Tài liệu hướng dẫn sử dụng
└── results/                     # Thư mục biểu đồ và kết quả chi tiết
    ├── eda_correlation_distribution.png  # Phân tích tương quan & phân phối dữ liệu
    ├── pca_variance.png                  # Biểu đồ Scree Plot phương sai tích lũy PCA
    ├── model_comparison.png              # So sánh hiệu năng R2 giữa 8 mô hình
    ├── actual_vs_predicted.png           # Biểu đồ tán xạ giá trị thực tế vs dự báo
    ├── pca_impact.png                    # Đánh giá tác động của PCA lên từng mô hình
    ├── pca_components_sweep.png          # Khảo sát biến thiên số thành phần PCA k=1..8
    └── results_summary.csv               # Bảng số liệu tổng hợp kết quả thực nghiệm
```

---

## 3. Tóm tắt kết quả thực nghiệm

### 3.1. Bảng so sánh hiệu năng các mô hình (5-Fold Cross-Validation)

| Mô hình | R² (Không PCA) | RMSE (Không PCA) | MAE (Không PCA) | R² (Có PCA 95%) | Tác động PCA ($\Delta R^2$) |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Dummy Baseline** | -0,0005 ± 0,0005 | 1,1539 | 0,9097 | -0,0005 ± 0,0005 | 0,0000 |
| **Linear Regression** | 0,6028 ± 0,0169 | 0,7277 | 0,5291 | 0,6001 ± 0,0159 | -0,0027 |
| **Ridge Regression** | 0,6028 ± 0,0170 | 0,7277 | 0,5291 | 0,6001 ± 0,0159 | -0,0027 |
| **Lasso Regression** | 0,6026 ± 0,0167 | 0,7278 | 0,5291 | 0,5999 ± 0,0157 | -0,0027 |
| **K-Nearest Neighbors** | 0,6985 ± 0,0128 | 0,6341 | 0,4422 | 0,6897 ± 0,0124 | -0,0088 |
| **Decision Tree** | 0,6114 ± 0,0126 | 0,7198 | 0,4669 | 0,5266 ± 0,0169 | -0,0848 |
| **Random Forest** | 0,8082 ± 0,0099 | 0,5050 | 0,3329 | 0,7597 ± 0,0112 | -0,0485 |
| **HistGradientBoosting** | **0,8364 ± 0,0079** | **0,4665** | **0,3118** | **0,7943 ± 0,0101** | -0,0421 |
| **Support Vector Regression** | 0,7337 ± 0,0108 | 0,5960 | 0,4046 | 0,7255 ± 0,0113 | -0,0082 |

### 3.2. Kết luận then chốt:
1. **Mô hình xuất sắc nhất:** **HistGradientBoosting** đạt $R^2 = 0,8364$, $RMSE = 0,4665$, $MAE = 0,3118$ và độ ổn định cao nhất qua 5 fold ($\sigma = 0,0079$), vượt trội hoàn toàn so với mô hình cơ sở Dummy Baseline ($R^2 \approx 0$).
2. **Tác động của PCA:** 
   - Với bộ dữ liệu có 8 chiều ban đầu, việc áp dụng PCA không giúp cải thiện độ chính xác dự báo do PCA là phép biến đổi tuyến tính không giám sát (*unsupervised*), có thể bỏ qua các mối quan hệ phi tuyến cục bộ có ích cho mô hình cây.
   - PCA làm giảm nhẹ hiệu năng của các mô hình phi tuyến mạnh (Random Forest giảm $0,0485$, HistGradientBoosting giảm $0,0421$).
   - Tuy nhiên, khi cần giảm chiều sâu (ví dụ nén về 4-5 thành phần), PCA vẫn bảo toàn được phần lớn hiệu năng với chi phí bộ nhớ thấp hơn.

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
Lệnh dưới đây sẽ tự động tải dữ liệu California Housing, thực hiện phân tích EDA, huấn luyện 8 mô hình qua 5-Fold CV (có và không có PCA), khảo sát sweep PCA từ 1 đến 8 thành phần và lưu toàn bộ biểu đồ vào thư mục `results/`:
```bash
python regression_pipeline.py
```

### 4.4. Tạo báo cáo Word học thuật
Lệnh dưới đây sẽ đọc bìa gốc `BÌA TIỂU LUẬN.docx` (giữ nguyên logo PTIT, định dạng và khung viền), ghép nối 8 chương nội dung hoàn chỉnh, chèn 6 biểu đồ, 3 bảng số liệu và lưu ra file `Bao_cao_Bai_tap_1.docx`:
```bash
python create_word_report.py
```

---

## 5. Giấy phép & Tác quyền
Dự án được xây dựng phục vụ nghiên cứu và hoàn thành Bài tập 1 môn Nhận dạng mẫu - Học viện Công nghệ Bưu chính Viễn thông.