"""
================================================================================
BÀI TẬP 1: XÂY DỰNG PIPELINE HỒI QUY VỚI TIỀN XỬ LÝ VÀ GIẢM CHIỀU PCA
Chương trình Thạc sĩ - Môn: Nhận dạng mẫu
================================================================================
Mục tiêu:
1. Xây dựng pipeline hoàn chỉnh gồm Tiền xử lý dữ liệu (Imputation, Standardization)
   và Giảm chiều đặc trưng bằng PCA (Principal Component Analysis).
2. So sánh mô hình cơ sở (Dummy Baseline) với 8 mô hình hồi quy đại diện cho các họ thuật toán:
   - Mô hình tuyến tính: Linear Regression, Ridge, Lasso
   - Mô hình dựa trên khoảng cách (Instance-based): K-Nearest Neighbors (KNN)
   - Mô hình cây quyết định: Decision Tree
   - Mô hình kết hợp (Ensemble): Random Forest, HistGradientBoosting
   - Mô hình máy vector hỗ trợ: Support Vector Regression (SVR)
3. Chia dữ liệu Train (80%) / Test hold-out (20%). Mọi phân tích, so sánh và lựa chọn
   mô hình chỉ dùng tập Train qua 5-Fold Cross-Validation (MAE, RMSE, R² và độ ổn định std).
4. Lựa chọn mô hình tự động theo hạng trung bình của MAE, RMSE, R² và độ ổn định qua CV,
   sau đó huấn luyện lại trên toàn bộ tập Train và đánh giá MỘT LẦN trên tập Test hold-out.
5. Đánh giá chuyên sâu tác động của PCA:
   - Phân tích phương sai tích lũy (Cumulative Explained Variance).
   - So sánh trực tiếp từng mô hình (Có PCA vs Không có PCA).
   - Khảo sát biến thiên hiệu năng theo số chiều PCA k in [1, 2, ..., 8].
================================================================================
"""

import os
import sys
import time
import numpy as np
import pandas as pd

# Thiết lập UTF-8 cho console Windows để in tiếng Việt chuẩn xác
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

# Sử dụng backend 'Agg' không giao diện để tránh xung đột luồng tkinter trên Windows
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.datasets import fetch_california_housing
from sklearn.model_selection import KFold, train_test_split
from sklearn.base import clone
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.metrics import mean_absolute_error, root_mean_squared_error, r2_score

# Danh sách các mô hình hồi quy
from sklearn.dummy import DummyRegressor
from sklearn.linear_model import LinearRegression, Ridge, Lasso
from sklearn.neighbors import KNeighborsRegressor
from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import RandomForestRegressor, HistGradientBoostingRegressor
from sklearn.svm import SVR

# Thiết lập giao diện đồ họa chuẩn học thuật
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams.update({
    'font.size': 11,
    'axes.labelsize': 12,
    'axes.titlesize': 13,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'figure.titlesize': 14,
    'figure.autolayout': True
})

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "results")
os.makedirs(OUTPUT_DIR, exist_ok=True)
RANDOM_STATE = 42
TEST_SIZE = 0.2


def load_and_explore_data():
    """
    Tải bộ dữ liệu California Housing và tiến hành phân tích khám phá dữ liệu (EDA).
    Target: Median House Value (đơn vị: 100,000 USD) - biến liên tục.
    """
    print("=" * 80)
    print("1. TẢI VÀ KHÁM PHÁ BỘ DỮ LIỆU (CALIFORNIA HOUSING)")
    print("=" * 80)
    
    housing = fetch_california_housing(as_frame=True)
    df = housing.frame
    X = housing.data
    y = housing.target

    print(f"[*] Kích thước dữ liệu: {df.shape[0]} dòng, {df.shape[1]} cột (8 features + 1 target)")
    print(f"[*] Danh sách đặc trưng (Features): {list(X.columns)}")
    print(f"[*] Tên biến mục tiêu (Target): {y.name} (Giá trị trung vị nhà ở, đơn vị: $100k)")
    print(f"[*] Kiểm tra missing values:\n{df.isnull().sum()}")
    print("\n[*] Thống kê mô tả 5 dòng đầu tiên:")
    print(df.head())
    print("\n[*] Thống kê phân bố:")
    print(df.describe().T[['mean', 'std', 'min', '50%', 'max']])

    # Vẽ biểu đồ EDA: Phân bố Target và Ma trận tương quan
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    
    # Phân bố biến mục tiêu
    sns.histplot(y, kde=True, ax=axes[0], color='#2b5c8f', bins=40)
    axes[0].set_title('Phân bố biến mục tiêu (MedHouseVal)')
    axes[0].set_xlabel('Giá trị nhà trung vị ($100,000)')
    axes[0].set_ylabel('Tần suất')

    # Ma trận tương quan Pearson
    corr = df.corr()
    sns.heatmap(corr, annot=True, fmt=".2f", cmap='coolwarm', cbar=True, ax=axes[1], annot_kws={"size": 8})
    axes[1].set_title('Ma trận hệ số tương quan Pearson giữa các thuộc tính')

    eda_path = os.path.join(OUTPUT_DIR, "eda_correlation_distribution.png")
    plt.savefig(eda_path, dpi=300)
    plt.close()
    print(f"[✓] Đã lưu biểu đồ EDA tại: {eda_path}")

    return X, y


def analyze_pca_variance(X):
    """
    Phân tích tỷ lệ phương sai giải thích (Explained Variance Ratio) của PCA
    trên dữ liệu sau khi chuẩn hóa bằng StandardScaler.
    Chỉ dùng tập Train để không rò rỉ thông tin từ tập Test vào việc chọn k.
    """
    print("\n" + "=" * 80)
    print("2. PHÂN TÍCH GIẢM CHIỀU PCA (PRINCIPAL COMPONENT ANALYSIS)")
    print("=" * 80)

    # Tiền xử lý chuẩn hóa trước khi thực hiện PCA
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    pca = PCA()
    pca.fit(X_scaled)

    var_ratio = pca.explained_variance_ratio_
    cum_var_ratio = np.cumsum(var_ratio)

    print(f"{'Thành phần (PC)':<15} | {'Phương sai giải thích':<25} | {'Phương sai tích lũy':<25}")
    print("-" * 70)
    for i in range(len(var_ratio)):
        print(f"PC{i+1:<13} | {var_ratio[i]:>10.4f} ({var_ratio[i]*100:>5.2f}%)   | {cum_var_ratio[i]:>10.4f} ({cum_var_ratio[i]*100:>5.2f}%)")

    # Xác định số chiều đạt ngưỡng 90% và 95%
    k_90 = int(np.argmax(cum_var_ratio >= 0.90) + 1)
    k_95 = int(np.argmax(cum_var_ratio >= 0.95) + 1)
    pd.DataFrame({
        "PC": [f"PC{i+1}" for i in range(len(var_ratio))],
        "Eigenvalue": pca.explained_variance_,
        "Variance_Ratio": var_ratio,
        "Cumulative_Ratio": cum_var_ratio,
    }).to_csv(os.path.join(OUTPUT_DIR, "pca_variance.csv"), index=False)
    print(f"\n[*] Để giữ lại >= 90% thông tin phương sai: cần k = {k_90} thành phần chính.")
    print(f"[*] Để giữ lại >= 95% thông tin phương sai: cần k = {k_95} thành phần chính.")

    # Vẽ Scree Plot và Cumulative Variance Plot
    fig, ax1 = plt.subplots(figsize=(9, 5))

    pcs = [f"PC{i+1}" for i in range(len(var_ratio))]
    bars = ax1.bar(pcs, var_ratio * 100, color='#4A90E2', alpha=0.7, label='Phương sai từng thành phần (%)')
    ax1.set_xlabel('Các thành phần chính (Principal Components)')
    ax1.set_ylabel('Phương sai giải thích từng thành phần (%)', color='#2C3E50')
    ax1.set_ylim(0, max(var_ratio * 100) + 5)

    # Trục phụ cho phương sai tích lũy
    ax2 = ax1.twinx()
    line = ax2.plot(pcs, cum_var_ratio * 100, color='#D0021B', marker='o', linewidth=2.5, label='Phương sai tích lũy (%)')
    ax2.set_ylabel('Phương sai tích lũy (%)', color='#D0021B')
    ax2.set_ylim(0, 105)
    ax2.axhline(90, color='gray', linestyle='--', alpha=0.7, label='Ngưỡng 90%')
    ax2.axhline(95, color='orange', linestyle=':', alpha=0.7, label='Ngưỡng 95%')

    # Thêm giá trị phần trăm lên thanh bar
    for bar in bars:
        h = bar.get_height()
        ax1.annotate(f'{h:.1f}%',
                     xy=(bar.get_x() + bar.get_width() / 2, h),
                     xytext=(0, 3), textcoords="offset points",
                     ha='center', va='bottom', fontsize=9)

    plt.title('Tỷ lệ phương sai giải thích của các thành phần chính (PCA Scree Plot)')
    # Ghép legend
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='center right')

    pca_plot_path = os.path.join(OUTPUT_DIR, "pca_variance.png")
    plt.savefig(pca_plot_path, dpi=300)
    plt.close()
    print(f"[✓] Đã lưu biểu đồ PCA Variance tại: {pca_plot_path}")

    # Chọn k nhỏ nhất giữ lại >= 90% phương sai cho các thực nghiệm so sánh chính
    return k_90


def get_models():
    """
    Định nghĩa từ điển chứa các mô hình hồi quy đại diện.
    Bao gồm Baseline và 7 mô hình thuộc các trường phái khác nhau.
    """
    models = {
        "Dummy Baseline": DummyRegressor(strategy="mean"),
        "Linear Regression": LinearRegression(),
        "Ridge Regression": Ridge(alpha=1.0, random_state=RANDOM_STATE),
        "Lasso Regression": Lasso(alpha=0.01, random_state=RANDOM_STATE),
        "K-Nearest Neighbors": KNeighborsRegressor(n_neighbors=7, n_jobs=2),
        "Decision Tree": DecisionTreeRegressor(max_depth=10, random_state=RANDOM_STATE),
        "Random Forest": RandomForestRegressor(n_estimators=70, max_depth=12, random_state=RANDOM_STATE, n_jobs=2),
        "HistGradientBoosting": HistGradientBoostingRegressor(max_iter=100, random_state=RANDOM_STATE),
        "Support Vector Regressor": SVR(C=1.0, epsilon=0.1, cache_size=1000)
    }
    return models


def evaluate_pipeline(pipeline, X, y, cv):
    """
    Đánh giá pipeline bằng K-Fold Cross-Validation.
    Trả về danh sách MAE, RMSE, R² cho từng fold, cùng thời gian huấn luyện.
    """
    fold_maes = []
    fold_rmses = []
    fold_r2s = []
    fit_times = []

    X_vals = X.values if hasattr(X, "values") else X
    y_vals = y.values if hasattr(y, "values") else y

    for train_idx, val_idx in cv.split(X_vals, y_vals):
        X_train, X_val = X_vals[train_idx], X_vals[val_idx]
        y_train, y_val = y_vals[train_idx], y_vals[val_idx]

        t0 = time.time()
        pipeline.fit(X_train, y_train)
        fit_times.append(time.time() - t0)

        preds = pipeline.predict(X_val)

        fold_maes.append(mean_absolute_error(y_val, preds))
        fold_rmses.append(root_mean_squared_error(y_val, preds))
        fold_r2s.append(r2_score(y_val, preds))

    return {
        "mae_mean": np.mean(fold_maes),
        "mae_std": np.std(fold_maes),
        "rmse_mean": np.mean(fold_rmses),
        "rmse_std": np.std(fold_rmses),
        "r2_mean": np.mean(fold_r2s),
        "r2_std": np.std(fold_r2s),
        "avg_fit_time": np.mean(fit_times)
    }


def run_model_comparison(X, y, k_pca=5):
    """
    Chạy thực nghiệm so sánh tất cả các mô hình:
    1. Tiền xử lý (Imputer + Scaler) -> Mô hình (Không dùng PCA)
    2. Tiền xử lý (Imputer + Scaler) -> PCA (k thành phần) -> Mô hình (Có PCA)
    Sử dụng 5-Fold Cross-Validation trên tập Train để đánh giá độ chính xác và độ ổn định.
    """
    print("\n" + "=" * 80)
    print(f"3. THỰC NGHIỆM SO SÁNH CÁC MÔ HÌNH HỒI QUY (5-FOLD CV TRÊN TẬP TRAIN, k_PCA = {k_pca})")
    print("=" * 80)

    cv = KFold(n_splits=5, shuffle=True, random_state=RANDOM_STATE)
    models = get_models()

    results = []

    for name, model in models.items():
        print(f"[*] Đang huấn luyện & đánh giá: {name}...")

        # 1. Pipeline KHÔNG DÙNG PCA
        pipe_no_pca = Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("scaler", StandardScaler()),
            ("regressor", model)
        ])
        res_no_pca = evaluate_pipeline(pipe_no_pca, X, y, cv)
        results.append({
            "Model": name,
            "PCA": "No (Full 8D)",
            "MAE_Mean": res_no_pca["mae_mean"],
            "MAE_Std": res_no_pca["mae_std"],
            "RMSE_Mean": res_no_pca["rmse_mean"],
            "RMSE_Std": res_no_pca["rmse_std"],
            "R2_Mean": res_no_pca["r2_mean"],
            "R2_Std": res_no_pca["r2_std"],
            "Fit_Time_Sec": res_no_pca["avg_fit_time"]
        })

        # 2. Pipeline CÓ DÙNG PCA
        pipe_with_pca = Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("scaler", StandardScaler()),
            ("pca", PCA(n_components=k_pca, random_state=RANDOM_STATE)),
            ("regressor", model)
        ])
        res_with_pca = evaluate_pipeline(pipe_with_pca, X, y, cv)
        results.append({
            "Model": name,
            "PCA": f"Yes (PCA {k_pca}D)",
            "MAE_Mean": res_with_pca["mae_mean"],
            "MAE_Std": res_with_pca["mae_std"],
            "RMSE_Mean": res_with_pca["rmse_mean"],
            "RMSE_Std": res_with_pca["rmse_std"],
            "R2_Mean": res_with_pca["r2_mean"],
            "R2_Std": res_with_pca["r2_std"],
            "Fit_Time_Sec": res_with_pca["avg_fit_time"]
        })

    df_results = pd.DataFrame(results)

    # In bảng tổng hợp kết quả
    print("\n" + "=" * 115)
    print("BẢNG TỔNG HỢP KẾT QUẢ ĐÁNH GIÁ (5-FOLD CROSS-VALIDATION)")
    print("=" * 115)
    header = f"{'Model':<25} | {'PCA':<15} | {'MAE (Mean ± Std)':<22} | {'RMSE (Mean ± Std)':<22} | {'R² (Mean ± Std)':<22} | {'Time (s)':<8}"
    print(header)
    print("-" * 115)
    for _, r in df_results.iterrows():
        mae_str = f"{r['MAE_Mean']:.4f} ± {r['MAE_Std']:.4f}"
        rmse_str = f"{r['RMSE_Mean']:.4f} ± {r['RMSE_Std']:.4f}"
        r2_str = f"{r['R2_Mean']:.4f} ± {r['R2_Std']:.4f}"
        print(f"{r['Model']:<25} | {r['PCA']:<15} | {mae_str:<22} | {rmse_str:<22} | {r2_str:<22} | {r['Fit_Time_Sec']:>7.2f}")
    print("=" * 115)

    csv_path = os.path.join(OUTPUT_DIR, "results_summary.csv")
    df_results.to_csv(csv_path, index=False)
    print(f"[✓] Đã xuất file CSV tổng kết tại: {csv_path}")

    return df_results


def select_best_model(df_results):
    """
    Lựa chọn mô hình dựa trên kết quả CV (chỉ dùng tập Train):
    xếp hạng mọi cấu hình (trừ Dummy Baseline) theo 4 tiêu chí
      - MAE_Mean  (thấp hơn tốt hơn)
      - RMSE_Mean (thấp hơn tốt hơn)
      - R2_Mean   (cao hơn tốt hơn)
      - R2_Std    (thấp hơn = ổn định hơn qua các fold)
    rồi chọn cấu hình có hạng trung bình nhỏ nhất (hòa thì ưu tiên RMSE_Mean thấp hơn).
    """
    print("\n" + "=" * 80)
    print("4. LỰA CHỌN MÔ HÌNH THEO MAE / RMSE / R² / ĐỘ ỔN ĐỊNH QUA CV")
    print("=" * 80)

    df_rank = df_results[df_results["Model"] != "Dummy Baseline"].copy()
    df_rank["Rank_MAE"] = df_rank["MAE_Mean"].rank(method="min")
    df_rank["Rank_RMSE"] = df_rank["RMSE_Mean"].rank(method="min")
    df_rank["Rank_R2"] = df_rank["R2_Mean"].rank(method="min", ascending=False)
    df_rank["Rank_Stability"] = df_rank["R2_Std"].rank(method="min")
    df_rank["Avg_Rank"] = df_rank[["Rank_MAE", "Rank_RMSE", "Rank_R2", "Rank_Stability"]].mean(axis=1)
    df_rank = df_rank.sort_values(by=["Avg_Rank", "RMSE_Mean"]).reset_index(drop=True)

    cols = ["Model", "PCA", "Rank_MAE", "Rank_RMSE", "Rank_R2", "Rank_Stability", "Avg_Rank"]
    print(df_rank[cols].to_string(index=False))

    rank_path = os.path.join(OUTPUT_DIR, "model_ranking.csv")
    df_rank[cols + ["MAE_Mean", "RMSE_Mean", "R2_Mean", "R2_Std"]].to_csv(rank_path, index=False)
    print(f"[✓] Đã xuất bảng xếp hạng mô hình tại: {rank_path}")

    best = df_rank.iloc[0]
    print(f"\n[★] Mô hình được chọn: {best['Model']} | {best['PCA']} "
          f"(hạng trung bình = {best['Avg_Rank']:.2f})")
    return best["Model"], best["PCA"].startswith("Yes")


def run_pca_dimension_sweep(X, y):
    """
    Thực nghiệm kiểm tra tác động của số chiều PCA k biến thiên từ 1 đến 8
    trên 3 mô hình tiêu biểu: Linear Regression, Random Forest, và HistGradientBoosting.
    Dùng đúng siêu tham số như thực nghiệm so sánh chính để kết quả có thể đối chiếu.
    """
    print("\n" + "=" * 80)
    print("5. KHẢO SÁT TÁC ĐỘNG CỦA SỐ LƯỢNG THÀNH PHẦN PCA (k = 1..8)")
    print("=" * 80)

    cv = KFold(n_splits=5, shuffle=True, random_state=RANDOM_STATE)
    all_models = get_models()
    models_to_test = {name: all_models[name]
                      for name in ["Linear Regression", "Random Forest", "HistGradientBoosting"]}

    n_features = X.shape[1]
    sweep_results = []

    for k in range(1, n_features + 1):
        print(f"[*] Đang khảo sát PCA với k = {k}...")
        for name, model in models_to_test.items():
            pipe = Pipeline([
                ("imputer", SimpleImputer(strategy="median")),
                ("scaler", StandardScaler()),
                ("pca", PCA(n_components=k, random_state=RANDOM_STATE)),
                ("regressor", model)
            ])
            res = evaluate_pipeline(pipe, X, y, cv)
            sweep_results.append({
                "k_components": k,
                "Model": name,
                "R2_Mean": res["r2_mean"],
                "R2_Std": res["r2_std"],
                "RMSE_Mean": res["rmse_mean"],
                "RMSE_Std": res["rmse_std"]
            })

    df_sweep = pd.DataFrame(sweep_results)
    df_sweep.to_csv(os.path.join(OUTPUT_DIR, "pca_sweep_results.csv"), index=False)

    # Vẽ biểu đồ PCA Component Sweep
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5.5))
    palette = {"Linear Regression": "#E74C3C", "Random Forest": "#27AE60", "HistGradientBoosting": "#2980B9"}

    for name in models_to_test.keys():
        sub = df_sweep[df_sweep["Model"] == name]
        ax1.plot(sub["k_components"], sub["R2_Mean"], marker='o', label=name, color=palette[name], linewidth=2)
        ax1.fill_between(sub["k_components"],
                         sub["R2_Mean"] - sub["R2_Std"],
                         sub["R2_Mean"] + sub["R2_Std"],
                         alpha=0.15, color=palette[name])

        ax2.plot(sub["k_components"], sub["RMSE_Mean"], marker='s', label=name, color=palette[name], linewidth=2)
        ax2.fill_between(sub["k_components"],
                         sub["RMSE_Mean"] - sub["RMSE_Std"],
                         sub["RMSE_Mean"] + sub["RMSE_Std"],
                         alpha=0.15, color=palette[name])

    ax1.set_title("Hệ số xác định R² theo số thành phần PCA (k)")
    ax1.set_xlabel("Số thành phần chính PCA (k)")
    ax1.set_ylabel("R² Score (Càng cao càng tốt)")
    ax1.set_xticks(range(1, n_features + 1))
    ax1.legend()

    ax2.set_title("Sai số RMSE theo số thành phần PCA (k)")
    ax2.set_xlabel("Số thành phần chính PCA (k)")
    ax2.set_ylabel("RMSE (Càng thấp càng tốt)")
    ax2.set_xticks(range(1, n_features + 1))
    ax2.legend()

    sweep_plot_path = os.path.join(OUTPUT_DIR, "pca_components_sweep.png")
    plt.savefig(sweep_plot_path, dpi=300)
    plt.close()
    print(f"[✓] Đã lưu biểu đồ khảo sát số chiều PCA tại: {sweep_plot_path}")

    return df_sweep


def build_pipeline(model, k_pca=None):
    """Pipeline chuẩn: Imputer -> Scaler -> [PCA] -> Regressor."""
    steps = [
        ("imputer", SimpleImputer(strategy="median")),
        ("scaler", StandardScaler()),
    ]
    if k_pca is not None:
        steps.append(("pca", PCA(n_components=k_pca, random_state=RANDOM_STATE)))
    steps.append(("regressor", model))
    return Pipeline(steps)


def plot_visualizations(df_results, k_pca):
    """
    Tạo các biểu đồ trực quan hóa kết quả CV trên tập Train:
    1. model_comparison.png: So sánh các mô hình trên R2, RMSE, MAE kèm sai số std (CV).
    2. pca_impact.png: Tác động của PCA trên từng mô hình.
    """
    print("\n" + "=" * 80)
    print("6. TRỰC QUAN HÓA KẾT QUẢ VÀ SO SÁNH MÔ HÌNH")
    print("=" * 80)

    # 1. So sánh tổng quan (Không có PCA)
    df_no_pca = df_results[df_results["PCA"].str.startswith("No")].copy()
    df_no_pca = df_no_pca.sort_values(by="R2_Mean", ascending=True)

    fig, axes = plt.subplots(1, 3, figsize=(18, 6))

    # Bar chart R2
    axes[0].barh(df_no_pca["Model"], df_no_pca["R2_Mean"], xerr=df_no_pca["R2_Std"],
                 color='#2ecc71', alpha=0.85, capsize=4, edgecolor='black', linewidth=0.7)
    axes[0].set_title("R² Score (Cao hơn là tốt hơn)")
    axes[0].set_xlabel("R² Mean ± CV Std")

    # Bar chart RMSE
    axes[1].barh(df_no_pca["Model"], df_no_pca["RMSE_Mean"], xerr=df_no_pca["RMSE_Std"],
                 color='#e74c3c', alpha=0.85, capsize=4, edgecolor='black', linewidth=0.7)
    axes[1].set_title("RMSE (Thấp hơn là tốt hơn)")
    axes[1].set_xlabel("RMSE Mean ± CV Std")

    # Bar chart MAE
    axes[2].barh(df_no_pca["Model"], df_no_pca["MAE_Mean"], xerr=df_no_pca["MAE_Std"],
                 color='#3498db', alpha=0.85, capsize=4, edgecolor='black', linewidth=0.7)
    axes[2].set_title("MAE (Thấp hơn là tốt hơn)")
    axes[2].set_xlabel("MAE Mean ± CV Std")

    comp_plot_path = os.path.join(OUTPUT_DIR, "model_comparison.png")
    plt.savefig(comp_plot_path, dpi=300)
    plt.close()
    print(f"[✓] Đã lưu biểu đồ so sánh mô hình tại: {comp_plot_path}")

    # 2. Biểu đồ tác động của PCA: So sánh No PCA vs With PCA
    fig, axes = plt.subplots(1, 2, figsize=(16, 6))

    model_names = df_results["Model"].unique()
    x_indices = np.arange(len(model_names))
    width = 0.38

    sub_no = df_results[df_results["PCA"].str.startswith("No")].set_index("Model").loc[model_names]
    sub_yes = df_results[df_results["PCA"].str.startswith("Yes")].set_index("Model").loc[model_names]
    label_no = "Không dùng PCA (8 chiều)"
    label_yes = f"Có dùng PCA ({k_pca} chiều)"

    # R2 Comparison
    axes[0].bar(x_indices - width/2, sub_no["R2_Mean"], width, yerr=sub_no["R2_Std"],
                capsize=3, label=label_no, color="#3498db", alpha=0.9)
    axes[0].bar(x_indices + width/2, sub_yes["R2_Mean"], width, yerr=sub_yes["R2_Std"],
                capsize=3, label=label_yes, color="#e67e22", alpha=0.9)
    axes[0].set_ylabel("R² Score")
    axes[0].set_title("Tác động của PCA lên Hệ số xác định R²")
    axes[0].set_xticks(x_indices)
    axes[0].set_xticklabels(model_names, rotation=35, ha="right")
    axes[0].legend()

    # RMSE Comparison
    axes[1].bar(x_indices - width/2, sub_no["RMSE_Mean"], width, yerr=sub_no["RMSE_Std"],
                capsize=3, label=label_no, color="#3498db", alpha=0.9)
    axes[1].bar(x_indices + width/2, sub_yes["RMSE_Mean"], width, yerr=sub_yes["RMSE_Std"],
                capsize=3, label=label_yes, color="#e67e22", alpha=0.9)
    axes[1].set_ylabel("RMSE ($100k)")
    axes[1].set_title("Tác động của PCA lên Sai số RMSE")
    axes[1].set_xticks(x_indices)
    axes[1].set_xticklabels(model_names, rotation=35, ha="right")
    axes[1].legend()

    pca_impact_path = os.path.join(OUTPUT_DIR, "pca_impact.png")
    plt.savefig(pca_impact_path, dpi=300)
    plt.close()
    print(f"[✓] Đã lưu biểu đồ tác động PCA tại: {pca_impact_path}")


def evaluate_on_test(X_train, X_test, y_train, y_test, best_name, best_use_pca, k_pca):
    """
    Đánh giá cuối cùng trên tập Test hold-out (chưa từng được dùng ở các bước trước):
    - Dummy Baseline
    - Mô hình được chọn qua CV (đúng cấu hình đã chọn)
    - Cùng mô hình đó ở cấu hình PCA còn lại (để kiểm chứng tác động của PCA trên dữ liệu mới)
    Mỗi pipeline được huấn luyện lại trên TOÀN BỘ tập Train rồi dự báo trên tập Test.
    """
    print("\n" + "=" * 80)
    print("7. ĐÁNH GIÁ CUỐI CÙNG TRÊN TẬP TEST HOLD-OUT")
    print("=" * 80)

    models = get_models()
    configs = [
        ("Dummy Baseline", None, "Baseline"),
        (best_name, k_pca if best_use_pca else None, "Mô hình được chọn"),
        (best_name, None if best_use_pca else k_pca, "Đối chứng PCA"),
    ]

    rows = []
    predictions = {}
    for name, k, role in configs:
        pipe = build_pipeline(clone(models[name]), k)
        pipe.fit(X_train, y_train)
        preds = pipe.predict(X_test)
        pca_label = f"Yes (PCA {k}D)" if k is not None else "No (Full 8D)"
        rows.append({
            "Role": role,
            "Model": name,
            "PCA": pca_label,
            "MAE": mean_absolute_error(y_test, preds),
            "RMSE": root_mean_squared_error(y_test, preds),
            "R2": r2_score(y_test, preds),
        })
        predictions[role] = preds

    df_test = pd.DataFrame(rows)
    print(df_test.to_string(index=False, float_format=lambda v: f"{v:.4f}"))
    test_path = os.path.join(OUTPUT_DIR, "test_results.csv")
    df_test.to_csv(test_path, index=False)
    print(f"[✓] Đã xuất kết quả tập Test tại: {test_path}")

    # Biểu đồ Actual vs Predicted trên tập Test
    y_pred_dummy = predictions["Baseline"]
    y_pred_best = predictions["Mô hình được chọn"]
    best_label = f"{best_name} ({'PCA ' + str(k_pca) + 'D' if best_use_pca else 'không PCA'})"

    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    for ax, preds, title, color, alpha in [
        (axes[0], y_pred_dummy, "Baseline (Dummy Mean)", '#e74c3c', 0.25),
        (axes[1], y_pred_best, f"Mô hình được chọn: {best_label}", '#2980b9', 0.3),
    ]:
        ax.scatter(y_test, preds, alpha=alpha, color=color, s=15)
        ax.plot([y_test.min(), y_test.max()], [y_test.min(), y_test.max()], 'k--', lw=2,
                label="Đường lý tưởng (y = y_pred)")
        ax.set_title(f"{title} — Tập Test\nMAE = {mean_absolute_error(y_test, preds):.3f}, "
                     f"R² = {r2_score(y_test, preds):.3f}")
        ax.set_xlabel("Giá trị thực tế ($100k)")
        ax.set_ylabel("Giá trị dự báo ($100k)")
        ax.legend()

    actual_pred_path = os.path.join(OUTPUT_DIR, "actual_vs_predicted.png")
    plt.savefig(actual_pred_path, dpi=300)
    plt.close()
    print(f"[✓] Đã lưu biểu đồ Actual vs Predicted (tập Test) tại: {actual_pred_path}")

    return df_test


def main():
    print("=" * 80)
    print("BẮT ĐẦU CHẠY CHƯƠNG TRÌNH PIPELINE HỒI QUY VỚI TIỀN XỬ LÝ VÀ PCA")
    print("=" * 80)
    total_t0 = time.time()

    # Bước 1: Load dữ liệu & EDA
    X, y = load_and_explore_data()

    # Bước 2: Tách tập Test hold-out (20%) — chỉ dùng ở bước đánh giá cuối cùng
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=TEST_SIZE, random_state=RANDOM_STATE)
    print(f"\n[*] Chia dữ liệu: Train = {len(X_train)} mẫu, Test hold-out = {len(X_test)} mẫu")

    # Bước 3: Phân tích PCA Variance (trên tập Train)
    k_pca = analyze_pca_variance(X_train)

    # Bước 4: Đánh giá so sánh các mô hình (có và không có PCA) bằng CV trên tập Train
    df_results = run_model_comparison(X_train, y_train, k_pca=k_pca)

    # Bước 5: Lựa chọn mô hình theo MAE/RMSE/R²/độ ổn định
    best_name, best_use_pca = select_best_model(df_results)

    # Bước 6: Khảo sát số chiều PCA k biến thiên từ 1 đến 8 (trên tập Train)
    run_pca_dimension_sweep(X_train, y_train)

    # Bước 7: Trực quan hóa kết quả CV
    plot_visualizations(df_results, k_pca)

    # Bước 8: Đánh giá cuối cùng trên tập Test hold-out
    evaluate_on_test(X_train, X_test, y_train, y_test, best_name, best_use_pca, k_pca)

    print("\n" + "=" * 80)
    print(f"HOÀN THÀNH TOÀN BỘ QUY TRÌNH TRONG: {time.time() - total_t0:.2f} GIÂY")
    print(f"Toàn bộ kết quả, bảng biểu và đồ thị đã được lưu trữ trong thư mục: {OUTPUT_DIR}")
    print("=" * 80)


if __name__ == "__main__":
    main()
