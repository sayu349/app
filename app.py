# Flaskライブラリ
from flask import Flask
from flask import render_template
from flask import request
from flask import send_file


# 統計等のライブラリ
import numpy as np
import pandas as pd
import scipy as sp
from scipy.stats import poisson
from scipy.stats import binom
import math

# ディレクトリ用のライブラリ
import os

# 可視化ライブラリ
import matplotlib.pyplot as plt
import seaborn as sns

# おまじない
app = Flask(__name__)


# ------------------------------------------------------

# データを読み込んでカラムを取得する
def columns_search(input_file, save_file_path, file_name):
    # .xlsx読み込み → データ保存 → カラム取得
    if file_name[-4:] == "xlsx":
        import_excel_df = pd.read_excel(input_file)
        print(type(input_file))
        import_excel_df.to_excel(save_file_path)
        columns = import_excel_df.columns
    # .csv読み込み → .xlsxとして保存 → .xlsx読み込み → カラム取得
    elif file_name[-3:] == "csv":
        import_csv = pd.read_csv(input_file)
        import_csv.to_excel(save_file_path)
        import_excel_df = pd.read_excel(save_file_path)
        columns = import_excel_df.columns
    return columns


# ------------------------------------------------------

# ポアソン分布による金額単位サンプリングによるサンプル数算定の関数
def sample_poisson(N, pm, ke, alpha, audit_risk, internal_control="依拠しない"):
    k = np.arange(ke + 1)
    pt = pm / N
    n = 1
    while True:
        mu = n * pt
        pmf_poi = poisson.cdf(k, mu)
        if pmf_poi.sum() < alpha:
            break
        n += 1
    if audit_risk == "SR":
        n = math.ceil(n)
    if audit_risk == "RMM-L":
        n = math.ceil(n / 10 * 2)
    if audit_risk == "RMM-H":
        n = math.ceil(n / 2)
    if internal_control == "依拠する":
        n = math.ceil(n / 3)
    return n


# ------------------------------------------------------

# エクセル読み込みページ
# http://127.0.0.1:5000/
# import.html
@app.route("/")
def index():
    return render_template("import.html")


# ------------------------------------------------------

# 変動パラメータ設定ページ
# http://127.0.0.1:5000/detail-option
# detail-opiton.html
@app.route("/detail-option", methods=["POST"])
def column_search():
    # uploadしたフォルダの保存先
    upload_path = "uploads/upload_file.xlsx"
    # requst.files['upload-file']でHTML側で入力したシートの情報を取得
    file = request.files["upload-file"]
    # ファイル名を取得（拡張子を見分ける為）
    file_title = file.filename

    # csvの場合
    if file_title[-3:] == "csv":
        # 以降の処理が.xlsxデータなので、csvをxlsx形式に変換する
        columns = columns_search(file, upload_path, file_title)
        return render_template(
            "detail-option.html",
            # 以下htmlに飛ばすデータ
            # html側の変数名 = py側の変数名
            columns=columns,
            file_title=file_title,
        )

    # xlsxの場合
    elif file_title[-4:] == "xlsx":
        columns = columns_search(file, upload_path, file_title)
        return render_template(
            "detail-option.html", columns=columns, file_title=file_title
        )

    # xlsの場合
    elif file_title[-3:] == "xls":
        columns = columns_search(file, upload_path, file_title)
        return render_template(
            "detail-option.html", columns=columns, file_title=file_title
        )

    # xlsx,csv,xls以外が出力された場合
    else:
        error_text = "入力されたファイルには対応していません。拡張子が.xlsx,.xls,.xlsのもので試してください。"
        return render_template("error.html", error_text=error_text)


# ------------------------------------------------------

# 結果ダウンロードページ
# http://127.0.0.1:5000/result
# result.html
@app.route("/result", methods=["POST"])
def calc_result():
    # ------------------------------------------------------

    sample_data = pd.read_excel("uploads/upload_file.xlsx")
    # request.form['sample'] = "金額列"
    amount = request.form["amount_column"]
    # 母集団の金額が正しいかチェック
    total_amount = sample_data[amount].sum()

    # ------------------------------------------------------

    # 変動パラメータの設定
    # 母集団の金額合計
    N = total_amount
    # 手続実施上の重要性
    pm = int(request.form["pm"])
    # ランダムシード　(サンプリングの並び替えのステータスに利用、任意の数を入力)
    random_state = int(request.form["random_state"])
    # 監査リスク
    audit_risk = request.form["audit_risk"]
    # 内部統制
    internal_control = request.form["internal_control"]
    # 予想虚偽表示金額（変更不要）
    ke = 0
    alpha = 0.05
    # サンプルサイズnの算定
    n = sample_poisson(N, pm, ke, alpha, audit_risk, internal_control)
    # サンプリングシートに記載用の、パラメータ一覧
    sampling_param = pd.DataFrame(
        [
            ["母集団合計", N],
            ["手続実施上の重要性", pm],
            ["リスク", audit_risk],
            ["内部統制", internal_control],
            ["random_state", random_state],
        ]
    )

    # ------------------------------------------------------

    # 母集団をまずは降順に並び替える（ここで並び替えるのは、サンプル出力の安定のため安定のため）
    sample_data = sample_data.sort_values(amount, ascending=False)
    # 母集団をシャッフル
    shuffle_data = sample_data.sample(
        frac=1, random_state=random_state
    )  # random_stateを使って乱数を固定化する
    # サンプリング区間の算定
    m = N / n

    # ------------------------------------------------------

    # 列の追加
    shuffle_data["cumsum"] = shuffle_data[amount].cumsum()  # 積み上げ合計
    shuffle_data["group"] = shuffle_data["cumsum"] // m  # サンプルのグループ化

    result_data = shuffle_data.loc[
        shuffle_data.groupby("group")["cumsum"].idxmin(),
    ]

    # ------------------------------------------------------

    # 保存先ディレクトリ
    file_name = "result/result.xlsx"
    # シート呼び出し
    writer = pd.ExcelWriter(file_name)
    # エクセルファイルに計算したデータを追加していく
    ## 全レコードを'全体'シートに出力
    sample_data.to_excel(writer, sheet_name="母集団", index=False)
    ## サンプリング結果を、サンプリングシートに記載
    result_data.to_excel(writer, sheet_name="サンプリング結果", index=False)
    ## サンプリングの情報追記
    sampling_param.to_excel(writer, sheet_name="サンプリングパラメータ", index=False, header=None)
    ## Excelファイルを保存
    writer.save()
    ## Excelファイルを閉じる
    writer.close()

    # ------------------------------------------------------

    return render_template("result.html")


# ------------------------------------------------------

# ファイルを保存する
# http://127.0.0.1:5000/result で出力する
# result.html
@app.route("/resultsave")
def export_action():
    return send_file("result/result.xlsx")
