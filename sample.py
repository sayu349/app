# Flaskライブラリ
from flask import Flask
from flask import render_template
from flask import request


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

# カラム行特定（行の空欄が0の行を探してくる）
def read_excel(filename):
    df = pd.read_excel(filename)
    row_count = 0
    while len(df.T.query("index.str.contains('Unnamed')",engine="python"))!=0:
        amount_search_file = pd.read_excel(filename,skiprows = row_count)
        df = amount_search_file
        row_count = row_count + 1
    return df.columns


# ポアソン分布による金額単位サンプリングによるサンプル数算定の関数
def sample_poisson(N, pm, ke, alpha, audit_risk, internal_control='依拠しない'):
    k = np.arange(ke+1)
    pt = pm/N
    n = 1
    while True:
        mu = n*pt
        pmf_poi = poisson.cdf(k, mu)
        if pmf_poi.sum() < alpha:
            break
        n += 1
    if audit_risk == 'SR':
        n = math.ceil(n)
    if audit_risk == 'RMM-L':
        n = math.ceil(n/10*2)
    if audit_risk == 'RMM-H':
        n = math.ceil(n/2)
    if internal_control == '依拠する':
        n = math.ceil(n/3)
    return n

# エクセル読み込みページ
@app.route("/")
def index():
    return render_template("import-menu.html")

# 変動パラメータ設定ページ
@app.route("/sampleform-post", methods=["POST"])
def column_search():
    file = request.files['upload-file']
    # データを保持したい場合は以下のコードを回す(オンライン上へデプロイする際には要検討)
    file.save("uploads/upload_file.xlsx")
    file_title = file.filename
    data_list = read_excel(request.files['upload-file'])
    return render_template("column_search_result.html",
                           data_list=data_list,
                           file_title=file_title)

# 結果ダウンロードページ
@app.route("/result", methods=["POST"])
def calc_result():
    print(request.form)
    # 金額列定義
    amount = request.form['sample']

    # シート内で列名を指定して、金額合計を算出
    df = pd.read_excel("uploads/upload_file.xlsx")
    sample_data = pd.read_excel("uploads/upload_file.xlsx")
    total_amount = df[request.form['sample']].sum()
    # 変動パラメータの設定
    # 母集団の金額合計
    N = total_amount
    # 手続実施上の重要性
    pm = int(request.form['pm'])
    # ランダムシード　(サンプリングの並び替えのステータスに利用、任意の数を入力)
    random_state = int(request.form['random_state'])
    # 監査リスク
    audit_risk = request.form['audit_risk']
    # 内部統制
    internal_control = request.form['internal_control']
    # 予想虚偽表示金額（変更不要）
    ke = 0
    alpha = 0.05
    # サンプルサイズnの算定
    n = sample_poisson(N, pm, ke, alpha, audit_risk, internal_control)

    # 母集団をまずは降順に並び替える（ここで並び替えるのは、サンプル出力の安定のため安定のため）
    sample_data = sample_data.sort_values(amount, ascending=False)

    # 母集団をシャッフル
    shuffle_data = sample_data.sample(frac=1, random_state=random_state) #random_stateを使って乱数を固定化する

     # サンプリング区間の算定
    m = N/n
    print(m)


    # 列の追加
    shuffle_data['cumsum'] = shuffle_data[amount].cumsum() # 積み上げ合計
    shuffle_data['group'] = shuffle_data['cumsum']//m # サンプルのグループ化

    result_data = shuffle_data.loc[shuffle_data.groupby('group')['cumsum'].idxmin(), ]

    #保存先ディレクトリ指定
    file_name = "result/result.xlsx"
    # result_data.to_excel(file_name, encoding="shift_jis", index=False)
    
    writer = pd.ExcelWriter(file_name)

    # 全レコードを'全体'シートに出力
    sample_data.to_excel(writer, sheet_name = '母集団', index=False)
    # サンプリング結果を、サンプリングシートに記載
    result_data.to_excel(writer, sheet_name = 'サンプリング結果', index=False)
    # サンプリングの情報追記
    #sampling_param.to_excel(writer, sheet_name = 'サンプリングパラメータ', index=False, header=None)
    # Excelファイルを保存
    writer.save()
    # Excelファイルを閉じる
    writer.close()

    return render_template("result.html", n=n)
