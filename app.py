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

#------------------------------------------------------

# 必要なし？
# excelとcsvを見分けて、ファイルを読み込む
def read_excel(file,filename):
    # csvの場合
    if filename[-1] == 'v':
        df = pd.read_csv(file)
    # excelの場合
    elif filename[-1] == 'x':
        df = pd.read_excel(file)
    return df.columns

#------------------------------------------------------
# 必要なし？
# カラムが全行含まれているかチェックする
def check_columns(file,filename):
    # csvの場合（最後の文字がvのファイル）
    if filename[-1] == 'v':
        df = pd.read_csv(file)
    # excelの場合（最後の文字がxのファイル）
    elif filename[-1] == 'x':
        df = pd.read_excel(file)
    # カラム（1行目）の要素が空であることを意味する「Unnamed」の個数を数える
    return df.T.query("index.str.contains('Unnamed')",engine="python")

#------------------------------------------------------

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

#------------------------------------------------------

# エクセル読み込みページ
# http://127.0.0.1:5000/
# import.html
@app.route("/")
def index():
    return render_template("import.html")

#------------------------------------------------------

# 変動パラメータ設定ページ
# http://127.0.0.1:5000/detail-option
# detail-opiton.html
@app.route("/detail-option", methods=["POST"])
def column_search():
    file = request.files['upload-file']
    file_title = file.filename
    print(file_title[-1])

    # csvの場合
    if file_title[-1] == "v":
        file.save("uploads/upload_file.csv")
        csv = pd.read_csv("uploads/upload_file.csv")
        csv.to_excel("uploads/upload_file.xlsx")
        excel_data = pd.read_excel("uploads/upload_file.xlsx")
        column = excel_data.columns

    # excelの場合
    else:
        file.save("uploads/upload_file.xlsx")
        excel_data = pd.read_excel("uploads/upload_file.xlsx")
        column = excel_data.columns
        
    return render_template("detail-option.html",
                            data_list=column,
                            file_title=file_title)

#------------------------------------------------------

# 結果ダウンロードページ
# http://127.0.0.1:5000/result
# result.html
@app.route("/result", methods=["POST"])
def calc_result():
    #------------------------------------------------------
    
    # 金額列定義
    sample_data = pd.read_excel("uploads/upload_file.xlsx")
    # request.form['sample'] = "金額列"
    amount = request.form['sample']
    # 母集団の金額が正しいかチェック
    total_amount = sample_data[request.form['sample']].sum()

    #------------------------------------------------------
    
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
    # サンプリングシートに記載用の、パラメータ一覧
    sampling_param = pd.DataFrame([['母集団合計', N],
                                ['手続実施上の重要性', pm],
                                ['リスク', audit_risk],
                                ['内部統制', internal_control],
                                ['random_state', random_state]])

    #------------------------------------------------------
    
    # 母集団をまずは降順に並び替える（ここで並び替えるのは、サンプル出力の安定のため安定のため）
    sample_data = sample_data.sort_values(amount, ascending=False)
    # 母集団をシャッフル
    shuffle_data = sample_data.sample(frac=1, random_state=random_state) #random_stateを使って乱数を固定化する
     # サンプリング区間の算定
    m = N/n

    #------------------------------------------------------

    # 列の追加
    shuffle_data['cumsum'] = shuffle_data[amount].cumsum() # 積み上げ合計
    shuffle_data['group'] = shuffle_data['cumsum']//m # サンプルのグループ化

    result_data = shuffle_data.loc[shuffle_data.groupby('group')['cumsum'].idxmin(), ]

    #------------------------------------------------------

    # 保存先ディレクトリ
    file_name = "result/result.xlsx"
    # シート呼び出し
    writer = pd.ExcelWriter(file_name)
    # エクセルファイルに計算したデータを追加していく
    ## 全レコードを'全体'シートに出力
    sample_data.to_excel(writer, sheet_name = '母集団', index=False)
    ## サンプリング結果を、サンプリングシートに記載
    result_data.to_excel(writer, sheet_name = 'サンプリング結果', index=False)
    ## サンプリングの情報追記
    sampling_param.to_excel(writer, sheet_name = 'サンプリングパラメータ', index=False, header=None)
    ## Excelファイルを保存
    writer.save()
    ## Excelファイルを閉じる
    writer.close()

    #------------------------------------------------------

    return render_template("result.html")

#------------------------------------------------------

# ファイルを保存する
# http://127.0.0.1:5000/result で出力する
# result.html
@app.route("/resultsave")
def export_action():
    return send_file('result/result.xlsx')