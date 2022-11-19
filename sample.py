import os
from flask import Flask
from flask import render_template
from flask import request
import math
import pandas as pd
import numpy as np


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

# 金額列調整方法


def read_excel(filename):
    df = pd.read_excel(filename)
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


@app.route("/")
def index():
    return render_template("import-menu.html")


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


@app.route("/result", methods=["POST"])
def calc_result():
    print(request.form)
    # シート内で列名を指定して、金額合計を算出
    df = pd.read_excel("uploads/upload_file.xlsx")
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
    print(n)
    return render_template("result.html",n=n)
