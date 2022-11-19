import os
from flask import Flask



app = Flask(__name__)

#Flaskでカラムの候補を出力する
import pandas as pd
import numpy as np
from flask import Flask,render_template, request

def read_excel(filename="sample.xlsx"):
    df = pd.read_excel(filename)
    return df.columns

data_list=read_excel()
@app.route("/")
def index():
    return render_template("index.html", data_list=data_list)


@app.route("/sampleform-post",methods=["POST"])
def sampleform():
    print("受け取りました。")
    amount = request.form["pm"]
    risk = request.form["audit_risk"]
    file = request.form["upload-file"]
    print(request.form)
    print(amount)
    print(risk)
    print(file)
    print(type(file))
    return f'受け取りました:{amount}'