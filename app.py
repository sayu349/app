import os
from flask import Flask
from flask import Flask, render_template


app = Flask(__name__)

#Flaskでカラムの候補を出力する
import pandas as pd
import numpy as np
from flask import Flask,render_template

def read_excel(filename="sample.xlsx"):
    df = pd.read_excel(filename)
    return df.columns

data_list=read_excel()

@app.route("/")
def index():
    return render_template("index.html", data_list=data_list)