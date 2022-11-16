import os
from flask import Flask, flash, request, redirect, url_for
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'images'
ALLOWED_EXTENSIONS = {'xlsx','csv'}

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

