## このアプリの流れ
- import.html("/")で、データを入力する。
    - カラムが全列入力されていなければ、エラーが出るようにする。
- detail-option.html("/detail-option")でカラム列や計算方法を指定する。
- result.html("/result")で完成したデータをダウンロードする

## 利用方法
### flask起動方法(windows)
- コマンドプロンプトを起動
- cd /app.pyのある場所まで移動
- set FLASK_APP=app.py
- set FLASK_ENV=development
- flask run
- 出力されたURLをCTRL+クリックでアプリ起動

### コマンドプロンプト以外の起動方法は以下を参照（公式）
- https://msiz07-flask-docs-ja.readthedocs.io/ja/latest/quickstart.html#a-minimal-application