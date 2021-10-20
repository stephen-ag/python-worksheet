import json
from flask import Flask,render_template, request, jsonify
import pandas as pd
# need to install xlrd and openpyxl to use read_excel function in pandas
import openpyxl

app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')

@app.route('/data', methods=['GET', 'POST'])

def data():
    if request.method == 'POST':
        file = request.form['upload-file']
        data1 = pd.read_excel(file)
        dat1 = pd.DataFrame(data1)
        df = dat1.head()
        return render_template('data.html', data=df.to_html())

if __name__   ==   '__main__':
    app.run(debug=True)