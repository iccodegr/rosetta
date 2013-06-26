from rosetta import convert
from flask import Flask, request

app = Flask(__name__)

@app.route('/convert', methods=['POST'])
def convert():
    data = request.form['data']
    from_type = request.form['from_type']
    to_type = request.form['to_type']
    return convert(data, from_type, to_type)

