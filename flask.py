from flask import Flask, render_template
import pandas as pd

app = Flask(__name__)

@app.route('/')
def index():
    data = pd.read_excel("assignments.xlsx")
    return render_template("canvas.html", data=data.to_html())

if __name__ == '__main__':
    app.run(debug=True)