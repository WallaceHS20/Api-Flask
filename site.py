from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/', methods=["GET", "POST"])
def index():

    variavel = "game do número"
    if request.method == "GET":
        return render_template("index.html", variavel=variavel)
    else:

        numero = 0
        palpite = int(request.form.get("name"))

        if numero == palpite:
            return '<h1>Ganhou</h1>'

@app.route('/<string:nome>')
def error(nome):
    return f'<h1> Ola ({nome})'

app.run()