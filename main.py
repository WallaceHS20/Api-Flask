from flask import Flask, make_response, jsonify, request
from bd import usuarios

app = Flask(__name__)

@app.route('/logs', methods=['GET'])
def get_logs():
    return make_response(
        jsonify(usuarios)
    )

@app.route('/logs', methods=['POST'])
def create_log():
    usuario = request.json
    usuarios.append(usuario)
    # faça algo com o objeto usuario
    return make_response(

        jsonify(
        
        mensagem = "Usuario Registrado com Sucesso!",
         usuario = usuario) 
        
        )

app.run()