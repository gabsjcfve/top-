from flask import Flask, render_template, request, send_file
import pandas as pd
import io
import requests

app = Flask(__name__)

# Chave da API do Gemini
API_KEY = "AIzaSyBzwbCvx_LMKbGu3OiVmJzveXmW25Hfuk0"

def classificar_acao_gemini(descricao):
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"
    headers = {'Content-Type': 'application/json'}
    data = {
        "contents": [{
            "parts": [{"text": f"Classifique a seguinte ação de manutenção em uma das seguintes categorias: Concluída, Em acompanhamento, Liberada por operador, Falta Informação, Pendente. Retorne apenas a categoria, sem nenhuma explicação adicional. Descrição da ação: {descricao}"}]
        }]
    }
    response = requests.post(url, headers=headers, json=data, params={"key": API_KEY})
    if response.status_code == 200:
        return response.json()['candidates'][0]['content']['parts'][0]['text'].split(":")[-1].strip()
    else:
        return None

def completar_classificacoes(acoes_manutencao):
    acoes_manutencao['Classificação'] = acoes_manutencao.apply(
        lambda row: classificar_acao_gemini(row['Descrição']),
        axis=1
    )
    return acoes_manutencao

@app.route('/')
def index():
    return render_template('index.html', message=None)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return render_template('index.html', message='Nenhum arquivo selecionado.')

    file = request.files['file']
    if file.filename == '':
        return render_template('index.html', message='Nenhum arquivo selecionado.')

    if file:
        excel_data = pd.ExcelFile(file)
        acoes_manutencao = pd.read_excel(excel_data, 'Ações de Manutenção')
        status_acoes = pd.read_excel(excel_data, 'Status das Ações')

        acoes_manutencao = completar_classificacoes(acoes_manutencao)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            acoes_manutencao.to_excel(writer, sheet_name='Ações de Manutenção', index=False)
            status_acoes.to_excel(writer, sheet_name='Status das Ações', index=False)
        output.seek(0)

        return send_file(output, download_name='planilha_atualizada.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
