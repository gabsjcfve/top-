from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import io
import requests
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

# Certifique-se de que a pasta de uploads existe
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Chave da API do Gemini
API_KEY = "SUA_CHAVE_AQUI"

# Função para classificar ações usando a API do Gemini
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
        return "Erro na classificação"

# Função para processar a planilha
def processar_planilha(filepath):
    excel_data = pd.ExcelFile(filepath)
    acoes_manutencao = pd.read_excel(excel_data, 'Ações de Manutenção')
    acoes_manutencao['Classificação'] = acoes_manutencao['Descrição'].apply(classificar_acao_gemini)
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'planilha_classificada.xlsx')
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        acoes_manutencao.to_excel(writer, sheet_name='Ações de Manutenção', index=False)
    return output_path

# Página inicial
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            output_path = processar_planilha(filepath)
            return redirect(url_for('download_file', filename='planilha_classificada.xlsx'))
    return render_template('upload.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
