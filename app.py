import os
import pandas as pd
import shutil
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import zipfile

app = Flask(__name__)

# Configurações de upload
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['PROCESSED_FOLDER'] = 'processed'
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # Limite de 500MB

# Cria as pastas se não existirem
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Verifica se os arquivos foram enviados
        if 'planilha' not in request.files or 'fotos_zip' not in request.files:
            return "Por favor, envie a planilha e a pasta de fotos compactada.", 400

        planilha_file = request.files['planilha']
        fotos_file = request.files['fotos_zip']

        # Salva os arquivos enviados
        planilha_filename = secure_filename(planilha_file.filename)
        fotos_filename = secure_filename(fotos_file.filename)

        planilha_path = os.path.join(app.config['UPLOAD_FOLDER'], planilha_filename)
        fotos_zip_path = os.path.join(app.config['UPLOAD_FOLDER'], fotos_filename)

        planilha_file.save(planilha_path)
        fotos_file.save(fotos_zip_path)

        # Extrai as fotos do arquivo zip
        fotos_extracted_path = os.path.join(app.config['UPLOAD_FOLDER'], 'fotos_extracted')
        os.makedirs(fotos_extracted_path, exist_ok=True)

        with zipfile.ZipFile(fotos_zip_path, 'r') as zip_ref:
            zip_ref.extractall(fotos_extracted_path)

        # Chama a função de processamento
        resultado = processar_arquivos(planilha_path, fotos_extracted_path)

        if resultado:
            # Compacta a pasta de fotos processadas
            output_zip_path = os.path.join(app.config['PROCESSED_FOLDER'], 'fotos_renomeadas.zip')
            shutil.make_archive(output_zip_path.replace('.zip', ''), 'zip', fotos_extracted_path)

            # Envia o arquivo zip para o usuário
            return send_file(output_zip_path, as_attachment=True)
        else:
            return "Ocorreu um erro durante o processamento.", 500

    return render_template('index.html')

def processar_arquivos(planilha_caminho, pasta_fotos_caminho):
    try:
        if not os.path.exists(planilha_caminho):
            raise FileNotFoundError(f"Planilha não encontrada: {planilha_caminho}")

        if not os.path.exists(pasta_fotos_caminho):
            raise FileNotFoundError(f"Pasta de fotos não encontrada: {pasta_fotos_caminho}")

        df = pd.read_excel(planilha_caminho, dtype=str)

        if 'Foto' not in df.columns or 'Equipamento' not in df.columns:
            raise ValueError("As colunas 'Foto' e 'Equipamento' não estão na planilha")

        def generate_next_filename(base_name, existing_files):
            suffix = 'b'
            new_filename = f"{base_name} {suffix}.jpg"
            while new_filename in existing_files:
                suffix = chr(ord(suffix) + 1)
                new_filename = f"{base_name} {suffix}.jpg"
            return new_filename

        existing_files = set(os.listdir(pasta_fotos_caminho))

        for index, row in df.iterrows():
            nome_atual = row['Foto']
            nova_numeracao = row['Equipamento']

            if pd.isna(nome_atual) or not isinstance(nome_atual, str) or not nome_atual.strip():
                print(f"Nome atual inválido ou vazio na linha {index + 1}")
                continue

            nome_atual = nome_atual.strip()
            nova_numeracao = nova_numeracao.strip()

            if '...' in nome_atual:
                try:
                    start, end = map(int, nome_atual.split('...'))
                except ValueError:
                    print(f"Intervalo inválido na linha {index + 1}: {nome_atual}")
                    continue
                nome_atual_range = range(start, end + 1)
            else:
                try:
                    nome_atual_range = [int(nome_atual)]
                except ValueError:
                    print(f"Erro na linha {index + 1}: {nome_atual}")
                    continue

            for num in nome_atual_range:
                num_str = str(num)
                encontrado = False

                for arquivo in existing_files:
                    if num_str in arquivo:
                        caminho_atual = os.path.join(pasta_fotos_caminho, arquivo)
                        novo_nome = f"{nova_numeracao}.jpg"
                        novo_caminho = os.path.join(pasta_fotos_caminho, novo_nome)

                        if os.path.exists(novo_caminho):
                            novo_caminho = os.path.join(pasta_fotos_caminho, generate_next_filename(nova_numeracao, existing_files))

                        shutil.copy2(caminho_atual, novo_caminho)

                        existing_files.add(os.path.basename(novo_caminho))

                        print(f"Duplicado: {caminho_atual} -> {novo_caminho}")
                        encontrado = True
                        break

                if not encontrado:
                    print(f"Arquivo não encontrado para o número: {num}")

        print("Renomeação concluída.")
        return True
    except Exception as e:
        print(f"Erro durante o processamento: {e}")
        return False

if __name__ == '__main__':
    app.run(debug=True)
