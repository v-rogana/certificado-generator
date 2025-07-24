# app.py
import os
import io
import base64
import zipfile
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'

# Criar diretórios necessários
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('temp_certificates', exist_ok=True)

class CertificateGenerator:
    def __init__(self, background_image=None):
        self.background_image = background_image
        self.width, self.height = landscape(A4)
        
    def generate_certificate(self, nome, texto_certificado, data_local):
        """Gera um certificado PDF em memória"""
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=landscape(A4))
        
        # Se houver imagem de fundo
        if self.background_image:
            try:
                img = ImageReader(self.background_image)
                c.drawImage(img, 0, 0, width=self.width, height=self.height)
            except:
                pass
        
        # Configurar fonte e desenhar texto
        # Título
        c.setFont("Helvetica-Bold", 36)
        c.drawCentredString(self.width/2, self.height - 100, "CERTIFICADO")
        
        # Subtítulo
        c.setFont("Helvetica", 24)
        c.drawCentredString(self.width/2, self.height - 150, "Certificamos que")
        
        # Nome do participante
        c.setFont("Helvetica-Bold", 30)
        c.drawCentredString(self.width/2, self.height - 200, nome)
        
        # Texto do certificado
        c.setFont("Helvetica", 18)
        # Quebrar texto longo em múltiplas linhas
        lines = self._wrap_text(texto_certificado, 80)
        y_position = self.height - 250
        for line in lines:
            c.drawCentredString(self.width/2, y_position, line)
            y_position -= 25
        
        # Data e local
        c.setFont("Helvetica", 16)
        c.drawRightString(self.width - 50, 100, data_local)
        
        # Linha para assinatura
        c.line(self.width/2 - 150, 150, self.width/2 + 150, 150)
        c.setFont("Helvetica", 12)
        c.drawCentredString(self.width/2, 135, "Assinatura")
        
        c.save()
        buffer.seek(0)
        return buffer
    
    def _wrap_text(self, text, max_chars):
        """Quebra texto em linhas menores"""
        words = text.split()
        lines = []
        current_line = []
        current_length = 0
        
        for word in words:
            if current_length + len(word) + 1 <= max_chars:
                current_line.append(word)
                current_length += len(word) + 1
            else:
                lines.append(' '.join(current_line))
                current_line = [word]
                current_length = len(word)
        
        if current_line:
            lines.append(' '.join(current_line))
        
        return lines

def processar_formulario_presenca(df):
    """Processa dados de presença do formulário"""
    # Buscar colunas relevantes
    nome_col = None
    email_col = None
    atividade_col = None
    
    for col in df.columns:
        col_lower = col.lower()
        if "nome" in col_lower and "ultimo" in col_lower:
            nome_col = col
        elif "e-mail" in col_lower and "certificado" in col_lower:
            email_col = col
        elif "atividade" in col_lower and "participou" in col_lower:
            atividade_col = col
    
    if not all([nome_col, email_col, atividade_col]):
        # Tentar usar índices como fallback
        if len(df.columns) >= 9:
            nome_col = df.columns[6]
            email_col = df.columns[7]
            atividade_col = df.columns[8]
        else:
            raise ValueError("Estrutura do arquivo não reconhecida")
    
    # Limpar dados
    df_limpo = df.dropna(subset=[nome_col])
    df_limpo['Horas'] = 2  # 2 horas por presença
    
    # Agrupar por pessoa
    resumo = df_limpo.groupby(nome_col).agg({
        'Horas': 'sum',
        email_col: 'first',
        atividade_col: lambda x: list(x)
    }).reset_index()
    
    resumo.columns = ['Nome', 'Horas_Total', 'Email', 'Atividades']
    resumo['Atividades_Texto'] = resumo['Atividades'].apply(
        lambda x: ', '.join(sorted(set(str(a) for a in x if pd.notna(a))))
    )
    
    return resumo

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/gerar_certificados_formulario', methods=['POST'])
def gerar_certificados_formulario():
    try:
        # Receber arquivo
        file = request.files['arquivo']
        if not file:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        # Dados do formulário
        data_inicio = request.form.get('data_inicio', '')
        data_fim = request.form.get('data_fim', '')
        local = request.form.get('local', 'Belo Horizonte')
        data_emissao = request.form.get('data_emissao', datetime.now().strftime('%d/%m/%Y'))
        
        # Processar imagem de fundo se enviada
        background = None
        if 'background' in request.files:
            bg_file = request.files['background']
            if bg_file and bg_file.filename:
                background = Image.open(bg_file)
        
        # Ler Excel
        df = pd.read_excel(file)
        resumo = processar_formulario_presenca(df)
        
        # Gerar certificados
        certificados = []
        generator = CertificateGenerator(background)
        
        for _, participante in resumo.iterrows():
            nome = participante['Nome']
            horas = participante['Horas_Total']
            atividades = participante['Atividades_Texto']
            
            # Texto do certificado
            if len(atividades) > 100:
                texto = f"Participou de atividades de formação ministradas pela Associação Allos, realizadas de {data_inicio} a {data_fim}, totalizando carga horária de {horas} horas."
            else:
                texto = f"Participou das seguintes atividades ministradas pela Associação Allos: {atividades}, realizadas de {data_inicio} a {data_fim}, totalizando carga horária de {horas} horas."
            
            data_local = f"{local}, {data_emissao}"
            
            # Gerar PDF
            pdf_buffer = generator.generate_certificate(nome, texto, data_local)
            
            certificados.append({
                'nome': nome,
                'email': participante['Email'],
                'pdf': pdf_buffer.getvalue()
            })
        
        # Criar arquivo ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for cert in certificados:
                nome_arquivo = f"{cert['nome'].replace(' ', '_')}_certificado.pdf"
                zip_file.writestr(nome_arquivo, cert['pdf'])
        
        zip_buffer.seek(0)
        
        # Retornar estatísticas e download
        stats = {
            'total_participantes': len(resumo),
            'total_horas': resumo['Horas_Total'].sum(),
            'media_horas': resumo['Horas_Total'].mean()
        }
        
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='certificados.zip'
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/gerar_certificados_personalizado', methods=['POST'])
def gerar_certificados_personalizado():
    try:
        # Receber arquivo
        file = request.files['arquivo']
        if not file:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        # Dados do formulário
        texto_padrao = request.form.get('texto_certificado', '')
        data_local = request.form.get('data_local', f"Belo Horizonte, {datetime.now().strftime('%d de %B de %Y')}")
        coluna_nome = request.form.get('coluna_nome', '')
        coluna_texto = request.form.get('coluna_texto', '')
        
        # Processar imagem de fundo
        background = None
        if 'background' in request.files:
            bg_file = request.files['background']
            if bg_file and bg_file.filename:
                background = Image.open(bg_file)
        
        # Ler Excel
        df = pd.read_excel(file)
        
        # Validar colunas
        if coluna_nome and coluna_nome not in df.columns:
            return jsonify({'error': f'Coluna "{coluna_nome}" não encontrada'}), 400
        
        # Se não especificou coluna, usar a primeira
        if not coluna_nome:
            coluna_nome = df.columns[0]
        
        # Gerar certificados
        certificados = []
        generator = CertificateGenerator(background)
        
        for _, row in df.iterrows():
            nome = str(row[coluna_nome])
            
            # Usar texto personalizado da coluna ou texto padrão
            if coluna_texto and coluna_texto in df.columns and pd.notna(row[coluna_texto]):
                texto = str(row[coluna_texto])
            else:
                texto = texto_padrao.replace('{nome}', nome)
            
            # Gerar PDF
            pdf_buffer = generator.generate_certificate(nome, texto, data_local)
            
            certificados.append({
                'nome': nome,
                'pdf': pdf_buffer.getvalue()
            })
        
        # Criar arquivo ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for cert in certificados:
                nome_arquivo = f"{cert['nome'].replace(' ', '_')}_certificado.pdf"
                zip_file.writestr(nome_arquivo, cert['pdf'])
        
        zip_buffer.seek(0)
        
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='certificados_personalizados.zip'
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/preview_excel', methods=['POST'])
def preview_excel():
    """Retorna preview das colunas do Excel"""
    try:
        file = request.files['arquivo']
        if not file:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        df = pd.read_excel(file)
        
        # Retornar informações sobre o arquivo
        info = {
            'colunas': list(df.columns),
            'total_linhas': len(df),
            'preview': df.head(5).to_dict('records')
        }
        
        return jsonify(info)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)