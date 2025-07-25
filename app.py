from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from fpdf import FPDF
import os
import zipfile
from datetime import datetime
import io
import tempfile
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

class PDFCustom(FPDF):
    def __init__(self, bg_image_path=None):
        super().__init__('L', 'mm', 'A4')  # Paisagem
        self.bg_image_path = bg_image_path
        self.set_auto_page_break(auto=False, margin=0)

    def header(self):
        if self.bg_image_path and os.path.exists(self.bg_image_path):
            self.image(self.bg_image_path, x=0, y=0, w=self.w, h=self.h)

def processar_participacoes(df):
    """Processa as participações e conta 2 horas por presença"""
    
    # Tenta encontrar as colunas corretas de forma mais flexível
    nome_coluna = None
    email_coluna = None
    atividade_coluna = None
    
    # Procura pelas colunas baseado em palavras-chave
    for col in df.columns:
        col_lower = col.lower()
        
        # Busca a coluna de nome
        if any(termo in col_lower for termo in ["nome completo", "name", "participante", "ultimo nome"]):
            if not nome_coluna:  # Pega a primeira coluna que corresponde
                nome_coluna = col
        
        # Busca a coluna de email
        elif any(termo in col_lower for termo in ["e-mail", "email", "mail"]):
            if not email_coluna:
                email_coluna = col
        
        # Busca a coluna de atividade
        elif any(termo in col_lower for termo in ["atividade", "evento", "workshop", "curso"]):
            if not atividade_coluna:
                atividade_coluna = col
    
    # Se não encontrou, usa as primeiras colunas
    if not nome_coluna and len(df.columns) > 0:
        nome_coluna = df.columns[0]
    
    if not nome_coluna:
        raise ValueError("Não foi possível identificar a coluna de nomes")
    
    # Remove linhas com valores vazios na coluna de nome
    df_limpo = df.dropna(subset=[nome_coluna])
    
    # Adiciona coluna de horas (2 horas por presença)
    df_limpo['Horas'] = 2
    
    # Agrupa por nome e soma as horas
    resumo = df_limpo.groupby(nome_coluna).agg({
        'Horas': 'sum'
    }).reset_index()
    
    # Adiciona email e atividades se as colunas existirem
    if email_coluna and email_coluna in df.columns:
        emails = df_limpo.groupby(nome_coluna)[email_coluna].first()
        resumo['Email'] = resumo[nome_coluna].map(emails)
    
    if atividade_coluna and atividade_coluna in df.columns:
        atividades = df_limpo.groupby(nome_coluna)[atividade_coluna].apply(
            lambda x: ', '.join(sorted(set(str(a) for a in x if pd.notna(a))))
        )
        resumo['Atividades'] = resumo[nome_coluna].map(atividades)
    
    # Renomeia colunas
    resumo.rename(columns={nome_coluna: 'Nome', 'Horas': 'Horas_Total'}, inplace=True)
    
    return resumo

def gerar_certificado_presenca(nome, horas, data_inicio, data_fim, local, data_emissao, bg_path=None):
    """Gera certificado para participação com horas acumuladas"""
    
    pdf = PDFCustom(bg_image_path=bg_path)
    pdf.set_margins(30, 25, 30)
    pdf.add_page()
    
    # Título
    pdf.set_y(50)
    pdf.set_font("Times", "", 28)
    pdf.cell(0, 15, "CERTIFICADO", align="C", ln=True)
    
    # Certificamos que
    pdf.ln(10)
    pdf.set_font("Times", "", 18)
    pdf.cell(0, 10, "Certificamos que", align="C", ln=True)
    
    # Nome
    pdf.ln(5)
    pdf.set_font("Times", "B", 26)
    pdf.cell(0, 15, nome, align="C", ln=True)
    
    # Texto do certificado
    pdf.ln(10)
    pdf.set_font("Times", "", 16)
    
    periodo = f"no período de {data_inicio} a {data_fim}"
    texto = f"Participou das atividades realizadas pela Associação Allos {periodo}, "
    texto += f"com carga horária total de {int(horas)} horas."
    
    pdf.multi_cell(0, 10, texto, align="C")
    
    # Local e data
    pdf.ln(20)
    pdf.set_font("Times", "", 14)
    pdf.cell(0, 10, f"{local}, {data_emissao}", align="R")
    
    return pdf

def gerar_certificado_personalizado(nome, texto, data_local, bg_path=None):
    """Gera certificado com texto personalizado"""
    
    pdf = PDFCustom(bg_image_path=bg_path)
    pdf.set_margins(30, 25, 30)
    pdf.add_page()
    
    # Título
    pdf.set_y(50)
    pdf.set_font("Times", "", 28)
    pdf.cell(0, 15, "CERTIFICADO", align="C", ln=True)
    
    # Certificamos que
    pdf.ln(10)
    pdf.set_font("Times", "", 18)
    pdf.cell(0, 10, "Certificamos que", align="C", ln=True)
    
    # Nome
    pdf.ln(5)
    pdf.set_font("Times", "B", 26)
    pdf.cell(0, 15, nome, align="C", ln=True)
    
    # Texto personalizado
    pdf.ln(10)
    pdf.set_font("Times", "", 16)
    
    # Substitui {nome} no texto
    texto_final = texto.replace("{nome}", nome)
    pdf.multi_cell(0, 10, texto_final, align="C")
    
    # Data e local
    if data_local:
        pdf.ln(20)
        pdf.set_font("Times", "", 14)
        pdf.cell(0, 10, data_local, align="R")
    
    return pdf

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/preview_excel', methods=['POST'])
def preview_excel():
    """Preview das primeiras linhas do Excel"""
    try:
        file = request.files['arquivo']
        df = pd.read_excel(file)
        
        # Pega as primeiras 5 linhas
        preview_data = df.head(5).fillna('').to_dict('records')
        
        return jsonify({
            'colunas': df.columns.tolist(),
            'preview': preview_data,
            'total_linhas': len(df)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/gerar_certificados_formulario', methods=['POST'])
def gerar_certificados_formulario():
    """Gera certificados baseados em presença (2h por participação)"""
    try:
        # Recebe os arquivos e dados
        arquivo = request.files['arquivo']
        background = request.files.get('background')
        data_inicio = request.form['data_inicio']
        data_fim = request.form['data_fim']
        local = request.form.get('local', 'Belo Horizonte')
        data_emissao = request.form.get('data_emissao')
        
        # Formata datas
        data_inicio_fmt = datetime.strptime(data_inicio, '%Y-%m-%d').strftime('%d/%m/%Y')
        data_fim_fmt = datetime.strptime(data_fim, '%Y-%m-%d').strftime('%d/%m/%Y')
        
        if data_emissao:
            data_emissao_fmt = datetime.strptime(data_emissao, '%Y-%m-%d')
        else:
            data_emissao_fmt = datetime.now()
        
        # Formata data de emissão em português
        meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
                 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
        data_emissao_texto = f"{data_emissao_fmt.day} de {meses[data_emissao_fmt.month-1]} de {data_emissao_fmt.year}"
        
        # Salva background temporariamente se fornecido
        bg_path = None
        if background:
            bg_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(background.filename))
            background.save(bg_path)
        
        # Carrega e processa dados
        df = pd.read_excel(arquivo)
        resumo = processar_participacoes(df)
        
        # Cria arquivo ZIP em memória
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for _, participante in resumo.iterrows():
                nome = participante['Nome']
                horas = participante['Horas_Total']
                
                # Gera o certificado
                pdf = gerar_certificado_presenca(
                    nome, horas, data_inicio_fmt, data_fim_fmt, 
                    local, data_emissao_texto, bg_path
                )
                
                # Salva em memória
                pdf_buffer = io.BytesIO()
                pdf_content = pdf.output()
                pdf_buffer.write(pdf_content)
                
                # Adiciona ao ZIP
                nome_arquivo = f"{nome.replace(' ', '_')}_certificado.pdf"
                zip_file.writestr(nome_arquivo, pdf_buffer.getvalue())
        
        # Limpa arquivo temporário
        if bg_path and os.path.exists(bg_path):
            os.remove(bg_path)
        
        # Retorna o ZIP
        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='certificados_presenca.zip'
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/gerar_certificados_personalizado', methods=['POST'])
def gerar_certificados_personalizado():
    """Gera certificados com texto personalizado"""
    try:
        # Recebe os arquivos e dados
        arquivo = request.files['arquivo']
        background = request.files.get('background')
        coluna_nome = request.form.get('coluna_nome', '')
        texto_certificado = request.form['texto_certificado']
        coluna_texto = request.form.get('coluna_texto', '')
        data_local = request.form.get('data_local', '')
        
        # Salva background temporariamente se fornecido
        bg_path = None
        if background:
            bg_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(background.filename))
            background.save(bg_path)
        
        # Carrega dados
        df = pd.read_excel(arquivo)
        
        # Determina a coluna de nomes
        if coluna_nome:
            nome_col = coluna_nome
        else:
            nome_col = df.columns[0]
        
        # Remove linhas vazias
        df_limpo = df.dropna(subset=[nome_col])
        
        # Cria arquivo ZIP em memória
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for _, row in df_limpo.iterrows():
                nome = str(row[nome_col]).strip()
                
                # Determina o texto do certificado
                if coluna_texto and coluna_texto in df.columns:
                    texto = str(row[coluna_texto])
                else:
                    texto = texto_certificado
                
                # Gera o certificado
                pdf = gerar_certificado_personalizado(nome, texto, data_local, bg_path)
                
                # Salva em memória
                pdf_buffer = io.BytesIO()
                pdf_content = pdf.output()
                pdf_buffer.write(pdf_content)
                
                # Adiciona ao ZIP
                nome_arquivo = f"{nome.replace(' ', '_')}_certificado.pdf"
                zip_file.writestr(nome_arquivo, pdf_buffer.getvalue())
        
        # Limpa arquivo temporário
        if bg_path and os.path.exists(bg_path):
            os.remove(bg_path)
        
        # Retorna o ZIP
        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='certificados_personalizados.zip'
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

    