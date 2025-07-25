# Gerador de Certificados - Associação Allos

Sistema web para geração automática de certificados em PDF a partir de arquivos Excel.

## 🚀 Funcionalidades

### 1. Certificados por Presença
- Conta 2 horas por cada participação registrada
- Agrupa automaticamente por participante
- Gera certificados com total de horas acumuladas

### 2. Certificados Personalizados
- Texto customizável para cada certificado
- Possibilidade de usar diferentes textos por linha
- Data e local de emissão configuráveis

## 📋 Pré-requisitos

- Python 3.8 ou superior
- Navegador web moderno (Chrome, Firefox, Safari, Edge)

## 🔧 Instalação

1. Clone ou baixe os arquivos do projeto:
   - `app.py`
   - `templates/index.html`
   - `requirements.txt`

2. Instale as dependências:
```bash
pip install -r requirements.txt
```

## 🎯 Como usar

1. Execute a aplicação:
```bash
python app.py
```

2. Acesse no navegador:
```
http://localhost:5000
```

3. Escolha o tipo de certificado:
   - **Por Presença**: Para eventos com múltiplas participações
   - **Personalizado**: Para certificados com texto específico

## 📊 Formato do arquivo Excel

### Para Certificados por Presença
O sistema procura automaticamente por colunas com palavras-chave como:
- Nome: "nome", "participante", "name"
- Email: "e-mail", "email", "mail"
- Atividade: "atividade", "evento", "workshop"

### Para Certificados Personalizados
- Primeira coluna: Nomes dos participantes (ou selecione outra coluna)
- Opcionalmente: Coluna com texto personalizado para cada pessoa

## 🎨 Personalização

### Imagem de Fundo
- Aceita formatos: JPG, PNG
- Recomendado: Imagem em paisagem (297 x 210 mm)
- A imagem será ajustada automaticamente

### Texto do Certificado
Use a variável `{nome}` onde desejar inserir o nome do participante.

Exemplo:
```
{nome} participou do workshop de Python realizado em julho de 2025.
```

## 📁 Estrutura do Projeto

```
projeto/
│
├── app.py              # Aplicação Flask principal
├── requirements.txt    # Dependências Python
└── templates/
    └── index.html      # Interface web
```

## 🔍 Solução de Problemas

### Erro ao ler arquivo Excel
- Verifique se o arquivo está no formato .xlsx ou .xls
- Certifique-se de que o arquivo não está corrompido
- O arquivo não deve estar aberto em outro programa

### Certificados não são gerados
- Verifique se há nomes válidos no arquivo
- Confirme que as colunas estão corretamente identificadas
- Veja o console do terminal para mensagens de erro

### Problemas com acentuação
- O sistema suporta caracteres UTF-8
- Se houver problemas, salve o Excel com codificação UTF-8

## 💡 Dicas

1. **Preview do Excel**: Na aba de certificados personalizados, você pode visualizar as primeiras linhas do arquivo antes de gerar.

2. **Múltiplas Execuções**: Você pode gerar certificados várias vezes sem reiniciar a aplicação.

3. **Download Automático**: Os certificados são baixados automaticamente em um arquivo ZIP.

## 🛠️ Desenvolvimento

Para modo de desenvolvimento com auto-reload:
```python
app.run(debug=True, port=5000)
```

Para produção, considere usar um servidor WSGI como Gunicorn:
```bash
pip install gunicorn
gunicorn app:app
```