# ⚖️ Gerador de Ata Jurídica — Família Pilatti

App Streamlit para geração automática de atas de reunião jurídica mensais.

## Como publicar (uma vez só)

### 1. Crie uma conta no GitHub
Acesse github.com → crie uma conta gratuita → crie um repositório novo chamado `ata-juridica`

### 2. Suba os arquivos
No repositório criado, faça upload dos 3 arquivos:
- `app.py`
- `requirements.txt`
- `README.md`

### 3. Publique no Streamlit Cloud
- Acesse share.streamlit.io
- Faça login com sua conta Google ou GitHub
- Clique em "New app"
- Selecione seu repositório `ata-juridica`
- Main file path: `app.py`
- Clique em "Deploy"

Em ~2 minutos seu app estará disponível em um link permanente como:
`https://ata-juridica-pilatti.streamlit.app`

## Uso mensal
1. Acesse o link do app
2. Preencha data e participantes
3. Faça upload do Excel exportado do sistema
4. Clique em "Gerar Ata"
5. Baixe o Word → revise encerrados e deliberações → envie ao cliente

## Atualizar a chave Gemini
Edite a linha no arquivo `app.py`:
```python
GEMINI_API_KEY = "sua_nova_chave_aqui"
```
Salve o arquivo no GitHub → o Streamlit atualiza automaticamente em segundos.
