# Automação Contábil

Pequeno projeto em Python para automatizar o lançamento de transações em uma aplicação web de exemplo usando uma planilha Excel como fonte de dados.

**O que é:**
- Interface gráfica em `gui.py` construída com `customtkinter` que permite selecionar uma planilha `.xlsx` e disparar a automação.
- Módulo `automacao_contabil.py` usa Playwright para abrir o site de demonstração, autenticar e incluir lançamentos a partir da planilha.

**Requisitos:**
- Python 3.9+ (recomendado 3.10/3.11)
- pacotes: `playwright`, `openpyxl`, `customtkinter`

Instale dependências:

```bash
python -m pip install --upgrade pip
pip install playwright openpyxl customtkinter
python -m playwright install
```

**Como usar:**
- Execute a interface gráfica:

```bash
python gui.py
```

- Se preferir executar o módulo de automação diretamente (abre diálogo para selecionar o arquivo):

```bash
python automacao_contabil.py
```

**Formato esperado da planilha:**
- A planilha deve conter uma aba chamada `Lançamentos` com colunas (a partir da linha 2): descrição, valor, data, tipo ("Receita"/"Despesa"), categoria, status.

**Observações:**
- `automacao_contabil.py` usa Playwright em modo não-headless e possui delays configuráveis (`action_delay`) para evitar problemas de sincronização.
- Não empurre arquivos sensíveis (por ex. planilhas com dados reais). Veja `.gitignore`.

**Subir para o GitHub:**
- Use o script `push_to_github.ps1` na raiz do projeto ou os comandos Git abaixo:

```powershell
git init
git add .
git commit -m "Initial commit - Automação Contábil"
git branch -M main
git remote add origin <URL-DO-REPOSITÓRIO>
git push -u origin main
```
