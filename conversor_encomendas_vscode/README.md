# Conversor de notas de encomenda — VS Code

Aplicação local para extrair linhas de artigos de PDFs de notas de encomenda, exportar Excel e criar etiquetas em PDF A4 com QR Codes verdadeiros.

## Privacidade

A aplicação não faz chamadas a APIs externas e não envia PDFs para sites. Os ficheiros são processados pelo Python no computador onde a aplicação está instalada. Ao instalar dependências, o `pip` precisa de descarregar os pacotes do repositório configurado, mas os PDFs não são enviados com esses pedidos.

## 1. Instalar Python e VS Code

- Instala Python 3.11 ou 3.12.
- Instala Visual Studio Code.
- Abre esta pasta no VS Code.

## 2. Criar ambiente virtual

No terminal integrado do VS Code:

### Windows (PowerShell)

```powershell
py -m venv .venv
.venv\\Scripts\\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Se o PowerShell bloquear a ativação, usa o terminal Command Prompt e executa:

```bat
.venv\\Scripts\\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## 3. Executar

```bash
streamlit run app.py
```

O navegador abre a aplicação local. Se não abrir automaticamente, o terminal indica o endereço local (normalmente `http://localhost:8501`).

## 4. Utilizar

1. Carrega um ou mais PDFs.
2. Revê a tabela extraída e corrige os campos necessários.
3. Descarrega o Excel consolidado.
4. Descarrega o PDF A4 com as etiquetas.
5. Imprime o PDF a 100% / tamanho real, sem “ajustar à página”, e testa uma folha antes de imprimir em quantidade.

## QR Code

Cada QR Code contém os campos exportados do respetivo artigo, em JSON UTF-8. O QR Code não é uma imagem decorativa: é gerado pela biblioteca `qrcode` e pode ser lido por um leitor compatível.

## Limitações da primeira versão

- Está ajustada ao layout do PDF de exemplo (nota de encomenda com texto selecionável).
- PDFs digitalizados necessitam de OCR local, ainda não incluído.
- Como descrições e tabelas variam entre sistemas, confirma sempre os dados antes de imprimir.
- A etiqueta mostra versões compactas de alguns textos para caber no papel; o QR Code mantém os valores completos.
- O nome do cliente é extraído do bloco “Exmo(s) Sr(s)”. Confirma que este corresponde ao cliente pretendido na tua utilização.

## Estrutura

```text
conversor_encomendas_vscode/
├── app.py
├── requirements.txt
├── README.md
└── output/
```
