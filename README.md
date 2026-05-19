# Automação MPP → XLSX

Aplicação em Python para localizar arquivos Microsoft Project (.mpp), criar cópias temporárias, abrir os arquivos no Microsoft Project e exportar as tarefas para planilhas Excel (.xlsx) com colunas configuráveis.

## Visão geral

O fluxo principal do projeto é:

1. Buscar arquivos .mpp em uma pasta base.
2. Criar cópias temporárias para processamento.
3. Abrir cada arquivo no Microsoft Project.
4. Ler os campos configurados em `COLUNAS`.
5. Gerar um .xlsx para cada arquivo processado.
6. Remover as cópias temporárias ao final.

## Recursos

- Busca arquivos .mpp em pasta raiz com opção recursiva.
- Cria cópias de trabalho com sufixo configurável.
- Lê tarefas diretamente do Microsoft Project via COM.
- Exporta para Excel com formatação de datas, percentuais, duração e trabalho.
- Permite configurar pastas e colunas por interface gráfica.
- Remove as cópias temporárias após a exportação.

## Requisitos

- Windows
- Python 3.9 ou superior
- Microsoft Project instalado

## Dependências

Instale os pacotes listados em [requeriments.txt](requeriments.txt).

## Instalação

1. Crie e ative um ambiente virtual:

```bash
python -m venv .venv
.venv\Scripts\activate
```

2. Instale as dependências:

```bash
pip install -r requeriments.txt
```

3. Ajuste as pastas e colunas em `config.py` ou pela interface gráfica.

## Configuração

As principais configurações ficam em `config.py`:

- `PASTA_RAIZ`: pasta onde os arquivos .mpp serão procurados.
- `PASTA_COPIA`: pasta onde as cópias temporárias serão criadas.
- `SUFIXO_COPIA`: sufixo adicionado ao nome da cópia.
- `PASTA_XLSX`: pasta onde os arquivos exportados serão salvos.
- `BUSCA_RECURSIVA`: define se a busca inclui subpastas.
- `PROJECT_VISIVEL`: define se o Microsoft Project abre visível durante o processamento.
- `COLUNAS`: lista de pares `(nome_no_xlsx, atributo_no_project)`.

Exemplo:

```python
PASTA_RAIZ = r"C:\Projetos\MPP"
PASTA_COPIA = r"C:\Projetos\MPP\TEMP"
PASTA_XLSX = r"C:\Projetos\MPP\Exportados"
BUSCA_RECURSIVA = True
PROJECT_VISIVEL = False

COLUNAS = [
   ("Nome", "Name"),
   ("Início", "Start"),
   ("Término", "Finish"),
   ("Duração", "Duration"),
]
```

## Como usar

### Execução principal

```bash
python main.py
```

Esse comando executa todo o fluxo: busca, cópia, leitura no Project, exportação e limpeza das cópias.

### Interface de configuração

```bash
python config_ui.py
```

A interface permite:

- alterar caminhos de origem, cópia e exportação;
- ligar ou desligar a busca recursiva;
- escolher se o Microsoft Project ficará visível;
- adicionar, editar e remover colunas exportadas.

## Estrutura do projeto

```text
microsoft-project-automation/
├── main.py
├── config.py
├── config_ui.py
├── buscar_arquivos.py
├── copiar_arquivos.py
├── exportar_xlsx.py
├── requeriments.txt
├── LICENSE
└── README.md
```

## Formato da exportação

Os arquivos gerados usam a aba `Tarefas` e aplicam tratamento automático para alguns tipos de campo:

- datas em `DD/MM/YYYY`;
- percentuais em `0.00%`;
- trabalho em `0.00h`;
- duração em `0.00d`.

As colunas recebem largura ajustada automaticamente, com destaque para campos de nome.

## Licença

O uso deste projeto é restrito e depende de autorização prévia e expressa. Consulte o arquivo [LICENSE](LICENSE) para os termos completos.

