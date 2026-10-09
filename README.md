# Negócios Dia - BTC

Script em Python que lê o CSV diário de negócios, aplica filtros, gera uma
planilha Excel formatada com tabela dinâmica e resumos (Top 10 Execuções,
Top 10 Papéis), e monta um e-mail de resumo diário via Outlook.

> ⚠️ Requer Windows com Microsoft Outlook instalado, pois usa `pywin32`
> para automatizar o Outlook.

## Pré-requisitos

- Python 3.9 ou superior
- Windows com Outlook configurado (para o envio de e-mail)
- Acesso de leitura ao arquivo CSV de entrada e de escrita na pasta de saída

## Instalação

1. Clone o repositório:
   ```
   git clone https://github.com/LeonardoLubas/Negocios_dia_btc.git
   cd Negocios_dia_btc
   ```

2. (Opcional, recomendado) Crie e ative um ambiente virtual:
   ```
   python -m venv venv
   venv\Scripts\activate
   ```

3. Instale as dependências:
   ```
   pip install -r requirements.txt
   ```

## Configuração

Este projeto não versiona dados sensíveis (caminhos de rede, e-mails,
códigos internos). Você precisa criar dois arquivos locais, que **não**
devem ser commitados (já estão no `.gitignore`):

### 1. Arquivo `.env`

Copie o modelo e preencha com os valores reais:

```
copy .env.example .env
```

Edite o `.env` com um editor de texto e preencha:

| Variável | Descrição |
|---|---|
| `ARQUIVO_ENTRADA` | Caminho completo do CSV de negócios de entrada |
| `PASTA_SAIDA` | Pasta onde o Excel e o CSV filtrado serão salvos |
| `EMAIL_DESTINATARIOS` | E-mails principais, separados por `;` |
| `EMAIL_COPIA` | E-mails em cópia (CC), separados por `;` |
| `TOMADORES_FILE` | Caminho do arquivo JSON com o mapa de tomadores (padrão: `tomadores.json`) |
| `EXEC_DOADOR_ID` | Código do executor doador usado no filtro |
| `LIMITE_CONTA_DOADOR` | Limite numérico usado no filtro da conta doadora |

### 2. Arquivo `tomadores.json`

Copie o modelo e preencha com o mapa real de códigos:

```
copy tomadores.example.json tomadores.json
```

O formato é `"código": "Nome da Corretora"`:

```json
{
    "3": "XP",
    "8": "UBS"
}
```

## Uso

Com o ambiente configurado, rode:

```
python negocios_dia.py
```

O script vai:

1. Ler e filtrar o CSV de entrada
2. Salvar `NegociosFiltrado.csv` na pasta de saída
3. Gerar `NegociosDia.xlsx` com abas Resumo, Tabela Dinâmica e Dados Filtrados (formatado)
4. Abrir um e-mail no Outlook (`mail.Display()`) pronto para conferência, com o resumo e o Excel em anexo

> O envio automático (`mail.Send()`) está comentado no código por padrão —
> o e-mail abre para você conferir antes de enviar manualmente.

## Estrutura do projeto

```
negocios_dia.py          # script principal
requirements.txt         # dependências Python
.env.example              # modelo de variáveis de ambiente (sem dados reais)
.env                       # suas variáveis reais (NÃO versionado)
tomadores.example.json    # modelo do mapa de tomadores (sem dados reais)
tomadores.json             # mapa real de tomadores (NÃO versionado)
.gitignore
```
