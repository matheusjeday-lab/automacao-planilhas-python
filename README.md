# Automação de Planilhas em Python

**Status:** ✅ Projeto concluído e testado

![Python](https://img.shields.io/badge/Python-3.x-blue)
![Pandas](https://img.shields.io/badge/Pandas-data%20processing-blue)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel-green)
![Matplotlib](https://img.shields.io/badge/Matplotlib-charts-orange)

Sistema desenvolvido em Python para automatizar o processamento de planilhas de vendas.

O programa lê automaticamente os arquivos recebidos, valida os dados, identifica erros e duplicidades, consolida as informações em um histórico e gera relatórios atualizados.

## Objetivo

O projeto simula uma necessidade comum em empresas que recebem planilhas de vendas diariamente.

Em vez de reunir e conferir os arquivos manualmente, o sistema automatiza o processo:

```text
Planilhas recebidas
        ↓
Leitura dos arquivos
        ↓
Validação dos dados
        ↓
Identificação de erros e duplicidades
        ↓
Atualização do histórico
        ↓
Geração dos relatórios
        ↓
Registro da execução
        ↓
Arquivos processados
```

## Funcionalidades

* Leitura automática de arquivos `.xlsx`
* Descoberta automática de novas planilhas na pasta de entrada
* Validação das colunas obrigatórias
* Identificação de campos vazios
* Validação de datas
* Validação de quantidade e valores
* Identificação e remoção de registros duplicados
* Consolidação de diferentes planilhas
* Manutenção de histórico das vendas
* Relatório geral de vendas
* Resumo de vendas por mês
* Resumo por produto
* Resumo por cliente
* Identificação do produto mais comprado por cada cliente
* Dashboard com indicadores e gráficos
* Relatório separado de erros
* Controle dos arquivos processados
* Registro das execuções em log
* Movimentação automática dos arquivos processados para uma pasta específica

## Estrutura do projeto

```text
AutomacaoPlanilhas/
│
├── entrada/
│   └── Planilhas novas para processamento
│
├── saida/
│   ├── relatorio_vendas.xlsx
│   ├── dados_historicos.csv
│   ├── log_execucoes.csv
│   ├── controle_arquivos.csv
│   └── relatorio_erros.xlsx
│
├── processados/
│   └── Arquivos que já foram processados
│
├── programa/
│   └── automacao.py
│
└── README.md
```
## Demonstração

O sistema foi testado com diferentes cenários de processamento, incluindo:

* Processamento de múltiplos arquivos
* Inclusão de novas vendas
* Registros duplicados
* Arquivos com colunas ausentes
* Datas inválidas
* Quantidades inválidas
* Valores inválidos
* Arquivos vazios
* Arquivos válidos contendo registros inválidos

Durante os testes, o sistema conseguiu continuar processando os arquivos válidos mesmo quando encontrou arquivos ou registros com problemas.

### Exemplo de fluxo

```text
entrada/
   ↓
Leitura das planilhas
   ↓
Validação
   ↓
Erros e duplicidades identificados
   ↓
Histórico atualizado
   ↓
Relatórios atualizados
   ↓
Controle e log registrados
   ↓
processados/
```

### Resultado

Ao final de cada execução, o terminal informa:

```text
Arquivos encontrados
Arquivos processados
Arquivos rejeitados
Linhas lidas
Linhas novas
Duplicidades ignoradas
Linhas consolidadas
Total das vendas
```

Os relatórios gerados permitem acompanhar tanto os resultados das vendas quanto os problemas encontrados durante o processamento.

## Principais funcionalidades

* Leitura automática de arquivos Excel (`.xlsx`)
* Consolidação de dados de múltiplas planilhas
* Validação de colunas e dados
* Identificação de campos vazios
* Validação de datas, quantidades e valores
* Identificação e tratamento de registros duplicados
* Continuidade do processamento mesmo quando existem arquivos ou registros inválidos
* Atualização automática do histórico de vendas
* Geração de relatórios em Excel
* Dashboard com indicadores e gráficos
* Resumo de vendas por mês, produto e cliente
* Identificação do produto mais comprado por cliente
* Relatório detalhado de erros encontrados
* Controle individual dos arquivos processados
* Registro das execuções em arquivo de log (`CSV`)
* Organização automática dos arquivos já processados

## Relatórios gerados

### Dashboard

Apresenta uma visão geral dos dados, incluindo:

* Total de vendas
* Número de vendas
* Quantidade de clientes
* Quantidade de produtos
* Gráficos de vendas

### Vendas por mês

Mostra o total vendido em cada mês.

### Vendas por produto

Apresenta o desempenho de cada produto.

### Vendas por cliente

Mostra quanto cada cliente comprou.

### Resumo por cliente

Além do valor total comprado, apresenta o produto mais comprado e a quantidade correspondente.

### Relatório de erros

Registra problemas encontrados durante o processamento, permitindo identificar quais arquivos ou registros precisam ser corrigidos.

### Controle de arquivos

Mantém um histórico dos arquivos processados, permitindo saber quais planilhas já foram recebidas e processadas.

## Formato esperado das planilhas

As planilhas de entrada devem possuir as seguintes colunas:

```text
Data
Cliente
Produto
Quantidade
Valor
```

Exemplo:

| Data       | Cliente | Produto | Quantidade | Valor |
| ---------- | ------- | ------- | ---------: | ----: |
| 02/01/2026 | Ana     | Teclado |          2 |   150 |
| 03/01/2026 | João    | Mouse   |          3 |    80 |
| 05/01/2026 | Maria   | Teclado |          1 |   150 |

O sistema também remove espaços extras dos nomes das colunas para evitar erros causados por pequenas diferenças de formatação.

## Tecnologias utilizadas

* Python
* Pandas
* OpenPyXL
* Matplotlib
* Excel `.xlsx`
* CSV

## Como executar

### 1. Instalar as dependências

No terminal, dentro da pasta do projeto, execute:

```bash
pip install -r requirements.txt
```

### 2. Adicionar as planilhas

Coloque os arquivos `.xlsx` que deseja processar na pasta:

```text
entrada/
```

As planilhas devem conter as colunas:

```text
Data | Cliente | Produto | Quantidade | Valor
```

### 3. Executar a automação

Execute:

```bash
python programa/automacao.py
```

### 4. Resultados

Após a execução:

* os relatórios serão gerados ou atualizados na pasta `saida/`;
* o histórico de vendas será atualizado;
* duplicidades serão identificadas e ignoradas;
* erros encontrados serão registrados;
* a execução será registrada no log;
* os arquivos processados serão movidos para `processados/`.

O terminal também apresenta um resumo da execução, incluindo arquivos processados, arquivos rejeitados, linhas novas, duplicidades ignoradas e total das vendas.


## Tratamento de erros

O sistema não interrompe todo o processamento quando encontra um arquivo inválido.

Arquivos com problemas são identificados e registrados no relatório de erros, enquanto os arquivos válidos continuam sendo processados.

Entre os problemas identificados estão:

* Arquivo inválido ou impossível de ler
* Colunas obrigatórias ausentes
* Planilha vazia
* Datas inválidas
* Quantidades inválidas
* Valores inválidos
* Campos vazios
* Registros duplicados

## Histórico e duplicidades

O sistema mantém um arquivo histórico com os dados já processados.

Quando uma nova planilha é recebida, seus registros são comparados com o histórico para evitar que uma mesma venda seja adicionada novamente.

Dessa forma, o processamento pode ser executado várias vezes sem duplicar os dados já existentes.

## Exemplo de resultado

Após processar diferentes planilhas, o sistema pode gerar uma estrutura consolidada como:

```text
Total das vendas: R$ 13.840,00
Linhas consolidadas: 29
Arquivos processados: 1
Arquivos rejeitados: 0
Duplicidades ignoradas: 0
```

Os valores acima são apenas um exemplo de execução durante os testes do projeto.

## Objetivo profissional

Este projeto foi desenvolvido como demonstração prática de automação de tarefas administrativas utilizando Python, com foco em:

* Manipulação de dados
* Automação de processos
* Validação de informações
* Tratamento de erros
* Geração de relatórios
* Organização de arquivos
* Persistência de histórico
* Visualização de dados

O projeto simula uma solução que poderia ser utilizada para reduzir tarefas manuais no processamento diário de planilhas.
