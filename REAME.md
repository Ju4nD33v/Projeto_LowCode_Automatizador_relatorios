# Gerador de Relatórios Automatizados — Excel para PDF

> Sistema Python que lê planilhas `.xlsx` e exporta relatórios profissionais em PDF com gráficos e tabelas de dados, de forma totalmente automatizada.

---

## Como o Sistema Funciona

O fluxo do sistema é simples e direto:

```
Planilha Excel (.xlsx)
        │
        ▼
  Leitura dos dados
  (pandas + openpyxl)
        │
        ▼
  Geração dos gráficos
  (matplotlib → imagens .png temporárias)
        │
        ▼
  Montagem do PDF
  (fpdf2 — cabeçalho, KPIs, gráficos, tabela, rodapé)
        │
        ▼
  Arquivo PDF final
  (salvo na mesma pasta da planilha)
```

### Etapas Detalhadas

**1. Detecção da Planilha**
Ao rodar `python gerador_relatorios.py` sem argumentos, o sistema varre a pasta atual em busca de arquivos `.xlsx`. Se encontrar apenas um, usa automaticamente. Se encontrar vários, lista e pede para você escolher.

**2. Leitura e Processamento**
A biblioteca `pandas` lê todos os dados da primeira aba da planilha (ou de uma aba específica, se indicada). As colunas são detectadas automaticamente — não é necessário nenhuma configuração manual.

**3. Geração dos Gráficos**
O `matplotlib` cria imagens PNG temporárias em memória:
- **Gráfico de barras** — agrupa por `Vendedor` e soma o `Total (R$)`
- **Gráfico de pizza** — distribui o `Total (R$)` por `Região`
- **Gráfico de barras** — agrupa por `Produto` e soma o `Total (R$)`

As imagens ficam em uma pasta temporária `__charts_temp__/` que é **apagada automaticamente** ao final.

**4. Montagem do PDF**
O `fpdf2` constrói o PDF página a página:
- **Página 1** — Cabeçalho, cards de KPI e gráficos analíticos
- **Página 2+** — Tabela com todos os registros da planilha

**5. Saída**
O PDF é salvo com o nome `relatorio_AAAAMMDD_HHMMSS.pdf` na mesma pasta, ou no caminho que você definir com `-o`.

---

## Estrutura do Projeto

```
Projeto_LowCode/
├── requirements.txt           # Dependências Python
├── criar_planilha_exemplo.py  # Gera um Excel fictício para teste
├── gerador_relatorios.py      # Script principal (Excel → PDF)
└── REAME.md                   # Este arquivo
```

---

## Instalação

```bash
pip install -r requirements.txt
```

---

## Como Usar

### Modo mais simples — detecção automática
Coloque qualquer arquivo `.xlsx` na mesma pasta e execute:

```bash
python gerador_relatorios.py
```

O sistema encontra a planilha sozinho e gera o PDF.

---

### Com várias planilhas na pasta
O sistema lista os arquivos encontrados e você escolhe:

```
Multiplas planilhas encontradas. Escolha uma:
  [1] vendas_janeiro.xlsx
  [2] vendas_fevereiro.xlsx
Digite o numero: 1
```

---

### Passando argumentos opcionais

```bash
# Especificando o arquivo manualmente
python gerador_relatorios.py minha_planilha.xlsx

# Definindo onde salvar o PDF
python gerador_relatorios.py minha_planilha.xlsx -o relatorio_mensal.pdf

# Especificando a aba da planilha (padrão: primeira aba)
python gerador_relatorios.py minha_planilha.xlsx --aba "Janeiro"
```

---

### Gerando dados de teste
Se você não tem uma planilha, rode este script para criar uma com 50 registros fictícios de vendas:

```bash
python criar_planilha_exemplo.py
```

---

## Colunas Reconhecidas Automaticamente

O sistema detecta as colunas abaixo para montar os gráficos e KPIs. Planilhas com outras colunas também funcionam — os dados são exibidos na tabela de detalhes.

| Coluna              | Tipo     | Usado em                          |
|---------------------|----------|-----------------------------------|
| `Vendedor`          | Texto    | Gráfico de barras por vendedor    |
| `Região`            | Texto    | Gráfico de pizza por região       |
| `Produto`           | Texto    | Gráfico de barras por produto     |
| `Total (R$)`        | Número   | KPIs e todos os gráficos          |
| `Quantidade`        | Número   | Tabela de dados detalhados        |
| `Meta Atingida`     | Sim/Não  | KPI de metas atingidas            |

> Qualquer coluna não listada aqui ainda aparece normalmente na tabela de dados detalhados do PDF.

---

## O que o PDF Contém

| Seção                | Conteúdo                                              |
|----------------------|-------------------------------------------------------|
| **Cabeçalho**        | Título, nome do arquivo fonte e data/hora de geração  |
| **Resumo Executivo** | 4 cards coloridos com KPIs principais                 |
| **Análise Gráfica**  | Gráfico de barras (vendedor) + pizza (região)         |
| **Vendas por Produto**| Gráfico de barras por produto                        |
| **Dados Detalhados** | Tabela completa com até 40 linhas por página          |
| **Rodapé**           | Numeração de páginas em todas as páginas              |

---

## Dependências

| Biblioteca   | Versão  | Função                           |
|--------------|---------|----------------------------------|
| `pandas`     | 2.2.2   | Leitura e processamento do Excel |
| `openpyxl`   | 3.1.2   | Engine de leitura de `.xlsx`     |
| `fpdf2`      | 2.7.9   | Geração do PDF                   |
| `matplotlib` | 3.9.0   | Geração dos gráficos             |
| `Pillow`     | 10.3.0  | Suporte a imagens no PDF         |
