# Extrator PDF para Excel - Ordens de Serviço

Sistema para extrair automaticamente dados de ordens de serviço de arquivos PDF e transferir para planilha Excel.

## O que faz

Este programa automatiza a tarefa chata de copiar dados de PDFs para Excel. Ele:

- Abre uma tela para você escolher vários PDFs de uma vez
- Lê os dados das tabelas dentro dos PDFs 
- Coloca tudo organizadinho na planilha Excel
- Mantém a formatação da planilha original
- Calcula algumas fórmulas automaticamente
- Mostra quantas ordens foram processadas e o tempo que levou

## O que você precisa instalar antes

### 1. Python (obrigatório)
Baixe e instale o Python pelo site oficial:
**https://www.python.org/downloads/**

Na hora de instalar, MARQUE a opção "Add Python to PATH" (muito importante!)

**Para testar se instalou certo:**
Abra o Prompt de Comando (cmd) e digite:
```
python --version
```
Se aparecer "Python 3.11.0"!

### 2. Bibliotecas Python
Depois de instalar o Python, abra o **Prompt de Comando** (cmd) e digite este comando:

```
pip install pdfplumber openpyxl pyperclip
```

**Para testar se instalou certo:**
```
pip list
```
Deve aparecer na lista: pdfplumber, openpyxl, pyperclip

### Links diretos caso precise:
- Python: https://www.python.org/downloads/
- Documentação pdfplumber: https://pypi.org/project/pdfplumber/
- Documentação openpyxl: https://pypi.org/project/openpyxl/

## IMPORTANTE - Configure antes de usar

**ATENÇÃO:** Você precisa alterar os caminhos no arquivo `projeto.py` antes de usar:

### Caminhos da planilha Excel
Procure estas 3 linhas no código e altere para o caminho da SUA planilha:

**Linha para carregar a planilha:**
```python
wb = load_workbook(r"C:\Users\joao.beserra\Documents\sala-tecnica.xlsx")
```

**Linha para salvar a planilha:**
```python
wb.save(r"C:\Users\joao.beserra\Documents\sala-tecnica.xlsx")
```

**Linha para abrir a planilha:**
```python
os.startfile(r"C:\Users\joao.beserra\Documents\sala-tecnica.xlsx")
```

**Exemplo de como alterar:**
```python
# Troque por:
wb = load_workbook(r"C:\MinhaPasta\minha-planilha.xlsx")
wb.save(r"C:\MinhaPasta\minha-planilha.xlsx")
os.startfile(r"C:\MinhaPasta\minha-planilha.xlsx")
```

### Nome da aba da planilha
Se sua planilha não tem uma aba chamada "Planilha1", altere esta linha:
```python
ws = wb["Planilha1"]
```
Para o nome da sua aba:
```python
ws = wb["MinhaAba"]
```

## Como usar (passo a passo)

### Jeito mais fácil - usando o .bat:
1. Baixe a pasta pai deste repositório, botão direito e clique em "extrair tudo".
2. Clique com o botão direito no arquivo "PDF_TO_EXCEL" e em "Enviar para"/"Área de tabalho (Criar Atalho).
3. Clique duas vezes no arquivo `PDF_TO_EXCEL.bat`
4. Vai abrir uma janela preta e depois uma tela para escolher os PDFs
5. Selecione os arquivos PDF e clique "Abrir"
6. Aguarde o processamento (vai aparecer na tela o progresso)
7. A planilha vai abrir automaticamente quando terminar

### Jeito manual - rodando pelo Python:
1. Abra o Prompt de Comando (cmd)
2. Navegue até a pasta do projeto:
   ```
   cd C:\caminho\para\sua\pasta
   ```
3. Execute o comando:
   ```
   python projeto.py
   ```

## O que o programa busca nos PDFs

O sistema procura estes campos específicos nas tabelas dos PDFs:
- **Data** - Data da ordem de serviço
- **Data/Prazo** - Prazo de execução  
- **Rodovia** - Identificação da rodovia
- **Trecho** - Trecho da rodovia
- **Sentido** - Sentido da rodovia
- **Classificação** - Classificação do serviço
- **Tipo** - Tipo de serviço
- **Descrição** - Descrição detalhada
- **Origem** - Órgão de origem
- **Executor** - Responsável pela execução
- **Ordem_Serviço** - Número da ordem (extraído da primeira célula da tabela)

## Como os dados são organizados na planilha

O programa organiza os dados nesta ordem específica:
```
Coluna 1: Data
Coluna 2: Origem (ARTESP ou ENGENHARIA)
Coluna 3: Campo especial (RETRO se detectado)
Coluna 4: Data (repetida)
Coluna 5: Data (repetida)
Coluna 6: Data/Prazo
Coluna 7-11: Campos vazios (para fórmulas)
Coluna 12: Rodovia (espaços viram hífens)
Coluna 13: Primeira parte do Trecho (dividido por <>)
Coluna 14: Segunda parte do Trecho (dividido por +)
Coluna 15: Sentido
Coluna 16: Classificação (após "E: ")
Coluna 17: Tipo
Coluna 18: Descrição (segunda linha)
Coluna 19: Campo vazio
Coluna 20: Ordem_Serviço (só número antes do -)
Coluna 21: Executor
```

## Coisas legais que o programa faz sozinho

**Fórmulas automáticas nas colunas:**
- **Coluna 8:** Fórmula para mês por extenso em maiúscula
- **Coluna 9:** Calcula prazo do edital (Coluna F - Coluna D)
- **Coluna 10:** Calcula prazo de execução (Coluna F - Coluna G)
- **Coluna 19:** Sempre coloca "X"

**Formatação inteligente de dados:**
- **Origem:** Se contém "artesp" → vira "ARTESP", senão vira "ENGENHARIA"
- **Campo especial:** Se contém "retro" → coloca "RETRO" na coluna 3
- **Rodovia:** Substitui espaços por hífens
- **Trecho:** Divide por "<>" e depois por "+" em colunas separadas
- **Classificação:** Pega só o que vem depois de "E: "
- **Descrição:** Pega a segunda linha (depois do \n)
- **Ordem de Serviço:** Pega só o número antes do "-"
- **Textos em maiúscula:** Sentido, Classificação, Tipo, Executor

**Conversões especiais:**
- Colunas 13, 14 e 20 são convertidas para números inteiros
- Mantém todas as formatações da planilha original (cores, bordas, fontes)
- Copia fórmulas existentes das linhas anteriores

## Cuidados importantes

**ANTES DE USAR:**
- Feche a planilha Excel (senão dá erro de permissão)
- Faça uma cópia de backup da planilha
- Teste com 1 ou 2 PDFs primeiro para validar
- Certifique-se que a planilha tem pelo menos uma linha com dados (para copiar formatação)

**Formato esperado dos PDFs:**
- Os PDFs precisam ter tabelas organizadas
- As informações devem estar no formato: `Campo\nValor` (separadas por quebra de linha)
- Primeira célula da tabela deve ter o número da ordem de serviço
- Campos obrigatórios: Data, Data/Prazo, Rodovia, Trecho, Sentido, Classificação, Tipo, Descrição, Origem, Executor

**Estrutura das células nos PDFs:**
```
Exemplo de célula esperada:
Data
11/06/2025

Rodovia  
SP-280

Trecho
KM 10<>KM 20+Sentido Norte
```

## Se algo der errado

**"Erro de permissão ao salvar":**
- Feche a planilha Excel completamente
- Veja se você tem permissão para escrever na pasta
- Tente executar como administrador (botão direito → "Executar como administrador")

**"PDF não foi processado" ou "Página ignorada":**
- Verifique se o PDF tem tabelas estruturadas
- Confirme se as células têm quebra de linha (\n) separando campo e valor
- Alguns PDFs podem ter layout diferente do esperado
- Primeira linha deve ter dados com quebra de linha

**"Não encontrou os dados" ou "dados em branco":**
- Confira se os nomes dos campos estão exatamente como esperado
- Verifique se as informações estão no formato `Campo\nValor`
- Campos obrigatórios: Data, Data/Prazo, Rodovia, Trecho, Sentido, Classificação, Tipo, Descrição, Origem, Executor

**"Erro ao processar Trecho/Classificação":**
- Trecho deve ter formato: `texto<>texto+texto` 
- Classificação deve ter: `texto E: valor`
- Descrição deve ter quebra de linha para pegar a segunda parte

**"'python' não é reconhecido":**
- Python não está instalado ou não está no PATH do Windows
- Reinstale o Python marcando "Add to PATH"
- Teste no cmd: `python --version`

## O que aparece na tela quando roda

O programa mostra:
- Quais PDFs estão sendo processados
- Se alguma página foi ignorada (e por quê)
- Quantas ordens foram processadas no total  
- Tempo que levou para fazer tudo
- Confirmação quando salvou e abriu a planilha

## Para desenvolvedores

# O script foi feito sob medida para o PDF das ordens de serviço exportadas com uma formatação específica, não vai funcionar para outro PDF ou outra planilha EXCEL.

Se você quiser mexer no código ou melhorar algo:
- O arquivo principal é `projeto.py`
- Use as bibliotecas: pdfplumber, openpyxl, tkinter
- Fique a vontade para aprender e facilitar algum processo da sua empresa.
- A lógica principal está no loop que processa cada PDF
- Formatações específicas estão no final do loop principal

---

**Programa feito para facilitar o trabalho com ordens de serviço**
