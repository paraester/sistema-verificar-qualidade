# sistema-verificar-qualidade
## How-To Detalhado: Sistema de Qualidade com Tkinter, Pandas e Openpyxl

### Visão Geral
O projeto "Sistema de Qualidade" consiste em scripts que realizam a manipulação e análise de arquivos CSV e planilhas Excel baseadas em um modelo predefinido. O sistema foi desenvolvido com as seguintes bibliotecas principais:
- **Tkinter**: Criação de interface gráfica (GUI).
- **Pandas**: Manipulação de dados tabulares (CSV).
- **Openpyxl**: Criação e edição de planilhas Excel.
- **Ttkbootstrap**: Personalização de temas visuais para a interface gráfica.

A ideia é automatizar a geração e tratamento de dados de qualidade, criando relatórios personalizados em planilhas Excel, onde cada pessoa ou área de responsabilidade tem sua própria aba.

### Estrutura do Projeto

#### 1. **interface_tk-2.py**
Este arquivo é responsável pela **interface gráfica** do sistema e pelas principais interações do usuário. Ele inclui os seguintes componentes:
- **Gerar arquivo de qualidade**: Aciona a criação do arquivo CSV com dados de qualidade.
- **Tratar arquivo de qualidade**: Aplica regras de tratamento e filtragem ao arquivo CSV.
- **Carregar áreas responsáveis**: Lê e lista as áreas disponíveis para filtragem no arquivo.
- **Criar abas no Excel**: Gera uma planilha Excel com abas individuais para cada nome listado no arquivo `lista_nomes.txt`.

#### 2. **tk_modelo_planilha_dado_lista-de-nomes_tabela-total.py**
Este arquivo gerencia a **criação das planilhas Excel** baseadas em um **modelo predefinido** (`modelo.xlsx`). Ele utiliza um arquivo CSV com dados (`total.csv`) e cria abas na planilha para cada nome listado no arquivo `lista_nomes.txt`. As principais funcionalidades incluem:
- **Criação de Abas**: Para cada nome na lista, é criada uma aba na planilha Excel.
- **Preenchimento de Dados**: A partir do CSV, os dados relacionados ao nome são preenchidos nas células da aba.
- **Tratamento de Nomes Duplicados**: Caso haja nomes iguais na lista, o sistema cria uma aba diferenciando com o segundo nome.
- **Substituição de Placeholder**: Substitui campos como `<Nome Completo>` e `<nome da área responsável>` no modelo de planilha.

### Estrutura de Arquivos e Pastas
A seguir, a estrutura dos arquivos e pastas que compõem o sistema:

```plaintext
.
├── Dados
│   ├── total.csv
│   ├── resultado_planilha.xlsx
├── lista_nomes.txt
├── area.txt
├── modelo.xlsx
├── interface_tk-2.py
├── tk_modelo_planilha_dado_lista-de-nomes_tabela-total.py
├── log_script.log
```

#### Arquivos Explicados:
- **`total.csv`**: Contém todos os dados que serão utilizados para gerar os relatórios e preencher as abas das planilhas.
- **`lista_nomes.txt`**: Lista de nomes dos indivíduos ou entidades para os quais serão criadas abas personalizadas.
- **`area.txt`**: Define a **área responsável** a ser incluída nas planilhas e no nome do arquivo de saída.
- **`modelo.xlsx`**: Modelo da planilha Excel que será replicado e modificado para gerar os relatórios.
- **`interface_tk-2.py`**: Script que controla a interface gráfica e as funções de tratamento e geração de arquivos.
- **`tk_modelo_planilha_dado_lista-de-nomes_tabela-total.py`**: Script que manipula os dados CSV e cria as planilhas Excel.
- **`log_script.log`**: Arquivo de log onde são registrados os eventos e erros durante a execução dos scripts.

### Detalhamento das Funcionalidades

#### 1. Interface Gráfica - `interface_tk-2.py`

##### Geração de Arquivo de Qualidade
- O botão **"Gerar arquivo Qualidade"** aciona um subprocesso que executa o script responsável pela coleta e geração dos dados de qualidade em formato CSV.
- Este arquivo é armazenado na pasta `Dados` com o nome `total.csv`.

##### Tratamento de Arquivo
- O botão **"Tratar arquivo Qualidade"** processa o arquivo `total.csv`, aplicando as regras de tratamento e filtragem. As funcionalidades incluem a **inversão de nomes**, **filtragem de áreas responsáveis**, e a **remoção de linhas desnecessárias**.

##### Carregar Áreas
- **"Carregar Áreas"** exibe na interface as áreas responsáveis disponíveis no arquivo CSV, possibilitando que o usuário escolha uma área específica para gerar relatórios personalizados.

##### Criar Planilhas Personalizadas
- **"Criar Abas no Excel"** lê o arquivo `lista_nomes.txt` e gera abas separadas para cada nome na lista, preenchendo-as com os dados correspondentes do `total.csv`.

##### Barra de Progresso
- Uma **barra de progresso** exibe o andamento da criação das planilhas à medida que as abas são geradas.

#### 2. Manipulação de Planilhas Excel - `tk_modelo_planilha_dado_lista-de-nomes_tabela-total.py`

##### Criação das Abas
- Para cada nome em `lista_nomes.txt`, o sistema verifica se o nome existe no arquivo `total.csv`.
  - Se existir, uma aba será criada e preenchida com os dados correspondentes.
  - Se não existir, uma aba vazia será criada, mantendo a personalização com o nome.

##### Tratamento de Nomes Duplicados
- Quando dois nomes na lista têm o mesmo primeiro nome, o sistema cria abas diferenciadas adicionando o **segundo nome** no título da aba, evitando duplicidades.

##### Modelo de Planilha
- As abas são criadas com base em um modelo pré-definido (`modelo.xlsx`), onde placeholders são substituídos pelos dados reais de cada indivíduo.
- O campo `<Nome Completo>` é substituído pelo nome da pessoa, e o campo `<nome da área responsável>` recebe o valor contido no arquivo `area.txt`.

##### Salvamento do Arquivo
- Ao final do processo, o arquivo Excel é salvo com o nome da **área responsável** (definido em `area.txt`) na pasta `Dados`, com o formato `resultado_planilha.xlsx`.

### Como Usar o Sistema

#### 1. Instalar Dependências
Para rodar o projeto, é necessário instalar as seguintes bibliotecas Python:
```bash
pip install pandas openpyxl ttkbootstrap tkinter
```

#### 2. Executar a Interface Gráfica
Com todas as dependências instaladas, execute o script `interface_tk-2.py` para abrir a interface gráfica do sistema:
```bash
python interface_tk-2.py
```

#### 3. Gerar o Arquivo de Qualidade
1. Clique em **"Gerar arquivo Qualidade"** para iniciar a geração dos dados CSV.
2. O arquivo será salvo na pasta `Dados` com o nome `total.csv`.

#### 4. Tratar o Arquivo
1. Após gerar o arquivo CSV, clique em **"Tratar arquivo Qualidade"** para aplicar as regras de tratamento.

#### 5. Carregar Áreas e Gerar Arquivos Filtrados
1. Clique em **"Carregar Áreas"** para visualizar as áreas disponíveis no arquivo CSV.
2. Selecione uma área e clique em **"Gerar Arquivo Filtrado"** para salvar um arquivo CSV filtrado com base na área escolhida.

#### 6. Criar Planilhas Excel Personalizadas
1. Certifique-se de que o arquivo `lista_nomes.txt` contém os nomes que deseja processar.
2. Clique em **"Criar Abas no Excel"** para gerar a planilha.
3. O sistema criará uma aba para cada nome, utilizando os dados do CSV e o modelo Excel.

#### 7. Arquivo Final
O arquivo Excel gerado será salvo com o nome da área responsável (definido em `area.txt`) na pasta `Dados`, ex.: `COS-ADM.xlsx`.

### Logs e Depuração
O arquivo `log_script.log` registra todas as operações realizadas durante a execução dos scripts, incluindo erros. Se algo der errado, verifique este arquivo para mais detalhes.

### Considerações Finais
Este sistema permite automatizar a criação de relatórios personalizados em planilhas Excel com base em dados de qualidade. A flexibilidade de leitura de listas de nomes e filtragem por áreas torna o sistema altamente customizável, útil para a criação de relatórios gerenciais, técnicos e administrativos.
