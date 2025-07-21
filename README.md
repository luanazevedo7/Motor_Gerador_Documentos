# Motor de Geração de Documentos Automatizados v1

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?style=for-the-badge&logo=python)

## 📄 Contexto

Este projeto nasceu de uma necessidade real: automatizar a criação de Termos de Compromisso para artistas de uma Cidade selecionados em editais culturais, uma tarefa manual, repetitiva e suscetível a erros. Para transformar a solução em uma ferramenta robusta e reutilizável, o script foi projetado como um **motor flexível**, capaz de gerar qualquer tipo de documento (`.docx`) a partir de dados em uma planilha Excel.

Esta versão aprimora o fluxo de trabalho para geração em massa, introduzindo um modo de preenchimento em grupo que torna o processo ainda mais rápido e eficiente.

## ✨ Funcionalidades

* **Altamente Configurável:** Todas as configurações principais (nomes de pastas, arquivos, colunas, etc.) são definidas em um "Painel de Controle" no topo do script, facilitando a adaptação para qualquer projeto.
* **Leitura Estruturada de Dados:** Utiliza a biblioteca Pandas para ler dados de uma aba específica de uma planilha Excel, garantindo consistência.
* **Interface Interativa:** Apresenta uma lista de itens da planilha e permite a seleção múltipla para processamento.
* **Processamento em Lote Inteligente:** Ao selecionar múltiplos itens, o script pergunta se os dados interativos (como categoria e valor) são os mesmos para todo o grupo, economizando tempo de digitação repetitiva.
* **Geração de Documentos com Formatação Preservada:** Preenche um documento Word modelo (`.docx`), substituindo *placeholders* (ex: `[nome]`) pelos dados corretos e mantendo a formatação original do modelo.
* **Organização de Saída:** Salva todos os documentos gerados em uma pasta dedicada, com nomes de arquivo únicos e prefixo customizável.

## 🚀 Demonstração

*Dica: Tire uma nova captura de tela mostrando a pergunta sobre o 'Modo de Preenchimento' em ação e salve na pasta `exemplos` como `screenshot_demo.png`.*
![Demonstração do Script](exemplos/screenshot_demo.png)

## 🛠️ Tecnologias Utilizadas

* **Python 3:** Linguagem principal do projeto.
* **Pandas:** Para leitura e manipulação dos dados da planilha Excel.
* **python-docx:** Para a criação e manipulação dos documentos Word (`.docx`).
* **openpyxl:** Dependência do Pandas para trabalhar com arquivos `.xlsx`.

## ⚙️ Configurando o Motor

A grande vantagem deste projeto é sua flexibilidade. Todas as customizações são feitas no **Painel de Controle** no topo do arquivo `src/motor_gerador.py`.

* `PASTA_DADOS`, `NOME_ARQUIVO_EXCEL`, `NOME_ABA_EXCEL`: Defina onde estão seus dados de entrada.
* `PASTA_MODELO`, `NOME_ARQUIVO_MODELO`: Defina onde está seu documento Word de modelo.
* `PASTA_SAIDA`, `PREFIXO_ARQUIVO_SAIDA`: Controle o nome e o local dos arquivos gerados.
* `MAPEAMENTO_COLUNAS`: Conecte o nome da coluna no seu Excel (ex: `'Nome do Cliente'`) com o placeholder no seu Word (ex: `'[cliente]'`).
* `COLUNA_IDENTIFICADORA`: Escolha qual coluna da planilha será usada para exibir a lista de seleção para o usuário.
* `CAMPOS_INTERATIVOS`: Uma lista que define quais perguntas o programa fará em tempo de execução.

## 📥 Instalação e Uso

### Pré-requisitos

* Python 3.9 ou superior

### Passos

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/luanazevedo7/Motor_Gerador_Documentos.git](https://github.com/luanazevedo7/Motor_Gerador_Documentos.git)
    cd Motor_Gerador_Documentos
    ```

2.  **Crie e ative um ambiente virtual:**
    ```bash
    # Windows
    python -m venv venv
    .\venv\Scripts\Activate
    ```

3.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Prepare a Estrutura de Pastas e Arquivos:**
    * Crie as pastas `dados` e `modelo` na raiz do projeto.
    * Coloque sua planilha Excel na pasta `dados/`.
    * Coloque seu documento Word modelo na pasta `modelo/`.
    * *Você pode usar os arquivos da pasta `exemplos/` como referência.*

/ (pasta raiz do seu projeto)
|
|-- dados/
|   └── modelo_banco_de_dados.xlsx  <-- COLOQUE SUA PLANILHA AQUI
|
|-- modelo/
|   └── modelo_doc.docx             <-- COLOQUE SEU MODELO WORD AQUI
|
|-- src/
|   └── motor_gerador_documentos.py                (o script que você executa)
|
|-- motor_gerador_documentos.exe                   (executável gerado após ajustar as configurações no "Painel de Controle" ['pyinstaller --onefile motor_gerador_documentos.py'])

5.  **Ajuste as Configurações:** Abra o arquivo `src/motor_gerador.py` e ajuste as variáveis no "Painel de Controle" para corresponder aos seus arquivos e necessidades.

6.  **Execute o programa a partir da pasta raiz do projeto:**

7.  Siga as instruções no terminal. **Se você selecionar mais de um item, o programa perguntará se deseja usar o modo de preenchimento 'Individual' ou 'Em Grupo'**, otimizando seu tempo. Os documentos gerados aparecerão na pasta de saída que você configurou.

## 👨‍💻 Autor

* **Luan Azevedo**
* **LinkedIn:** `https://linkedin.com/in/euluangomes`
* **GitHub:** `https://github.com/luanazevedo7`
