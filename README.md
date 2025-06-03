# Automatizador de Lembretes de Vencimento via WhatsApp

Este script Python automatiza o envio de lembretes de vencimento para contatos listados em uma planilha Excel, utilizando o WhatsApp Web para o envio das mensagens. Ele verifica quais vencimentos estão a 5 dias de expirar e envia uma mensagem personalizada.

## Funcionalidades

* Lê dados (Nome, Telefone, Data de Vencimento) de uma planilha Excel.
* Calcula se o vencimento ocorre em exatamente 5 dias.
* Abre o WhatsApp Web e envia uma mensagem personalizada para os contatos selecionados.
* Utiliza `pyautogui` para interagir com a interface do WhatsApp Web (localizar e clicar no botão de enviar).
* Registra erros de envio ou de dados em um arquivo `erros.csv`.

## Tecnologias Utilizadas

* Python 3
* `openpyxl`
* `pyautogui`
* `webbrowser`
* Módulos padrão: `urllib.parse`, `time`, `datetime`

## Pré-requisitos

* Python 3.6 ou superior instalado.
* Navegador web padrão configurado.
* Sessão do WhatsApp Web **ativa e logada** no navegador padrão.
* A tela do computador deve estar **ligada e desbloqueada** durante a execução do script, e a janela do navegador onde o WhatsApp Web abrir deve estar visível.

## Configuração do Projeto

1.  **Clone o Repositório:**
    ```bash
    # Exemplo se estiver usando git
    # git clone SEU_LINK_PARA_O_REPOSITORIO_AQUI.git
    # cd NOME_DA_PASTA_DO_PROJETO
    ```

2.  **Crie um Ambiente Virtual (Recomendado):**
    ```bash
    python -m venv venv
    ```
    Ative o ambiente:
    * Windows: `venv\Scripts\activate`
    * macOS/Linux: `source venv/bin/activate`

3.  **Instale as Dependências:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Ajuste os Caminhos no Script (IMPORTANTE):**
    O script `automatizador_wpp.py` contém caminhos fixos para os arquivos. **Você precisará editar o script e alterar esses caminhos** para que correspondam a localização dos arquivos no seu computador:
    * Linha do `openpyxl.load_workbook(...)`: Altere `'D:/PyProjetos/pjt007/clientes_exemplo.xlsx'` para o caminho completo do seu arquivo Excel.
    * Linha do `pyautogui.locateOnScreen(...)`: Altere `'D:/PyProjetos/pjt007/img/1.PNG'` para o caminho completo da sua imagem do botão "enviar".

5.  **Prepare o Arquivo de Clientes (`clientes_exemplo.xlsx`:**
    * O script, como fornecido, está configurado para ler o arquivo `D:/PyProjetos/pjt007/clientes_exemplo.xlsx`. Certifique-se de que este arquivo exista no caminho especificado ou altere o caminho no script.
    * Este arquivo deve conter as seguintes colunas na primeira aba (por padrão, 'Folha1'):
        * **Coluna A:** Nome do Contato
        * **Coluna B:** Telefone do Contato (com código do país e DDD, ex: `55119XXXXXXXX`)
        * **Coluna C:** Data de Vencimento (formatada como data no Excel, ex: `DD/MM/AAAA`)
    * A leitura dos dados começa da segunda linha (a primeira linha é assumida como cabeçalho).

6.  **Configure a Imagem do Botão de Envio (`img/1.PNG`):**
    * O script procura pela imagem em `'D:/PyProjetos/pjt007/img/1.PNG'`.
    * **IMPORTANTE:** É **altamente provável** que você precise substituir esta imagem por uma captura de tela do botão "Enviar" feita no **seu próprio computador**, com a sua resolução de tela, tema do sistema operacional e zoom do navegador.
    * Salve sua captura de tela no caminho que você especificou no script.
    * O script usa uma confiança de `0.9` para localizar a imagem. Se tiver problemas, você pode tentar ajustar este valor diretamente no código (parâmetro `confidence`) ou garantir que sua captura de tela seja bem nítida.

## Como Executar

1.  Certifique-se de que todos os pré-requisitos e configurações (especialmente os caminhos no script) foram atendidos.
2.  Abra um terminal ou prompt de comando.
3.  Navegue até a pasta onde você salvou o script `automatizador_wpp.py`.
4.  Se estiver usando um ambiente virtual, ative-o.
5.  Execute o script:
    ```bash
    python automatizador_wpp.py
    ```
6.  Acompanhe os `print`s no console. O arquivo `erros.csv` será criado na mesma pasta do script (ou onde o script tiver permissão para escrever) caso ocorram erros de dados ou de envio.
7.  **Lembrete:** Para que a lógica de "lembrete com 5 dias de antecedência" funcione continuamente, o script precisa ser executado **diariamente**.

## Observações e Limitações

* **Dependência da Interface do WhatsApp Web:** Mudanças na interface do WhatsApp podem quebrar a localização da imagem do botão de enviar.
* **Visibilidade da Tela:** A janela do navegador com o WhatsApp Web não deve estar minimizada ou obstruída.
* **Caminhos Fixos:** Este script usa caminhos de arquivo fixos. Para usá-lo em outro computador ou para maior portabilidade, esses caminhos dentro do arquivo Python precisam ser editados manualmente.
* **Uso Responsável:** Utilize com ética e responsabilidade.

---

