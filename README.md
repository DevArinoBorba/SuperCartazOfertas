# Super Cartaz Ofertas 🏷️

O **Super Cartaz Ofertas** é um sistema web e PWA (Progressive Web App) desenvolvido para simplificar a criação, gerenciamento e impressão de cartazes de ofertas para redes de supermercados ou lojas físicas. Ele foi projetado para permitir que o setor de marketing ou a administração cadastre ofertas de forma centralizada e que cada filial (loja) acesse e imprima seus respectivos cartazes com layouts padronizados e profissionais.

---

## 🛠️ Tecnologias Utilizadas

O sistema foi construído utilizando as seguintes tecnologias:

*   **Backend (Python/Flask)**: Servidor leve e ágil utilizando o microframework [Flask](https://flask.palletsprojects.com/) para controle de rotas, sessões e lógica de negócios.
*   **Banco de Dados (SQLite)**: Banco de dados relacional embarcado e rápido, ideal para implantações locais e gerenciamento simplificado.
*   **Frontend (HTML5, CSS3, Javascript)**: Interface responsiva, moderna e dinâmica com templates [Jinja2](https://jinja.palletsprojects.com/) e folha de estilo customizada para impressão precisa.
*   **PWA (Progressive Web App)**: Service Worker local (`sw.js`) e arquivo de manifesto (`manifest.json`) que permitem que o aplicativo seja instalado em computadores e celulares e funcione offline para consulta.
*   **Importação e Processamento de Arquivos**: Utilização da biblioteca `openpyxl` e módulo `csv` para leitura e importação ágil de planilhas Excel (.xlsx) e CSV de ofertas.

---

## 🚀 Funcionalidades do Sistema

O sistema possui controle de acesso baseado em níveis de permissão (Administrador vs. Filial/Loja):

### 🔑 Para Administradores

*   **Painel Administrativo (Dashboard)**: Estatísticas rápidas de ofertas cadastradas, ofertas expirando hoje, ofertas vencidas e distribuição de cartazes por filial.
*   **Gerenciamento Manual de Ofertas**: Criação, edição, exclusão e duplicação de ofertas individuais.
*   **Duplicação em Lote (Lote)**: Seleção de múltiplas ofertas de um dia e duplicação para qualquer outra data com um clique.
*   **Importação via Planilha (CSV / Excel)**: Importação em lote de ofertas a partir de arquivos `.csv` ou `.xlsx` com mapeamento automático inteligente de colunas (com suporte a sinônimos de colunas).
*   **Exportação**: Download de relatórios de ofertas em formato CSV.
*   **Editor Visual de Layouts**: Ajuste interativo (posicionamento, fontes, cores e dimensões) para os formatos de cartazes **A4 (Grande)** e **Pequeno**, permitindo salvar um layout padrão global ou customizar especificamente para uma oferta.
*   **Trilha de Auditoria (Logs)**: Histórico completo de ações executadas no sistema (criações, edições, exclusões) contendo usuário, data/hora e detalhes para fins de segurança.
*   **Gerenciamento de Usuários**: Cadastro e exclusão de funcionários, definição de senhas e atribuição a lojas específicas.

### 🏪 Para Filiais (Usuários Comuns)

*   **Painel de Ofertas Filtrado**: Visualização restrita às ofertas ativas destinadas à sua loja específica ou marcadas para "todos".
*   **Impressão de Cartazes**:
    *   Visualização individual de cartaz.
    *   **Impressão em Lote**: Seleção de múltiplos itens para gerar um PDF ou folha de impressão conjunta.
    *   Ajuste dinâmico entre o formato **A4 (Grande)** e **Pequeno**.
*   **Gestão de Perfil**: Alteração de senha de acesso direto pela filial.

---

## 💻 Como Executar o Projeto

1.  **Pré-requisitos**: Certifique-se de ter o Python 3.8+ instalado em sua máquina.
2.  **Instalar dependências**:
    ```bash
    pip install -r requirements.txt
    ```
3.  **Iniciar o Servidor**:
    *   No Windows, você pode rodar o arquivo executável local `iniciar_servidor.bat` ou executar diretamente no terminal:
    ```bash
    python app.py
    ```
4.  **Acesso**: O sistema iniciará por padrão em `http://localhost:5000` (acessível também na rede local através do IP da máquina).
5.  **Credenciais Iniciais (Admin padrão)**:
    *   **Usuário**: `admin`
    *   **Senha**: `admin`

*(Nota: Recomenda-se alterar a senha do administrador após o primeiro acesso).*
