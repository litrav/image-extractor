# 🖼️ Extrator de Imagens do Excel (por Aba)

Um script Python simples, mas poderoso, que descompacta arquivos `.xlsx` ou `.xlsm`, analisa a complexa estrutura de XML interna e extrai todas as imagens, organizando-as em pastas separadas com base na aba (planilha) a que pertencem.

Chega de "Salvar como página da web" ou de caçar imagens manualmente!

---

## 🚀 Funcionalidades Principais

* **Organização Automática:** Cria uma pasta principal e, dentro dela, subpastas para cada aba da planilha que contém imagens.
* **Mapeamento Avançado:** O script não se limita a links diretos. Ele também navega pelas referências de "Drawings" (`xl/drawings/`) para encontrar imagens inseridas de forma indireta.
* **Fallback Inteligente (Opcional):** Imagens que não podem ser mapeadas a uma aba específica (como logotipos em cabeçalhos, rodapés ou imagens "fantasma" deixadas pelo Excel) são salvas na pasta principal, garantindo que nada seja perdido.
* **Nomenclatura Customizável:** Você pode definir um nome base para os arquivos de imagem extraídos.
* **Standalone:** O script pode ser facilmente compilado em um único arquivo `.exe` usando o PyInstaller, permitindo o uso em qualquer máquina Windows sem a necessidade de instalar Python ou qualquer biblioteca.

---

## ⚙️ Como Usar (Versão `.py`)

### 1. Pré-requisitos

O script utiliza as seguintes bibliotecas Python:

* `openpyxl`: Para ler os nomes das abas da planilha.
* `pyinstaller` (Opcional): Apenas se você quiser compilar o `.exe`.

### 2. Instalação

1.  Clone este repositório:
    ```bash
    git clone [https://github.com/seu-usuario/seu-repositorio.git](https://github.com/seu-usuario/seu-repositorio.git)
    cd seu-repositorio
    ```

2.  (Recomendado) Crie um ambiente virtual:
    ```bash
    python -m venv venv
    venv\Scripts\activate
    ```

3.  Instale as dependências:
    ```bash
    pip install openpyxl
    ```

### 3. Executando o Script

Com seu ambiente virtual ativo, basta rodar:

```bash
python seu_script.py