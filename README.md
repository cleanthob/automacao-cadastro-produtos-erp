# 📊 Automacao de Cadastro de Produtos no ERP

## ✨ Índice

- [Descrição](#-descri%C3%A7%C3%A3o)
- [Funcionalidades](#-funcionalidades)
- [Pré-requisitos](#-pr%C3%A9-requisitos)
- [Abrir e rodar o projeto](#%EF%B8%8F-abrir-e-rodar-o-projeto)
- [Personalização](#-personaliza%C3%A7%C3%A3o)
- [Licença](#-licen%C3%A7a)

## 🔍 Descrição

Este projeto automatiza o cadastro de produtos em um sistema **ERP**, preenchendo automaticamente os campos necessários com base em um arquivo **Excel**. Ele utiliza a biblioteca **PyAutoGUI** para interagir com a interface gráfica do sistema e agilizar o processo de inserção de dados.

## 🚀 Funcionalidades
✅ Abertura automática do ERP  
✅ Leitura de dados do arquivo **produtos.xlsx**  
✅ Preenchimento automático dos campos de cadastro de produtos  
✅ Inserção de imagens dos produtos  
✅ Salvamento automático após cada cadastro  

## 📌 Pré-requisitos

Antes de rodar a automação, certifique-se de ter instalado:
- O software **ERP** instalado no caminho correto
- As bibliotecas Python necessárias:
  ```bash
  pip install pyautogui pandas pyperclip openpyxl
  ```

## 🛠️ Abrir e rodar o projeto

1. **Clone o repositório**  
   ```bash
   git clone https://github.com/cleanthob/automacao_cadastro_erp.git
   ```

2. **Acesse o diretório do projeto**  
   ```bash
   cd automacao_cadastro_erp
   ```

3. **Prepare o arquivo `produtos.xlsx`** com os seguintes campos:  
   - Nome  
   - ID  
   - Categoria  
   - GTIN  
   - Supplier  
   - Descrição  
   - Preço  
   - Custo  
   - Estoque  
   - Imagem (caminho do arquivo)  

4. **Execute o script:**  
   ```bash
   python main.py
   ```

## 🎨 Personalização

Caso o software **ERP** esteja instalado em outro local, altere a linha correspondente no código:
```python
subprocess.Popen([r"C:\Caminho\Para\Seu\ERP.exe"])
```
Se precisar ajustar a precisão das detecções de imagem, modifique o parâmetro `confidence` na função `locateOnScreen`:
```python
encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.8)
```

## 📝 Licença

Este projeto está licenciado sob a GPL-3.0. Consulte o arquivo [LICENSE](LICENSE) para mais informações sobre os termos de licenciamento.
