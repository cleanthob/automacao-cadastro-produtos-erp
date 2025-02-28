import pandas as pd

# Ajustando os dados para garantir que todas as listas tenham o mesmo tamanho
data = {
    "ID": [1, 2, 3, 4, 5, 6, 7, 8, 9],
    "Nome": [
        "Câmera",
        "Teclado",
        "Microfone",
        "Notebook",
        "Monitor",
        "Hub USB",
        "Adaptador de Rede",
        "Ring Light",
        "Celular",
    ],
    "Categoria": [
        "Acessórios",
        "Acessórios",
        "Acessórios",
        "Principal",
        "Principal",
        "Acessórios",
        "Acessórios",
        "Acessórios",
        "Principal",
    ],
    "GTIN": [
        66020915,
        80642885,
        12309120,
        41230927,
        69104107,
        80521723,
        97591129,
        36957231,
        48447371,
    ],
    "Supplier": [127, 350, 132, 127, 320, 208, 127, 350, 382],
    "Descrição": [
        "Câmera 1080p Logitech USB",
        "Teclado tkl Branco Sem Fio",
        "HyperX Solocast - Microfone condensador USB",
        "Notebook Acer Aspire 5 Intel Core i5-10210U, 4GB RAM, 256 SSD",
        "Monitor 27 LG Full hd, LED, IPS, 75 Hz, 5ms, Freesync, hdmi, vga",
        "Dock Hub Usb Tipo C Para Usb 3.0 Com Hdmi 4k",
        "Adaptador de rede USB 4AB femea W1272 Multilaser CX 1 UN",
        "Kit Iluminador Led Ring Light Tripé Suporte Celular Estetica",
        "Iphone 8 Plus, 128GB, Cinza Espacial",
    ],
    "Imagem": [
        r"imgs_produtos\camera.jpg",
        r"imgs_produtos\teclado.jpg",
        r"imgs_produtos\mic.jpg",
        r"imgs_produtos\notebook.jpg",
        r"imgs_produtos\monitor.jpg",
        r"imgs_produtos\hubusb.jpg",
        r"imgs_produtos\adaptadordere.jpg",
        r"imgs_produtos\ringlight.jpg",
        r"imgs_produtos\iphone.jpg",
    ],
    "Preço": [500, 300, 800, 3500, 1200, 200, 100, 300, 3000],
    "Custo": [55.6, 50, 200, 1200, 600, 40, 20, 80, 1000],
    "Estoque": [800, 500, 300, 100, 50, 600, 800, 500, 100],
}

# Criando o DataFrame corrigido
df = pd.DataFrame(data)

# Caminho para salvar o arquivo
file_path = "produtos.xlsx"

# Salvando a planilha Excel
df.to_excel(file_path, index=False)

# Retornar o caminho do arquivo gerado
file_path
