import pyautogui
import subprocess
import time
import pandas as pd
import pyperclip


def encontrar_imagem(imagem):
    while True:
        try:
            encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9)
            time.sleep(1)

            if encontrou:
                break
        except pyautogui.ImageNotFoundException:
            pass

    encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9)
    return encontrou


def clicar_na_direta(posicoes_imagem):
    return (
        posicoes_imagem[0] + posicoes_imagem[2],
        posicoes_imagem[1] + posicoes_imagem[3] / 2,
    )


def escrever_texto(texto):
    pyperclip.copy(texto)
    pyautogui.hotkey("ctrl", "v")


subprocess.Popen([r"C:\Program Files\Fakturama2\Fakturama.exe"])

encontrou = encontrar_imagem(r"imgs\fakturama.png")

tabela_produtos = pd.read_excel("produtos.xlsx")

for linha in tabela_produtos.index:
    nome = tabela_produtos.loc[linha, "Nome"]
    id_ = tabela_produtos.loc[linha, "ID"]
    categoria = tabela_produtos.loc[linha, "Categoria"]
    gtin = tabela_produtos.loc[linha, "GTIN"]
    supplier = tabela_produtos.loc[linha, "Supplier"]
    descricao = tabela_produtos.loc[linha, "Descrição"]
    preco = tabela_produtos.loc[linha, "Preço"]
    custo = tabela_produtos.loc[linha, "Custo"]
    estoque = tabela_produtos.loc[linha, "Estoque"]
    imagem = tabela_produtos.loc[linha, "Imagem"]

    # Menu New
    encontrou = encontrar_imagem(r"imgs\menu_new.png")
    pyautogui.click(pyautogui.center(encontrou))

    # Menu New Product
    encontrou = encontrar_imagem(r"imgs\new_product.png")
    pyautogui.click(pyautogui.center(encontrou))

    # Preencher todos os campos
    encontrou = encontrar_imagem(r"imgs\item_number.png")
    pyautogui.click(clicar_na_direta(encontrou))
    escrever_texto(str(id_))

    encontrou = encontrar_imagem(r"imgs\name.png")
    pyautogui.click(clicar_na_direta(encontrou))
    escrever_texto(str(nome))

    encontrou = encontrar_imagem(r"imgs\categoria.png")
    pyautogui.click(clicar_na_direta(encontrou))
    escrever_texto(str(categoria))

    encontrou = encontrar_imagem(r"imgs\gtin.png")
    pyautogui.click(clicar_na_direta(encontrou))
    escrever_texto(str(gtin))

    encontrou = encontrar_imagem(r"imgs\supplier.png")
    pyautogui.click(clicar_na_direta(encontrou))
    escrever_texto(str(supplier))

    encontrou = encontrar_imagem(r"imgs\descricao.png")
    pyautogui.click(clicar_na_direta(encontrou))
    escrever_texto(str(descricao))

    encontrou = encontrar_imagem(r"imgs\preco.png")
    pyautogui.click(clicar_na_direta(encontrou))
    preco_texto = f"{preco:.2f}".replace(".", ",")
    escrever_texto(str(preco_texto))

    encontrou = encontrar_imagem(r"imgs\custo.png")
    pyautogui.click(clicar_na_direta(encontrou))
    custo_texto = f"{custo:.2f}".replace(".", ",")
    escrever_texto(str(custo_texto))

    encontrou = encontrar_imagem(r"imgs\estoque.png")
    pyautogui.click(clicar_na_direta(encontrou))
    estoque_texto = f"{estoque:.2f}".replace(".", ",")
    escrever_texto(str(estoque_texto))

    # Selecionar imagem

    encontrou = encontrar_imagem(r"imgs\selecionar_imagem.png")
    pyautogui.click(pyautogui.center(encontrou))

    escrever_texto(
        rf"C:\Users\C\Documents\projetos\automacao_erp\{str(imagem)}"
    )  # Caminho da imagem do produto
    pyautogui.press("enter")

    # Clicar em salvar

    encontrou = encontrar_imagem(r"imgs\save.png")
    pyautogui.click(pyautogui.center(encontrou))
