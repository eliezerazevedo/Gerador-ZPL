# type: ignore
import win32print
import random
import string

# Função para gerar um código aleatório Code128
def gerar_code128():
    # Gera uma string aleatória de 10 caracteres alfanuméricos
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=10))

def imprimir_zpl(zpl_content):
    try:
        # Obtém a impressora padrão
        impressora_padrao = win32print.GetDefaultPrinter()
        
        # Abre a impressora
        handle_impressora = win32print.OpenPrinter(impressora_padrao)
        
        # Inicia o trabalho de impressão
        job_id = win32print.StartDocPrinter(handle_impressora, 1, ("ImpressaoZPL", None, "RAW"))
        win32print.StartPagePrinter(handle_impressora)
        
        # Envia o conteúdo ZPL
        win32print.WritePrinter(handle_impressora, zpl_content.encode('utf-8'))
        
        # Finaliza o trabalho
        win32print.EndPagePrinter(handle_impressora)
        win32print.EndDocPrinter(handle_impressora)
        
        print("Impressão enviada com sucesso!")
    except Exception as e:
        print(f"Erro ao imprimir: {e}")
    finally:
        if 'handle_impressora' in locals():
            win32print.ClosePrinter(handle_impressora)

# Gera o código aleatório para o Code 128
conteudo_code128 = gerar_code128()

# Substitui as variáveis do ZPL com os conteúdos desejados
zpl_code = f"""
^XA
^PW400  ; Define a largura da etiqueta para 400 pontos (50mm)
^LL400  ; Define a altura da etiqueta para 400 pontos (50mm)

^FX Top section with logo, name and address.
^CF0,60
^FO10,50^FDPRINT MASTER.^FS
^CF0,30
^FO50,250^GB300,3,3^FS  ; Ajustado para caber na largura de 400

^FX Second section with recipient address and permit information.
^CFA,30

^XZ
^XA

^FX Third section with bar code.
^BY5,2,270
^FO50,150^BC^FD{conteudo_code128}^FS  ; Código de barras na posição correta
^FO50,50^FD{conteudo_code128}^FS  ; Texto do código de barras abaixo dele, ajustado para a largura da etiqueta
^XZ
"""

# Chama a função para imprimir
imprimir_zpl(zpl_code)