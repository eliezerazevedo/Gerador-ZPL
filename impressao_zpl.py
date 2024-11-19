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
conteudo_pagina2 = gerar_code128()

# Substitui as variáveis do ZPL com os conteúdos desejados
zpl_code = f"""
^XA
^FWB                     // Define a orientação para horizontal
^PW800                   // Largura (10 cm convertidos para dots, supondo 203 dpi)
^LL440                   // Altura (5.5 cm convertidos para dots, supondo 203 dpi)
^FX Primeira Página
^FO50,50                 // Posição do texto
^A0N,50,50               // Fonte com altura e largura ajustadas
^FDPrimeira Página: produto teste^FS
^XZ

^XA
^FWB                     // Define a orientação para horizontal
^PW800                   // Largura (10 cm convertidos para dots, supondo 203 dpi)
^LL440                   // Altura (5.5 cm convertidos para dots, supondo 203 dpi)
^FX Segunda Página
^FO50,50
^A0N,50,50
^FDSerial:^FS
^FO50,100
^A0N,50,50
^FD{conteudo_pagina2}^FS
^FO50,150
^BCR,100,Y,N,N          // Código de barras Code128
^FD>: {conteudo_pagina2}^FS
^XZ
"""

# Chama a função para imprimir
imprimir_zpl(zpl_code)
