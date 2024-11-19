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

^FX Top section with logo, name and address.
^CF0,60
^FO50,50^GB100,100,100^FS
^FO75,75^FR^GB100,100,100^FS
^FO93,93^GB40,40,40^FS
^FO220,50^FDIntershipping, Inc.^FS
^CF0,30
^FO220,115^FD1000 Shipping Lane^FS
^FO220,155^FDShelbyville TN 38102^FS
^FO220,195^FDUnited States (USA)^FS
^FO50,250^GB700,3,3^FS

^FX Second section with recipient address and permit information.
^CFA,30
^FO50,300^FDJohn Doe^FS
^FO50,340^FD100 Main Street^FS
^FO50,380^FDSpringfield TN 39021^FS
^FO50,420^FDUnited States (USA)^FS
^CFA,15
^FO600,300^GB150,150,3^FS
^FO638,340^FDPermit^FS
^FO638,390^FD123456^FS
^FO50,500^GB700,3,3^FS

^XZ
^XA

^FX Third section with bar code.
^BY5,2,270
^FO100,550^BC^FD{conteudo_code128}^FS

^FX Fourth section (the two boxes on the bottom).
^FO50,900^GB700,250,3^FS
^FO400,900^GB3,250,3^FS
^CF0,40
^FO100,960^FDCtr. X34B-1^FS
^FO100,1010^FDREF1 F00B47^FS
^FO100,1060^FDREF2 BL4H8^FS
^CF0,190
^FO470,955^FDCA^FS

^XZ
"""

# Chama a função para imprimir
imprimir_zpl(zpl_code)
