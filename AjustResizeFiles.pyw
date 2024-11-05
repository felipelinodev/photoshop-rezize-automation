from win32com.client import Dispatch
import os 
import time
#AQUI GUARDA O DIRETORIO ONDE ESTÁ O ARQUIVO PYTHON
caminho = os.path.dirname(os.path.realpath(__file__))

#INICIALIZA O PHOTOSHOP E PSD FILE
ps = Dispatch("Photoshop.Application")

def ExeGeral(manga_cacom, smasc, sinf, sfem):
    #AQUI ADICIONA (MANGA _D.psd) O ARQUIVO AO CAMINHO
    caminho_com = caminho + manga_cacom
    doc = ps.Open(caminho_com)

    def executador(acao, caminho_func):
        ps.DoAction(acao, "REDIMENCIONAR")
        time.sleep(1.5)
        #TEMPO DE ESPERA PARA A ACTION SER EXECUTADA
        if acao == "FEM":
            docl = ps.ActiveDocument.ArtLayers["ESCRITA_MANGA"]
            docl.Translate(0, 1.4)
        else:
            print("DEU ERRO")
        
        options = Dispatch('Photoshop.PNGSaveOptions')
        options.compression = 4
        pngfile = caminho + caminho_func
        doc.SaveAs(pngfile, options, True)
    #CHAMO FUNÇÃO E MUDO OS ARGUMENTOS
    executador("MASC", smasc)
    executador("INF", sinf) #Essa inf é depedende da masc, ela reaproveita a (altura) da mesma.
    executador("FEM",  sfem)
    time.sleep(1.5)
    doc.Close(2)
 #CHAMO FUNÇÃO GERAL, COM OS ARGUMENTOS 1: DO PSD PARA PODER ABRIR O ARQUIVO, E OS OUTROS PARA EXPORTAÇÃO DOS AQUIVOS
ExeGeral("\\molde_manga_longa_d.psd", "\\molde_manga_longa_d.png", "\\molde_manga_longa_d_INF.png", "\\molde_manga_longa_d_FEM.png")
ExeGeral("\\molde_manga_longa_e.psd", "\\molde_manga_longa_e.png", "\\molde_manga_longa_e_INF.png", "\\molde_manga_longa_e_FEM.png")

ps.Quit()




