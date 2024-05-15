import re
import sap
import os
import logging
import env
import time
import utils

# ? Informações: módulo responsável pelo tratamento dos arquivos de background (conferência de jobs, exportar/excluir arquivos)


def __get_spool_list():
    """  
    Este método realiza a consulta no SAP Gui das ordens spool do usuário [SP02] para obter número e título dos itens. É essencial para identificar os dados resultantes de cada transação executada em background (neste contexto, a partir do campo "Título" setado a cada parametrização da execução em background, contendo o código da transação)
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)

    Returns:
        object: contendo o número e nome (título) das ordens spool [Exemplo: {56699555: 09_2023_IW67}]
    """

    while True:
        try:
            session = sap.get_session_by_number('ERP', 1)
            columnIndex = __get_sap_collection_column_index(['Nº spool', 'Título'])
            session.StartTransaction("SP02")
            session.findById("wnd[0]/tbar[1]/btn[45]").press()  # refresh
            spoolList = {}

            row = 3
            while True:
                number = session.FindById(f'wnd[0]/usr/lbl[{columnIndex["Nº spool"]},{row}]', False)
                name = session.FindById(f'wnd[0]/usr/lbl[{columnIndex["Título"]},{row}]', False)
                if number == None or name == None:
                    break
                spoolList[number.text.strip()] = name.text.strip()
                row += 1

            if len(spoolList) != 0:
                return spoolList

        except Exception:
            continue


def __get_quantity_sap_spool():

    while True:
        try:
            partialText = ["Ordens spool exibidas", "Ordem spool exibida"]
            session = sap.get_session_by_number('ERP', 1)
            session.findById("wnd[0]/tbar[0]/btn[83]").press()  # last page
            data = session.FindById("wnd[0]/usr").Children
            for item in data:
                itemName = str(item.text.strip())
                if any(text.upper() in itemName.upper() for text in partialText):
                    return int(itemNamePrev)

                itemNamePrev = itemName

            return 0

        except Exception:
            continue


def __get_sap_collection_column_index(columnTitle: list):
    """  
    Este método realiza a consulta da lista de jobs do usuário [SMX] para identificar título e indíce da coluna de cada item a partir de uma verificação via regex. A exibição dos dados da transação SMX não implementa o objeto GuiGridView, o que impede a consulta direta e específica dos dados.
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)

    Args:
        columnTitle (list): lista de títulos para consulta

    Returns:
        object: contendo título e índice da coluna (index) dos jobs [Exemplo: {Status: 4}]
    """

    while True:
        try:
            columnIndexRegex = r'\[(\d+),\d+\]'
            session = sap.get_session_by_number('ERP', 1)
            data = session.FindById("wnd[0]/usr").Children
            columnIndex = {}
            for item in data:
                for title in columnTitle:
                    itemName = item.text.strip()
                    if itemName.upper() == title.upper():
                        columnIndex[title] = int(re.search(columnIndexRegex, item.id).group(1))
                        columnTitle.remove(title)
                        break
                if not columnTitle:
                    break

            return columnIndex

        except Exception:
            continue


def __export_all_spool_files():
    """  
    Este método realiza a exportação dos arquivos das ordens spool para arquivos .txt [SP02] para o diretório default do SAP Gui. Segue a ordens de passos: Ordem de spool > Transferir > Exportar como texto
    ATENÇÃO: os arquivos exportados constam no diretório default do SAP para exportação (ver env.DIR_SPOOL_DATA)
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)

    Returns:
        bool: True se executado com sucesso
    """

    while True:
        try:
            start = time.time()
            __remove_all_exported_files()
            print('Starting exportation spool files')
            session = sap.get_session_by_number('ERP', 1)
            session.StartTransaction("SP02")
            quantitySapSpool = __get_quantity_sap_spool()
            session.findById("wnd[0]/tbar[1]/btn[45]").press()
            session.findById("wnd[0]/tbar[1]/btn[48]").press()
            session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[1]").select()

            quantityExportedSpool = len(os.listdir(env.DIR_SPOOL_DATA))
            if quantitySapSpool == quantityExportedSpool:
                print(f'{utils.CustomMessage.prGreen("Successfully")} spool files exportation [{quantityExportedSpool}] [{time.strftime("%H:%M:%S", time.gmtime(time.time()-start))}]')
                return True
            else:
                print(utils.CustomMessage.prRed(f'Error to export spool files: {quantityExportedSpool}/{quantitySapSpool}'))
                raise Exception

        except Exception:
            continue


def __rename_exported_files():
    """  
    Este método realiza a renomeação dos arquivos .txt exportados [SP02] para o padrão `numeroOrdemSpool_tituloOrdemSpool.txt`. É essencial para identificar, dentro os arquivos exportados, qual transação originou os dados (para isso o campo "Título" deve ser setado durante a parametrização da execução em background contendo o nome da transação)
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)

    Returns:
        bool: True se executado com sucesso
    """

    while True:
        try:
            print('Starting renaming of exported spool files')
            spoolList = __get_spool_list()
            with os.scandir(env.DIR_SPOOL_DATA) as entries:
                for entry in entries:
                    for spool in spoolList:
                        if spool in entry.name:
                            newPathName = entry.path.replace(entry.name, f'{spool}_{spoolList[spool]}.txt')
                            os.rename(entry.path, newPathName)
            print(f"{utils.CustomMessage.prGreen('Successfully')} spool files rename")
            return True

        except Exception:
            logging.exception('Exception ocurred')
            continue


def __remove_all_jobs():
    """  
    Este método realiza a remoção de todos os jobs do usuário [SMX]
    ATENÇÃO: jobs com status "ativo" não podem ser excluídos, deve-se aguardar a conclusão ou realizar o cancelamento
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)

    Returns:
        bool: True se executado com sucesso
    """

    while True:
        try:
            print('Removing all SAP user jobs')
            session = sap.get_session_by_number('ERP', 1)
            session.StartTransaction("SMX")
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[0]/mbar/menu[1]/menu[11]").select()
            session.findById("wnd[0]/tbar[1]/btn[14]").press()
            btnRemoveOne = session.findById("wnd[1]/usr/btnSPOP-OPTION1", False)
            if btnRemoveOne != None:
                btnRemoveOne.press()

            info = session.findById("wnd[0]/usr/lbl[2,3]", False)
            if info != None and info.text == 'Lista não contém dados':
                return True

        except Exception:
            continue


def __remove_all_spools():
    """  
    Este método realiza a remoção de todos as ordens spool do usuário [SP02]
    ATENÇÃO: dados das ordens spool removidas não poderão ser recuperados, mesmo com jobs concluídos (necessário executar novamente em background)
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)

    Returns:
        bool: True se executado com sucesso
    """

    while True:
        try:
            print('Removing all SAP user spools')
            session = sap.get_session_by_number('ERP', 1)
            session.StartTransaction("SP02")
            session.findById("wnd[0]/tbar[1]/btn[45]").press()
            session.findById("wnd[0]/tbar[1]/btn[48]").press()
            session.findById("wnd[0]/tbar[1]/btn[14]").press()
            btnRemoveOne = session.findById("wnd[1]/usr/btnSPOP-OPTION1", False)
            btnRemoveAll = session.findById("wnd[1]/usr/btnSPOP-OPTION3", False)
            if btnRemoveAll != None:
                btnRemoveAll.press()
            elif btnRemoveOne != None:
                btnRemoveOne.press()

            info = session.findById("wnd[0]/usr/lbl[2,3]", False)
            if info != None and info.text == 'Lista não contém dados':
                return True

        except Exception:
            continue


def __remove_all_exported_files():
    """  
    Este método realiza a remoção dos arquivos .txt exportados para o diretório default do SAP Gui
    ATENÇÃO: os arquivos exportados constam no diretório default do SAP para exportação (ver env.DIR_SPOOL_DATA)
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)

    Returns:
        bool: True se executado com sucesso
    """

    while True:
        try:
            print('Removing all SAP exported files')
            with os.scandir(env.DIR_SPOOL_DATA) as entries:
                for entry in entries:
                    os.remove(entry.path)
            return True

        except Exception:
            continue


def __await_all_job_conclusion():
    """  
    Este método realiza a conferência, a cada x segundos,  se todos os jobs do usuário [SMX] estão concluídos (status "Concl."). É necessário assegurar que todos os jobs estão concluídos antes de realizar a exportação dos dados pela transação SP02
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)
    """

    while True:
        try:
            start = time.time()
            print('Starting unfinished jobs conference')
            awaitSeconds = 3
            while not __is_all_job_concluded():
                for sec in reversed(range(awaitSeconds)):
                    print(
                        f'There are still {utils.CustomMessage.prYellow("unfinished")} jobs. Reassessing in {utils.CustomMessage.prYellow(sec)} seconds', end="\r")
                    time.sleep(1)
            print(
                f'{utils.CustomMessage.prGreen("Successfully")} jobs finished [{time.strftime("%H:%M:%S", time.gmtime(time.time()-start))}]')
            break

        except Exception:
            continue


def __is_all_job_concluded():
    """
    Este método realiza a conferência se todos os jobs do usuário [SMX] estão concluídos (status "Concl.")
    ATENÇÃO: a lista completa de jobs da transação SMX devem estar visíveis na tela do SAP (sem scroll)
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)

    Returns:
        bool: True se executado com sucesso
    """

    while True:
        try:
            session = sap.get_session_by_number('ERP', 1)
            if session.Info.Transaction != 'SMX':
                session.StartTransaction("SMX")

            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            columnIndex = __get_sap_collection_column_index(['Status'])

            row = 3
            while True:
                status = session.FindById(f'wnd[0]/usr/lbl[{columnIndex["Status"]},{row}]', False)
                if status == None:
                    break
                if status.text.strip().upper() != "Concl.".upper():
                    return False
                row += 1

            return True

        except Exception:
            continue

# ? ==========================================================================================

def export_files():
    """  
    Este método realiza a conferência e exportação dos dados das ordens spool do usuário [SP02]
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)

    Returns:
        bool: True se executado com sucesso
    """

    while True:
        try:
            __await_all_job_conclusion()
            __export_all_spool_files()
            __rename_exported_files()
            return True

        except Exception:
            continue


def remove_trash():
    """  
    Este método realiza a remoção de jobs, ordens spools e arquivos exportados em preparação para execução em background
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)

    Returns:
        bool: True se executado com sucesso
    """

    while True:
        try:
            __remove_all_jobs()
            __remove_all_spools()
            __remove_all_exported_files()
            return True

        except Exception:
            continue
