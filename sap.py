import win32com.client
import time
import logging
import utils

# ? Informações: módulo responsável pelo tratamento da conexão e comunicação com SAP GUI Script


def __get_session_by_number(systemName: str, sessionNumber: int):
    """  
    Este método realiza a conexão com o SAP Gui para uma determinada sessão/tela (ses[x])
    ATENÇÃO: não hevará conexão se não identificada a sessão/tela informada. Não está parametrizada a tentativa de abertura de novas sessões/telas (deverá ocorrer manualmente)

    Args:
        systemName (str): nome do sistema SAP (por default ERP)
        sessionNumber (int): número da sessão/tela para realizar conexão

    Returns:
        object: conexão com o SAP Gui
    """

    try:
        sapApp = win32com.client.GetObject('SAPGUI').getScriptingEngine
        for con in sapApp.Connections:
            for ses in con.Sessions:
                if ses.Info.SystemName == systemName.upper() and ses.Info.User and ses.Info.SessionNumber == sessionNumber:
                    return ses

    except Exception as e:
        logging.exception('Exception occurred')
        print(f"Error: {str(e)} [ses {sessionNumber}]")
        return None

    finally:
        sapApp = None
        con = None


def get_session_by_number(systemName: str, sessionNumber: int):
    """  
    Este método realiza uma tentativa, a cada x segundos, de conexão com uma sessão/tela do SAP Gui
    IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de número de tentativas)

    Args:
        systemName (str): nome do sistema SAP (por default ERP)
        sessionNumber (int): número da sessão/tela para realizar conexão

    Returns:
        object: conexão com o SAP Gui
    """

    while True:
        try:
            ses = None
            awaitSeconds = 10
            while ses == None:
                ses = __get_session_by_number(systemName, sessionNumber)
                if ses == None:
                    for sec in reversed(range(awaitSeconds)):
                        print(
                            f'{utils.CustomMessage.prRed("Failed")} to get SAP GUI session [{systemName} - {sessionNumber}]. Retrying in {utils.CustomMessage.prPurple(sec)} seconds', end="\r")
                        time.sleep(1)

            return ses

        except Exception:
            continue

# ? ==========================================================================================


def clear_all_input_text(session: object):
    """ 
    Este método realiza a limpeza de todos os dados preenchidos em uma transação do SAP Gui
    ATENÇÃO: a chamada deste método somente deve ocorrer após a transação estar aberta no SAP

    Args:
        session (object): objeto de conexão com o SAP Gui
    """

    try:
        for obj in session.findById("wnd[0]/usr").Children:
            if (obj.Type == "GuiCTextField" or obj.type == "GuiTextField"):
                try:
                    obj.text = ''
                except:
                    pass

    except Exception:
        logging.exception('Exception occurred')
        return None


def set_user_variant(session: object, variant: str):
    """ 
    Este método realiza a chamada de uma variante de layout para uma transação do SAP Gui
    ATENÇÃO: deve-se informar o nome da variante e não o RE do criador

    Args:
        session (object): objeto de conexão com o SAP Gui
        variant (str): nome da variant de layout
    """

    try:
        clear_all_input_text(session)
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ''
        session.findById("wnd[1]/usr/txtV-LOW").text = variant
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

    except Exception:
        logging.exception('Exception occurred')
        return None


def create_background_job(session: object, jobName: str):
    """ 
    Este método realiza a parametrização de execução em background de uma transação [F9]. O título do job será informado para facilitar a identificação dos dados na transação SP02

    Args:
        session (object): objeto de conexão com o SAP Gui
        jobName (str): nome do job (título)

    Returns:
        bool: True se executado com sucesso
    """

    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select()
        session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").text = ''  # clean local output (LOCAL) to allow set max rows
        session.findById("wnd[1]/tbar[0]/btn[6]").press()  # 'características'
        session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").DoubleClickItem('PAART', 'Column1')
        session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-LINCT").text = '60000'  # max rows allowed for SAP
        session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ExpandNode('SPOOLREQUEST')
        session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").DoubleClickItem('PRTXT', 'Column1')
        session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").Text = jobName.upper()

        session.findById("wnd[2]/tbar[0]/btn[13]").press()
        session.findById("wnd[1]/tbar[0]/btn[13]").press()
        session.findById("wnd[1]/usr/btnSOFORT_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        return True

    except Exception:
        logging.exception('Exception occurred')
        return False
