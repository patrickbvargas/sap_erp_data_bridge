import multiprocessing
import threading
import utils
import logging

# ? Informações: módulo principal responsável por executar tarefas simultâneas ou paralelas


def run_multiprocess(arrTaskConfig: list):
    """  
    Este método realiza a criação e start de processos para execução simultânea de tarefas (multiprocessing)
    ATENÇÃO: ações em multiprocessing permitem execução em várias telas do SAP simultaneamente, contudo é necessário que cada método receba um objeto Session diferente (ou um indicador de qual Session conectar - sessionNumber)

    Args:
        arrTaskConfig (list): lista de objetos ThreadConfig com métodos e argumentos para execução
    """

    try:
        multiprocessing.freeze_support()

        arrProcess = []
        for tc in arrTaskConfig:
            p = multiprocessing.Process(target=tc.callback, args=tc.args)
            p.start()
            arrProcess.append(p)
            print(f'{utils.CustomMessage.prGreen("Successfully")} created process [PID {p.pid}]')

        for p in arrProcess:
            p.join()

    except Exception:
        logging.exception('Exception occurred')
        raise Exception('Exception occurred')


def run_multithread(arrTaskConfig: list):
    """  
    Este método realiza a criação e start de threads para execução paralela de tarefas (multithreading)
    ATENÇÃO: ações em multitreading permitem execução em várias telas do SAP paralelamente, contudo é necessário que cada método receba um objeto Session diferente (ou um indicador de qual Session conectar - sessionNumber)

    Args:
        arrTaskConfig (list): lista de objetos ThreadConfig com métodos e argumentos para execução
    """

    try:
        arrThreads = []
        for tc in arrTaskConfig:
            t = threading.Thread(target=tc.callback, args=tc.args)
            t.start()
            arrThreads.append(t)
            print(f'{utils.CustomMessage.prGreen("Successfully")} created thread [TID {t.native_id}]')

        for t in arrThreads:
            t.join()

    except Exception:
        logging.exception('Exception occurred')
        raise Exception('Exception occurred')
