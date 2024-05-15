import datetime
import fileCrud
import logging
import time
import numpy
import math
import background
import env
import utils
import multitask
import calendar
from model import TaskConfig, ReferenceInfo
from parameters import  IW67ByMeasurementMEDL, IW67ByMeasurementMEDE

# ? Informações: módulo principal responsável por executar os scripts

logging.basicConfig(filename="log.txt", level=logging.DEBUG, encoding='utf-8', datefmt='%d-%m-%Y %H:%M:%S', format="%(asctime)s - %(levelname)s - %(message)s")

# ? ==========================================================================================

MERGE_TABLES = True
MEASUREMENT_MEDL = list(set([380, 30, 150, 20, 590, 113]))
MEASUREMENT_MEDE = list(set([10, 640, 130, 310, 81, 380]))

# ? ==========================================================================================

def __run_update_background(updateObject: object, referenceInfo: object, arrParam: list, qtdSessions: int, maxConcurrentSpools: int):

    try:
        start = time.time()
        utils.print_start_block(f'Starting background query {utils.CustomMessage.prYellow(updateObject(referenceInfo).name)} [{len(arrParam)}] [{referenceInfo.name}]')

        background.remove_trash()
        __run_update_into_file(updateObject, referenceInfo, arrParam, qtdSessions, maxConcurrentSpools)

        background.export_files()
        __run_update_from_file(updateObject, referenceInfo)

        utils.print_end_block(f'Executed data update in {time.strftime("%H:%M:%S", time.gmtime(time.time()-start))}')

    except Exception:
        raise Exception('Exception occurred')


def __run_update_into_file(updateObject: object, referenceInfo: object, arrParam: list, qtdSessions: int, maxConcurrentSpools: int):

    try:
        maxSesSpool = math.floor(maxConcurrentSpools / qtdSessions)
        splitedArrParam = numpy.array_split(arrParam, qtdSessions)
        if len(splitedArrParam) > 6:
            raise Exception('Array parameters with length greater than 6 [sessions].')

    # ---------------------------

        arrTaskConfig = []
        for index, arr in enumerate(splitedArrParam):
            u = updateObject(referenceInfo)
            m = math.ceil(len(arr) / maxSesSpool)
            t = TaskConfig(u.execute, [arr, index + 1, m])
            arrTaskConfig.append(t)

        multitask.run_multiprocess(arrTaskConfig)

    except Exception:
        raise Exception('Exception occurred')


def __run_update_from_file(updateObject: object, referenceInfo: object):

    try:
        u = updateObject(referenceInfo)
        u.import_file_data()
        u.export_file_data()

    except Exception:
        raise Exception('Exception occurred')


def __get_reference_info(referenceName: str):
    """
    Este método realiza a conversão de um período referência para um objeto ReferenceInfo, contendo informações do período. Os dados de referência são essenciais para a consulta correta no SAP e atualização de dados na banco Access.

    Args:
        referenceName (str): nome do período de referência, no formato MM_AAAA [Exemplo: 09_2023]

    Returns:
        ReferenceInfo: objeto com dados do período
    """

    try:
        year, month = list(map(lambda x: int(x), referenceName.split("_")))
        dateIni = datetime.date(year, month, 1)
        dateEnd = datetime.date(year, month, calendar.monthrange(year, month)[1])

        return ReferenceInfo(dateIni.strftime("%Y_%m"), month, year, dateIni, dateEnd)

    except Exception:
        return None


def __merge_table_data(transactionName: str | None):

    if not MERGE_TABLES:
        return

    start = time.time()
    utils.print_start_block(f'Starting table data merge for {transactionName}')

    names = {
        'IW67': 'TB_ECC_MEDIDA',
        'IW39': 'TB_ECC_ORDEM',
    }

    if transactionName is not None:
        transactionName = transactionName.upper()

    for partialFileName, outputFileName in names.items():
        if transactionName == partialFileName or transactionName is None:
            print(utils.CustomMessage.prYellow(f'Exporting data from *{partialFileName}* to file {outputFileName}'))
            entries = utils.get_file_entries(env.DIR_EXPORTED_DATA, 'txt', [partialFileName])
            fileCrud.merge_text_file_data(entries, outputFileName)

    utils.print_end_block(f'Finished table data merge in {time.strftime("%H:%M:%S", time.gmtime(time.time()-start))}')


# ? ===================

def __get_measurement_medl():

    arrMeasurementMEDL = list(map(lambda x: f'{x:04d}', MEASUREMENT_MEDL))

    return list(set(arrMeasurementMEDL))


def __get_measurement_mede():

    arrMeasurementMEDE = list(map(lambda x: f'{x:04d}', MEASUREMENT_MEDE))

    return list(set(arrMeasurementMEDE))


# ? ==========================================================================================

def run_update_IW67_MEDL(referenceInfo: object, qtdSessions: int, maxConcurrentSpools: int):

    arrMeasurementMEDL = __get_measurement_medl()

    __run_update_background(IW67ByMeasurementMEDL, referenceInfo, arrMeasurementMEDL, qtdSessions, maxConcurrentSpools)
    __merge_table_data('IW67')


def run_update_IW67_MEDE(referenceInfo: object, qtdSessions: int, maxConcurrentSpools: int):

    arrMeasurementMEDE = __get_measurement_mede()

    __run_update_background(IW67ByMeasurementMEDE, referenceInfo, arrMeasurementMEDE, qtdSessions, maxConcurrentSpools)
    __merge_table_data('IW67')


# ! ----------------------------------------------------------------------------------------------------

if __name__ == '__main__':

    start = time.time()
    referenceName = datetime.date.today().strftime("%Y_%m")
    referenceInfo = __get_reference_info(referenceName)
    qtdSessions = 6  # Max qtd sessions = 6
    maxConcurrentSpools = 48  # Max visible spool rows [Monitor = 48 | Notebook = 27]

    utils.print_start_block(f'Starting global update for {referenceName} [{qtdSessions} sessions - max {maxConcurrentSpools} spools]')
    run_update_IW67_MEDL(referenceInfo, qtdSessions, maxConcurrentSpools)
    run_update_IW67_MEDE(referenceInfo, qtdSessions, maxConcurrentSpools)    
    utils.print_end_block(f'Finished global update in {time.strftime("%H:%M:%S", time.gmtime(time.time()-start))}')
