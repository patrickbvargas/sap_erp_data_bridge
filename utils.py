import os
import pyperclip
import logging
import datetime


class CustomMessage:

    @staticmethod
    def prRed(message: str): return f"\033[91m{message}\033[00m" if os.environ.get(
        'TERM_PROGRAM') == 'vscode' else message

    @staticmethod
    def prGreen(message: str): return f"\033[92m{message}\033[00m" if os.environ.get(
        'TERM_PROGRAM') == 'vscode' else message

    @staticmethod
    def prYellow(message: str): return f"\033[93m{message}\033[m" if os.environ.get(
        'TERM_PROGRAM') == 'vscode' else message

    @staticmethod
    def prLightPurple(message: str): return f"\033[94m{message}\033[00m" if os.environ.get(
        'TERM_PROGRAM') == 'vscode' else message

    @staticmethod
    def prPurple(message: str): return f"\033[95m{message}\033[00m" if os.environ.get(
        'TERM_PROGRAM') == 'vscode' else message

    @staticmethod
    def prCyan(message: str): return f"\033[96m{message}\033[00m" if os.environ.get(
        'TERM_PROGRAM') == 'vscode' else message

    @staticmethod
    def prLightGray(message: str): return f"\033[97m{message}\033[00m" if os.environ.get(
        'TERM_PROGRAM') == 'vscode' else message

    @staticmethod
    def prBlack(message: str): return f"\033[98m{message}\033[00m" if os.environ.get(
        'TERM_PROGRAM') == 'vscode' else message

    @staticmethod
    def prInfo(message: str): return f"\033[2;37;43m{message}\033[00m" if os.environ.get(
        'TERM_PROGRAM') == 'vscode' else message

# ? ==========================================================================================

def get_referenceName_year(referenceInfo: object):

    return [f'{referenceInfo.year}_{month+1:02d}' for month in range(12)]


def split_array(arrData: list, size: int):
    """  
    Este método realiza a divisão de um array em N partes de tamanho X (size)

    Args:
        arrData (list): array de dados a serem divididos
        size (int): tamanho máximo de cada parte dividida

    Returns:
        list: array dividido
    """

    try:
        arrRes = []
        for i in range(0, len(arrData), size):
            arrRes.append(arrData[i:i+size])
        return arrRes
    except Exception:
        logging.exception('Exception occurred')
        return []


def copy_to_clipboard(arrParam: list):
    """  
    Este método realiza a cópia para a área de transferência de um conjunto de dados, em formato de tabela (com quebra de linha)
    ATENÇÃO: o array de dados deve ser unidimensional (não usar matriz) e utilizar dados primitivos (int, str, double...)

    Args:
        arrParam (list): array de dados para cópia
    """

    try:
        text = '\r\n'.join(list(map(lambda x: str(x), arrParam)))
        while (True):
            pyperclip.copy(text)
            if (text == pyperclip.paste()):
                break

    except Exception:
        logging.exception('Exception occurred')


def print_start_block(message: str):
    """  
    Este método realiza a escrita no terminal de um padrão estruturado para indicar o início de execução de um determinado bloco de código

    Args:
        message (str): mensagem a ser incluída durante a escrita
    """

    maxCaracteres = 80
    caractere = '─'
    print(caractere*maxCaracteres)
    print(CustomMessage.prLightGray(f'⋙  {message}'))


def print_end_block(message: str):
    """  
    Este método realiza a escrita no terminal de um padrão estruturado para indicar o fim de execução de um determinado bloco de código

    Args:
        message (str): mensagem a ser incluída durante a escrita
    """
    maxCaracteres = 80
    caractere = '─'
    print(CustomMessage.prLightGray(f'⋘  {message}'))
    print(caractere*maxCaracteres)


def print_progress_bar(message: str, maxNumCaracteres: int, index: int, length: int):
    """  
    Este método realiza a escrita no terminal de uma barra de progresso de uma determinada tarefa

    Args:
        message (str): mensagem que precede a barra de progresso
        maxNumCaracteres (int): tamanho da barra de progresso em caracteres
        index (int): índice/posição atual do conjunto de dados
        length (int): tamanho máximo do conjunto de dados
    """

    try:
        percent = index / length
        caractere = "▰ "
        barLength = int(percent * maxNumCaracteres)
        progressBar = f'{(caractere * barLength).ljust(len(caractere) * maxNumCaracteres)}'
        progressPercent = f'{index}/{length}'.ljust(12)
        progressMessage = f'{message}: {CustomMessage.prGreen(progressBar)} {progressPercent} [{percent:00.1%} complete]'
        endChar = '\n' if percent == 1 else '\r'
        print(progressMessage, end=endChar)

    except Exception:
        print(f'{CustomMessage.prYellow("Failed")} to print progress bar.')


def get_file_entries(directoryPath: str, fileExtension: str, partialTexts: list):

    entries = [entry for entry in os.scandir(directoryPath) if entry.is_file() and fileExtension.lower() in entry.name.lower() and any(
        partialText.lower() in entry.name.lower() for partialText in partialTexts)]
    return entries

# ? ==========================================================================================


def __get_fixed_string(value: str):
    """ 
    Este método realiza o ajuste da string exportada SAP que contenha um "X" predecendo uma informação no arquivo .txt exportado, decorrente de uma incidência única no arquivo

    Args:
        value (str): valor original

    Returns:
        str: valor ajustado
    """

    replacements = {
        "X": "",
    }

    for old, new in replacements.items():
        value = value.upper().replace(old, new)

    return value.strip()


def format_float(value: str):
    """ 
    Este método realiza o ajuste de valor importado do SAP para padrões compreensíveis
    ATENÇÃO: o objetivo é corrigir falhas de formatação, não realizar casting de tipo de dado

    Args:
        value (str): valor original

    Returns:
        str: valor formatado
    """

    try:
        if value == None or value == '':
            return ''
        fixedDigit = __get_fixed_string(value).replace(".", "").replace(",", ".")
        fixedSignal = f'-{fixedDigit.replace("-", "")}' if fixedDigit.find("-") != -1 else fixedDigit
        formattedValue = str(float(fixedSignal)).replace(".", ",")
        return formattedValue

    except Exception:
        return None


def format_integer(value: str):
    """ 
    Este método realiza o ajuste de valor importado do SAP para padrões compreensíveis
    ATENÇÃO: o objetivo é corrigir falhas de formatação, não realizar casting de tipo de dado

    Args:
        value (str): valor original

    Returns:
        int: valor formatado
    """

    try:
        if value == None or value == '':
            return ''
        formattedValue = int(__get_fixed_string(value))
        return str(formattedValue)

    except Exception:
        return None


def format_date(value: str):
    """ 
    Este método realiza o ajuste de valor importado do SAP para padrões compreensíveis
    ATENÇÃO: o objetivo é corrigir falhas de formatação, não realizar casting de tipo de dado

    Args:
        value (str): valor original

    Returns:
        str: valor formatado
    """

    try:
        if value == None or value == '':
            return ''
        day, month, year = value.split(".")
        datetime.datetime(int(year), int(month), int(day))
        return f'{day}/{month}/{year}'

    except Exception:
        return None


def format_string(value: str):
    """ 
    Este método realiza o ajuste de valor importado do SAP para padrões compreensíveis
    ATENÇÃO: o objetivo é corrigir falhas de formatação, não realizar casting de tipo de dado

    Args:
        value (str): valor original

    Returns:
        str: valor formatado
    """

    try:
        if value == None or value == '':
            return ''
        formattedValue = value.strip()
        return formattedValue

    except Exception:
        return None


def format_currency(value: str):
    """ 
    Este método realiza o ajuste de valor importado do SAP para padrões compreensíveis
    ATENÇÃO: o objetivo é corrigir falhas de formatação, não realizar casting de tipo de dado

    Args:
        value (str): valor original
    """

    try:
        return format_float(value)

    except Exception:
        return None
