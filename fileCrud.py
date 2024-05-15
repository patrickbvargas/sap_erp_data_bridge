import logging
import env
import json
import utils
import pandas as pd

# ? Informações: módulo responsável pelo tratamento dos arquivos .txt (leitura, inserção e remoção)


def __is_spool_header_line(line: str, fields: list):
    """  
    Este método realiza a conferência se os dados de uma linha corresponde ao cabeçalho da transação, conferindo se todos os itens fileColumnName de fields possui ao menos uma correspondência na linha.    

    Args:
        line (str): dados da linha do arquivo
        fields (list): array de fields do objeto SapImportConfig

    Returns:
        bool: True se corresponde ao cabeçalho
    """

    try:
        uppercaseLine = line.upper()
        return all(any(name.upper() in uppercaseLine for name in field.fileColumnNames) for field in fields)

    except Exception:
        return False


def __is_spool_data_line(inputString: str, positions: list):
    try:
        return bool(positions) and all(inputString[pos] == '|' for pos in positions)
    except:
        return False


def __get_header_column_index(header: list, fields: list):
    """  
    Este método realiza a identificação do indíce da coluna de cada item field a partir da ordem dos dados no cabeçalho

    Args:
        header (list): array com os dados do cabeçalho
        fields (list): array de fields do objeto SapImportConfig

    Returns:
        object: contendo nome e índice da coluna (index) [Exemplo: {Nota: 1}]
    """

    try:
        columnIndex = {}
        for field in fields:
            for index, name in enumerate(header):
                if name is not None and name.upper() in map(str.upper, field.fileColumnNames):
                    columnIndex[field.name] = index
                    break

        return columnIndex

    except Exception:
        return None


def __get_fixed_string(inputString: str):
    """
    Este método realiza o ajuste da string linha do arquivo

    Args:
        inputString (str): valor original

    Returns:
        str: valor ajustado
    """

    if not inputString.endswith('|'):
        inputString += '|'

    return inputString


def __get_char_positions(inputString: str, targetChar: str):
    positions = [index for index, char in enumerate(inputString) if char == targetChar]
    return positions


def __get_split_string_by_positions(inputString: str, positions: list):
    splits = [inputString[start:end] for start, end in zip([0] + positions, positions + [None])]
    fixed = [substring.replace('|', '').strip() for substring in splits]
    return fixed


# ? ==========================================================================================

def import_spool_file_data(entries: list, fields: list):
    """
    Este método realiza a importação de dados de arquivos .txt de arquivos spool. É necessário que os nomes das colunas do arquivo estejam definidas para cada transação: SapImportConfig -> Fields -> fileColumnName
    ATENÇÃO: dependendo da quantidade de dados exportados o título das colunas do arquivo exportado pode variar [Exemplo: transação IW67 coluna "Texto das Medidas" pode conter os títulos a seguir no arquivo: ['Texto das medidas', 'TextoMedid', 'Texto medidas']

    Args:
        entries (list): lista de entradas de arquivos
        fields (list): array de fields do objeto SapImportConfig

    Returns:
        list: array contendo dictionary com os dados
    """

    try:

        data = []
        for index, entry in enumerate(entries):
            utils.print_progress_bar(f'Importing data from {len(entries)} files', 20, index + 1, len(entries))

            header = None
            headerSplitPositions = []
            with open(entry.path) as file:
                rows = file.readlines()

            for row in rows:
                row = __get_fixed_string(row)

                if __is_spool_header_line(row, fields):
                    headerSplitPositions = __get_char_positions(row, '|')
                    splitData = [item.strip() or '' for item in row.split('|')]
                    header = __get_header_column_index(splitData, fields)

                elif __is_spool_data_line(row, headerSplitPositions):
                    splitData = __get_split_string_by_positions(row, headerSplitPositions)
                    tempRow = {}
                    for field in fields:
                        tempRow[field.name] = field.getValue(splitData[header[field.name]])
                    if any(tempRow.values()):  # ? Por que não funciona all(tempRow.values()) para IW67_MEDL, ZPLM0152 e MB25
                        data.append(tempRow.copy())

        return data

    except Exception:
        return None


def import_text_file_data(entries: list):
    """  
    Este método realiza a importação de dados de arquivos .txt já tratados. Utiliza a primeira linha como cabeçalho dos objetos    

    Args:
        entries (list): lista de entradas de arquivos

    Returns:
        list: array contendo dictionary com os dados
    """

    try:

        data = []
        for index, entry in enumerate(entries):
            utils.print_progress_bar(f'Importing data from {len(entries)} files', 20, index + 1, len(entries))

            header = None
            with open(entry.path, encoding='utf-8') as file:
                rows = file.readlines()

            for row in rows:
                splitedData = [item.strip() or '' for item in row.split('|')]

                if header is None:  # first row = header
                    header = splitedData
                else:
                    tempRow = {}
                    for index, value in enumerate(splitedData):
                        tempRow[header[index]] = value
                    if any(tempRow.values()):
                        data.append(tempRow.copy())

        return data

    except Exception:
        logging.exception('Exception occurred')
        return None


def import_json_file_data(entries: list):

    try:
        data = []
        for index, entry in enumerate(entries):
            utils.print_progress_bar(f'Importing data from {len(entries)} files', 20, index + 1, len(entries))

            with open(entry.path, 'r') as file:
                data = json.load(file)

        return data

    except Exception:
        return None


# ? ==========================================================================================

def export_json_file_data(dirPath: str, outputFileName: str, data: list):

    try:
        print(f'Exporting data to file {outputFileName}.json [{len(data)} rows]')
        filePathName = f'{dirPath}/{outputFileName}.json'
        with open(filePathName, 'w') as file:
            json.dump(data, file, indent=None)

        return True

    except Exception:
        return False


def export_text_file_data(dirPath: str, outputFileName: str, data: list):

    try:
        print(f'Exporting data to file {outputFileName}.txt [{len(data)} rows]')
        separator = '|'
        filePathName = f'{dirPath}/{outputFileName}.txt'
        with open(filePathName, 'w+', encoding='utf-8') as file:
            header = separator.join(data[0].keys())
            file.write(f'{header}\n')

            for row in data:
                values = separator.join(map(str, row.values()))
                file.write(f'{values}\n')

        return True

    except Exception:
        return False


# ? ==========================================================================================

def convert_json_to_text(entries: list, outputFileName: str):

    try:
        data = import_json_file_data(entries)
        export_text_file_data(env.DIR_EXPORTED_DATA, outputFileName, data)

        return True

    except Exception:
        logging.exception('Exception occurred')
        return False


def convert_text_to_json(entries: list, outputFileName: str):

    try:
        data = import_text_file_data(entries)
        export_json_file_data(outputFileName, data)

        return True

    except Exception:
        logging.exception('Exception occurred')

# ? ==========================================================================================


def merge_text_file_data(entries: list, outputFileName: str):

    try:
        data = import_text_file_data(entries)
        export_text_file_data(env.DIR_TABLE_DATA, outputFileName, data)

        return True

    except Exception:
        logging.exception('Exception occurred')


# ? ==========================================================================================

def query_data(dataColumnName: str, entriesData: list, filters: list | None):

    try:
        data = set()
        for row in entriesData:
            if filters is None or all(filter.isValid(row[filter.columnName]) for filter in filters):
                data.add(row[dataColumnName])

        return list(data)

    except Exception:
        logging.exception('Exception occurred')
        return None

# ? ==========================================================================================


def query_plans_note():

    try:
        dataSet = set()
        filePath = fr'{env.DIR_TABLE_DATA}\NOTAS_PLANOS.xlsm'

        print(f'Querying plans notes from {filePath}')
        excelData = pd.read_excel(filePath, sheet_name='NOTA_PLANO')
        dataColumnIndex = 3

        for row in excelData.values:
            value = str(row[dataColumnIndex])
            if value.isdigit():
                dataSet.add(value)

        data = list(dataSet)
        print(f'{utils.CustomMessage.prGreen("Successfully")} queried {len(data)} notes')

        return data

    except Exception:
        logging.exception('Exception occurred')
        return None


def query_emergency_events():

    try:
        dataSet = set()
        filePath = fr'{env.DIR_TABLE_DATA}\SPIR.xlsm'

        print(f'Querying emergency events from {filePath}')
        excelData = pd.read_excel(filePath, sheet_name='AM')
        dataColumnIndex = 3

        for row in excelData.values:
            value = str(row[dataColumnIndex]).replace('OS', '').strip()
            if value.isdigit():
                dataSet.add(value)

        data = list(dataSet)
        print(f'{utils.CustomMessage.prGreen("Successfully")} queried {len(data)} emergency events')

        return data

    except Exception:
        logging.exception('Exception occurred')
        return None
