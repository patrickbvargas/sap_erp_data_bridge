import sap
import fileCrud
import logging
import time
import env
import utils
import datetime

# ? Informações: módulo responsável pela gestão das Classes em uso no script (contém as regras de negócio principais)


class ReferenceInfo:

    """
    Esta classe representa a parametrização de perído de execução da atualização

    A classe ReferenceInfo faz o seguinte:
        - Estrutura os dados necessários para executar os relatórios de acordo com as datas necessárias
        - Não possui métodos/funções próprias
    """

    def __init__(self, name: str, month: int, year: int, dateIni: any, dateEnd: any):
        self.name = name
        self.month = month
        self.year = year
        self.dateIni = dateIni
        self.dateEnd = dateEnd


class TaskConfig:

    """
    Esta classe representa a parametrização de tarefas a serem executadas em multithreading ou multiprocessing

    A classe TaskConfig faz o seguinte:
        - Estrutura os dados necessários para executar um método/função em multithread/multiprocess
        - Não possui métodos/funções próprias
    """

    def __init__(self, callback: any, args: list):
        """
        Este é o método construtor da classe TaskConfig.

        Args:
            callback (any): referência para o método/função do objeto reference que será executado (ponteiro)
            args (list): array com os argumentos necessários para executar o método/função (se não houver argumentos, informar array vazio [])
        """

        self.callback = callback
        self.args = args


class FilterConfig:

    def __init__(self, columnName: str, type: str, args: list):

        self.columnName = columnName
        self.type = type
        self.args = args

    def isValid(self, value: str) -> bool:

        filterFunctions = {
            'equalsAny': lambda x: x in self.args,
            'equalsAll': lambda x: all(arg == x for arg in self.args),
            'containsAll': lambda x: all(arg in x for arg in self.args),
            'containsAny': lambda x: any(arg in x for arg in self.args),
        }

        filterFunction = filterFunctions.get(self.type)
        if filterFunction:
            return filterFunction(value)
        else:
            print('Invalid filter type')
            return False


class FieldConfig:

    """
    Esta classe representa a parametrização de colunas que representam a estrutura dos dados de uma transação SAP, utilizados para importar/exportar corretamente os dados entre diversos contexto. Cada transação do SAP possui colunas específicas de acordo com o layout definido, os parâmetros dos nomes de coluna devem ser estruturados utilizando esta classe

    A classe Field faz o seguinte:
        - Estrutura os dados necessários para identificar as colunas de uma transação em diversos contextos
        - Não possui métodos/funções próprias
    """

    def __init__(self, name: str, fileColumnNames: list, getValueCallback: any):
        """
        Este é o método construtor da classe Field.

        Args:
            name (str): nome da campo            
            sapColumnName (str): nome da coluna no contexto de tabela SAP (GuiGridView)
            fileColumnNames (list): array de nomes da coluna no contexto de arquivo .txt (exportado background)
            getValueCallback (any): função para tratar o dado original e retornar formatado 
        """

        self.name = name
        self.fileColumnNames = fileColumnNames
        self.getValue = getValueCallback


class SapImportConfig:

    """
    Esta classe representa a parametrização da transação SAP. Cada transação possui parâmetros específicos que devem ser estruturados utilizando esta classe.
    ATENÇÃO: transações SAP que resultam em tabelas que não sejam instâncias de GuiGridView (ver documentação SAP) não são aplicáveis

    A classe SapImportConfig faz o seguinte:
        - Estrutura os dados necessários para executar uma transação no SAP
        - Não possui métodos/funções próprias
    """

    def __init__(self, sapTransaction: str, sapVariant: str, fields: list):
        """
        Este é o método construtor da classe SapImportConfig.

        Args:
            sapTransaction (str): código da transação do SAP
            sapVariant (str): nome da variante de execução do SAP         
            fields (list): array contendo a parametrização das colunas das tabelas de dados da transação 
        """

        self.sapTransaction = sapTransaction
        self.sapVariant = sapVariant
        self.fields = fields


class UpdateData(SapImportConfig):

    """
    Esta classe representa a parametrização completa para atualização dos dados a partir de transações SAP. Cada parametrização de consulta devem ser estruturada utilizando esta classe.

    A classe UpdateData faz o seguinte:
        - Estrutura os dados necessários para atualizar dados via transação SAP
    """

    def __init__(self, name: str, sapConfig: SapImportConfig, referenceInfo: ReferenceInfo):
        """
        Este é o método construtor da classe UpdateData.

        Args:
            name (str): nome que descreve a atualização de dados [Exemplo: IW67_MEDL]
            sapConfig (SapImportConfig): objeto que contém os parâmetros da transação SAP            
            referenceInfo (ReferenceInfo): objeto que representa os dados do período para execução
        """

        super().__init__(sapConfig.sapTransaction, sapConfig.sapVariant, sapConfig.fields)
        self.name = name
        self.referenceInfo = referenceInfo
        self.data = []
        self.printSapLog = False

    def _initialize_sap_transaction(self):
        # ! must be overridden by inherited class
        """
        Este método inicia a transação no SAP e parametriza os campos para realizar a consulta. É um método abstrato, cada instância da classe deve implementar sua lógica de consulta no SAP.

        Args:
            session (object): objeto de conexão com o SAP Script
            arrParam (list): lista com os parâmetros para consulta (ordens, notas, materiais ...)
        """

        pass

    def __consult_sap_data(self, sessionNumber: int, arrParam: list):
        """     
        Este método realiza a consulta de dados no SAP conforme os parâmetros informados, na tela SAP indicada
        IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de númeto de tentativas)

        Args:
            sessionNumber (int): número da sessão SAP (tela) para criação da conexão
            arrParam (list): lista com os parâmetros para consulta (ordens, notas, materiais ...)          

        Returns:
            bool: True se executado com sucesso
        """

        while True:
            try:
                session = sap.get_session_by_number('ERP', sessionNumber)
                if self.printSapLog:
                    print(f'Querying data from {self.name} in {session.name} [{len(arrParam)}]')

                start = time.time()
                session.StartTransaction(self.sapTransaction)
                if self.name != 'ZIM151':
                    sap.set_user_variant(session, self.sapVariant)
                self._initialize_sap_transaction(session, arrParam)

                if (sap.create_background_job(session, self.name)):
                    infoText = f'{utils.CustomMessage.prGreen("Successfully")} created background job for {self.name} in {session.name} [{time.strftime("%H:%M:%S", time.gmtime(time.time()-start))}]'
                else:
                    raise Exception

                return True

            except Exception:
                infoText = f'{utils.CustomMessage.prRed("Failed")} to query data from {self.name} in {session.name}'
                logging.exception('Exception occurred')
                continue

            finally:
                if self.printSapLog:
                    print(infoText)

    def import_file_data(self):
        """      
        Este método realiza a importação dos dados a partir de arquivos .txt resultantes da execução em background (spool). 
        ATENÇÃO: arquivos .txt devem estar exportados no diretório [env.DIR_SPOOL_DATA] e no nome do arquivo deve conter (em qualquer posição) o valor da variável "name".   
        IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de númeto de tentativas)

        Returns:
            bool: True se executado com sucesso
        """

        while True:
            try:
                start = time.time()
                print(f'Reading data from {self.name} exported files')

                nowDatetime = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                entries = utils.get_file_entries(env.DIR_SPOOL_DATA, 'txt', [self.name])
                fileData = fileCrud.import_spool_file_data(entries, self.fields)

                for data in fileData:
                    data['REFERENCIA'] = self.referenceInfo.name
                    data['DATA_HORA_CONSULTA'] = nowDatetime

                self.data.extend(fileData)

                infoText = f'{utils.CustomMessage.prGreen("Successfully")} data imported from {self.name} [{len(self.data)} rows] [{time.strftime("%H:%M:%S", time.gmtime(time.time()-start))}]'
                return True

            except Exception:
                infoText = f'{utils.CustomMessage.prRed("Failed")} to import data from {self.name} in exported file.'
                logging.exception('Exception occurred')
                continue

            finally:
                print(infoText)

    def export_file_data(self):
        """      
        Este método realiza a exportação dos dados para um arquivo .json        
        IMPORTANTE: será executado repetidamente até que haja sucesso (loop infinito - sem delimitação de númeto de tentativas)
        """

        while True:
            try:
                start = time.time()
                if not self.data:
                    infoText = f'{utils.CustomMessage.prYellow("No data")} to export from {self.name}'
                    return True

                fileCrud.export_text_file_data(env.DIR_EXPORTED_DATA, self.name, self.data)
                # fileCrud.export_json_file_data(env.DIR_EXPORTED_DATA, self.name, self.data)

                infoText = f'{utils.CustomMessage.prGreen("Successfully")} exported data from {self.name} to file [{len(self.data)} rows] [{time.strftime("%H:%M:%S", time.gmtime(time.time()-start))}]'
                return True

            except Exception:
                infoText = (f'{utils.CustomMessage.prRed("Failed")} to export data from {self.name} to file [{len(self.data)} rows]')
                logging.exception('Exception occurred')
                continue

            finally:
                print(infoText)

    def execute(self, arrParam: list, sessionNumber: int, maxConcurrentData: int):
        """ 
        Este método dá início a atualização completa dos dados, a partir de métodos específicos.
        """

        try:
            delay = 0 if sessionNumber == 1 else sessionNumber
            time.sleep(delay)  # Aguardar X segundos para minimizar concorrência no uso do clipboard
            for arr in utils.split_array(arrParam, maxConcurrentData):
                while not (self.__consult_sap_data(sessionNumber, arr)):
                    continue

            return True

        except Exception:
            logging.exception('Exception occurred')
            return False
