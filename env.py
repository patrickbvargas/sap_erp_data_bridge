import os

# ? Informações: módulo responsável pela gestão das variáveis globais de ambiente

USERNAME = os.environ['USERNAME']

# Diretório onde os dados importados serão armazenados
DIR_EXPORTED_DATA = '.\\data'

# Diretório onde os arquivos spool do SAP são exportados por padrão
DIR_SPOOL_DATA = 'DEFINIR_DIRETORIO_SPOOL'

# Diretório onde os dados finais (tabelas) serão armazenados para uso no BI
DIR_TABLE_DATA = 'DEFINIR_DIRETORIO_TABELAS'
