import pyodbc
import uuid
import datetime
import json

"""
    Python Automation Tools
    Versão: 1.0
    Data: 02/01/2019
    Autor: Reginaldo Luciano de Araujo
    
    Implementações:
    
    Constantes de conexão:
    strSanofiReportsConnectionString -> Sanofi_Reports
    strProjetoBPOConnectionString -> Projeto_BPO

    Funções:
    eventos_automacao -> Registra erros ocorridos em scripts de automação Python
    GetSheetLoaderFile ->

"""

strSanofiReportsConnectionString = 'Driver={SQL Server};Server=SOALV3BPOAUT01\\SQLEXPRESS;Database=Sanofi_Reports;UID=sa;PWD=Bp0stf19'
strProjetoBPOConnectionString = 'Driver={SQL Server};Server=SOALV3BPOAUT01\\SQLEXPRESS;Database=Projeto_BPO;UID=sa;PWD=Bp0stf19'

def eventos_automacao(id_batch, id_evento, dt_evento, st_fonte, vl_tipo_evento, st_evento, vl_execucao):

    """
        Stored Procedure: 
        SP_EVENTOS_AUTOMACAO
        Função: Registrar erros de automação.
        
        Parâmetros:

        @ID_BATCH			
        Tipo: UNIQUEIDENTIFIER
        Função: 

        @ID_EVENTO			
        Tipo: UNIQUEIDENTIFIER
        Função: 

        @DT_EVENTO			
        Tipo: DATETIME
        Função: 
        
        @ST_FONTE			
        Tipo: VARCHAR(100)
        Função: 

        @VL_TIPO_EVENTO	
        Tipo: NUMERIC(1,0)
        Função: 

        @ST_EVENTO			
        Tipo: TEXT
        Função: 

        @VL_EXECUCAO		Tipo>NUMERIC(1,0)

    """

    valores = (id_batch, id_evento, dt_evento, st_fonte, vl_tipo_evento, st_evento, vl_execucao)
    procedure = ('EXEC SP_EVENTOS_AUTOMACAO '
                '@ID_BATCH          = ?,' + 
                '@ID_EVENTO         = ?,' + 
                '@DT_EVENTO         = ?,' + 
                '@ST_FONTE          = ?,' + 
                '@VL_TIPO_EVENTO    = ?,' + 
                '@ST_EVENTO	        = ?,' + 
                '@VL_EXECUCAO       = ?' )

    conn = pyodbc.connect(strProjetoBPOConnectionString)
    cursor = conn.cursor()
    cursor.execute(procedure,(valores)) #Executa a procedure passando parâmetros
    conn.commit()
    conn.close()

def CreateGUID():
    objGUID = uuid.uuid4()
    return objGUID

def NowTime():
    return datetime.datetime.now().time()

def NowDate():
    return datetime.datetime.date()

def NowDateTime():
    return datetime.datetime.now()

def GetSheetLoaderFile(SheetLoaderHeader):
    with open("SheetLoader.json", "r",  encoding="utf8") as SheetLoader:
        data = json.load(SheetLoader)
        for SheetLoader in data['RegisteredLoaders']:
            if (SheetLoader.get('Loader') == SheetLoaderHeader):
                return SheetLoader.get('File')
