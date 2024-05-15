from model import SapImportConfig, FieldConfig, UpdateData, ReferenceInfo
import utils

# ? ==========================================================================================

IW67Config = SapImportConfig(
    'IW67',
    '/SAP_DATA_BRIDGE',
    [
        FieldConfig('NOTA', ['Nota'], utils.format_integer),
        FieldConfig('MEDIDA', ['CóMd'], utils.format_integer),
        FieldConfig('STATUS', ['StatSist'], utils.format_string),
        FieldConfig('RESPONSAVEL', ['Exec.por'], utils.format_string),
        FieldConfig('TEXTO', ['Texto das medidas', 'TextoMedid', 'Texto medidas'], utils.format_string),
        FieldConfig('LOCALIZACAO', ['Localiz.'], utils.format_integer),
        FieldConfig('USUARIO_CRIACAO', ['Criado/a'], utils.format_string),
        FieldConfig('DATA_CRIACAO', ['Dt.criação'], utils.format_date),
        FieldConfig('USUARIO_CONCLUSAO', ['por', 'Concl.por'], utils.format_string),
        FieldConfig('DATA_CONCLUSAO', ['Concluído'], utils.format_date),
        FieldConfig('DATA_PLANEJAMENTO_INICIO', ['Iníc.planj'], utils.format_date),
        FieldConfig('DATA_PLANEJAMENTO_FIM', ['Fim plan.'], utils.format_date),
        FieldConfig('INDICE', ['Medi'], utils.format_integer),
        FieldConfig('EQUIPAMENTO', ['LocInstal.'], utils.format_string)
    ]
)


# ? ==========================================================================================

class IW67ByMeasurementMEDE(UpdateData):

    def __init__(self, referenceInfo: ReferenceInfo):
        super().__init__(f'IW67_MEDE_{referenceInfo.name}', IW67Config, referenceInfo)

    def _initialize_sap_transaction(self, session: object, arrParam: list):

        while True:
            session.findById("wnd[0]/usr/chkDY_QMSM").selected = False
            session.findById("wnd[0]/usr/btn%_MNCOD_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[16]").press()
            utils.copy_to_clipboard(arrParam)
            session.findById("wnd[1]/tbar[0]/btn[24]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/usr/ctxtERLDAT-LOW").text = self.referenceInfo.dateIni.strftime("%d.%m.%Y")
            session.findById("wnd[0]/usr/ctxtERLDAT-HIGH").text = self.referenceInfo.dateEnd.strftime("%d.%m.%Y")
            if session.findById("wnd[0]/usr/ctxtMNCOD-LOW").text == str(arrParam[0]):
                break


class IW67ByMeasurementMEDL(UpdateData):

    def __init__(self, referenceInfo: ReferenceInfo):
        super().__init__('IW67_MEDL', IW67Config, referenceInfo)

    def _initialize_sap_transaction(self, session: object, arrParam: list):

        while True:
            session.findById("wnd[0]/usr/chkDY_QMSM").selected = True
            session.findById("wnd[0]/usr/btn%_MNCOD_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[16]").press()
            utils.copy_to_clipboard(arrParam)
            session.findById("wnd[1]/tbar[0]/btn[24]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = '01.01.2021'
            session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text = '31.12.9999'
            if session.findById("wnd[0]/usr/ctxtMNCOD-LOW").text == str(arrParam[0]):
                break

