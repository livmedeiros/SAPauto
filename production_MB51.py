import win32com.client
import subprocess
import time
import datetime
from datetime import timedelta

# Manipulação de datas
dataHoje = datetime.datetime.now()
diaSemana = dataHoje.weekday()   # Cada dia da semana é representado por um número sendo o intervalo: "0 == Segunda"; "6 == Domingo"


# Informações necessárias para realizar o login no SAP
SAP_SID = 'codigo-SID'
SAP_MANDANT = 'numero_mandante_ou_cliente'
SAP_USER = 'seu_usuario'
SAP_PASS = 'sua_senha'
SAP_EXE = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe"


try:

    def sap_process():
        # Realizando login na interface
        subprocess.check_call([SAP_EXE,
                                f'-user={SAP_USER}',
                                f'-pw={SAP_PASS}',
                                f'-system={SAP_SID}',
                                f'-client={SAP_MANDANT}',
                                '-language=PT'])

        time.sleep(10)

        SapGuiAuto = win32com.client.GetObject("SAPGui")
        if not isinstance(SapGuiAuto, win32com.client.CDispatch):
            raise ValueError("Não foi possível obter o objeto SAPGui")

        application = SapGuiAuto.GetScriptingEngine
        if not isinstance(application, win32com.client.CDispatch):
            raise ValueError("Não foi possível obter o objeto ScriptingEngine")

        connection = application.Children(0)
        if not isinstance(connection, win32com.client.CDispatch):
            raise ValueError("Não foi possível estabelecer a conexão SAP")

        session = connection.Children(0)
        if not isinstance(session, win32com.client.CDispatch):
            raise ValueError("Não foi possível obter a sessão SAP")

        session.findById("wnd[0]").maximize()
        # Seleciona a transação MB51
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB51"
        session.findById("wnd[0]").sendVKey(0)
        # Movimentos utilizados para verificar o apontamento de produção
        session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "131"
        session.findById("wnd[0]/usr/ctxtBWART-HIGH").text = "132"
        session.findById("wnd[0]/usr/txtUSNAM-LOW").text = "campo_usuario1"
        session.findById("wnd[0]/usr/txtUSNAM-HIGH").text = "campo_usuario2"
        session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = f"{dataVar}"
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = f"{dataVar}"
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 10
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[48]").press()
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(10)
        # Após realizar o processo, faz o log off na aplicação
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
        session.findById("wnd[0]").sendVKey(0)


    # Em caso de segunda-feira, o programa buscará o apontamento de produção na sexta e no sábado;
    if diaSemana == 0:
        for i in range(1):
            dataVar = (dataHoje - timedelta(days=3)).strftime('%d.%m.%Y')
            sap_process()
            dataVar = (dataHoje - timedelta(days=2)).strftime('%d.%m.%Y')
            sap_process()
    # em qualquer outro dia, irá buscar o apontamento do dia anterior.
    else:
        dataVar = (dataHoje - timedelta(days=1)).strftime('%d.%m.%Y')
        sap_process()


except Exception as e:
    print(f"Erro: {e}")