import os, ctypes, sys,subprocess,platform
from PyQt6.QtWidgets import *
from PyQt6.uic import loadUi
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        
        return False

class TrellixMenu(QMainWindow):
    current_dir = os.getcwd()
    def __init__(self):
        
        super().__init__()
        self.ui = loadUi('Ventana.ui', self)
        self.ui.progressBar.setValue(0)
        self.show()
        self.ValExistenciaKaspersky()
        self.InstalacionTrellix()
        self.ValidarTrellix()
        
    def ActualizacionSistema(self):
        if '10' in platform.version():
        
            os.system('wuauclt /detectnow /updatenow')
            os.system('shutdown -f -r -t 0')
        if '11' in platform.version():
            os.system("UsoClient ScanInstallWait")
            os.system('shutdown -f -r -t 0')
    ContadorValidaciones=0
    def ValidarTrellix(self):
        while self.ContadorValidaciones<2:
            
            self.ui.Lbl_N_Estado.setText("Fase 4")
            self.ui.Lbl_Estado.setText("Validando consolidacion de Trellix...")
            agente=self.get_product_id('Trellix Agent')
            Trellix=self.get_product_id('Trellix Data Exchange Layer for TA')
            PrevencionAMenazas=self.get_product_id('Protección adaptable frente a amenazas de Trellix Endpoint Security')
            WebControl=self.get_product_id('Control web de Trellix Endpoint Security')
            self.ui.progressBar.setValue(85)
            if agente is None:
                self.ui.Lbl_Estado.setText("Reinstalando Trellix Agent...")
                self.BarraProgreso(5)
                os.chdir(os.path.join(self.current_dir, '1'))
                self.execute_command('McAfeeSmartInstall.exe -s -f')
            if Trellix is None:
                self.ui.Lbl_Estado.setText("Reinstalando Trellix...")
                self.BarraProgreso(5)
                os.chdir(os.path.join(self.current_dir, '2'))
                self.execute_command('start /realtime setupEP.exe ADDLOCAL="tp,wc" /qb! >nul 2>&1')
            if PrevencionAMenazas is None and WebControl is None:
                self.ui.Lbl_Estado.setText("Reinstalando los paquetes Control Web\n y Prevención de amenazas...")
                self.BarraProgreso(5)
                os.chdir(os.path.join(self.current_dir, '2'))
                self.execute_command('start /realtime setupEP.exe ADDLOCAL="tp,wc" /qb! >nul 2>&1')
            elif PrevencionAMenazas is None and WebControl is not None:
                self.ui.Lbl_Estado.setText("Reinstalando Paquete Prevencion de amenzas...")
                os.chdir(os.path.join(self.current_dir, '2'))
                self.execute_command('start /realtime setupEP.exe ADDLOCAL="tp" /qb! >nul 2>&1')
            elif PrevencionAMenazas is not None and WebControl is None:
                self.ui.Lbl_Estado.setText("Reinstalando Paquete Control Web...")
                os.chdir(os.path.join(self.current_dir, '2'))
                self.execute_command('start /realtime setupEP.exe ADDLOCAL="wc" /qb! >nul 2>&1')
                self.ContadorValidaciones+=1
        else:
            if agente is None or Trellix is None or PrevencionAMenazas is None or WebControl is None:
        
                ctypes.windll.user32.MessageBoxW(0, "Persiste una problema al instalar el antivirus y/o sus paquetes; se forzaran las actualizaciones pendientes y se reiniciara el equipo. ", "Andrés >    Advertencia", 0x30)
                ctypes.windll.user32.MessageBoxW(0, "Despues de reiniciar el equipo vuelva a ejecutar Trellix.exe ", "Andrés >    Advertencia", 0x30)
                self.ActualizacionSistema()
                
            else:
                self.ui.progressBar.setValue(100)
                ctypes.windll.user32.MessageBoxW(0, "Proceso de cambio de solución End-Point Terminado con exito, recuerde aplicar la actualización y las 4 primeras opciones del Trellix Monitor", "Andrés >    Información",0)
                sys.exit(0)
        
        
    def InstalacionTrellix(self):
        current_dir = os.getcwd()
        self.ui.Lbl_N_Estado.setText("Fase 3")
        self.ui.Lbl_Estado.setText("Instalando Agente Trellix McAfee...")
        
        os.chdir(os.path.join(current_dir, '1'))
        self.BarraProgreso(10)
        self.execute_command('McAfeeSmartInstall.exe -s -f')
        self.BarraProgreso(10)
        self.ui.Lbl_Estado.setText("Instalando Solución End-Point Trellix McAfee...")
        os.chdir(os.path.join(current_dir, '2'))
        self.BarraProgreso(10)
        self.execute_command('start /realtime setupEP.exe ADDLOCAL="tp,wc" /qb! >nul 2>&1')
        self.BarraProgreso(10)
        os.chdir(current_dir)
        
        name = subprocess.check_output('wmic computersystem get name /value', shell=True)
        nombre = name.decode('utf-8').split('=')[1].strip()
        with open('nombres.txt', 'a') as f:
            f.write(nombre + '\n')
        
        
        
        
    valorset=0   
    def BarraProgreso(self,valor):
        self.valorset+=valor
        self.ui.progressBar.setValue(self.valorset)    
        QApplication.processEvents()
    def get_product_id(self,product_name):
        
        try:
            salida = subprocess.check_output(f'wmic product where "Name like \'%{product_name}%\'" get IdentifyingNumber /value', shell=True)
            
            self.BarraProgreso(10)
            
            return salida.decode('utf-8').split('=')[1].strip()
            
        
        except:
            return None
    def execute_command(self,command):
        
        try:
            subprocess.check_output(command, shell=True)
            
            self.BarraProgreso(10)
            
        except Exception as e:
            print(f"Error ejecutando comando: {e}")
    def ValExistenciaKaspersky(self):
       
        
        
        self.setFixedSize(798, 279)
        self.ui.Lbl_N_Estado.setText("Fase 1")
        self.ui.Lbl_Estado.setText("Validando Existencia de Kaspersky...")
        
        uid = self.get_product_id('Kaspersky Endpoint Security')
        
        uidd = self.get_product_id('Agente de red de Kaspersky Security Center')
        
        
        

        if uid is None:
            uid = self.get_product_id('Kaspersky Endpoint Security para Windows')
            

        if uidd is None:
            uidd = self.get_product_id('Kaspersky Security Center Network Agent')
            

        if uid is not None:
            self.ui.Lbl_N_Estado.setText("Fase 2")
            self.ui.Lbl_Estado.setText("Desinstalando Kaspersky...")
            self.execute_command(f'msiexec.exe /x {uid} KLLOGIN=administrador KLPASSWD=Camara2022** /qn')
            
            self.execute_command(f'msiexec.exe /x {uid} KLLOGIN=Admin KLPASSWD=Camara2022** /qn')
            

        if uidd is not None:
            
            self.ui.Lbl_Estado.setText("Desinstalando Agende de red Kaspersky...")
            self.execute_command(f'msiexec.exe /x {uidd} /qn')
         

    

if __name__ == '__main__':
    if not is_admin():
        ctypes.windll.user32.MessageBoxW(0,"Tenga en cuenta que este programa debe ejecutarse desde un usuario administrador.", "Andrés >    Advertencia",0x30)
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        
        sys.exit(0)
    app = QApplication([])
    main_window = TrellixMenu()
    
    app.exec()
