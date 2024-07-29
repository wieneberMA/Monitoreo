import sys
import time
import requests
import smtplib
import os
import shutil
import pandas as pd
from email.mime.text import MIMEText
from PyQt5 import QtWidgets, QtCore, QtGui
import pythoncom
from win32com.shell import shell, shellcon
from PyQt5.QtWidgets import QFileDialog

# Configuración del servidor SMTP y credenciales de correo electrónico
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SMTP_USER = 'mailing@shopeando.mx'
SMTP_PASSWORD = 'cftu kpnc fgpk lomg'

# Configuración del destinatario
TO_EMAILS = ['wieneber76@gmail.com']
CC_EMAILS = ['programacion@shopeando.mx']
BCC_EMAILS = ['']

# Archivos de registro
LOG_FILE = './service_logs.txt'
SERVICES_FILE = './services_list.txt'

# Función para enviar el correo electrónico
def send_email(subject, body, to_emails, cc_emails=None, bcc_emails=None):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = SMTP_USER
    msg['To'] = ", ".join(to_emails)
    
    if cc_emails:
        msg['Cc'] = ", ".join(cc_emails)
    
    if bcc_emails:
        msg['Bcc'] = ", ".join(bcc_emails)

    recipients = to_emails
    if cc_emails:
        recipients += cc_emails
    if bcc_emails:
        recipients += bcc_emails

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(SMTP_USER, recipients, msg.as_string())

class Monitor(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.services = []
        self.initUI()
        self.load_services()
        self.load_logs()
        self.checkServices()  # Hacer consulta al iniciar

    def initUI(self):
        self.setWindowTitle('Servicio de Monitoreo')
        self.setWindowIcon(QtGui.QIcon('logo.jpg'))
        self.resize(900, 400)  # Ajuste del tamaño para la nueva barra

        # Crear un widget central y establecerlo
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        
        self.layout = QtWidgets.QVBoxLayout(central_widget)

        # Menú de la aplicación
        self.menuBar = self.menuBar()

        # Menú de Notificaciones
        self.notificationMenu = self.menuBar.addMenu('Notificaciones')
        self.checkManualAction = QtWidgets.QAction('Consultar Manualmente', self)
        self.checkManualAction.triggered.connect(self.checkServices)
        self.notificationMenu.addAction(self.checkManualAction)

        self.notifySlowAction = QtWidgets.QAction('Notificar Servicios Lentos', self)
        self.notifySlowAction.triggered.connect(self.notifySlowServices)
        self.notificationMenu.addAction(self.notifySlowAction)

        self.notifyErrorsAction = QtWidgets.QAction('Notificar Errores', self)
        self.notifyErrorsAction.triggered.connect(self.notifyErrors)
        self.notificationMenu.addAction(self.notifyErrorsAction)

        # Barra de Archivos
        self.fileBar = self.menuBar.addMenu('Barra de Archivos')
        self.loadFilesAction = QtWidgets.QAction('Cargar Archivos', self)
        self.loadFilesAction.triggered.connect(self.load_files)
        self.fileBar.addAction(self.loadFilesAction)

        # Resto de la interfaz de usuario
        self.serviceInput = QtWidgets.QLineEdit(self)
        self.serviceInput.setPlaceholderText('Ingrese la URL del servicio')
        self.layout.addWidget(self.serviceInput)

        self.buttonLayout = QtWidgets.QHBoxLayout()

        self.addButton = QtWidgets.QPushButton('Agregar Servicio', self)
        self.addButton.clicked.connect(self.addService)
        self.buttonLayout.addWidget(self.addButton)

        self.deleteButton = QtWidgets.QPushButton('Eliminar Servicio', self)
        self.deleteButton.clicked.connect(self.deleteService)
        self.buttonLayout.addWidget(self.deleteButton)

        self.layout.addLayout(self.buttonLayout)

        self.serviceLayout = QtWidgets.QHBoxLayout()

        self.serviceList = QtWidgets.QListWidget(self)
        self.serviceList.setStyleSheet('background-color: #333; color: #fff;')  # Fondo oscuro para la lista
        self.serviceLayout.addWidget(self.serviceList)

        self.logTable = QtWidgets.QTableWidget(self)
        self.logTable.setColumnCount(4)  # Aumenta el número de columnas a 4
        self.logTable.setHorizontalHeaderLabels(['Servicio', 'Estado', 'Código de Estado', 'Última Consulta'])  # Añade la nueva columna
        self.logTable.setStyleSheet('QTableWidget::item { padding: 5px; }')  # Añade padding para una mejor visualización
        self.serviceLayout.addWidget(self.logTable)

        self.layout.addLayout(self.serviceLayout)

        self.trayIcon = QtWidgets.QSystemTrayIcon(QtGui.QIcon('icon.png'), self)
        self.trayIcon.setToolTip('Servicio de Monitoreo')
        self.trayIcon.activated.connect(self.trayIconActivated)

        self.trayMenu = QtWidgets.QMenu(self)
        self.showAction = QtWidgets.QAction('Mostrar', self)
        self.showAction.triggered.connect(self.show)
        self.trayMenu.addAction(self.showAction)

        self.quitAction = QtWidgets.QAction('Salir', self)
        self.quitAction.triggered.connect(QtWidgets.qApp.quit)
        self.trayMenu.addAction(self.quitAction)

        self.trayIcon.setContextMenu(self.trayMenu)
        self.trayIcon.show()

        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.checkServices)
        self.timer.start(30 * 60000)  # 30 minutos

    def is_auto_start_enabled(self):
        startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
        return os.path.exists(os.path.join(startup_folder, 'ServicioMonitoreo.lnk'))

    def toggle_auto_start(self, checked):
        startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
        shortcut_path = os.path.join(startup_folder, 'ServicioMonitoreo.lnk')
        if checked:
            target = sys.executable
            script = os.path.abspath(__file__)
            icon = os.path.abspath('icon.png')
            self.create_shortcut(shortcut_path, target, script, icon)
        else:
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)

    def create_shortcut(self, shortcut_path, target, script, icon):
        shortcut = pythoncom.CoCreateInstance(
            shell.CLSID_ShellLink, None,
            pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink
        )
        shortcut.SetPath(target)
        shortcut.SetArguments(f'"{script}"')
        shortcut.SetIconLocation(icon, 0)
        persist_file = shortcut.QueryInterface(pythoncom.IID_IPersistFile)
        persist_file.Save(shortcut_path, 0)

    def addService(self):
        url = self.serviceInput.text()
        if url:
            self.services.append(url)
            self.serviceList.addItem(url)
            self.serviceInput.clear()
            self.checkService(url)  # Hacer consulta al agregar el servicio
            self.save_service(url)

    def deleteService(self):
        selected_items = self.serviceList.selectedItems()
        if not selected_items:
            QtWidgets.QMessageBox.warning(self, 'Advertencia', 'Seleccione un servicio para eliminar.')
            return
        for item in selected_items:
            url = item.text()
            self.services.remove(url)
            self.serviceList.takeItem(self.serviceList.row(item))
            self.delete_service_from_file(url)

    def delete_service_from_file(self, url):
        if os.path.exists(SERVICES_FILE):
            with open(SERVICES_FILE, 'r') as f:
                lines = f.readlines()
            with open(SERVICES_FILE, 'w') as f:
                for line in lines:
                    if line.strip() != url:
                        f.write(line)

    def logServiceStatus(self, url, status, code=None):
        rowPosition = self.logTable.rowCount()
        current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
        self.logTable.insertRow(rowPosition)
        self.logTable.setItem(rowPosition, 0, QtWidgets.QTableWidgetItem(url))
        status_item = QtWidgets.QTableWidgetItem(status)
        if status == 'OK':
            status_item.setBackground(QtGui.QColor('green'))
            status_item.setForeground(QtGui.QColor('white'))
        elif status == 'Lento':
            status_item.setBackground(QtGui.QColor('orange'))
            status_item.setForeground(QtGui.QColor('white'))
        elif status == 'Caído':
            status_item.setBackground(QtGui.QColor('red'))
            status_item.setForeground(QtGui.QColor('white'))
        else:
            status_item.setBackground(QtGui.QColor('gray'))
            status_item.setForeground(QtGui.QColor('white'))
        self.logTable.setItem(rowPosition, 1, status_item)
        self.logTable.setItem(rowPosition, 2, QtWidgets.QTableWidgetItem(code if code else 'N/A'))
        self.logTable.setItem(rowPosition, 3, QtWidgets.QTableWidgetItem(current_time))
        self.save_log(url, status, code)

    def checkService(self, url):
        start_time = time.time()
        try:
            response = requests.get(url, timeout=10)
            elapsed_time = time.time() - start_time
            code = response.status_code
            if response.status_code == 200:
                # color = QtGui.QColor('green' if elapsed_time < 1 else 'orange')
                status = 'OK' if elapsed_time < 1 else 'Lento'
            else:
                self.notifyDown(url)
                # color = QtGui.QColor('red')
                status = 'Caído'
            self.logServiceStatus(url, status, code)
        except requests.exceptions.RequestException:
            self.notifyDown(url)
            # color = QtGui.QColor('red')
            status = 'Caído'
            self.logServiceStatus(url, status, 'Error')

        for i in range(self.serviceList.count()):
            if self.serviceList.item(i).text() == url:
                # self.serviceList.item(i).setBackground(color)
                break

    def checkServices(self):
        self.show_waiting_dialog()
        for i in range(self.serviceList.count()):
            url = self.serviceList.item(i).text()
            self.checkService(url)
        self.close_waiting_dialog()

    def notifyDown(self, url):
        send_email('Alerta: Servicio Caído', f'No se pudo acceder a {url}.', TO_EMAILS, CC_EMAILS, BCC_EMAILS)
        self.logServiceStatus(url, 'Caído')

    def notifySlowServices(self):
        slow_services = set()
        for row in range(self.logTable.rowCount()):
            if self.logTable.item(row, 1).text() == 'Lento':
                service = self.logTable.item(row, 0).text()
                slow_services.add(service)

        if slow_services:
            body = 'Los siguientes servicios están lentos:\n' + '\n'.join(slow_services)
            send_email('Alerta: Servicios Lentos', body, TO_EMAILS, CC_EMAILS, BCC_EMAILS)
            QtWidgets.QMessageBox.information(self, 'Notificación Enviada', 'Se ha enviado una notificación sobre los servicios lentos.')
        else:
            QtWidgets.QMessageBox.information(self, 'Sin Servicios Lentos', 'No hay servicios lentos en este momento.')

    def notifyErrors(self):
        errors = set()
        for row in range(self.logTable.rowCount()):
            if self.logTable.item(row, 1).text() == 'Caído':
                service = self.logTable.item(row, 0).text()
                errors.add(service)

        if errors:
            body = 'Los siguientes servicios están caídos:\n' + '\n'.join(errors)
            send_email('Alerta: Servicios Caídos', body, TO_EMAILS, CC_EMAILS, BCC_EMAILS)
            QtWidgets.QMessageBox.information(self, 'Notificación Enviada', 'Se ha enviado una notificación sobre los servicios caídos.')
        else:
            QtWidgets.QMessageBox.information(self, 'Sin Servicios Caídos', 'No hay servicios caídos en este momento.')

    def trayIconActivated(self, reason):
        if reason == QtWidgets.QSystemTrayIcon.Trigger:
            self.show()

    def closeEvent(self, event):
        self.trayIcon.hide()
        QtWidgets.qApp.quit()

    def changeEvent(self, event):
        if event.type() == QtCore.QEvent.WindowStateChange:
            if self.isMinimized():
                self.hide()
                self.trayIcon.showMessage(
                    'Servicio de Monitoreo',
                    'La aplicación sigue ejecutándose en segundo plano.',
                    QtWidgets.QSystemTrayIcon.Information,
                    2000
                )
                event.accept()

    def save_log(self, url, status, code=None):
        with open(LOG_FILE, 'a') as f:
            current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
            f.write(f"{url} | {status} | {code if code else 'N/A'} | {current_time}\n")

    def save_service(self, url):
        with open(SERVICES_FILE, 'a') as f:
            f.write(f"{url}\n")
    def load_logs(self):
        print("Cargando logs...")
        if os.path.exists(LOG_FILE):
            print(f"Archivo de log encontrado: {LOG_FILE}")
            with open(LOG_FILE, 'r') as f:
                for line in f:
                    # Verifica que la línea se divide correctamente
                    try:
                        url, status, code, timestamp = line.strip().split(' | ')
                        rowPosition = self.logTable.rowCount()
                        self.logTable.insertRow(rowPosition)

                        # Configura el item del URL
                        self.logTable.setItem(rowPosition, 0, QtWidgets.QTableWidgetItem(url))

                        # Configura el item del estado con color
                        status_item = QtWidgets.QTableWidgetItem(status)
                        if status == 'OK':
                            status_item.setBackground(QtGui.QColor('green'))
                            status_item.setForeground(QtGui.QColor('white'))
                        elif status == 'Lento':
                            status_item.setBackground(QtGui.QColor('orange'))
                            status_item.setForeground(QtGui.QColor('white'))
                        elif status == 'Caído':
                            status_item.setBackground(QtGui.QColor('red'))
                            status_item.setForeground(QtGui.QColor('white'))
                        else:
                            status_item.setBackground(QtGui.QColor('gray'))
                            status_item.setForeground(QtGui.QColor('white'))
                        self.logTable.setItem(rowPosition, 1, status_item)

                        # Configura el código de estado
                        self.logTable.setItem(rowPosition, 2, QtWidgets.QTableWidgetItem(code))

                        # Configura la última consulta
                        self.logTable.setItem(rowPosition, 3, QtWidgets.QTableWidgetItem(timestamp))
                    except ValueError as e:
                        print(f"Error al procesar la línea: {line.strip()} - {e}")
        else:
            print(f"Archivo de log no encontrado: {LOG_FILE}")

    def load_services(self):
        if os.path.exists(SERVICES_FILE):
            with open(SERVICES_FILE, 'r') as f:
                for line in f:
                    url = line.strip()
                    self.services.append(url)
                    self.serviceList.addItem(url)

    def export_to_excel(self):
        filename, _ = QFileDialog.getSaveFileName(self, 'Guardar Archivo Excel', '', 'Excel Files (*.xlsx)')
        if filename:
            data = []
            for row in range(self.logTable.rowCount()):
                service = self.logTable.item(row, 0).text()
                status = self.logTable.item(row, 1).text()
                code = self.logTable.item(row, 2).text()
                timestamp = self.logTable.item(row, 3).text()
                data.append([service, status, code, timestamp])

            df = pd.DataFrame(data, columns=['Servicio', 'Estado', 'Código de Estado', 'Última Consulta'])
            df.to_excel(filename, index=False)

    def import_from_excel(self):
        filename, _ = QFileDialog.getOpenFileName(self, 'Abrir Archivo Excel', '', 'Excel Files (*.xlsx)')
        if filename:
            df = pd.read_excel(filename)
            for index, row in df.iterrows():
                url = row['Servicio']
                status = row['Estado']
                code = row['Código de Estado']
                timestamp = row['Última Consulta']
                self.logServiceStatus(url, status, code)

    def load_files(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Seleccionar Archivos de Texto", "", "Text Files (*.txt)", options=options)
        for file_path in file_paths:
            with open(file_path, 'r') as file:
                for line in file:
                    url = line.strip()
                    if url and url not in self.services:
                        self.services.append(url)
                        self.serviceList.addItem(url)
                        self.checkService(url)
                        self.save_service(url)

    def show_waiting_dialog(self):
        self.waiting_dialog = QtWidgets.QDialog(self)
        self.waiting_dialog.setWindowTitle("Espere")
        self.waiting_dialog.setModal(True)
        self.waiting_dialog.setWindowFlags(QtCore.Qt.Window | QtCore.Qt.WindowTitleHint | QtCore.Qt.CustomizeWindowHint)
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(QtWidgets.QLabel("Realizando consultas a los servicios..."))
        self.waiting_dialog.setLayout(layout)
        self.waiting_dialog.show()

    def close_waiting_dialog(self):
        self.waiting_dialog.accept()

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    monitor = Monitor()
    monitor.show()
    sys.exit(app.exec_())
