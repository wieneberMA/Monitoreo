# monitoreo_de_servicios_web

Esta aplicación permite monitorear el estado de varios servicios web, enviando notificaciones por correo electrónico en caso de que alguno de los servicios esté lento o caído.

## Características

- **Monitoreo de servicios web**: Agrega y elimina servicios web para monitorear su estado.
- **Notificaciones por correo electrónico**: Recibe alertas cuando un servicio está lento o caído.
- **Registro de estado**: Mantén un registro del estado de los servicios monitoreados.
- **Importar/Exportar a Excel**: Importa y exporta la lista de servicios y sus estados desde/hacia un archivo de Excel.
- **Bandeja del sistema**: La aplicación se ejecuta en segundo plano y se puede acceder desde la bandeja del sistema.

## Instalación

### Requisitos

- Python 3.8+
- pip

### Pasos de Instalación

1. Clona este repositorio:

   ```bash
   git clone https://github.com/tu_usuario/monitoreo_de_servicios_web.git
   cd monitoreo_de_servicios_web
   ```
2. Crea y activa un entorno virtual:
   En windows: - python -m venv env
   .\env\Scripts\Activate.ps1
En linux: - python -m venv env
source env/bin/activate
3. Instalar dependencias
   -pip install -r requirements.txt
4. Correr
   - py monitoreo.py
