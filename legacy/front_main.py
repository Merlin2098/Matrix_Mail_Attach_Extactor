from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import win32com.client
import json
import os
import sys
import time
import winsound
from datetime import datetime
import ctypes
from ctypes import wintypes
from pathlib import Path

# ⭐ Agregar el directorio raíz al path para importar desde config/ y ui/
# Como front_main.py está en legacy/, necesitamos subir un nivel
sys.path.insert(0, str(Path(__file__).parent.parent))

# Importar adaptadores (misma carpeta)
from extractor_adapter import ExtractorWorker, validar_parametros_extractor
from clasificador_adapter import ClasificadorWorker, validar_carpeta_clasificar

# Importar ConfigManager (desde carpeta raíz/config/)
from config.config_manager import ConfigManager

# ⭐ Importar Estilos (desde carpeta raíz/ui/)
from ui.estilos import Estilos

class AplicacionCorreosPyQt(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # ⭐ Cargar configuración usando el nuevo ConfigManager
        self.config_manager = ConfigManager()

        # ⭐ Configurar icono de la aplicación
        self._configurar_icono()
        
        self.setWindowTitle("Gestión de Correos Outlook")
        self.setGeometry(100, 100, 1200, 750)

        # Aplicar tema inicial
        self.aplicar_tema(self.config_manager.get_tema())
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Barra de menú con tema
        self.crear_menu_tema()
        
        # Título
        titulo = QLabel("Gestión de Correos Outlook")
        titulo.setFont(QFont("Segoe UI", 22, QFont.Bold))
        titulo.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(titulo)
        
        # Pestañas
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Crear pestañas
        self.tabs.addTab(self.crear_pestana_descarga(), "📥 Descarga de Adjuntos")
        self.tabs.addTab(self.crear_pestana_clasificador(), "📂 Clasificar Documentos")
        
    def _configurar_icono(self):
    
        icon_path = self.config_manager.get_icon_path()
        if icon_path and icon_path.exists():
            try:
                self.setWindowIcon(QIcon(str(icon_path)))
                print(f"✅ Icono cargado desde: {icon_path}")
            except Exception as e:
                print(f"⚠️ Error al cargar icono: {e}")
        else:
            print(f"⚠️ Icono no encontrado en: {self.config_manager.icon_path}")
            print("📝 Coloca 'icon.ico' en la carpeta 'config/' para que aparezca")
    
    def crear_menu_tema(self):
        """Crea el menú con opción de cambiar tema"""
        menubar = self.menuBar()
        menu_ver = menubar.addMenu("Ver")
        
        # Acción para tema claro
        accion_claro = QAction("☀️ Tema Claro", self)
        accion_claro.triggered.connect(lambda: self.cambiar_tema("light"))
        menu_ver.addAction(accion_claro)
        
        # Acción para tema oscuro
        accion_oscuro = QAction("🌙 Tema Oscuro", self)
        accion_oscuro.triggered.connect(lambda: self.cambiar_tema("dark"))
        menu_ver.addAction(accion_oscuro)
    
    def cambiar_tema(self, tema):
        """Cambia entre tema claro y oscuro"""
        self.config_manager.set_tema(tema)
        self.aplicar_tema(tema)
    
    def aplicar_tema(self, tema):
        """Aplica el estilo del tema seleccionado"""
        self.setStyleSheet(Estilos.obtener_estilo(tema))
    
    # ==================== PESTAÑA DESCARGA ====================
    
    def crear_pestana_descarga(self):
        """Crea la pestaña de descarga con progreso diferenciado"""
        widget = QWidget()
        layout_principal = QHBoxLayout()
        
        # ===== PANEL IZQUIERDO: Configuración =====
        panel_config = QWidget()
        layout_config = QVBoxLayout()
        
        titulo = QLabel("📥 Descarga de Adjuntos")
        titulo.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout_config.addWidget(titulo)
        
        # Frases de referencia
        grupo_frases = QGroupBox("Frases de referencia")
        layout_frases = QVBoxLayout()
        
        self.txt_frases = QTextEdit()
        self.txt_frases.setPlaceholderText("Ingrese las frases de referencia (una por línea)...")
        self.txt_frases.setMaximumHeight(100)
        layout_frases.addWidget(self.txt_frases)
        
        grupo_frases.setLayout(layout_frases)
        layout_config.addWidget(grupo_frases)
        
        # Carpeta destino
        grupo_destino = QGroupBox("Carpeta de destino")
        layout_destino = QHBoxLayout()
        
        self.txt_destino = QLineEdit()
        self.txt_destino.setPlaceholderText("Seleccione carpeta de destino...")
        self.txt_destino.setReadOnly(True)
        
        btn_explorar_destino = QPushButton("📁 Explorar")
        btn_explorar_destino.clicked.connect(self.seleccionar_carpeta_destino)
        
        layout_destino.addWidget(self.txt_destino)
        layout_destino.addWidget(btn_explorar_destino)
        grupo_destino.setLayout(layout_destino)
        layout_config.addWidget(grupo_destino)
        
        # Bandeja de correo
        grupo_bandeja = QGroupBox("Bandeja de correo")
        layout_bandeja = QHBoxLayout()
        
        self.txt_bandeja = QLineEdit()
        self.txt_bandeja.setPlaceholderText("Seleccione bandeja de Outlook...")
        self.txt_bandeja.setReadOnly(True)
        
        btn_explorar_bandeja = QPushButton("📧 Explorar")
        btn_explorar_bandeja.clicked.connect(self.seleccionar_bandeja_outlook)
        
        layout_bandeja.addWidget(self.txt_bandeja)
        layout_bandeja.addWidget(btn_explorar_bandeja)
        grupo_bandeja.setLayout(layout_bandeja)
        layout_config.addWidget(grupo_bandeja)
        
        # Fechas
        grupo_fechas = QGroupBox("Rango de fechas")
        layout_fechas = QFormLayout()
        
        self.fecha_inicio = QDateEdit()
        self.fecha_inicio.setCalendarPopup(True)
        self.fecha_inicio.setDate(QDate.currentDate().addDays(-7))
        
        self.fecha_fin = QDateEdit()
        self.fecha_fin.setCalendarPopup(True)
        self.fecha_fin.setDate(QDate.currentDate())
        
        layout_fechas.addRow("Fecha inicio:", self.fecha_inicio)
        layout_fechas.addRow("Fecha fin:", self.fecha_fin)
        grupo_fechas.setLayout(layout_fechas)
        layout_config.addWidget(grupo_fechas)
        
        # Botón procesar
        self.btn_procesar = QPushButton("🚀 Iniciar Descarga")
        self.btn_procesar.clicked.connect(self.procesar_descarga)
        self.btn_procesar.setMinimumHeight(40)
        layout_config.addWidget(self.btn_procesar)
        
        layout_config.addStretch()
        panel_config.setLayout(layout_config)
        
        # ===== PANEL DERECHO: Progreso Diferenciado =====
        panel_progreso = QWidget()
        layout_progreso = QVBoxLayout()
        
        titulo_progreso = QLabel("📊 Estado del Proceso")
        titulo_progreso.setFont(QFont("Segoe UI", 14, QFont.Bold))
        layout_progreso.addWidget(titulo_progreso)
        
        # FASE 1: Filtrado
        label_filtrado = QLabel("🔍 Fase 1: Filtrado de Correos")
        label_filtrado.setFont(QFont("Segoe UI", 11, QFont.Bold))
        layout_progreso.addWidget(label_filtrado)
        
        self.log_filtrado_descarga = QTextEdit()
        self.log_filtrado_descarga.setReadOnly(True)
        self.log_filtrado_descarga.setMaximumHeight(250)
        self.log_filtrado_descarga.setPlaceholderText("Los logs de filtrado aparecerán aquí...")
        layout_progreso.addWidget(self.log_filtrado_descarga)
        
        # FASE 2: Descarga
        label_descarga = QLabel("📦 Fase 2: Descarga de Adjuntos")
        label_descarga.setFont(QFont("Segoe UI", 11, QFont.Bold))
        layout_progreso.addWidget(label_descarga)
        
        self.progress_descarga = QProgressBar()
        self.progress_descarga.setTextVisible(True)
        self.progress_descarga.setFormat("%p% - %v/%m archivos")
        self.progress_descarga.setEnabled(False)
        layout_progreso.addWidget(self.progress_descarga)
        
        self.log_descarga = QTextEdit()
        self.log_descarga.setReadOnly(True)
        self.log_descarga.setPlaceholderText("Los logs de descarga aparecerán aquí cuando inicie la fase 2...")
        layout_progreso.addWidget(self.log_descarga)
        
        # Botones de control
        layout_botones = QHBoxLayout()
        
        self.btn_pausar_descarga = QPushButton("⏸️ Pausar")
        self.btn_pausar_descarga.setEnabled(False)
        self.btn_pausar_descarga.clicked.connect(self.pausar_descarga)
        
        self.btn_cancelar_descarga = QPushButton("🛑 Cancelar")
        self.btn_cancelar_descarga.setEnabled(False)
        self.btn_cancelar_descarga.clicked.connect(self.cancelar_descarga)
        
        layout_botones.addWidget(self.btn_pausar_descarga)
        layout_botones.addWidget(self.btn_cancelar_descarga)
        layout_progreso.addLayout(layout_botones)
        
        panel_progreso.setLayout(layout_progreso)
        
        # Agregar paneles
        layout_principal.addWidget(panel_config, 1)
        layout_principal.addWidget(panel_progreso, 2)
        
        widget.setLayout(layout_principal)
        return widget
    
    # ==================== MÉTODOS DESCARGA ====================
    
    def procesar_descarga(self):
        """Inicia el proceso de descarga"""
        # Obtener parámetros
        frases_texto = self.txt_frases.toPlainText().strip()
        frases = [f.strip() for f in frases_texto.split('\n') if f.strip()]
        destino = self.txt_destino.text().strip()
        outlook_folder = self.txt_bandeja.text().strip()
        
        fecha_inicio = datetime.combine(
            self.fecha_inicio.date().toPyDate(),
            datetime.min.time()
        )
        fecha_fin = datetime.combine(
            self.fecha_fin.date().toPyDate(),
            datetime.max.time()
        )
        
        # Validar
        es_valido, mensaje_error = validar_parametros_extractor(
            frases, destino, outlook_folder, fecha_inicio, fecha_fin
        )
        
        if not es_valido:
            QMessageBox.warning(self, "Validación", mensaje_error)
            return
        
        # Limpiar logs
        self.log_filtrado_descarga.clear()
        self.log_descarga.clear()
        self.progress_descarga.setValue(0)
        self.progress_descarga.setEnabled(False)
        
        # Crear worker y thread
        self.thread_descarga = QThread()
        self.worker_descarga = ExtractorWorker()
        self.worker_descarga.moveToThread(self.thread_descarga)
        
        # Conectar señales
        self.worker_descarga.signal_log_filtrado.connect(self.actualizar_log_filtrado)
        self.worker_descarga.signal_log_descarga.connect(self.actualizar_log_descarga)
        self.worker_descarga.signal_progreso.connect(self.actualizar_progreso_descarga)
        self.worker_descarga.signal_inicio_descarga.connect(self.activar_fase_descarga)
        self.worker_descarga.signal_completado.connect(self.descarga_completada)
        self.worker_descarga.signal_error.connect(self.descarga_error)
        
        self.thread_descarga.started.connect(self.worker_descarga.ejecutar)
        
        # Inicializar
        params = {
            'frases': frases,
            'destino': destino,
            'outlook_folder': outlook_folder,
            'fecha_inicio': fecha_inicio,
            'fecha_fin': fecha_fin
        }
        
        self.worker_descarga.inicializar(params)
        
        # Habilitar botones
        self.btn_procesar.setEnabled(False)
        self.btn_cancelar_descarga.setEnabled(True)
        
        # Iniciar
        winsound.Beep(800, 200)  # Inicio del proceso
        self.thread_descarga.start()
    
    def actualizar_log_filtrado(self, mensaje):
        """Actualiza log de filtrado"""
        self.log_filtrado_descarga.append(mensaje)
        scrollbar = self.log_filtrado_descarga.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def actualizar_log_descarga(self, mensaje):
        """Actualiza log de descarga"""
        self.log_descarga.append(mensaje)
        scrollbar = self.log_descarga.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def actualizar_progreso_descarga(self, actual, total, porcentaje):
        """Actualiza barra de progreso"""
        self.progress_descarga.setMaximum(total)
        self.progress_descarga.setValue(actual)
    
    def activar_fase_descarga(self):
        """Activa la fase de descarga"""
        self.progress_descarga.setEnabled(True)
        self.progress_descarga.setValue(0)
        self.btn_pausar_descarga.setEnabled(True)
        
        self.log_filtrado_descarga.append("")
        self.log_filtrado_descarga.append("=" * 60)
        self.log_filtrado_descarga.append("✅ Filtrado completado. Iniciando descarga...")
        winsound.Beep(1000, 150)  # Cambio de fase: filtrado → descarga
        self.log_filtrado_descarga.append("=" * 60)
    
    def descarga_completada(self, estadisticas):
        """Proceso completado"""
        self.thread_descarga.quit()
        self.thread_descarga.wait()
        
        self.btn_procesar.setEnabled(True)
        self.btn_pausar_descarga.setEnabled(False)
        self.btn_cancelar_descarga.setEnabled(False)
        winsound.Beep(1400, 150)
        self.flash_taskbar()
        
        # Ya NO agregamos las estadísticas aquí porque el adapter ya las mostró
        # Solo resetear la UI
    
    def descarga_error(self, mensaje):
        """Error en descarga"""
        self.thread_descarga.quit()
        self.thread_descarga.wait()
        
        self.btn_procesar.setEnabled(True)
        self.btn_pausar_descarga.setEnabled(False)
        self.btn_cancelar_descarga.setEnabled(False)
        
        winsound.Beep(400, 500)
        self.flash_taskbar()
        QMessageBox.critical(self, "Error", mensaje)
            
    def pausar_descarga(self):
        """Pausa/reanuda"""
        if self.btn_pausar_descarga.text() == "⏸️ Pausar":
            self.worker_descarga.pausar()
            self.btn_pausar_descarga.setText("▶️ Reanudar")
        else:
            self.worker_descarga.reanudar()
            self.btn_pausar_descarga.setText("⏸️ Pausar")
    
    def cancelar_descarga(self):
        """Cancela descarga"""
        respuesta = QMessageBox.question(
            self,
            "Confirmar",
            "¿Cancelar el proceso?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if respuesta == QMessageBox.Yes:
            self.worker_descarga.cancelar()
            self.thread_descarga.quit()
            self.thread_descarga.wait()
            
            self.btn_procesar.setEnabled(True)
            self.btn_pausar_descarga.setEnabled(False)
            self.btn_cancelar_descarga.setEnabled(False)
            self.progress_descarga.setEnabled(False)
    
    # ==================== MÉTODOS AUXILIARES ====================
    
    def seleccionar_carpeta_destino(self):
        """Selecciona carpeta de destino"""
        carpeta = QFileDialog.getExistingDirectory(
            self,
            "Seleccionar Carpeta de Destino",
            "",
            QFileDialog.ShowDirsOnly
        )
        
        if carpeta:
            self.txt_destino.setText(carpeta)
    
    def seleccionar_bandeja_outlook(self):
        """Abre diálogo para seleccionar bandeja de Outlook"""
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            dialogo = DialogoBandejasOutlook(self, namespace)
            
            if dialogo.exec_() == QDialog.Accepted:
                bandeja = dialogo.obtener_bandeja_seleccionada()
                if bandeja:
                    self.txt_bandeja.setText(bandeja)
        
        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"No se pudo conectar con Outlook:\n{str(e)}"
            )


# ==================== DIÁLOGO BANDEJA OUTLOOK ====================


    # ==================== PESTAÑA CLASIFICADOR ====================
    
    def crear_pestana_clasificador(self):
        """Crea la pestaña de clasificación de documentos"""
        tab = QWidget()
        layout_principal = QHBoxLayout()
        
        # Panel izquierdo - Configuración
        panel_config = QGroupBox("⚙️ Configuración")
        layout_config = QVBoxLayout()
        
        # Selector de carpeta
        lbl_carpeta = QLabel("📁 Carpeta a clasificar:")
        layout_config.addWidget(lbl_carpeta)
        
        layout_carpeta = QHBoxLayout()
        self.txt_carpeta_clasificar = QLineEdit()
        self.txt_carpeta_clasificar.setPlaceholderText("Selecciona la carpeta con los documentos...")
        btn_carpeta = QPushButton("📂 Seleccionar")
        btn_carpeta.clicked.connect(self.seleccionar_carpeta_clasificar)
        layout_carpeta.addWidget(self.txt_carpeta_clasificar)
        layout_carpeta.addWidget(btn_carpeta)
        layout_config.addLayout(layout_carpeta)
        
        # Información
        info_text = QLabel(
            "ℹ️ <b>El clasificador creará 2 carpetas:</b><br>"
            "• <b>Documentos Firmados:</b> Archivos con 'firmado' en el nombre<br>"
            "• <b>Documentos sin Firmar:</b> Archivos con 'sin firmar' en el nombre<br>"
            "• Los demás archivos se omitirán"
        )
        info_text.setWordWrap(True)
        info_text.setStyleSheet("padding: 10px; background-color: #e3f2fd; border-radius: 5px;")
        layout_config.addWidget(info_text)
        
        layout_config.addStretch()
        
        # Botones
        self.btn_clasificar = QPushButton("▶️ Clasificar Documentos")
        self.btn_clasificar.clicked.connect(self.ejecutar_clasificacion)
        self.btn_clasificar.setMinimumHeight(40)
        layout_config.addWidget(self.btn_clasificar)
        
        self.btn_cancelar_clasificar = QPushButton("⏹️ Cancelar")
        self.btn_cancelar_clasificar.clicked.connect(self.cancelar_clasificacion)
        self.btn_cancelar_clasificar.setEnabled(False)
        layout_config.addWidget(self.btn_cancelar_clasificar)
        
        panel_config.setLayout(layout_config)
        
        # Panel derecho - Progreso
        panel_progreso = QGroupBox("📊 Progreso")
        layout_progreso = QVBoxLayout()
        
        self.progress_clasificar = QProgressBar()
        layout_progreso.addWidget(self.progress_clasificar)
        
        self.log_clasificar = QTextEdit()
        self.log_clasificar.setReadOnly(True)
        layout_progreso.addWidget(self.log_clasificar)
        
        panel_progreso.setLayout(layout_progreso)
        
        # Agregar paneles al layout principal
        layout_principal.addWidget(panel_config, 40)
        layout_principal.addWidget(panel_progreso, 60)
        
        tab.setLayout(layout_principal)
        return tab
    
    def seleccionar_carpeta_clasificar(self):
        """Selecciona carpeta para clasificar"""
        carpeta = QFileDialog.getExistingDirectory(
            self,
            "Seleccionar Carpeta a Clasificar",
            "",
            QFileDialog.ShowDirsOnly
        )
        
        if carpeta:
            self.txt_carpeta_clasificar.setText(carpeta)
    
    def ejecutar_clasificacion(self):
        """Ejecuta la clasificación de documentos"""
        carpeta = self.txt_carpeta_clasificar.text().strip()
        
        # Validar
        es_valido, mensaje_error = validar_carpeta_clasificar(carpeta)
        
        if not es_valido:
            QMessageBox.warning(self, "Validación", mensaje_error)
            return
        
        # Limpiar log
        self.log_clasificar.clear()
        self.progress_clasificar.setValue(0)
        
        # Crear thread y worker
        self.thread_clasificar = QThread()
        self.worker_clasificar = ClasificadorWorker()
        self.worker_clasificar.moveToThread(self.thread_clasificar)
        
        # Configurar
        self.worker_clasificar.inicializar(carpeta)
        
        # Conectar señales
        self.worker_clasificar.signal_progreso.connect(self.actualizar_progreso_clasificar)
        self.worker_clasificar.signal_log.connect(self.actualizar_log_clasificar)
        self.worker_clasificar.signal_completado.connect(self.clasificacion_completada)
        self.worker_clasificar.signal_error.connect(self.clasificacion_error)
        
        self.thread_clasificar.started.connect(self.worker_clasificar.ejecutar)
        
        # Deshabilitar botón
        self.btn_clasificar.setEnabled(False)
        self.btn_cancelar_clasificar.setEnabled(True)
        
        # Iniciar
        self.thread_clasificar.start()
        winsound.Beep(800, 200)
    
    def actualizar_progreso_clasificar(self, actual, total, porcentaje):
        """Actualiza barra de progreso"""
        self.progress_clasificar.setMaximum(total)
        self.progress_clasificar.setValue(actual)
    
    def actualizar_log_clasificar(self, mensaje):
        """Actualiza log"""
        self.log_clasificar.append(mensaje)
        scrollbar = self.log_clasificar.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def clasificacion_completada(self, estadisticas):
            """Clasificación completada"""
            self.thread_clasificar.quit()
            self.thread_clasificar.wait()
            
            self.btn_clasificar.setEnabled(True)
            self.btn_cancelar_clasificar.setEnabled(False)
            
            # Imprimir estadísticas en consola
            print("\n" + "="*50)
            print("RESUMEN DE CLASIFICACIÓN")
            print("="*50)
            print(f"📊 Total de documentos procesados: {estadisticas['total']}")
            print(f"✅ Documentos firmados: {estadisticas['firmados']}")
            print(f"⚠️  Documentos sin firmar: {estadisticas['sin_firmar']}")
            print(f"⏭️  Documentos omitidos: {estadisticas['omitidos']}")
            print(f"❌ Errores encontrados: {estadisticas['errores']}")
            print(f"⏱️  Tiempo total: {estadisticas['tiempo_total']:.2f} segundos")
            print("="*50 + "\n")
            winsound.Beep(1400, 150)
            self.flash_taskbar()

            # Agregar resumen al log visual
            self.log_clasificar.append("\n" + "="*50)
            self.log_clasificar.append("📋 RESUMEN FINAL DE CLASIFICACIÓN")
            self.log_clasificar.append("="*50)
            self.log_clasificar.append(f"📊 Total procesados: {estadisticas['total']}")
            self.log_clasificar.append(f"✅ Firmados: {estadisticas['firmados']}")
            self.log_clasificar.append(f"⚠️  Sin firmar: {estadisticas['sin_firmar']}")
            self.log_clasificar.append(f"⏭️  Omitidos: {estadisticas['omitidos']}")
            self.log_clasificar.append(f"❌ Errores: {estadisticas['errores']}")
            self.log_clasificar.append(f"⏱️  Tiempo: {estadisticas['tiempo_total']:.2f}s")
            self.log_clasificar.append("="*50)
            
            # Scroll al final para ver el resumen
            scrollbar = self.log_clasificar.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())
    def clasificacion_error(self, mensaje):
        """Error en clasificación"""
        self.thread_clasificar.quit()
        self.thread_clasificar.wait()
        
        self.btn_clasificar.setEnabled(True)
        self.btn_cancelar_clasificar.setEnabled(False)
        
        winsound.Beep(400, 500)
        self.flash_taskbar()
        QMessageBox.critical(self, "Error", mensaje)
    
    def cancelar_clasificacion(self):
        """Cancela clasificación"""
        respuesta = QMessageBox.question(
            self,
            "Confirmar",
            "¿Cancelar la clasificación?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if respuesta == QMessageBox.Yes:
            self.worker_clasificar.cancelar()
            self.thread_clasificar.quit()
            self.thread_clasificar.wait()
            
            self.btn_clasificar.setEnabled(True)
            self.btn_cancelar_clasificar.setEnabled(False)
    def flash_taskbar(self):
            """Hace que el icono de la aplicación parpadee en la barra de tareas"""
            try:
                # Constantes de Windows
                FLASHW_ALL = 3  # Parpadea tanto el icono como el caption
                FLASHW_TIMERNOFG = 12  # Parpadea hasta que la ventana vuelva al frente
                
                # Obtener el handle de la ventana
                hwnd = int(self.winId())
                
                # Estructura FLASHWINFO
                class FLASHWINFO(ctypes.Structure):
                    _fields_ = [
                        ('cbSize', wintypes.UINT),
                        ('hwnd', wintypes.HANDLE),
                        ('dwFlags', wintypes.DWORD),
                        ('uCount', wintypes.UINT),
                        ('dwTimeout', wintypes.DWORD)
                    ]
                
                # Configurar el flash
                flash_info = FLASHWINFO()
                flash_info.cbSize = ctypes.sizeof(FLASHWINFO)
                flash_info.hwnd = hwnd
                flash_info.dwFlags = FLASHW_ALL | FLASHW_TIMERNOFG
                flash_info.uCount = 5  # Número de parpadeos (5 veces)
                flash_info.dwTimeout = 0  # Usar velocidad predeterminada del sistema
                
                # Ejecutar el flash
                ctypes.windll.user32.FlashWindowEx(ctypes.byref(flash_info))
            except Exception as e:
                print(f"Error al hacer flash en taskbar: {e}")

class DialogoBandejasOutlook(QDialog):
    """Diálogo para seleccionar bandejas de Outlook - Versión con Carga Diferida"""
    
    def __init__(self, parent, namespace):
        super().__init__(parent)
        self.namespace = namespace
        self.bandeja_seleccionada = None
        
        # Cache para evitar recargar carpetas ya expandidas
        self.carpetas_cargadas = set()
        
        # Referencia a objetos de Outlook para lazy loading
        self.outlook_folders_map = {}  # item_id -> outlook_folder_object
        
        self.inicializar_ui()
        self.cargar_bandejas_inicial()
    
    def inicializar_ui(self):
        """Inicializa la interfaz"""
        self.setWindowTitle("Seleccionar Bandeja de Outlook")
        self.setMinimumSize(900, 550)
        
        layout = QVBoxLayout()
        
        # Instrucciones mejoradas
        label_info = QLabel(
            "⚠️ <b>Importante:</b> Navega hasta la carpeta exacta donde están los correos<br>"
            "💡 <b>Tip:</b> Haz clic en el ➕ para expandir carpetas. Pasa el mouse para ver ruta completa<br>"
            "⚡ <b>Carga rápida:</b> Las subcarpetas se cargan al expandir (lazy loading)"
        )
        label_info.setWordWrap(True)
        layout.addWidget(label_info)
        
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Carpeta", "Ruta completa"])
        self.tree.setColumnWidth(0, 350)
        self.tree.setColumnWidth(1, 500)
        
        # Conectar eventos
        self.tree.itemDoubleClicked.connect(self.aceptar)
        self.tree.itemClicked.connect(self.mostrar_ruta_seleccionada)
        self.tree.itemExpanded.connect(self.cargar_subcarpetas_bajo_demanda)  # ← LAZY LOADING
        
        layout.addWidget(self.tree)
        
        # Label para mostrar la ruta seleccionada
        self.label_ruta = QLabel("📍 Ruta seleccionada: (ninguna)")
        self.label_ruta.setWordWrap(True)
        self.label_ruta.setStyleSheet(
            "padding: 8px; "
            "background-color: #e8f4f8; "
            "border: 1px solid #16A085; "
            "border-radius: 4px; "
            "font-family: 'Consolas', 'Courier New', monospace; "
            "font-size: 11px;"
        )
        layout.addWidget(self.label_ruta)
        
        # Botones
        layout_botones = QHBoxLayout()
        
        btn_aceptar = QPushButton("✓ Aceptar")
        btn_aceptar.clicked.connect(self.aceptar)
        btn_aceptar.setMinimumHeight(35)
        
        btn_cancelar = QPushButton("✗ Cancelar")
        btn_cancelar.clicked.connect(self.reject)
        btn_cancelar.setMinimumHeight(35)
        
        layout_botones.addStretch()
        layout_botones.addWidget(btn_aceptar)
        layout_botones.addWidget(btn_cancelar)
        
        layout.addLayout(layout_botones)
        self.setLayout(layout)
    
    def cargar_bandejas_inicial(self):
        """
        Carga SOLO las cuentas principales (nivel superior)
        Las subcarpetas se cargan bajo demanda
        """
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            
            for carpeta in self.namespace.Folders:
                nombre_cuenta = str(carpeta.Name)
                
                # Crear item de cuenta
                item = QTreeWidgetItem([f"📧 {nombre_cuenta}", nombre_cuenta])
                item.setData(0, Qt.UserRole, nombre_cuenta)
                
                # Tooltips
                item.setToolTip(0, nombre_cuenta)
                item.setToolTip(1, nombre_cuenta)
                
                # Guardar referencia al objeto Outlook para lazy loading
                item_id = id(item)
                self.outlook_folders_map[item_id] = carpeta
                item.setData(0, Qt.UserRole + 1, item_id)  # Guardar ID en el item
                
                # Verificar si tiene subcarpetas
                try:
                    if carpeta.Folders.Count > 0:
                        # Agregar un item "dummy" para mostrar el ➕ de expansión
                        dummy = QTreeWidgetItem(["⏳ Cargando...", ""])
                        item.addChild(dummy)
                except:
                    pass
                
                self.tree.addTopLevelItem(item)
            
            # Expandir solo las cuentas (sin cargar subcarpetas aún)
            for i in range(self.tree.topLevelItemCount()):
                item = self.tree.topLevelItem(i)
                # NO expandir automáticamente para evitar carga
                # El usuario expandirá manualmente cuando lo necesite
            
            QApplication.restoreOverrideCursor()
            
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "Error", f"Error cargando bandejas: {str(e)}")
    
    def cargar_subcarpetas_bajo_demanda(self, item):
        """
        LAZY LOADING: Carga subcarpetas SOLO cuando el usuario expande un nodo
        """
        # Obtener ID del item
        item_id = item.data(0, Qt.UserRole + 1)
        
        # Si ya fue cargado, no hacer nada
        if item_id in self.carpetas_cargadas:
            return
        
        # Marcar como cargado
        self.carpetas_cargadas.add(item_id)
        
        # Obtener objeto de Outlook
        carpeta_outlook = self.outlook_folders_map.get(item_id)
        if not carpeta_outlook:
            return
        
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            
            # Remover items dummy ("Cargando...")
            while item.childCount() > 0:
                item.removeChild(item.child(0))
            
            # Obtener ruta acumulada
            ruta_acumulada = item.data(0, Qt.UserRole)
            
            # Cargar subcarpetas de primer nivel (sin recursión)
            self._agregar_subcarpetas_primer_nivel(carpeta_outlook, item, ruta_acumulada)
            
            QApplication.restoreOverrideCursor()
            
        except Exception as e:
            QApplication.restoreOverrideCursor()
            # Mostrar error solo si es relevante
            if "permission" not in str(e).lower():
                item.addChild(QTreeWidgetItem([f"⚠️ Error: {str(e)[:50]}", ""]))
    
    def _agregar_subcarpetas_primer_nivel(self, carpeta_outlook, item_tree, ruta_acumulada):
        """
        Agrega SOLO el primer nivel de subcarpetas (sin recursión)
        Cada subcarpeta tendrá su propio lazy loading
        """
        try:
            for subcarpeta in carpeta_outlook.Folders:
                nombre_subcarpeta = str(subcarpeta.Name)
                
                # Construir ruta completa
                ruta_completa = f"{ruta_acumulada}\\{nombre_subcarpeta}"
                
                # Crear item con icono
                icono = self._obtener_icono_carpeta(nombre_subcarpeta)
                item_sub = QTreeWidgetItem([f"{icono} {nombre_subcarpeta}", ruta_completa])
                item_sub.setData(0, Qt.UserRole, ruta_completa)
                
                # Tooltips
                item_sub.setToolTip(0, ruta_completa)
                item_sub.setToolTip(1, ruta_completa)
                
                # Guardar referencia para lazy loading
                item_id = id(item_sub)
                self.outlook_folders_map[item_id] = subcarpeta
                item_sub.setData(0, Qt.UserRole + 1, item_id)
                
                item_tree.addChild(item_sub)
                
                # Verificar si tiene subcarpetas (sin cargarlas)
                try:
                    if subcarpeta.Folders.Count > 0:
                        # Agregar dummy para mostrar ➕
                        dummy = QTreeWidgetItem(["⏳ Cargando...", ""])
                        item_sub.addChild(dummy)
                except:
                    pass
                    
        except Exception as e:
            # Carpetas inaccesibles se ignoran silenciosamente
            pass
    
    def _obtener_icono_carpeta(self, nombre):
        """Devuelve un icono según el nombre de la carpeta"""
        nombre_lower = nombre.lower()
        
        if "inbox" in nombre_lower or "bandeja" in nombre_lower or "entrada" in nombre_lower:
            return "📥"
        elif "sent" in nombre_lower or "enviados" in nombre_lower or "enviadas" in nombre_lower:
            return "📤"
        elif "draft" in nombre_lower or "borrador" in nombre_lower:
            return "📝"
        elif "trash" in nombre_lower or "deleted" in nombre_lower or "papelera" in nombre_lower or "eliminados" in nombre_lower:
            return "🗑️"
        elif "spam" in nombre_lower or "junk" in nombre_lower or "no deseado" in nombre_lower or "correo no deseado" in nombre_lower:
            return "🚫"
        elif "archive" in nombre_lower or "archivo" in nombre_lower:
            return "📦"
        else:
            return "📁"
    
    def mostrar_ruta_seleccionada(self, item):
        """Muestra la ruta completa del item seleccionado"""
        ruta = item.data(0, Qt.UserRole)
        if ruta:
            self.label_ruta.setText(f"📍 Ruta seleccionada: {ruta}")
            self.label_ruta.setStyleSheet(
                "padding: 8px; "
                "background-color: #d4edda; "
                "border: 1px solid #28a745; "
                "border-radius: 4px; "
                "font-family: 'Consolas', 'Courier New', monospace; "
                "font-size: 11px;"
            )
    
    def aceptar(self):
        """Acepta la selección"""
        item = self.tree.currentItem()
        if item:
            ruta = item.data(0, Qt.UserRole)
            
            # Validar que no sea el dummy "Cargando..."
            if ruta and "Cargando" not in item.text(0):
                self.bandeja_seleccionada = ruta
                self.accept()
            else:
                QMessageBox.warning(
                    self, 
                    "Selección inválida", 
                    "La carpeta seleccionada no tiene una ruta válida. "
                    "Por favor seleccione una carpeta diferente."
                )
        else:
            QMessageBox.warning(
                self, 
                "Selección requerida", 
                "Por favor seleccione una carpeta de correo antes de continuar."
            )
    
    def obtener_bandeja_seleccionada(self):
        """Retorna la bandeja seleccionada"""
        return self.bandeja_seleccionada


# ==================== SPLASH SCREEN INTEGRADO ====================

class SplashScreenIntegrado(QSplashScreen):
    """
    Splash screen integrado en front_main.py
    Muestra progreso mientras se inicializa la aplicación
    """
    
    def __init__(self, config_manager):
        self.config = config_manager
        self.tema = self.config.get_tema()
        
        # Crear pixmap de fondo
        self.pixmap = self._crear_pixmap()
        super().__init__(self.pixmap, Qt.WindowStaysOnTopHint)
        
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        
        # Crear widgets
        self._crear_widgets()
        self._aplicar_estilos()
    
    def _crear_pixmap(self):
        """Crea el pixmap de fondo"""
        ancho, alto = 500, 300
        pixmap = QPixmap(ancho, alto)
        
        # Colores según tema
        if self.tema == 'dark':
            bg_color = QColor('#1E1E1E')
            border_color = QColor('#16A085')
        else:
            bg_color = QColor('#F0F8F5')
            border_color = QColor('#16A085')
        
        pixmap.fill(bg_color)
        
        # Agregar borde
        painter = QPainter(pixmap)
        painter.setPen(border_color)
        painter.drawRect(0, 0, ancho - 1, alto - 1)
        
        # Cargar icono
        icon_path = self.config.get_icon_path()
        if icon_path and icon_path.exists():
            icon_pixmap = QPixmap(str(icon_path))
            icon_pixmap = icon_pixmap.scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            x = (ancho - icon_pixmap.width()) // 2
            y = 40
            painter.drawPixmap(x, y, icon_pixmap)
        
        painter.end()
        return pixmap
    
    def _crear_widgets(self):
        """Crea los widgets sobre el splash"""
        container = QWidget(self)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(40, 120, 40, 40)
        layout.setSpacing(15)
        
        # Título
        self.label_titulo = QLabel("📧 Gestión de Correos Outlook")
        self.label_titulo.setAlignment(Qt.AlignCenter)
        self.label_titulo.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(self.label_titulo)
        
        # Versión
        self.label_version = QLabel("v1.0 - Cargando...")
        self.label_version.setAlignment(Qt.AlignCenter)
        self.label_version.setFont(QFont("Segoe UI", 9))
        layout.addWidget(self.label_version)
        
        layout.addStretch()
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(25)
        layout.addWidget(self.progress_bar)
        
        # Estado
        self.label_estado = QLabel("Iniciando aplicación...")
        self.label_estado.setAlignment(Qt.AlignCenter)
        self.label_estado.setFont(QFont("Segoe UI", 9))
        layout.addWidget(self.label_estado)
        
        container.setGeometry(0, 0, 500, 300)
    
    def _aplicar_estilos(self):
        """Aplica estilos según el tema"""
        if self.tema == 'dark':
            color_texto = '#E0E0E0'
            color_secundario = '#888888'
            progress_bg = '#2D2D2D'
        else:
            color_texto = '#2C3E50'
            color_secundario = '#7F8C8D'
            progress_bg = '#E0E0E0'
        
        self.label_titulo.setStyleSheet(f"color: {color_texto};")
        self.label_version.setStyleSheet(f"color: {color_secundario};")
        self.label_estado.setStyleSheet(f"color: {color_secundario};")
        
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: 2px solid #16A085;
                border-radius: 5px;
                background-color: {progress_bg};
                text-align: center;
                color: {color_texto};
                font-weight: bold;
            }}
            QProgressBar::chunk {{
                background-color: #16A085;
                border-radius: 3px;
            }}
        """)
    
    def actualizar_progreso(self, valor, mensaje):
        """Actualiza progreso y mensaje"""
        self.progress_bar.setValue(valor)
        self.label_estado.setText(mensaje)
        self.repaint()
    
    def cerrar_con_fade(self):
        """Cierra con efecto fade"""
        self.fade_timer = QTimer()
        self.fade_opacity = 1.0
        
        def fade_step():
            self.fade_opacity -= 0.1
            if self.fade_opacity <= 0:
                self.fade_timer.stop()
                self.close()
            else:
                self.setWindowOpacity(self.fade_opacity)
        
        self.fade_timer.timeout.connect(fade_step)
        self.fade_timer.start(30)


# ==================== LOADER THREAD ====================

class LoaderThread(QThread):
    """Thread que simula carga progresiva"""
    progreso = pyqtSignal(int, str)
    completado = pyqtSignal()
    
    def run(self):
        """Simula carga de componentes"""
        pasos = [
            (0, "Iniciando aplicación..."),
            (20, "Cargando configuración..."),
            (40, "Inicializando componentes..."),
            (60, "Preparando interfaz..."),
            (80, "Finalizando carga..."),
            (100, "¡Aplicación lista!")
        ]
        
        for progreso, mensaje in pasos:
            self.progreso.emit(progreso, mensaje)
            time.sleep(0.3)  # Simular carga
        
        self.completado.emit()


# ==================== MAIN ====================

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Configurar high DPI
    app.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    app.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    # Cargar configuración
    config = ConfigManager()
    
    # Crear y mostrar splash
    splash = SplashScreenIntegrado(config)
    splash.show()
    app.processEvents()
    
    # Contenedor para la ventana principal (evita problema con nonlocal)
    ventana = {'main': None}
    
    # Crear loader thread
    loader = LoaderThread()
    
    def on_progreso(valor, mensaje):
        splash.actualizar_progreso(valor, mensaje)
    
    def on_completado():
        # Crear ventana principal
        ventana['main'] = AplicacionCorreosPyQt()
        
        # Cerrar splash con fade
        splash.cerrar_con_fade()
        
        # Mostrar ventana principal con delay
        QTimer.singleShot(300, ventana['main'].show)
    
    # Conectar señales
    loader.progreso.connect(on_progreso)
    loader.completado.connect(on_completado)
    
    # Iniciar carga
    loader.start()
    
    sys.exit(app.exec_())