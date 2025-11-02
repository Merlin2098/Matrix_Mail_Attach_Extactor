from PyQt5.QtCore import QObject, pyqtSignal
from backend_extractor import (
    ExtractorAdjuntosOutlook,
    FaseProceso,
    NivelMensaje,
    EstadoProceso
)
from datetime import datetime

class ExtractorWorker(QObject):
    """
    Worker thread-safe para integrar ExtractorAdjuntosOutlook con PyQt5.
    
    FASE 1 - MEJORAS:
    - Adaptado al nuevo sistema de callbacks unificados
    - Backend ahora emite fase explícitamente (no hay que adivinarla)
    - Más simple y mantenible
    """
    
    # Señales PyQt5
    signal_log_filtrado = pyqtSignal(str)      # Logs de fase de filtrado
    signal_log_descarga = pyqtSignal(str)      # Logs de fase de descarga
    signal_progreso = pyqtSignal(int, int, float)  # (actual, total, porcentaje)
    signal_inicio_descarga = pyqtSignal()      # Cuando inicia la descarga
    signal_completado = pyqtSignal(dict)       # Estadísticas finales
    signal_error = pyqtSignal(str)             # Errores
    
    def __init__(self):
        super().__init__()
        self.extractor = None
        self.params = None
        
    def inicializar(self, params: dict):
        """
        Inicializa el extractor con los nuevos callbacks unificados.
        
        Args:
            params: Diccionario con parámetros de extracción
        """
        self.params = params
        
        # Crear extractor con callbacks unificados
        self.extractor = ExtractorAdjuntosOutlook(
            callback_mensaje=self._callback_mensaje,
            callback_progreso=self._callback_progreso,
            callback_estado=self._callback_estado
        )
        
        self.signal_log_filtrado.emit("✓ Extractor inicializado correctamente")
    
    def _callback_mensaje(self, fase: FaseProceso, nivel: NivelMensaje, texto: str):
        """
        Callback unificado para mensajes.
        El backend ahora emite la fase explícitamente, no hay que adivinarla.
        
        Args:
            fase: Fase actual del proceso (emitida por el backend)
            nivel: Nivel del mensaje (info, success, warning, error)
            texto: Contenido del mensaje
        """
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Iconos según nivel
        iconos = {
            NivelMensaje.DEBUG: "🔍",
            NivelMensaje.INFO: "ℹ️",
            NivelMensaje.SUCCESS: "✅",
            NivelMensaje.WARNING: "⚠️",
            NivelMensaje.ERROR: "❌"
        }
        icono = iconos.get(nivel, "ℹ️")
        
        msg_formateado = f"[{timestamp}] {icono} {texto}"
        
        # Routing según fase (ahora el backend nos dice en qué fase estamos)
        if fase == FaseProceso.DESCARGA or fase == FaseProceso.FINALIZACION:
            self.signal_log_descarga.emit(msg_formateado)
        else:
            # INICIAL, FILTRADO
            self.signal_log_filtrado.emit(msg_formateado)
        
        # Detectar cambio de fase a DESCARGA para emitir señal especial
        if fase == FaseProceso.DESCARGA and texto.startswith("Iniciando fase"):
            self.signal_log_filtrado.emit("")
            self.signal_log_filtrado.emit("=" * 60)
            self.signal_log_filtrado.emit("✅ Filtrado completado. Iniciando descarga...")
            self.signal_log_filtrado.emit("=" * 60)
            self.signal_inicio_descarga.emit()
    
    def _callback_progreso(self, actual: int, total: int, porcentaje: float):
        """
        Callback para actualización de progreso.
        Solo se usa durante la fase de descarga.
        
        Args:
            actual: Cantidad actual procesada
            total: Cantidad total
            porcentaje: Porcentaje completado
        """
        if total > 0:
            self.signal_progreso.emit(actual, total, porcentaje)
    
    def _callback_estado(self, estado: EstadoProceso):
        """
        Callback para cambios de estado.
        
        Args:
            estado: Nuevo estado del proceso
        """
        # Mapeo de estados a mensajes
        mensajes_estado = {
            EstadoProceso.DETENIDO: "⏹️ Proceso detenido",
            EstadoProceso.INICIANDO: "🚀 Iniciando proceso...",
            EstadoProceso.FILTRANDO: "🔍 Filtrando correos en Outlook...",
            EstadoProceso.PROCESANDO: "📦 Procesando adjuntos...",
            EstadoProceso.PAUSADO: "⏸️ Proceso pausado",
            EstadoProceso.COMPLETADO: "✅ Proceso completado exitosamente",
            EstadoProceso.ERROR: "❌ Error en el proceso",
            EstadoProceso.CANCELADO: "🛑 Proceso cancelado"
        }
        
        mensaje = mensajes_estado.get(estado, estado.value)
        timestamp = datetime.now().strftime("%H:%M:%S")
        msg_completo = f"[{timestamp}] {mensaje}"
        
        # Los estados se emiten al log de filtrado (son estados generales)
        self.signal_log_filtrado.emit(msg_completo)
    
    def ejecutar(self):
        """Ejecuta el proceso completo de extracción"""
        try:
            if not self.params:
                self.signal_error.emit("No se han configurado los parámetros de extracción")
                return
            
            self.signal_log_filtrado.emit("🔍 Iniciando proceso de extracción...")
            self.signal_log_filtrado.emit("")
            
            # Ejecutar extracción
            estadisticas = self.extractor.extraer_adjuntos(
                frases=self.params['frases'],
                destino=self.params['destino'],
                outlook_folder=self.params['outlook_folder'],
                fecha_inicio=self.params['fecha_inicio'],
                fecha_fin=self.params['fecha_fin']
            )
            
            # Mostrar resumen final
            self.signal_log_descarga.emit("")
            self.signal_log_descarga.emit("=" * 60)
            self.signal_log_descarga.emit("🎉 PROCESO COMPLETADO")
            self.signal_log_descarga.emit("=" * 60)
            self.signal_log_descarga.emit("📊 Estadísticas:\n")
            self.signal_log_descarga.emit(f"   📧 Correos procesados: {estadisticas.get('correos_procesados', 0)}")
            self.signal_log_descarga.emit(f"   📎 Adjuntos descargados: {estadisticas.get('adjuntos_descargados', 0)}")
            
            adjuntos_fallidos = estadisticas.get('adjuntos_fallidos', 0)
            if adjuntos_fallidos > 0:
                self.signal_log_descarga.emit(f"   ⚠️ Adjuntos fallidos: {adjuntos_fallidos}")
            
            self.signal_log_descarga.emit(f"   💾 Tamaño total: {estadisticas.get('tamaño_total_mb', 0):.2f} MB")
            self.signal_log_descarga.emit(f"   📈 Tasa de éxito: {estadisticas.get('tasa_exito', 0):.1f}%")
            
            tiempo_total = estadisticas.get('tiempo_total', 0)
            tiempo_str = f"{int(tiempo_total // 60)}min {tiempo_total % 60:.1f}s" if tiempo_total >= 60 else f"{tiempo_total:.1f}s"
            self.signal_log_descarga.emit(f"   ⏱️ Tiempo total: {tiempo_str}")
            self.signal_log_descarga.emit("\n" + "=" * 60)
            
            self.signal_completado.emit(estadisticas)
            
        except ValueError as e:
            # Errores de validación
            error_msg = f"Error de validación: {str(e)}"
            self.signal_error.emit(error_msg)
            self.signal_log_filtrado.emit(f"❌ {error_msg}")
            
        except Exception as e:
            # Otros errores
            error_msg = f"Error durante la extracción: {str(e)}"
            self.signal_error.emit(error_msg)
            self.signal_log_descarga.emit(f"❌ {error_msg}")
    
    def pausar(self):
        """Pausa el proceso"""
        if self.extractor:
            self.extractor.pausar()
    
    def reanudar(self):
        """Reanuda el proceso pausado"""
        if self.extractor:
            self.extractor.reanudar()
    
    def cancelar(self):
        """Cancela el proceso"""
        if self.extractor:
            self.extractor.cancelar()


def validar_parametros_extractor(frases, destino, outlook_folder, fecha_inicio, fecha_fin):
    """
    Función de validación simple para el frontend.
    La validación real ahora está en el backend.
    
    Args:
        frases: Lista de frases de búsqueda
        destino: Carpeta de destino
        outlook_folder: Carpeta de Outlook
        fecha_inicio: Fecha inicial
        fecha_fin: Fecha final
        
    Returns:
        (bool, str): (es_valido, mensaje_error)
    """
    # Crear instancia temporal para validar
    extractor = ExtractorAdjuntosOutlook()
    
    try:
        es_valido, mensaje = extractor.validar_parametros(
            frases, destino, outlook_folder, fecha_inicio, fecha_fin
        )
        return es_valido, mensaje
    except Exception as e:
        return False, f"Error en validación: {str(e)}"