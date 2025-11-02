from PyQt5.QtCore import QObject, pyqtSignal
from backend_clasificador import (
    ClasificadorDocumentos,
    FaseProceso,
    NivelMensaje,
    EstadoProceso
)
from datetime import datetime


class ClasificadorWorker(QObject):
    """
    Worker thread-safe para integrar ClasificadorDocumentos con PyQt5.
    
    FASE 1 - MEJORAS:
    - Adaptado al nuevo sistema de callbacks unificados
    - Backend ahora emite fase explícitamente
    - Arquitectura consistente con ExtractorWorker
    """
    
    # Señales PyQt5
    signal_progreso = pyqtSignal(int, int, float)  # (actual, total, porcentaje)
    signal_log = pyqtSignal(str)                   # Logs generales
    signal_completado = pyqtSignal(dict)           # Estadísticas finales
    signal_error = pyqtSignal(str)                 # Errores
    
    def __init__(self):
        super().__init__()
        self.clasificador = None
        self.carpeta = None
    
    def inicializar(self, carpeta: str):
        """
        Inicializa el clasificador con los nuevos callbacks unificados.
        
        Args:
            carpeta: Carpeta con documentos a clasificar
        """
        self.carpeta = carpeta
        
        # Crear clasificador con callbacks unificados
        self.clasificador = ClasificadorDocumentos(
            callback_mensaje=self._callback_mensaje,
            callback_progreso=self._callback_progreso,
            callback_estado=self._callback_estado
        )
    
    def _callback_mensaje(self, fase: FaseProceso, nivel: NivelMensaje, texto: str):
        """
        Callback unificado para mensajes.
        El backend ahora emite la fase explícitamente.
        
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
        self.signal_log.emit(msg_formateado)
    
    def _callback_progreso(self, actual: int, total: int, porcentaje: float):
        """
        Callback para actualización de progreso.
        
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
            EstadoProceso.CLASIFICANDO: "📂 Clasificando documentos...",
            EstadoProceso.PAUSADO: "⏸️ Proceso pausado",
            EstadoProceso.COMPLETADO: "✅ Proceso completado exitosamente",
            EstadoProceso.ERROR: "❌ Error en el proceso",
            EstadoProceso.CANCELADO: "🛑 Proceso cancelado"
        }
        
        mensaje = mensajes_estado.get(estado, estado.value)
        timestamp = datetime.now().strftime("%H:%M:%S")
        msg_completo = f"[{timestamp}] {mensaje}"
        
        self.signal_log.emit(msg_completo)
    
    def ejecutar(self):
        """Ejecuta el proceso completo de clasificación"""
        try:
            self.signal_log.emit("📂 Iniciando proceso de clasificación...")
            self.signal_log.emit("")
            
            # Ejecutar clasificación
            estadisticas = self.clasificador.clasificar(self.carpeta)
            
            # Mostrar resumen final
            self.signal_log.emit("")
            self.signal_log.emit("=" * 60)
            self.signal_log.emit("🎉 PROCESO COMPLETADO")
            self.signal_log.emit("=" * 60)
            self.signal_log.emit("📊 Estadísticas:\n")
            self.signal_log.emit(f"   📄 Total de archivos: {estadisticas.get('total', 0)}")
            self.signal_log.emit(f"   ✅ Documentos firmados: {estadisticas.get('firmados', 0)}")
            self.signal_log.emit(f"   ⚠️ Documentos sin firmar: {estadisticas.get('sin_firmar', 0)}")
            self.signal_log.emit(f"   ⏭️ Archivos omitidos: {estadisticas.get('omitidos', 0)}")
            
            errores = estadisticas.get('errores', 0)
            if errores > 0:
                self.signal_log.emit(f"   ❌ Errores: {errores}")
            
            tiempo_total = estadisticas.get('tiempo_total', 0)
            tiempo_str = f"{int(tiempo_total // 60)}min {tiempo_total % 60:.1f}s" if tiempo_total >= 60 else f"{tiempo_total:.1f}s"
            self.signal_log.emit(f"   ⏱️ Tiempo total: {tiempo_str}")
            self.signal_log.emit("\n" + "=" * 60)
            
            self.signal_completado.emit(estadisticas)
            
        except ValueError as e:
            # Errores de validación
            error_msg = f"Error de validación: {str(e)}"
            self.signal_error.emit(error_msg)
            self.signal_log.emit(f"❌ {error_msg}")
            
        except Exception as e:
            # Otros errores
            error_msg = f"Error durante la clasificación: {str(e)}"
            self.signal_error.emit(error_msg)
            self.signal_log.emit(f"❌ {error_msg}")
    
    def cancelar(self):
        """Cancela el proceso"""
        if self.clasificador:
            self.clasificador.cancelar()


def validar_carpeta_clasificar(carpeta: str) -> tuple[bool, str]:
    """
    Función de validación simple para el frontend.
    La validación real ahora está en el backend.
    
    Args:
        carpeta: Carpeta a validar
        
    Returns:
        (bool, str): (es_valido, mensaje_error)
    """
    # Crear instancia temporal para validar
    clasificador = ClasificadorDocumentos()
    
    try:
        es_valido, mensaje = clasificador.validar_parametros(carpeta)
        return es_valido, mensaje
    except Exception as e:
        return False, f"Error en validación: {str(e)}"