"""
Clase base abstracta para backends de procesamiento.
Proporciona funcionalidad común para todos los backends.

FASE 2 - Base Común
"""

import logging
import os
from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime
from enum import Enum
from pathlib import Path
from threading import Event
from typing import Optional, Callable, Any


# ========================================
# TIPOS Y ENUMS COMPARTIDOS
# ========================================

class FaseProceso(Enum):
    """Fases del proceso (pueden ser extendidas por subclases)"""
    INICIAL = "inicial"
    PROCESANDO = "procesando"
    FINALIZACION = "finalizacion"


class NivelMensaje(Enum):
    """Niveles de mensajes de log"""
    DEBUG = "debug"
    INFO = "info"
    SUCCESS = "success"
    WARNING = "warning"
    ERROR = "error"


class EstadoProceso(Enum):
    """Estados posibles del proceso"""
    DETENIDO = "detenido"
    INICIANDO = "iniciando"
    FILTRANDO = "filtrando"  # Para fase de filtrado (extractor)
    PROCESANDO = "procesando"  # Para fase de procesamiento (extractor)
    CLASIFICANDO = "clasificando"  # Para fase de clasificación (clasificador)
    EN_EJECUCION = "en_ejecucion"  # Estado genérico
    PAUSADO = "pausado"
    COMPLETADO = "completado"
    CANCELADO = "cancelado"
    ERROR = "error"


@dataclass
class EstadisticasBase:
    """Estadísticas base del proceso"""
    tiempo_inicio: Optional[datetime] = None
    tiempo_fin: Optional[datetime] = None
    
    @property
    def tiempo_total(self) -> float:
        """Tiempo total en segundos"""
        if self.tiempo_inicio and self.tiempo_fin:
            return (self.tiempo_fin - self.tiempo_inicio).total_seconds()
        return 0.0


# ========================================
# CLASE BASE ABSTRACTA
# ========================================

class BackendBase(ABC):
    """
    Clase base abstracta para todos los backends de procesamiento.
    
    Proporciona:
    - Sistema de callbacks unificado
    - Control de estados (pausar/reanudar/cancelar)
    - Manejo de fases
    - Utilidades comunes
    - Logging a archivo
    - Validación base
    """
    
    def __init__(self,
                 callback_mensaje: Optional[Callable] = None,
                 callback_progreso: Optional[Callable] = None,
                 callback_estado: Optional[Callable] = None):
        """
        Inicializa el backend con callbacks unificados.
        
        Args:
            callback_mensaje: Función para mensajes (fase, nivel, texto)
            callback_progreso: Función para progreso (actual, total, porcentaje)
            callback_estado: Función para cambios de estado (EstadoProceso)
        """
        self.callback_mensaje = callback_mensaje or self._callback_default
        self.callback_progreso = callback_progreso or self._callback_default
        self.callback_estado = callback_estado or self._callback_default
        
        self.estado_actual = EstadoProceso.DETENIDO
        self.fase_actual = FaseProceso.INICIAL
        self.log_file: Optional[Path] = None
        
        # Control de pausa/cancelación
        self._event_pausa = Event()
        self._event_cancelar = Event()
        self._event_pausa.set()  # No pausado por defecto
        
        self._configurar_logging()
    
    def _callback_default(self, *args, **kwargs):
        """Callback por defecto que no hace nada"""
        pass
    
    def _configurar_logging(self):
        """Configura el sistema de logging interno"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(self.__class__.__name__)
    
    # ========================================
    # MÉTODOS ABSTRACTOS (deben ser implementados por subclases)
    # ========================================
    
    @abstractmethod
    def validar_parametros(self, *args, **kwargs) -> tuple[bool, str]:
        """
        Valida los parámetros específicos del backend.
        
        Returns:
            (bool, str): (es_valido, mensaje_error)
        """
        pass
    
    @abstractmethod
    def _procesar_principal(self, *args, **kwargs) -> dict:
        """
        Método principal de procesamiento (lógica específica).
        
        Returns:
            dict: Estadísticas del proceso
        """
        pass
    
    @abstractmethod
    def _generar_reporte(self) -> dict:
        """
        Genera reporte final de estadísticas.
        
        Returns:
            dict: Reporte con estadísticas
        """
        pass
    
    # ========================================
    # MÉTODOS DE COMUNICACIÓN
    # ========================================
    
    def _enviar_mensaje(self, fase: FaseProceso, nivel: NivelMensaje, texto: str):
        """
        Envía un mensaje con contexto de fase y nivel.
        
        Args:
            fase: Fase actual del proceso
            nivel: Nivel del mensaje
            texto: Contenido del mensaje
        """
        self.callback_mensaje(fase, nivel, texto)
        
        # También escribir en log si está configurado
        if self.log_file:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._escribir_log(f"[{timestamp}] [{nivel.value.upper()}] {texto}")
    
    def _actualizar_progreso(self, actual: int, total: int):
        """
        Actualiza el progreso del proceso.
        
        Args:
            actual: Cantidad actual procesada
            total: Cantidad total a procesar
        """
        porcentaje = (actual / total * 100) if total > 0 else 0.0
        self.callback_progreso(actual, total, porcentaje)
    
    def _cambiar_estado(self, nuevo_estado: EstadoProceso):
        """
        Cambia el estado del proceso y notifica.
        
        Args:
            nuevo_estado: Nuevo estado del proceso
        """
        self.estado_actual = nuevo_estado
        self.callback_estado(nuevo_estado)
        
        if self.log_file:
            self._escribir_log(f"Estado cambiado a: {nuevo_estado.value}")
    
    def _cambiar_fase(self, nueva_fase: FaseProceso):
        """
        Cambia la fase del proceso.
        
        Args:
            nueva_fase: Nueva fase del proceso
        """
        self.fase_actual = nueva_fase
        self._enviar_mensaje(
            nueva_fase,
            NivelMensaje.INFO,
            f"Iniciando fase: {nueva_fase.value}"
        )
    
    # ========================================
    # CONTROL DE FLUJO
    # ========================================
    
    def pausar(self):
        """Pausa el proceso"""
        estados_pausables = (
            EstadoProceso.EN_EJECUCION,
            EstadoProceso.FILTRANDO,
            EstadoProceso.PROCESANDO,
            EstadoProceso.CLASIFICANDO
        )
        if self.estado_actual in estados_pausables:
            self._estado_antes_pausa = self.estado_actual  # Guardar estado anterior
            self._event_pausa.clear()
            self._cambiar_estado(EstadoProceso.PAUSADO)
            self._enviar_mensaje(
                self.fase_actual,
                NivelMensaje.WARNING,
                "Proceso pausado"
            )
    
    def reanudar(self):
        """Reanuda el proceso pausado"""
        if self.estado_actual == EstadoProceso.PAUSADO:
            self._event_pausa.set()
            # Restaurar estado anterior si existe, sino EN_EJECUCION
            estado_restaurar = getattr(self, '_estado_antes_pausa', EstadoProceso.EN_EJECUCION)
            self._cambiar_estado(estado_restaurar)
            self._enviar_mensaje(
                self.fase_actual,
                NivelMensaje.INFO,
                "Proceso reanudado"
            )
    
    def cancelar(self):
        """Cancela el proceso"""
        self._event_cancelar.set()
        self._event_pausa.set()
        self._cambiar_estado(EstadoProceso.CANCELADO)
        self._enviar_mensaje(
            self.fase_actual,
            NivelMensaje.WARNING,
            "Proceso cancelado por el usuario"
        )
    
    def _verificar_cancelacion(self):
        """Verifica si se solicitó cancelación"""
        if self._event_cancelar.is_set():
            raise InterruptedError("Proceso cancelado por el usuario")
    
    def _verificar_pausa(self):
        """Espera si el proceso está pausado"""
        self._event_pausa.wait()
    
    def _resetear_control(self):
        """Resetea los eventos de control para un nuevo proceso"""
        self._event_cancelar.clear()
        self._event_pausa.set()
    
    # ========================================
    # UTILIDADES COMUNES
    # ========================================
    
    def _manejar_nombre_duplicado(self, ruta_archivo: Path) -> Path:
        """
        Maneja nombres de archivo duplicados agregando sufijos _1, _2, etc.
        
        Args:
            ruta_archivo: Ruta del archivo a verificar
            
        Returns:
            Path: Ruta del archivo sin duplicados
        """
        if not ruta_archivo.exists():
            return ruta_archivo
        
        carpeta = ruta_archivo.parent
        nombre_base = ruta_archivo.stem
        extension = ruta_archivo.suffix
        contador = 1
        
        while ruta_archivo.exists():
            nuevo_nombre = f"{nombre_base}_{contador}{extension}"
            ruta_archivo = carpeta / nuevo_nombre
            contador += 1
            
            # Protección contra loops infinitos
            if contador > 1000:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                ruta_archivo = carpeta / f"{nombre_base}_{timestamp}{extension}"
                break
        
        return ruta_archivo
    
    def _crear_carpeta_segura(self, ruta: Path) -> bool:
        """
        Crea una carpeta de forma segura.
        
        Args:
            ruta: Ruta de la carpeta a crear
            
        Returns:
            bool: True si se creó exitosamente
        """
        try:
            ruta.mkdir(parents=True, exist_ok=True)
            return True
        except Exception as e:
            self._enviar_mensaje(
                self.fase_actual,
                NivelMensaje.ERROR,
                f"Error al crear carpeta {ruta}: {str(e)}"
            )
            return False
    
    def _verificar_permisos_escritura(self, ruta: Path) -> bool:
        """
        Verifica permisos de escritura en una ruta.
        
        Args:
            ruta: Ruta a verificar
            
        Returns:
            bool: True si tiene permisos de escritura
        """
        return os.access(str(ruta), os.W_OK)
    
    # ========================================
    # LOGGING A ARCHIVO
    # ========================================
    
    def _crear_log_archivo(self, carpeta_destino: str, prefijo: str = "log"):
        """
        Crea el archivo de log con timestamp.
        
        Args:
            carpeta_destino: Carpeta donde crear el log
            prefijo: Prefijo del nombre del archivo
        """
        ahora = datetime.now()
        fecha_str = ahora.strftime("%d.%m.%Y")
        hora_str = ahora.strftime("%H.%M.%S")
        nombre_log = f"{prefijo}_fecha({fecha_str})_hora({hora_str}).log"
        
        self.log_file = Path(carpeta_destino) / nombre_log
        
        # Escribir encabezado
        self._escribir_log("=" * 80)
        self._escribir_log(f"LOG DE PROCESAMIENTO - {self.__class__.__name__.upper()}")
        self._escribir_log("=" * 80)
        self._escribir_log(f"Inicio: {ahora.strftime('%Y-%m-%d %H:%M:%S')}")
        self._escribir_log("")
    
    def _escribir_log(self, mensaje: str):
        """
        Escribe un mensaje en el archivo de log.
        
        Args:
            mensaje: Mensaje a escribir
        """
        if self.log_file:
            try:
                # Crear el archivo si no existe (modo 'a' crea automáticamente)
                with open(self.log_file, 'a', encoding='utf-8') as f:
                    f.write(f"{mensaje}\n")
            except Exception as e:
                self.logger.error(f"Error al escribir log: {e}")
    
    def _finalizar_log_archivo(self, estadisticas: dict):
        """
        Finaliza el archivo de log con estadísticas.
        
        Args:
            estadisticas: Diccionario con estadísticas finales
        """
        if not self.log_file:
            return
        
        self._escribir_log("")
        self._escribir_log("=" * 80)
        self._escribir_log("RESUMEN FINAL")
        self._escribir_log("=" * 80)
        
        # Escribir estadísticas
        for clave, valor in estadisticas.items():
            if clave != 'tiempo_total':
                self._escribir_log(f"{clave}: {valor}")
        
        # Tiempo total
        tiempo_total = estadisticas.get('tiempo_total', 0)
        if tiempo_total < 60:
            tiempo_str = f"{tiempo_total:.1f}s"
        else:
            minutos = int(tiempo_total // 60)
            segundos = tiempo_total % 60
            tiempo_str = f"{minutos}min {segundos:.1f}s"
        
        self._escribir_log(f"Tiempo total: {tiempo_str}")
        
        # Estado final
        if self.estado_actual == EstadoProceso.COMPLETADO:
            self._escribir_log("Estado: ✅ Completado exitosamente")
        elif self.estado_actual == EstadoProceso.CANCELADO:
            self._escribir_log("Estado: 🛑 Cancelado")
        elif self.estado_actual == EstadoProceso.ERROR:
            self._escribir_log("Estado: ❌ Error")
        
        self._escribir_log("=" * 80)
        self._escribir_log(f"Log guardado en: {self.log_file.absolute()}")
        self._escribir_log("=" * 80)
    
    # ========================================
    # VALIDACIÓN BASE
    # ========================================
    
    def _validar_carpeta_existe(self, carpeta: str, nombre: str = "carpeta") -> tuple[bool, str]:
        """
        Valida que una carpeta existe.
        
        Args:
            carpeta: Ruta de la carpeta
            nombre: Nombre descriptivo de la carpeta
            
        Returns:
            (bool, str): (es_valido, mensaje_error)
        """
        if not carpeta or not carpeta.strip():
            return False, f"Debe seleccionar una {nombre}"
        
        if not os.path.exists(carpeta):
            return False, f"La {nombre} no existe"
        
        if not os.path.isdir(carpeta):
            return False, f"La ruta no es una {nombre} válida"
        
        return True, ""
    
    def _validar_rango_fechas(self, fecha_inicio: datetime, fecha_fin: datetime) -> tuple[bool, str]:
        """
        Valida un rango de fechas.
        
        Args:
            fecha_inicio: Fecha inicial
            fecha_fin: Fecha final
            
        Returns:
            (bool, str): (es_valido, mensaje_error)
        """
        if not fecha_inicio or not fecha_fin:
            return False, "Debe seleccionar fechas de inicio y fin"
        
        if fecha_inicio > fecha_fin:
            return False, "La fecha de inicio no puede ser posterior a la fecha fin"
        
        return True, ""
    
    # ========================================
    # FLUJO DE EJECUCIÓN BASE
    # ========================================
    
    def ejecutar(self, *args, **kwargs) -> dict:
        """
        Método template para ejecutar el proceso completo.
        
        Este método coordina el flujo general:
        1. Validación de parámetros
        2. Inicialización
        3. Procesamiento principal (delegado a subclase)
        4. Finalización
        
        Returns:
            dict: Estadísticas del proceso
        """
        try:
            # Validar parámetros
            es_valido, mensaje_error = self.validar_parametros(*args, **kwargs)
            if not es_valido:
                self._enviar_mensaje(
                    FaseProceso.INICIAL,
                    NivelMensaje.ERROR,
                    f"Validación fallida: {mensaje_error}"
                )
                raise ValueError(mensaje_error)
            
            # Resetear estado
            self._resetear_control()
            self._cambiar_estado(EstadoProceso.INICIANDO)
            
            # Inicializar tiempo
            tiempo_inicio = datetime.now()
            
            # Procesamiento principal (implementado por subclase)
            resultado = self._procesar_principal(*args, **kwargs)
            
            # Finalizar
            tiempo_fin = datetime.now()
            resultado['tiempo_total'] = (tiempo_fin - tiempo_inicio).total_seconds()
            
            if not self._event_cancelar.is_set():
                self._cambiar_estado(EstadoProceso.COMPLETADO)
            
            return resultado
            
        except InterruptedError:
            self._cambiar_estado(EstadoProceso.CANCELADO)
            return self._generar_reporte()
            
        except Exception as e:
            self._cambiar_estado(EstadoProceso.ERROR)
            self._enviar_mensaje(
                self.fase_actual,
                NivelMensaje.ERROR,
                f"Error durante el proceso: {str(e)}"
            )
            raise
    
    # ========================================
    # REPRESENTACIÓN
    # ========================================
    
    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(estado={self.estado_actual.value}, fase={self.fase_actual.value})"