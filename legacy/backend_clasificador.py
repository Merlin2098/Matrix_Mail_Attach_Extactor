"""
Clasificador de documentos según estado de firma.

FASE 2 - Hereda de BackendBase
"""

import os
import shutil
from pathlib import Path
from dataclasses import dataclass
from datetime import datetime
from enum import Enum
from typing import Optional

from backend_base import (
    BackendBase,
    FaseProceso as FaseBase,
    NivelMensaje,
    EstadoProceso,
    EstadisticasBase
)


# ========================================
# RE-EXPORTAR PARA USO EXTERNO
# ========================================
__all__ = [
    'ClasificadorDocumentos',
    'FaseProceso',
    'NivelMensaje', 
    'EstadoProceso',  # Re-exportado desde backend_base
    'EstadisticasClasificacion'
]


# ========================================
# ENUMS ESPECÍFICOS
# ========================================

class FaseProceso(Enum):
    """Fases específicas del proceso de clasificación"""
    INICIAL = "inicial"
    CLASIFICANDO = "clasificando"
    FINALIZACION = "finalizacion"


# ========================================
# ESTADÍSTICAS ESPECÍFICAS
# ========================================

@dataclass
class EstadisticasClasificacion(EstadisticasBase):
    """Estadísticas del proceso de clasificación"""
    total: int = 0
    firmados: int = 0
    sin_firmar: int = 0
    omitidos: int = 0
    errores: int = 0


# ========================================
# CLASE PRINCIPAL
# ========================================

class ClasificadorDocumentos(BackendBase):
    """
    Clasificador de documentos según estado de firma.
    
    FASE 2 - Hereda de BackendBase:
    - Reutiliza callbacks, control de flujo, utilidades
    - Solo implementa lógica específica de clasificación
    - Arquitectura consistente con ExtractorAdjuntosOutlook
    """
    
    def __init__(self, 
                 callback_mensaje=None,
                 callback_progreso=None,
                 callback_estado=None):
        """Inicializa el clasificador"""
        super().__init__(callback_mensaje, callback_progreso, callback_estado)
        
        self.estadisticas = EstadisticasClasificacion()
        self.cancelado = False
    
    # ========================================
    # IMPLEMENTACIÓN DE MÉTODOS ABSTRACTOS
    # ========================================
    
    def validar_parametros(self, carpeta_origen: str) -> tuple[bool, str]:
        """Valida los parámetros de clasificación"""
        
        # Validar carpeta existe
        es_valido, mensaje = self._validar_carpeta_existe(carpeta_origen, "carpeta de origen")
        if not es_valido:
            return False, mensaje
        
        # Verificar permisos de escritura
        carpeta_path = Path(carpeta_origen)
        if not self._verificar_permisos_escritura(carpeta_path):
            return False, "No tiene permisos de escritura en la carpeta"
        
        return True, ""
    
    def _procesar_principal(self, carpeta_origen: str) -> dict:
        """Procesamiento principal de clasificación"""
        
        # Resetear estadísticas
        self.estadisticas = EstadisticasClasificacion()
        self.estadisticas.tiempo_inicio = datetime.now()
        self.cancelado = False
        
        try:
            # Cambiar a fase inicial
            self._cambiar_fase(FaseProceso.INICIAL)
            
            carpeta_path = Path(carpeta_origen)
            
            # Crear carpetas de destino
            carpeta_firmados = carpeta_path / "Documentos Firmados"
            carpeta_sin_firmar = carpeta_path / "Documentos sin Firmar"
            
            if not self._crear_carpeta_segura(carpeta_firmados):
                raise Exception("No se pudo crear carpeta de firmados")
            
            if not self._crear_carpeta_segura(carpeta_sin_firmar):
                raise Exception("No se pudo crear carpeta de sin firmar")
            
            self._enviar_mensaje(
                FaseProceso.INICIAL,
                NivelMensaje.SUCCESS,
                "Carpetas de destino creadas correctamente"
            )
            
            # Obtener archivos a procesar
            archivos = [f for f in carpeta_path.iterdir() if f.is_file()]
            total = len(archivos)
            self.estadisticas.total = total
            
            if total == 0:
                self._enviar_mensaje(
                    FaseProceso.INICIAL,
                    NivelMensaje.WARNING,
                    "No se encontraron archivos para clasificar"
                )
                self.estadisticas.tiempo_fin = datetime.now()
                return self._generar_reporte()
            
            self._enviar_mensaje(
                FaseProceso.INICIAL,
                NivelMensaje.INFO,
                f"Se encontraron {total} archivos para clasificar"
            )
            
            # Iniciar clasificación
            self._cambiar_fase(FaseProceso.CLASIFICANDO)
            self._cambiar_estado(EstadoProceso.CLASIFICANDO)  # ← Cambiado
            
            procesados = 0
            
            for archivo in archivos:
                if self.cancelado:
                    break
                
                self._verificar_cancelacion()
                self._verificar_pausa()
                
                # Clasificar archivo
                self._clasificar_archivo(archivo, carpeta_firmados, carpeta_sin_firmar)
                
                procesados += 1
                self._actualizar_progreso(procesados, total)
                
                # Log cada 10 archivos o al final
                if procesados % 10 == 0 or procesados == total:
                    self._enviar_mensaje(
                        FaseProceso.CLASIFICANDO,
                        NivelMensaje.INFO,
                        f"Procesados: {procesados}/{total} ({(procesados/total)*100:.1f}%)"
                    )
            
            # Finalizar
            self._cambiar_fase(FaseProceso.FINALIZACION)
            self.estadisticas.tiempo_fin = datetime.now()
            
            if self.cancelado:
                self._cambiar_estado(EstadoProceso.CANCELADO)
            else:
                self._cambiar_estado(EstadoProceso.COMPLETADO)
                self._enviar_mensaje(
                    FaseProceso.FINALIZACION,
                    NivelMensaje.SUCCESS,
                    f"Clasificación completada: {self.estadisticas.firmados} firmados, {self.estadisticas.sin_firmar} sin firmar"
                )
            
            return self._generar_reporte()
            
        except InterruptedError:
            self.estadisticas.tiempo_fin = datetime.now()
            raise
    
    def _generar_reporte(self) -> dict:
        """Genera reporte final de estadísticas"""
        return {
            'total': self.estadisticas.total,
            'firmados': self.estadisticas.firmados,
            'sin_firmar': self.estadisticas.sin_firmar,
            'omitidos': self.estadisticas.omitidos,
            'errores': self.estadisticas.errores,
            'tiempo_total': self.estadisticas.tiempo_total
        }
    
    # ========================================
    # SOBRESCRITURA DE CANCELAR
    # ========================================
    
    def cancelar(self):
        """Cancela el proceso (sobrescribe método base)"""
        super().cancelar()
        self.cancelado = True
    
    # ========================================
    # LÓGICA ESPECÍFICA DE CLASIFICACIÓN
    # ========================================
    
    def _clasificar_archivo(self, archivo: Path, 
                           carpeta_firmados: Path, 
                           carpeta_sin_firmar: Path) -> str:
        """
        Clasifica un archivo individual según su nombre.
        
        Args:
            archivo: Archivo a clasificar
            carpeta_firmados: Carpeta de destino para firmados
            carpeta_sin_firmar: Carpeta de destino para sin firmar
            
        Returns:
            str: Resultado de la clasificación ('firmado', 'sin_firmar', 'omitido', 'error')
        """
        nombre_lower = archivo.name.lower()
        
        try:
            # Verificar si es "sin firmar" (prioridad)
            # Detectar tanto "sin firmar", "sin_firmar", "sinfirmar", "not signed", "not_signed"
            es_sin_firmar = (
                "sin firmar" in nombre_lower or 
                "sin_firmar" in nombre_lower or
                "sinfirmar" in nombre_lower or
                "not signed" in nombre_lower or
                "not_signed" in nombre_lower or
                "notsigned" in nombre_lower
            )
            
            if es_sin_firmar:
                destino = carpeta_sin_firmar / archivo.name
                shutil.move(str(archivo), str(destino))
                self.estadisticas.sin_firmar += 1
                self._enviar_mensaje(
                    FaseProceso.CLASIFICANDO,
                    NivelMensaje.WARNING,
                    f"⚠️ Sin firmar: {archivo.name}"
                )
                return 'sin_firmar'
            
            # Verificar si es "firmado"
            elif "firmado" in nombre_lower or "signed" in nombre_lower:
                destino = carpeta_firmados / archivo.name
                shutil.move(str(archivo), str(destino))
                self.estadisticas.firmados += 1
                self._enviar_mensaje(
                    FaseProceso.CLASIFICANDO,
                    NivelMensaje.SUCCESS,
                    f"✅ Firmado: {archivo.name}"
                )
                return 'firmado'
            
            # No coincide con ningún criterio
            else:
                self.estadisticas.omitidos += 1
                return 'omitido'
                
        except PermissionError:
            self.estadisticas.errores += 1
            self._enviar_mensaje(
                FaseProceso.CLASIFICANDO,
                NivelMensaje.ERROR,
                f"❌ Archivo bloqueado: {archivo.name}"
            )
            return 'error'
            
        except Exception as e:
            self.estadisticas.errores += 1
            self._enviar_mensaje(
                FaseProceso.CLASIFICANDO,
                NivelMensaje.ERROR,
                f"❌ Error con {archivo.name}: {str(e)}"
            )
            return 'error'
    
    # ========================================
    # MÉTODO PÚBLICO PRINCIPAL
    # ========================================
    
    def clasificar(self, carpeta_origen: str) -> dict:
        """
        Método principal para clasificar documentos.
        Usa el template method pattern heredado de BackendBase.
        """
        return self.ejecutar(carpeta_origen)


# ========================================
# EJEMPLO DE USO
# ========================================
if __name__ == "__main__":
    import time
    
    def callback_mensaje(fase, nivel, texto: str):
        print(f"[{fase.value.upper()}] [{nivel.value.upper()}] {texto}")
    
    def callback_progreso(actual: int, total: int, porcentaje: float):
        print(f"Progreso: {actual}/{total} ({porcentaje:.1f}%)")
    
    def callback_estado(estado):
        print(f"Estado: {estado.value}")
    
    clasificador = ClasificadorDocumentos(
        callback_mensaje=callback_mensaje,
        callback_progreso=callback_progreso,
        callback_estado=callback_estado
    )
    
    print("\n" + "="*50)
    print("EJEMPLO DE USO - Clasificador de Documentos")
    print("="*50)
    print("\n⚠️ Modifica esta ruta según tus necesidades:\n")
    
    carpeta_ejemplo = r"C:\Users\usuario\Downloads\Documentos"
    print(f"Carpeta: {carpeta_ejemplo}")
    print("\n" + "="*50)
    
    respuesta = input("\n¿Deseas ejecutar la clasificación? (s/n): ")
    
    if respuesta.lower() == 's':
        try:
            print("\nIniciando clasificación...")
            stats = clasificador.clasificar(carpeta_ejemplo)
            
            print("\n" + "="*50)
            print("RESUMEN FINAL")
            print("="*50)
            print(f"Total de archivos: {stats['total']}")
            print(f"Documentos firmados: {stats['firmados']}")
            print(f"Documentos sin firmar: {stats['sin_firmar']}")
            print(f"Archivos omitidos: {stats['omitidos']}")
            print(f"Errores: {stats['errores']}")
            print(f"Tiempo total: {stats['tiempo_total']:.1f} segundos")
            print("="*50)
                
        except KeyboardInterrupt:
            print("\n\nCancelando...")
            clasificador.cancelar()
            time.sleep(1)
        except Exception as e:
            print(f"Error: {e}")
    else:
        print("\nEjecución cancelada.")
    
    print("\nPresiona Enter para salir...")
    input()