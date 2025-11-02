"""
Tests unitarios para BackendBase

FASE 2 - Test de la clase base común
"""

import unittest
from unittest.mock import Mock, patch, MagicMock
from pathlib import Path
from datetime import datetime
import tempfile
import os
import sys

# Agregar el directorio raíz y legacy al path
sys.path.insert(0, str(Path(__file__).parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "legacy"))

from legacy.backend_base import (
    BackendBase,
    FaseProceso,
    NivelMensaje,
    EstadoProceso,
    EstadisticasBase
)


# ========================================
# BACKEND MOCK PARA TESTING
# ========================================

class MockBackend(BackendBase):
    """Backend mock para testing de la clase base"""
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.procesar_llamado = False
        self.validar_llamado = False
    
    def validar_parametros(self, valor: int) -> tuple[bool, str]:
        """Validación mock"""
        self.validar_llamado = True
        if valor > 0:
            return True, ""
        return False, "Valor debe ser mayor a 0"
    
    def _procesar_principal(self, valor: int) -> dict:
        """Procesamiento mock"""
        self.procesar_llamado = True
        return {"procesados": valor, "tiempo_total": 1.0}
    
    def _generar_reporte(self) -> dict:
        """Reporte mock"""
        return {"test": "ok"}


# ========================================
# TESTS
# ========================================

class TestBackendBase(unittest.TestCase):
    """Tests para BackendBase"""
    
    def setUp(self):
        """Configuración antes de cada test"""
        self.callback_mensaje = Mock()
        self.callback_progreso = Mock()
        self.callback_estado = Mock()
        
        self.backend = MockBackend(
            callback_mensaje=self.callback_mensaje,
            callback_progreso=self.callback_progreso,
            callback_estado=self.callback_estado
        )
    
    def tearDown(self):
        """Limpieza después de cada test"""
        pass
    
    # ========================================
    # TESTS DE INICIALIZACIÓN
    # ========================================
    
    def test_inicializacion_con_callbacks(self):
        """Test: Backend se inicializa correctamente con callbacks"""
        self.assertIsNotNone(self.backend.callback_mensaje)
        self.assertIsNotNone(self.backend.callback_progreso)
        self.assertIsNotNone(self.backend.callback_estado)
        self.assertEqual(self.backend.estado_actual, EstadoProceso.DETENIDO)
    
    def test_inicializacion_sin_callbacks(self):
        """Test: Backend funciona sin callbacks"""
        backend = MockBackend()
        self.assertIsNotNone(backend.callback_mensaje)
        # No debe lanzar error al llamar callback default
        backend._enviar_mensaje(FaseProceso.INICIAL, NivelMensaje.INFO, "test")
    
    # ========================================
    # TESTS DE COMUNICACIÓN
    # ========================================
    
    def test_enviar_mensaje(self):
        """Test: _enviar_mensaje llama al callback correctamente"""
        self.backend._enviar_mensaje(
            FaseProceso.INICIAL,
            NivelMensaje.INFO,
            "Test mensaje"
        )
        
        self.callback_mensaje.assert_called_once_with(
            FaseProceso.INICIAL,
            NivelMensaje.INFO,
            "Test mensaje"
        )
    
    def test_actualizar_progreso(self):
        """Test: _actualizar_progreso calcula porcentaje correctamente"""
        self.backend._actualizar_progreso(50, 100)
        
        self.callback_progreso.assert_called_once_with(50, 100, 50.0)
    
    def test_actualizar_progreso_division_cero(self):
        """Test: _actualizar_progreso maneja total=0"""
        self.backend._actualizar_progreso(0, 0)
        
        self.callback_progreso.assert_called_once_with(0, 0, 0.0)
    
    def test_cambiar_estado(self):
        """Test: _cambiar_estado actualiza estado y notifica"""
        self.backend._cambiar_estado(EstadoProceso.EN_EJECUCION)
        
        self.assertEqual(self.backend.estado_actual, EstadoProceso.EN_EJECUCION)
        self.callback_estado.assert_called_once_with(EstadoProceso.EN_EJECUCION)
    
    def test_cambiar_fase(self):
        """Test: _cambiar_fase actualiza fase y envía mensaje"""
        self.backend._cambiar_fase(FaseProceso.PROCESANDO)
        
        self.assertEqual(self.backend.fase_actual, FaseProceso.PROCESANDO)
        self.callback_mensaje.assert_called_once()
    
    # ========================================
    # TESTS DE CONTROL DE FLUJO
    # ========================================
    
    def test_pausar(self):
        """Test: pausar cambia estado correctamente"""
        self.backend._cambiar_estado(EstadoProceso.EN_EJECUCION)
        self.backend.pausar()
        
        self.assertEqual(self.backend.estado_actual, EstadoProceso.PAUSADO)
        self.assertFalse(self.backend._event_pausa.is_set())
    
    def test_pausar_solo_si_en_ejecucion(self):
        """Test: pausar solo funciona si está EN_EJECUCION"""
        self.backend._cambiar_estado(EstadoProceso.DETENIDO)
        self.backend.pausar()
        
        # No debe cambiar a PAUSADO
        self.assertEqual(self.backend.estado_actual, EstadoProceso.DETENIDO)
    
    def test_reanudar(self):
        """Test: reanudar cambia estado correctamente"""
        self.backend._cambiar_estado(EstadoProceso.EN_EJECUCION)
        self.backend.pausar()
        self.backend.reanudar()
        
        self.assertEqual(self.backend.estado_actual, EstadoProceso.EN_EJECUCION)
        self.assertTrue(self.backend._event_pausa.is_set())
    
    def test_cancelar(self):
        """Test: cancelar cambia estado y establece flag"""
        self.backend.cancelar()
        
        self.assertEqual(self.backend.estado_actual, EstadoProceso.CANCELADO)
        self.assertTrue(self.backend._event_cancelar.is_set())
        self.assertTrue(self.backend._event_pausa.is_set())
    
    def test_verificar_cancelacion_lanza_excepcion(self):
        """Test: _verificar_cancelacion lanza InterruptedError si cancelado"""
        self.backend.cancelar()
        
        with self.assertRaises(InterruptedError):
            self.backend._verificar_cancelacion()
    
    # ========================================
    # TESTS DE UTILIDADES
    # ========================================
    
    def test_manejar_nombre_duplicado_sin_duplicado(self):
        """Test: _manejar_nombre_duplicado devuelve mismo path si no existe"""
        with tempfile.TemporaryDirectory() as tmpdir:
            ruta = Path(tmpdir) / "test.txt"
            resultado = self.backend._manejar_nombre_duplicado(ruta)
            
            self.assertEqual(ruta, resultado)
    
    def test_manejar_nombre_duplicado_con_duplicado(self):
        """Test: _manejar_nombre_duplicado agrega sufijo si existe"""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Crear archivo original
            ruta = Path(tmpdir) / "test.txt"
            ruta.touch()
            
            # Manejar duplicado
            resultado = self.backend._manejar_nombre_duplicado(ruta)
            
            self.assertEqual(resultado, Path(tmpdir) / "test_1.txt")
            self.assertNotEqual(ruta, resultado)
    
    def test_crear_carpeta_segura(self):
        """Test: _crear_carpeta_segura crea carpeta correctamente"""
        with tempfile.TemporaryDirectory() as tmpdir:
            nueva_carpeta = Path(tmpdir) / "test_carpeta"
            
            resultado = self.backend._crear_carpeta_segura(nueva_carpeta)
            
            self.assertTrue(resultado)
            self.assertTrue(nueva_carpeta.exists())
            self.assertTrue(nueva_carpeta.is_dir())
    
    def test_verificar_permisos_escritura(self):
        """Test: _verificar_permisos_escritura detecta permisos correctamente"""
        with tempfile.TemporaryDirectory() as tmpdir:
            ruta = Path(tmpdir)
            
            resultado = self.backend._verificar_permisos_escritura(ruta)
            
            self.assertTrue(resultado)
    
    # ========================================
    # TESTS DE VALIDACIÓN
    # ========================================
    
    def test_validar_carpeta_existe_valida(self):
        """Test: _validar_carpeta_existe con carpeta válida"""
        with tempfile.TemporaryDirectory() as tmpdir:
            es_valido, mensaje = self.backend._validar_carpeta_existe(tmpdir)
            
            self.assertTrue(es_valido)
            self.assertEqual(mensaje, "")
    
    def test_validar_carpeta_existe_vacia(self):
        """Test: _validar_carpeta_existe con string vacío"""
        es_valido, mensaje = self.backend._validar_carpeta_existe("")
        
        self.assertFalse(es_valido)
        self.assertIn("seleccionar", mensaje.lower())
    
    def test_validar_carpeta_existe_no_existe(self):
        """Test: _validar_carpeta_existe con carpeta inexistente"""
        es_valido, mensaje = self.backend._validar_carpeta_existe("/ruta/inexistente/123")
        
        self.assertFalse(es_valido)
        self.assertIn("no existe", mensaje.lower())
    
    def test_validar_rango_fechas_valido(self):
        """Test: _validar_rango_fechas con rango válido"""
        fecha_inicio = datetime(2025, 1, 1)
        fecha_fin = datetime(2025, 12, 31)
        
        es_valido, mensaje = self.backend._validar_rango_fechas(fecha_inicio, fecha_fin)
        
        self.assertTrue(es_valido)
        self.assertEqual(mensaje, "")
    
    def test_validar_rango_fechas_invalido(self):
        """Test: _validar_rango_fechas con inicio > fin"""
        fecha_inicio = datetime(2025, 12, 31)
        fecha_fin = datetime(2025, 1, 1)
        
        es_valido, mensaje = self.backend._validar_rango_fechas(fecha_inicio, fecha_fin)
        
        self.assertFalse(es_valido)
        self.assertIn("posterior", mensaje.lower())
    
    # ========================================
    # TESTS DE LOGGING
    # ========================================
    
    def test_crear_log_archivo(self):
        """Test: _crear_log_archivo crea archivo correctamente"""
        with tempfile.TemporaryDirectory() as tmpdir:
            self.backend._crear_log_archivo(tmpdir, "test_log")
            
            self.assertIsNotNone(self.backend.log_file)
            self.assertTrue(self.backend.log_file.exists())
            self.assertTrue(str(self.backend.log_file).startswith(str(tmpdir)))
    
    def test_escribir_log(self):
        """Test: _escribir_log escribe mensaje correctamente"""
        with tempfile.TemporaryDirectory() as tmpdir:
            self.backend._crear_log_archivo(tmpdir, "test_log")
            self.backend._escribir_log("Test mensaje")
            
            with open(self.backend.log_file, 'r', encoding='utf-8') as f:
                contenido = f.read()
            
            self.assertIn("Test mensaje", contenido)
    
    # ========================================
    # TESTS DE FLUJO COMPLETO
    # ========================================
    
    def test_ejecutar_flujo_completo_exitoso(self):
        """Test: ejecutar completa flujo exitoso"""
        resultado = self.backend.ejecutar(10)
        
        self.assertTrue(self.backend.validar_llamado)
        self.assertTrue(self.backend.procesar_llamado)
        self.assertEqual(resultado["procesados"], 10)
        self.assertIn("tiempo_total", resultado)
    
    def test_ejecutar_con_parametros_invalidos(self):
        """Test: ejecutar lanza ValueError con parámetros inválidos"""
        with self.assertRaises(ValueError):
            self.backend.ejecutar(-1)
    
    def test_ejecutar_con_cancelacion(self):
        """Test: ejecutar maneja cancelación correctamente"""
        def procesar_con_cancelacion(valor):
            self.backend.cancelar()
            self.backend._verificar_cancelacion()
            return {"test": "ok"}
        
        self.backend._procesar_principal = procesar_con_cancelacion
        
        resultado = self.backend.ejecutar(10)
        self.assertEqual(self.backend.estado_actual, EstadoProceso.CANCELADO)
    
    def test_resetear_control(self):
        """Test: _resetear_control limpia eventos correctamente"""
        self.backend._event_cancelar.set()
        self.backend._event_pausa.clear()
        
        self.backend._resetear_control()
        
        self.assertFalse(self.backend._event_cancelar.is_set())
        self.assertTrue(self.backend._event_pausa.is_set())
    
    # ========================================
    # TESTS DE REPRESENTACIÓN
    # ========================================
    
    def test_repr(self):
        """Test: __repr__ devuelve representación correcta"""
        repr_str = repr(self.backend)
        
        self.assertIn("MockBackend", repr_str)
        self.assertIn("estado=", repr_str)
        self.assertIn("fase=", repr_str)


# ========================================
# TESTS DE ESTADISTICAS BASE
# ========================================

class TestEstadisticasBase(unittest.TestCase):
    """Tests para EstadisticasBase"""
    
    def test_tiempo_total_calcula_correctamente(self):
        """Test: tiempo_total calcula diferencia correctamente"""
        stats = EstadisticasBase()
        stats.tiempo_inicio = datetime(2025, 1, 1, 10, 0, 0)
        stats.tiempo_fin = datetime(2025, 1, 1, 10, 1, 30)
        
        self.assertEqual(stats.tiempo_total, 90.0)
    
    def test_tiempo_total_sin_tiempos(self):
        """Test: tiempo_total devuelve 0 si no hay tiempos"""
        stats = EstadisticasBase()
        
        self.assertEqual(stats.tiempo_total, 0.0)


# ========================================
# RUNNER
# ========================================

if __name__ == '__main__':
    unittest.main(verbosity=2)