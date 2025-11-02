"""
Tests unitarios para ExtractorAdjuntosOutlook

FASE 2 - Test del extractor heredando de BackendBase
"""

import unittest
from unittest.mock import Mock, MagicMock, patch
from pathlib import Path
from datetime import datetime
import tempfile
import sys

# Agregar el directorio raíz y legacy al path
sys.path.insert(0, str(Path(__file__).parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "legacy"))

from legacy.backend_extractor import (
    ExtractorAdjuntosOutlook,
    FaseProceso,
    NivelMensaje,
    EstadoProceso,
    EstadisticasExtraccion
)


class TestExtractorAdjuntosOutlook(unittest.TestCase):
    """Tests para ExtractorAdjuntosOutlook"""
    
    def setUp(self):
        """Configuración antes de cada test"""
        self.callback_mensaje = Mock()
        self.callback_progreso = Mock()
        self.callback_estado = Mock()
        
        self.extractor = ExtractorAdjuntosOutlook(
            callback_mensaje=self.callback_mensaje,
            callback_progreso=self.callback_progreso,
            callback_estado=self.callback_estado
        )
    
    # ========================================
    # TESTS DE INICIALIZACIÓN
    # ========================================
    
    def test_inicializacion(self):
        """Test: Extractor se inicializa correctamente"""
        self.assertIsNotNone(self.extractor)
        self.assertIsInstance(self.extractor.estadisticas, EstadisticasExtraccion)
        self.assertEqual(self.extractor.estado_actual, EstadoProceso.DETENIDO)
    
    def test_configuracion_por_defecto(self):
        """Test: Extractor tiene configuración por defecto"""
        self.assertIn("min_lote", self.extractor.config)
        self.assertIn("max_lote", self.extractor.config)
        self.assertGreater(self.extractor.config["max_lote"], 0)
    
    # ========================================
    # TESTS DE VALIDACIÓN
    # ========================================
    
    def test_validar_parametros_validos(self):
        """Test: validar_parametros acepta parámetros válidos"""
        with tempfile.TemporaryDirectory() as tmpdir:
            frases = ["test"]
            fecha_inicio = datetime(2025, 1, 1)
            fecha_fin = datetime(2025, 12, 31)
            
            es_valido, mensaje = self.extractor.validar_parametros(
                frases, tmpdir, "Test\\Inbox", fecha_inicio, fecha_fin
            )
            
            self.assertTrue(es_valido)
            self.assertEqual(mensaje, "")
    
    def test_validar_parametros_destino_vacio(self):
        """Test: validar_parametros rechaza destino vacío"""
        fecha_inicio = datetime(2025, 1, 1)
        fecha_fin = datetime(2025, 12, 31)
        
        es_valido, mensaje = self.extractor.validar_parametros(
            ["test"], "", "Test\\Inbox", fecha_inicio, fecha_fin
        )
        
        self.assertFalse(es_valido)
        self.assertIn("destino", mensaje.lower())
    
    def test_validar_parametros_outlook_folder_vacio(self):
        """Test: validar_parametros rechaza outlook_folder vacío"""
        with tempfile.TemporaryDirectory() as tmpdir:
            fecha_inicio = datetime(2025, 1, 1)
            fecha_fin = datetime(2025, 12, 31)
            
            es_valido, mensaje = self.extractor.validar_parametros(
                ["test"], tmpdir, "", fecha_inicio, fecha_fin
            )
            
            self.assertFalse(es_valido)
            self.assertIn("bandeja", mensaje.lower())
    
    def test_validar_parametros_fechas_invalidas(self):
        """Test: validar_parametros rechaza fecha_inicio > fecha_fin"""
        with tempfile.TemporaryDirectory() as tmpdir:
            fecha_inicio = datetime(2025, 12, 31)
            fecha_fin = datetime(2025, 1, 1)
            
            es_valido, mensaje = self.extractor.validar_parametros(
                ["test"], tmpdir, "Test\\Inbox", fecha_inicio, fecha_fin
            )
            
            self.assertFalse(es_valido)
            self.assertIn("posterior", mensaje.lower())
    
    def test_validar_parametros_sin_frases_genera_warning(self):
        """Test: validar_parametros con frases vacías genera warning"""
        with tempfile.TemporaryDirectory() as tmpdir:
            fecha_inicio = datetime(2025, 1, 1)
            fecha_fin = datetime(2025, 12, 31)
            
            es_valido, mensaje = self.extractor.validar_parametros(
                [], tmpdir, "Test\\Inbox", fecha_inicio, fecha_fin
            )
            
            self.assertTrue(es_valido)
            # Debe haber enviado mensaje de warning
            self.callback_mensaje.assert_called()
    
    # ========================================
    # TESTS DE ESTADÍSTICAS
    # ========================================
    
    def test_estadisticas_iniciales(self):
        """Test: Estadísticas inician en 0"""
        stats = EstadisticasExtraccion()
        
        self.assertEqual(stats.total_correos, 0)
        self.assertEqual(stats.correos_procesados, 0)
        self.assertEqual(stats.adjuntos_descargados, 0)
        self.assertEqual(stats.adjuntos_fallidos, 0)
        self.assertEqual(stats.tamaño_total_mb, 0.0)
    
    def test_tasa_exito_calculo(self):
        """Test: tasa_exito calcula porcentaje correctamente"""
        stats = EstadisticasExtraccion()
        stats.adjuntos_descargados = 80
        stats.adjuntos_fallidos = 20
        
        self.assertEqual(stats.tasa_exito, 80.0)
    
    def test_tasa_exito_sin_adjuntos(self):
        """Test: tasa_exito devuelve 0 sin adjuntos"""
        stats = EstadisticasExtraccion()
        
        self.assertEqual(stats.tasa_exito, 0.0)
    
    # ========================================
    # TESTS DE GENERACIÓN DE REPORTE
    # ========================================
    
    def test_generar_reporte(self):
        """Test: _generar_reporte devuelve dict con claves correctas"""
        self.extractor.estadisticas.total_correos = 10
        self.extractor.estadisticas.adjuntos_descargados = 5
        
        reporte = self.extractor._generar_reporte()
        
        self.assertIn("total_correos", reporte)
        self.assertIn("correos_procesados", reporte)
        self.assertIn("adjuntos_descargados", reporte)
        self.assertIn("adjuntos_fallidos", reporte)
        self.assertIn("tamaño_total_mb", reporte)
        self.assertIn("tiempo_total", reporte)
        self.assertIn("tasa_exito", reporte)
    
    # ========================================
    # TESTS DE CONTROL DE FLUJO (heredados de BackendBase)
    # ========================================
    
    def test_pausar_y_reanudar(self):
        """Test: pausar y reanudar funcionan correctamente"""
        self.extractor._cambiar_estado(EstadoProceso.EN_EJECUCION)
        
        self.extractor.pausar()
        self.assertEqual(self.extractor.estado_actual, EstadoProceso.PAUSADO)
        
        self.extractor.reanudar()
        self.assertEqual(self.extractor.estado_actual, EstadoProceso.EN_EJECUCION)
    
    def test_cancelar(self):
        """Test: cancelar establece flag correctamente"""
        self.extractor.cancelar()
        
        self.assertEqual(self.extractor.estado_actual, EstadoProceso.CANCELADO)
        self.assertTrue(self.extractor._event_cancelar.is_set())
    
    # ========================================
    # TESTS DE UTILIDADES (heredadas de BackendBase)
    # ========================================
    
    def test_manejar_nombre_duplicado(self):
        """Test: _manejar_nombre_duplicado funciona heredado de BackendBase"""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Crear archivo
            ruta = Path(tmpdir) / "test.txt"
            ruta.touch()
            
            # Manejar duplicado
            resultado = self.extractor._manejar_nombre_duplicado(ruta)
            
            self.assertNotEqual(ruta, resultado)
            self.assertEqual(resultado.name, "test_1.txt")
    
    # ========================================
    # TESTS DE EXCEL (específico del extractor)
    # ========================================
    
    def test_generar_excel_sin_archivos(self):
        """Test: _generar_excel_listado no falla sin archivos"""
        with tempfile.TemporaryDirectory() as tmpdir:
            # No debe lanzar error
            self.extractor._generar_excel_listado(tmpdir)
            
            # Debe haber enviado mensaje informativo
            self.callback_mensaje.assert_called()
    
    def test_generar_excel_con_archivos(self):
        """Test: _generar_excel_listado genera archivo con datos"""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Simular archivos descargados
            self.extractor.estadisticas.archivos_descargados = [
                {
                    'nombre': 'test1.pdf',
                    'fecha_descarga': datetime.now(),
                    'fecha_correo': '01/01/2025',
                    'hora_correo': '10:00:00'
                },
                {
                    'nombre': 'test2.pdf',
                    'fecha_descarga': datetime.now(),
                    'fecha_correo': '02/01/2025',
                    'hora_correo': '11:00:00'
                }
            ]
            
            self.extractor._generar_excel_listado(tmpdir)
            
            # Verificar que se creó un archivo Excel
            archivos = list(Path(tmpdir).glob("*.xlsx"))
            self.assertEqual(len(archivos), 1)
            self.assertTrue(archivos[0].name.startswith("lista_documentos"))
    
    # ========================================
    # TESTS DE MÉTODO PÚBLICO
    # ========================================
    
    @patch('backend_extractor.pythoncom')
    @patch('backend_extractor.win32com.client')
    def test_extraer_adjuntos_valida_parametros_primero(self, mock_win32, mock_pythoncom):
        """Test: extraer_adjuntos valida parámetros antes de procesar"""
        fecha_inicio = datetime(2025, 12, 31)
        fecha_fin = datetime(2025, 1, 1)
        
        with self.assertRaises(ValueError):
            self.extractor.extraer_adjuntos(
                ["test"], "/tmp", "Test\\Inbox", fecha_inicio, fecha_fin
            )


# ========================================
# TESTS DE INTEGRACIÓN (mockeados)
# ========================================

class TestExtractorIntegracion(unittest.TestCase):
    """Tests de integración con mocks de Outlook"""
    
    @patch('backend_extractor.pythoncom')
    @patch('backend_extractor.win32com.client')
    def test_conectar_outlook_exitoso(self, mock_win32, mock_pythoncom):
        """Test: _conectar_outlook establece conexión exitosamente"""
        # Mock de Outlook
        mock_outlook = MagicMock()
        mock_namespace = MagicMock()
        mock_outlook.GetNamespace.return_value = mock_namespace
        mock_win32.Dispatch.return_value = mock_outlook
        
        extractor = ExtractorAdjuntosOutlook()
        namespace = extractor._conectar_outlook()
        
        self.assertIsNotNone(namespace)
        mock_win32.Dispatch.assert_called_once_with("Outlook.Application")
        mock_outlook.GetNamespace.assert_called_once_with("MAPI")
    
    @patch('backend_extractor.pythoncom')
    @patch('backend_extractor.win32com.client')
    def test_conectar_outlook_falla(self, mock_win32, mock_pythoncom):
        """Test: _conectar_outlook lanza ConnectionError si falla"""
        mock_win32.Dispatch.side_effect = Exception("Outlook no disponible")
        
        extractor = ExtractorAdjuntosOutlook()
        
        with self.assertRaises(ConnectionError):
            extractor._conectar_outlook()


# ========================================
# RUNNER
# ========================================

if __name__ == '__main__':
    unittest.main(verbosity=2)