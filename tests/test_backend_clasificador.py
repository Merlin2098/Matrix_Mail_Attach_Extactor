"""
Tests unitarios para ClasificadorDocumentos

FASE 2 - Test del clasificador heredando de BackendBase
"""

import unittest
from unittest.mock import Mock, patch
from pathlib import Path
from datetime import datetime
import tempfile
import shutil
import sys

# Agregar el directorio raíz y legacy al path
sys.path.insert(0, str(Path(__file__).parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "legacy"))

from legacy.backend_clasificador import (
    ClasificadorDocumentos,
    FaseProceso,
    NivelMensaje,
    EstadoProceso,  # Ahora se re-exporta desde backend_clasificador
    EstadisticasClasificacion
)


class TestClasificadorDocumentos(unittest.TestCase):
    """Tests para ClasificadorDocumentos"""
    
    def setUp(self):
        """Configuración antes de cada test"""
        self.callback_mensaje = Mock()
        self.callback_progreso = Mock()
        self.callback_estado = Mock()
        
        self.clasificador = ClasificadorDocumentos(
            callback_mensaje=self.callback_mensaje,
            callback_progreso=self.callback_progreso,
            callback_estado=self.callback_estado
        )
    
    # ========================================
    # TESTS DE INICIALIZACIÓN
    # ========================================
    
    def test_inicializacion(self):
        """Test: Clasificador se inicializa correctamente"""
        self.assertIsNotNone(self.clasificador)
        self.assertIsInstance(self.clasificador.estadisticas, EstadisticasClasificacion)
        self.assertEqual(self.clasificador.estado_actual, EstadoProceso.DETENIDO)
        self.assertFalse(self.clasificador.cancelado)
    
    # ========================================
    # TESTS DE VALIDACIÓN
    # ========================================
    
    def test_validar_parametros_validos(self):
        """Test: validar_parametros acepta carpeta válida"""
        with tempfile.TemporaryDirectory() as tmpdir:
            es_valido, mensaje = self.clasificador.validar_parametros(tmpdir)
            
            self.assertTrue(es_valido)
            self.assertEqual(mensaje, "")
    
    def test_validar_parametros_carpeta_vacia(self):
        """Test: validar_parametros rechaza carpeta vacía"""
        es_valido, mensaje = self.clasificador.validar_parametros("")
        
        self.assertFalse(es_valido)
        self.assertIn("seleccionar", mensaje.lower())
    
    def test_validar_parametros_carpeta_no_existe(self):
        """Test: validar_parametros rechaza carpeta inexistente"""
        es_valido, mensaje = self.clasificador.validar_parametros("/ruta/inexistente/123")
        
        self.assertFalse(es_valido)
        self.assertIn("no existe", mensaje.lower())
    
    def test_validar_parametros_no_es_carpeta(self):
        """Test: validar_parametros rechaza si no es carpeta"""
        with tempfile.NamedTemporaryFile() as tmpfile:
            es_valido, mensaje = self.clasificador.validar_parametros(tmpfile.name)
            
            self.assertFalse(es_valido)
            self.assertIn("carpeta", mensaje.lower())
    
    # ========================================
    # TESTS DE ESTADÍSTICAS
    # ========================================
    
    def test_estadisticas_iniciales(self):
        """Test: Estadísticas inician en 0"""
        stats = EstadisticasClasificacion()
        
        self.assertEqual(stats.total, 0)
        self.assertEqual(stats.firmados, 0)
        self.assertEqual(stats.sin_firmar, 0)
        self.assertEqual(stats.omitidos, 0)
        self.assertEqual(stats.errores, 0)
    
    # ========================================
    # TESTS DE GENERACIÓN DE REPORTE
    # ========================================
    
    def test_generar_reporte(self):
        """Test: _generar_reporte devuelve dict con claves correctas"""
        self.clasificador.estadisticas.total = 10
        self.clasificador.estadisticas.firmados = 5
        self.clasificador.estadisticas.sin_firmar = 3
        
        reporte = self.clasificador._generar_reporte()
        
        self.assertIn("total", reporte)
        self.assertIn("firmados", reporte)
        self.assertIn("sin_firmar", reporte)
        self.assertIn("omitidos", reporte)
        self.assertIn("errores", reporte)
        self.assertIn("tiempo_total", reporte)
        
        self.assertEqual(reporte["total"], 10)
        self.assertEqual(reporte["firmados"], 5)
        self.assertEqual(reporte["sin_firmar"], 3)
    
    # ========================================
    # TESTS DE CLASIFICACIÓN
    # ========================================
    
    def test_clasificar_archivo_firmado(self):
        """Test: _clasificar_archivo detecta documento firmado"""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Crear archivo con nombre "firmado"
            archivo = Path(tmpdir) / "documento_firmado.pdf"
            archivo.touch()
            
            carpeta_firmados = Path(tmpdir) / "Firmados"
            carpeta_sin_firmar = Path(tmpdir) / "Sin_Firmar"
            carpeta_firmados.mkdir()
            carpeta_sin_firmar.mkdir()
            
            resultado = self.clasificador._clasificar_archivo(
                archivo, carpeta_firmados, carpeta_sin_firmar
            )
            
            self.assertEqual(resultado, 'firmado')
            self.assertEqual(self.clasificador.estadisticas.firmados, 1)
            self.assertTrue((carpeta_firmados / "documento_firmado.pdf").exists())
    
    def test_clasificar_archivo_sin_firmar(self):
        """Test: _clasificar_archivo detecta documento sin firmar"""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Crear archivo con nombre "sin firmar"
            archivo = Path(tmpdir) / "documento_sin_firmar.pdf"
            archivo.touch()
            
            carpeta_firmados = Path(tmpdir) / "Firmados"
            carpeta_sin_firmar = Path(tmpdir) / "Sin_Firmar"
            carpeta_firmados.mkdir()
            carpeta_sin_firmar.mkdir()
            
            resultado = self.clasificador._clasificar_archivo(
                archivo, carpeta_firmados, carpeta_sin_firmar
            )
            
            self.assertEqual(resultado, 'sin_firmar')
            self.assertEqual(self.clasificador.estadisticas.sin_firmar, 1)
            self.assertTrue((carpeta_sin_firmar / "documento_sin_firmar.pdf").exists())
    
    def test_clasificar_archivo_omitido(self):
        """Test: _clasificar_archivo omite documento sin criterio"""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Crear archivo con nombre genérico
            archivo = Path(tmpdir) / "documento.pdf"
            archivo.touch()
            
            carpeta_firmados = Path(tmpdir) / "Firmados"
            carpeta_sin_firmar = Path(tmpdir) / "Sin_Firmar"
            carpeta_firmados.mkdir()
            carpeta_sin_firmar.mkdir()
            
            resultado = self.clasificador._clasificar_archivo(
                archivo, carpeta_firmados, carpeta_sin_firmar
            )
            
            self.assertEqual(resultado, 'omitido')
            self.assertEqual(self.clasificador.estadisticas.omitidos, 1)
            self.assertTrue(archivo.exists())  # No fue movido
    
    def test_clasificar_prioridad_sin_firmar(self):
        """Test: _clasificar_archivo prioriza 'sin firmar' sobre 'firmado'"""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Crear archivo con ambas palabras (sin firmar tiene prioridad)
            archivo = Path(tmpdir) / "documento_firmado_sin_firmar.pdf"
            archivo.touch()
            
            carpeta_firmados = Path(tmpdir) / "Firmados"
            carpeta_sin_firmar = Path(tmpdir) / "Sin_Firmar"
            carpeta_firmados.mkdir()
            carpeta_sin_firmar.mkdir()
            
            resultado = self.clasificador._clasificar_archivo(
                archivo, carpeta_firmados, carpeta_sin_firmar
            )
            
            self.assertEqual(resultado, 'sin_firmar')
            self.assertTrue((carpeta_sin_firmar / "documento_firmado_sin_firmar.pdf").exists())
    
    def test_clasificar_archivo_ingles(self):
        """Test: _clasificar_archivo detecta términos en inglés"""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Crear archivos con términos en inglés
            archivo_signed = Path(tmpdir) / "document_signed.pdf"
            archivo_not_signed = Path(tmpdir) / "document_not_signed.pdf"
            archivo_signed.touch()
            archivo_not_signed.touch()
            
            carpeta_firmados = Path(tmpdir) / "Firmados"
            carpeta_sin_firmar = Path(tmpdir) / "Sin_Firmar"
            carpeta_firmados.mkdir()
            carpeta_sin_firmar.mkdir()
            
            # Test "signed"
            resultado1 = self.clasificador._clasificar_archivo(
                archivo_signed, carpeta_firmados, carpeta_sin_firmar
            )
            self.assertEqual(resultado1, 'firmado')
            
            # Resetear estadísticas para segundo test
            self.clasificador.estadisticas = EstadisticasClasificacion()
            
            # Test "not signed"
            resultado2 = self.clasificador._clasificar_archivo(
                archivo_not_signed, carpeta_firmados, carpeta_sin_firmar
            )
            self.assertEqual(resultado2, 'sin_firmar')
    
    # ========================================
    # TESTS DE CONTROL DE FLUJO
    # ========================================
    
    def test_cancelar(self):
        """Test: cancelar establece flag y cambia estado"""
        self.clasificador.cancelar()
        
        self.assertTrue(self.clasificador.cancelado)
        self.assertEqual(self.clasificador.estado_actual, EstadoProceso.CANCELADO)
    
    def test_pausar_y_reanudar(self):
        """Test: pausar y reanudar funcionan (heredado de BackendBase)"""
        self.clasificador._cambiar_estado(EstadoProceso.EN_EJECUCION)
        
        self.clasificador.pausar()
        self.assertEqual(self.clasificador.estado_actual, EstadoProceso.PAUSADO)
        
        self.clasificador.reanudar()
        self.assertEqual(self.clasificador.estado_actual, EstadoProceso.EN_EJECUCION)
    
    # ========================================
    # TESTS DE INTEGRACIÓN
    # ========================================
    
    def test_clasificar_carpeta_completa(self):
        """Test: clasificar procesa carpeta completa correctamente"""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Crear archivos de prueba
            (Path(tmpdir) / "doc1_firmado.pdf").touch()
            (Path(tmpdir) / "doc2_sin_firmar.pdf").touch()
            (Path(tmpdir) / "doc3_firmado.pdf").touch()
            (Path(tmpdir) / "doc4.pdf").touch()  # Omitido
            
            resultado = self.clasificador.clasificar(tmpdir)
            
            # Verificar estadísticas
            self.assertEqual(resultado["total"], 4)
            self.assertEqual(resultado["firmados"], 2)
            self.assertEqual(resultado["sin_firmar"], 1)
            self.assertEqual(resultado["omitidos"], 1)
            self.assertEqual(resultado["errores"], 0)
            
            # Verificar que se crearon las carpetas
            self.assertTrue((Path(tmpdir) / "Documentos Firmados").exists())
            self.assertTrue((Path(tmpdir) / "Documentos sin Firmar").exists())
    
    def test_clasificar_carpeta_vacia(self):
        """Test: clasificar maneja carpeta vacía correctamente"""
        with tempfile.TemporaryDirectory() as tmpdir:
            resultado = self.clasificador.clasificar(tmpdir)
            
            self.assertEqual(resultado["total"], 0)
            self.callback_mensaje.assert_called()
    
    def test_clasificar_con_parametros_invalidos(self):
        """Test: clasificar lanza ValueError con parámetros inválidos"""
        with self.assertRaises(ValueError):
            self.clasificador.clasificar("")
    
    # ========================================
    # TESTS DE MANEJO DE ERRORES
    # ========================================
    
    def test_clasificar_archivo_error_permiso(self):
        """Test: _clasificar_archivo maneja PermissionError"""
        with tempfile.TemporaryDirectory() as tmpdir:
            archivo = Path(tmpdir) / "documento_firmado.pdf"
            archivo.touch()
            
            carpeta_firmados = Path(tmpdir) / "Firmados"
            carpeta_sin_firmar = Path(tmpdir) / "Sin_Firmar"
            carpeta_firmados.mkdir()
            carpeta_sin_firmar.mkdir()
            
            # Simular PermissionError
            with patch('shutil.move', side_effect=PermissionError("Archivo bloqueado")):
                resultado = self.clasificador._clasificar_archivo(
                    archivo, carpeta_firmados, carpeta_sin_firmar
                )
            
            self.assertEqual(resultado, 'error')
            self.assertEqual(self.clasificador.estadisticas.errores, 1)
    
    # ========================================
    # TESTS DE UTILIDADES HEREDADAS
    # ========================================
    
    def test_crear_carpeta_segura(self):
        """Test: _crear_carpeta_segura funciona (heredado de BackendBase)"""
        with tempfile.TemporaryDirectory() as tmpdir:
            nueva_carpeta = Path(tmpdir) / "test_carpeta"
            
            resultado = self.clasificador._crear_carpeta_segura(nueva_carpeta)
            
            self.assertTrue(resultado)
            self.assertTrue(nueva_carpeta.exists())


# ========================================
# RUNNER
# ========================================

if __name__ == '__main__':
    unittest.main(verbosity=2)