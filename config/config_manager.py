# config_manager.py
"""
Gestor centralizado de configuración para la aplicación.
Maneja rutas dinámicas para funcionar tanto en desarrollo como en ejecutable.

UBICACIÓN: Este archivo debe estar en config/config_manager.py
"""

import sys
import json
from pathlib import Path
from datetime import datetime


class ConfigManager:
    """
    Gestor singleton de configuración con búsqueda dinámica de archivos.
    
    Funciona tanto en desarrollo (python script) como en ejecutable (.exe)
    Busca config.json e icon.ico en la carpeta config/ relativa al ejecutable.
    """
    
    _instance = None
    
    # Nombres de archivos
    CONFIG_FILENAME = "config.json"
    ICON_FILENAME = "ico.ico"
    CONFIG_DIR_NAME = "config"
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        
        self._base_path = self._get_base_path()
        self.config_dir = self._base_path / self.CONFIG_DIR_NAME
        self.config_path = self.config_dir / self.CONFIG_FILENAME
        self.icon_path = self.config_dir / self.ICON_FILENAME
        
        # Crear directorio config si no existe
        self.config_dir.mkdir(exist_ok=True)
        
        self._config = self._load_config()
        self._initialized = True
    
    @staticmethod
    def _get_base_path() -> Path:
        """
        Obtiene la ruta base de la aplicación de forma dinámica.
        
        Returns:
            Path: Ruta base del proyecto (carpeta raíz)
            
        Comportamiento:
            - En .exe (PyInstaller): retorna directorio del ejecutable
            - En desarrollo: retorna directorio raíz del proyecto
            
        Nota: Como config_manager.py está en config/, necesitamos subir un nivel
        """
        if getattr(sys, 'frozen', False):
            # Ejecutando como .exe (PyInstaller/cx_Freeze)
            # sys.executable apunta al .exe, queremos su directorio padre
            return Path(sys.executable).parent
        else:
            # Ejecutando como script Python
            # Este archivo está en config/config_manager.py
            # Necesitamos subir un nivel para llegar a la raíz del proyecto
            return Path(__file__).parent.parent
    
    def _load_config(self) -> dict:
        """
        Carga la configuración desde config.json.
        Si no existe, crea una configuración por defecto.
        """
        if not self.config_path.exists():
            return self._create_default_config()
        
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            return config
        except json.JSONDecodeError as e:
            print(f"⚠️ Error al leer config.json: {e}")
            print("📝 Creando configuración por defecto...")
            return self._create_default_config()
        except Exception as e:
            print(f"❌ Error inesperado: {e}")
            return self._get_default_config()
    
    def _create_default_config(self) -> dict:
        """
        Crea un archivo de configuración por defecto.
        """
        default_config = self._get_default_config()
        self.save_config(default_config)
        return default_config
    
    @staticmethod
    def _get_default_config() -> dict:
        """
        Retorna la configuración por defecto.
        """
        return {
            "tema": "light",
            "ui": {
                "colors": {
                    "primary": "#16A085",
                    "primary_hover": "#138D75",
                    "primary_pressed": "#0E6655"
                },
                "window": {
                    "width": 1000,
                    "height": 700
                }
            },
            "app": {
                "beep_on_complete": True,
                "flash_taskbar": True
            },
            "ultima_actualizacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    
    def get(self, key: str, default=None):
        """
        Obtiene un valor de configuración usando notación de punto.
        
        Args:
            key: Clave en formato "seccion.subseccion.valor"
            default: Valor por defecto si no existe
            
        Returns:
            Valor de configuración o default
            
        Ejemplos:
            config.get('tema')  # 'light'
            config.get('ui.colors.primary')  # '#16A085'
        """
        keys = key.split('.')
        value = self._config
        
        for k in keys:
            if isinstance(value, dict):
                value = value.get(k)
            else:
                return default
            
            if value is None:
                return default
        
        return value
    
    def set(self, key: str, value):
        """
        Establece un valor de configuración usando notación de punto.
        
        Args:
            key: Clave en formato "seccion.subseccion.valor"
            value: Valor a establecer
            
        Ejemplos:
            config.set('tema', 'dark')
            config.set('ui.colors.primary', '#FF0000')
        """
        keys = key.split('.')
        config = self._config
        
        # Navegar hasta el penúltimo nivel
        for k in keys[:-1]:
            config = config.setdefault(k, {})
        
        # Establecer el valor final
        config[keys[-1]] = value
        self.save_config(self._config)
    
    def save_config(self, config: dict = None):
        """
        Guarda la configuración en disco.
        
        Args:
            config: Diccionario de configuración. Si es None, usa self._config
        """
        if config is None:
            config = self._config
        
        try:
            config["ultima_actualizacion"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            
        except Exception as e:
            print(f"❌ Error al guardar configuración: {e}")
    
    def get_icon_path(self) -> Path | None:
        """
        Obtiene la ruta del icono de la aplicación.
        
        Returns:
            Path del icono si existe, None en caso contrario
        """
        if self.icon_path.exists():
            return self.icon_path
        return None
    
    def get_tema(self) -> str:
        """
        Obtiene el tema actual de la aplicación.
        
        Returns:
            'light' o 'dark'
        """
        return self.get('tema', 'light')
    
    def set_tema(self, tema: str):
        """
        Establece el tema de la aplicación.
        
        Args:
            tema: 'light' o 'dark'
        """
        if tema in ['light', 'dark']:
            self.set('tema', tema)
        else:
            print(f"⚠️ Tema inválido: {tema}. Usar 'light' o 'dark'")
    
    def reload(self):
        """
        Recarga la configuración desde el archivo.
        Útil si el archivo fue modificado externamente.
        """
        self._config = self._load_config()
    
    def __repr__(self):
        return f"ConfigManager(config_path='{self.config_path}')"


# Para uso directo y testing
if __name__ == "__main__":
    config = ConfigManager()
    print("=" * 60)
    print("🔍 INFORMACIÓN DEL CONFIG MANAGER")
    print("=" * 60)
    print(f"📂 Ruta base del proyecto: {config._base_path}")
    print(f"📁 Directorio config: {config.config_dir}")
    print(f"📄 Config path: {config.config_path}")
    print(f"🎨 Icon path: {config.icon_path}")
    print()
    print("⚙️  CONFIGURACIÓN ACTUAL:")
    print(f"   Tema: {config.get_tema()}")
    print(f"   Color primario: {config.get('ui.colors.primary')}")
    print(f"   Beep al completar: {config.get('app.beep_on_complete')}")
    print()
    
    icon = config.get_icon_path()
    if icon:
        print(f"✅ Icono encontrado en: {icon}")
    else:
        print(f"⚠️  Icono NO encontrado. Esperado en: {config.icon_path}")
    
    print("=" * 60)