# 📧 Gestión de Correos Outlook - MatrixMAE

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![PyQt5](https://img.shields.io/badge/PyQt5-5.15+-green.svg)](https://pypi.org/project/PyQt5/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

## 📋 Descripción

**MatrixMAE** es una aplicación de escritorio desarrollada en Python con PyQt5 que automatiza la gestión de correos electrónicos de Microsoft Outlook. Permite extraer adjuntos de manera masiva y clasificar documentos según su estado de firma, optimizando significativamente los flujos de trabajo empresariales.

### 🎯 Características Principales

- **📥 Extracción Masiva de Adjuntos**
  - Filtrado por frases clave, rangos de fecha y carpetas específicas
  - Detección automática del rango de fechas disponible en la bandeja
  - Sistema anti-duplicados inteligente
  - Generación de reportes en Excel con metadata completa
  - Log detallado de todas las operaciones

- **📁 Clasificación de Documentos**
  - Organización automática según estado de firma
  - Detección de patrones: `firmado`, `signed`, `sin_firmar`, `not_signed`
  - Estadísticas en tiempo real del proceso
  - Manejo seguro de archivos duplicados

- **🎨 Interfaz Moderna y Amigable**
  - Tema claro/oscuro configurable
  - Indicadores de progreso detallados por fase
  - Logs en tiempo real con códigos de color
  - Notificaciones visuales y sonoras al completar tareas
  - Selector inteligente de carpetas con lazy loading

## 🚀 Instalación

### Requisitos Previos

- Windows 10/11
- Python 3.8 o superior
- Microsoft Outlook instalado y configurado
- Permisos de administrador (recomendado)

### Instalación desde Código Fuente

1. **Clonar el repositorio:**
```bash
git clone https://github.com/Merlin2098/Matrix_Mail_Attach_Extactor.git
cd Matrix_Mail_Attach_Extactor
```

2. **Crear entorno virtual:**
```bash
python -m venv venv
venv\Scripts\activate
```

3. **Instalar dependencias:**
```bash
pip install -r requirements.txt
```

4. **Ejecutar la aplicación:**
```bash
python legacy/front_main.py
```

## 🛠️ Desarrollo

### Estructura del Proyecto

```
proyecto/
├── config/                    # Configuración y recursos
│   ├── config_manager.py     # Gestor singleton de configuración
│   ├── config.json            # Configuración persistente
│   └── ico.ico                # Icono de la aplicación
│
├── legacy/                    # Módulos principales
│   ├── front_main.py         # Interfaz gráfica PyQt5
│   ├── backend_base.py       # Clase base abstracta
│   ├── backend_extractor.py  # Lógica de extracción
│   ├── backend_clasificador.py # Lógica de clasificación
│   ├── extractor_adapter.py  # Worker para threading
│   ├── clasificador_adapter.py # Worker para threading
│   └── logs/                 # Logs generados (auto-creada)
│
├── ui/                        # Componentes de interfaz
│   ├── __init__.py           
│   └── estilos.py            # Estilos CSS para temas
│
├── tests/                     # Tests unitarios
│   ├── test_backend_base.py
│   ├── test_backend_extractor.py
│   └── test_backend_clasificador.py
│
├── venv/                      # Entorno virtual (ignorado en git)
├── 1.generar_onedir.py       # Script para generar ejecutable
├── requirements.txt           # Dependencias del proyecto
├── README.md                  # Este archivo
└── .gitignore                # Archivos ignorados por git
```

### Arquitectura

El proyecto implementa un patrón de arquitectura en 3 capas:

1. **Capa de Presentación** (`front_main.py`): Interfaz gráfica PyQt5
2. **Capa de Adaptación** (`*_adapter.py`): Workers para threading y señales
3. **Capa de Lógica** (`backend_*.py`): Procesamiento y reglas de negocio

```
┌─────────────────────────────────────────┐
│         BackendBase (Abstracta)         │
│  • Callbacks unificados                 │
│  • Control de estados                   │
│  • Utilidades comunes                   │
└─────────────────────────────────────────┘
              ▲                    ▲
              │                    │
    ┌─────────┴────────┐  ┌───────┴─────────┐
    │ ExtractorOutlook │  │  Clasificador   │
    └──────────────────┘  └─────────────────┘
              ▲                    ▲
              │                    │
    ┌─────────┴────────┐  ┌───────┴─────────┐
    │ ExtractorAdapter │  │ ClasificadorAdpt│
    └──────────────────┘  └─────────────────┘
              ▲                    ▲
              └────────────────────┘
                        │
              ┌─────────┴─────────┐
              │   front_main.py   │
              └───────────────────┘
```

### Generar Ejecutable

Para crear un ejecutable distribuible:

1. **Activar entorno virtual:**
```bash
venv\Scripts\activate
```

2. **Ejecutar script de generación:**
```bash
python 1.generar_onedir.py
```

3. **Distribución:**
   - El ejecutable se generará en `dist/MatrixMAE/`
   - Distribuir la carpeta completa, no solo el .exe
   - El archivo `config.json` puede editarse después de la distribución

## 🧪 Testing

Ejecutar todos los tests:
```bash
pytest tests/ -v
```

Ejecutar con cobertura:
```bash
pytest tests/ --cov=legacy --cov-report=html
```

### Tests Disponibles
- **33 tests** para `BackendBase`
- **17 tests** para `ExtractorAdjuntosOutlook`
- **19 tests** para `ClasificadorDocumentos`

## 🔧 Configuración

### config.json

```json
{
  "tema": "claro",
  "ui": {
    "splash_duration": 2000,
    "window_size": [1200, 700]
  },
  "extractor": {
    "max_intentos": 3,
    "timeout": 30
  },
  "clasificador": {
    "patrones_firmado": ["firmado", "signed", "firm"],
    "crear_subcarpetas": true
  }
}
```

### Variables de Entorno

No se requieren variables de entorno específicas. La aplicación detecta automáticamente las rutas necesarias.

## 📝 Uso

### Extracción de Adjuntos

1. Abrir la pestaña **"Descarga de Adjuntos"**
2. Seleccionar carpeta de Outlook con el botón **"📧 Explorar"**
3. Configurar:
   - **Frases de búsqueda** (separadas por coma)
   - **Rango de fechas** (inicio y fin)
   - **Carpeta de destino** para guardar adjuntos
4. Hacer clic en **"▶️ Iniciar Descarga"**
5. Monitorear el progreso en las áreas de log

### Clasificación de Documentos

1. Abrir la pestaña **"Clasificar Documentos"**
2. Seleccionar **carpeta origen** con documentos a clasificar
3. Seleccionar **carpeta destino** para documentos organizados
4. Hacer clic en **"▶️ Iniciar Clasificación"**
5. Revisar estadísticas en tiempo real

## 🐛 Solución de Problemas

### Problema: "No se puede conectar a Outlook"
**Solución:** 
- Verificar que Outlook esté instalado y configurado
- Ejecutar la aplicación como administrador
- Asegurarse de que Outlook no esté ejecutándose en modo seguro

### Problema: "Warning de High DPI"
**Estado:** Warning conocido que no afecta la funcionalidad
**Nota:** La interfaz funciona perfectamente a pesar del warning

### Problema: "No se encuentran correos en el rango especificado"
**Solución:**
- Verificar el rango real disponible mostrado en los logs
- Ajustar las fechas según lo disponible en la bandeja
- Revisar que las frases de búsqueda sean correctas

## 👥 Contribuir

Las contribuciones son bienvenidas. Por favor:

1. Fork el proyecto
2. Crear una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abrir un Pull Request

### Guías de Estilo

- Seguir PEP 8 para código Python
- Documentar funciones con docstrings
- Mantener cobertura de tests > 80%
- Actualizar README.md con cambios significativos

## 📄 Licencia

Este proyecto está licenciado bajo la Licencia MIT - ver el archivo [LICENSE]([LICENSE](https://github.com/Merlin2098/Matrix_Mail_Attach_Extactor/blob/main/LICENSE)) para más detalles.


---

**Última actualización:** Noviembre 2025 | **Versión:** 2.0.0
