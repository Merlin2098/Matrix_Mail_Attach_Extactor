import os
import sys
import pkg_resources
import subprocess
import shutil

# ==========================================================
# CONFIGURACIÓN
# ==========================================================
NOMBRE_EXE = "MatrixMAE.exe"
MAIN_SCRIPT = "legacy/front_main.py"

DIST_PATH = "dist"
BUILD_PATH = "build"
SPEC_PATH = "spec"

EXCLUSIONES = [
    "pip", "wheel", "setuptools", "pkg_resources",
    "distutils", "ensurepip", "test", "tkinter.test",
    "pytest", "pytest_cov", "coverage"
]

# ==========================================================
# VALIDAR ENTORNO VIRTUAL
# ==========================================================
def validar_entorno_virtual():
    print("=" * 60)
    print("🔍 VALIDACIÓN DE ENTORNO VIRTUAL")
    print("=" * 60)

    if sys.prefix == sys.base_prefix:
        print("❌ ERROR: No estás dentro de un entorno virtual (venv).")
        print("   Activa uno antes de continuar.")
        print("   Ejemplo: venv\\Scripts\\activate")
        sys.exit(1)

    print(f"✅ Entorno virtual detectado: {sys.prefix}\n")

    paquetes = sorted([(pkg.key, pkg.version) for pkg in pkg_resources.working_set])
    print(f"📦 Librerías instaladas ({len(paquetes)}):")
    for nombre, version in paquetes:
        flag = "🧹 (excluir)" if nombre in EXCLUSIONES else "✅"
        print(f"   {flag} {nombre:<20} {version}")
    print("\n")

# ==========================================================
# CONFIRMACIÓN MANUAL
# ==========================================================
def confirmar_ejecucion():
    print("=" * 60)
    print("⚠️  CONFIRMACIÓN DE EJECUCIÓN FINAL")
    print("=" * 60)
    respuesta = input("¿Deseas generar el ejecutable ahora? (S/N): ").strip().lower()

    if respuesta not in ("s", "si", "sí"):
        print("\n🛑 Proceso cancelado por el usuario.")
        sys.exit(0)

    print("\n✅ Confirmado. Continuando con la generación...\n")

# ==========================================================
# LIMPIAR BUILDS ANTERIORES
# ==========================================================
def limpiar_builds():
    for carpeta in [DIST_PATH, BUILD_PATH, SPEC_PATH]:
        if os.path.exists(carpeta):
            try:
                shutil.rmtree(carpeta)
                print(f"🧹 Limpiado: {carpeta}")
            except Exception as e:
                print(f"⚠️ No se pudo limpiar {carpeta}: {e}")
    print()

# ==========================================================
# CONSTRUIR COMANDO PYINSTALLER
# ==========================================================
def construir_comando():
    base_dir = os.getcwd()
    legacy_dir = os.path.join(base_dir, "legacy") # ⭐ Definir ruta a legacy

    comando = [
        sys.executable, "-m", "PyInstaller",
        "--onedir",
        "--windowed",
        "--clean",
        "--log-level", "WARN",
        "--distpath", DIST_PATH,
        "--workpath", BUILD_PATH,
        "--specpath", SPEC_PATH,
        "--name", NOMBRE_EXE.replace(".exe", ""),

        # ⭐ INICIO DE LA CORRECCIÓN ⭐
        
        # 1. Añadir el directorio RAÍZ (para encontrar 'config' y 'ui')
        "--paths", base_dir,
        
        # 2. Añadir el directorio LEGACY (para encontrar 'extractor_adapter', etc.)
        "--paths", legacy_dir,
        
        # ⭐ FIN DE LA CORRECCIÓN ⭐

        # Dependencias ocultas necesarias para Outlook y pywin32
        "--hidden-import=win32timezone",
        "--hidden-import=pythoncom",
        "--hidden-import=win32com",
        "--hidden-import=win32com.client",
        "--hidden-import=win32com.gen_py",
        # Imports adicionales para PyQt5
        "--hidden-import=PyQt5.QtCore",
        "--hidden-import=PyQt5.QtGui",
        "--hidden-import=PyQt5.QtWidgets",
    ]

    # Excluir módulos innecesarios
    for excl in EXCLUSIONES:
        comando += ["--exclude-module", excl]

    # ======================================================
    # 📁 ESTRUCTURA DE CARPETAS Y ARCHIVOS (DATOS)
    # ======================================================
    
    # 1. Archivo config.json (DATOS)
    config_json_path = os.path.join(base_dir, "config", "config.json")
    if os.path.exists(config_json_path):
        comando += ["--add-data", f"{config_json_path};config"]
    else:
        print("⚠️ Advertencia: no se encontró 'config/config.json'")

    # 2. Icono de la aplicación (DATOS Y RECURSO EXE)
    ico_path = os.path.join(base_dir, "config", "ico.ico")
    if os.path.exists(ico_path):
        comando += ["--icon", ico_path]
        comando += ["--add-data", f"{ico_path};config"]
    else:
        print("⚠️ Advertencia: no se encontró 'config/ico.ico'")

    # 3. Crear carpeta logs vacía en el bundle (LÓGICA ORIGINAL)
    logs_placeholder = os.path.join(legacy_dir, "logs", ".keep")
    if not os.path.exists(logs_placeholder):
        os.makedirs(os.path.dirname(logs_placeholder), exist_ok=True)
        with open(logs_placeholder, 'w') as f:
            f.write("# Placeholder para mantener la carpeta logs\n")
    
    comando += ["--add-data", f"{logs_placeholder};legacy/logs"]

    # Script principal con ruta completa
    main_path = os.path.join(base_dir, MAIN_SCRIPT)
    comando.append(main_path)
    
    return comando

# ==========================================================
# GENERAR EXE
# ==========================================================
def generar_exe():
    print("=" * 60)
    print("🚀 INICIANDO GENERACIÓN DEL EJECUTABLE (MODO ONEDIR)")
    print("=" * 60)

    verificar_main()
    verificar_estructura()
    limpiar_builds()

    cmd = construir_comando()
    print("⚙️  Comando PyInstaller:")
    print("   ", " ".join(cmd))
    print("\n🔨 Compilando, por favor espera...\n")

    result = subprocess.run(cmd)

    print("=" * 60)
    if result.returncode == 0:
        carpeta_exe = os.path.join(DIST_PATH, NOMBRE_EXE.replace(".exe", ""))
        print(f"✅ Generación completada correctamente.")
        print(f"📂 Carpeta de salida: {carpeta_exe}")
        print(f"📦 Ejecutable: {os.path.join(carpeta_exe, NOMBRE_EXE)}")
    else:
        print("❌ Error: PyInstaller no se ejecutó correctamente.")
        print("💡 Revisa los mensajes de error arriba para más detalles.")
    print("=" * 60)

# ==========================================================
# VERIFICAR SCRIPT PRINCIPAL
# ==========================================================
def verificar_main():
    ruta = os.path.join(os.getcwd(), MAIN_SCRIPT)
    if not os.path.isfile(ruta):
        print(f"❌ ERROR: No se encontró '{MAIN_SCRIPT}' en el directorio actual.")
        sys.exit(1)
    else:
        print(f"✅ Archivo principal encontrado: {MAIN_SCRIPT}\n")

# ==========================================================
# VERIFICAR ESTRUCTURA DE CARPETAS
# ==========================================================
def verificar_estructura():
    print("📁 Verificando estructura del proyecto:")
    
    carpetas_requeridas = [
        "config",
        "legacy", 
        "ui"
    ]
    
    archivos_requeridos = [
        "config/config_manager.py",
        "config/config.json",
        "config/ico.ico",
        "ui/estilos.py",
        "legacy/backend_base.py",
        "legacy/backend_extractor.py",
        "legacy/backend_clasificador.py",
        "legacy/extractor_adapter.py",
        "legacy/clasificador_adapter.py"
    ]
    
    todo_ok = True
    
    for carpeta in carpetas_requeridas:
        if os.path.exists(carpeta):
            print(f"   ✅ Carpeta '{carpeta}' encontrada")
        else:
            print(f"   ❌ Carpeta '{carpeta}' NO encontrada")
            todo_ok = False
    
    # ⭐ NOTA: He vuelto a poner los archivos .py aquí para verificar
    # que existen, aunque ahora sabemos que se compilan
    # gracias a --paths y no a --add-data
    for archivo in archivos_requeridos:
        if os.path.exists(archivo):
            print(f"   ✅ Archivo '{archivo}' encontrado")
        else:
            print(f"   ⚠️ Archivo '{archivo}' NO encontrado")
            # No marcamos como error crítico, solo advertencia
    
    if not todo_ok:
        print("\n❌ ERROR: Estructura del proyecto incompleta.")
        print("   Asegúrate de ejecutar este script desde la raíz del proyecto.")
        sys.exit(1)
    
    print()

# ==========================================================
# EJECUCIÓN PRINCIPAL
# ==========================================================
if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("   GENERADOR DE EJECUTABLE - GESTIÓN DE CORREOS OUTLOOK")
    print("=" * 60 + "\n")
    
    validar_entorno_virtual()
    confirmar_ejecucion()
    generar_exe()
    
    print("\n🎉 Proceso completado. ¡Gracias por usar el generador!")