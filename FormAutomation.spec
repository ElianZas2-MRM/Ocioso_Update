# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules


a = Analysis(
    ['run.py'],
    # Todo el codigo cuelga del paquete osocio/, asi que alcanza con la raiz: ya no hace
    # falta listar una ruta por carpeta para que resuelvan los imports planos.
    pathex=['.'],
    binaries=[],
    datas=[
        ('Asset/Fullheader.png', 'Asset'),
        ('Asset/icon.ico', 'Asset'),
        ('Asset/osopng.png', 'Asset'),
        ('Asset/tabler_icons', 'Asset/tabler_icons'),
    ],
    # lambdatest_mac/android ya no van en datas: eran carpetas sueltas que se cargaban por
    # sys.path desde el disco. Ahora son subpaquetes de osocio.providers y entran como
    # codigo, igual que el resto.
    hiddenimports=collect_submodules('osocio') + [
        # collect_submodules ignora los modulos que arrancan con guion bajo, y este trae
        # get_runner(), que es por donde corre CADA pais. Sin esta linea el .exe compila
        # igual y revienta recien al ejecutar un pais.
        'osocio.forms._runner_common',
        'pytz',
        'openpyxl',
        'truststore',
        # Revisión Masiva en paralelo: se importa dentro de la función, no al tope del módulo.
        'concurrent.futures',
        # Icono de la bandeja (SysTrayIcon en interface/main_interface.py)
        'win32gui',
        'win32con',
        'win32api',
        'selenium.webdriver.chrome.webdriver',
        'selenium.webdriver.firefox.webdriver',
        'selenium.webdriver.edge.webdriver',
    ] + collect_submodules('selenium.webdriver'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['unittest', 'pydoc', 'pdb', 'difflib', 'tkinter.test'],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

# Build en modo ONEFILE: un unico OsocioFormAutomation.exe con todo comprimido adentro,
# sin carpeta _internal\ con las librerias sueltas al lado.
# Contra: el bootloader descomprime el bundle a %TEMP%\_MEIxxxxx en cada arranque, asi que
# la app tarda unos segundos mas en abrir que en onedir. Se acepta a cambio de entregar un
# solo archivo (lo que se reparte es el .exe + las carpetas de datos, nada mas).
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='OsocioFormAutomation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['Asset\\icon.ico'],
)
