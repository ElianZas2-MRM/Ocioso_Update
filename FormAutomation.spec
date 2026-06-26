# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules


a = Analysis(
    ['run.py'],
    pathex=['.', 'forms', 'core'],
    binaries=[],
    datas=[
        ('Asset/Fullheader.png', 'Asset'),
        ('Asset/icon.ico', 'Asset'),
        ('Asset/osopng.png', 'Asset'),
    ],
    hiddenimports=[
        'autonomous_runner',
        '_runner_common',
        'generic_country_base',
        'country_configs',
        'field_dependencies',
        'base_form_filler',
        'browser_manager',
        'screenshot_manager',
        'utils.data_generator',
        'utils.fixed_field_mapping_store',
        'utils.paths',
        'utils.popup_logger',
        'utils.scheduling',
        'interface.main_interface',
        'interface.helpers_interface',
        'interface.field_validation_ui',
        'interface.console_widget',
        'pytz',
        'openpyxl',
        'selenium.webdriver.chrome.webdriver',
        'selenium.webdriver.firefox.webdriver',
        'selenium.webdriver.edge.webdriver',
    ] + collect_submodules('selenium.webdriver'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    exclude_binaries=False,
    name='OsocioFormAutomation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['Asset\\icon.ico'],
)
