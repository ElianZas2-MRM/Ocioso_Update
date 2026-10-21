"""
Proveedor BrowserStack Automate (web): iPhone, Android y Mac en la nube.

Reemplazó a LambdaTest en la app. Corre los forms con el MISMO motor que los navegadores
locales (BaseFormFiller), así respeta todas las reglas de la app; el iPhone usa ese motor
con el llenado por la capa nativa, porque Safari en iPhone no deja entrar al iframe.

Módulos:
  runner.py           punto de entrada (UI, CLI, runner autónomo): run(...) -> resumen
  dispositivos.py     tabla única iPhone BS / Android BS / Mac BS y tope de paralelo
  bs_driver.py        credenciales, capabilities, creación del driver, estado y video
  sesion.py           ciclo de vida de una sesión (estado por API REST, video)
  fabrica.py          navegador de BrowserStack para el motor (browser_factory)
  iphone_runner.py    FormFillerIphone: el motor con el llenado por la capa nativa
  ios_nativo.py       la capa nativa del iPhone (detectar, tocar, tipear, TY)
  prueba_conexion.py  herramienta de desarrollo contra una landing real
"""
