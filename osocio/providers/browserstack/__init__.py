"""
Proveedor BrowserStack Automate (web) — reemplazo de LambdaTest.

Vive separado de lambdatest_mac/ y lambdatest_android/ a propósito: mientras dure la
migración conviven los dos, y LambdaTest no se toca hasta que BrowserStack esté validado
en los tres dispositivos (iPhone 14 Pro Max, Samsung Galaxy S26 Ultra y Mac Safari).

Módulos:
  bs_driver.py        credenciales, capabilities, creación del driver, estado y video
  prueba_conexion.py  prueba de humo contra una landing real (fase 1, no envía leads)
"""
