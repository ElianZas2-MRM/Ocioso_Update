# Historial de cambios del build

Este historial vivia adentro de `build.bat`, donde ocupaba 116 de sus 280 lineas: quien
abria el script para entender como se arma el `.exe` se comia todo esto primero.

Se conserva tal cual estaba, solo sin el prefijo `REM`.

---

Ago 2026: Fix: la app no arrancaba (NameError sched_mode_btns al armar la pestana de
          programacion). Ademas: los paises de "Envio de Leads" no eran clickeables
          (faltaba el bind), y "Desactivar"/"Programado" del footer nunca se
          actualizaban (el guard usaba globals()/locals(), siempre False)
          Pestana "Programacion de Tests" -> "Testeo Programado", con sub-pestanas
          "Envio de Leads Normales" y "Revision Masiva de Insercion (Smoke Test)".
          Cada sub-pestana tiene su guia rapida, su card de configuracion (mercados,
          navegadores, secuencial/paralelo) y su resumen con badges. Los dias y
          horarios se configuran desde el boton del footer
          Revision Masiva: ya NO genera un Excel por pais. Sale UN solo archivo por
          corrida, resultados/resultado_urlsinsertas/Resultados_Revision_Masiva_
          <matriz>_<fecha>.xlsx: es una COPIA del Excel matriz ya editada, con las
          columnas de resultado al principio de cada hoja/mercado y el resto de las
          columnas originales atras. El Excel matriz NO se modifica
          Revision Masiva: espera hasta 5s a que aparezca el iframe del form antes de
          darlo por ausente (landings que lo cargan en diferido daban FAIL falso)
          Revision Masiva: "Rerun fails previos" ahora reintenta cualquier fila con
          FAIL en CUALQUIER columna (antes solo miraba 3) y las SKIPPED
          Revision Masiva: se saca el check "Ejecucion en paralelo" del footer (el modo
          se elige en la card); los mercados se recorren de a uno porque todos escriben
          en el mismo libro. Boton "INICIAR REVISION MASIVA" -> "Iniciar ahora"
          Build: vuelve a ONEFILE (un solo .exe con todo comprimido adentro, sin la
          carpeta _internal\ con las librerias sueltas) y se rearma el ZIP del portable
          Build: el portable ya no viaja con resultados/capturas/reportes de esta PC;
          esas carpetas se crean vacias. Tampoco viaja ningun Excel de datos: solo
          data\Field_Validation_URLs.xlsx (los de leads los genera la app, el matriz de
          Revision Masiva lo elige el usuario). json/: se excluyen tambien
          programacion_leads y programacion_masivo
Jul 2026: Excel de resultados: columnas de RESULTADO primero y los datos de entrada al
          final (antes era al reves). Se reordena al cerrar la corrida, no antes: los
          runners leen los datos del lead de la misma hoja (URL en A, Formulario en B)
          y moverlas durante la corrida desalinearia esa lectura
          Si la landing no cargo (404 / sin respuesta / redirect), "Form coincide" y
          "Formulario Inserto" ya no dicen "no coincide con el esperado" — el form nunca
          se busco. Dicen "N/D — la landing no cargo (...)" con el estado real
Jul 2026: Excel: SIEMPRE se dice POR QUE fallo un lead. Nueva columna "Motivo" con la
          causa corta (form no inserto / no coincide / validacion del form + el campo /
          sin TY page / event_id / 404). Antes habia que leer el Resultado entero
          "Formulario Inserto" pasa a TRES estados en vez de dos:
            verde   "Form inserto"
            AMBAR   "Form inserto NO coincide con el esperado, se envio lead igualmente"
            rojo    "Form NO inserto (iframe sin src / sin form en la landing)"
          "Form coincide" acompana: FAIL en ambar si el lead viajo igual (hay dato en la
          base, entro por otro form), FAIL en rojo si no salio o no habia form
          Nuevas columnas "Estado URL landing" / "Estado URL form" (utils/url_status.py):
          status HTTP real de cada URL — 200 verde, redirect ambar, 404/503/sin respuesta
          rojo. Selenium no expone el status, asi que una landing caida se veia solo como
          "form no encontrado". Si la URL falla, se antepone al Motivo como causa raiz
          Aplica a browsers (escritorio) y a LambdaTest Mac/Android
Jul 2026: Resultados = lo que REALMENTE se envio: antes del click en Enviar se relee el
          DOM (1 solo execute_script) y se pisa el tracking con el valor efectivo. El
          reintento (recarga la landing y rellena de cero) re-sorteaba modelo/ciudad/
          concesionario y el Excel quedaba con lo del PRIMER intento -> no coincidia con
          la base de datos. Los dropdowns random sin trackear quedan como Final::<campo>
          Nueva columna "Datos vs Excel": verde OK si los dropdowns quedaron como se
          pidio; ambar con el detalle (pedido 'X' -> quedo 'Y') si no, y ahi el Resultado
          tampoco queda verde. Solo dropdowns: en texto el flujo transforma a proposito
          (CPF/CNPJ regenerados, maxlength) y darian avisos falsos
          Fix Android: el.clear() no vaciaba el campo NI lanzaba excepcion -> el
          re-ingreso concatenaba ('ApellidoApellido', email duplicado) y el lead fallaba
          por validacion. Vaciado verificado en cascada; si no se puede, no escribe encima
          Fix modelo por ?model=: se leia driver.current_url, que dentro del iframe es la
          landing (sin el parametro) -> esas filas quedaban sin modelo registrado. Ahora
          usa la URL del form del Excel, con filtro de tokens que no son modelo
          Aplica a LambdaTest Android/Mac (lt_runner.py) y al runner de escritorio
Jul 2026: iframe GM: siempre priorizar src con gm_forms/gm_front/gm_admin (evita agarrar
          el iframe equivocado cuando la landing tiene varios) — browsers y LambdaTest
          CTA: mas selectores por texto ES/PT + barrido generico del <form> (fix "no
          encuentra el boton que si esta"); retry recargando SOLO el iframe (no la landing)
          Capturas: landing_inicial -> form_vacio -> form_errores -> form_completado -> TY
          -> landing_final; sin capturas de landing completa en el medio; multipaso por paso
          Foco: ventana real off-screen SIN robar foco via Win32 SW_SHOWNOACTIVATE (no headless)
          Modal: progreso por FORMULARIOS (no sesiones) + mercados en ejecucion + filas con
          error y motivo corto (form ausente / incorrecto / sin TYP / landing 404 / ...)
          Errores: cualquier fallo del form = FAIL en Excel/UI/email (TYP no vista, form
          incorrecto, landing 404, campo/dropdown sin completar). "FORMULARIO AUSENTE" si no
          hay form; "Formulario incorrecto (distinto al esperado)" si es otro
          Resultados: columna Modelo = modelo elegido en el dropdown o ?model= real del form
          LambdaTest: NUNCA capturas (evidencia = video); mismo iframe GM / CTA / modelo
Jul 2026: Nueva pestaña "Comparador Dealers": chequea region/ciudad/dealer/BAC/modelos
          contra un Excel de dealers (fila de encabezado y columnas configurables,
          múltiples forms por pasada, detección de duplicados y extras de forma jerárquica,
          desbloqueo de nivel dealer, columnas adicionales como píldoras-checkbox en vivo,
          ejecución en dos fases (comparación rápida primero + capturas ZIP opcionales después),
          corrección de screenshots en Chrome dentro de iframe, e icono de la app propio.
Jun 2026: Generar Excels: panel Brasil visible al seleccionar pais (pack before fix)
          Generar Excels: boton "Regenerar datos (URLs actuales)" sin ingresar URLs
          Generar Excels: radio buttons modo envio (redondelitos, sin indicatoron=0)
          Ejecucion: forms standalone (sin landing URL) ya no se saltean
          Browser: viewport desktop garantizado (1366x768) en Chrome/FF/Edge
          Browser: ventana movida fuera de pantalla (x=10000) en los 3 navegadores
          Build: lambdatest_android incluido en portable y dist
Jun 2026: LambdaTest: fix terms checkbox (React fiber onChange + shadow DOM cookies)
          LambdaTest: zoom 80% en URLs comprar-carro para visibilidad completa
          LambdaTest: scroll al fondo tras 2do campo en forms comprar-carro
          LambdaTest: cookie popup GM via shadowRoot (gb-legal-notification)
Jun 2026: Fix Firefox: screenshots faltantes en form_errores/completado/typage
          CPF/CNPJ/CEP: cero inicial recuperado si Excel lo omio (zfill)
          Chevrolet BR: click en #contact-by-form antes de entrar al form
          Error COM de Outlook muestra mensaje real (antes generico)
          Tracking por paso en Excel de resultados (columnas PasoN::campo)
          Seleccion de dropdowns ahora solo exacta (fuzzy matching eliminado)
          Generacion de documentos Brasil via API 4devs
          Retirado xlsxwriter (reemplazado por openpyxl)
          Soporte checkboxes custom con opacity:0 (fake-terms, stat-radio)
          Retry generico: cualquier fallo de envio recarga y reintenta (max 2)
          Sanitizacion 'test' a 'prueba' en campos de texto
          Email auto-generado: formato {nombre}{apellido}_{pais}{nn}@mrm.com
          Fix URLs form esperado/encontrado por fila en capturas
          pywin32 requerido para envio de email via Outlook (win32com)
          libro-reclamaciones (chevrolet.com.pe): IDs fijos cc_name/
          cc_telephone/cc_ci/cc_email + ciudad/dealer aleatorio + envio real
          Deteccion exacta de error transitorio de envio: BR (Desculpe...) y
          resto de mercados (Lo siento, ocurrio un inconveniente...)
