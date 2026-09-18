"""Pre-scroll de contenido diferido: una sola implementacion para los cuatro usos.

Muchos formularios montan sus campos recien cuando entran en pantalla
(IntersectionObserver, lazy-loading). Por eso todo lo que abre un navegador baja primero
por la pagina: si no, la mitad del form todavia no existe en el DOM cuando se lo mira.

Esto estaba escrito cuatro veces, una por cada cosa que abre un navegador (Envio de Leads,
Comparar Dealers, Validacion de Campos y LambdaTest), y las cuatro copias arrastraban el
mismo bug: median la altura de la pagina UNA SOLA VEZ, antes de empezar a bajar. Al bajar,
una pagina con carga diferida crece, asi que el recorrido cortaba con la altura inicial y
se salteaba todo el medio - justo el contenido que esta funcion existe para cargar. El
salto final al fondo lo disimulaba: llegaba abajo sin haber pasado por las secciones
intermedias.

Dos de las copias, ademas, solo movian la ventana sin disparar el evento de scroll, que es
lo que escucha el codigo de la pagina para montar lo que falta.

Vive en utils porque lo usan `core`, `validation` y `providers`, y utils es el unico
paquete del que los tres ya importan.
"""

import time

# Techo de pasadas. Una pagina con scroll infinito crece cada vez que se baja, asi que sin
# un limite el bucle no termina nunca y deja la corrida colgada.
MAX_PASOS = 60

# Posicion con la que se pide "hasta el fondo" sin saber cuanto mide: el navegador la
# recorta al maximo real.
FONDO = 999999

# Mover la ventana no alcanza: hay frameworks que montan contenido escuchando el evento,
# no la posicion. Se dispara en window y en document porque no todos escuchan en el mismo.
_SCROLL_JS = (
    "window.scrollTo(0, arguments[0]);"
    "window.dispatchEvent(new Event('scroll', {bubbles:true,cancelable:false}));"
    "document.dispatchEvent(new Event('scroll', {bubbles:true}));"
)
_ALTO_JS = "return document.body.parentNode.scrollHeight"
_VIEWPORT_JS = "return window.innerHeight"


def pre_scroll(driver, step_wait=0.15, end_wait=0.3, max_pasos=MAX_PASOS):
    """Baja por la pagina disparando eventos de scroll, para que cargue lo diferido.

    La altura se vuelve a medir en CADA paso, y esa es la parte que importa: al bajar, la
    pagina crece. Termina dejando la pagina arriba, porque lo que viene despues (llenar,
    capturar, comparar) arranca desde el principio.

    Los tiempos de espera los elige quien llama: en headless el navegador necesita mas
    para procesar los eventos, y una sesion remota de LambdaTest tiene su propia latencia.
    """
    try:
        viewport = driver.execute_script(_VIEWPORT_JS) or 800
    except Exception:
        return
    # Piso de 800px: con un viewport chico los pasos se vuelven tantos que el techo de
    # pasadas corta antes de llegar al fondo.
    paso = max(viewport * 0.8, 800)

    posicion = 0
    for _ in range(max_pasos):
        try:
            alto = driver.execute_script(_ALTO_JS) or 0
            if posicion >= alto:
                break
            driver.execute_script(_SCROLL_JS, posicion)
        except Exception:
            # Se corta el recorrido, pero igual se intenta el fondo y la vuelta al tope:
            # dejar la pagina a mitad de camino es peor que no haber bajado.
            break
        time.sleep(step_wait)
        posicion += paso

    try:
        driver.execute_script(_SCROLL_JS, FONDO)
        time.sleep(end_wait)
        driver.execute_script(_SCROLL_JS, 0)
        time.sleep(end_wait)
    except Exception:
        pass
