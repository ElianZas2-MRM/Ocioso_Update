"""
fabrica.py
==========
Le da al motor de los navegadores (BaseFormFiller) un navegador de BrowserStack.

El motor crea su navegador con un callable de la misma firma que
BrowserManager.create_browser (browser_type, viewport, headless, background), y se le
puede pasar otro (GenericCountryBase(..., browser_factory=...)). Esta fábrica es ese
otro: abre una sesión de BrowserStack y guarda la SesionBrowserStack para que el runner
pueda cerrarla, marcar el estado y traer el video al terminar.
"""

from typing import Callable, Optional

from osocio.providers.browserstack.sesion import SesionBrowserStack


class FabricaNavegadorBrowserStack:
    def __init__(self, plataforma: str, build_name: str, session_name: str,
                 log: Callable = print):
        self.plataforma = plataforma
        self.build_name = build_name
        self.session_name = session_name
        self.log = log
        self.sesion: Optional[SesionBrowserStack] = None

    def __call__(self, browser_type=None, viewport=None, headless=False, background=True):
        # viewport/headless/background son de los navegadores locales: en la nube el
        # tamaño lo define el dispositivo (DISPOSITIVOS en bs_driver).
        self.sesion = SesionBrowserStack(self.plataforma, self.build_name,
                                         self.session_name, log=self.log)
        return self.sesion.abrir()
