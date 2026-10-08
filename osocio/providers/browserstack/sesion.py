"""
sesion.py
=========
Ciclo de vida de UNA sesión de BrowserStack: abrirla, y al cerrarla marcarla como
pasada/fallada en el dashboard y traer el link del video.

Lo usan los dos motores (el de los navegadores para Android/Mac y el nativo del
iPhone), así que el cierre y el video funcionan igual para los tres dispositivos.
"""

from typing import Callable, Dict, Optional

from osocio.providers.browserstack.bs_driver import (
    create_bs_driver, fetch_bs_session, load_credentials, mark_bs_status_api,
)


class SesionBrowserStack:
    def __init__(self, plataforma: str, build_name: str, session_name: str,
                 log: Callable = print):
        self.plataforma = plataforma
        self.build_name = build_name
        self.session_name = session_name
        self.log = log
        self.driver = None
        self.session_id = ""
        self._credenciales: Optional[tuple] = None

    def abrir(self):
        """Crea la sesión remota. Falla con un mensaje claro si no hay credenciales."""
        self._credenciales = load_credentials()
        self.driver = create_bs_driver(*self._credenciales, platform=self.plataforma,
                                       build_name=self.build_name,
                                       session_name=self.session_name)
        self.session_id = self.driver.session_id
        self.log(f"✓ Sesión BrowserStack ({self.plataforma}): {self.session_id}")
        return self.driver

    def cerrar(self, passed: bool, reason: str = "") -> Dict[str, str]:
        """Marca el estado, cierra el driver y devuelve los datos de la sesión (video_url...).

        Se puede llamar aunque el driver ya esté cerrado (el motor de los navegadores lo
        cierra en su cleanup): por eso el estado se marca por la API REST y no por el driver.
        """
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None
        if not (self.session_id and self._credenciales):
            return {"video_url": "", "device": "", "os_version": "", "status": ""}
        mark_bs_status_api(self.session_id, *self._credenciales, passed, reason)
        return fetch_bs_session(self.session_id, *self._credenciales)
