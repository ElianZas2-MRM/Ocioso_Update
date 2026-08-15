"""
generic_country_base.py — Clase base genérica que inyecta la configuración de país.
Reemplaza los 9 archivos individuales (Formulario_*_Base.py) del proyecto original.
Recibe el nombre del país, carga su config desde country_configs.py y crea BaseFormFiller.
"""
import copy
from base_form_filler import BaseFormFiller
from country_configs import get_country_config


class GenericCountryBase(BaseFormFiller):
    """
    Clase base genérica para todos los países.
    Reemplaza los 9 archivos Formulario_*_Base.py individuales.
    """

    def __init__(self, country_name: str, browser="chrome", viewport="fullscreen", headless=False, background=True, is_scheduled=False, pausar_autenticacion=False, preview_visible_browser=False, excel_suffix=""):
        base_config = get_country_config(country_name)
        config = copy.deepcopy(base_config)
        config['browser'] = browser
        config['viewport'] = viewport
        config['headless'] = headless
        config['background'] = background
        config['is_scheduled'] = is_scheduled
        config['pausar_autenticacion'] = pausar_autenticacion
        config['preview_visible_browser'] = preview_visible_browser

        # Determinar el nombre del excel basado en el navegador/dispositivo
        dev_name = "Chrome"
        if browser == "firefox":
            dev_name = "Firefox"
        elif browser == "edge":
            dev_name = "Edge"
        # excel_suffix permite apuntar a una variante del mismo mercado sin duplicar
        # configuración (hoy sólo "_T3", los formularios 2.0 de Adobe AEM).
        config['excel_file'] = f"Lead_information_Formulario_{country_name}_{dev_name}{excel_suffix or ''}.xlsx"

        super().__init__(config)
