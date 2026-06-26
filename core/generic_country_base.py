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

    def __init__(self, country_name: str, browser="chrome", viewport="fullscreen", headless=False, background=True):
        base_config = get_country_config(country_name)
        config = copy.deepcopy(base_config)
        config['browser'] = browser
        config['viewport'] = viewport
        config['headless'] = headless
        config['background'] = background
        super().__init__(config)
