import copy
from base_form_filler import BaseFormFiller
from country_configs import get_country_config


class GenericCountryBase(BaseFormFiller):
    """
    Clase base genérica para todos los países.
    Reemplaza los 9 archivos Formulario_*_Base.py individuales.
    """

    def __init__(self, country_name: str, browser="chrome", viewport="fullscreen", headless=False):
        base_config = get_country_config(country_name)
        config = copy.deepcopy(base_config)
        config['browser'] = browser
        config['viewport'] = viewport
        config['headless'] = headless
        super().__init__(config)
