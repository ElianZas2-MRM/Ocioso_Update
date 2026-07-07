"""
Configuraciones específicas por país para el motor de formularios.
Cada entrada es la config que recibe BaseFormFiller.__init__() sin browser/viewport/headless,
que se inyectan en tiempo de ejecución por GenericCountryBase.
"""

COUNTRY_CONFIGS = {
    "Argentina": {
        'pais': 'Argentina',
        'excel_file': 'Lead_information_Formulario_Argentina_Main.xlsx',
        'data_start_index': 2,
        'wait_kits_after_model': True,
        'country_fields': {
            'required_fields': [
                "firstname", "lastname", "ci", "cellphone", "email",
                "models", "region", "city", "dealer", "estimated-date-purchase", "patent"
            ]
        },
        'field_mapping': [
            {'type': 'select', 'id': 'models',                              'data_index': 0,  'name': 'Modelo'},
            {'type': 'text',   'id': 'firstname',                           'data_index': 1,  'name': 'Nombre'},
            {'type': 'text',   'id': 'lastname',                            'data_index': 2,  'name': 'Apellido'},
            {'type': 'text',   'id': 'ci',                                  'data_index': 3,  'name': 'Documento'},
            {'type': 'text',   'id': 'cellphone',                           'data_index': 4,  'name': 'Celular'},
            {'type': 'text',   'id': 'email',                               'data_index': 5,  'name': 'Email'},
            {'type': 'select', 'id': 'region',                              'data_index': 6,  'name': 'Región'},
            {'type': 'select', 'id': 'city',                                'data_index': 7,  'name': 'Ciudad'},
            {'type': 'select', 'id': 'dealer',                              'data_index': 8,  'name': 'Concesionario'},
            {'type': 'select', 'id': 'estimated-date-purchase',             'data_index': 9,  'name': 'Fecha estimada'},
            {'type': 'text',   'id': 'patent',                              'data_index': 11, 'name': 'Patente'},
            {'type': 'text',   'id': 'event',                               'data_index': 12, 'name': 'Evento'},
            {'type': 'text',   'id': ['vin', 'chassis'],                    'data_index': 13, 'name': 'Chasis'},
            {'type': 'text',   'id': 'adviser',                             'data_index': 14, 'name': 'Código asesor'},
            {'type': 'text',   'id': ['comment', 'disapproval', 'suggestions'], 'data_index': 15, 'name': 'Comentario'},
        ],
    },

    "Bolivia": {
        'pais': 'Bolivia',
        'excel_file': 'Lead_information_Formulario_Bolivia_Main.xlsx',
        'data_start_index': 2,
        'country_fields': {
            'required_fields': ["name", "lastname", "document", "phone", "email", "model"]
        },
        'field_mapping': [
            {'type': 'select', 'id': 'model',                               'data_index': 0,  'name': 'Modelo'},
            {'type': 'text',   'id': 'name',                                'data_index': 1,  'name': 'Nombre'},
            {'type': 'text',   'id': 'lastname',                            'data_index': 2,  'name': 'Apellido'},
            {'type': 'text',   'id': 'document',                            'data_index': 3,  'name': 'Documento'},
            {'type': 'text',   'id': 'phone',                               'data_index': 4,  'name': 'Celular'},
            {'type': 'text',   'id': 'email',                               'data_index': 5,  'name': 'Email'},
            {'type': 'select', 'id': 'region',                              'data_index': 6,  'name': 'Región'},
            {'type': 'select', 'id': 'city',                                'data_index': 7,  'name': 'Ciudad'},
            {'type': 'select', 'id': 'dealer',                              'data_index': 8,  'name': 'Concesionario'},
            {'type': 'select', 'id': 'estimated-day',                       'data_index': 9,  'name': 'Fecha estimada'},
            {'type': 'text',   'id': 'patent',                              'data_index': 11, 'name': 'Patente'},
            {'type': 'select', 'id': 'event',                               'data_index': 12, 'name': 'Evento'},
            {'type': 'text',   'id': ['vin', 'chassis'],                    'data_index': 13, 'name': 'Chasis'},
            {'type': 'text',   'id': 'adviser',                             'data_index': 14, 'name': 'Código asesor'},
            {'type': 'text',   'id': ['comment', 'disapproval', 'suggestions'], 'data_index': 15, 'name': 'Comentario'},
        ],
    },

    "Brasil": {
        'pais': 'Brasil',
        'excel_file': 'Lead_information_Formulario_Brasil_Main.xlsx',
        'data_start_index': 2,
        'country_fields': {
            'required_fields': ["firstname", "lastname", "cpf", "cellphone", "email", "models"]
        },
        'field_mapping': [
            {'type': 'select', 'id': 'models',                                        'data_index': 0,  'name': 'Modelo'},
            {'type': 'text',   'id': ['firstname', 'fullname'],                       'data_index': 1,  'name': 'Nombre'},
            {'type': 'text',   'id': 'lastname',                                      'data_index': 2,  'name': 'Apellido'},
            {'type': 'text',   'id': ['cpf'],                                         'data_index': 3,  'name': 'CPF'},
            {'type': 'text',   'id': ['cnpj'],                                        'data_index': 4,  'name': 'CNPJ'},
            {'type': 'text',   'id': ['cep', 'customer-cep', 'zip', 'postal'],         'data_index': 5,  'name': 'CEP'},
            {'type': 'text',   'id': ['telephone', 'telephone-mask'],                 'data_index': 6,  'name': 'Celular'},
            {'type': 'text',   'id': 'email',                                         'data_index': 7,  'name': 'Email'},
            {'type': 'select', 'id': 'region',                                        'data_index': 8,  'name': 'Región'},
            {'type': 'select', 'id': 'city',                                          'data_index': 9,  'name': 'Ciudad'},
            {'type': 'select', 'id': 'dealer',                                        'data_index': 10, 'name': 'Concesionario'},
            {'type': 'select', 'id': 'estimated-date-purchase',                       'data_index': 11, 'name': 'Fecha estimada'},
            {'type': 'text',   'id': 'patent',                                        'data_index': 13, 'name': 'Patente'},
            {'type': 'select', 'id': ['event', 'event-name'],                         'data_index': 14, 'name': 'Evento'},
            {'type': 'text',   'id': ['chassis', 'vin', 'vin-code'],                  'data_index': 15, 'name': 'VIN', 'data_key': 'vin'},
            {'type': 'text',   'id': 'adviser',                                       'data_index': 16, 'name': 'Código asesor'},
            {'type': 'text',   'id': ['comment', 'disapproval', 'suggestions', 'requirement'], 'data_index': 17, 'name': 'Comentario'},
        ],
    },

    "Chile": {
        'pais': 'Chile',
        'excel_file': 'Lead_information_Formulario_Chile_Main.xlsx',
        'data_start_index': 2,
        'country_fields': {
            'required_fields': ["firstname", "lastname", "rut", "telephone", "email", "models"]
        },
        'field_mapping': [
            {'type': 'select', 'id': 'models',                              'data_index': 0,  'name': 'Modelo'},
            {'type': 'text',   'id': 'firstname',                           'data_index': 1,  'name': 'Nombre'},
            {'type': 'text',   'id': 'lastname',                            'data_index': 2,  'name': 'Apellido'},
            {'type': 'text',   'id': 'rut',                                 'data_index': 3,  'name': 'Documento'},
            {'type': 'text',   'id': 'telephone',                           'data_index': 4,  'name': 'Celular'},
            {'type': 'text',   'id': 'email',                               'data_index': 5,  'name': 'Email'},
            {'type': 'select', 'id': 'region',                              'data_index': 6,  'name': 'Región'},
            {'type': 'select', 'id': 'city',                                'data_index': 7,  'name': 'Ciudad'},
            {'type': 'select', 'id': 'dealer',                              'data_index': 8,  'name': 'Concesionario'},
            {'type': 'select', 'id': 'estimated-date-purchase',             'data_index': 9,  'name': 'Fecha estimada'},
            {'type': 'text',   'id': 'patent',                              'data_index': 11, 'name': 'Patente'},
            {'type': 'select', 'id': 'event',                               'data_index': 12, 'name': 'Evento'},
            {'type': 'text',   'id': 'chassis',                             'data_index': 13, 'name': 'Chasis'},
            {'type': 'text',   'id': 'adviser',                             'data_index': 14, 'name': 'Código asesor'},
            {'type': 'text',   'id': ['comment', 'disapproval', 'suggestions'], 'data_index': 15, 'name': 'Comentario'},
        ],
    },

    "Colombia": {
        'pais': 'Colombia',
        'excel_file': 'Lead_information_Formulario_Colombia_Main.xlsx',
        'data_start_index': 2,
        'country_fields': {
            'required_fields': ["firstname", "lastname", "ci", "telephone", "email", "models"]
        },
        'field_mapping': [
            {'type': 'select', 'id': 'models',                              'data_index': 0,  'name': 'Modelo'},
            {'type': 'text',   'id': 'firstname',                           'data_index': 1,  'name': 'Nombre'},
            {'type': 'text',   'id': 'lastname',                            'data_index': 2,  'name': 'Apellido'},
            {'type': 'text',   'id': 'ci',                                  'data_index': 3,  'name': 'Documento'},
            {'type': 'text',   'id': 'telephone',                           'data_index': 4,  'name': 'Celular'},
            {'type': 'text',   'id': 'email',                               'data_index': 5,  'name': 'Email'},
            {'type': 'select', 'id': 'region',                              'data_index': 6,  'name': 'Región'},
            {'type': 'select', 'id': 'city',                                'data_index': 7,  'name': 'Ciudad'},
            {'type': 'select', 'id': 'dealer',                              'data_index': 8,  'name': 'Concesionario'},
            {'type': 'select', 'id': 'estimated-date-purchase',             'data_index': 9,  'name': 'Fecha estimada'},
            {'type': 'text',   'id': 'patent',                              'data_index': 11, 'name': 'Patente'},
            {'type': 'select', 'id': 'event',                               'data_index': 12, 'name': 'Evento'},
            {'type': 'text',   'id': 'chassis',                             'data_index': 13, 'name': 'Chasis'},
            {'type': 'text',   'id': 'adviser',                             'data_index': 14, 'name': 'Código asesor'},
            {'type': 'text',   'id': ['comment', 'disapproval', 'suggestions'], 'data_index': 15, 'name': 'Comentario'},
        ],
    },

    "Ecuador": {
        'pais': 'Ecuador',
        'excel_file': 'Lead_information_Formulario_Ecuador_Main.xlsx',
        'data_start_index': 2,
        'country_fields': {
            'required_fields': ["firstname", "lastname", "ci", "telephone", "email", "models"]
        },
        'field_mapping': [
            {'type': 'select', 'id': 'models',                              'data_index': 0,  'name': 'Modelo'},
            {'type': 'text',   'id': 'firstname',                           'data_index': 1,  'name': 'Nombre'},
            {'type': 'text',   'id': 'lastname',                            'data_index': 2,  'name': 'Apellido'},
            {'type': 'text',   'id': 'ci',                                  'data_index': 3,  'name': 'Documento'},
            {'type': 'text',   'id': 'telephone',                           'data_index': 4,  'name': 'Celular'},
            {'type': 'text',   'id': 'email',                               'data_index': 5,  'name': 'Email'},
            {'type': 'select', 'id': 'region',                              'data_index': 6,  'name': 'Región'},
            {'type': 'select', 'id': 'city',                                'data_index': 7,  'name': 'Ciudad'},
            {'type': 'select', 'id': 'dealer',                              'data_index': 8,  'name': 'Concesionario'},
            {'type': 'select', 'id': 'estimated-date-purchase',             'data_index': 9,  'name': 'Fecha estimada'},
            {'type': 'text',   'id': 'patent',                              'data_index': 11, 'name': 'Patente'},
            {'type': 'select', 'id': 'event',                               'data_index': 12, 'name': 'Evento'},
            {'type': 'text',   'id': 'chassis',                             'data_index': 13, 'name': 'Chasis'},
            {'type': 'text',   'id': 'adviser',                             'data_index': 14, 'name': 'Código asesor'},
            {'type': 'text',   'id': ['comment', 'disapproval', 'suggestions'], 'data_index': 15, 'name': 'Comentario'},
        ],
    },

    "Paraguay": {
        'pais': 'Paraguay',
        'excel_file': 'Lead_information_Formulario_Paraguay_Main.xlsx',
        'data_start_index': 2,
        'country_fields': {
            'required_fields': ["firstname", "lastname", "ci", "telephone", "email", "models"]
        },
        'field_mapping': [
            {'type': 'select', 'id': 'models',                              'data_index': 0,  'name': 'Modelo'},
            {'type': 'text',   'id': 'firstname',                           'data_index': 1,  'name': 'Nombre'},
            {'type': 'text',   'id': 'lastname',                            'data_index': 2,  'name': 'Apellido'},
            {'type': 'text',   'id': 'ci',                                  'data_index': 3,  'name': 'Documento'},
            {'type': 'text',   'id': 'telephone',                           'data_index': 4,  'name': 'Celular'},
            {'type': 'text',   'id': 'email',                               'data_index': 5,  'name': 'Email'},
            {'type': 'select', 'id': 'region',                              'data_index': 6,  'name': 'Región'},
            {'type': 'select', 'id': 'city',                                'data_index': 7,  'name': 'Ciudad'},
            {'type': 'select', 'id': 'dealer',                              'data_index': 8,  'name': 'Concesionario'},
            {'type': 'select', 'id': 'estimated-date-purchase',             'data_index': 9,  'name': 'Fecha estimada'},
            {'type': 'text',   'id': 'patent',                              'data_index': 11, 'name': 'Patente'},
            {'type': 'select', 'id': 'event',                               'data_index': 12, 'name': 'Evento'},
            {'type': 'text',   'id': 'chassis',                             'data_index': 13, 'name': 'Chasis'},
            {'type': 'text',   'id': 'adviser',                             'data_index': 14, 'name': 'Código asesor'},
            {'type': 'text',   'id': ['comment', 'disapproval', 'suggestions'], 'data_index': 15, 'name': 'Comentario'},
        ],
    },

    "Peru": {
        'pais': 'Peru',
        'excel_file': 'Lead_information_Formulario_Peru_Main.xlsx',
        'data_start_index': 2,
        'country_fields': {
            'required_fields': ["firstname", "lastname", "ci", "telephone", "email", "models"]
        },
        'field_mapping': [
            {'type': 'select', 'id': 'models',                              'data_index': 0,  'name': 'Modelo'},
            {'type': 'text',   'id': 'firstname',                           'data_index': 1,  'name': 'Nombre'},
            {'type': 'text',   'id': 'lastname',                            'data_index': 2,  'name': 'Apellido'},
            {'type': 'select', 'id': 'document-type',                       'data_index': 10, 'name': 'Tipo de Documento'},
            {'type': 'text',   'id': 'ci',                                  'data_index': 3,  'name': 'Documento'},
            {'type': 'text',   'id': 'telephone',                           'data_index': 4,  'name': 'Celular'},
            {'type': 'text',   'id': 'email',                               'data_index': 5,  'name': 'Email'},
            {'type': 'select', 'id': 'region',                              'data_index': 6,  'name': 'Región'},
            {'type': 'select', 'id': 'city',                                'data_index': 7,  'name': 'Ciudad'},
            {'type': 'select', 'id': 'dealer',                              'data_index': 8,  'name': 'Concesionario'},
            {'type': 'select', 'id': 'estimated-date-purchase',             'data_index': 9,  'name': 'Fecha estimada'},
            {'type': 'text',   'id': 'patent',                              'data_index': 10, 'name': 'Patente'},
            {'type': 'select', 'id': 'event',                               'data_index': 11, 'name': 'Evento'},
            {'type': 'text',   'id': ['vin', 'chassis'],                    'data_index': 12, 'name': 'Chasis'},
            {'type': 'text',   'id': 'adviser',                             'data_index': 13, 'name': 'Código asesor'},
            {'type': 'text',   'id': ['comment', 'disapproval', 'suggestions'], 'data_index': 14, 'name': 'Comentario'},
        ],
    },

    "Uruguay": {
        'pais': 'Uruguay',
        'excel_file': 'Lead_information_Formulario_Uruguay_Main.xlsx',
        'data_start_index': 2,
        'country_fields': {
            'required_fields': ["firstname", "lastname", "ci", "telephone", "email", "models"]
        },
        'field_mapping': [
            {'type': 'select', 'id': 'models',                              'data_index': 0,  'name': 'Modelo'},
            {'type': 'text',   'id': 'firstname',                           'data_index': 1,  'name': 'Nombre'},
            {'type': 'text',   'id': 'lastname',                            'data_index': 2,  'name': 'Apellido'},
            {'type': 'text',   'id': 'ci',                                  'data_index': 3,  'name': 'Documento'},
            {'type': 'text',   'id': 'telephone',                           'data_index': 4,  'name': 'Celular'},
            {'type': 'text',   'id': 'email',                               'data_index': 5,  'name': 'Email'},
            {'type': 'select', 'id': 'region',                              'data_index': 6,  'name': 'Región'},
            {'type': 'select', 'id': 'city',                                'data_index': 7,  'name': 'Ciudad'},
            {'type': 'select', 'id': 'dealer',                              'data_index': 8,  'name': 'Concesionario'},
            {'type': 'select', 'id': 'estimated-date-purchase',             'data_index': 9,  'name': 'Fecha estimada'},
            {'type': 'text',   'id': 'patent',                              'data_index': 11, 'name': 'Patente'},
            {'type': 'select', 'id': 'event',                               'data_index': 12, 'name': 'Evento'},
            {'type': 'text',   'id': ['chassis', 'vin'],                    'data_index': 13, 'name': 'Chasis'},
            {'type': 'text',   'id': 'adviser',                             'data_index': 14, 'name': 'Código asesor'},
            {'type': 'text',   'id': ['comment', 'disapproval', 'suggestions'], 'data_index': 15, 'name': 'Comentario'},
        ],
    },
}

AVAILABLE_COUNTRIES = list(COUNTRY_CONFIGS.keys())


def get_country_config(country_name: str) -> dict:
    config = COUNTRY_CONFIGS.get(country_name)
    if config is None:
        raise ValueError(f"País '{country_name}' no encontrado. Disponibles: {AVAILABLE_COUNTRIES}")
    return config
