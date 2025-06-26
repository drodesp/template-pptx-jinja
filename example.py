# example.py
# Este script sirve como ejemplo principal para generar una presentación PPTX dinámica
# desde una plantilla utilizando Jinja2 y la biblioteca "template_pptx_jinja".

import jinja2
from template_pptx_jinja.render import PPTXRendering

def main():
    # Ruta al archivo PPTX de plantilla
    input_path = 'example/template.pptx'

    # Diccionario de datos para rellenar los placeholders en la plantilla
    model = {
        "name": "John",
        "number": 3,
        "step": [
            {"name": "analysis"},
            {"name": "design"},
            {"name": "production"},
            {"name": "production2"},
        ],
        "size": [10, 9000],
        "my_table_name": "My filling table",
        "my_table": [
            ["Hello", "World"],
            ["Python", "Programming"],
            ["Data", "Science"],
            ["Machine", "Learning"],
            ["Artificial", "Intelligence"]
        ]
    }

    # Diccionario de imágenes a reemplazar: hash original -> ruta nueva
    pictures = {
        "example/model.jpg": "example/image.jpg"
    }

    # Contexto para Jinja2: modelo de datos + imágenes
    data = {
        'model': model,
        'pictures': pictures
    }

    # Definimos un filtro personalizado para Jinja2 (ej. pluralizar)
    def plural(input, word_ending):
        return word_ending if input > 0 else ''

    jinja2_env = jinja2.Environment()
    jinja2_env.filters['plural'] = plural

    # Ruta de salida del fichero generado
    output_path = 'example/presentation_generated.pptx'

    # Instancia del motor de renderizado
    rendering = PPTXRendering(input_path, data, output_path, jinja2_env)
    message = rendering.process()
    print(message)
# 

if __name__ == '__main__': main()
# 
