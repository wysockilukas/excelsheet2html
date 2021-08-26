from pathlib import Path


__all__ = ["get_js_script"]


def _add_variables_to_script(input):
    js = 'const table_data = {};\n'
    js += f'table_data.primary_cells = {input.primary_cells};\n'
    js += f'table_data.secondary_cells = {input.secondary_cells};\n'
    js += f'table_data.formula_cells = {input.formula_cells};\n'
    js += f'table_data.secondary_cells_1 = {input.secondary_cells_1};\n'
    js += f'table_data.secondary_cells_2 = {input.secondary_cells_2};\n'
    return js


def get_js_script(class_name_prefix, input,  number_format, precission):
    f = open(Path(str(Path(__file__).parent.resolve()) +
                  "/app.js"), 'r')
    # f = open('app.js', 'r')
    file_content = f.read()
    f.close()
    js = file_content.replace('[PREF]', class_name_prefix)
    js = js.replace('[FORMAT]', str(number_format))
    js = js.replace('[PRECISION]', str(precission))

    vars = _add_variables_to_script(input)

    script = f"""
    <script>
    {vars}
    {js}
    </script>
    """
    return script
