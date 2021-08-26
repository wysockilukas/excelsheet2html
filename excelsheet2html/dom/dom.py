from openpyxl.utils.cell import column_index_from_string, get_column_letter
from excelsheet2html.config.constants import *
import re

__all__ = "DOM"


def _is_in_range(cell: 'object', rng: 'str') -> bool:
    """
    Check if cell is in given range, return bool
    """
    cell_c = cell.column
    cell_r = cell.row
    a = rng.split(':')
    cl = int(column_index_from_string(re.findall('[A-Z]+', a[0])[0]))
    cr = int(column_index_from_string(re.findall('[A-Z]+', a[1])[0]))
    rt = int(re.findall('\d+', a[0])[0])
    rb = int(re.findall('\d+', a[1])[0])
    return (cl <= cell_c <= cr) and (rt <= cell_r <= rb)


def _get_border_style_from_cell(cell):
    class_name = []
    for border_direction in ["right", "left", "top", "bottom"]:
        b_s = getattr(cell.border, border_direction)
        if not b_s:
            continue
        if b_s.style in {'thick', 'thin', 'double', 'medium'}:
            class_name.append(f'b_{border_direction}')
    return class_name


def _select_cells(num_of_rows, num_of_columns, ws):
    tbl = []
    for row_i in range(1, num_of_rows + 1):
        one_row = []
        for col_i in range(1, num_of_columns + 1):
            one_row.append(ws.cell(row_i, col_i))
        tbl.append(tuple(one_row))
    return tuple(tbl)


def _render_colgroup(glb, class_prefix):
    html = '<colgroup>\n'
    html += f""" <col class="{class_prefix}row_num" style="width:{glb.grid_column_width}px">\n"""
    for k, v in glb.column_width.items():
        html += f""" <col class="{class_prefix}col_{k}" style="width:{glb.column_width.get(k)}px">\n"""
    html += '</colgroup>\n'
    return html


def _render_header(glb, class_prefix):
    html = '<thead>\n'
    html += '<tr>\n'
    html += f""" <th class="{class_prefix}row_num"></th>\n"""
    for k, v in glb.column_width.items():
        html += f""" <th class="{class_prefix}col_{k}">{k}<div class="{class_prefix}resizer"></div></th>\n"""
    html += '</tr>\n'
    html += '</thead>\n'
    return html


def _get_class(glb, cell: object, stylesheet: object) -> str:
    class_name = []
    pref = stylesheet.class_name_prefix

    if glb.read_borders:
        cls = _get_border_style_from_cell(cell)
        if len(cls):
            ext_cls = [f'{pref}{x}' for x in cls]
            class_name.extend(ext_cls)

    if cell.alignment.indent > 0:
        cname = f"{pref}label_padding_" + str(int(cell.alignment.indent))
        class_name.append(cname)
        stylesheet.add_class(f"table.{pref}table td.{cname}",
                             "padding-left: " + str(int(cell.alignment.indent) * 10) + "px;")

    if cell.fill.start_color.type == 'rgb':
        if cell.fill.start_color.rgb != '00000000':
            rgb = cell.fill.start_color.rgb[2:]
            cname = f"{pref}bgcolor_" + rgb
            stylesheet.add_class(
                f".{cname}", "background-color: #" + rgb + ";")
            class_name.append(cname)

    if cell.fill.start_color.type == 'indexed':
        rgb = COLOR_INDEX[cell.fill.start_color.indexed][2:]
        cname = f"{pref}bgcolor_" + str(cell.fill.start_color.indexed)
        stylesheet.add_class(f".{cname}", "background-color: #" + rgb + ";")
        class_name.append(cname)
    if cell.fill.start_color.type == 'theme':
        rgb = glb.theme_colors[cell.fill.start_color.theme]
        cname = f"{pref}bgcolor_theme" + str(cell.fill.start_color.theme)
        stylesheet.add_class(f".{cname}", "background-color: #" + rgb + ";")
        class_name.append(cname)

    if cell.font.b:
        class_name.append(f"{pref}font_b")
        stylesheet.add_class(f".{pref}font_b", "font-weight: bold;")

    if glb.ontop:
        if cell.value is None and _is_in_range(cell, glb.ontop):
            class_name.append(f'{pref}ontop')

    if glb.onleft:
        if cell.value is None and _is_in_range(cell, glb.onleft):
            class_name.append(f'{pref}onleft')

    if glb.onright:
        if cell.value is None and _is_in_range(cell, glb.onright):
            class_name.append(f'{pref}onright')

    if glb.onbottom:
        if cell.value is None and _is_in_range(cell, glb.onbottom):
            class_name.append(f'{pref}onbottom')

    if _is_in_range(cell, glb.labels) and cell.value is not None:
        class_name.append(f'{pref}label')

    if _is_in_range(cell, glb.labels) and glb.merged_cells.get(cell.coordinate, None):
        mcell = glb.merged_cells.get(cell.coordinate, None)
        rowspan = mcell.get('rowspan', None)
        if rowspan:
            class_name.append(f'{pref}label_rowspan')

    if glb.headers:
        if _is_in_range(cell, glb.headers) and cell.value is not None:
            class_name.append(f'{pref}header')

    if glb.ondead_area:
        if _is_in_range(cell, glb.ondead_area):
            class_name.append(f'{pref}deadarea')

    if _is_in_range(cell, glb.values):
        class_name.append(f'{pref}values')
    if cell.row == glb.last_row and (_is_in_range(cell, glb.labels) or _is_in_range(cell, glb.values)):
        class_name.append(f'{pref}last_row')

    if cell.row == glb.first_row and _is_in_range(cell, glb.labels):
        class_name.append(f'{pref}first_row')

    if cell.column == glb.right_column and (_is_in_range(cell, glb.headers) or _is_in_range(cell, glb.values)):
        class_name.append(f'{pref}right_column')

    return ' '.join(class_name)


def _render_td(glb, cell: object,  stylesheet: object) -> str:
    cls = _get_class(glb, cell, stylesheet)
    value = cell.value
    if cell.value is None:
        value = ''
    cell_type = type(cell).__name__
    if cell_type == 'MergedCell':
        return None

    mcell = glb.merged_cells.get(cell.coordinate, None)
    if mcell:
        rowspan = mcell.get('rowspan', None)
        colspan = mcell.get('colspan', None)
        colspan_atr = f'colspan="{colspan}"' if colspan else ''
        rowspan_atr = f'rowspan="{rowspan}"' if rowspan else ''
        return f'<td id="{cell.coordinate}" class="{cls}" {rowspan_atr} {colspan_atr}>{value}</td>'

    if _is_in_range(cell, glb.values) and cell.coordinate in glb.formula_cells:
        glb.formula_cells[cell.coordinate] = cell.value
        return f'<td id="{cell.coordinate}" class="{cls}" data-is-formula="true">{value}</td>'

    return f'<td id="{cell.coordinate}" class="{cls}">{value}</td>'


def _render_table(glb, stylesheet, all_cells):
    class_prefix = stylesheet.class_name_prefix
    html = '<tbody>\n'
    table_cells = []
    for row in range(1, glb.num_of_rows + 1):
        table_cells.append([f'<td class="{class_prefix}row_num">{row}</td>'])

    for row_id in range(0, glb.num_of_rows):
        for col_id in range(0, glb.num_of_columns):
            _html = _render_td(glb, all_cells[row_id][col_id], stylesheet)
            if _html:
                table_cells[row_id].append(_html)
    for r in table_cells:
        html += '<tr>\n'
        for e in r:
            html += e + '\n'
        html += '</tr>\n'
    html += '</tbody>\n'
    return html


class DOM():
    def __init__(self, glb, ws: object, stylesheet: object) -> None:
        self.glb = glb
        self.ws = ws
        self.stylesheet = stylesheet
        self.all_cells = _select_cells(
            glb.num_of_rows, glb.num_of_columns, ws)
        # print(props)
        # print(self.all_cells)

    def render_table(self):
        colgroup = _render_colgroup(
            self.glb, self.stylesheet.class_name_prefix)
        header = _render_header(self.glb, self.stylesheet.class_name_prefix)
        table = _render_table(self.glb, self.stylesheet, self.all_cells)

        self.stylesheet.create_global_css()

        html = f"""
        <table class="{self.stylesheet.class_name_prefix}table" style="width:{self.glb.total_width}px">
        {colgroup}
        {header}
        {table}
        </table>
        """

        return html
