class MyGlobals():
    """
    Class - container for global parameters.
    The object of type MyGlobals is passed to all functions so object properties are avialable globaly on project scope

    :param area: eg. 'A1:G16', holds range of excel sheet to be converted as htmls. The value is taken form excel named rnge - `area`
    :type area: string 
    :param column_width: dict with key as column letter and value its width
    :type column_width: dict, like {'A':85}
    :param first_row: row number of first row of table, where labels starts
    :type first_row: int   
    :param formula_cell: dict with key: formula cell coordinate and value its value
    :type formula_cell: dict, like {'E4':345}   
    :param grid_column_width: width of first column in table html. One with row numbers
    :type grid_column_width: int  
    :param headers: eg. 'A1:G16', range of table headers. The value is taken form excel named rnge - `headers`
    :type headers: string
    :param labels: eg. 'A1:G16', range of table labels. The value is taken form excel named rnge - `labels`
    :type labels: string 
    :param last_row: row number of last row of table, where labels ends
    :type last_row: int  
    :param merged_cells: dict with all merged cells. eg {'B14':{'colspan':3}}
    :type merged_cells: dict  
    :param num_of_columns: numnber of columns - taken from area
    :type num_of_columns: int 
    :param num_of_rows: numnber of rows - taken from area
    :type num_of_rows: int   
    :param onbottom: eg. 'A1:G16', range of cells - part of area under the table.
    :type onbottom: string
    :param ondead_area: eg. 'A1:G16', range of cells - on left of headers but over the labels
    :type ondead_area: string 
    :param onleft: eg. 'A1:G16', range of cells - part of area on left of the table.
    :type onleft: string  
    :param onright: eg. 'A1:G16', range of cells - part of area on right of the table.
    :type onright: string 
    :param ontop: eg. 'A1:G16', range of cells - part of area on top of the table.
    :type ontop: string     
    :param read_borders: indicates if border style is taken from excel sheet or default style is used
    :type read_borders: bool  
    :param right_column: last column of table. 
    :type right_column: int  
    :param theme+colors: excel theme color list eg [`FFFFFF`]
    :type theme+colors: list 
    :param total_width: total html table width
    :type total_width: int   
    :param values: eg. 'A1:G16', range of table values. The value is taken form excel named rnge - `values`
    :type values: string                                                                             
    """

    def __init__(self) -> None:
        """
        Constructor method
        """
        self.read_borders = False
        self.theme_colors = []

        self.area = None
        self.labels = None
        self.headers = None
        self.values = None
        self.ontop = None
        self.onbottom = None
        self.onleft = None
        self.onright = None
        self.ondead_area = None

        self.num_of_rows = None
        self.num_of_columns = None
        self.last_row = None
        self.first_row = None
        self.right_column = None
        self.merged_cells = None
        self.grid_column_width = None
        self.column_width = {}

        self.total_width = None
        self.formula_cells = {}


def set_theme_colors(wb):
    """
    Method returns theme colors from workbook as a list
    """
    global theme_colors
    """Gets theme colors from the workbook"""
    # see: https://groups.google.com/forum/#!topic/openpyxl-users/I0k3TfqNLrc
    from openpyxl.xml.functions import QName, fromstring
    xlmns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    root = fromstring(wb.loaded_theme)
    themeEl = root.find(QName(xlmns, 'themeElements').text)
    colorSchemes = themeEl.findall(QName(xlmns, 'clrScheme').text)
    firstColorScheme = colorSchemes[0]

    colors = []

    for c in ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']:
        accent = firstColorScheme.find(QName(xlmns, c).text)
        # walk all child nodes, rather than assuming [0]
        for i in list(accent):
            if 'window' in i.attrib['val']:
                colors.append(i.attrib['lastClr'])
            else:
                colors.append(i.attrib['val'])
    return colors
