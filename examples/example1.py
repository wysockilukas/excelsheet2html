from excelsheet2html import excelsheet2html
from pathlib import Path


if __name__ == "__main__":
    root_path = Path(__file__).resolve().parent
    output = excelsheet2html(
        filepath=Path(str(root_path) + "/Book1.xlsx"), class_name_prefix='c_', read_borders=True, number_format='comma')
    f = open(Path(str(root_path) + "/Book2.html"), 'w')
    f.write(output.html)
    f.close()
