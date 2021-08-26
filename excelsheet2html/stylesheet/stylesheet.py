from pathlib import Path


__all__ = ['StyleSheet']


class StyleSheet():
    def __init__(self, glb, class_name_prefix) -> None:
        self.glb = glb
        self.class_name_prefix = class_name_prefix
        self.classes = []
        self.dynamic_classes_names = []
        self.font_color = 'rgb(68, 68, 68)'
        self.sheet_border = '1px solid rgba(205, 205, 205, 0.8)'

    def create_global_css(self):
        f = open(Path(str(Path(__file__).parent.resolve()) +
                 "/global_css_template.txt"), 'r')
        file_content = f.read()
        f.close()

        for c in file_content.split('[END]'):
            c_name, c_content = c.split('[SEP]')
            c_name = c_name.replace('[PREF]', self.class_name_prefix)
            c_content = c_content.replace('[BORDER1]', self.sheet_border)
            if self.glb.read_borders:
                c_content = c_content.replace('[BORDER2]', 'unset')
            else:
                c_content = c_content.replace('[BORDER2]', '1px solid black')
            c_content = c_content.replace('[COLOR1]', self.font_color)
            self.classes.append(c_name + "\n{"+c_content+"}")

    def add_class(self, class_name, css):
        if not class_name in self.dynamic_classes_names:
            self.classes.append(f"""
                {class_name} {{
                    {css}
                }}        
            """)
            self.dynamic_classes_names.append(class_name)

    def get_global_css(self):
        return """
        <style class="{0}style">
        {1}
        </style>
        """.format(self.class_name_prefix, '\n'.join(self.classes))
