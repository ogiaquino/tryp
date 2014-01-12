import re
from xlrd import open_workbook
from xlwt import easyxf, Borders, Pattern, Style

#TODO: Create a mapping of excel default colours.
colour = {
    8: "black",
    10: "red",
    13: "yellow",
    40: "#00CCFF",
    43: "#FFFF99",
    46: "#CC99FF",
    50: "lime",
    50: "#99CC00",
    51: "#FFCC00",
    64: "white",
     9: "white",
    17: "green",
    15: "cyan",
    14: "magenta",
    12: "blue",
}


class XStyle:
    pass


def font(wb, xf):
    font = wb.font_list[xf.font_index]
    bold = True if font.weight == 700 else False
    colour_index = colour.get(font.colour_index, 'black')
    return (('font_name', font.name), ('font_size', font.height / 20),
            ('bold', bold), ('font_color', colour_index))


def pattern(xf):
    forecolour = colour[xf.background.pattern_colour_index]
    return (('bg_color', forecolour),)


def alignment(xf):
    horz_align = Style.xf_dict['alignment']['horz']
    horz_align = dict(zip(horz_align.values(), horz_align.keys()))
    vert_align = Style.xf_dict['alignment']['vert']
    vert_align = dict(zip(vert_align.values(), vert_align.keys()))
    horizontal = horz_align[xf.alignment.hor_align]
    vertical = vert_align[xf.alignment.vert_align]
    if vertical == 'centre':
        vertical = 'vcenter'
    text_wrapped = str(xf.alignment.text_wrapped)
    wrap = Style.xf_dict['alignment']['wrap'][text_wrapped]
    return (('align', horizontal), ('valign', vertical),
            ('text_wrap', wrap))


def borders(xf):
    bottom = xf.border.bottom_line_style.real
    left = xf.border.left_line_style.real
    right = xf.border.right_line_style.real
    top = xf.border.top_line_style.real
    return (('bottom', bottom), ('left', left), ('right', right),
            ('top', top))


def number_format(wb, xf):
    return (('num_format', wb.format_map[xf.format_key].format_str),)


class Template(object):
    def __init__(self, ct=None, wb=None):
        self.ct = ct
        self.wb = wb
        self.wbt = open_workbook(ct.excel['template'], formatting_info=True)
        self.ws = self.wbt.sheet_by_index(0)

        note_map = self.ws.cell_note_map
        note_map = dict([(note_map[k].text, k) for k in note_map.keys()])
        self.crosstab_loc = (note_map['CROSSTAB'][0], note_map['CROSSTAB'][1])

        self.styles = {}
        self.styles['values'] = self.__get_values_styles()
        self.styles['index'] = self.__get_index_styles()
        self.styles['column'] = self.__get_column_styles()
        self.styles['values_labels'] = self.__get_values_labels_styles()
        self.styles['corner'] = self.__get_ct_corner_styles()
        self.styles['header'] = self.__get_header_styles()

    def get_styles(self, row, col, overwrite={}, num_format=None):
        crosstab_row, crosstab_col = self.crosstab_loc
        xf = self.wbt.xf_list[self.ws.cell_xf_index(row + crosstab_row,
                                                   col + crosstab_col)]
        xstyle = XStyle()
        xformat = dict(font(self.wbt, xf) +
                       borders(xf) +
                       pattern(xf) +
                       alignment(xf) +
                       number_format(self.wbt, xf))
        xstyle = self.wb.add_format(xformat)
        xstyle.row_height = float(self.ws.rowinfo_map[row +
                                  crosstab_row].height / 20)
        xstyle.column_width = float(self.ws.computed_column_width(col +
                                    crosstab_col) / 256)
        xstyle.row = row  # CAN BE USE TO OVERWRITE STYLE IN excel.py
        xstyle.col = col  # CAN BE USE TO OVERWRITE STYLE IN excel.py
        xstyle.label = self.__get_label(row + crosstab_row, col + crosstab_col)
        return xstyle

    def __get_label(self, row, col):
        label = unicode(self.ws.cell(row, col).value)
        match = re.search('\[(.*?)\]', label)
        if match:
            label = re.search('\[(.*?)\]', label).group(1)
        else:
            label = ''
        return label

    def __get_header_styles(self):
        headers = []
        crosstab_row = self.crosstab_loc[0]
        for row in range(0, crosstab_row):
            for col in range(len(self.ws.row(row))):
                h = (row, col, self.ws.row(row)[col].value,
                     self.get_styles(row - crosstab_row, col))
                headers.append(h)
        return headers

    def __get_ct_corner_styles(self):
        styles = {}
        return self.get_styles(0, 0)

    def __get_values_styles(self):
        yaxis = [''] + self.ct.visible_yaxis_summary + [self.ct.yaxis[-1]]
        xaxis = [''] + self.ct.xaxis
        styles = {}

        for i, y in enumerate(yaxis):
            col = -1
            for x in xaxis:
                for z in self.ct.zaxis:
                    col = col + 1
                    sty = self.get_styles(i + len(self.ct.xaxis) + 1,
                                          col + len(self.ct.yaxis))
                    sty.z = z
                    styles[(y, x, z)] = sty
        return styles

    def __get_index_styles(self):
        yaxis = [''] + self.ct.visible_yaxis_summary + [self.ct.yaxis[-1]]
        xaxis = [''] + self.ct.xaxis
        styles = {}

        for i, y in enumerate(yaxis):
            for j in range(len(yaxis)):
                col = j
                row = i + len(self.ct.xaxis) + 1
                sty = self.get_styles(row, col)
                styles[(y, j)] = sty
        return styles

    def __get_values_labels_styles(self):
        styles = {}
        for i in range(len(self.ct.zaxis)):
            sty = self.get_styles(len(self.ct.xaxis), i + len(self.ct.yaxis))
            styles[self.ct.zaxis[i]] = sty
        return styles

    def __get_column_styles(self):
        yaxis = [''] + self.ct.visible_yaxis_summary + [self.ct.yaxis[-1]]
        xaxis = [''] + self.ct.xaxis
        styles = {}

        for h in range(len(self.ct.xaxis)):
            col = len(self.ct.yaxis) - 1
            for i, x in enumerate(xaxis):
                for j, z in enumerate(self.ct.zaxis):
                    col = col + 1
                    styles[(h, x, z)] = self.get_styles(h, col)

        return styles
