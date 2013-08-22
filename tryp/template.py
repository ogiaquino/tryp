from xlrd import open_workbook
from xlwt import easyxf, Borders, Pattern, Style

colour = {
    8: "black",
    10: "red",
    13: "yellow",
    40: "sky-blue",
    43: "light-yellow",
    46: "lavender",
    50: "lime",
    51: "gold",
    64: "white",
}


class Template(object):
    def __init__(self, ct):
        self.ct = ct
        self.wb = open_workbook(ct.excel['template'], formatting_info=True)
        self.ws = self.wb.sheet_by_index(0)

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

    def get_styles(self, row, col, overwrite={}):
        crosstab_row, crosstab_col = self.crosstab_loc
        xf = self.wb.xf_list[self.ws.cell_xf_index(row + crosstab_row,
                                                   col + crosstab_col)]
        xfval = dict(self.__font(xf) +
                     self.__pattern(xf) +
                     self.__alignment(xf) +
                     self.__borders(xf))
        for k in overwrite:
            xfval[k] = overwrite[k]

        xfstr = 'font: name %(name)s, colour %(colour_index)s,' \
                'height %(height)s, bold %(bold)s;'\
                'pattern: pattern solid, fore-colour %(forecolour)s;'\
                'alignment: vertical %(vertical)s, horizontal %(horizontal)s;'\
                'borders : bottom %(bottom)s, left %(left)s,'\
                'right %(right)s, top %(top)s' % xfval
        style = easyxf(xfstr)

        style.label = ''
        value = self.ws.cell(row + crosstab_row, col + crosstab_col).value
        if isinstance(value, basestring):
            if '[' in value and ']' in value:
                value = value.split('[')[1]
                value = value.replace(']', '')
                style.label = value

        style.row = row
        style.col = col
        style.row_height = self.ws.rowinfo_map[row + crosstab_row].height
        style.column_width = self.ws.computed_column_width(col + crosstab_col)
        style.num_format_str = self.__number_format(xf)
        return style

    def __font(self, xf):
        font = self.wb.font_list[xf.font_index]
        bold = 'on' if font.weight == 700 else 'off'
        colour_index = colour.get(font.colour_index, 8)
        return (('name', font.name), ('height', font.height), ('bold', bold),
                ('colour_index', colour_index))

    def __pattern(self, xf):
        forecolour = colour[xf.background.pattern_colour_index]
        return (('forecolour', forecolour),)

    def __alignment(self, xf):
        horz_align = Style.xf_dict['alignment']['horz']
        horz_align = dict(zip(horz_align.values(), horz_align.keys()))
        vert_align = Style.xf_dict['alignment']['vert']
        vert_align = dict(zip(vert_align.values(), vert_align.keys()))
        horizontal = horz_align[xf.alignment.hor_align]
        vertical = vert_align[xf.alignment.vert_align]
        return (('horizontal', horizontal), ('vertical', vertical))

    def __borders(self, xf):
        brd = Borders()
        bottom = xf.border.bottom_line_style.real
        left = xf.border.left_line_style.real
        right = xf.border.right_line_style.real
        top = xf.border.top_line_style.real
        return (('bottom', bottom), ('left', left), ('right', right),
                ('top', top))

    def __number_format(self, xf):
        return self.wb.format_map[xf.format_key].format_str
