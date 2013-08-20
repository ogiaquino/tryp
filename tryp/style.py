from xlrd import open_workbook
from xlwt import easyxf, Borders, Pattern, Style as XStyle

colour = {
    46: "lavender",
    43: "light-yellow",
    51: "gold",
    64: "white",
    50: "lime",
    13: "yellow",
    40: "sky-blue"
}


class Style(object):
    def __init__(self, ct):
        self.wb = open_workbook(ct.excel['template'], formatting_info=True)
        self.ws = self.wb.sheet_by_index(0)
        note_map = self.ws.cell_note_map
        note_map = dict([(note_map[k].text, k) for k in note_map.keys()])
        self.crosstab_row = note_map['CROSSTAB'][0]
        self.crosstab_col = note_map['CROSSTAB'][1]

        self.ytotal_label = self.ws.cell(self.crosstab_row + len(ct.xaxis) + 1,
                            self.crosstab_col).value

        self.xtotal_label = self.ws.cell(self.crosstab_row,
                            self.crosstab_col + len(ct.yaxis)).value

        print self.xtotal_label, self.ytotal_label

        self.values = self.get_values_styles(ct)
        self.index = self.get_index_styles(ct)
        self.column = self.get_column_styles(ct)
        self.values_labels = self.get_values_labels_styles(ct)
        self.corner = self.get_corner_styles(ct)

    def get_corner_styles(self, ct):
        styles = {}
        return self.get_styles(self.crosstab_row - 1, self.crosstab_col - 1)

    def get_values_styles(self, ct):
        yaxis = [''] + ct.visible_yaxis_summary + [ct.yaxis[-1]]
        xaxis = [''] + ct.xaxis
        styles = {}

        for i, y in enumerate(yaxis):
            col = -1
            for x in xaxis:
                for z in ct.zaxis:
                    col = col + 1
                    styles[(y, x, z)] = self.get_styles(i + len(ct.xaxis) + 1,
                                                        col + len(ct.yaxis))
        return styles

    def get_index_styles(self, ct):
        yaxis = [''] + ct.visible_yaxis_summary + [ct.yaxis[-1]]
        xaxis = [''] + ct.xaxis
        styles = {}

        for i, y in enumerate(yaxis):
            for j in range(len(yaxis)):
                styles[(y, j)] = self.get_styles(i + len(ct.xaxis) + 1, j)
        return styles

    def get_values_labels_styles(self, ct):
        styles = {}
        for i in range(len(ct.zaxis)):
            styles[ct.zaxis[i]] = self.get_styles(len(ct.xaxis), i +
                                                  len(ct.yaxis))
        return styles

    def get_column_styles(self, ct):
        yaxis = [''] + ct.visible_yaxis_summary + [ct.yaxis[-1]]
        xaxis = [''] + ct.xaxis
        styles = {}

        for h in range(len(ct.xaxis)):
            col = len(ct.yaxis) - 1
            for i, x in enumerate(xaxis):
                for j, z in enumerate(ct.zaxis):
                    col = col + 1
                    styles[(h, x, z)] = self.get_styles(h, col)

        return styles

    def get_styles(self, row, col):
        row = row + self.crosstab_row
        col = col + self.crosstab_col
        xf = self.wb.xf_list[self.ws.cell_xf_index(row, col)]
        xfval = dict(self.font(xf) + self.pattern(xf) + self.alignment(xf) +
                     self.borders(xf))
        xfstr = 'font: name %(name)s, height %(height)s, bold %(bold)s;'\
                'pattern: pattern solid, fore-colour %(forecolour)s;'\
                'alignment: vertical %(vertical)s, horizontal %(horizontal)s;'\
                'borders : bottom %(bottom)s, left %(left)s,'\
                'right %(right)s, top %(top)s' % xfval
        style = easyxf(xfstr)
        style.num_format_str = self.number_format(xf)
        return style

    def font(self, xf):
        name = self.wb.font_list[xf.font_index].name
        height = self.wb.font_list[xf.font_index].height
        bold = self.wb.font_list[xf.font_index].weight
        bold = 'on' if bold == 700 else 'off'
        return (('name', name), ('height', height), ('bold', bold))

    def pattern(self, xf):
        forecolour = colour[xf.background.pattern_colour_index]
        return (('forecolour', forecolour),)

    def alignment(self, xf):
        horz_align = XStyle.xf_dict['alignment']['horz']
        horz_align = dict(zip(horz_align.values(), horz_align.keys()))
        vert_align = XStyle.xf_dict['alignment']['vert']
        vert_align = dict(zip(vert_align.values(), vert_align.keys()))
        horizontal = horz_align[xf.alignment.hor_align]
        vertical = vert_align[xf.alignment.vert_align]
        return (('horizontal', horizontal), ('vertical', vertical))

    def borders(self, xf):
        brd = Borders()
        bottom = xf.border.bottom_line_style.real
        left = xf.border.left_line_style.real
        right = xf.border.right_line_style.real
        top = xf.border.top_line_style.real
        return (('bottom', bottom), ('left', left), ('right', right),
                ('top', top))

    def number_format(self, xf):
        return self.wb.format_map[xf.format_key].format_str
