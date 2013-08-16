from xlrd import open_workbook
from xlwt import easyxf, Borders, Pattern, Style

template = "report/inventory_by_sku_templates.xls"
wb = open_workbook(template, formatting_info=True)
ws = wb.sheet_by_index(0)

colour = {
    46: "lavender",
    43: "light-yellow",
    51: "gold",
    64: "white"
}


def get_values_styles(ct):
    yaxis = [''] + ct.visible_yaxis_summary + [ct.yaxis[-1]]
    xaxis = [''] + ct.xaxis
    styles = {}

    for i, y in enumerate(yaxis):
        col = -1
        for x in xaxis:
            for z in ct.zaxis:
                col = col + 1
                styles[(y, x, z)] = get_styles(i + len(ct.xaxis) + 1,
                                               col + len(ct.yaxis))
    return styles


def get_styles(row, col):
    xfi = wb.xf_list[ws.cell_xf_index(row, col)]
    xf = 'font: name %(name)s, height %(height)s;' \
         'pattern: pattern solid, fore-colour %(forecolour)s;' \
         'alignment: vertical %(vertical)s, horizontal %(horizontal)s' \
          % font(xfi)
    style = easyxf(xf)
    style.borders = borders(xfi)
    style.num_format_str = number_format(xfi)
    return style


def font(xfi):
    name = wb.font_list[xfi.font_index].name
    forecolour = colour[xfi.background.pattern_colour_index]
    height = wb.font_list[xfi.font_index].height
    horizontal = alignment(xfi)['horizontal']
    vertical = alignment(xfi)['vertical']
    return {'name': name, 'forecolour': forecolour, 'height': height,
            'horizontal': horizontal, 'vertical': vertical}


def borders(xfi):
    brd = Borders()
    bottom = xfi.border.bottom_line_style.real
    left = xfi.border.left_line_style.real
    right = xfi.border.right_line_style.real
    top = xfi.border.top_line_style.real
    borders = {'bottom': bottom, 'left': left, 'right': right, 'top': top}
    brd.bottom = borders["bottom"]
    brd.left = borders["left"]
    brd.right = borders["right"]
    brd.top = borders["top"]
    return brd


def number_format(xfi):
    return wb.format_map[xfi.format_key].format_str


def alignment(xfi):
    horz_align = Style.xf_dict['alignment']['horz']
    horz_align = dict(zip(horz_align.values(), horz_align.keys()))
    vert_align = Style.xf_dict['alignment']['vert']
    vert_align = dict(zip(vert_align.values(), vert_align.keys()))
    horizontal = horz_align[xfi.alignment.hor_align]
    vertical = vert_align[xfi.alignment.vert_align]
    return {'horizontal': horizontal, 'vertical': vertical}
