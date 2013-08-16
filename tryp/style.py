from xlrd import open_workbook
from xlwt import easyxf, Borders, Pattern, Style

template = "report/inventory_by_sku_templates.xls"
wb = open_workbook(template, formatting_info=True)
ws = wb.sheet_by_index(0)

colour = {
    46: "lavender",
    51: "gold",
    43: "light-yellow",
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
    xf = 'font: name %(name)s;' \
         'pattern: pattern solid, fore-colour %(forecolour)s;' % font(xfi)
    style = easyxf(xf)
    return style


def font(xfi):
    name = wb.font_list[xfi.font_index].name
    forecolour = colour[xfi.background.pattern_colour_index]
    return {'name': name, 'forecolour': forecolour}
