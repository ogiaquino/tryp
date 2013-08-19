from xlwt import easyxf, Borders, Pattern, Style

colour = {
    46: "lavender",
    43: "light-yellow",
    51: "gold",
    64: "white",
    50: "lime",
    13: "yellow",
    40: "sky-blue"
}


def styles(ct, wbt, wst):
    style = {}
    style['values'] = get_values_styles(ct, wbt, wst)
    style['index'] = get_index_styles(ct, wbt, wst)
    style['column'] = get_column_styles(ct, wbt, wst)
    style['values_labels'] = get_values_labels_styles(ct, wbt, wst)
    return style


def get_values_styles(ct, wbt, wst):
    yaxis = [''] + ct.visible_yaxis_summary + [ct.yaxis[-1]]
    xaxis = [''] + ct.xaxis
    styles = {}

    for i, y in enumerate(yaxis):
        col = -1
        for x in xaxis:
            for z in ct.zaxis:
                col = col + 1
                styles[(y, x, z)] = get_styles(i + len(ct.xaxis) + 1,
                                               col + len(ct.yaxis), wbt, wst)
    return styles


def get_index_styles(ct, wbt, wst):
    yaxis = [''] + ct.visible_yaxis_summary + [ct.yaxis[-1]]
    xaxis = [''] + ct.xaxis
    styles = {}

    for i, y in enumerate(yaxis):
        for j in range(len(yaxis)):
            styles[(y, j)] = get_styles(i + len(ct.xaxis) + 1, j, wbt, wst)
    return styles


def get_values_labels_styles(ct, wbt, wst):
    styles = {}
    for i in range(len(ct.zaxis)):
        styles[ct.zaxis[i]] = get_styles(len(ct.xaxis), i + len(ct.yaxis),
                                         wbt, wst)
    return styles


def get_column_styles(ct, wbt, wst):
    yaxis = [''] + ct.visible_yaxis_summary + [ct.yaxis[-1]]
    xaxis = [''] + ct.xaxis
    styles = {}

    for h in range(len(ct.xaxis)):
        col = len(ct.yaxis) - 1
        for i, x in enumerate(xaxis):
            for j, z in enumerate(ct.zaxis):
                col = col + 1
                styles[(h, x, z)] = get_styles(h, col, wbt, wst)

    return styles


def get_styles(row, col, wbt, wst):
    xf = wbt.xf_list[wst.cell_xf_index(row, col)]
    xfval = dict(font(xf, wbt) + pattern(xf) + alignment(xf) + borders(xf))
    xfstr = 'font: name %(name)s, height %(height)s, bold %(bold)s;' \
            'pattern: pattern solid, fore-colour %(forecolour)s;' \
            'alignment: vertical %(vertical)s, horizontal %(horizontal)s;' \
            'borders : bottom %(bottom)s, left %(left)s,'\
            'right %(right)s, top %(top)s' % xfval
    style = easyxf(xfstr)
    style.num_format_str = number_format(xf, wbt)
    return style


def font(xf, wbt):
    name = wbt.font_list[xf.font_index].name
    height = wbt.font_list[xf.font_index].height
    bold = wbt.font_list[xf.font_index].weight
    bold = 'on' if bold == 700 else 'off'
    return (('name', name), ('height', height), ('bold', bold))


def pattern(xf):
    forecolour = colour[xf.background.pattern_colour_index]
    return (('forecolour', forecolour),)


def alignment(xf):
    horz_align = Style.xf_dict['alignment']['horz']
    horz_align = dict(zip(horz_align.values(), horz_align.keys()))
    vert_align = Style.xf_dict['alignment']['vert']
    vert_align = dict(zip(vert_align.values(), vert_align.keys()))
    horizontal = horz_align[xf.alignment.hor_align]
    vertical = vert_align[xf.alignment.vert_align]
    return (('horizontal', horizontal), ('vertical', vertical))


def borders(xf):
    brd = Borders()
    bottom = xf.border.bottom_line_style.real
    left = xf.border.left_line_style.real
    right = xf.border.right_line_style.real
    top = xf.border.top_line_style.real
    return (('bottom', bottom), ('left', left), ('right', right), ('top', top))


def number_format(xf, wbt):
    return wbt.format_map[xf.format_key].format_str
