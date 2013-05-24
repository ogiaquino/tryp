from xlwt import easyxf, Borders, Pattern, Style


def borders_style(borders):
    brd = Borders()
    brd.bottom = borders["bottom"]
    brd.left = borders["left"]
    brd.right= borders["right"]
    brd.top = borders["top"]
    return brd

def get_styles(reportname, axis):
    module = __import__('%s.styles' % reportname, fromlist=[axis])
    styles = getattr(module, axis)
    crosstab_styles = {}
    for name, style in styles.iteritems():
        xf_str = 'font: name %(name)s, color %(color)s,' \
                 'bold %(bold)s, height %(height)s;' \
                 'pattern: pattern solid, fore-colour %(background)s;' % styles[name]['font']
        xf_str = xf_str + styles[name]['alignment']
        exf = easyxf(xf_str)
        exf.borders = borders_style(styles[name]['border'])
        exf.num_format_str = styles[name]['number_format_str']
        crosstab_styles[name] = exf
    return crosstab_styles

def get_labels(reportname, axis):
    try:
        module = __import__('%s.styles' % reportname, fromlist=[axis])
        labels = getattr(module, axis)
        return labels
    except AttributeError:
        return {}
