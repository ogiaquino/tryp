from time import gmtime, strftime
from xlwt import easyxf, Borders, Workbook, Pattern, Style, XFStyle, Font
plus_row = 3

values = {
"!!!!!%!BSR":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "lavender",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x01,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!!!%!percentage":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "lavender",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "0.0\\%",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!!!%BSR":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "lavender",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x01,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!!!%percentage":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "lavender",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "0.0\\%",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!!%!BSR":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "sky-blue",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!!%BSR":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "sky-blue",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!!%!percentage":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "sky-blue",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "0.0\\%",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!!%percentage":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "sky-blue",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "0.0\\%",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%!BSR":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "gold",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%BSR":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "gold",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%!percentage":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "gold",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "0.0\\%",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%percentage":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "gold",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "0.0\\%",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%!BSR":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "light-yellow",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%BSR":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "light-yellow",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%!percentage":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "light-yellow",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "0.0\\%",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%percentage":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "light-yellow",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "0.0\\%",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%percentage":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "0.0\\%",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%BSR":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%!percentage":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "0.0\\%",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%!BSR":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
}



rows_total = {
"!!!!!":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "lavender",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x02,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Center"
    },
"!!!!":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "sky-blue",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Left"
    },
"!!!":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "gold",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Left"
    },
"!!":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "light-yellow",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Left"
    }
}

rows_labels= {
0:
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Top, horizontal Center"
    },
1:
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Top, horizontal Left"
    },
2:
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Top, horizontal Left"
    },
3:
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Left"
    },
4:
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Left"
    }
}

values_labels= {
0:
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x01,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Center"
    },
1:
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x02,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Left"
    },
}

columns_labels= {
0:
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x02,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Center"
    },
}

columns_total = {
"!":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "lime",
       "bold": "on",
       "height": "160"
      },
     "border":
      {"left": 0x02,
       "right": 0x02,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Center"
    },
}

def headers(ws, connection=None, crosstab=None):
    ws.row(3).height = 700
    ws.row(4).height = 700
    ws.set_panes_frozen(True)
    ws.set_horz_split_pos(5)
    ws.set_vert_split_pos(5)
    ws.col(0).width = 3200
    ws.col(2).width = 4700
    ws.col(3).width = 4000
    ws.col(4).width = 7600
    ws.show_grid = False
    xf_str = 'font: name sans-serif, color black,' \
             'bold on, height 160;' \
             'pattern: pattern solid, fore-colour white; '
    #xf_str = xf_str + styles[name]['alignment']
    exf = easyxf(xf_str)
    ws.write(0,0,'BAD STOCK RETURN', exf)
    ws.write(1,0,'WEST MALAYSIA & EAST MALAYSIA', exf)
    now = strftime("%d-%b-%Y")
    ws.write(2,0,now, exf)
