from __future__ import division
import pandas.io.sql as psql
from time import gmtime, strftime
from xlwt import easyxf, Borders, Workbook, Pattern, Style, XFStyle, Font

plus_row = 3

values = {
"!!!!!%!Target":
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
"!!!!!%!Sell Out Actual":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "lavender",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x01,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!!!%!Ach":
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
"!!!!!%Target":
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
"!!!!!%Sell Out Actual":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "lavender",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x01,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!!!%Ach":
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
"!!!!%!Target":
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
"!!!!%Target":
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
"!!!!%!Sell Out Actual":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "sky-blue",
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
"!!!!%!Ach":
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
"!!!!%Sell Out Actual":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "sky-blue",
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
"!!!!%Ach":
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
"!!!%!Target":
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
"!!!%Target":
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
"!!!%!Sell Out Actual":
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
"!!!%!Ach":
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
"!!!%Sell Out Actual":
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
"!!!%Ach":
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
"!!%!Target":
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
"!!%Target":
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
"!!%!Sell Out Actual":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "light-yellow",
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
"!!%!Ach":
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
"!!%Sell Out Actual":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "light-yellow",
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
"!!%Ach":
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
"%Sell Out Actual":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
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
"%Ach":
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
"%Target":
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
"%!Sell Out Actual":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "white",
       "bold": "off",
       "height": "160"
      },
     "border":
      {"left": 0x01,
       "right": 0x01,
       "top": 0x01,
       "bottom": 0x01
      },
     "number_format_str": "0.0\\%",
     "number_format_str": "#,###0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%!Ach":
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
"%!Target":
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
     "alignment": "alignment: vertical Center, horizontal Center; align: wrap on"
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
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Center; align: wrap on"
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
       "right": 0x02,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Center; align: wrap on"
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


def conditional_formatting(xf, header_info, value):
    if xf.num_format_str == '0.0\%':
        ep = float(header_info['elapse_percentage'])
        xf.num_format_str = "[color10][>{0}]0.0\\%;[Red][<{1}]0.0\\%".format(ep, ep)
    return xf


def headers(ws, tryp):
    def selling_days():
        return 1

    def elapse_days():
        return 1
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
    exf = easyxf(xf_str)
    ws.write_merge(0,0,0,2,'SELL OUT - DAILY SALES REPORT', exf)
    ws.write_merge(1,1,0,2,'DISTRIBUTOR', exf)
    now = strftime("%d-%b-%Y")
    ws.write_merge(2,2,0,2,now, exf)

    selling_days = selling_days()
    elapse_days = elapse_days()
    elapse_percentage = (elapse_days / selling_days) * 100
    ws.write(0,3,'Selling Days', exf)
    ws.write(1,3,'Days Elapse', exf)
    ws.write(2,3,'Elapse %', exf)
    ws.write(0,4, str(selling_days), exf)
    ws.write(1,4, str(elapse_days), exf)
    ws.write(2,4, str(round(elapse_percentage,1)) + '%', exf)

    #merge the corner
    style = easyxf('borders: top medium;')
    ws.write_merge(0 + tryp.plus_row, len(tryp.columns)+tryp.plus_row, 0,
                   len(tryp.rows)-1, '', style)

    #borderize thick the last row
    for i in range(len(tryp.rows) + len(tryp.crosstab.values[0])):
        ws.write(len(tryp.crosstab.values)+ len(tryp.columns) + tryp.plus_row + 1, i, '', style)
    return {'elapse_percentage': elapse_percentage}
