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


def headers(ws, connection=None, crosstab=None):
    def selling_days(connection):
        if connection:
            now = strftime("%Y-%m-%d")
            query = """
                SELECT
                    count(*)
                FROM
                    dim_calendar
                WHERE
                    date BETWEEN date_trunc('mon', coalesce('%s', now())::date)
                    AND date_trunc('mon', coalesce('%s', now())::date) +
                        '1 mon'::interval - '1 day'::interval
                AND principal_code = 'kraft'
                AND holiday = 'N'
            """ % (now, now)
            df = psql.frame_query(query, con=connection)
            return df['count'][0]
        return 1

    def elapse_days(connection):
        if connection:
            now = strftime("%Y-%m-%d")
            query = """
                SELECT
                    count(*)
                FROM
                    dim_calendar
                WHERE
                    date
                    BETWEEN date_trunc('mon', coalesce('%s', now())::date)
                    AND coalesce('%s', now())::date
                    AND principal_code = 'kraft'
                    AND holiday = 'N'
            """ % (now, now)
            df = psql.frame_query(query, con=connection)
            return df['count'][0]
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
    #xf_str = xf_str + styles[name]['alignment']
    exf = easyxf(xf_str)
    ws.write(0,0,'SELL OUT - DAILY SALES REPORT', exf)
    ws.write(1,0,'DISTRIBUTOR', exf)
    now = strftime("%d-%b-%Y")
    ws.write(2,0,now, exf)

    selling_days = selling_days(connection)
    elapse_days = elapse_days(connection)
    elapse_percentage = (elapse_days / selling_days) * 100
    ws.write(0,3,'Selling Days', exf)
    ws.write(1,3,'Days Elapse', exf)
    ws.write(2,3,'Elapse %', exf)
    ws.write(0,4, str(selling_days), exf)
    ws.write(1,4, str(elapse_days), exf)
    ws.write(2,4, str(round(elapse_percentage,1)) + '%', exf)
    return {'elapse_percentage': elapse_percentage}
