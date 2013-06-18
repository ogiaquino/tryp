from __future__ import division
import psycopg2
import pandas.io.sql as psql
from time import gmtime, strftime
from xlwt import easyxf, Borders, Workbook, Pattern, Style, XFStyle, Font

plus_row = 5

columns_total_labels = {
'!!': ' TOTAL',
'!' : ''
}


values = {
"!!!!%Sales Qty":
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
"!!!!%SOH Qty":
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
"!!!!%DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!!!%Sales Value":
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
"!!!!%SOH Value":
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
"!!!!%DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!!!%Sales Volume":
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
"!!!!%SOH Volume":
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
"!!!!%DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
    
    
"!!!%Sales Qty":
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
"!!!%SOH Qty":
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
"!!!%DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!!%Sales Value":
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
"!!!%SOH Value":
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
"!!!%DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!!%Sales Volume":
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
"!!!%SOH Volume":
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
"!!!%DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!%Sales Qty":
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
"!!%SOH Qty":
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
"!!%DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!%Sales Value":
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
"!!%SOH Value":
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
"!!%DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!%Sales Volume":
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
"!!%SOH Volume":
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
"!!%DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
    
"!!!!%!Sales Qty":
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
"!!!!%!SOH Qty":
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
"!!!!%!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!!!%!Sales Value":
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
"!!!!%!SOH Value":
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
"!!!!%!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!!!%!Sales Volume":
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
"!!!!%!SOH Volume":
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
"!!!!%!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
    
    
"!!!%!Sales Qty":
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
"!!!%!SOH Qty":
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
"!!!%!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!!%!Sales Value":
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
"!!!%!SOH Value":
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
"!!!%!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!!%!Sales Volume":
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
"!!!%!SOH Volume":
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
"!!!%!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!%!Sales Qty":
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
"!!%!SOH Qty":
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
"!!%!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!%!Sales Value":
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
"!!%!SOH Value":
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
"!!%!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!%!Sales Volume":
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
"!!%!SOH Volume":
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
"!!%!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
    
"!!!!%!!Sales Qty":
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
"!!!!%!!SOH Qty":
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
"!!!!%!!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!!!%!!Sales Value":
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
"!!!!%!!SOH Value":
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
"!!!!%!!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!!!%!!Sales Volume":
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
"!!!!%!!SOH Volume":
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
"!!!!%!!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
    
    
"!!!%!!Sales Qty":
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
"!!!%!!SOH Qty":
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
"!!!%!!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!!%!!Sales Value":
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
"!!!%!!SOH Value":
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
"!!!%!!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!!%!!Sales Volume":
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
"!!!%!!SOH Volume":
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
"!!!%!!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!%!!Sales Qty":
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
"!!%!!SOH Qty":
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
"!!%!!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!%!!Sales Value":
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
"!!%!!SOH Value":
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
"!!%!!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!%!!Sales Volume":
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
"!!%!!SOH Volume":
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
"!!%!!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
"!!!!%!!!Sales Qty":
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
"!!!!%!!!SOH Qty":
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
"!!!!%!!!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!!!%!!!Sales Value":
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
"!!!!%!!!SOH Value":
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
"!!!!%!!!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!!!%!!!Sales Volume":
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
"!!!!%!!!SOH Volume":
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
"!!!!%!!!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
    
    
"!!!%!!!Sales Qty":
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
"!!!%!!!SOH Qty":
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
"!!!%!!!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!!%!!!Sales Value":
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
"!!!%!!!SOH Value":
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
"!!!%!!!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!!%!!!Sales Volume":
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
"!!!%!!!SOH Volume":
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
"!!!%!!!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!%!!!Sales Qty":
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

"!!%!!!SOH Qty":
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
"!!%!!!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"!!%!!!Sales Value":
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
"!!%!!!SOH Value":
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
"!!%!!!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"!!%!!!Sales Volume":
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
"!!%!!!SOH Volume":
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
"!!%!!!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
"%!!!Sales Qty":
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

"%!!!SOH Qty":
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

"%!!!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"%!!!Sales Value":
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

"%!!!SOH Value":
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

"%!!!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
"%!!!Sales Volume":
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

"%!!!SOH Volume":
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

"%!!!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

"%!!Sales Qty":
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

"%!!SOH Qty":
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

"%!!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"%!!Sales Value":
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

"%!!SOH Value":
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

"%!!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
"%!!Sales Volume":
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

"%!!SOH Volume":
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

"%!!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
    
    
"%!Sales Qty":
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

"%!SOH Qty":
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

"%!DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"%!Sales Value":
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

"%!SOH Value":
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

"%!DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
"%!Sales Volume":
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

"%!SOH Volume":
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

"%!DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
"%Sales Qty":
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

"%SOH Qty":
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

"%DOSC Qty":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

    
"%Sales Value":
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

"%SOH Value":
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

"%DOSC Value":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
"%Sales Volume":
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

"%SOH Volume":
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

"%DOSC Volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
}



rows_total = {
"!!!!":
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
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Center; align: wrap on"
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
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Center; align: wrap on"
    },
5:
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
6:
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
7:
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
8:
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
1:
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
2:
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
"!!!":
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
"!!":
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
     "alignment": "alignment: vertical Center, horizontal Center",
    },
"!":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "yellow",
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


def conditional_rows_label(connection, xf):
    xf_str = 'font: name sans-serif, color red,' \
             'bold on, height 160;' \
             'pattern: pattern solid, fore-colour white; '
    borders = Borders()
    borders.top = 0x01
    borders.bottom = 0x01
    borders.left = 0x01
    borders.right = 0x01
    xf_new = easyxf(xf_str)
    xf_new.borders = borders
    return {'labels': {
                '0030599'    :   '0030599 *',
                '0030604'    :   '0030604 *',
                '0030607'    :   '0030607 *',
                '0030589'    :   '0030589 *',
                '0030585'    :   '0030585 *',
                '0030587'    :   '0030587 *',
                '0030580'    :   '0030580 *',
                '0030542'    :   '0030542 *',
                '0788340'    :   '0788340 *',
                '0788342'    :   '0788342 *',
                '0788344'    :   '0788344 *',
                '0030590'    :   '0030590 *',
                '0030591'    :   '0030591 *',
                '0030593'    :   '0030593 *',
                '0030616'    :   '0030616 *',
                '0763700'    :   '0763700 *',
                '0030614'    :   '0030614 *',
                '0030540'    :   '0030540 *',
                '0030541'    :   '0030541 *',
                '0030538'    :   '0030538 *',
                '616126'     :   '616126 *',
                '0603436'    :   '0603436 *',
                '0603442'    :   '0603442 *',
                '0603444'    :   '0603444 *',
                '0784482'    :   '0784482 *',
                '0603438'    :   '0603438 *',
                '0030633'    :   '0030633 *',
                '0030623'    :   '0030623 *',
                '0030628'    :   '0030628 *',
                '0961983'    :   '0961983 *',
                '0030661'    :   '0030661 *',
                '0030651'    :   '0030651 *',
                '0856087'    :   '0856087 *',
                '0856083'    :   '0856083 *',
                '0030621'    :   '0030621 *',
                '0030625'    :   '0030625 *',
                '0030630'    :   '0030630 *',
                '0030643'    :   '0030643 *',
                '0030648'    :   '0030648 *',
                '0030656'    :   '0030656 *',
                '0921512'    :   '0921512 *',
                '0030646'    :   '0030646 *',
                '0030682'    :   '0030682 *',
                '0030694'    :   '0030694 *',
                '0030639'    :   '0030639 *',
                '0030664'    :   '0030664 *',
                '616129'     :   '616129 *',
                '616130'     :   '616130 *',
                '616131'     :   '616131 *',
                '616133'     :   '616133 *',
                '616134'     :   '616134 *',
                '616137'     :   '616137 *',
                '616135'     :   '616135 *',
                '616136'     :   '616136 *',
                '616091'     :   '616091 *',
                '616095'     :   '616095 *',
                '616140'     :   '616140 *',
                '616141'     :   '616141 *',
                '616143'     :   '616143 *',
                '859469'     :   '859469 *',
                '917904'     :   '917904 *',
                '859471'     :   '859471 *',
                '616114'     :   '616114 *',
                '616115'     :   '616115 *',
                '616107'     :   '616107 *',
                '859477'     :   '859477 *',
                '917968'     :   '917968 *',
                '917971'     :   '917971 *',
                '615481'     :   '615481 *',
                '960175'     :   '960175 *',
                '615476'     :   '615476 *',
                '610131'     :   '610131 *',
                '610132'     :   '610132 *',
                '610133'     :   '610133 *',
                '917987'     :   '917987 *',
                '673748'     :   '673748 *',
                '673749'     :   '673749 *',
                '673750'     :   '673750 *',
                '610139'     :   '610139 *',
                '0030599'    :   '0030599 *',
                '0030607'    :   '0030607 *',
                '0030589'    :   '0030589 *',
                '0030585'    :   '0030585 *',
                '0030580'    :   '0030580 *',
                '0030542'    :   '0030542 *',
                '0788340'    :   '0788340 *',
                '0030590'    :   '0030590 *',
                '0030591'    :   '0030591 *',
                '0030593'    :   '0030593 *',
                '0030616'    :   '0030616 *',
                '0763700'    :   '0763700 *',
                '0030614'    :   '0030614 *',
                '0030540'    :   '0030540 *',
                '0030541'    :   '0030541 *',
                '0030538'    :   '0030538 *',
                '616126'     :   '616126 *',
                '0603436'    :   '0603436 *',
                '0603442'    :   '0603442 *',
                '0603444'    :   '0603444 *',
                '0784482'    :   '0784482 *',
                '0603438'    :   '0603438 *',
                '0030633'    :   '0030633 *',
                '0030623'    :   '0030623 *',
                '0030628'    :   '0030628 *',
                '0030661'    :   '0030661 *',
                '0030651'    :   '0030651 *',
                '0856087'    :   '0856087 *',
                '0856083'    :   '0856083 *',
                '0030621'    :   '0030621 *',
                '0030625'    :   '0030625 *',
                '0030630'    :   '0030630 *',
                '0030648'    :   '0030648 *',
                '0030656'    :   '0030656 *',
                '0921512'    :   '0921512 *',
                '0030646'    :   '0030646 *',
                '0030682'    :   '0030682 *',
                '616129'     :   '616129 *',
                '616130'     :   '616130 *',
                '616131'     :   '616131 *',
                '616133'     :   '616133 *',
                '616134'     :   '616134 *',
                '616137'     :   '616137 *',
                '616135'     :   '616135 *',
                '616136'     :   '616136 *',
                '616091'     :   '616091 *',
                '616095'     :   '616095 *',
                '616140'     :   '616140 *',
                '616141'     :   '616141 *',
                '616143'     :   '616143 *',
                '859469'     :   '859469 *',
                '616115'     :   '616115 *',
                '616107'     :   '616107 *',
                '917807'     :   '917807 *',
                '859477'     :   '859477 *',
                '917968'     :   '917968 *',
                '917971'     :   '917971 *',
                '615481'     :   '615481 *',
                '960175'     :   '960175 *',
                '610131'     :   '610131 *',
                '610132'     :   '610132 *',
                '610133'     :   '610133 *',
                '917987'     :   '917987 *',
                '673748'     :   '673748 *',
                '673749'     :   '673749 *',
                '673750'     :   '673750 *',
                '610139'     :   '610139 *',
                '0030599'    :   '0030599 *',
                '0030607'    :   '0030607 *',
                '0030590'    :   '0030590 *',
                '0030591'    :   '0030591 *',
                '0030593'    :   '0030593 *',
                '0030619'    :   '0030619 *',
                '0784052'    :   '0784052 *',
                '0766211'    :   '0766211 *',
                '0766200'    :   '0766200 *',
                '0030616'    :   '0030616 *',
                '0030614'    :   '0030614 *',
                '0603436'    :   '0603436 *',
                '0603442'    :   '0603442 *',
                '0030633'    :   '0030633 *',
                '0030623'    :   '0030623 *',
                '0030628'    :   '0030628 *',
                '0030630'    :   '0030630 *',
                '0030646'    :   '0030646 *',
                '616105'     :   '616105 *',
                '616129'     :   '616129 *',
                '616130'     :   '616130 *',
                '616134'     :   '616134 *',
                '616135'     :   '616135 *',
                '859469'     :   '859469 *',
                '616107'     :   '616107 *',
                '917807'     :   '917807 *',
                '859477'     :   '859477 *',
                '917968'     :   '917968 *',
                '616144'     :   '616144 *',
                '673748'     :   '673748 *',
                '673749'     :   '673749 *',
                '673750'     :   '673750 *',
                '610139'     :   '610139 *',
                '610146'     :   '610146 *',
                '610145'     :   '610145 *',
                '0030590'    :   '0030590 *',
                '0030593'    :   '0030593 *',
                '0030619'    :   '0030619 *',
                '0030595'    :   '0030595 *',
                '0030592'    :   '0030592 *',
                '0784052'    :   '0784052 *',
                '0766211'    :   '0766211 *',
                '0766200'    :   '0766200 *',
                '0766203'    :   '0766203 *',
                '0604509'    :   '0604509 *',
                '0030631'    :   '0030631 *',
                '0030633'    :   '0030633 *',
                '0030623'    :   '0030623 *',
                '0030628'    :   '0030628 *',
                '0030646'    :   '0030646 *',
                '616105'     :   '616105 *',
                '616146'     :   '616146 *',
                '616106'     :   '616106 *',
                '616144'     :   '616144 *',
                '673748'     :   '673748 *',
                '673749'     :   '673749 *',
                '673750'     :   '673750 *',
                '610146'     :   '610146 *',
                '610145'     :   '610145 *',
                '0030619'    :   '0030619 *',
                '0030592'    :   '0030592 *',
                '0766211'    :   '0766211 *',
                '0766200'    :   '0766200 *',
                '0604509'    :   '0604509 *',
                '0030631'    :   '0030631 *',
                '0030627'    :   '0030627 *',
                '0030633'    :   '0030633 *',
                '616146'     :   '616146 *',
                '616106'     :   '616106 *',
                '616144'     :   '616144 *',
                '673750'     :   '673750 *',
                '0030580'    :   '0030580 *',
                '0030542'    :   '0030542 *',
                '0030590'    :   '0030590 *',
                '0030619'    :   '0030619 *',
                '0766211'    :   '0766211 *',
                '0766200'    :   '0766200 *',
                '616128'     :   '616128 *',
                '0603436'    :   '0603436 *',
                '0030633'    :   '0030633 *',
                '0030623'    :   '0030623 *',
                '0030657'    :   '0030657 *',
                '0030621'    :   '0030621 *',
                '0030625'    :   '0030625 *',
                '0030630'    :   '0030630 *',
                '0030646'    :   '0030646 *',
                '616129'     :   '616129 *',
                '616130'     :   '616130 *',
                '616131'     :   '616131 *',
                '616133'     :   '616133 *',
                '616134'     :   '616134 *',
                '616135'     :   '616135 *',
                '616136'     :   '616136 *',
                '859471'     :   '859471 *',
                '859474'     :   '859474 *',
                '616106'     :   '616106 *',
                '917807'     :   '917807 *',
                '615481'     :   '615481 *',
                '960175'     :   '960175 *',
                '615476'     :   '615476 *',
                '610131'     :   '610131 *',
                '610132'     :   '610132 *',
                '610133'     :   '610133 *',
                '673748'     :   '673748 *',
                '673749'     :   '673749 *',
                '673750'     :   '673750 *',
                '610139'     :   '610139 *',
                '617709'     :   '617709 *',
                '617712'     :   '617712 *',
            
           }, 'xf': xf_new}


def headers(ws, tryp):
    ws.row(7).height = 700
    ws.set_panes_frozen(True)
    ws.set_horz_split_pos(9)
    ws.set_vert_split_pos(4)
    ws.col(1).width = 6300
    ws.col(2).width = 3500
    ws.col(3).width = 10500
    ws.row(8).height = 1100 
    for i in range(len(tryp.crosstab.values[0])):
        ws.col(i + 4).width = 2800
    ws.show_grid = False
    xf_str = 'font: name sans-serif, color black,' \
             'bold on, height 160;' \
             'pattern: pattern solid, fore-colour white; '
    exf = easyxf(xf_str)
    ws.write_merge(0,0,0,2,'Inventory by SKU', exf)
    ws.write_merge(1,1,0,2,'Days Of Stock Covered = Stock On Hand/Avg Daily Sales', exf)
    now = strftime("%d-%b-%Y")
    ws.write_merge(2,2,0,2,now, exf)
    ws.write_merge(3,3,0,2,'Last 75 days', exf)

    
    xf_str = 'font: name sans-serif, color red,' \
             'bold on, height 160;' \
             'pattern: pattern solid, fore-colour white; '
    exf = easyxf(xf_str)
    ws.write_merge(4,4,0,2,'* MSL', exf)

    
    #merge the corner
    style = easyxf('borders: top medium;')
    ws.write_merge(0 + tryp.plus_row, len(tryp.columns)+tryp.plus_row, 0,
                   len(tryp.rows)-1, '', style)

    #borderize thick the last row
    for i in range(len(tryp.rows) + len(tryp.crosstab.values[0])):
        ws.write(len(tryp.crosstab.values)+ len(tryp.columns) + tryp.plus_row + 1, i, '', style)
