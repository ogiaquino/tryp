from __future__ import division
import pandas.io.sql as psql
from time import gmtime, strftime
from xlwt import easyxf, Borders, Workbook, Pattern, Style, XFStyle, Font

plus_row = 5

columns_total_labels = {
'!!': ' TOTAL',
'!' : ''
}


values = {
"!!!%Sales_rm":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

"!!!%Average Sales RM":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

"!!!%SOH_rm":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
"!!!%DOSC RM":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%Sales_ctn":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%Average Sales CTN":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%SOH_ctn":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%DOSC CTN":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%Sales_volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%Average Sales TONNES":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%SOH_volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!!%DOSC TONNES":
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
"!!%Sales_rm":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

"!!%Average Sales RM":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

"!!%SOH_rm":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
"!!%DOSC RM":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%Sales_ctn":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%Average Sales CTN":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%SOH_ctn":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%DOSC CTN":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%Sales_volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%Average Sales TONNES":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%SOH_volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!!%DOSC TONNES":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!%Sales_rm":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

"!%Average Sales RM":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

"!%SOH_rm":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
"!%DOSC RM":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!%Sales_ctn":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!%Average Sales CTN":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!%SOH_ctn":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!%DOSC CTN":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!%Sales_volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!%Average Sales TONNES":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!%SOH_volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"!%DOSC TONNES":
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
"%Sales_rm":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

"%Average Sales RM":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },

"%SOH_rm":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
    
"%DOSC RM":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%Sales_ctn":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%Average Sales CTN":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%SOH_ctn":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%DOSC CTN":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%Sales_volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%Average Sales TONNES":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%SOH_volume":
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
     "number_format_str": "#,###0.0",
     "alignment": "alignment: vertical Center, horizontal right"
    },
"%DOSC TONNES":
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
"!!!":
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
"!!":
    {"font":
      {"name": "sans-serif",
       "color": "black",
       "background": "sky-blue",
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
"!":
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
       "right": 0x01,
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
       "right": 0x01,
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
       "right": 0x01,
       "top": 0x02,
       "bottom": 0x02
      },
     "number_format_str": "",
     "alignment": "alignment: vertical Center, horizontal Center; align: wrap on"
    },
9:
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
10:
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
11:
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


def conditional_formatting(xf, header_info):
    if xf.num_format_str == '0.0\%':
        ep = header_info['elapse_percentage']
        xf.num_format_str = "[GREEN][>{0}]0.0\\%;[RED][<{1}]0.0\\%".format(ep, ep)
    return xf


def conditional_rows_label(connection):
    if connection:
        query = """
                SELECT stock_keeping_unit
                FROM 
                (
                SELECT
                    DISTINCT product_id,            
                    mv_star.principal_code,
                    product.stock_keeping_unit
                FROM
                    mv_star
                JOIN dim_products as product            
                ON
                    product.id = product_id
                AND 
                    product.principal_code = mv_star.principal_code
                
                WHERE
                    now()::date <= valid_to
                    AND now()::date >= valid_from
                ) as tmp
        """
        df = psql.frame_query(query, con=connection)
        return dict([(x[0], x[0] + ' *') for x in df.values])
    else:
        return {}
    

def headers(ws, connection=None, crosstab=None):
    def working_days(connection):
        if connection:
            now = strftime("%Y-%m-%d")
            query = """
                    SELECT 
                        count(*)
                    FROM 
                        dim_calendar 
                    WHERE 
                        date BETWEEN (now()::date - '1 day'::interval - '74 day'::interval) 
                    AND 
                        (now()::date - '1 day'::interval) 
                    AND 
                        principal_code='kraft' 
                    AND 
                        holiday='f';
                    """
            df = psql.frame_query(query, con=connection)
            return df['count'][0]
        return 1
    ws.row(5).height = 700
    ws.set_panes_frozen(True)
    ws.set_horz_split_pos(7)
    ws.set_vert_split_pos(3)
    ws.col(0).width = 3000
    ws.col(2).width = 5000
    #ws.row(7).height = 1100 
    #for i in range(len(crosstab.values[0])):
    #    ws.col(i + 4).width = 2400
    ws.show_grid = False
    xf_str = 'font: name sans-serif, color black,' \
             'bold on, height 160;' \
             'pattern: pattern solid, fore-colour white; '
    exf = easyxf(xf_str)
    ws.write(0,0,'Inventory by Dist', exf)
    wd = str(working_days(connection))
    ws.write(1,0,'Avg Daily Sales = Sales / %s' % wd, exf)
    ws.write(2,0,'Days Of Stock Covered = Stock On Hand / Avg Daily Sales', exf)
    now = strftime("%d-%b-%Y")
    ws.write(3,0,now, exf)
    ws.write(4,0,'Last 75 days (%s working days)' % wd, exf)
