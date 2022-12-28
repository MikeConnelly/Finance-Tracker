from typing import Dict


DEFAULT_EXPENSES_STYLES = [
    {
        "header": {
            "bg_color": "#e16463"
        },
        "data": {
            "bg_color": "#f4cdcc"
        },
        "alt": {
            "bg_color": "#d3b1af"
        },
        "total": {
            "bg_color": "#e16463"
        },
        "line": {
            "color": "#e16463"
        }
    },
    {
        "header": {
            "bg_color": "#92c57c"
        },
        "data": {
            "bg_color": "#daf2d8"
        },
        "alt": {
            "bg_color": "#acbfaa"
        },
        "total": {
            "bg_color": "#92c57c"
        },
        "line": {
            "color": "#92c57c"
        }
    },
    {
        "header": {
            "bg_color": "#f3b369"
        },
        "data": {
            "bg_color": "#fae5cd"
        },
        "alt": {
            "bg_color": "#c6b6a3"
        },
        "total": {
            "bg_color": "#f3b369"
        },
        "line": {
            "color": "#f3b369"
        }
    },
    {
        "header": {
            "bg_color": "#6f9eeb"
        },
        "data": {
            "bg_color": "#c9dbf9"
        },
        "alt": {
            "bg_color": "#9faec5"
        },
        "total": {
            "bg_color": "#6f9eeb"
        },
        "line": {
            "color": "#6f9eeb"
        }
    },
    {
        "header": {
            "bg_color": "#fad866"
        },
        "data": {
            "bg_color": "#fcf2cd"
        },
        "alt": {
            "bg_color": "#c9c1a3"
        },
        "total": {
            "bg_color": "#fad866"
        },
        "line": {
            "color": "#fad866"
        }
    },
    {
        "header": {
            "bg_color": "#8e7cc2"
        },
        "data": {
            "bg_color": "#dbd2ea"
        },
        "alt": {
            "bg_color": "#aba4b6"
        },
        "total": {
            "bg_color": "#8e7cc2"
        },
        "line": {
            "color": "#8e7cc2"
        }
    }
]


DEFAULT_OVERALL_STYLES = {
    "income": {
        "header": {
            "bg_color": "#92c57c"
        },
        "data": {
            "bg_color": "#daf2d8"
        },
        "alt": {
            "bg_color": "#acbfaa"
        },
        "total": {
            "bg_color": "#92c57c"
        },
        "line": {
            "color": "#92c57c"
        }
    },
    "expenses": {
        "header": {
            "bg_color": "#e16463"
        },
        "data": {
            "bg_color": "#f4cdcc"
        },
        "alt": {
            "bg_color": "#d3b1af"
        },
        "total": {
            "bg_color": "#e16463"
        },
        "line": {
            "color": "#e16463"
        }
    },
    "surplus": {
        "header": {
            "bg_color": "#6f9eeb"
        },
        "total": {
            "bg_color": "#6f9eeb"
        },
        "line": {
            "color": "#6f9eeb"
        }
    }
}


"""
Styles Type Structure
{
    minor_category: {
        cell_type: {
            property: value
        }
    }
}
"""
Styles = Dict[str, Dict[str, Dict[str, str]]]


"""Should really accept the entire FinanceData object. Otherwise I have to repeate this whenever I get an expenses map"""
def merge_styles_with_default_expenses_styles(data: Dict[str, Dict[str, float]], custom_styles: Styles) -> Styles:
    styles: Styles = {}
    timespan = list(data.keys())[0]
    num_defaults_used = 0
    default_expenses_styles_length = len(DEFAULT_EXPENSES_STYLES)
    for category in data[timespan].keys():
        if category in custom_styles.keys():
            styles[category] = custom_styles[category]
        else:
            styles[category] = DEFAULT_EXPENSES_STYLES[num_defaults_used % default_expenses_styles_length]
            num_defaults_used += 1
    return styles


def create_styles_map_for_overall_data(data: Dict[str, Dict[str, Dict[str, float]]]) -> Styles:
    styles = {}
    timespan = list(data.keys())[0]
    for major_category in data[timespan].keys():
        for minor_category in data[timespan][major_category].keys():
            styles[minor_category] = DEFAULT_OVERALL_STYLES.get(major_category)
    styles['Total Income'] = DEFAULT_OVERALL_STYLES.get('income')
    styles['Total Expenses'] = DEFAULT_OVERALL_STYLES.get('expenses')
    styles['Total Surplus'] = DEFAULT_OVERALL_STYLES.get('surplus')
    return styles
