from typing import Dict


DEFAULT_STYLES_LIST = [
    {
        "header": {
            "bg_color": "#e16463",
            "bold": "true"
        },
        "data": {
            "bg_color": "#f4cdcc"
        },
        "alt": {
            "bg_color": "#d3b1af"
        },
        "total": {
            "bg_color": "#e16463",
            "bold": "true"
        },
        "line": {
            "color": "#e16463"
        }
    },
    {
        "header": {
            "bg_color": "#92c57c",
            "bold": "true"
        },
        "data": {
            "bg_color": "#daf2d8"
        },
        "alt": {
            "bg_color": "#acbfaa"
        },
        "total": {
            "bg_color": "#92c57c",
            "bold": "true"
        },
        "line": {
            "color": "#92c57c"
        }
    },
    {
        "header": {
            "bg_color": "#f3b369",
            "bold": "true"
        },
        "data": {
            "bg_color": "#fae5cd"
        },
        "alt": {
            "bg_color": "#c6b6a3"
        },
        "total": {
            "bg_color": "#f3b369",
            "bold": "true"
        },
        "line": {
            "color": "#f3b369"
        }
    },
    {
        "header": {
            "bg_color": "#6f9eeb",
            "bold": "true"
        },
        "data": {
            "bg_color": "#c9dbf9"
        },
        "alt": {
            "bg_color": "#9faec5"
        },
        "total": {
            "bg_color": "#6f9eeb",
            "bold": "true"
        },
        "line": {
            "color": "#6f9eeb"
        }
    },
    {
        "header": {
            "bg_color": "#fad866",
            "bold": "true"
        },
        "data": {
            "bg_color": "#fcf2cd"
        },
        "alt": {
            "bg_color": "#c9c1a3"
        },
        "total": {
            "bg_color": "#fad866",
            "bold": "true"
        },
        "line": {
            "color": "#fad866"
        }
    },
    {
        "header": {
            "bg_color": "#8e7cc2",
            "bold": "true"
        },
        "data": {
            "bg_color": "#dbd2ea"
        },
        "alt": {
            "bg_color": "#aba4b6"
        },
        "total": {
            "bg_color": "#8e7cc2",
            "bold": "true"
        },
        "line": {
            "color": "#8e7cc2"
        }
    }
]


DEFAULT_OVERALL_STYLES = {
    "income": {
        "header": {
            "bg_color": "#92c57c",
            "bold": "true"
        },
        "data": {
            "bg_color": "#daf2d8"
        },
        "alt": {
            "bg_color": "#acbfaa"
        },
        "total": {
            "bg_color": "#92c57c",
            "bold": "true"
        },
        "line": {
            "color": "#92c57c"
        }
    },
    "expenses": {
        "header": {
            "bg_color": "#e16463",
            "bold": "true"
        },
        "data": {
            "bg_color": "#f4cdcc"
        },
        "alt": {
            "bg_color": "#d3b1af"
        },
        "total": {
            "bg_color": "#e16463",
            "bold": "true"
        },
        "line": {
            "color": "#e16463"
        }
    },
    "surplus": {
        "header": {
            "bg_color": "#6f9eeb",
            "bold": "true"
        },
        "total": {
            "bg_color": "#6f9eeb",
            "bold": "true"
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


def merge_styles_with_defaults(minor_categories: list[str], custom_styles: Styles) -> Styles:
    """Merge styles in `custom_styles` with styles in `DEFAULT_STYLES_LIST` using `categories` as keys."""
    styles: Styles = {}
    num_defaults_used = 0
    default_styles_list_length = len(DEFAULT_STYLES_LIST)
    for category in minor_categories:
        if category in custom_styles.keys():
            styles[category] = custom_styles[category]
        else:
            styles[category] = DEFAULT_STYLES_LIST[num_defaults_used % default_styles_list_length]
            num_defaults_used += 1
    return styles


def create_styles_map_for_overall_data(categories: Dict[str, list[str]]) -> Styles:
    """Create styles map using `DEFAULT_OVERALL_STYLES` for each minor category in `categories`."""
    styles: Styles = {}
    for major_category in categories.keys():
        for minor_category in categories[major_category]:
            styles[minor_category] = DEFAULT_OVERALL_STYLES.get(major_category)
    styles['Total Income'] = DEFAULT_OVERALL_STYLES.get('income')
    styles['Total Expenses'] = DEFAULT_OVERALL_STYLES.get('expenses')
    styles['Total Surplus'] = DEFAULT_OVERALL_STYLES.get('surplus')
    return styles
