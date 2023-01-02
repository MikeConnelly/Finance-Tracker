import os
from typing import Dict
import plotly.graph_objects as go
from plotly.graph_objects import Figure

IMAGE_DIR = 'images'
show_interactive_figure = False


def create_sankey_plot_for_overall_data(category_overall_data: Dict[str, Dict[str, float]]):
    """Create sankey plot for `category_overall_data`. Get the file path to the generated image."""
    # data
    label = []
    source = []
    target = []
    value = []

    # parse data into sankey lists
    total_income_index = len(list(category_overall_data['income'].keys()))
    for index, minor_category in enumerate(category_overall_data['income'].keys()):
        label.append(minor_category)
        source.append(index)
        target.append(total_income_index)
        value.append(category_overall_data['income'][minor_category])

    label.append('Total Income')

    for index, minor_category in enumerate(category_overall_data['expenses'].keys(), start=total_income_index + 1):
        label.append(minor_category)
        source.append(total_income_index)
        target.append(index)
        value.append(category_overall_data['expenses'][minor_category])

    # data to dict, dict to sankey
    link = dict(source=source, target=target, value=value)
    node = dict(label=label, pad=50, thickness=5)
    data = go.Sankey(link=link, node=node)
    # plot
    fig = go.Figure(data)
    if show_interactive_figure:
        fig.show()
    return write_image_file(fig)


def write_image_file(figure: Figure) -> str:
    """Write image file and get the image path."""
    if not os.path.exists(IMAGE_DIR):
        os.mkdir(IMAGE_DIR)
    path = f'{IMAGE_DIR}/sankey1.png'
    figure.write_image(path)
    return path
