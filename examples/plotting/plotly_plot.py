"""
PyXLL Examples: Plotting with plotly
"""
import plotly.express as px
import pyxll

def plot(xs,
         ys,
         xlabel=None,
         ylabel=None,
         title=None,
         width=None,
         height=None,
         zoom=None,
         allow_html=None,
         allow_svg=None):
    """Plot a line chart using plotly"""

    # Create the figure and plot the line chart
    fig = px.line(x=xs, y=ys, title=title, labels={"x": xlabel, "y": ylabel})

    fig.update_layout(
        margin=dict(l=5, r=10, t=30, b=10),
    )

    # Plot the figure in Excel
    pyxll.plot(fig,
               width=width,
               height=height,
               zoom=zoom,
               allow_html=allow_html,
               allow_svg=allow_svg)