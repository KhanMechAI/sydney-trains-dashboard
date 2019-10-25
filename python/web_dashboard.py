import os

# dash libs
import dash
from dash.dependencies import Input, Output
import dash_core_components as dcc
import dash_html_components as html
import plotly.figure_factory as ff
import plotly.graph_objs as go
import dash_table

server = Flask(__name__)

@server.route('/')
def index():
    return render_template('start_page.html')

app = dash.Dash(
    __name__,
    server=server,
    routes_pathname_prefix='/dash/'
)

app.layout = html.Div("My Dash app")


if __name__ == '__main__':
    app.run_server(debug=True)

