from dash import Dash, html, dash_table, dcc, callback, Output, Input
import plotly.express as px
import pandas as pd
import win32com.client as win32
from openpyxl.workbook import Workbook
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import dash_bootstrap_components as dbc

app = Dash(__name__)

# Importando os dados
df = pd.read_excel(r'C:\Users\mateu\Downloads\Apuama\Projeto Trainee\KPI 2.0\Day_15.xlsx')

# Criando a figura do Subplots
fig1 = make_subplots(
    rows=1, cols=3,
    subplot_titles=("Temperature [°C]", "Engine Temperature [°C]", "Oil Temperature [°C]"),
    shared_xaxes=False,
    horizontal_spacing=0.06,
    specs=[[{"type": "scatter"},{"type": "scatter"},{"type": "scatter"}]]
    )

#Add o gráfico a figura
# Logger Temperature
fig1.add_trace(
    go.Scatter(
        x=df["Time (s)"],
        y=df["Logger Temperature [°C]"],
        mode="lines",
        name="Temperature",
    ),
    row=1, col=1
)

##Engine Temperature
fig1.add_trace(
    go.Scatter(
        x=df["Time (s)"],
        y=df["Engine Temp [°C]"],
        mode='lines',
        name="Engine Temperature"
    ),
    row=1, col=2
)

#Oil Temperature
fig1.add_trace(
    go.Scatter(
        x=df["Time (s)"],
        y=df["Oil Temperature [°C]"],
        mode='lines',
        name="Oil Temperature"
    ),
    row=1, col=3
)

# MUdando os nomes dos eixos X e Y
fig1.update_xaxes(title_text="Time (s)", range=[0,300], row=1, col=1)
fig1.update_xaxes(title_text="Time (s)", range=[0,300], row=1, col=2)
fig1.update_xaxes(title_text="Time (s)", range=[0,300], row=1, col=3)

fig1.update_yaxes(title_text="Temperature [°C]", row=1, col=1)
fig1.update_yaxes(title_text="Engine Temperature [°C]", row=1, col=2)
fig1.update_yaxes(title_text="Oil Temperature [°C]", row=1, col=3)

# Configurações da figura
fig1.update_layout(
    height=400,
    width=1500,
    showlegend=False,
    title_text="Vital Sign",
    template="plotly_dark"
)

# Criando a segunda figura
fig2 = make_subplots(
    rows=1, cols=3,
    subplot_titles=("Fuel Pressure [bar]", "Oil Pressure [bar]", "RPM [rpm]"),
    shared_xaxes=False,
    horizontal_spacing=0.06,
    specs=[[{"type": "scatter"},{"type": "scatter"},{"type": "scatter"}]]
)

#Add gráficos a segunda figura
#Fuel Pressure
fig2.add_trace(
    go.Scatter(
        x=df["Time (s)"],
        y=df["Fuel Pressure [PSI]"],
        mode="lines",
        name="Fuel Pressure",
    ),
    row=1, col=1
)

#Oil Pressure
fig2.add_trace(
    go.Scatter(
        x=df["Time (s)"],
        y=df["Oil Pressure [PSI]"],
        mode="lines",
        name="Oil PRessure",
    ),
    row=1, col=2
)

#RPM[rpm]
fig2.add_trace(
    go.Scatter(
        x=df["Time (s)"],
        y=df["RPM [rpm]"],
        mode="lines",
        name="RPM",
    ),
    row=1, col=3
)

fig2.update_xaxes(title_text="Time (s)", range=[0,300], row=1, col=1)
fig2.update_xaxes(title_text="Time (s)", range=[0,300], row=1, col=2)
fig2.update_xaxes(title_text="Time (s)", range=[0,300], row=1, col=3)

fig2.update_yaxes(title_text="Fuel PRessure [bar]", row=1, col=1)
fig2.update_yaxes(title_text="Oil Pressure [bar]", row=1, col=2)
fig2.update_yaxes(title_text="RPM [rpm]", row=1, col=3)

# Configurações da segunda figura
fig2.update_layout(
    height=400,
    width=1500,
    showlegend=False,
    template="plotly_dark"
)

#Criando a terceira figura
fig3 = make_subplots(
    rows=1, cols=1,
    subplot_titles=("External Voltage [V]]"),
    shared_xaxes=False,
    horizontal_spacing=0.06,
    specs=[[{"type": "scatter"}]]
)

#Add gráficos a terceira figura
#External Voltagae
fig3.add_trace(
    go.Scatter(
        x=df["Time (s)"],
        y=df["External Voltage [V]"],
        mode="lines",
        name="Voltage",
    ),
    row=1, col=1
)

#Alterando os nomes dos eixos X e Y
fig3.update_xaxes(title_text="Time (s)",  range=[0,300], row=1, col=1)

fig3.update_yaxes(title_text="External Voltage [V]", row=1, col=1)

# Configurações da terceira figura
fig3.update_layout(
    height=400,
    width=500,
    showlegend=False,
    template="plotly_dark"
)



app = Dash()

app.layout = html.Div(style={'backgroundColor':'#111111'}, children=[
    html.H1(children='APUAMA RACING ', 
         style={'textAlign': 'center','color':'#ffcd24'}),

    html.Div(className='row', children='''
        AF Data
    ''',
        style={'textAlign': 'center','color':'#ffcd24', 'fontSize': 25}),

    dcc.Graph(
        id='row1',
        figure=fig1
    ),
    dcc.Graph(
        id='row2',
        figure=fig2
    ),
    dcc.Graph(
        id='row3',
        figure=fig3
    )
])


if __name__ == '__main__':
    app.run(debug=True)












