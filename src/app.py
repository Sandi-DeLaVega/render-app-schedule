#Added Export Excel Button
#Changed the Format of the Hourly Transaction Data
#fix the editable cells of Cycle Time per TC in minutes Confirmed
#Editable Cells for Allowance for Manpower and Weights Penalty
#With Model Generation

#from jupyter_dash import JupyterDash
#import subprocess
#import webbrowser
import os
from dash import dcc, html, dash_table, Dash, no_update #ctx, 
from dash.dependencies import Input, Output, State
from dash.exceptions import PreventUpdate
import io
import dash_bootstrap_components as dbc
import base64
#from flask import Flask
#import os
#import re
import pandas as pd
import random
import numpy as np
#from openpyxl import Workbook
#from io import BytesIO
import warnings
warnings.filterwarnings("ignore")

# Initialization
#---------------------------------------------------
# -------------------------------------------------
style_data_conditional = [
    {
        "if": {"state": "active"},
         "backgroundColor": "rgba(255,255,0,0.3)",
        "border": "1px solid blue",
    },
    {
        "if": {"state": "selected"},
        "backgroundColor": "rgba(255,255,0,0.3)",
        "border": "1px solid blue",
    },
]

tabs_styles = {'zIndex': 99, 'display': 'inlineBlock', 
               'border': 'grey', 'border-radius': '4px'}





included_col = ["Unnamed: 0", "Unnamed: 1",
                "7.00-8.00", "8.00-9.00", "9.00-10.00", 
                "10.00-11.00", "11.00-12.00",
                "12.00-13.00", "13.00-14.00", "14.00-15.00", "15.00-16.00",
                "16.00-17.00", "17.00-18.00", "18.00-19.00", 
                "19.00-20.00", "20.00-21.00"]

order_list = ["Date", "Day of Week", 
              "00.00-1.00", "1.00-2.00", "2.00-3.00", "3.00-4.00",
              "4.00-5.00", "5.00-6.00", "6.00-7.00", "7.00-8.00",
              "8.00-9.00", "9.00-10.00", "10.00-11.00", "11.00-12.00",
              "12.00-13.00", "13.00-14.00", "14.00-15.00", "15.00-16.00",
              "16.00-17.00", "17.00-18.00", "18.00-19.00", "19.00-20.00",
              "20.00-21.00", "21.00-22.00", "22.00-23.00", "23.00-24.00"] 
# -------------------------------------------------
cycle_time_dict = {
    "Cycle Time in Mins": [1],
    "+ Cycle Time in Secs": [50],
    "Transition Time in Secs":[30]}

cycle_time_df = pd.DataFrame(cycle_time_dict)


# Initial data for the Store
initial_data = cycle_time_df.to_dict('records')



#weekly_sched_df = pd.DataFrame(weekly_sched)
weekly_sched_df = pd.DataFrame()

penalty_weights = {
    'Penalty': ['weight'],
    'Sunday': [2],
    'Monday': [1],
    'Tuesday': [1],
    'Wednesday': [1],
    'Thursday': [1],
    'Friday': [2],
    'Saturday': [2],  }

penalty_weights_df = pd.DataFrame(penalty_weights)
penalty_weights_series = penalty_weights_df.iloc[:,1:].T[0]

manpower_allowance = {
    'Manpower': ['Allowance'],
    '0700': [1],
    '0800': [2],
    '0900': [2],
    '1000': [2],
    '1100': [2],
    '1200': [2],
    '1300': [2],
    '1400': [2],
    '1500': [2],
    '1600': [2],
    '1700': [2],
    '1800': [2],
    '1900': [2],
    '2000': [2],
    '2100': [1],
    
    }

penalty_weights_df = pd.DataFrame(penalty_weights)

manpower_allowance_df = pd.DataFrame(manpower_allowance)

order_list1 = ["Date", "Day of Week", 
              "8.00-9.00", "9.00-10.00", "10.00-11.00", "11.00-12.00",
              "12.00-13.00", "13.00-14.00", "14.00-15.00", "15.00-16.00",
              "16.00-17.00", "17.00-18.00", "18.00-19.00", "19.00-20.00",
              "20.00-21.00", "21.00-22.00"] 

numeric_col = ["8.00-9.00", 
                        "9.00-10.00", "10.00-11.00", "11.00-12.00",
            "12.00-13.00", "13.00-14.00", "14.00-15.00", "15.00-16.00",
            "16.00-17.00", "17.00-18.00", "18.00-19.00", "19.00-20.00",
            "20.00-21.00", "21.00-22.00"]


            
full_time_matrix_df = pd.read_excel('./Data/Full Time and Part Time Matrix.xlsx', sheet_name="Full Time Matrix")
part_time_matrix_df = pd.read_excel('./Data/Full Time and Part Time Matrix.xlsx', sheet_name="Part Time Matrix")

day_off_matrix = {
   'Dayoff': [ 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday',
              'Friday', 'Saturday'],
   'Sunday': [0, 1, 1, 1, 1, 1, 1],
   'Monday': [1, 0, 1, 1, 1, 1, 1],
   'Tuesday': [1, 1, 0, 1, 1, 1, 1],
   'Wednesday': [1, 1, 1, 0, 1, 1, 1],
   'Thursday': [1, 1, 1, 1, 0, 1, 1],
   'Friday': [1, 1, 1, 1, 1, 0, 1],
   'Saturday': [1, 1, 1, 1, 1, 1, 0]}

day_off_matrix_df = pd.DataFrame(day_off_matrix)
day_off_matrix_df.set_index('Dayoff', inplace =True)
download_component = dcc.Download(id="download_component")
download_component_daily = dcc.Download(id = "download_component_daily")
download_component_cdata = dcc.Download(id = "download_component_cdata")
# -------------------------------------------------

style_data_conditional = [
    {
        "if": {"state": "active"},
         "backgroundColor": "rgba(179, 196, 53, 0.5)",
        "border": "1px solid blue",
    },
    {
        "if": {"state": "selected"},
        "backgroundColor": "rgba(179, 196, 53, 0.5)",
        "border": "1px solid blue",
    },
]

style_cell_option_script = {'text-align': 'center', 
                                 'border': '1px solid blue', 
                                 'font-size': '12px'}

style_header_option_script = {'backgroundColor': '#ade2f0', 
                               'font-weight': 'bold', 
                               'text-align': 'center', 
                               'border': '1px solid blue', 
                               'font-size': '12px',
                               'whiteSpace': 'normal'}

# -------------------------------------------------
Keywords, score, playbook_tagging, decision_set = [], [], [], []
#FoundLibrary = refresh_found_library(Keywords, score, playbook_tagging, decision_set)

# -------------------------------------------------

#Creates a path to app's primary css specifications
assets_path = os.getcwd() +'\\assets'
dbc_css = os.path.join("assets_path", "scheduler_v1_3.css")

#server = Flask(__name__)
#changed by uncommenting server = Flask
app = Dash(__name__,  external_stylesheets=[dbc.themes.BOOTSTRAP, dbc_css])
server = app.server
#app = JupyterDash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc_css])

app.title = 'Scheduler App'



hourly_sched_df_sun,  hourly_sched_df_mon,\
 hourly_sched_df_tue, hourly_sched_df_wed,\
 hourly_sched_df_thu, hourly_sched_df_fri, \
     hourly_sched_df_sat = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), \
                             pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
     
hourly_sched_df_sun_summary,  hourly_sched_df_mon_summary,\
 hourly_sched_df_tue_summary, hourly_sched_df_wed_summary,\
 hourly_sched_df_thu_summary, hourly_sched_df_fri_summary, \
     hourly_sched_df_sat_summary = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), \
                                 pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

cashiers_reporting, headcount_per_hour1, headcount_per_hour2 = pd.DataFrame(), None, None
     
dataframes = {
    'Sunday': hourly_sched_df_sun,
    'Monday': hourly_sched_df_mon,
    'Tuesday': hourly_sched_df_tue,
    'Wednesday': hourly_sched_df_wed,
    'Thursday': hourly_sched_df_thu,
    'Friday': hourly_sched_df_fri,
    'Saturday': hourly_sched_df_sat,
    'Weekly Schedule': weekly_sched_df,
}

dataframes_summary = {
    'Sunday': hourly_sched_df_sun_summary,
    'Monday': hourly_sched_df_mon_summary,
    'Tuesday': hourly_sched_df_tue_summary,
    'Wednesday': hourly_sched_df_wed_summary,
    'Thursday': hourly_sched_df_thu_summary,
    'Friday': hourly_sched_df_fri_summary,
    'Saturday': hourly_sched_df_sat_summary,
    'Weekly Schedule': cashiers_reporting,
    'HC Per Hour': headcount_per_hour1,
    'HC Per Step': headcount_per_hour2,
}

#------------------------------------------   

app.layout = html.Div(
    [
#------------------------------------------   
        #Title Portion
        html.Div(
            className = "grid-container-header",
            children = [
                html.Div(
                    className="item1",
                    children=[
                        html.Div('Scheduler App Version 5.4.3.1'), 
                        ], ), 
                html.Div(
                    className="grid-item2",
                    children=[], ), 

                 ]
             ),
        
        dbc.Container([
            #Tabs for Data and Output
            dcc.Tabs(className = "dbc", id = "main_tab",
                                     style = {"backgroundColor": "#FFCD00",
                                                'color': '#173759',
                                                'font-family': "Montserrat",
                                                'font-size': '20px',
                                              'font-weight': 'bold',
                                                    },
                     
                     children = [
                    
                #First Tab
                
                dbc.Tab(label = "1. Input/Upload Data", 
                        id = "input-tab-main",
                        style = {
                                "backgroundColor": "#5791CE",
                                'color': 'white',
                                'font-family': "Montserrat",
                                'font-size': '20px',
                                    },

                        
                        
                        #style = tab_style_main, 
                        #selected_style = tab_selected_style_main_1st,
                        children = [
                            
                        html.Div(className="grid-item-tab-main",
                        children = [
                        
                        dbc.Row([
                            #Input Hourly Sales Penetration
                            dbc.Col([
                                    html.Div( 
                                        className="grid-item-main-data",
                                     #className = "grid-item-kbarticle", 
                                     children = [
                                        html.Div( className="header-article", 
                                                     children=[ html.Div('1. Cycle Time Per TC in Minutes')]), 

                                         #html.Div( className="header-subtitle", 
                                                            #children=[ html.Div('Change the Parameters for the Schedule Generation')]), 
                                         
                                          
                                         dbc.Row([
                                             dbc.Col([
                                                 html.Br(),
                                                 #html.Div( className="header-article", 
                                                  #children=[ html.Div('Cycle Time Per TC in Minutes')]), 
                                                 
                                                 html.Div(id = 'textarea-output-table-cycle-time', 
                                                          style = {'whiteSpace': 'pre-line', 
                                                                                    'border-style': 'inset', 
                                                                                    'width': 'auto' }, 
                                                          
                                                       children = [dash_table.DataTable(id = 'cycle_time_table', 
                                                                            columns = [{'name': i, 'id': i} \
                                                                                       for i in cycle_time_df.columns], 
                                                                            data = cycle_time_df.to_dict('records'), 
                                                                            editable=True,
                                                                            style_data = {
                                                                                         'whiteSpace': 'normal',
                                                                                         'height': 'auto',
                                                                                         },
                                                                            style_table={'overflowY': 'auto'},
                                                                            #page_size = 10,
                                                                             style_data_conditional = style_data_conditional,
                                                                            style_cell = style_cell_option_script, 
                                                                            style_header = style_header_option_script), 
                                                                  ],
                                                           ),
                                                 ], width = 9),
                                                 dbc.Col([
                                                     html.Br(),
                                                     html.Div(className="header-article", 
                                                      children=[ html.Div('TC per HC')]), 
                                                     
                                                     html.Div(id = "transaction-output", className="header-article", 
                                                      children=[]), 
                                                     
                                                     ], width = 3)
                                                 
                                             ]),
                                             
                                        dbc.Row([
                                            html.Br(),
                                            ]),
                                        
                                        dbc.Row([
                                            dbc.Col([
                                                html.Div(className = 'header-subtitle', 
                                                         children=[ \
                                                    html.Div('Enter Time Above Between 0 to 60 then press Enter to Save'),
                                                    html.Div('For Cycle Time in Mins, min value is 1.')]),
                                                ])
                                            
                                            ]),
                                        
                                        dbc.Row([
                                            html.Br(),
                                            html.Br(),
                                            ]),
                                         
                                         
                                             dbc.Row([
                                                 dbc.Col([
                                                     dbc.Row([
                                                         dbc.Col([
                                                             ], width = 2),
                                                         dbc.Col([
                                                             html.Div(
                                                                 className = "button_gen_sched_div",
                                                                 children = [html.Button("Generate Schedule", id="generate--model-button", 
                                                                                         className = "button_gen_sched",
                                                                                         disabled=True),
                                                                             
                                                                             ]),
                                                             
                                                             ], width = 8),
                                                         dbc.Col([
                                                             ], width = 2),
                                                         ])    
                                                     ])
                                                 ]),
                                             
                                             dbc.Row([
                                                 html.Br(),
                                                 ]),
                                              
                                             dbc.Row([
                                                 
                                                 dbc.Col([
                                                     html.Div(className = 'header-subtitle', 
                                                              id='status-message', children = [],style={'margin-top': '10px'}),
                                                     
                                                     html.Div(className = 'header-subtitle', 
                                                              id='status-message-model-generation', children = [],style={'margin-top': '10px'}),
                                                     ])
                                                 ])
                                         
                                                  ],
                                                 ), #end of sub Tabs of Input Sale Penetration
                                    
                                        
                                        
                                        
                                        ], 
                                width=4), #end of Input Hourly Sales Penetration column 
                                
                            #Input Personnel
                            dbc.Col([
                                html.Div(
                                    className="grid-item-main-data",
                                         children = [
                                             html.Div( className="header-article", 
                                                                  children=[ html.Div('2. Personnel Data')]),

                                             html.Div( className="header-subtitle", 
                                                            children=[ html.Div('Upload the Personnel Data Below')]), 
                                             
                                            html.Div(className = "grid-item-tab-main-data",
                                            children = [
                                               
                                            
                                            dbc.Row(children = [
                                            dbc.Col([
                                                     html.Div(children = [dcc.Upload(
                                                              id = 'upload-data-personnel',
                                                              children = html.Div([
                                                                  'Drag and Drop or ',
                                                                  html.A('Select File')]),
                                                            style={
                                                                'color': "#173759",
                                                            'width': '95%',
                                                            'height': '60px',
                                                            'lineHeight': '60px',
                                                            'borderWidth': '1px',
                                                            'borderStyle': 'dashed',
                                                            'borderRadius': '5px',
                                                                'font-weight': 'bold',
                                                            'textAlign': 'center',
                                                            'margin': '5px'}, multiple = False
                                                                                    ),
                                                                         ],
                                                             ),
                                                    ]),
                                                     ],
                                                    ),
                                            
                                            dbc.Row(
                                            children = [
                                            dbc.Col([
                                            html.Div(className = "grid-item-output",
                                                     children = [
                                                         html.Div(id='output-div-data-personnel'),
                                                         html.Div(id='output-datatable-personnel', 
                                                                  children = []),
                                                     ]),
                                            ]),
                                            
                                            ],
                                                
                                            ),
                                            #status of Data Upload
                                            dbc.Row(
                                                children = [
                                                    dbc.Col([
                                                        html.Div( className="header-article", 
                                                         children=[ html.Div('Status of Data Upload')]), 
                                                        html.Div(id = "output-upload-personnel-data" , 
                                                                 className="header-article-upload", 
                                                         children=[ ]), 
                                                    ]),
                                                ]),  
                                            ]), #End of Main Data
                                                     ],
                                         ), 
                                    ], width=4), #end of input Personnel column

                            #Input Daily Sales
                            dbc.Col([
                                    html.Div( 
                                        className="grid-item-main-data",
                                        #className = "grid-item-kbarticle", 
                                        children = [
                                            html.Div( className="header-article", 
                                                            children=[ html.Div('3. Hourly Transaction Data')]), 

                                            html.Div( className="header-subtitle", 
                                                            children=[ html.Div('Choose Between Upload or Manual')]), 
                                            
                                            html.Div(className = "grid-item-tab-main-data",
                                            children = [
                                               
                                            
                                            dbc.Row(children = [
                                            dbc.Col([
                                                     html.Div(children = [dcc.Upload(
                                                              id = 'upload-data-daily-sales',
                                                              children = html.Div([
                                                                  'Drag and Drop or ',
                                                                  html.A('Select File')]),
                                                            style={
                                                                'color': "#173759",
                                                            'width': '95%',
                                                            'height': '60px',
                                                            'lineHeight': '60px',
                                                            'borderWidth': '1px',
                                                            'borderStyle': 'dashed',
                                                            'borderRadius': '5px',
                                                                'font-weight': 'bold',
                                                            'textAlign': 'center',
                                                            'margin': '5px'}, multiple = False
                                                                                    ),
                                                                         ],
                                                             ),
                                                    ]),
                                                     ],
                                                    ),
                                            
                                            dbc.Row(
                                            children = [
                                            dbc.Col([
                                            html.Div(className = "grid-item-output",
                                                     children = [
                                                         html.Div(id='output-div-data-daily-sales'),
                                                         html.Div(id='output-datatable-daily-sales'),
                                                         html.Br(),
                                                         html.Div(id='output-datatable-headcount_per_hour1'),
                                                         html.Br(),
                                                         html.Div(id='output-datatable-headcount_per_hour2'),
                                                     ]),
                                            ]),
                                            
                                            ],
                                                
                                            ),
                                            #status of Data Upload
                                            dbc.Row(
                                                children = [
                                                    dbc.Col([
                                                        html.Div( className="header-article", 
                                                         children=[ html.Div('Status of Data Upload')]), 
                                                        
                                                        html.Div(id = "output-upload-data-daily-sales" , 
                                                                 className="header-article-upload", 
                                                         children=[ ]), 
                                                        
                                                    ]),
                                                ]),  
                                            ]), #End of Main Data
                                                                    
                                                   
                                            ]), #end of grid-item-tagging 
                                    ], width=4), #end of Input Daily Sales column 

                           
                        ]), #end of Row1
                    
                        
                        ], #end of children of Html.Div
                                ), #End of HTML.DIV
                            
                        
                    
                ]),#end of Tab 1
                
                
                dbc.Tab(label = "2. Weekly and Daily Schedule Generation", 
                        id = "sched-tab-main",
                        style = {
                                "backgroundColor": "#ED213B",
                                'color': 'white',
                                'font-family': "Montserrat",
                                'font-size': '20px',
                                    },

                        #style = tab_style_main, 
                        #selected_style = tab_selected_style_main_2nd,
                        children = [
                        html.Div(className="grid-item-tab-main",
                                children = [
                                     #Status of the Model Generation
                                    dbc.Row(),
                                    #Weekly Schedule Generation
                                    dbc.Row(
                                        children = [
                                            dbc.Col(
                                                
                                                [
                                                    html.Div(
                                                        className="grid-item-main-weekly-sched",
                                                    children = [html.Div(className="header-article",
                                                                         children = ["Weekly Schedule Generated"]),
                                                                
                                                                html.Div(className="grid-weekly-sched-data",
                                                                         children = [
                                                                             
                                                                             #Cashier Reporting and the Weekly Schedule
                                                                             html.Div(id = 'textarea-output-table-weekly-sched', 
                                                                                      className = "output-grid-daily-sched",
                                                                                      style = {'whiteSpace': 'pre-line', 
                                                                                                                'border-style': 'inset', 
                                                                                                                'width': '100%',
                                                                                                                'overflow': 'scroll'}, 
                                                                                      
                                                                                   children = [dash_table.DataTable(id = 'weekly_sched_table', 
                                                                                                        columns = [{'name': i, 'id': i} \
                                                                                                                   for i in weekly_sched_df.columns], 
                                                                                                        data = weekly_sched_df.to_dict('records'), 
                                                                                                        style_data = {
                                                                                                                     'whiteSpace': 'normal',
                                                                                                                     'height': 'auto',
                                                                                                                     },
                                                                                                        style_table={'overflowY': 'auto'},
                                                                                                        page_size = 6,
                                                                                                         style_data_conditional = style_data_conditional,
                                                                                                        style_cell = style_cell_option_script, 
                                                                                                        style_header = style_header_option_script), 
                                                                                              ],
                                                                                       ),
                                                                               
                                                                            
                                                                            
                                                                             ]),
                                                                
                                                                    
                                                               ]),
                                                
                                                ])
                                            ],
                                            ),
                                    #Daily Schedule Generation
                                    dbc.Row(
                                            children = [
                                                dbc.Col([html.Div(
                                                className="grid-item-main-hourly-sched",
                                                children = [html.Div(className="header-article",
                                                                         children = ["Daily Schedule Generated"]),
                                                dcc.Tabs(className = "dbc_sub4",
                                                         style = {
                                                            'font-family': "Montserrat",
                                                            'font-size': '16px',
                                                            'border': '2px solid white',
                                                             'border-radius': '4px',
                                                             'font-weight': 'bold',
                                                        
                                                                },
                                                     children = [
                                                                #Sunday Tab
                                                                dbc.Tab(label = "Sun",
                                                                        style = {
                                                                                "backgroundColor": "#FECE10",
                                                                                'color': '#ED213B',
                                                                                'font-family': "Montserrat",
                                                                                'font-size': '16px',
                                                                                'font-weight': 'bold',
                                                                            'border': '2px solid white',
                                                                            
                                                                                    },
                                                                     children = [
                                                                       html.Div(className = "grid-item-tab-sub",
                                                                               children = [
                                                                                   html.Div(id = 'textarea-output-table-daily-sched-sun', 
                                                                                            style = {'whiteSpace': 'pre-line', 
                                                                                                             'border-style': 'inset', 
                                                                                                             'width': '100%',
                                                                                                             'overflow': 'scroll'}, 
                                                                                children = [dash_table.DataTable(id = 'hourly_sched_df_sun_table', 
                                                                                                     columns = [{'name': i, 'id': i} \
                                                                                                                for i in hourly_sched_df_sun.columns], 
                                                                                                     data = hourly_sched_df_sun.to_dict('records'), 
                                                                                                     style_data={
                                                                                                                  'whiteSpace': 'normal',
                                                                                                                  'height': 'auto',
                                                                                                                  },
                                                                                                     style_table={'overflowY': 'auto'},
                                                                                                     page_size = 10,
                                                                                                      style_data_conditional = style_data_conditional,
                                                                                                     style_cell = style_cell_option_script, 
                                                                                                     style_header = style_header_option_script), 
                                                                                            ],
                                                                                        ),
                                                                                   
                                                                                   
                                                                                   ])],),
                                                                #Monday Tab
                                                                dbc.Tab(label = "Mon",
                                                                        style = {
                                                                                "backgroundColor": "#FECE10",
                                                                                'color': '#ED213B',
                                                                                'font-family': "Montserrat",
                                                                                'font-size': '16px',
                                                                                'font-weight': 'bold',
                                                                             'border': '2px solid white',
                                                                                    },
                                                                     children = [
                                                                       html.Div(className = "grid-item-tab-sub",
                                                                               children = [
                                                                                       html.Div(id = 'textarea-output-table-daily-sched-mon', 
                                                                                                style = {'whiteSpace': 'pre-line', 
                                                                                                                 'border-style': 'inset', 
                                                                                                                 'width': '100%',
                                                                                                                 'overflow': 'scroll'}, 
                                                                                                children = [dash_table.DataTable(id = 'hourly_sched_df_mon_table', 
                                                                                                                     columns = [{'name': i, 'id': i} \
                                                                                                                                for i in hourly_sched_df_mon.columns], 
                                                                                                                     data = hourly_sched_df_mon.to_dict('records'), 
                                                                                                                     style_data={
                                                                                                                                  'whiteSpace': 'normal',
                                                                                                                                  'height': 'auto',
                                                                                                                                  },
                                                                                                                     style_table={'overflowY': 'auto'},
                                                                                                                     page_size = 10,
                                                                                                                      style_data_conditional = style_data_conditional,
                                                                                                                     style_cell = style_cell_option_script, 
                                                                                                                     style_header = style_header_option_script), 
                                                                                                            ],
                                                                                                        ),
                                                                                   
                                                                                   ])],),
                                                                #Tuesday Tab
                                                                dbc.Tab(label = "Tue",
                                                                        style = {
                                                                                "backgroundColor": "#FECE10",
                                                                                'color': '#ED213B',
                                                                                'font-family': "Montserrat",
                                                                                'font-size': '16px',
                                                                                'font-weight': 'bold',
                                                                             'border': '2px solid white',
                                                                                    },
                                                                     children = [
                                                                       html.Div(className = "grid-item-tab-sub",
                                                                               children = [
                                                                                       html.Div(id = 'textarea-output-table-daily-sched-tue', 
                                                                                                style = {'whiteSpace': 'pre-line', 
                                                                                                                 'border-style': 'inset', 
                                                                                                                 'width': '100%',
                                                                                                                 'overflow': 'scroll'}, 
                                                                                                children = [dash_table.DataTable(id = 'hourly_sched_df_tue_table', 
                                                                                                                     columns = [{'name': i, 'id': i} \
                                                                                                                                for i in hourly_sched_df_tue.columns], 
                                                                                                                     data = hourly_sched_df_tue.to_dict('records'), 
                                                                                                                     style_data={
                                                                                                                                  'whiteSpace': 'normal',
                                                                                                                                  'height': 'auto',
                                                                                                                                  },
                                                                                                                     style_table={'overflowY': 'auto'},
                                                                                                                     page_size = 10,
                                                                                                                      style_data_conditional = style_data_conditional,
                                                                                                                     style_cell = style_cell_option_script, 
                                                                                                                     style_header = style_header_option_script), 
                                                                                                            ],
                                                                                                        ),
                                                                                   
                                                                                   ])],),
                                                                #Wednesday Tab
                                                                 dbc.Tab(label = "Wed",
                                                                         style = {
                                                                                "backgroundColor": "#FECE10",
                                                                                'color': '#ED213B',
                                                                                'font-family': "Montserrat",
                                                                                'font-size': '16px',
                                                                                'font-weight': 'bold',
                                                                              'border': '2px solid white',
                                                                                    },
                                                                     children = [
                                                                       html.Div(className = "grid-item-tab-sub",
                                                                               children = [
                                                                                           html.Div(id = 'textarea-output-table-daily-sched-wed', 
                                                                                                    style = {'whiteSpace': 'pre-line', 
                                                                                                                     'border-style': 'inset', 
                                                                                                                     'width': '100%',
                                                                                                                     'overflow': 'scroll'}, 
                                                                                                    children = [dash_table.DataTable(id = 'hourly_sched_df_wed_table', 
                                                                                                                         columns = [{'name': i, 'id': i} \
                                                                                                                                    for i in hourly_sched_df_wed.columns], 
                                                                                                                         data = hourly_sched_df_wed.to_dict('records'), 
                                                                                                                         style_data={
                                                                                                                                      'whiteSpace': 'normal',
                                                                                                                                      'height': 'auto',
                                                                                                                                      },
                                                                                                                         style_table={'overflowY': 'auto'},
                                                                                                                         page_size = 10,
                                                                                                                          style_data_conditional = style_data_conditional,
                                                                                                                         style_cell = style_cell_option_script, 
                                                                                                                         style_header = style_header_option_script), 
                                                                                                                ],
                                                                                                            ),
                                                                                   
                                                                                   ])],),
                                                                 #Thursday Tab
                                                                dbc.Tab(label = "Thu",
                                                                        style = {
                                                                                "backgroundColor": "#FECE10",
                                                                                'color': '#ED213B',
                                                                                'font-family': "Montserrat",
                                                                                'font-size': '16px',
                                                                                'font-weight': 'bold',
                                                                             'border': '2px solid white',
                                                                                    },
                                                                     children = [
                                                                       html.Div(className = "grid-item-tab-sub",
                                                                               children = [
                                                                                   
                                                                                               html.Div(id = 'textarea-output-table-daily-sched-thu', 
                                                                                                        style = {'whiteSpace': 'pre-line', 
                                                                                                                         'border-style': 'inset', 
                                                                                                                         'width': '100%',
                                                                                                                         'overflow': 'scroll'}, 
                                                                                                        children = [dash_table.DataTable(id = 'hourly_sched_df_thu_table', 
                                                                                                                             columns = [{'name': i, 'id': i} \
                                                                                                                                        for i in hourly_sched_df_thu.columns], 
                                                                                                                             data = hourly_sched_df_thu.to_dict('records'), 
                                                                                                                             style_data={
                                                                                                                                          'whiteSpace': 'normal',
                                                                                                                                          'height': 'auto',
                                                                                                                                          },
                                                                                                                             style_table={'overflowY': 'auto'},
                                                                                                                             page_size = 10,
                                                                                                                              style_data_conditional = style_data_conditional,
                                                                                                                             style_cell = style_cell_option_script, 
                                                                                                                             style_header = style_header_option_script), 
                                                                                                                    ],
                                                                                                                ),
                                                                                                   
                                                                                   
                                                                                           ])],),
                                                                 #Friday Tab
                                                                 dbc.Tab(label = "Fri",
                                                                         style = {
                                                                                "backgroundColor": "#FECE10",
                                                                                'color': '#ED213B',
                                                                                'font-family': "Montserrat",
                                                                                'font-size': '16px',
                                                                                'font-weight': 'bold',
                                                                              'border': '2px solid white',
                                                                                    },
                                                                     children = [
                                                                       html.Div(className = "grid-item-tab-sub",
                                                                               children = [
                                                                                                html.Div(id = 'textarea-output-table-daily-sched-fri', 
                                                                                                         style = {'whiteSpace': 'pre-line', 
                                                                                                                          'border-style': 'inset', 
                                                                                                                          'width': '100%',
                                                                                                                          'overflow': 'scroll'}, 
                                                                                                         children = [dash_table.DataTable(id = 'hourly_sched_df_fri_table', 
                                                                                                                              columns = [{'name': i, 'id': i} \
                                                                                                                                         for i in hourly_sched_df_fri.columns], 
                                                                                                                              data = hourly_sched_df_fri.to_dict('records'), 
                                                                                                                              style_data={
                                                                                                                                           'whiteSpace': 'normal',
                                                                                                                                           'height': 'auto',
                                                                                                                                           },
                                                                                                                              style_table={'overflowY': 'auto'},
                                                                                                                              page_size = 10,
                                                                                                                               style_data_conditional = style_data_conditional,
                                                                                                                              style_cell = style_cell_option_script, 
                                                                                                                              style_header = style_header_option_script), 
                                                                                                                     ],
                                                                                                                 ),                                                                                   
                                                                                   
                                                                                   ])],),
                                                         
                                                                  #Saturday Tab
                                                                 dbc.Tab(label = "Sat",
                                                                         style = {
                                                                                "backgroundColor": "#FECE10",
                                                                                'color': '#ED213B',
                                                                                'font-family': "Montserrat",
                                                                                'font-size': '16px',
                                                                                'font-weight': 'bold',
                                                                              'border': '2px solid white',
                                                                                    },
                                                                     children = [
                                                                       html.Div(className = "grid-item-tab-sub",
                                                                               children = [
                                                                                               html.Div(id = 'textarea-output-table-daily-sched-sat', 
                                                                                                        style = {'whiteSpace': 'pre-line', 
                                                                                                                         'border-style': 'inset', 
                                                                                                                         'width': '100%',
                                                                                                                         'overflow': 'scroll'}, 
                                                                                                        children = [dash_table.DataTable(id = 'hourly_sched_df_sat_table', 
                                                                                                                             columns = [{'name': i, 'id': i} \
                                                                                                                                        for i in hourly_sched_df_sat.columns], 
                                                                                                                             data = hourly_sched_df_sat.to_dict('records'), 
                                                                                                                             style_data={
                                                                                                                                          'whiteSpace': 'normal',
                                                                                                                                          'height': 'auto',
                                                                                                                                          },
                                                                                                                             style_table={'overflowY': 'auto'},
                                                                                                                             page_size = 10,
                                                                                                                              style_data_conditional = style_data_conditional,
                                                                                                                             style_cell = style_cell_option_script, 
                                                                                                                             style_header = style_header_option_script), 
                                                                                                                    ],
                                                                                                                ),   
                                                                                   
                                                                                           ])],),
                                                         
                                                                 ],
                                                    ),
                                                ]
                                                
                                                )]),
                                            
                                            ],

                                            ), #end of dbc.Row of both daily and weekly sched
                                        
                                dbc.Row([
                                    dbc.Col([
                                        dbc.Row([
                                            dbc.Col([
                                                html.Div(
                                                    className = "button_download_div",
                                                    children = [html.Button("Prepare to Export", id="export-button", 
                                                                            className = "button_download"),]),
                                                download_component,
                                                #download_component_daily,
                                                #download_component_cdata,
                                                ], width = 2),
                                            
                                            dbc.Col([
                                                
                                                html.Div(id = 'export-button-status-text',
                                                         children = []),
                                                html.A("Download Hourly.csv", 
                                                       id="download-link", 
                                                       download="hourly.csv", href="",
                                                       style={'display': 'none'},
                                                       target="_blank"),
                                                ], width = 10),
                                            
                                            
                                            
                                            ])    
                                    
                                        
                                           
                                        
                                        
                                        ])
                                    
                                    
                                   
                                    ]),

                        
                                ]), #end of grid-item-tab-main
                       ]), #end of Tab 2
            ]), #end of dcc Tabs
                
            
        ], 
        fluid=True),  #end of children of dbc-container               
    #Footer
    html.Div(
            className = "grid-container-footer",
            children = [
                html.Div(
                    className="item2",
                    children=[
                        html.Div('Updated as of Dec 3, 2023'), 
                        ], ), 
                html.Div(
                    className="grid-item2",
                    children=[], ), 

                 ]
             ),
                #to store data for sharing between select_from_table and search
                dcc.Store(id='headcount_per_hour1', data=[], storage_type='memory'),
                dcc.Store(id='headcount_per_hour2', data=[], storage_type='memory'),
                dcc.Store(id='personnel-data', data=[], storage_type='memory'),
                dcc.Store(id = "cycle-time-data", data = [], storage_type = "memory"),
                dcc.Store(id = "cashier-reporting-data", data = [], storage_type = "memory"),
                dcc.Store(id = "weekly-sched-df-data", data = [], storage_type = "memory"),
                dcc.Store(id = "weekly_sched_gen-data", data = [], storage_type = "memory"),
                dcc.Store(id = "all-data", data = [], storage_type = "memory"),
                
             ] #end of main layout children
         ) #end of main layout
    
    
    
    
    

@app.callback(
    [Output('generate--model-button', 'disabled'),
        Output('generate--model-button', 'style'),
        Output('status-message', 'children')],
    [Input('upload-data-personnel', 'contents'),
     Input('upload-data-daily-sales', 'contents')]
)
def update_button_status(file1_contents, file2_contents):
    if file1_contents is not None and file2_contents is not None:
        # Both files are uploaded successfully
        return False,  {'background-color': '#3aa803'}, html.Div('Both files uploaded successfully')
    
    elif file1_contents is not None:
        # File 1 is uploaded, prompt to upload File 2
        return True, {'background-color': '#ddd'}, html.Div('Personnel Data has been uploaded. Please upload Sales Data.')
    elif file2_contents is not None:
        # File 2 is uploaded, prompt to upload File 1
        return True, {'background-color': '#ddd'}, html.Div('Sales Data has been uploaded. Please upload Personnel Data.')
    
    else:
        # At least one file is missing
        return True, {'background-color': '#ddd'}, html.Div('Please upload Personnel Data and Sales Data.')
    

#Enable Export Button if Model has been generated
@app.callback(
    [Output('export-button', 'disabled'),
     Output('export-button', 'style')],
    [Input('generate--model-button', 'n_clicks')],
    [State('generate--model-button', 'disabled')]
)
def update_export_button(generate_clicks, generate_disabled):
    if generate_disabled is True:
        # The generate model button is disabled or hasn't been clicked yet
        return True, {'background-color': '#ddd'}
    else:
        
        # The generate model button has been clicked at least once
        return False, {'background-color': '#3aa803'}



#Model Generation Logic
@app.callback(
    [Output('textarea-output-table-weekly-sched', 'children'),
     Output('textarea-output-table-daily-sched-sun', 'children'),
     Output('textarea-output-table-daily-sched-mon', 'children'),
     Output('textarea-output-table-daily-sched-tue', 'children'),
     Output('textarea-output-table-daily-sched-wed', 'children'),
     Output('textarea-output-table-daily-sched-thu', 'children'),
     Output('textarea-output-table-daily-sched-fri', 'children'),
     Output('textarea-output-table-daily-sched-sat', 'children'),
     
     Output('cashier-reporting-data', 'data'),
     Output('weekly-sched-df-data', 'data'),
     Output('weekly_sched_gen-data', 'data'),
     Output('all-data', 'data'),
     Output('status-message-model-generation', 'children')],
    
    [Input('generate--model-button', 'n_clicks')],
    [State('personnel-data', 'data'),
     State('headcount_per_hour1', 'data'),
     State('headcount_per_hour2', 'data')]
)
def generate_table(n_clicks, pdata, data1, data2):
    if n_clicks is None or n_clicks == 0:
        # The button has not been clicked
        return None, None,None,None,None,None,None,None,[], [], [], [], ""
    
    else:
        
        headcount_per_hour1 = pd.DataFrame(data1)
        headcount_per_hour2 = pd.DataFrame(data2)
        personnel_data_df = pd.DataFrame(pdata)
        
        weekly_sched_gen = headcount_per_hour2[["Day of Week", "New HC Needed"]].T
        weekly_sched_gen.columns = weekly_sched_gen.iloc[0]
        weekly_sched_gen = weekly_sched_gen.iloc[1:,].T

        personnel_df_sched = personnel_data_df[["Personnel Name"]]
        personnel_df_sched["Person #"] = ["Person " + str(i) for i in range(1,len(personnel_df_sched) + 1)]

        # --------------------- Daily
        day_off_counter = {
            'Sunday': 0,
            'Monday': 0,
            'Tuesday': 0,
            'Wednesday': 0,
            'Thursday': 0,
            'Friday': 0,
            'Saturday': 0,  }

        day_off_counter_series = pd.Series(day_off_counter)
        day_of_list = []
        #start_time = time.time()
        for i in range(1,len(personnel_df_sched) + 1):
            iteration_key = "Iteration " + str(i)
            if iteration_key == "Iteration 1":
                iteration_key_prev = "New HC Needed"
            else:
                iteration_key_prev = "Iteration " + str(i - 1)
            
            min_hc =  weekly_sched_gen[iteration_key_prev].min()
            #What are the Days with min HC Needed
            day_with_min_HC = weekly_sched_gen[weekly_sched_gen[iteration_key_prev] == min_hc].index.tolist()

            if len(day_with_min_HC) == 1:
                choose_day_off = day_with_min_HC[0]
            else:
                
                # check which one has lower count in day_off_counter
                counts = day_off_counter_series[day_with_min_HC]

                # Find all days with the minimum count
                min_count = counts.min()
                days_with_min_count = counts[counts == min_count]
                
                if len(days_with_min_count) == 1:
                    choose_day_off = days_with_min_count.index[0]
                    
                else:
                    
                    # Get the penalty weights for the days with tied counts
                    penalty_weights_tied_days = penalty_weights_series[days_with_min_count.index]
                    
                    # Sort by penalty weights to choose the day with the highest weight
                    penalty_weights_tied_days = penalty_weights_tied_days.sort_values(ascending=False)
                    
                    # If there are multiple days with the lowest  weight, randomly choose one
                    min_penalty_days = penalty_weights_tied_days[penalty_weights_tied_days == penalty_weights_tied_days.min()]
                    
                    if len(min_penalty_days) == 1:
                        choose_day_off = min_penalty_days.index[-1]
                    else:
                        days_list = min_penalty_days.index.tolist()
                        choose_day_off = random.choice(days_list)
                   
            day_of_list.append(choose_day_off)
            day_off_counter_series[choose_day_off] += 1
            weekly_sched_gen[iteration_key] = weekly_sched_gen[iteration_key_prev] - day_off_matrix_df[choose_day_off]
            
        #end_time = time.time()    
        #total_time = end_time - start_time
        #print(f"Total Time Generated for Daily {total_time}")
        personnel_data_df["Day Off"] = day_of_list

        weekly_sched_df = personnel_data_df.merge(day_off_matrix_df,
                                                   left_on = "Day Off", 
                                                   right_on = day_off_matrix_df.index,
                                                   how = 'left')
        #Output
        cashiers_reporting = weekly_sched_df.groupby("Employment Type").sum()
        cashiers_reporting = cashiers_reporting.reset_index(drop = False)
        cashiers_reporting = cashiers_reporting[["Employment Type", "Sunday",
                                                 "Monday", "Tuesday", "Wednesday",
                                                 "Thursday", "Friday", "Saturday"]]
        #Removed Cashier Reporting Display
        weekly_sched_df_dash =  html.Div([\
                                html.H5('Schedule'),
                                
                                dash_table.DataTable(id = 'weekly_sched_table', 
                                columns = [{'name': i, 'id': i} \
                                    for i in weekly_sched_df.columns], 
                                    data = weekly_sched_df.to_dict('records'), 
                                   style_data = {
                                        'whiteSpace': 'normal',
                                       'height': 'auto',
                                                              },
                                                 style_table={'overflowY': 'auto'},
                                                 page_size = 8,
                                                  style_data_conditional = style_data_conditional,
                                                 style_cell = style_cell_option_script, 
                                                 style_header = style_header_option_script), 
                                       ],
                                ),
            
        store_cashiers_reporting = cashiers_reporting.reset_index(drop = True)
        store_cashiers_reporting = store_cashiers_reporting.to_dict('records')
        
        store_weekly_sched_df = weekly_sched_df.reset_index(drop = True)
        store_weekly_sched_df = store_weekly_sched_df.to_dict('records')
        
        store_weekly_sched_gen = weekly_sched_gen.reset_index()
        store_weekly_sched_gen = store_weekly_sched_gen.to_dict('records')
        
        #--------------------------Hourly Generation
        hourly_sched_df_sun,  hourly_sched_df_mon,\
         hourly_sched_df_tue, hourly_sched_df_wed,\
         hourly_sched_df_thu, hourly_sched_df_fri, \
             hourly_sched_df_sat = None, None, None, None, None, None, None
             
        hourly_sched_df_sun_summary,  hourly_sched_df_mon_summary,\
         hourly_sched_df_tue_summary, hourly_sched_df_wed_summary,\
         hourly_sched_df_thu_summary, hourly_sched_df_fri_summary, \
             hourly_sched_df_sat_summary = None, None, None, None, None, None, None

        main_hourly_df_basis_sun,  main_hourly_df_basis_mon, \
        main_hourly_df_basis_tue, main_hourly_df_basis_wed, \
         main_hourly_df_basis_thur,  main_hourly_df_basis_fri, \
        main_hourly_df_basis_sat = None, None, None, None, None, None, None
             
        dataframes = {
            'Sunday': hourly_sched_df_sun,
            'Monday': hourly_sched_df_mon,
            'Tuesday': hourly_sched_df_tue,
            'Wednesday': hourly_sched_df_wed,
            'Thursday': hourly_sched_df_thu,
            'Friday': hourly_sched_df_fri,
            'Saturday': hourly_sched_df_sat,
            #"Weekly Iteration": weekly_sched_gen
        }

        dataframes_summary = {
            'Sunday': hourly_sched_df_sun_summary,
            'Monday': hourly_sched_df_mon_summary,
            'Tuesday': hourly_sched_df_tue_summary,
            'Wednesday': hourly_sched_df_wed_summary,
            'Thursday': hourly_sched_df_thu_summary,
            'Friday': hourly_sched_df_fri_summary,
            'Saturday': hourly_sched_df_sat_summary,
            #'Weekly Schedule': cashiers_reporting,
            #'HC Per Hour': headcount_per_hour1,
            #'HC Per Step': headcount_per_hour2,
        }

        #main_hourly_df_basis
        dataframe_hourly_basis = \
            {
             'Sunday': main_hourly_df_basis_sun,
             'Monday': main_hourly_df_basis_mon,
             'Tuesday': main_hourly_df_basis_tue,
             'Wednesday': main_hourly_df_basis_wed,
             'Thursday': main_hourly_df_basis_thur,
             'Friday': main_hourly_df_basis_fri,
             'Saturday': main_hourly_df_basis_sat,
            }

        for weekday_to_sched in ["Sunday", "Monday", "Tuesday", "Wednesday", 
                                 "Thursday", "Friday", "Saturday"]:
            
            #print(f"Scheduling for {weekday_to_sched}")
            avail_to_allocate = weekly_sched_df[weekly_sched_df[weekday_to_sched] == 1][["Personnel Name", "Employment Type"]]
            
            
            full_time_allocate = avail_to_allocate[avail_to_allocate["Employment Type"] == "Full-Time"].reset_index(drop = True)
            full_time_allocate["Allocated"] = 0
            full_time_allocate["To Allocate"] = 1
            
            part_time_allocate = avail_to_allocate[avail_to_allocate["Employment Type"] == "Part-Time"].reset_index(drop = True)
            part_time_allocate["Allocated"] = 0
            part_time_allocate["To Allocate"] = 1
            
            #Filter to Sunday
            main_hourly_df = \
                headcount_per_hour1[headcount_per_hour1["Day of Week"] == weekday_to_sched][numeric_col].T.iloc[:,0]
            
            main_hourly_df_basis = \
                main_hourly_df.fillna(0).copy()
            
            main_hour_df_basis_2 = \
                headcount_per_hour2[headcount_per_hour2["Day of Week"] == weekday_to_sched].T.iloc[:,0]
            
            allocated_personnel_count = 0
            
            allocated_hourly_sched_full_time = []
            allocated_hourly_sched_part_time = []
            
            Part1_loop = ["Step 1 Must Start at 8", 
                          "Step 3 Must Start at 12",
                          "Step 5 Must Start at 10",
                          "Step 2 Must Start at 9",
                          "Step 4 Must Start at 11"]
            
            start_dict = {"Step 1 Must Start at 8": 8, 
                          "Step 3 Must Start at 12": 12,
                          "Step 5 Must Start at 10": 10,
                          "Step 2 Must Start at 9": 9,
                          "Step 4 Must Start at 11": 11}
            
            
            
            main_hour_df_basis_2_part1 = main_hour_df_basis_2[Part1_loop] 
            
            above0 = main_hour_df_basis_2_part1 > 0
            
            main_hour_df_basis_2_part1_allocate = \
                pd.DataFrame(main_hour_df_basis_2_part1[above0])
                
            main_hour_df_basis_2_part1_allocate.columns = ["To Allocate"]
            
            min_hc = int(main_hour_df_basis_2_part1_allocate["To Allocate"].sum())
            
            main_hour_df_basis_2_part1_allocate["Start"] = \
                main_hour_df_basis_2_part1_allocate.index.map(start_dict)
            
            main_hour_df_basis_2_part1_allocate_basis = \
                main_hour_df_basis_2_part1_allocate.copy()
            
            
            for count in range(0, min(min_hc,len(avail_to_allocate))):
                
                above0 = \
                    main_hour_df_basis_2_part1_allocate_basis["To Allocate"] > 0
                
                main_hour_df_basis_2_part1_allocate_basis = \
                                main_hour_df_basis_2_part1_allocate_basis[above0]
                
                allocate_hours_limit_constraint = \
                    main_hour_df_basis_2_part1_allocate_basis["Start"].tolist()
                
                
                
                if allocated_personnel_count < full_time_allocate["To Allocate"].sum():
                    full_time_matrix_df_red = \
                        full_time_matrix_df[full_time_matrix_df["Start"].isin(allocate_hours_limit_constraint)]
                
                    max_coverage_df_full_time = \
                        pd.concat([full_time_matrix_df_red[["Start","Hour"]], 
                                   (full_time_matrix_df_red[numeric_col].fillna(0)*main_hourly_df_basis)], axis = 1)
                        
                    
                
                    max_coverage_df_full_time["Coverage"] = max_coverage_df_full_time[numeric_col].sum(axis = 1)
                
                    max_coverage_df_full_time = max_coverage_df_full_time.set_index(["Start", "Hour"])
                    start_hour, hour_break = max_coverage_df_full_time["Coverage"].idxmax()
                    
                    #Choose between 2 Breaks
                    start_at = full_time_matrix_df["Start"] == start_hour
                    
                    h3_hc = full_time_matrix_df["Hour"] == 3
                    h4_hc = full_time_matrix_df["Hour"] == 4
                    
                    hour_coverage_3rd = full_time_matrix_df[start_at][h3_hc].T.iloc[3:,0].fillna(0)
                    hour_coverage_4th = full_time_matrix_df[start_at][h4_hc].T.iloc[3:,0].fillna(0)
                    
                    if hour_break == 3 :
                        main_hourly_df_basis = main_hourly_df_basis - hour_coverage_3rd
                        
                        hour_basis = full_time_matrix_df[start_at][h3_hc].T.iloc[3:,0]
                    else: #hour_break == 4
                        main_hourly_df_basis = main_hourly_df_basis - hour_coverage_4th
                        
                        hour_basis = full_time_matrix_df[start_at][h4_hc].T.iloc[3:,0]
            
                    #print(f"Personnel: {full_time_allocate.loc[allocated_personnel_count, 'Personnel Name']}")
                    full_time_allocate.loc[allocated_personnel_count, "Allocated"] = 1
                    full_time_allocate.loc[allocated_personnel_count, "To Allocate"] = 1
                    full_time_allocate.loc[allocated_personnel_count, "Start at"] = start_hour
                    full_time_allocate.loc[allocated_personnel_count, "Hour Break"] = hour_break
                    full_time_allocate.loc[allocated_personnel_count, "Phase"] = "Phase 1"
                    full_time_allocate.loc[allocated_personnel_count, "Sched"] = "Start at " + str(start_hour) + ", Hour Break at hour " + str(hour_break)
                    #print(f"Full_Time to Allocate : {full_time_allocate['Allocated'].sum()}")
                    temp = pd.DataFrame(hour_basis).T
                    temp.reset_index(drop = True)
                    
                    allocated_hourly_sched_full_time.append(temp)
                    
                else: 
                    part_time_matrix_df_red = \
                        part_time_matrix_df[part_time_matrix_df["Start"].isin(allocate_hours_limit_constraint)]
            
                
                    max_coverage_df_part_time = \
                        pd.concat([part_time_matrix_df_red[["Start","Hour"]], 
                                   (part_time_matrix_df_red[numeric_col].fillna(0)*main_hourly_df_basis)], axis = 1)
                
                    max_coverage_df_part_time["Coverage"] = max_coverage_df_part_time[numeric_col].sum(axis = 1)
                
                    max_coverage_df_part_time = max_coverage_df_part_time.set_index(["Start", "Hour"])
                    start_hour, hour_break = max_coverage_df_part_time["Coverage"].idxmax()
                
                    start_at = part_time_matrix_df["Start"] == start_hour
               
                    h3_hc = part_time_matrix_df["Hour"] == 3
                    h4_hc = part_time_matrix_df["Hour"] == 4
                    
                    hour_coverage_3rd = part_time_matrix_df[start_at][h3_hc].T.iloc[3:,0].fillna(0)
                    hour_coverage_4th = part_time_matrix_df[start_at][h4_hc].T.iloc[3:,0].fillna(0)
                    
                    if hour_break == 3 :
                        main_hourly_df_basis = main_hourly_df_basis - hour_coverage_3rd
                        
                        hour_basis = part_time_matrix_df[start_at][h3_hc].T.iloc[3:,0]
                    else: 
                        main_hourly_df_basis = main_hourly_df_basis - hour_coverage_4th
                        
                        hour_basis = part_time_matrix_df[start_at][h4_hc].T.iloc[3:,0]
                
                    part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Allocated"] = 1
                    part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Start at"] = start_hour
                    part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Hour Break"] = hour_break        
                    part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Phase"] = "Phase 1"
                    part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Sched"] = "Start at " + str(start_hour) + ", Hour Break at hour " + str(hour_break)
                    
                    temp = pd.DataFrame(hour_basis).T
                    temp.reset_index(drop = True)
                
                    allocated_hourly_sched_part_time.append(temp)
                   
                allocated_personnel_count += 1
                main_hour_df_basis_2_part1_allocate_basis.loc[main_hour_df_basis_2_part1_allocate_basis["Start"] == start_hour,
                                                              "To Allocate"] -= 1
            

            if allocated_personnel_count < len(avail_to_allocate): 
                remaining = len(avail_to_allocate) - (allocated_personnel_count)
                
                #-------------------------Part 2 Loop
                Part2_loop = ["Step 6 Add To Cover 1800H", 
                              "Step 7 Add To Cover 1700H",
                              "Step 8 Add To Cover 1600H",
                              "Step 9 Add To Cover 1000H",
                              "Step 10 Add To Cover 1100H",
                              "Step 11 Add To Cover 1200H",
                              "Step 12 Add To Cover 1300H",
                              "Step 13 Add To Cover 1400H",
                              "Step 14 Add To Cover 1500H"]
                
                start_end_dict = {"Step 6 Add To Cover 1800H": (9, 12), 
                              "Step 7 Add To Cover 1700H": (8, 12),
                              "Step 8 Add To Cover 1600H": (8, 12),
                              "Step 9 Add To Cover 1000H": (8, 10),
                              "Step 10 Add To Cover 1100H": (8, 11),
                              "Step 11 Add To Cover 1200H": (8, 12),
                              "Step 12 Add To Cover 1300H": (8, 12),
                              "Step 13 Add To Cover 1400H": (8, 12),
                              "Step 14 Add To Cover 1500H": (8, 12)}      
                
                    
                main_hour_df_basis_2_part2 = main_hour_df_basis_2[Part2_loop] 
                
                above0 = main_hour_df_basis_2_part2 > 0
                
                main_hour_df_basis_2_part2_allocate = \
                    pd.DataFrame(main_hour_df_basis_2_part2[above0])
                
                if len(main_hour_df_basis_2_part2_allocate) > 0:
                    main_hour_df_basis_2_part2_allocate.columns = ["To Allocate"]
                    
                    min_hc = int(main_hour_df_basis_2_part2_allocate["To Allocate"].sum())
                    
                    main_hour_df_basis_2_part2_allocate["Start and End"] = \
                        main_hour_df_basis_2_part2_allocate.index.map(start_end_dict)
                    
                    main_hour_df_basis_2_part2_allocate["Start"], \
                        main_hour_df_basis_2_part2_allocate["End"] = \
                            zip(*main_hour_df_basis_2_part2_allocate["Start and End"])
                    
                    main_hour_df_basis_2_part2_allocate_basis = \
                        main_hour_df_basis_2_part2_allocate.copy()
                        
                            
                    for count in range(0, min(min_hc, remaining)):
                        above0 = \
                            main_hour_df_basis_2_part2_allocate_basis["To Allocate"] > 0
                        
                        main_hour_df_basis_2_part2_allocate_basis = \
                                        main_hour_df_basis_2_part2_allocate_basis[above0]
                        
                        part = main_hour_df_basis_2_part2_allocate_basis.index.tolist()[0]
                         
                        allocate_hours_limit_constraint_Start = \
                            main_hour_df_basis_2_part2_allocate_basis.loc[part, "Start"]
                        
                        allocate_hours_limit_constraint_End = \
                            main_hour_df_basis_2_part2_allocate_basis.loc[part, "End"]
                        
                        #-----------------------------Stopped Here--------------------------
                        if allocated_personnel_count < full_time_allocate["To Allocate"].sum():

                            full_time_matrix_df_red = \
                                full_time_matrix_df[full_time_matrix_df["Start"] >= allocate_hours_limit_constraint_Start]\
                                    [full_time_matrix_df["Start"] <= allocate_hours_limit_constraint_End]
                        
                            max_coverage_df_full_time = \
                                pd.concat([full_time_matrix_df_red[["Start","Hour"]], 
                                           (full_time_matrix_df_red[numeric_col].fillna(0)*main_hourly_df_basis)], axis = 1)
                                
                            max_coverage_df_full_time["Coverage"] = max_coverage_df_full_time[numeric_col].sum(axis = 1)
                        
                            max_coverage_df_full_time = max_coverage_df_full_time.set_index(["Start", "Hour"])
                            start_hour, hour_break = max_coverage_df_full_time["Coverage"].idxmax()
                            
                            #Choose between 2 Breaks
                            start_at = full_time_matrix_df["Start"] == start_hour
                            
                            h3_hc = full_time_matrix_df["Hour"] == 3
                            h4_hc = full_time_matrix_df["Hour"] == 4
                            
                            hour_coverage_3rd = full_time_matrix_df[start_at][h3_hc].T.iloc[3:,0].fillna(0)
                            hour_coverage_4th = full_time_matrix_df[start_at][h4_hc].T.iloc[3:,0].fillna(0)
                            
                            if hour_break == 3 :
                                main_hourly_df_basis = main_hourly_df_basis - hour_coverage_3rd
                                
                                hour_basis = full_time_matrix_df[start_at][h3_hc].T.iloc[3:,0]
                            else: #hour_break == 4
                                main_hourly_df_basis = main_hourly_df_basis - hour_coverage_4th
                                
                                hour_basis = full_time_matrix_df[start_at][h4_hc].T.iloc[3:,0]
                    
                            
                            full_time_allocate.loc[allocated_personnel_count, "Allocated"] = 1
                            full_time_allocate.loc[allocated_personnel_count, "To Allocate"] = 1
                            full_time_allocate.loc[allocated_personnel_count, "Start at"] = start_hour
                            full_time_allocate.loc[allocated_personnel_count, "Hour Break"] = hour_break
                            full_time_allocate.loc[allocated_personnel_count, "Phase"] = "Phase 2"
                            full_time_allocate.loc[allocated_personnel_count, "Sched"] = "Start at " + str(start_hour) + ", Hour Break at hour " + str(hour_break)
                            
                            temp = pd.DataFrame(hour_basis).T
                            temp.reset_index(drop = True)
                            
                            allocated_hourly_sched_full_time.append(temp)
                            
                        else: #Choose Part-Time
                            
                            part_time_matrix_df_red = \
                                part_time_matrix_df[part_time_matrix_df["Start"] >= allocate_hours_limit_constraint_Start]\
                                    [part_time_matrix_df["Start"] <= allocate_hours_limit_constraint_End]
                                    
                        
                            max_coverage_df_part_time = \
                                pd.concat([part_time_matrix_df_red[["Start","Hour"]], 
                                           (part_time_matrix_df_red[numeric_col].fillna(0)*main_hourly_df_basis)], axis = 1)
                        
                            
                            max_coverage_df_part_time["Coverage"] = max_coverage_df_part_time[numeric_col].sum(axis = 1)
                        
                            max_coverage_df_part_time = max_coverage_df_part_time.set_index(["Start", "Hour"])
                            start_hour, hour_break = max_coverage_df_part_time["Coverage"].idxmax()
                        
                            start_at = part_time_matrix_df["Start"] == start_hour
            
                            h3_hc = part_time_matrix_df["Hour"] == 3
                            h4_hc = part_time_matrix_df["Hour"] == 4
                            
                            hour_coverage_3rd = part_time_matrix_df[start_at][h3_hc].T.iloc[3:,0].fillna(0)
                            hour_coverage_4th = part_time_matrix_df[start_at][h4_hc].T.iloc[3:,0].fillna(0)
                            
                            if hour_break == 3 :
                                main_hourly_df_basis = main_hourly_df_basis - hour_coverage_3rd
                                
                                hour_basis = part_time_matrix_df[start_at][h3_hc].T.iloc[3:,0]
                            else: #hour_break == 4
                                main_hourly_df_basis = main_hourly_df_basis - hour_coverage_4th
                                
                                hour_basis = part_time_matrix_df[start_at][h4_hc].T.iloc[3:,0]
                        
                        
                            part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Allocated"] = 1
                            part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Start at"] = start_hour
                            part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Hour Break"] = hour_break
                            part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Phase"] = "Phase 2"
                            part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Sched"] = "Start at " + str(start_hour) + ", Hour Break at hour " + str(hour_break)
                            
                            temp = pd.DataFrame(hour_basis).T
                            temp.reset_index(drop = True)
                        
                            allocated_hourly_sched_part_time.append(temp)
                           
                        allocated_personnel_count += 1
                        main_hour_df_basis_2_part2_allocate_basis.loc[part,
                                                                      "To Allocate"] -= 1
                    
                    
            if allocated_personnel_count < len(avail_to_allocate): 
                remaining = len(avail_to_allocate) - (allocated_personnel_count)
                
                for count in range(0, remaining):
                    if allocated_personnel_count < full_time_allocate["To Allocate"].sum():
                        #print("Choose Full-Time")
                        max_coverage_df_full_time = \
                            pd.concat([full_time_matrix_df[["Start","Hour"]], 
                                       (full_time_matrix_df[numeric_col].fillna(0)*main_hourly_df_basis)], axis = 1)
                            
                        max_coverage_df_full_time["Coverage"] = max_coverage_df_full_time[numeric_col].sum(axis = 1)
                    
                        max_coverage_df_full_time = max_coverage_df_full_time.set_index(["Start", "Hour"])
                        start_hour, hour_break = max_coverage_df_full_time["Coverage"].idxmax()
                        
                        #Choose between 2 Breaks
                        start_at = full_time_matrix_df["Start"] == start_hour
                        
                        h3_hc = full_time_matrix_df["Hour"] == 3
                        h4_hc = full_time_matrix_df["Hour"] == 4
                        
                        hour_coverage_3rd = full_time_matrix_df[start_at][h3_hc].T.iloc[3:,0].fillna(0)
                        hour_coverage_4th = full_time_matrix_df[start_at][h4_hc].T.iloc[3:,0].fillna(0)
                        
                        if hour_break == 3 :
                            main_hourly_df_basis = main_hourly_df_basis - hour_coverage_3rd
                            
                            hour_basis = full_time_matrix_df[start_at][h3_hc].T.iloc[3:,0]
                        else: #hour_break == 4
                            main_hourly_df_basis = main_hourly_df_basis - hour_coverage_4th
                            
                            hour_basis = full_time_matrix_df[start_at][h4_hc].T.iloc[3:,0]
            
                        full_time_allocate.loc[allocated_personnel_count, "Allocated"] = 1
                        full_time_allocate.loc[allocated_personnel_count, "To Allocate"] = 1
                        full_time_allocate.loc[allocated_personnel_count, "Start at"] = start_hour
                        full_time_allocate.loc[allocated_personnel_count, "Hour Break"] = hour_break
                        full_time_allocate.loc[allocated_personnel_count, "Phase"] = "Phase 3"
                        full_time_allocate.loc[allocated_personnel_count, "Sched"] = "Start at " + str(start_hour) + ", Hour Break at hour " + str(hour_break)
                        
                        temp = pd.DataFrame(hour_basis).T
                        temp.reset_index(drop = True)
                        
                        allocated_hourly_sched_full_time.append(temp)
                        
                    else: #Choose Part-Time
                        #print("Choose Part-Time")
                        max_coverage_df_part_time = \
                            pd.concat([part_time_matrix_df[["Start","Hour"]], 
                                       (part_time_matrix_df[numeric_col].fillna(0)*main_hourly_df_basis)], axis = 1)
                        
                        max_coverage_df_part_time["Coverage"] = max_coverage_df_part_time[numeric_col].sum(axis = 1)
                    
                        max_coverage_df_part_time = max_coverage_df_part_time.set_index(["Start", "Hour"])
                        start_hour, hour_break = max_coverage_df_part_time["Coverage"].idxmax()
                    
                        start_at = part_time_matrix_df["Start"] == start_hour
            
                   
                        h3_hc = part_time_matrix_df["Hour"] == 3
                        h4_hc = part_time_matrix_df["Hour"] == 4
                        
                        hour_coverage_3rd = part_time_matrix_df[start_at][h3_hc].T.iloc[3:,0].fillna(0)
                        hour_coverage_4th = part_time_matrix_df[start_at][h4_hc].T.iloc[3:,0].fillna(0)
                        
                        if hour_break == 3 :
                            main_hourly_df_basis = main_hourly_df_basis - hour_coverage_3rd
                            
                            hour_basis = part_time_matrix_df[start_at][h3_hc].T.iloc[3:,0]
                        else: #hour_break == 4
                            main_hourly_df_basis = main_hourly_df_basis - hour_coverage_4th
                            
                            hour_basis = part_time_matrix_df[start_at][h4_hc].T.iloc[3:,0]
                    
                    
                        
                        part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Allocated"] = 1
                        part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Start at"] = start_hour
                        part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Hour Break"] = hour_break
                        part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Phase"] = "Phase 3"
                        part_time_allocate.loc[allocated_personnel_count - len(full_time_allocate), "Sched"] = "Start at " + str(start_hour) + ", Hour Break at hour " + str(hour_break)
                        #print(f"Part_Time to Allocate : {part_time_allocate['Allocated'].sum()}")
                        temp = pd.DataFrame(hour_basis).T
                        temp.reset_index(drop = True)
                    
                        allocated_hourly_sched_part_time.append(temp)
                       
                    allocated_personnel_count += 1
                    
                
            if len(allocated_hourly_sched_full_time) > 0:
                allocated_hourly_sched_time_combine_ft = pd.concat(allocated_hourly_sched_full_time).reset_index(drop = True)
                full_time_allocate = full_time_allocate.join(allocated_hourly_sched_time_combine_ft)
            
            if len(allocated_hourly_sched_part_time) > 0:
                allocated_hourly_sched_time_combine_pt = pd.concat(allocated_hourly_sched_part_time).reset_index(drop = True)
                part_time_allocate = part_time_allocate.join(allocated_hourly_sched_time_combine_pt)
                
            
            hourly_schedule_gen = pd.concat([full_time_allocate, part_time_allocate]).reset_index(drop = True)
            
            hourly_summary_sched_gen = pd.DataFrame()
            hourly_summary_sched_gen["HC Needed per Hour"] = main_hourly_df
            hourly_summary_sched_gen["HC Allocated"]  = hourly_schedule_gen[numeric_col].sum()
            hourly_summary_sched_gen["HC Remaining"] = main_hourly_df_basis
            hourly_summary_sched_gen = hourly_summary_sched_gen.T
            
            #Convert to Dictionary
            hourly_schedule_gen_store = hourly_schedule_gen.reset_index(drop = True)
            hourly_schedule_gen_store = hourly_schedule_gen_store.to_dict('records')
            
            hourly_summary_sched_gen_store = hourly_summary_sched_gen.reset_index()
            hourly_summary_sched_gen_store = hourly_summary_sched_gen_store.to_dict('records')
            
            
            dataframes[weekday_to_sched] = hourly_schedule_gen_store
            dataframes_summary[weekday_to_sched] = hourly_summary_sched_gen_store
            dataframe_hourly_basis[weekday_to_sched] = main_hourly_df_basis
            
            
        
        
        all_data = {"Main": dataframes,
                    "Summary": dataframes_summary,
                    "Hourly Basis": dataframe_hourly_basis}
        
        main_dict_results = all_data["Main"]
        
        columns_to_remove = ['Allocated', 'To Allocate',
                             'Start at', 'Hour Break',
                             'Phase'] 
        hourly_sched_df_sun = pd.DataFrame(main_dict_results["Sunday"])
        hourly_sched_df_sun = hourly_sched_df_sun.drop(columns=columns_to_remove)
        
        hourly_sched_df_mon = pd.DataFrame(main_dict_results["Monday"])
        hourly_sched_df_mon = hourly_sched_df_mon.drop(columns=columns_to_remove)
        
        hourly_sched_df_tue = pd.DataFrame(main_dict_results["Tuesday"])
        hourly_sched_df_tue = hourly_sched_df_tue.drop(columns=columns_to_remove)
        
        hourly_sched_df_wed = pd.DataFrame(main_dict_results["Wednesday"])
        hourly_sched_df_wed = hourly_sched_df_wed.drop(columns=columns_to_remove)
        
        hourly_sched_df_thu = pd.DataFrame(main_dict_results["Thursday"])
        hourly_sched_df_thu = hourly_sched_df_thu.drop(columns=columns_to_remove)
        
        hourly_sched_df_fri  = pd.DataFrame(main_dict_results["Friday"])
        hourly_sched_df_fri = hourly_sched_df_fri.drop(columns=columns_to_remove)
        
        hourly_sched_df_sat = pd.DataFrame(main_dict_results["Saturday"])
        hourly_sched_df_sat = hourly_sched_df_sat.drop(columns=columns_to_remove)
        
        sun_df_dash =  html.Div([dash_table.DataTable(id = 'hourly_sched_df_sun_table', 
                                         columns = [{'name': i, 'id': i} \
                                                    for i in hourly_sched_df_sun.columns], 
                                         data = hourly_sched_df_sun.to_dict('records'), 
                                         style_data={
                                                      'whiteSpace': 'normal',
                                                      'height': 'auto',
                                                      },
                                         style_table={'overflowY': 'auto'},
                                         page_size = 8,
                                          style_data_conditional = style_data_conditional,
                                         style_cell = style_cell_option_script, 
                                         style_header = style_header_option_script)
                                       ],
                                )
            
        mon_df_dash =  html.Div([dash_table.DataTable(id = 'hourly_sched_df_mon_table', 
                                         columns = [{'name': i, 'id': i} \
                                                    for i in hourly_sched_df_mon.columns], 
                                         data = hourly_sched_df_mon.to_dict('records'), 
                                         style_data={
                                                      'whiteSpace': 'normal',
                                                      'height': 'auto',
                                                      },
                                         style_table={'overflowY': 'auto'},
                                         page_size = 8,
                                          style_data_conditional = style_data_conditional,
                                         style_cell = style_cell_option_script, 
                                         style_header = style_header_option_script)
                                       ],
                                )
            
        tue_df_dash =  html.Div([dash_table.DataTable(id = 'hourly_sched_df_tue_table', 
                                             columns = [{'name': i, 'id': i} \
                                                        for i in hourly_sched_df_tue.columns], 
                                             data = hourly_sched_df_tue.to_dict('records'), 
                                             style_data={
                                                          'whiteSpace': 'normal',
                                                          'height': 'auto',
                                                          },
                                             style_table={'overflowY': 'auto'},
                                             page_size = 8,
                                              style_data_conditional = style_data_conditional,
                                             style_cell = style_cell_option_script, 
                                             style_header = style_header_option_script)
                                           ],
                                    )
        wed_df_dash =  html.Div([dash_table.DataTable(id = 'hourly_sched_df_wed_table', 
                                             columns = [{'name': i, 'id': i} \
                                                        for i in hourly_sched_df_wed.columns], 
                                             data = hourly_sched_df_wed.to_dict('records'), 
                                             style_data={
                                                          'whiteSpace': 'normal',
                                                          'height': 'auto',
                                                          },
                                             style_table={'overflowY': 'auto'},
                                             page_size = 8,
                                              style_data_conditional = style_data_conditional,
                                             style_cell = style_cell_option_script, 
                                             style_header = style_header_option_script)
                                           ],
                                    )
            
        thu_df_dash =  html.Div([dash_table.DataTable(id = 'hourly_sched_df_thu_table', 
                                             columns = [{'name': i, 'id': i} \
                                                        for i in hourly_sched_df_thu.columns], 
                                             data = hourly_sched_df_thu.to_dict('records'), 
                                             style_data={
                                                          'whiteSpace': 'normal',
                                                          'height': 'auto',
                                                          },
                                             style_table={'overflowY': 'auto'},
                                             page_size = 8,
                                              style_data_conditional = style_data_conditional,
                                             style_cell = style_cell_option_script, 
                                             style_header = style_header_option_script)
                                           ],
                                    )
            
        fri_df_dash =  html.Div([dash_table.DataTable(id = 'hourly_sched_df_fri_table', 
                                             columns = [{'name': i, 'id': i} \
                                                        for i in hourly_sched_df_fri.columns], 
                                             data = hourly_sched_df_fri.to_dict('records'), 
                                             style_data={
                                                          'whiteSpace': 'normal',
                                                          'height': 'auto',
                                                          },
                                             style_table={'overflowY': 'auto'},
                                             page_size = 8,
                                              style_data_conditional = style_data_conditional,
                                             style_cell = style_cell_option_script, 
                                             style_header = style_header_option_script)
                                           ],
                                    ) 
            
        sat_df_dash =  html.Div([dash_table.DataTable(id = 'hourly_sched_df_sat_table', 
                                             columns = [{'name': i, 'id': i} \
                                                        for i in hourly_sched_df_sat.columns], 
                                             data = hourly_sched_df_sat.to_dict('records'), 
                                             style_data={
                                                          'whiteSpace': 'normal',
                                                          'height': 'auto',
                                                          },
                                             style_table={'overflowY': 'auto'},
                                             page_size = 8,
                                              style_data_conditional = style_data_conditional,
                                             style_cell = style_cell_option_script, 
                                             style_header = style_header_option_script)
                                           ],
                                    ) 
            
        return weekly_sched_df_dash, \
            sun_df_dash, mon_df_dash, tue_df_dash, wed_df_dash, \
                thu_df_dash, fri_df_dash, sat_df_dash, \
            store_cashiers_reporting, store_weekly_sched_df,\
                store_weekly_sched_gen, all_data, \
            html.H5("Schedule Has Been Generated. Go to Tab 2. Weekly and Daily Schedule Generation.")
    
    






@app.callback(
    Output('cycle_time_table', 'data'),
    Input('cycle_time_table', 'data_previous'),
    State('cycle_time_table', 'data'),
    prevent_initial_call=True  # Prevents the initial call when the page loads
)
def update_cycle_time_data(data_previous, data):
    if data_previous is None or data == data_previous:
        return data  # No changes
    # Validate and limit values to the range [1, 60]
    for i, row in enumerate(data):
        for col, value in row.items():
            try:
                value = int(value)
                if col == 'Cycle Time in Mins':
                    if value <= 0:
                        data[i][col] = 1  # Ensure a minimum value of 1 for this column
                    
                elif not (0 <= value <= 60):
                    # Revert to the previous value if it's out of range
                    data[i][col] = data_previous[i][col]
                    
            except (ValueError, TypeError):
                # Revert to the previous value if the input is not a valid integer
                data[i][col] = data_previous[i][col]
    return data





@app.callback(
    [Output("transaction-output", 'children'),
     Output('cycle-time-data', 'data')],
    Input('cycle_time_table', "data")
)
def calcu_tc(data):
    if not data:
        return "No data available."

    # Convert the data to a DataFrame
    df = pd.DataFrame(data).astype(int)

    try:
        # Calculate transaction rates (tc) based on provided formula
        df['Transaction Rate'] = 60 / (df['Cycle Time in Mins'] + \
                                   (df['+ Cycle Time in Secs'] / 60) + \
                                       (df['Transition Time in Secs'] / 60))
            
        cycle_time_ = df[['Transaction Rate']].reset_index(drop = True)
        cycle_time_ = cycle_time_.to_dict('records')
        
        return html.Div([np.round(df['Transaction Rate'][0],2) ]),\
                    cycle_time_ 
    
    except Exception as e:
        return f"Error in calculation: {str(e)}", []
    
  

#Upload Data Table Personnel

def process_excel(file_contents):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(io.BytesIO(file_contents))

    # Check if the required columns exist
    if 'Personnel Name' not in df.columns or 'Employment Type' not in df.columns:
        return None, "Error: The Excel file must contain 'Personnel Name' and 'Employment Type' columns."

    # Convert the data type of columns to text
    df['Personnel Name'] = df['Personnel Name'].astype(str)
    df['Employment Type'] = \
            df['Employment Type'].apply(lambda x: 'Full-Time' if \
                                        pd.isna(x) or (x.lower() != 'part-time' \
                                            and x.lower() != 'part time'
                                            and x.lower() != 'Part-Time') \
                                            else 'Part-Time')
    return df, None



@app.callback(
    [Output('output-datatable-personnel', 'children'),
     Output('output-upload-personnel-data', 'children'),
     Output('personnel-data', 'data')],
    
    [Input('upload-data-personnel', 'contents')],
    [State('upload-data-personnel', 'filename')]
)
def update_output(contents, filename):
    if contents is None:
        return [html.Div(className = "header-article-upload" , 
                         children = ['No file selected.'])], \
            [html.Div("Please upload personnel data file with columns Personnel Name and Employment Type")],\
            []

    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)

    df, error_message = process_excel(decoded)

    if error_message:
        return [html.Div(className = "header-article-upload" , 
                         children = [error_message])], \
                [html.Div("Error in Upload")], []

    table = html.Div([
        html.H5(f'Uploaded Excel File: {filename}'),
        dash_table.DataTable(
            id = 'table',
            columns = [{'name': col, 'id': col} for col in df.columns],
            data = df.to_dict('records'),
            style_data = {
                              'whiteSpace': 'normal',
                              'height': 'auto',
                              },
                style_table={'overflowY': 'auto'},
                page_size = 12,
                style_data_conditional = style_data_conditional,
                style_cell = style_cell_option_script, 
                style_header = style_header_option_script
        )
    ])
    
    df_count = df.groupby("Employment Type").count().reset_index()\
                        [["Employment Type", "Personnel Name"]]
    
    try:
        full_time_count = df_count[df_count["Employment Type"] == "Full-Time"]["Personnel Name"][0]
    except KeyError:
        full_time_count = 0 
    
    #try:
        #part_time_count = df_count[df_count["Employment Type"] == "Part-Time"]["Personnel Name"][0]
    #except KeyError:
        #part_time_count = 0 
    
    part_time_count = len(df) - full_time_count
    
        
    store_data_df = df.reset_index(drop = True)
    store_data_df = store_data_df.to_dict('records')
    
    return [table], [html.Div(className = "header-article-upload" , 
                              children = [html.Div("Upload Successful"),
                                          html.Br(),
                                          html.Div(f"Total to Sched: {full_time_count + part_time_count}"),
                                          html.Div(f"Full-Time: {full_time_count}"),
                                          html.Div(f"Part-Time: {part_time_count}"),
                                          ])], \
                                store_data_df





def custom_sort(order, items):
    order_dict = {item: index for index, item in enumerate(order)}
    return sorted(items, key=lambda x: order_dict.get(x, float('inf')))





def process_excel_main_data(file_contents):
    
    global included_col
    
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(io.BytesIO(file_contents), skiprows=12).iloc[2:-2,:]
        
        if len(df) > 0:
            unclean_col = df.columns.tolist()
            new_col = []
            c_count = 0
            for c in unclean_col:
                c_new = str(c).lstrip(" ").rstrip(" ").lstrip("\xa0")
                new_col.append(c_new)
                c_count += 1
            df.columns = new_col
            
            # Check if the required columns exist"
            total_columns = []
            for col in included_col:
                if col not in df.columns:
                    total_columns.append(col)
        
            total_columns_string = ", ".join(total_columns)
            if len(total_columns) > 1:
                
                string_error = f"Error: The Excel file must contain {total_columns_string} columns."
                return None, string_error
            
            elif len(total_columns) == 1:
                string_error = f"Error: The Excel file must contain {total_columns_string} column."
                return None, string_error
                    
            return df, None
        
        else:
            string_error = f"Error: The Excel file must contain {', '.join(included_col)} column."
            return None, string_error
            
    except TypeError or AttributeError:
        string_error = f"Error: The Excel file must contain {', '.join(included_col)} column."
        return None, string_error
        
    
    
    


@app.callback(
    [Output('output-datatable-daily-sales', 'children'),
     Output('output-upload-data-daily-sales', 'children'),
     Output('headcount_per_hour1', 'data'),
     Output('headcount_per_hour2', 'data'),
     Output('output-datatable-headcount_per_hour1', 'children'),
     Output('output-datatable-headcount_per_hour2', 'children'),
     ],
    
    [Input('upload-data-daily-sales', 'contents'),
     Input('cycle-time-data', "data")],
    [State('upload-data-daily-sales', 'filename')]
)
def update_output_sales(contents, data, filename):
    global order_list
    global order_list1
    global numeric_col
    
    try:
        if contents is None:
            return [html.Div(className = "header-article-upload" , 
                             children = ['No file selected.'])], \
                [html.Div(f"Please upload personnel data file with columns {', '.join(included_col)}")], \
                [], [], [], []
    
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
    
        
        df, error_message = process_excel_main_data(decoded)
    
        if df is not None:
            col_new = df.columns.tolist()
            new_sales_data_col = ["Date", "Day of Week"] + col_new[2:]    
            df.columns = new_sales_data_col 
            
            col_new1 = custom_sort(order_list, new_sales_data_col)
            df = df[col_new1]
            
            
            if error_message:
                return [html.Div(className = "header-article-upload" , 
                                 children = [error_message])], \
                        [html.Div("Error in Upload")], [], [], [], []
        
            #cycle_df = pd.DataFrame(data).astype(int)   
            
            #from the back-end logic
            df["Total"] = df[["7.00-8.00",
                          "8.00-9.00", "9.00-10.00", "10.00-11.00", "11.00-12.00",
                          "12.00-13.00", "13.00-14.00", "14.00-15.00", "15.00-16.00",
                          "16.00-17.00", "17.00-18.00", "18.00-19.00", "19.00-20.00",
                          "20.00-21.00", "21.00-22.00"]].sum(axis = 1)

            df["Max"] = df[["7.00-8.00",
                          "8.00-9.00", "9.00-10.00", "10.00-11.00", "11.00-12.00",
                          "12.00-13.00", "13.00-14.00", "14.00-15.00", "15.00-16.00",
                          "16.00-17.00", "17.00-18.00", "18.00-19.00", "19.00-20.00",
                          "20.00-21.00", "21.00-22.00"]].max(axis = 1)
            
            trans_rate = data[0]["Transaction Rate"]
            
            df["HC Needed"] = np.ceil(df["Max"]/trans_rate)


            
            
            #add Headcount1
            headcount_per_hour1 = df[order_list1]
            #Convert to HC per Hourly
            for col in numeric_col:
                headcount_per_hour1[col] = headcount_per_hour1[col].astype(float).fillna(0)
                headcount_per_hour1[col] = np.ceil(headcount_per_hour1[col]/trans_rate)
                
            #Minimum in Hourly Sched
            #Step 1 = Must Start at 8
            headcount_per_hour2 = headcount_per_hour1[["Day of Week", "8.00-9.00"]]
            headcount_per_hour2.columns = ["Day of Week", "Step 1 Must Start at 8"]

            #Step 2 = Must Start at 9
            headcount_per_hour2["Step 2 Must Start at 9"] = \
                    (headcount_per_hour1["9.00-10.00"] - headcount_per_hour2["Step 1 Must Start at 8"]).clip(lower=0)

            #Step 3 = Must Start at 12 #To cover 2100 onwards
            headcount_per_hour2["Step 3 Must Start at 12"] = headcount_per_hour1["21.00-22.00"]
                
            #Step 4 = Must Start at 11 #To cover 2000 onwards
            headcount_per_hour2["Step 4 Must Start at 11"] = \
                    (headcount_per_hour1["20.00-21.00"] - headcount_per_hour2["Step 3 Must Start at 12"]).clip(lower=0)


            #Step 5 = Must Start at 10 #To cover 1900 onwards
            headcount_per_hour2["Step 5 Must Start at 10"] = \
                    (headcount_per_hour1["19.00-20.00"] - headcount_per_hour2["Step 4 Must Start at 11"] - \
                     headcount_per_hour2["Step 3 Must Start at 12"]).clip(lower=0)

            #Step 6 Additional to cover 1800H that can Start at 9 to 12
            headcount_per_hour2["Step 6 Add To Cover 1800H"] = \
                    (headcount_per_hour1["18.00-19.00"] - \
                     headcount_per_hour2["Step 2 Must Start at 9"]  - \
                     headcount_per_hour2["Step 3 Must Start at 12"] - \
                     headcount_per_hour2["Step 4 Must Start at 11"] - \
                     headcount_per_hour2["Step 5 Must Start at 10"]).clip(lower=0)
                        
                        
            #Step 7 Additional to cover 1700H that can Start at 8 to 12
            headcount_per_hour2["Step 7 Add To Cover 1700H"] = \
                    (headcount_per_hour1["17.00-18.00"] - \
                     headcount_per_hour2["Step 1 Must Start at 8"] - \
                     headcount_per_hour2["Step 2 Must Start at 9"]  - \
                     headcount_per_hour2["Step 3 Must Start at 12"] - \
                     headcount_per_hour2["Step 4 Must Start at 11"] - \
                     headcount_per_hour2["Step 5 Must Start at 10"]).clip(lower=0)

            #Step 8 Additional to cover 1600H that can Start at 8 to 12
            headcount_per_hour2["Step 8 Add To Cover 1600H"] = \
                    (headcount_per_hour1["16.00-17.00"] - \
                     headcount_per_hour2["Step 1 Must Start at 8"] - \
                     headcount_per_hour2["Step 2 Must Start at 9"]  - \
                     headcount_per_hour2["Step 3 Must Start at 12"] - \
                     headcount_per_hour2["Step 4 Must Start at 11"] - \
                     headcount_per_hour2["Step 5 Must Start at 10"]).clip(lower=0)

            #Step 9 Additional to cover 1000H
            headcount_per_hour2["Step 9 Add To Cover 1000H"] = \
                    (headcount_per_hour1["10.00-11.00"] - \
                     np.ceil(headcount_per_hour2["Step 1 Must Start at 8"]*0.5) - \
                     headcount_per_hour2["Step 2 Must Start at 9"]  - \
                     headcount_per_hour2["Step 5 Must Start at 10"] - \
                         headcount_per_hour2["Step 6 Add To Cover 1800H"] - \
                         headcount_per_hour2["Step 7 Add To Cover 1700H"] - \
                         headcount_per_hour2["Step 8 Add To Cover 1600H"] ).clip(lower=0)

            #Step 10 Additional to cover 1100H
            headcount_per_hour2["Step 10 Add To Cover 1100H"] = \
                    (headcount_per_hour1["11.00-12.00"] - \
                     np.ceil(headcount_per_hour2["Step 1 Must Start at 8"]*0.5) - \
                      np.ceil(headcount_per_hour2["Step 2 Must Start at 9"]*0.5)  - \
                          headcount_per_hour2["Step 4 Must Start at 11"] - \
                     headcount_per_hour2["Step 5 Must Start at 10"] - \
                         headcount_per_hour2["Step 6 Add To Cover 1800H"] - \
                         headcount_per_hour2["Step 7 Add To Cover 1700H"] - \
                         headcount_per_hour2["Step 8 Add To Cover 1600H"] ).clip(lower=0)
                        
            #Step 11 Additional to cover 1200H
            headcount_per_hour2["Step 11 Add To Cover 1200H"] = \
                    (headcount_per_hour1["12.00-13.00"] - \
                     headcount_per_hour2["Step 1 Must Start at 8"] - \
                      np.ceil(headcount_per_hour2["Step 2 Must Start at 9"]*0.5)  - \
                          headcount_per_hour2["Step 3 Must Start at 12"] - \
                          headcount_per_hour2["Step 4 Must Start at 11"] - \
                      np.ceil(headcount_per_hour2["Step 5 Must Start at 10"]*0.5) - \
                         headcount_per_hour2["Step 6 Add To Cover 1800H"] - \
                         headcount_per_hour2["Step 7 Add To Cover 1700H"] - \
                         headcount_per_hour2["Step 8 Add To Cover 1600H"] ).clip(lower=0)
                        
            #Step 12 Additional to cover 1300H
            headcount_per_hour2["Step 12 Add To Cover 1300H"] = \
                    (headcount_per_hour1["13.00-14.00"] - \
                     headcount_per_hour2["Step 1 Must Start at 8"] - \
                      headcount_per_hour2["Step 2 Must Start at 9"] - \
                          headcount_per_hour2["Step 3 Must Start at 12"]  - \
                         np.ceil( headcount_per_hour2["Step 4 Must Start at 11"]*0.5) - \
                      np.ceil(headcount_per_hour2["Step 5 Must Start at 10"]*0.5) - \
                         np.ceil(headcount_per_hour2["Step 6 Add To Cover 1800H"]*0.75) - \
                         np.ceil(headcount_per_hour2["Step 7 Add To Cover 1700H"]*0.8) - \
                         np.ceil(headcount_per_hour2["Step 8 Add To Cover 1600H"]*0.8) ).clip(lower=0)
                        
            #Step 13 Additional to cover 1400H
            headcount_per_hour2["Step 13 Add To Cover 1400H"] = \
                    (headcount_per_hour1["14.00-15.00"] - \
                     headcount_per_hour2["Step 1 Must Start at 8"] - \
                      headcount_per_hour2["Step 2 Must Start at 9"] - \
                          np.ceil(headcount_per_hour2["Step 3 Must Start at 12"]*0.5)  - \
                         np.ceil( headcount_per_hour2["Step 4 Must Start at 11"]*0.5) - \
                      headcount_per_hour2["Step 5 Must Start at 10"] - \
                         np.ceil(headcount_per_hour2["Step 6 Add To Cover 1800H"]*0.75) - \
                         np.ceil(headcount_per_hour2["Step 7 Add To Cover 1700H"]*0.8) - \
                         np.ceil(headcount_per_hour2["Step 8 Add To Cover 1600H"]*0.8) ).clip(lower=0)
                        
            #Step 14 Additional to cover 1500H
            headcount_per_hour2["Step 14 Add To Cover 1500H"] = \
                    (headcount_per_hour1["15.00-16.00"] - \
                     headcount_per_hour2["Step 1 Must Start at 8"] - \
                      headcount_per_hour2["Step 2 Must Start at 9"] - \
                          np.ceil(headcount_per_hour2["Step 3 Must Start at 12"]*0.5)  - \
                          headcount_per_hour2["Step 4 Must Start at 11"] - \
                      headcount_per_hour2["Step 5 Must Start at 10"] - \
                    np.ceil(headcount_per_hour2["Step 6 Add To Cover 1800H"]*0.875) - \
                    np.ceil(headcount_per_hour2["Step 7 Add To Cover 1700H"]*0.9) - \
                    np.ceil(headcount_per_hour2["Step 8 Add To Cover 1600H"]*0.9) ).clip(lower=0)

            headcount_per_hour2["Total Required"] = \
                headcount_per_hour2[headcount_per_hour2.columns.tolist()[1:]].sum(axis = 1)


            headcount_per_hour2["HC Needed"] = df["HC Needed"]
            headcount_per_hour2["New HC Needed"] = headcount_per_hour2[["HC Needed", "Total Required"]].max(axis=1)
            
            store_data_df_hour1 = headcount_per_hour1.reset_index(drop = True)
            store_data_df_hour2 = headcount_per_hour2.reset_index(drop = True)
            
            store_data_df_hour1 = store_data_df_hour1.to_dict('records')
            store_data_df_hour2 = store_data_df_hour2.to_dict('records')
            df["New HC Needed"] = headcount_per_hour2["New HC Needed"]
            
            #Reduced Columns
            columns_to_remove = ["HC Needed",
                                 "Total",
                                 "Max"]
            df_red = df.drop(columns=columns_to_remove)
            
            table = html.Div([
                html.H5(f'Uploaded Excel File: {filename}'),
                dash_table.DataTable(
                    id = 'table',
                    columns = [{'name': col, 'id': col} for col in df_red.columns],
                    data = df_red.to_dict('records'),
                    style_data = {
                                      'whiteSpace': 'normal',
                                      'height': 'auto',
                                      },
                        style_table={'overflowY': 'auto'},
                        #page_size = 10,
                        style_data_conditional = style_data_conditional,
                        style_cell = style_cell_option_script, 
                        style_header = style_header_option_script
                )
            ])
            
            return [table], [html.Div(className = "header-article-upload" , 
                                      children = ["Upload Successful"])], \
                    store_data_df_hour1, store_data_df_hour2, \
                    [], []
        
        else:
            return [html.Div(className = "header-article-upload" , 
                             children = [f"Error. Change File. Please upload sales data file with columns {', '.join(included_col)}"])], \
                [html.Div(f"Please upload sales data file with columns {', '.join(included_col)}")], [], [], [], []
                
    except TypeError or AttributeError:
        return [html.Div(className = "header-article-upload" , 
                         children = [f"Error. Change File. Please upload sales data file with columns {', '.join(included_col)}"])], \
            [html.Div(f"Please upload sales data file with columns {', '.join(included_col)}")], [], [], [], []






@app.callback(
    [Output("download_component", "data"),
     Output("export-button-status-text", "children"),
     Output("download-link", "style")],
    [Input("export-button", "n_clicks")],
    [State('hourly_sched_df_sun_table', 'derived_virtual_data'),
     State('hourly_sched_df_mon_table', 'derived_virtual_data'),
     State('hourly_sched_df_tue_table', 'derived_virtual_data'),
     State('hourly_sched_df_wed_table', 'derived_virtual_data'),
     State('hourly_sched_df_thu_table', 'derived_virtual_data'),
     State('hourly_sched_df_fri_table', 'derived_virtual_data'),
     State('hourly_sched_df_sat_table', 'derived_virtual_data')],
    prevent_initial_call=True,
)
def export_to_excel(n_clicks, sun_data, mon_data, tue_data,
                    wed_data, thu_data, fri_data, sat_data):
    
    reorder_col = ["Day","Personnel Name","Employment Type","Sched",
                       "8.00-9.00","9.00-10.00","10.00-11.00","11.00-12.00",
                       "12.00-13.00","13.00-14.00", "14.00-15.00","15.00-16.00",
                       "16.00-17.00","17.00-18.00","18.00-19.00","19.00-20.00",
                       "20.00-21.00","21.00-22.00"]
    
    #reordered_columns = [{'name': col, 'id': col} for col in desired_order]
    days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    dataframes = [pd.DataFrame(day_data) if day_data else pd.DataFrame(columns=reorder_col) \
                  for day_data in [sun_data, mon_data, tue_data, wed_data, thu_data, fri_data, sat_data]]

     # Add "Day" column and reorder columns
    for day, df in zip(days, dataframes):
        df["Day"] = day
        df = df[reorder_col]
        
    # Concatenate all DataFrames
    dff = pd.concat(dataframes)
    """
    sun_dff = pd.DataFrame(sun_data)
    sun_dff["Day"] = "Sun"
    sun_dff = sun_dff.reindex(columns=[reorder_col])
    
    mon_dff = pd.DataFrame(mon_data)
    mon_dff["Day"] = "Mon"
    mon_dff = mon_dff.reindex(columns=[reorder_col])
    
    tue_dff = pd.DataFrame(tue_data)
    tue_dff["Day"] = "Tue"
    tue_dff = tue_dff.reindex(columns=[reorder_col])
    
    wed_dff = pd.DataFrame(wed_data)
    wed_dff["Day"] = "Wed"
    wed_dff = wed_dff.reindex(columns=[reorder_col])
    
    thu_dff = pd.DataFrame(thu_data)
    thu_dff["Day"] = "Thu"
    thu_dff = thu_dff.reindex(columns=[reorder_col])
    
    fri_dff = pd.DataFrame(fri_data)
    fri_dff["Day"] = "Fri"
    fri_dff = fri_dff.reindex(columns=[reorder_col])
    
    sat_dff = pd.DataFrame(sat_data)
    sat_dff["Day"] = "Sat"
    sat_dff = sat_dff.reindex(columns=[reorder_col])
    
    dff = pd.concat([sun_dff,
                     mon_dff,
                     tue_dff,
                     wed_dff, thu_dff,
                     fri_dff, sat_dff])
    
    """
    # Specify the filename in the to_csv method
    csv_string = dff.to_csv(index=False, encoding="utf-8")
    
    return csv_string, html.H3("Click on Download Link"), {'display': 'block'}


@app.callback(
    Output("download-link", "href"),
    Input(download_component, "data"),
    prevent_initial_call=True,
)
def update_download_link(data):
    if not data:
        raise PreventUpdate

    return f"data:text/csv;charset=utf-8,{data}"


if __name__ == "__main__":
    app.run_server(debug = True)
    
