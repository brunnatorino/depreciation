import matplotlib.pyplot as plt
import pandas as pd
from gooey import Gooey, GooeyParser
import numpy as np
import xlrd
from datetime import datetime, date
import datetime as dt
from datetime import datetime, timedelta
import plotly.graph_objects as go
from plotly.subplots import make_subplots

@Gooey(program_name="DEPRECIATION",navigation='TABBED', header_bg_color = '#ffedfe',default_size=(710, 700))
def parse_args():

    parser = GooeyParser()
    
    parser.add_argument('Fixed_Assets_File',
                        action='store',
                        widget='FileChooser',
                        help="Excel file with all fixed assets to depreciate")
    parser.add_argument('Current_Year',
                    action='store',
                    help="Current Financial Year in format YYYY")


    parser.add_argument('Country',
                        widget='Dropdown',
                        choices=['Germany','France','Poland','United Kingdom','Netherlands','Sweden','Spain',],
                        help="Choose the country")
    
    parser.add_argument('Method_of_Depreciation',
                        widget='Dropdown',
                        action = 'store',
                        choices=['Useful Life','Depreciation Rate'],
                        help='Choose Depreciation Method')
    
    tables = parser.add_argument_group('Supporting Graphs and Tables')

    tables.add_argument('Tables',
                        widget='Dropdown',
                        action = 'store',
                        choices=['YES','NO'],
                        help="Produce Depreciation Table?")
    
    tables.add_argument('Graphs',
                        widget='Dropdown',
                        choices=['YES','NO'],
                        action = 'store',
                        help="Produce Depreciation Graph?")
    output = parser.add_argument_group('Depreciation File Name')
    
    output.add_argument('Output_File_Name',
                        action='store',
                        help="Name of the output file with .xlsx",
                        gooey_options={
                             'validator': {
                                 'test': 'user_input.endswith(".xlsx") == True',
                                 'message': 'Must contain .xlsx at the end!'
                                 }
                             })



    
    args = parser.parse_args()
    return args

def country_select(assets):

    assets = pd.read_excel(assets)
    assets['Asset_Class'] = assets['Asset Class']
    assets['Initial_Value'] = assets['APC FY start']
    assets['Acquisition Date'] = assets['Capitalized on']
    assets['Value'] = assets['Initial_Value'] + assets['Acquisition'] + assets['Retirement'] + assets['Transfer']
    dfA = assets.dropna(subset = ['Acquisition Date'])
    dfA = dfA[(dfA.Initial_Value != 0) | (dfA.Value != 0) | (dfA.Transfer != 0) | (dfA.Acquisition != 0)|(dfA.Retirement != 0)]
 

    dfA['Acquisition Date'] = pd.to_datetime(dfA['Acquisition Date'],format='%d%m%Y')
    dfA['year'], dfA['month'], dfA['day'] = dfA['Acquisition Date'].dt.year, dfA['Acquisition Date'].dt.month,dfA['Acquisition Date'].dt.day
    
    dfA['Closing'] = pd.to_datetime(dfA['Closing'],format='%d%m%Y')
    dfA['Beginning'] = pd.to_datetime(dfA['Beginning'] ,format='%d%m%Y')

    dfA.loc[dfA["day"] != 1,'Depreciation_Start'] = dfA['Acquisition Date'] + pd.offsets.MonthBegin(1)
    dfA.loc[dfA["day"] == 1,'Depreciation_Start'] = dfA['Acquisition Date']

    dfA['Months_Past'] = ((dfA['Beginning'] - dfA['Depreciation_Start'])/np.timedelta64(1, 'M')).astype(int) + 1
    ## sets to zero all the assets bought this year// they haven't been depreciated yet 
    dfA.loc[dfA["Months_Past"] <= 0,'Months_Past'] = 0
    
    if 'Germany' in args.Country:
        master_file = pd.read_excel("germany_depreciation.xlsx")
        master_file = master_file.dropna(subset = ['Years'])
        
        master_years = master_file[['Asset','Years']]
        master_rate = master_file[['Asset','Depreciation Rate']]
        
        master_years = master_years.set_index('Asset').to_dict()['Years']
        dfA['Years'] = dfA['Asset']
        dfA['Years'] = dfA['Years'].map(master_years)
        
        master_rate = master_rate.set_index('Asset').to_dict()['Depreciation Rate']
        dfA['Depreciation Rate'] = dfA['Asset']
        dfA['Depreciation Rate'] = dfA['Depreciation Rate'].map(master_rate)
        
        dfA.loc[dfA["Depreciation_Start"].dt.year > 2017,'Threshold'] = 250                                                                                                                                                                                                                                                         
        dfA.loc[dfA["Depreciation_Start"].dt.year < 2008 ,'Threshold'] = 410
        dfA.loc[(dfA["Depreciation_Start"].dt.year >= 2008) & (dfA["Depreciation_Start"].dt.year <= 2017),'Threshold'] = 150

        dfA = dfA[(dfA.Value >= dfA.Threshold) | (dfA.Value == 0)]
        
    elif 'France' in args.Country:
        master_file = pd.read_excel("france_depreciation.xlsx")
        master_file = master_file.dropna(subset = ['Years'])
        master_years = master_file[['Asset','Years']]
        master_rate = master_file[['Asset','Depreciation Rate']]

        master_years = master_years.set_index('Asset').to_dict()['Years']
        dfA['Years'] = dfA['Asset']
        dfA['Years'] = dfA['Years'].map(master_years)
        
        master_rate = master_rate.set_index('Asset').to_dict()['Depreciation Rate']
        dfA['Depreciation Rate'] = dfA['Asset']
        dfA['Depreciation Rate'] = dfA['Depreciation Rate'].map(master_rate)

        dfA['Threshold'] = 500
        dfA = dfA[(dfA.Value >= dfA.Threshold) | (dfA.Value == 0)]

    elif 'Spain' in args.Country:
        master_file = pd.read_excel("spain_depreciation.xlsx")
        master_file = master_file.dropna(subset = ['Years'])
        master_years = master_file[['Asset','Years']]
        master_rate = master_file[['Asset','Depreciation Rate']]

        master_years = master_years.set_index('Asset').to_dict()['Years']
        dfA['Years'] = dfA['Asset']
        dfA['Years'] = dfA['Years'].map(master_years)
        
        master_rate = master_rate.set_index('Asset').to_dict()['Depreciation Rate']
        dfA['Depreciation Rate'] = dfA['Asset']
        dfA['Depreciation Rate'] = dfA['Depreciation Rate'].map(master_rate)


        dfA['Threshold'] = 300
        dfA = dfA[(dfA.Value >= dfA.Threshold) | (dfA.Value == 0)]

    elif 'Poland' in args.Country:
        master_file = pd.read_excel("poland_depreciation.xlsx")
        master_file = master_file.dropna(subset = ['Years'])
        master_years = master_file[['Asset','Years']]
        master_rate = master_file[['Asset','Depreciation Rate']]

        master_years = master_years.set_index('Asset').to_dict()['Years']
        dfA['Years'] = dfA['Asset']
        dfA['Years'] = dfA['Years'].map(master_years)
        
        master_rate = master_rate.set_index('Asset').to_dict()['Depreciation Rate']
        dfA['Depreciation Rate'] = dfA['Asset']
        dfA['Depreciation Rate'] = dfA['Depreciation Rate'].map(master_rate)

        dfA.loc[dfA["Depreciation_Start"].dt.year < 2018,'Threshold'] = 3500
        dfA.loc[dfA["Depreciation_Start"].dt.year >= 2018 ,'Threshold'] = 10000
        
        dfA = dfA[(dfA.Value >= dfA.Threshold) | (dfA.Value == 0)]


    elif 'United Kindgom' in args.Country:
        master_file = pd.read_excel("uk_depreciation.xlsx")
        master_file = master_file.dropna(subset = ['Years'])

        master_years = master_file[['Asset','Years']]
        master_rate = master_file[['Years','Depreciation Rate']]

        master_years = master_years.set_index('Asset').to_dict()['Years']
        dfA['Years'] = dfA['Asset_Class'].map(master_years)

        master_rate = master_rate.set_index('Years').to_dict()['Depreciation Rate']

        dfA['Depreciation Rate'] = dfA['Years'].map(master_rate)


    elif 'Netherlands' in args.Country:
        master_file = pd.read_excel("nl_depreciation.xlsx")
        master_file = master_file.dropna(subset = ['Years'])
        master_years = master_file[['Asset','Years']]
        master_rate = master_file[['Asset','Depreciation Rate']]

        master_years = master_years.set_index('Asset').to_dict()['Years']
        dfA['Years'] = dfA['Asset']
        dfA['Years'] = dfA['Years'].map(master_years)
        
        master_rate = master_rate.set_index('Asset').to_dict()['Depreciation Rate']
        dfA['Depreciation Rate'] = dfA['Asset']
        dfA['Depreciation Rate'] = dfA['Depreciation Rate'].map(master_rate)

        dfA['Threshold'] = 450
        dfA = dfA[(dfA.Value >= dfA.Threshold) | (dfA.Value == 0)]

    elif 'Sweden' in args.Country:
        master_file = pd.read_excel("sweden_depreciation.xlsx")
        master_file = master_file.dropna(subset = ['Years'])
        master_years = master_file[['Asset','Years']]
        master_rate = master_file[['Asset','Depreciation Rate']]

        master_years = master_years.set_index('Asset').to_dict()['Years']
        dfA['Years'] = dfA['Asset']
        dfA['Years'] = dfA['Years'].map(master_years)
        
        master_rate = master_rate.set_index('Asset').to_dict()['Depreciation Rate']
        dfA['Depreciation Rate'] = dfA['Asset']
        dfA['Depreciation Rate'] = dfA['Depreciation Rate'].map(master_rate)

        dfA['Threshold'] = 20000
        dfA = dfA[(dfA.Value >= dfA.Threshold) | (dfA.Value == 0)]
        

    return dfA
        

def time_calc(dfA, current):

    prev = current - 1
    dfA['Life_In_Months'] = dfA['Years'].mul(12)
    dfA['Depreciation_Past'] = dfA['APC FY start'].div(dfA['Life_In_Months'])
    dfA['Depreciated_Amount'] = dfA['Depreciation_Past'].mul(dfA['Months_Past'])
    dfA['Balance_Start'] = dfA['APC FY start'] - dfA['Depreciated_Amount']
    dfA['Months_Left'] = dfA['Life_In_Months']- dfA['Months_Past']
    dfA.loc[dfA["Months_Left"] > 12,'Months_Left'] = 12
    dfA['Value_Left_To_Depreciate'] = dfA['Value'] - dfA['Depreciated_Amount']
    dfA['Depreciation_Per_Month'] = dfA['Value_Left_To_Depreciate'].div(dfA['Months_Left'])
    dfA = dfA.query('Balance_Start >= 0')
    

    if dfA['Depreciation_Start'].dt.year.eq(current).sum():

        dfA.loc[dfA['Depreciation_Start'].dt.month == 1,'Depreciation_This_Year_In_Months'] = 9
        dfA.loc[dfA['Depreciation_Start'].dt.month == 2,'Depreciation_This_Year_In_Months'] = 8
        dfA.loc[dfA['Depreciation_Start'].dt.month == 3,'Depreciation_This_Year_In_Months'] = 7
        dfA.loc[dfA['Depreciation_Start'].dt.month == 4,'Depreciation_This_Year_In_Months'] = 6
        dfA.loc[dfA['Depreciation_Start'].dt.month == 5,'Depreciation_This_Year_In_Months'] = 5
        dfA.loc[dfA['Depreciation_Start'].dt.month == 6,'Depreciation_This_Year_In_Months'] = 4
        dfA.loc[dfA['Depreciation_Start'].dt.month == 7,'Depreciation_This_Year_In_Months'] = 3
        dfA.loc[dfA['Depreciation_Start'].dt.month == 8,'Depreciation_This_Year_In_Months'] = 2
        dfA.loc[dfA['Depreciation_Start'].dt.month == 9,'Depreciation_This_Year_In_Months'] = 1
        dfA.loc[dfA['Depreciation_Start'].dt.month == 10,'Depreciation_This_Year_In_Months'] = 12
        dfA.loc[dfA['Depreciation_Start'].dt.month == 11,'Depreciation_This_Year_In_Months'] = 11
        dfA.loc[dfA['Depreciation_Start'].dt.month == 12,'Depreciation_This_Year_In_Months'] = 10

    elif dfA['Depreciation_Start'].dt.year == prev:

        if dfA['Depreciation_Start'].dt.month >= 10:
            dfA.loc[dfA['Depreciation_Start'].dt.month == 1,'Depreciation_This_Year_In_Months'] = 9
            dfA.loc[dfA['Depreciation_Start'].dt.month == 2,'Depreciation_This_Year_In_Months'] = 8
            dfA.loc[dfA['Depreciation_Start'].dt.month == 3,'Depreciation_This_Year_In_Months'] = 7
            dfA.loc[dfA['Depreciation_Start'].dt.month == 4,'Depreciation_This_Year_In_Months'] = 6
            dfA.loc[dfA['Depreciation_Start'].dt.month == 5,'Depreciation_This_Year_In_Months'] = 5
            dfA.loc[dfA['Depreciation_Start'].dt.month == 6,'Depreciation_This_Year_In_Months'] = 4
            dfA.loc[dfA['Depreciation_Start'].dt.month == 7,'Depreciation_This_Year_In_Months'] = 3
            dfA.loc[dfA['Depreciation_Start'].dt.month == 8,'Depreciation_This_Year_In_Months'] = 2
            dfA.loc[dfA['Depreciation_Start'].dt.month == 9,'Depreciation_This_Year_In_Months'] = 1
            dfA.loc[dfA['Depreciation_Start'].dt.month == 10,'Depreciation_This_Year_In_Months'] = 12
            dfA.loc[dfA['Depreciation_Start'].dt.month == 11,'Depreciation_This_Year_In_Months'] = 11
            dfA.loc[dfA['Depreciation_Start'].dt.month == 12,'Depreciation_This_Year_In_Months'] = 10
        else:
            dfA.loc[dfA['Depreciation_Start'].dt.month < 10,'Depreciation_This_Year_In_Months'] = 12

    elif dfA['Depreciation_Start'].dt.year < prev:
        dfA['Depreciation_This_Year_In_Months'] = 12

    dfA['Amount_To_Depreciate'] = dfA['Depreciation_This_Year_In_Months'].mul(dfA['Depreciation_Per_Month'])

    dfA.loc[dfA["Balance_Start"] < dfA['Amount_To_Depreciate'],'Depreciation_This_Year_In_Months'] = dfA['Balance_Start'].div(dfA['Depreciation_Per_Month']).round(0)
    dfA.loc[dfA["Balance_Start"] < dfA['Amount_To_Depreciate'],'Amount_To_Depreciate'] = dfA['Balance_Start']
    
    dfA['End_Book_Value'] = dfA['Value'] - dfA['Amount_To_Depreciate'] - dfA['Depreciated_Amount']
    
    dfA.loc[dfA["End_Book_Value"] < 1,'Note'] = 'End of Life in Current fiscal Year'
    dfA.loc[dfA["Value"] == 0,'Note'] = 'Historical Asset'

    dfA.loc[dfA["End_Book_Value"] < 0,'End_Book_Value'] = 0
    dfA.loc[dfA["End_Book_Value"] < 1,'End_Book_Value'] = 0
    dfA['Reclass'] = 'No'

    dfA.loc[dfA["Amount_To_Depreciate"] == 0,'Reclass'] = 'Possible Reclass - Check' 
    dfA.loc[(dfA["Transfer"] != 0), 'Reclass'] = 'Yes'

    dfA = dfA.round(2)

    return dfA

def graphs_make(dfA):

    if args.Graphs == 'YES':

        fig = go.Figure()
        fig.add_trace(go.Bar(x=dfA['Asset Class'],
                        y=dfA['Depreciated_Amount'],
                        name='Amount Depreciated',
                        marker_color='rgb(55, 83, 109)',
                        ))
        fig.add_trace(go.Bar(x=dfA['Asset Class'],
                        y=dfA['Value'],
                        name='Total Value',
                        marker_color='rgb(26, 118, 255)'
                        ))
        fig.update_layout(
            title='Depreciation Chart',
            xaxis_tickfont_size=14,
            yaxis=dict(
                title='USD',
                titlefont_size=16,
                tickfont_size=14,
            ),
            legend=dict(
                x=0,
                y=1.0,
                bgcolor='rgba(255, 255, 255, 0)',
                bordercolor='rgba(255, 255, 255, 0)'
            ),
            barmode='group',
            bargap=0.15,
            bargroupgap=0.1 
        )
        fig.show()

def tables_make(dfA):
        
    if args.Tables == 'YES':
        dfA['Asset_Description'] = dfA['Asset description']
        del dfA['Asset description']
        headerColor = 'royalblue'
        rowEvenColor = 'aliceblue'
        rowOddColor = 'white'

        fig1 = go.Figure(data=[go.Table(
          header=dict(
            values=['<b>Asset</b>','<b>Asset Description</b>','<b>Value</b>','<b>Amount Depreciated</b>','<b>Amount To Depreciate</b>','<b>Final Book Value</b>','<b>Reclass?</b>','<b>Note</b>'],
            line_color='darkslategray',
            fill_color=headerColor,
            align=['left','center'],
            font=dict(color='white', size=12)
          ),
          cells=dict(
            values=[dfA.Asset,dfA.Asset_Description,dfA.Value,dfA.Depreciated_Amount,dfA.Amount_To_Depreciate,dfA.End_Book_Value,dfA.Reclass ,dfA.Note],
            line_color = 'darkslategray',
            fill_color = [[rowOddColor,rowEvenColor,rowOddColor, rowEvenColor,rowOddColor,rowEvenColor]*200],
            align = ['left', 'center'],
            font = dict(color = 'darkslategray', size = 11)
                ))
          ])

        fig1.show()


def save_files(dfA, output):


    dfA = dfA.round(2)

    writer = pd.ExcelWriter(output,
                            engine='xlsxwriter',
                            datetime_format='yyyymmdd',
                            date_format='yyyymmdd')

    dfA.to_excel(writer, index = False, sheet_name = ('Assets'))


    workbook  = writer.book
    worksheet = writer.sheets['Assets']
    worksheet.set_column('B:AV', 40)
    writer.save()


if __name__ == '__main__':

    args = parse_args()

    assets = args.Fixed_Assets_File
    country = args.Country
    output = args.Output_File_Name
    graphs = args.Graphs
    tables = args.Tables
    year = int(args.Current_Year)

    assets_country = country_select(assets)
    assets_calc = time_calc(assets_country, year)
    tables_make(assets_calc)
    graphs_make(assets_calc)
    save_files(assets_calc, output)
    print('Depreciated assets were saved!')


    
