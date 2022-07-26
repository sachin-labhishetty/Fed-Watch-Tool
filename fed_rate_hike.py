from datetime import datetime
from openpyxl import load_workbook
import xlwings as xw
import pandas as pd
import numpy as np
import math
import blpapi
import pdblp

def generate_tickers(meeting_dates):
    """
    returns a list of tickers following the bbg 
    generic tickers convention US0ACR MMMYYYY Index.
    
    Inputs:
    --------
    
    meeting_dates: list of meeting dates in string formats
    """
    # ticker for current overnight implied rate 
    dict_tickers = {'US0ACR Index':datetime.today().date()}
    for date in pd.to_datetime(meeting_dates):
        dict_tickers['US0AFR ' + date.strftime('%b%Y').upper() + ' Index'] = date.date()
    
    return dict_tickers 

def pull_data(tickers, flds, target_rate):
    """
    returns a dataframe of the meeting dates and
    the implied rates for that date
    
    Inputs:
    --------
    
    tickers: dictionary of tickers referncing to the 
    implied rates for different meeting months
    
    flds: list of fields for which you want to pull the data
    
    target_rate: ticker FDTR Index from bbg, the current target rate of the country
    """
    # pulls date using the bloomberg's API
    con = pdblp.BCon()
    con.start()
    
    df_implied_rates = con.ref(list(tickers.keys()), flds)
    df_implied_rates['Dates'] = df_implied_rates['ticker'].map(tickers)
    df_implied_rates = df_implied_rates.sort_values(by='Dates').reset_index().drop(columns=['index','Dates'])
    t_rate = con.ref(target_rate, flds)
    return df_implied_rates, t_rate.iloc[:,2].values[0]

def format_datatable(tickers, df):
    """
    returns a formatted datatable
    
    Inputs:
    --------
    tickers: dictionary of tickers referncing to the 
    implied rates for different meeting months
    
    df: dataframe containing only the implied rates
    calculated from bloomberg
    """
    df['meeting dates'] = df['ticker'].map(tickers)
    df.drop(columns=['ticker','field'], inplace=True)
    
    return df

def calculate_bbg_wirp(df, arm):
    """
    returns the table that is being displayed
    in the bbg WIRP screen
    
    Inputs:
    --------
    df: dataframe containg the implied rates
    pulled from bbg
    
    arm: assumed rate move, usually is 0.25
    """
    df['implied rate change'] = df['value'].diff()
    df['#Hike/Cut'] = (df['value'] - df['value'].iloc[0])/arm
    df['%Hike/Cut'] = 100*(df['#Hike/Cut'].diff(1))
    df.rename(columns = {'value':'implied rate'}, inplace=True)
    
    return df[['meeting dates', '#Hike/Cut', '%Hike/Cut', 'implied rate change', 'implied rate']]

def calc_hike_cut_p(df, arm):
    """
    returns the probability of a hike or a cut 
    for every meeting date
    
    Inputs
    -------
    df: dataframe with the WIRP table information
    
    arm: assumed rate move, usually is 0.25
    """
    df = df.iloc[1:].copy(deep=True)
    # +1 indicates hike & -1 indicates cut
    df['change_sign'] = np.where(df['%Hike/Cut'] > 0, 1, -1)
    
    # higher change in magnitude, could be rate cut or a hike
    df['p_change_h'] = abs(df['%Hike/Cut']/100) - (abs(df['%Hike/Cut'])/100).apply(np.floor) 
    df['amt_change_h'] = df['change_sign']*(abs(df['%Hike/Cut'])/100).apply(np.ceil)*arm
    
    # lower change in magnitude, could be rate cut or a hike
    df['p_change_l'] = 1 - df['p_change_h']
    df['amt_change_l'] = df['amt_change_h'] - df['change_sign']*arm
    
    return df

def probability_table(target_rate, df):
    """
    returns a conditional probability table similar
    to the one displayed in CME website.
    
    Inputs:
    --------
    target_rate: the current fed's (CB's) target rate. Pulled through the bbg API
    
    df: (params) dataframe which contains the probabilities of rate hikes & cut every meeting
    """
    h_delta_probabilities = list(df['p_change_h'])
    rates = {target_rate:1}
    ls_prob = []

    for i in range(len(h_delta_probabilities)):
        h_delta_probability = h_delta_probabilities[i]
        new_rates = {}
        for rate, prior_probability in rates.items():
            rate_with_h_delta = rate + df['amt_change_h'].iloc[i]
            rate_with_l_delta = rate + df['amt_change_l'].iloc[i]

            p_with_h_delta = rates[rate]*h_delta_probability
            p_with_l_delta = rates[rate]*(1-h_delta_probability)

            new_rates[rate_with_h_delta] = new_rates.get(rate_with_h_delta,0) + p_with_h_delta
            new_rates[rate_with_l_delta] = new_rates.get(rate_with_l_delta,0) + p_with_l_delta   
        
        rates = new_rates
        ls_prob.append(rates)

    df_probabilities = pd.DataFrame(ls_prob)
    
    return df_probabilities.reindex(sorted(df_probabilities.columns), axis=1)

def format_column_names(df):
    """
    formats the column names to match the cme table
    note** the column names here are in float format
    
    Inputs:
    -------
    df: dataframe for which the columns name have to be formatted
    """
    ls_new_labels = []
    
    for rate_label in df.columns:
        new_label = str(int(100*(rate_label-0.25)))+"-"+str(int(100*rate_label))
        ls_new_labels.append(new_label)
        
    return ls_new_labels

def format_table(df, df_wirp):
    """
    returns a formatted table similar to cme
    and the modified ones
    Inputs:
    --------
    df: probabilities datatable
    
    df_wirp: bbg wirp table
    """
    df.replace(np.nan, 0, inplace=True)
    new_labels = format_column_names(df)
    df.columns = new_labels
    df = df.mul(100)
    df = df.round(3)
    df.insert(0, 'Meeting Date', df_wirp['meeting dates'].iloc[1:].values)
    
    # modified table
    df_mod = pd.DataFrame()
    start = 1
    for column in df.columns[1:]:
        df_mod[column]=df.iloc[:,start:].sum(axis=1)
        start += 1
    df_mod.insert(0, 'Meeting Date', df['Meeting Date'])
    df_mod = df_mod.round(2)
    
    return df, df_mod

def write_dataframes_to_excel(sheet, top, left, df):
    """
    Funtion to write dataframes starting from 
    the top left cell.
    """
    position = (top, left)
    top_left_cell = sheet.range(position)
    top_left_cell.expand('table').clear()
    top_left_cell.value = df.copy()

    #Format the data, column width, height
    all_data_range = top_left_cell.expand('table')
    all_data_range.column_width = 10.76
    
    #Format the borders
    data_top_left_cell = (top+1, left+1)
    pos_top_left_data = sheet.range(data_top_left_cell)
    data_ex_headers = pos_top_left_data.expand('table')
    for border_id in range(7,13):
        data_ex_headers.api.Borders(border_id).Weight = 2

    #Format the headers
    header_range = top_left_cell.expand('right')
    header_range.color = (112,173,71)
    header_range.api.Font.Color = 0xFFFFFF
    header_range.api.Font.Bold = True

    top_left_cell.expand('down').api.Font.Bold = True
    top_left_cell.expand("right").api.Borders.Weight = 2
    top_left_cell.expand("down").api.Borders.Weight = 2
    bottom_right_cell = top_left_cell.expand('table').last_cell
    right = bottom_right_cell.column

    return right

@xw.func
def main():
#if __name__ == '__main__':
    """
    writes the wirp, cme, mod_cme dataframes
    into the excel "BBG_Data".
    """
    filename = "Z:\Enterprise Shares\Risk Management\Market_Risk\Fifth Third Securities\Reports\Daily Risk Report 3\Troubleshooting Tools\Fed Hike Probabilities_WIRP_CME.xlsx"
    meeting_dates = ['7/27/2022', '9/21/2022','11/2/2022', '12/14/2022', '2/1/2023', '3/15/2023', '5/3/2023',
                '6/14/2023', '7/26/2023', '9/20/2023', '11/1/2023', '12/13/2023', '1/31/2024']
    flds = ['PX_LAST']

    tickers = generate_tickers(meeting_dates[:12])
    print(tickers)
    df_data, target_rate = pull_data(tickers, flds, "FDTR Index")
    df_data = format_datatable(tickers, df_data)
    df_wirp = calculate_bbg_wirp(df_data, arm=0.25)
    df_params = calc_hike_cut_p(df_wirp, arm=0.25)
    df_prob = probability_table(target_rate=target_rate, df=df_params)
    df_cme, df_mod = format_table(df_prob, df_wirp)

    
    # Writing final outputs to the excel file
    wk_bk = xw.Book.caller()
    sheet = wk_bk.sheets['BBG_Data']
    sheet.range((1,1),(5000,15000)).clear() 
    
    #Writing the data frames into the specified (sheet, top, left, dataframe)
    rng = sheet.range((1,2))
    rng.expand('right').value = "The data is refreshed as of " + str(datetime.today())
    #1. wirp
    write_dataframes_to_excel(sheet, 3, 2, df_wirp)
    #2. cme
    write_dataframes_to_excel(sheet, 33, 2, df_cme)
    #3. mod
    write_dataframes_to_excel(sheet, 63, 2, df_mod)
