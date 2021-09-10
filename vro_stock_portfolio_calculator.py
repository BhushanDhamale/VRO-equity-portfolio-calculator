#!/usr/bin/env python
# coding: utf-8

# import libraries
import pandas as pd
from sys import argv
from os import system

# read excel into df
df = pd.read_excel(argv[1], skipfooter = 1)

# get index of row containing 'Stocks Investments'
skiprows_index = df[df['Value Research'] == 'Stocks Investments'].index

# convert the index to list and then to int
# +2 is done to get the correct index
skiprows_index = list(skiprows_index)[0] + 2

# reload the excel into df
df = pd.read_excel(argv[1],
                   skiprows = range(skiprows_index),
                   skipfooter = 1)

# drop irrelevant columns
drop_columns = ['Demat a/c', 'Last price date', '1D Change (INR)',
                '1D Change (%)']
df.drop(drop_columns, axis = 1, inplace = True)

# drop unnamed columns
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

# drop rows of exited companies
df = df[df['Shares'] != 0]

# drop rows of ETFs
df = df[~df['Company name'].str.contains('Fund')]

# recalculate 'Total Return'
# needed because VRO's calculations bug out sometimes
df['Total Return'] = df['Current Value'] - df['Total cost']

# calculate absolute return
df['Absolute Return %'] = df['Total Return'] / df['Total cost'] * 100

# recalculate '% of Total'
tot_curr_val = df['Current Value'].sum()
df['% of Total'] = df['Current Value'] / tot_curr_val * 100

# sort by '% of Total'
df.sort_values(by = '% of Total', ascending = False, inplace = True)

# reset index
df.reset_index(drop = True, inplace = True)

# reorder columns
columns_order = ['Company name', 'Shares', 'Cost per share',
                 'Total cost', 'Last price', 'Current Value',
                 'Total Return', 'Absolute Return %',
                 'Return % pa', '% of Total']
df = df[columns_order]

# round to one decimal
df_percent = df.iloc[:, 7:].round(1)

# round to zero decimals
# convert float to int
df_round = df.iloc[:, 1:7].round(0).astype(int)

# replace column values in df from df_percent and df_round
df.iloc[:, 7:] = df_percent
df.iloc[:, 1:7] = df_round

# calculate total values for portfolio
curr_value = df['Current Value'].sum().astype(int)
tot_cost = df['Total cost'].sum().astype(int)
tot_return = df['Total Return'].sum().astype(int)
abs_return = ((tot_return / tot_cost) * 100).round(0)
prcnt_total = df['% of Total'].sum().astype(int)

# round to one decimal
df_percent = df.iloc[:, 7:].round(1)

# round to zero decimals
# convert float to int
df_round = df.iloc[:, 1:7].round(0).astype(int)

# replace column values in df from df_percent and df_round
df.iloc[:, 7:] = df_percent
df.iloc[:, 1:7] = df_round

# append 'Total' row to df
col_values = ['Total','NA','NA',tot_cost,'NA',curr_value,tot_return,
              abs_return,'NA',prcnt_total]
df_total = pd.DataFrame([col_values], columns=columns_order)
df = pd.concat([df,df_total], ignore_index=True)

# save without index to excel file
outfile = 'stock_portfolio.xlsx'
df.to_excel(outfile, index = False)
print("Output written to {}".format(outfile))

# open excel file
system('libreoffice --calc stock_portfolio.xlsx')
