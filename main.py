import numpy as np
import pandas as pd
import time


""" filename and sheet locations"""
filename_time = "." + time.strftime("%Y%m%d-%H%M%S")
filenames = {'int': './internal.xlsx', 'pb': './pb.xlsx', 'admin': './admin_falsebreak.xlsx'}
sheets = {'int': 'Internal', 'pb': 'PB', 'admin': 'Admin'}

""" setup internal dataframe """
df_int = pd.read_excel(
    filenames['int'],
    sheets['int'],
    header=0,
    usecols=[0, 2, 4, 5],
    names=['Type', 'ISIN', 'Quantity', 'Price']
)
df_int = df_int.sort_values(by=['Type', 'ISIN']).reset_index()
df_int['Price'] = df_int['Price'] * 100
df_int_agg = df_int.groupby(['Type', 'ISIN']).agg({
    'Quantity': 'sum',
    'Price': 'mean'
})

""" setup pb dataframe """
df_pb = pd.read_excel(
    filenames['pb'],
    sheets['pb'],
    header=3,
    usecols=[4, 6, 10, 11],
    names=['Type', 'ISIN', 'Quantity', 'Price']
)
df_pb = df_pb.sort_values(by=['Type', 'ISIN']).reset_index()
df_pb_agg = df_pb.groupby(['Type', 'ISIN']).agg({
    'Quantity': 'sum',
    'Price': 'mean'
})

""" setup admin dataframe """
df_admin = pd.read_excel(
    filenames['admin'],
    sheets['admin'],
    header=0,
    usecols=[3, 5, 6, 7],
    names=['ISIN', 'Name', 'Price', 'Quantity']
)
conditions_admin = [
    ((df_admin['Name'].str.startswith('STOCK LOAN FEE', na=False)) | (df_admin['Name'].str.startswith('TICKET CHARGES', na=False))),
    ((df_admin['Name'].str.startswith('EURO', na=False)) | (df_admin['Name'].str.startswith('U S DOLLARS', na=False))),
    (df_admin['Name'].str.startswith('FFX', na=False)),
    ((df_admin['Name'].str.startswith('RP', na=False)) | (df_admin['Name'].str.startswith('RV', na=False))),
    ((df_admin['Name'].str.startswith('BUY', na=False)) | (df_admin['Name'].str.startswith('SELL', na=False))),
    (((~df_admin['Name'].str.startswith('RP', na=False)) | (~df_admin['Name'].str.startswith('RV', na=False))) & (~df_admin['ISIN'].isnull())),  # Catch-all Bond
    df_admin['ISIN'].isnull()  # Catch-all "Other"
]
values_admin = ['Expenses', 'Cash', 'FX Forward', 'Repurchase Agreement', 'CDS', 'Bond', 'Other']
df_admin['Type'] = np.select(conditions_admin, values_admin)
df_admin = df_admin.sort_values(by=['Type', 'ISIN']).reset_index()
df_admin_agg = df_admin.groupby(['Type', 'ISIN']).agg({
    'Quantity': 'sum',
    'Price': 'mean'
})

""" setup comparison tool """
diff_cols = ['Quantity', 'Price']
df_agg_diffs_int_pb = df_int_agg[diff_cols] - df_pb_agg[diff_cols]
df_agg_diffs_int_pb = df_agg_diffs_int_pb.loc[df_agg_diffs_int_pb['Quantity'].fillna(0) != 0]
df_agg_diffs_int_admin = df_int_agg[diff_cols] - df_admin_agg[diff_cols]
df_agg_diffs_int_admin = df_agg_diffs_int_admin.loc[df_agg_diffs_int_admin['Quantity'].fillna(0) != 0]

""" run comparison and write csv """
df_agg_results = pd.concat({'Int vs. PB': df_agg_diffs_int_pb, 'Int vs. Admin': df_agg_diffs_int_admin}, names=['Source'])
df_agg_results.to_csv(r'csv_results' + filename_time + '.csv')

""" print results """
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)
print('\n', '\n', '  Master Break List  ', '\n', df_agg_results)
