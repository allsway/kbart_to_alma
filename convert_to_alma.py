#!/usr/bin/env python3.5

import sys
import re
import pandas as pd
import numpy as np
import xlrd
import uuid
from pandas import ExcelWriter
from pandas import ExcelFile
import dateparser
import datetime

# Calculate the date columns for each date
def get_year(df, column):
    date_triple = {'year_list': [],'month_list': [],'day_list': []}
    df[column] = df[column].astype('str')
    temp_df = df[column].map(lambda x: x.split('.')[0])
    print (temp_df)
    #print (df[column])
    temp_df = pd.to_datetime(temp_df)
    print (temp_df)
    #  datetime.datetime.fromtimestamp(fl)
    i = 0
    for row in temp_df:
        original =  re.sub("[^0-9]", "", str(df[column][i]))
        i += 1
        if (not pd.isnull(row)):
            if len(original) > 7:
                date_triple['month_list'].append(row.month)
                date_triple['day_list'].append(row.day)
                date_triple['year_list'].append(row.year)
            elif len(original) > 5:
                date_triple['month_list'].append(row.month)
                date_triple['year_list'].append(row.year)
                date_triple['day_list'].append('')
            else:
                date_triple['month_list'].append('')
                date_triple['year_list'].append(row.year)
                date_triple['day_list'].append('')
        else:
            date_triple['year_list'].append('')
            date_triple['month_list'].append('')
            date_triple['day_list'].append('')

    return date_triple

# Parses kbart file headers
def read_kbart(file):
    output_file = 'alma_import_' + str(uuid.uuid4()) + '.xlsx'
    df = pd.read_excel(file)# parse_dates??
    # Get the column names for the Alma expected import columns
    cols = ['LOCALIZED','ISSN','ISSN2','ISSN3','ISBN','ISBN2','ISBN3','PORTFOLIO_PID','MMS','TITLE','FROM_YEAR','TO_YEAR','FROM_MONTH','TO_MONTH','FROM_DAY','TO_DAY','FROM_VOLUME','TO_VOLUME','FROM_ISSUE','TO_ISSUE','WARNINGS','PUBLICATION_DATE_OPERATOR','PUBLICATION_DATE_YEAR','PUBLICATION_DATE_MONTH', 'GLOBAL_FROM_YEAR', 'GLOBAL_TO_YEAR', 'GLOBAL_FROM_MONTH', 'GLOBAL_TO_MONTH', 'GLOBAL_FROM_DAY', 'GLOBAL_TO_DAY', 'GLOBAL_FROM_VOLUME', 'GLOBAL_TO_VOLUME','GLOBAL_FROM_ISSUE', 'GLOBAL_TO_ISSUE', 'GLOBAL_WARNINGS', 'GLOBAL_PUBLICATION_DATE_OPERATOR','GLOBAL_PUBLICATION_DATE_YEAR','GLOBAL_PUBLICATION_DATE_MONTH','AVAILABILITY','PUBLISHER','PLACE_OF_PUBLICATION','DATE_OF_PUBLICATION','URL','PARSER_PARAMETERS','PROXY_ENABLE','PROXY_SELECTED','PROXY_LEVEL','AUTHOR','ELECTRONIC_MATERIAL_TYPE','OWNERSHIP','GROUP_NAME','AUTHENTICATION_NOTES','PUBLIC_NOTES','INTERNAL_DESCRIPTION','COVERAGE_STATEMENT','ACTIVATION_DATE','EXPECTED_ACTIVATION_DATE','LICENSE','LICENSE_NAME','PDA','NOTES' ]
    df_out = pd.DataFrame(index=df.index)
    if df['date_first_issue_online'].dtype == 'datetime64[ns]':
        from_date_format = True
    else:
        first_issue = get_year(df,'date_first_issue_online')
        from_date_format = False
    if df['date_last_issue_online'].dtype == 'datetime64[ns]':
        last_date_format = True
    else:
        last_date_format = False
        last_issue = get_year(df,'date_last_issue_online')
    for col in cols:
        if col == 'ISSN':
           df_out.loc[:,col] = df['online_identifier']
        elif col == 'ISSN2':
            df_out.loc[:,col] = df['print_identifier']
        elif col == 'TITLE':
            df_out.loc[:,col] = df['publication_title']
        elif col == 'FROM_YEAR':
            if from_date_format:
                df_out.loc[:, col] = df['date_first_issue_online'].dt.year
            else:
               df_out.loc[:, col] = first_issue['year_list']
        elif col == 'FROM_MONTH':
            if from_date_format:
                df_out.loc[:, col] = df['date_first_issue_online'].dt.month
            else:
                df_out.loc[:,col] = first_issue['month_list']
        elif col == 'FROM_DAY':
            if from_date_format:
                df_out.loc[:, col] = df['date_first_issue_online'].dt.day
            else:
                df_out.loc[:,col] = first_issue['day_list']
        elif col == 'TO_YEAR':
            if last_date_format:
                df_out.loc[:, col] = df['date_last_issue_online'].dt.year
            else:
                df_out.loc[:,col] = last_issue['year_list']
        elif col == 'TO_MONTH':
            if last_date_format:
                df_out.loc[:, col] = df['date_last_issue_online'].dt.month
            else:
                df_out.loc[:, col] = last_issue['month_list']
        elif col == 'TO_DAY':
            if last_date_format:
                df_out.loc[:, col] = df['date_last_issue_online'].dt.day
            else:
                df_out.loc[:, col] = last_issue['day_list']
        elif col == 'FROM_VOLUME':
            df_out.loc[:, col] = df['num_first_vol_online']
        elif col == 'FROM_ISSUE':
            df_out.loc[:, col] = df['num_first_issue_online']
        elif col == 'TO_VOLUME':
            df_out.loc[:, col] = df['num_last_vol_online']
        elif col == 'TO_ISSUE':
            df_out.loc[:, col] = df['num_last_issue_online']
        elif col == 'AVAILABILITY':
            df_out.loc[:, col] = 'ACTIVE'
        elif col == 'URL':
      	    df_out.loc[:,col] = df['title_url']
        elif col == 'TITLE':
            df_out.loc[:, col] = df['publication_title']
#        elif col == 'NOTES':
#            df_out.loc[:, col] = df['embargo_info']
        elif col == 'INTERNAL_DESCRIPTION':
            if 'coverage_notes' in df.columns and 'title_change_history' in df.columns:
                df_out.loc[:, col] = df['coverage_notes'] + ', Title change history: ' + df['title_change_history']
            elif 'coverage_notes' in df.columns:
                df_out.loc[:, col] = df['coverage_notes']
            elif 'title_change_history' in df.columns:
                df_out.loc[:, col] = df['title_change_history']
#        elif col == 'PUBLIC_NOTES':
#            df_out.loc[:, col] = df['coverage_depth']
        else:
            df_out.loc[:,col] = ''
    df_out = df_out.fillna('')

    df_out.rename(columns={'ISSN2': 'ISSN'}, inplace=True)
    df_out.rename(columns={'ISSN3': 'ISSN'}, inplace=True)
    df_out.rename(columns={'ISBN2': 'ISBN'}, inplace=True)
    df_out.rename(columns={'ISBN3': 'ISBN'}, inplace=True)

    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    print (output_file)
    df_out.to_excel(writer,sheet_name='Sheet1',index=False)
    writer.save()


kbart_file = sys.argv[1]
read_kbart(kbart_file)
