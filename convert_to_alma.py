import sys
import re
import pandas as pd
import xlrd
from pandas import ExcelWriter
from pandas import ExcelFile


'''
['publication_title', TITLE
'print_identifier',ISSN(2)
'online_identifier',ISSN
'date_first_issue_online' FROM_YEAR,FROM_MONTH,FROM_DAY
'num_first_issue_online', FROM_ISSUE
'num_first_vol_online', FROM_VOLUME
'date_last_issue_online', TO_YEAR, TO_MONTH,TO_DAY
'num_last_issue_online', TO_ISSUE
'num_last_vol_online', TO_VOLUME
'title_url', URL
'first_author',
'title_id',
'embargo_info', INTERNAL_DESCRIPTION
'coverage_depth', COVERAGE_STATEMENT
'coverage_note', PUBLIC_NOTES (?)
'publisher_name'],
)
AVAILABILITY = ACTIVE
'''

def read_kbart(file):
    df = pd.read_excel(file)
    #df = pd.read_html(file, sheetname='Sheet1')
    print (df.columns)
    print (df['publication_title'])
    cols = ['LOCALIZED','ISSN','ISSN2','ISSN3','ISBN','ISBN2','ISBN3','PORTFOLIO_PID','MMS','TITLE','FROM_YEAR','TO_YEAR','FROM_MONTH','TO_MONTH','FROM_DAY','TO_DAY','FROM_VOLUME','TO_VOLUME','FROM_ISSUE','TO_ISSUE','WARNINGS','PUBLICATION_DATE_OPERATOR','PUBLICATION_DATE_YEAR','PUBLICATION_DATE_MONTH', 'GLOBAL_FROM_YEAR', 'GLOBAL_TO_YEAR', 'GLOBAL_FROM_MONTH', 'GLOBAL_TO_MONTH', 'GLOBAL_FROM_DAY', 'GLOBAL_TO_DAY', 'GLOBAL_FROM_VOLUME', 'GLOBAL_TO_VOLUME','GLOBAL_FROM_ISSUE', 'GLOBAL_TO_ISSUE', 'GLOBAL_WARNINGS', 'GLOBAL_PUBLICATION_DATE_OPERATOR','GLOBAL_PUBLICATION_DATE_YEAR','GLOBAL_PUBLICATION_DATE_MONTH','AVAILABILITY','PUBLISHER','PLACE_OF_PUBLICATION','DATE_OF_PUBLICATION','URL','PARSER_PARAMETERS','PROXY_ENABLE','PROXY_SELECTED','PROXY_LEVEL','AUTHOR','ELECTRONIC_MATERIAL_TYPE','OWNERSHIP','GROUP_NAME','AUTHENTICATION_NOTES','PUBLIC_NOTES','INTERNAL_DESCRIPTION','COVERAGE_STATEMENT','ACTIVATION_DATE','EXPECTED_ACTIVATION_DATE','LICENSE','LICENSE_NAME','PDA','NOTES' ]
    print (cols)
    print(len(cols))
    df_out = pd.DataFrame(index=df.index)
    for col in cols:
      if col == 'ISSN':
         df_out.loc[:,col] = df['online_identifier']
      elif col == 'ISSN2':
          df_out.loc[:,col] = df['print_identifier']
      elif col == 'TITLE':
          df_out.loc[:,col] = df['publication_title']
      elif col == 'FROM_YEAR':
          df_out.loc[:, col] = df['date_first_issue_online'].dt.year
      elif col == 'FROM_MONTH':
          df_out.loc[:, col] = df['date_first_issue_online'].dt.month
      elif col == 'FROM_DAY':
          df_out.loc[:, col] = df['date_first_issue_online'].dt.day
      elif col == 'TO_YEAR':
          df_out.loc[:, col] = df['date_last_issue_online'].dt.year
      elif col == 'TO_MONTH':
          df_out.loc[:, col] = df['date_last_issue_online'].dt.month
      elif col == 'TO_DAY':
          df_out.loc[:, col] = df['date_last_issue_online'].dt.day
      elif col == 'FROM_VOL':
          df_out.loc[:, col] = df['num_first_vol_online']
      elif col == 'FROM_ISSUE':
          df_out.loc[:, col] = df['num_first_issue_online']
      elif col == 'TO_VOL':
          df_out.loc[:, col] = df['num_last_vol_online']
      elif col == 'TO_ISSUE':
          df_out.loc[:, col] = df['num_last_issue_online']
      elif col == 'AVAILABILITY':
        df_out.loc[:, col] = 'ACTIVE'
      elif col == 'TITLE':
          df_out.loc[:, col] = df['publication_title']
      elif col == 'NOTES':
          df_out.loc[:, col] = df['embargo_info']
      elif col == 'INTERNAL_DESCRIPTION':
          df_out.loc[:, col] = df['coverage_note']
      elif col == 'PUBLIC_NOTES':
          df_out.loc[:, col] = df['coverage_depth']
      else:
          df_out.loc[:,col] = ''
    df_out = df_out.fillna('')

    df_out.rename(columns={'ISSN2': 'ISSN'}, inplace=True)
    df_out.rename(columns={'ISSN3': 'ISSN'}, inplace=True)
    df_out.rename(columns={'ISBN2': 'ISBN'}, inplace=True)
    df_out.rename(columns={'ISBN3': 'ISBN'}, inplace=True)

    writer = pd.ExcelWriter('alma_import.xlsx', engine='xlsxwriter')
    df_out.to_excel(writer,sheet_name='Sheet1',index=False)
    writer.save()


kbart_file = sys.argv[1]
read_kbart(kbart_file)
