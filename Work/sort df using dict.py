""" test sorting a multiindex df by dict using agg values from df """

from pathlib import Path

import numpy as np
import pandas as pd
from factory_and_campaign_subtotals import factory_and_campaign_subtotals


##
request_df = pd.read_csv("/Users/Denise/Downloads/all-parent-campaigns-requests-2024-08-05.csv")
request_df.rename(columns={'factory_name': 'Factory',}, inplace=True)

request_df = request_df[(request_df['Factory'].notna())]
request_df = request_df[(~request_df['Factory'].str.lower().str.contains("zz"))]
request_df = request_df[(request_df['Factory'].str.contains("-2024") & (request_df['Factory'].str.lower().str.contains("ga") | request_df[
    'Factory'].str.lower().str.contains("az")))]

# Create a dictionary of the earliest(min) date in a Factory to use for sorting
Factory_dict = request_df.groupby(by=['Factory'], dropna=False)['created_at'].min().apply(pd.to_datetime).to_dict()
##

factory_csv = Path("~/Downloads/parent-campaign-address-counts-2024-08-01.csv").expanduser()
FACTORY_FILTER_STRING = '6-2024'

# factory_tots, factory_pull_date = factory_and_campaign_subtotals(factory_csv, FACTORY_FILTER_STRING,
#                                                                  break_fields="[('Factory', True), ]")

df = pd.read_csv(factory_csv, usecols=['Factory', 'Name', 'Total Addresses' ])
df = df[(df['Factory'].notna() & ~df['Name'].str.lower().str.contains("test"))]
df = df[(~df['Factory'].str.lower().str.contains("zz"))]
df = df[(~df['Name'].str.lower().str.contains("xx"))]
df['Name'] = df['Name'].map(lambda s: s.split('-')[0] + '-' + s.split('-')[1])
df = df[(df['Name'].str.lower().str.contains("d"))]


df = df[(df['Factory'].str.contains("-2024") & (df['Factory'].str.lower().str.contains("ga") | df[
    'Factory'].str.lower().str.contains("az")))]
# df = df[(df['Factory'].str.contains("-2024"))]

def set_election_type(name):
    if 'general' in name.lower():
        election_type = 'g12eneral'
    else:
        election_type = 'P21rimary'
    return election_type

df['Election'] = df['Factory'].apply(lambda fact: set_election_type(fact))

df.rename(columns={'Total Addresses': 'Total',}, inplace=True)

# summed = df.groupby(by=['Election', 'Factory', 'Name'], dropna=False).sum()
summed = df[['Election', 'Total']].groupby(by=['Election'], dropna=False).sum()
# summed = df.groupby(by=['Factory', 'Election', 'Name'], dropna=False).sum()
# summed = df.groupby(by=['Factory',], dropna=False).sum()

if False:
    sincere_data_file_name = factory_csv.stem
    wks_file = factory_csv.with_suffix('.xlsx')
    summed.to_excel(wks_file)

# Factory = {
# "GA Primary 1-2024": 2,
# "GA General GOTV 6-2024": 1,
# }


Name = {
"GA-Clayton-General GOTV 6-2024 **mail Sept 15 to 30 (revised)**": 1,
"GA-Worth-General GOTV 6-2024 **mail Sept 15 to 30 (revised)**": 2
}

## add dicts to dict lookup
level_dicts = dict()
# Key value in level_dicts is matched to level name in index (case-insensitive)
level_dicts['Factory'] = Factory_dict
# level_dicts['Election'] = lambda s: s[2:3]
level_dicts['Election'] = lambda s: s.lower()
# level_dicts['Election'] = (dict())
# level_dicts['name'] = Name
level_dicts = {k.lower(): v for k, v in level_dicts.items()}
##

# def multiindex_df_sorter(level, keep_index=lambda x: x):
def multiindex_df_sorter(level, default_level_dict=dict(), **kwargs):
    """ function for sorting a dataframe's multiindex using a dictionary for each level.  if name of index
    (case-insensitive) field matches key in dict hardcoded as 'level_dicts' then dictionary is used,
    otherwise index level is left as is via use of an empty dictionary.  """
    level_dict = level_dicts.get(level.name.lower(), default_level_dict)

    use_expression_1 = False

    print(f"{use_expression_1=}")
    print(f"'isinstance(level_dict, dict)' for {level.name}= {isinstance(level_dict, dict)}")
    print(f"'level_dict' for {level.name}= {level_dict}")
    print()

    # ret = level.map( lambda s: d.get(s, np.NaN))
    # mapped_index = level.map(lambda index_item: level_dict.get(index_item, index_item))
    if use_expression_1:
        mapped_index = level.map(level_dict if not isinstance(level_dict, dict)
                                 else lambda index_item: level_dict.get(index_item, index_item))
    else:
        mapped_index = level.map(lambda index_item: level_dict.get(index_item, index_item) if isinstance(level_dict,
                                                                                                         dict) else level_dict)
    return mapped_index

# a_sorted_df = summed.sort_index(level=['Factory', 'Election', 'Name'], key=multiindex_df_sorter)
# a_sorted_df = summed.sort_index(level=['Election', 'Factory', 'Name'], key=multiindex_df_sorter)
a_sorted_df = summed.sort_index(level=['Election'], key=multiindex_df_sorter)
print("\nsorted df:")
pd.set_option('display.max_rows', None)
print(a_sorted_df)

a=1

### This works but not for passed functions, only dicts
# def multiindex_df_sorter(level, default_level_dict=dict(), **kwargs):
#     level_dict = level_dicts.get(level.name.lower(), default_level_dict)
#     # mapped_index = level.map(lambda index_item: level_dict.get(index_item, index_item))
#     return mapped_index

#def multiindex_df_sorter(level, default=lambda x: x):
    # return {
    #   'one': lambda s: s.str.lower(),
    #   'two': np.abs,
    # }.get(level.name, default)(level)

# def multiindex_df_sorter(level, default=lambda x: x):
#     ret = {
#       'Factory': lambda s: order.get(s, 0),
#     }.get(level.name, default)
#     return ret(level)

# def multiindex_df_sorter(level, default=lambda x: x):
#     ret = level.map( lambda s: order.get(s, np.NaN))
#     # ret = lambda s: order.get(s, 0)(level)

# order = {'size': ['small', 'medium', 'tall']}
#
# def multiindex_df_sorter(s):
#     if s.name in order:
#         return s.map({k:v for v,k in enumerate(order[s.name])})
#     return s
#
# out = df_0.sort_index(level=["number", "size", "letter"], key=multiindex_df_sorter)
