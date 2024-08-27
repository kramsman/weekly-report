""" it appears file factory_and_campaign_subtotals.py was taken out of git accidentally during an edit of a foreign git.
this is to see if that is the case and if I can get back in """

# from factory_and_campaign_subtotals import factory_and_campaign_subtotals
from bekutils import test_foreign_git_edit
# from bekutils import exit_yes, get_file_name

test_foreign_git_edit()

a=1
#
#
# factory_csv = Path("~/Downloads/parent-campaign-address-counts-2024-08-01.csv").expanduser()
# FACTORY_FILTER_STRING = '6-2024'
#
# # factory_tots, factory_pull_date = factory_and_campaign_subtotals(factory_csv, FACTORY_FILTER_STRING,
# #                                                                  break_fields="[('Factory', True), ]")
#
# df = pd.read_csv(factory_csv, usecols=['Factory', 'Name', 'Total Addresses' ])
# df = df[(df['Factory'].notna() & ~df['Name'].str.lower().str.contains("test"))]
# df = df[(~df['Factory'].str.lower().str.contains("zz"))]
# df = df[(~df['Name'].str.lower().str.contains("xx"))]
# df['Name'] = df['Name'].map(lambda s: s.split('-')[0] + '-' + s.split('-')[1])
# df = df[(df['Name'].str.lower().str.contains("d"))]
#
#
# df = df[(df['Factory'].str.contains("-2024") & (df['Factory'].str.lower().str.contains("ga") | df[
#     'Factory'].str.lower().str.contains("az")))]
# # df = df[(df['Factory'].str.contains("-2024"))]
#
# factory_tots, factory_pull_date = factory_and_campaign_subtotals(factory_csv, FACTORY_FILTER_STRING,
#                                                                  break_fields="[('Election', True), "
#                                                                               "('Factory', False), ]")
