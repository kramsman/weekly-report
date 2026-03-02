"""Orchestrate local creation of all Sincere summary Excel reports."""

import os
import shutil
from pathlib import Path
from pathlib import Path

import pandas as pd
from uvbekutils import pyautobek
from uvbekutils import pyautobek
from uvbekutils import select_file
from uvbekutils import select_file

from weekly_report.utils.read_sincere_request_file import read_sincere_request_file
from weekly_report.constants import FACTORY_FILTER_STRING
from weekly_report.constants import OUTPUT_DIR_ADMIN
from weekly_report.constants import OUTPUT_DIR_ADMIN
from weekly_report.constants import OUTPUT_DIR_REPORTS
from weekly_report.constants import OUTPUT_DIR_REPORTS
from weekly_report.constants import SINCERE_DOWNLOAD_DIR
from weekly_report.constants import SINCERE_DOWNLOAD_DIR
from weekly_report.create_admin_report import create_admin_report
from weekly_report.create_room_reports import create_room_reports


def create_report_files() -> None:
    """Produce all local Excel report files from user-selected Sincere data.

    Prompts for weekly vs monthly format, input CSV files, and output
    directories. Applies factory and org filters, derives sort dictionaries,
    then delegates to create_admin_report() and create_room_reports().
    """

    report_by = pyautobek.confirm("\n\n[W]eekly or [M]onthly report?", "Date Format", ["W", "M", 'Exit'])
    report_by = report_by.upper()
    if report_by == 'M':
        report_var = "month"
    elif report_by == 'W':
        report_var = "data_week_ending"
    else:
        exit()

    if True:  # False out prompt for debugging
        input_file = select_file("Pick a Sincere Requests File",
                                    SINCERE_DOWNLOAD_DIR,
                                 'all-parent-campaigns-requests*.csv',
                                 ["Select", "Cancel"],
                                 "file",
                                 "Select Sincere address export file 'all-parent-campaign-REQUESTS-yyyy-mm-dd.csv'"
                                 )
    else:
        input_file = "/Users/Denise/Downloads/all-parent-campaigns-requests-2026-02-16.csv"

    if input_file is None:
        exit()
    # input_file = Path('/Users/Denise/Downloads/all-parent-campaigns-requests-2025-08-01.csv')

    if True:  # False out for debugging
        factory_csv = select_file("Pick File",
                              SINCERE_DOWNLOAD_DIR,
                              'parent-campaign-address-counts*.csv',
                              ["Select", "Cancel"],
                              'file',
                              f"Pick a parent-campaign file to summarize (eg "
                                f"'parent-campaign-address-counts-2023-08-03.csv'."
                                f"\n\nCreate via ROV > Reports > New Report > Parent Campaign Address COUNTS",
                              )
    else:
        factory_csv = "/Users/Denise/Downloads/parent-campaign-address-counts-2026-02-16.csv"

    if factory_csv is None:
        exit()
    # factory_csv = Path('/Users/Denise/Downloads/parent-campaign-address-counts-2025-08-01.csv')

    # Create dataframe of excel sheet data
    sincere_df = read_sincere_request_file(input_file)
    sincere_data_file_name = os.path.basename(input_file)

    # only take factories from this year - many old and some new are locked
    sincere_df = sincere_df.loc[sincere_df['factory_name']
        .apply(lambda x: (FACTORY_FILTER_STRING in x.lower()))]
    # remove any unwanted requests - training and sample rooms
    sincere_df = sincere_df.loc[sincere_df['factory_name']
        .apply(lambda x: not ('zzz' in x.lower() or 'xxx' in x.lower() or 'test' in x.lower()))]
    # remove more unwanted requests - training and sample rooms
    sincere_df = sincere_df.loc[sincere_df['org_name']
        .apply(lambda x: not ('zzz' in x.lower() or 'xxx' in x.lower() or 'training' in x.lower() or 'sample' in
                         x.lower()))]

    # Set Election field to one of two values: General if 'general' found in name else Primary
    sincere_df['election'] = sincere_df['factory_name'].apply(lambda fact: ('General' if 'general' in fact.lower()
                                                                            else 'Primary'))

    # Create a dictionary of the earliest(min) date in a Factory to use for sorting
    factory_dict = sincere_df.groupby(by=['factory_name'], dropna=False)['created_at'].min().apply(
        pd.to_datetime).dt.strftime('%Y%m%d').to_dict()
    factory_dict['_TOTAL'] = 999999

    election_dict = {
    "General": 1,
    "Primary": 2,
    }

    file_date = Path(input_file).stem[-10:]

    master_campaigns = sincere_df['master_campaign'].unique()
    master_campaigns = sorted(master_campaigns)

    ## add dicts to 'level_dicts'; dict will be used for sort values for and dict name matching level name in multiindex
    ## working version to play can be found in sort df using dict.py
    level_dicts = dict()  # name of 'level_dicts' is hardcoded in sort function
    # Key value in level_dicts is matched to level name in index (case-insensitive)
    # level_dicts['Factory'] = factory_dict
    level_dicts['Election'] = election_dict
    level_dicts = {k.lower(): v for k, v in level_dicts.items()}

    # define multiindex_df_sorter function used to sort multiiindex output of groupby by dictionaries
    def multiindex_df_sorter(level: pd.Index, default_level_dict: dict = {}) -> pd.Index:
    # def multiindex_df_sorter(level, default_level_dict={'_TOTAL': 999999}):
        """Map a MultiIndex level to sort-key values using a level-specific dictionary.

        Looks up the level name (case-insensitive) in the outer-scope ``level_dicts``
        dict. If a match is found the corresponding dict maps each index item to a
        sort key; otherwise items sort as-is (empty dict fallback).

        Args:
            level: One level of a pandas MultiIndex to be sorted.
            default_level_dict: Fallback mapping when the level name is not found
                in ``level_dicts``. Defaults to {}.

        Returns:
            The level mapped to sort-key values.
        """
        level_dict = level_dicts.get(level.name.lower(), default_level_dict)
        mapped_index = level.map(lambda index_item: level_dict.get(index_item, index_item))
        return mapped_index

    ### Create admin report
    admin_rpt_filename = f"ROVWide Sincere Summary {file_date}-{report_by}.xlsx"
    # create_admin_report(sincere_df, sincere_data_file_name, report_by, OUTPUT_DIR_ADMIN, admin_rpt_filename,
    #                     factory_csv)
    create_admin_report(sincere_df, sincere_data_file_name, report_by, OUTPUT_DIR_ADMIN, admin_rpt_filename,
                        factory_csv, multiindex_df_sorter)


    # Base of report file; input file name is used to create a sub dir under this
    output_dir = os.path.expanduser(OUTPUT_DIR_REPORTS)
    output_dir_w_file = os.path.join(output_dir, Path(input_file).stem + "-" + report_by)
    # CHECK and create dir if not exists
    # bad_path_create(output_dir_w_file)
    # if not os.path.isdir(output_dir_w_file):
    #     print(f"Adding Directory {output_dir_w_file}\n")
    #     os.makedirs(output_dir_w_file)
    if os.path.isdir(output_dir_w_file):
        shutil.rmtree(output_dir_w_file)
        print(f"Removing directory {output_dir_w_file}\n")
    os.makedirs(output_dir_w_file)
    # Run reports at the room/org level
    create_room_reports(sincere_df, sincere_data_file_name, file_date, report_by, output_dir_w_file)

    print("\nDone with all orgs.")
    pyautobek.alert(f"Org reports produced. In:"
               f"\n\n{OUTPUT_DIR_REPORTS}"
               f"\n\n\nAdmin reports produced. In:"
               f"\n\n{OUTPUT_DIR_ADMIN}","Done!",
                    )
