"""Read and clean the Sincere all-parent-campaigns-requests CSV export."""

import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd

from weekly_report.utils.utils import check_df_headers


def read_sincere_request_file(input_file: str | Path) -> pd.DataFrame:
    """Read a Sincere requests CSV and return a cleaned, enriched DataFrame.

    Validates column headers, normalises org and team name strings (replaces
    /, ', and , characters), computes week-ending Sunday dates and month
    period columns, fills missing team names with ' No Team', and derives
    master_campaign and shortened parent_campaign_name fields.

    Args:
        input_file: Path to the Sincere all-parent-campaigns-requests CSV.

    Returns:
        Cleaned request data with additional derived columns including
        data_week_ending, month, master_campaign, and election.
    """
    request_df = pd.read_csv(input_file, na_filter=False)

    # Check heading fields in Sincere request file to ensure fields didn't move/change
    check_df_headers(df=request_df,
                     vals=['parent_campaign_id', 'parent_campaign_name', 'request_id', 'created_at', 'writer_name',
                           'writer_email', 'addresses_count', 'org_name', 'org_id', 'team_name', 'factory_name',
                           'factory_id'])

    # replace '/' in some input fields (like org name) which cause issues when used as directories
    request_df['org_name'] = request_df['org_name'].str.replace("/", "-").str.replace("'", "-").str.replace(",", "-")
    request_df['team_name'] = request_df['team_name'].str.replace("/", "-").str.replace("'", "-").str.replace(",", "-")

    # todo: Sort orgs somewhere.  Here?

    # create a date field as a datetime object
    request_df['date2'] = pd.to_datetime(request_df['created_at'])

    # 'daysoffset' will contain the weekday 'day of week' as integers so we can step back to Monday
    request_df['daysoffset'] = request_df['date2'].apply(lambda x: x.weekday())
    # TODO Explain this timedelta operation
    request_df['data_week_ending'] = request_df.apply(lambda x: x['date2'] - dt.timedelta(days=x['daysoffset'] - 6), axis=1).dt.date
    request_df['month'] = pd.to_datetime(request_df['date2']).dt.to_period('M')

    # request_df.fillna(" No Team", inplace = True) # note space in front for sorting # change to only replace one col
    request_df['team_name'] = np.where(request_df['team_name'] == "", " No Team", request_df['team_name'])
    request_df['team_name'] = request_df['team_name'].str.replace(":", "-", regex=False).str.replace("\\", "-", regex=False).str.replace("/", "-", regex=False)

    request_df = request_df[~request_df['org_name'].isin(['ROV Test Silo', 'ROV Training Silo', 'ROV Sample Silo'])]  # select records for one Org

    request_df['master_campaign'] = np.where(request_df['factory_name'].isnull() == False, request_df['factory_name'],
                                     request_df['parent_campaign_name'].apply(lambda x: x[0:2]))

    # if factory name exists, trim parent campaign name  to just state/county
    request_df['parent_campaign_name'] = np.where(request_df['factory_name'].isnull() == True, request_df['parent_campaign_name'],
                                          request_df['parent_campaign_name'].str.split("-").str[0] + '-' +
                                          request_df['parent_campaign_name'].str.split("-").str[1])
    return request_df
