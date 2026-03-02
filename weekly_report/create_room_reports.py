import os

import pandas as pd
from openpyxl.styles import Font
from openpyxl.styles import Font
from openpyxl.styles import Font
from openpyxl.styles import Font
from openpyxl.styles import Font
from uvbekutils import autosize_xls_cols
from uvbekutils import exit_yes

from weekly_report.__main__ import make_pivot
from weekly_report.__main__ import make_pivot
from weekly_report.__main__ import make_pivot
from weekly_report.__main__ import make_pivot
from weekly_report.__main__ import make_pivot
from weekly_report.__main__ import make_pivot
from weekly_report.__main__ import make_pivot
from weekly_report.constants import noteLines


def create_room_reports(sincere_df, sincere_data_file, file_date, report_by, str_output_dir_rooms ):
    """ create all the room reports """
    report_by = report_by.upper()
    if report_by == 'M':
        report_var = "month"
    elif report_by == 'W':
        report_var = "data_week_ending"
    else:
        exit_yes(f"Report_by not W or M: {report_by}")

    orgs = sorted(sincere_df['org_name'].unique(), reverse=True)

    # master_campaigns = sincere_df['master_campaign'].unique()
    # master_campaigns = sorted(master_campaigns)


    for org in orgs:
        print("Org: " + org)
        xlo = sincere_df[sincere_df['org_name'] == org]  # select records for one Org

        # Name workbook for each org with like: "CA-North Coast VoterLetters Summary 2021-08-02.xlsx"
        room_file_name = f"{org} Sincere Summary {file_date}-{report_by}.xlsx"

        excel_output_file = os.path.join(os.path.expanduser(str_output_dir_rooms),
                            org + " Sincere Summary " + file_date + "-" + report_by + ".xlsx")
#    outputDirWFile = os.path.join(outputDir,Path(InputFile).stem + "-" + reportBy)


        teams = sorted(xlo['team_name'].unique())

        writer = pd.ExcelWriter(excel_output_file, engine='openpyxl')

        # pivot on campaigns
        # make_pivot(writer, xlo, [report_var], ['master_campaign', 'parent_campaign_name'], np.sum, 'Campaigns', 'C10')
        # make_pivot(writer, xlo, [report_var], ['master_campaign', 'parent_campaign_name'], 'sum', 'Campaigns', 'C10')
        make_pivot(writer=writer, df=xlo, report_var=[report_var],
                   index_vars=['election', 'master_campaign', 'parent_campaign_name'], aggfunc='sum',
                   sheet_name='Campaigns', freeze_cell='D10')

        # pivot on writers/teams
        # make_pivot(writer, xlo, [report_var], ['team_name'], np.sum, 'Team Sums', 'B10')
        make_pivot(writer=writer, df=xlo, report_var=[report_var], index_vars=['team_name'], aggfunc='sum',
                   sheet_name='Team Sums', freeze_cell='B10')

        # make_pivot(writer, xlo, [report_var], ['team_name', 'writer_name'], np.sum, 'Teams w Writers', 'C10')
        make_pivot(writer=writer, df=xlo, report_var=[report_var], index_vars=['team_name', 'writer_name'],
                   aggfunc='sum', sheet_name='Teams w Writers', freeze_cell='C10')

        # show COUNT (rather than np.sum) of requests to identify potential organizers
        make_pivot(writer=writer, df=xlo, report_var=[report_var], index_vars=['team_name', 'writer_name'],
                   aggfunc='count', sheet_name='Teams w Counts', freeze_cell='C10')
        # TODO: change var label on address_count to request_count

        # Meika's two reports - cols by campaign not date
        # make_pivot(writer, xlo, ['master_campaign', 'parent_campaign_name'], ['team_name'], np.sum, 'Team by Campaigns', 'B11')
        # make_pivot(writer, xlo, ['master_campaign', 'parent_campaign_name'], ['team_name'], 'sum', 'Team by Campaigns', 'B11')
        make_pivot(writer=writer, df=xlo, report_var=['election', 'master_campaign', 'parent_campaign_name'],
                   index_vars=['team_name'], aggfunc='sum', sheet_name='Team by Campaigns', freeze_cell='B11')

        # make_pivot(writer, xlo, ['master_campaign', 'parent_campaign_name'], ['team_name', 'writer_name'], np.sum,
        #            'Teams w Writers by Campaign', 'B11')
        # make_pivot(writer, xlo, ['master_campaign', 'parent_campaign_name'], ['team_name', 'writer_name'], 'sum',
        #            'Teams w Writers by Campaign', 'B11')
        make_pivot(writer=writer, df=xlo, report_var=['election', 'master_campaign', 'parent_campaign_name'],
                   index_vars=['team_name', 'writer_name'], aggfunc='sum', sheet_name='Teams w Writers by Campaign',
                   freeze_cell='C11')
        # # TODO: sort campaigns by master latest date (which?) then campaign latest date, not alpha

        for team in teams:
            xlt = xlo[xlo['team_name'] == team]
            shname = team[:30]
            # make_pivot(writer, xlt, [report_var], ['team_name', 'writer_name'], np.sum, shname, 'C10')
            make_pivot(writer=writer, df=xlt, report_var=[report_var], index_vars=['team_name', 'writer_name'],
                       aggfunc='sum', sheet_name=shname, freeze_cell='C10')

        # Insert a notes tab as first one. Must be openpyxl because excelwriter is dataframe only
        wb = writer.book
        wb.create_sheet('Notes', 0)
        ws = wb["Notes"]
        r = 8
        for line in noteLines:
            ws.cell(row=r, column=1).value = line
            r = r + 1

        for sh in wb.worksheets:
            autosize_xls_cols(sh)
            # wb.active = 0 # did not work - still first two sheets activated!

        min_date2 = sincere_df['created_at'].min()
        max_date2 = sincere_df['created_at'].max()
        for sh in wb.worksheets:
            sh['A1'].value = "Room: " + org
            sh['A1'].font = Font(b=True, size=20)
            if sh.title not in ['Notes', 'Campaigns', 'Team Sums', 'Teams w Writers', 'Teams w Counts']:
                sh['A2'].value = "Team: " + sh.title
            sh['A2'].font = Font(b=True, size=16)

            if sh.title != 'All Teams':
                sh['A3'].value = ("By Month" if report_by.lower() == "m" else "By Week")
                sh['A3'].font = Font(b=True, size=12)

            sh['A4'].value = "Date range, inclusive: " + min_date2 + " to " + max_date2
            sh['A4'].font = Font(size=12)
            sh['A5'].value = "Source: " + sincere_data_file
            sh['A5'].font = Font(size=12)

            sh.sheet_view.tabSelected = False  # Deselects all via loop; defaults to index = 0 selected (active doesn't deselect any, only  activates!).

        writer.close()
        # writer.save()
