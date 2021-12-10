import os
import sys
print(os.path.dirname(sys.executable))
import glob
import pathlib
from numpy import remainder
#import numpy as np
import pandas as pd
import datetime as dt
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name    # https://xlsxwriter.readthedocs.io/working_with_cell_notation.html

# Folder locations
current_directory = pathlib.Path(__file__).parent.resolve()
print(current_directory)
#current_directory = 'C:/Users/Calvin/OneDrive/Dev/ManaloMachineShop/Payroll'
attendance_directory = os.path.join(current_directory, 'Attendance')
zkteco_logs_directory = os.path.join(current_directory, 'ZKT Eco logs')

# Finding latest zkteco attendance log
list_of_excels = glob.glob(os.path.join(zkteco_logs_directory, '*.xls*'))
latest_attendance = max(list_of_excels, key=os.path.getctime)

attendance_log = pd.read_excel(latest_attendance)
attendance_log['Date And Time'] = pd.to_datetime(attendance_log['Date And Time'])
attendance_log['Employee Name'] = attendance_log['Last Name'] + ", " + attendance_log['First Name']

multi_index = []
for index, row in attendance_log[['Personnel ID', 'Employee Name']].drop_duplicates('Personnel ID').sort_values('Employee Name').iterrows():
    index_names = ['IN', 'OUT', 'Total', 'OT']
    for index_name in index_names:
        multi_index.append((row['Personnel ID'], row['Employee Name'], index_name))

index = pd.MultiIndex.from_tuples(multi_index, names=['ID', 'Employee Name', 'Time'])

# pd.Series(np.random.randn(8), index=index)

dates = attendance_log['Date And Time'].dt.date.unique()
personnel_ids = attendance_log['Personnel ID'].unique()
attendance = pd.DataFrame(columns=dates, index=index)

for id in personnel_ids:
    for date in dates:
        print('\n')
        print(date)
        filter = (attendance_log['Date And Time'].dt.date == date) & (attendance_log['Personnel ID'] == id)
        time_in_or_out = attendance_log[filter]['Date And Time']
        print(time_in_or_out)

        if len(time_in_or_out) == 0:        # Absent
            print("Empty!")
            continue
        elif len(time_in_or_out) == 1:      # Only 1 Check-In/Out
            in_out_status = attendance_log[filter]['In/Out Status'].values[0].split('-')[1].upper()
            attendance.loc[(id, attendance_log[attendance_log['Personnel ID'] == id]['Employee Name'], in_out_status), date] = min(time_in_or_out.dt.time)
            continue

        attendance.loc[(id, attendance_log[attendance_log['Personnel ID'] == id]['Employee Name'], 'IN'), date] = min(time_in_or_out.dt.time)
        attendance.loc[(id, attendance_log[attendance_log['Personnel ID'] == id]['Employee Name'], 'OUT'), date] = max(time_in_or_out.dt.time)

print("Saving Attendance Excel")
date_today = dt.date.today()
dated_attendance_directory = os.path.join(attendance_directory, date_today.strftime('%Y'), date_today.strftime('%B'))
os.makedirs(dated_attendance_directory, exist_ok=True)
attendance_location = os.path.join(dated_attendance_directory, date_today.strftime('%m-%d-%Y') + '.xlsx')
print(attendance_location)
attendance['Total'] = ''
attendance.fillna('-', inplace=True)

with pd.ExcelWriter(attendance_location, engine='xlsxwriter') as writer:
    attendance.to_excel(writer, sheet_name='Attendance', index=True)
    workbook = writer.book
    worksheet = writer.sheets['Attendance']
    # Formatting table
    time_format = workbook.add_format({'num_format': '[h]:mm;@', 'align': 'center'})
    number_format = workbook.add_format({'num_format': '0'})
    worksheet.set_column('A:A', 3.5)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('D:I', 12, time_format)
    worksheet.freeze_panes(1, 0)
    worksheet.set_zoom(150)
    worksheet.set_landscape()
    red_fill_format = workbook.add_format({'pattern': 3, 'bg_color': 'red'})
    worksheet.conditional_format('D1:Z100', {'type':    'cell',
                                        'criteria': '==',
                                        'value':    '"-"',
                                        'format':   red_fill_format})
    # Writing Formulas
    bottom_border_format        = workbook.add_format({'num_format': '[h]:mm;@', 'align': 'center', 'bottom': 1})
    right_border_format         = workbook.add_format({'num_format': '[h]:mm;@', 'align': 'center', 'bold': True, 'right': 1})
    bottom_border_edge_format   = workbook.add_format({'num_format': '[h]:mm;@', 'align': 'center', 'bold': True, 'bottom': 1, 'right': 1})
    percentage_format           = workbook.add_format({'num_format': '0%', 'align': 'center'})
    total_rows, total_cols = attendance.shape

    for row_number in range(1, total_rows+1):
        start_col   = len(attendance.index.names)
        end_col     = total_cols + len(attendance.index.names)
        for col_number in range(start_col, end_col):
            index       = row_number%4    # 1:IN, 2:OUT 3:Total 4:OT
            row_adj     = {1:0, 2:0, 3:2, 0:3}
            cell_in     = xl_rowcol_to_cell(row_number - row_adj[index], col_number)
            cell_out    = xl_rowcol_to_cell(row_number - row_adj[index] + 1, col_number)
            cell_total  = xl_rowcol_to_cell(row_number, col_number)
            cell_ot     = xl_rowcol_to_cell(row_number , col_number)
            print(f'{index}: {cell_in} {cell_out} {cell_ot} {cell_ot}')

            if index == 3:      # Total Hours
                worksheet.write(cell_total, f'=IFERROR(MIN({cell_out} - {cell_in} - TIME(1,0,0), TIME(8,0,0)), "")')
            
            elif index == 0:    # Total OT
                worksheet.write(cell_ot, f'=IFERROR(MAX({cell_out} - {cell_in} - TIME(9,0,0), TIME(0,0,0)), "")', bottom_border_format)

            # Last column
            if col_number == total_cols + start_col - 1:
                print(cell_total)
                first = xl_rowcol_to_cell(row_number, col_number - total_cols + 1)
                last  = xl_rowcol_to_cell(row_number, col_number - 1)
                worksheet.write(cell_ot, "", right_border_format)
                if index == 3:    # Total Regular Hours for the week
                    print(f'=SUM({first}:{last} in {cell_total})')
                    worksheet.write(cell_total, f'=SUM({first}:{last})', right_border_format)
                elif index == 0:    # Total overtime Regular Hours for the week
                    print(f'=SUM({first}:{last} in {cell_total})')
                    worksheet.write(cell_ot, f'=SUM({first}:{last})', bottom_border_edge_format)

            # Adding normal/holidays rate at the bottom
            if row_number == total_rows and col_number != end_col - 1:
                cell_end = xl_rowcol_to_cell(row_number+1, col_number)
                print(f"End {cell_end}")
                worksheet.write(cell_end, '100%', percentage_format)

    # PAYROLL WORKSHEET
    master_file = pd.read_excel(os.path.join(current_directory, 'Master File.xlsx'))
    payroll = attendance.reset_index()[['ID', 'Employee Name']].drop_duplicates().reset_index(drop=True)
    payroll = payroll.merge(master_file, how='left', left_on='ID', right_on='Personnel ID')
    payroll = payroll.drop(columns = ['First Name', 'Last Name', 'Personnel ID', 'ID'])
    payroll[['Days Worked', 'Regular Hours', 'Overtime Hours', 'Holiday Pay', 'Regular Wage Pay', 'Overtime Pay', 'Allowance Pay', 'Incentive Pay', 'Net Pay', 'Signature']] = ""
    payroll.fillna('-', inplace=True)
    payroll.to_excel(writer, sheet_name='Payroll', startrow=3, index=False)

    # Formatting col width & table
    worksheet = writer.sheets['Payroll']
    worksheet.freeze_panes(0, 1)
    worksheet.set_column('A:A', 23)
    worksheet.set_column('B:B', 6.5)
    worksheet.set_column('C:C', 8.5)
    worksheet.set_column('D:D', 9.5)
    worksheet.set_column('E:G', 5.5)
    worksheet.set_column('H:K', 8.5)
    worksheet.set_column('L:O', 9)
    worksheet.set_column('P:P', 10)
    worksheet.set_column('Q:Q', 15)
    worksheet.set_zoom(85)
    worksheet.set_landscape()
    worksheet.set_paper(5)  # Legal paper format (8 1/2 x 14 in)
    worksheet.set_margins(left=0.25, right=0.25, top=0.75, bottom=0.75)

    # PAYROLL title
    worksheet.set_row(0, 45)
    merged_title_format = workbook.add_format({'font': 'Britannic Bold', 'font_size': 45, 'bold': 1, 'align': 'center', 'valign': 'vcenter'})
    worksheet.merge_range('A1:Q1', 'PAYROLL', merged_title_format)

    # Date period
    date_period = f"For the period {min(attendance_log['Date And Time'].dt.date).strftime('%m/%d/%Y')} - {max(attendance_log['Date And Time'].dt.date).strftime('%m/%d/%Y')}"
    merged_date_period_format = workbook.add_format({'font': 'Times New Roman', 'font_size': 15, 'align': 'center', 'valign': 'bottom'})
    worksheet.merge_range('F2:L2', date_period, merged_date_period_format)

    # Formatting headers
    worksheet.set_row(3, 30)
    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center', 'border': 1}) 
    for col_num, value in enumerate(payroll.columns.values):
        worksheet.write(3, col_num, value, header_format)

    # Writing formulas to table
    name_format     = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
    default_format  = workbook.add_format({'num_format': '0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    hour_format     = workbook.add_format({'num_format': '[h]:mm;@', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    net_pay_format  = workbook.add_format({'num_format': '₱#,##0_);($#,##0)', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    
    for row_num, value in payroll.iterrows():
        adjusted_row = row_num + 4
        worksheet.set_row(adjusted_row, 20)

        for col_num, col_name in enumerate(payroll.columns.values):
            cell = xl_rowcol_to_cell(adjusted_row, col_num)

            if col_name in ['Days Worked', 'Regular Hours', 'Overtime Hours']:
                start_cell      = xl_rowcol_to_cell(row_num * 4 + 3, start_col)
                end_cell        = xl_rowcol_to_cell(row_num * 4 + 3, total_cols + 1)
                hour_cell       = xl_rowcol_to_cell(row_num * 4 + 3, total_cols + 2)
                overtime_cell   = xl_rowcol_to_cell(row_num * 4 + 4, total_cols + 2)
                formulas = {'Days Worked':      f'=COUNTIF(Attendance!{start_cell}:{end_cell}, "<>"&"*")',
                            'Regular Hours':     f'=Attendance!{hour_cell}',
                            'Overtime Hours':   f'=Attendance!{overtime_cell}'}
                worksheet.write(cell, formulas[col_name], hour_format if 'Hours' in col_name else default_format)

            elif 'Pay' in col_name:
                daily_wage_cell         = xl_rowcol_to_cell(adjusted_row, 1)
                daily_incentive_cell    = xl_rowcol_to_cell(adjusted_row, 2)
                daily_allowance_cell    = xl_rowcol_to_cell(adjusted_row, 3)
                days_worked_cell        = xl_rowcol_to_cell(adjusted_row, 7)
                hours_worked_cell       = xl_rowcol_to_cell(adjusted_row, 8)
                overtime_hours_cell     = xl_rowcol_to_cell(adjusted_row, 9)
                holiday_pay_cell        = xl_rowcol_to_cell(adjusted_row, 10)
                incentive_pay_cell      = xl_rowcol_to_cell(adjusted_row, 14)
                formulas = {
                                'Regular Wage Pay':         f'={daily_wage_cell}*{hours_worked_cell}*3',
                                'Overtime Pay':     f'={daily_wage_cell}*{overtime_hours_cell}*3*125%',
                                'Allowance Pay':    f'={daily_allowance_cell}*{days_worked_cell}',
                                'Incentive Pay':    f'={daily_incentive_cell}*{days_worked_cell}',
                                'Holiday Pay':      '', # Empty cell. Manually calculated.
                                'Net Pay':          f'=SUM({holiday_pay_cell}:{incentive_pay_cell})'
                            }
                worksheet.write(cell, formulas[col_name], default_format if col_name != 'Net Pay' else net_pay_format)

            else:
                worksheet.write(cell, value[col_name], default_format if 'Name' not in col_name else name_format)


print("Opening excel file")
os.system(f"start EXCEL.exe {attendance_location}")
