import os
import pandas as pd
from tkinter import Tk, filedialog, Label, Entry, Button, messagebox
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from tkinter import filedialog
from tkinter import *
from PIL import Image, ImageTk
from tkinter.ttk import Progressbar
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from pprint import pprint
import csv
import tempfile
from openpyxl.styles import PatternFill
from tkinter import ttk
from tkinter import Toplevel, ttk
from tkinter import font



SERVICE_ACCOUNT_FILE='calender-api-463402-37c84dbc6a80.json'
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

#border
thick_border = Border(
    top=Side(style='thick'),
    bottom=Side(style='thick')
)


# Yellow Color
fill_yellow = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')


# Define fill colors
green_fill = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')  # light green
red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')    # very light red



def fetch_philippine_holidays(start_date, end_date):
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )

    service = build('calendar', 'v3', credentials=credentials)

    calendar_id = 'en.philippines#holiday@group.v.calendar.google.com'

    events = service.events().list(
        calendarId=calendar_id,
        timeMin=start_date.isoformat() + 'Z',
        timeMax=end_date.isoformat() + 'Z',
        singleEvents=True,
        orderBy='startTime'
    ).execute()

    holiday_dates = set()
    for event in events.get('items', []):
        date = event['start'].get('date')
        if date:
            holiday_dates.add(datetime.strptime(date, "%Y-%m-%d").strftime("%Y/%m/%d"))

    return holiday_dates

start=datetime(2020,1,1)
end=datetime(2030,1,1)
holidays = fetch_philippine_holidays(start, end)


def read_csv(filename):
    if filename.lower().endswith('.csv'):
        with open(filename, newline='', encoding='utf-8') as csvfile:
            return list(csv.reader(csvfile))
    elif filename.lower().endswith(('.xlsx', '.xls')):
        df = pd.read_excel(filename)
        return [df.columns.tolist()] + df.values.tolist()
    else:
        raise ValueError("Unsupported file type. Please upload a .csv or .xlsx/.xls file.")





def adjust_to_next_business_day(date_obj, holidays):
    while date_obj.weekday() >= 5 or date_obj.strftime("%Y/%m/%d") in holidays:
        date_obj += timedelta(days=1)
    return date_obj



def extract(filename):
    add = ""
    for elem in filename:
        if elem == '_':
            break
        else:
            add += elem

    sheet_map = {
        "Loan Data": f"{add}_selected_loans.csv",
    }

    for sheet_name, output_csv in sheet_map.items():
        df = pd.read_excel(filename, sheet_name=sheet_name)
        for col in df.select_dtypes(include=['float', 'int']).columns:
            if (df[col] % 1 == 0).all():
                df[col] = df[col].astype(int)
        df.to_csv(output_csv, index=False)


def edit_merge(file):
    extract(file)
    base_name = file.split('_')[0]
    transactions = f'{base_name}_selected_loans.csv'
    

    
    rows = read_csv(file)
    if not rows or len(rows) < 2:
        return []

    rows = rows[1:]  # Skip header
    head = ''
    head2 = ''
    for row in rows:
        if not pd.isna(row[0]) and row[0] != '':
            head = row[0]
        if pd.isna(row[0]) or row[0] == '':
            row[0] = head

        if not pd.isna(row[1]) and row[1] != '':
            head2 = row[1]
        if pd.isna(row[1]) or row[1] == '':
            row[1] = head2
    return rows


####################################################################
##################################################################





def add_weeks(start, tenor):
    date_obj = datetime.strptime(start, "%Y/%m/%d")
    new_date_obj = date_obj + relativedelta(weeks=tenor)
    return new_date_obj.strftime("%Y/%m/%d")

def add_months(start, tenor):
    date_obj = datetime.strptime(start, "%Y/%m/%d")
    new_date_obj = date_obj + relativedelta(months=tenor)
    return new_date_obj.strftime("%Y/%m/%d")

def generate_due_dates(start_date_str, tenor, freq_days, last):
    start_date = datetime.strptime(start_date_str, "%Y/%m/%d")
    end_date_str = add_months(start_date_str, tenor)
    max_end_date = last
    end_date = min(datetime.strptime(end_date_str, "%Y/%m/%d"),
                   datetime.strptime(max_end_date, "%Y/%m/%d"))
    dates = []
    current = start_date
    while current <= end_date:
        adjusted = adjust_to_next_business_day(current, holidays)
        if adjusted<=end_date:
            dates.append(adjusted.strftime("%Y/%m/%d"))
        current =adjusted + timedelta(days=freq_days)
    return dates

def mapping(payments_with_indices, dates, used,before,after):
    missed = []
    for date in dates:
        date_obj = datetime.strptime(date, "%Y/%m/%d")
        mini = (date_obj - timedelta(days=before)).strftime("%Y/%m/%d") #Window here        
        counter=0
        maxi_obj=date_obj
        while(counter<after):
            maxi_obj += timedelta(days=1)
            if maxi_obj.strftime("%Y/%m/%d") in holidays or maxi_obj.weekday()>=5:
                continue
            counter+=1
        maxi = maxi_obj.strftime("%Y/%m/%d")
              
        matched = False
        matched_date = ''
        for global_idx, (pdate, _),cheque in payments_with_indices:
            if global_idx in used:
                continue
            if mini <= pdate <= maxi:
                used.add(global_idx)
                matched = True
                matched_date = pdate
                break
        if matched:
            missed.append([date, matched_date,global_idx,cheque,"No Default"])
        else:
            missed.append([date, '', '',"DEFAULT",maxi])
    return missed

def default(all_rows, start, tenor, repayment, freq, last, used,before,after):
    if freq == 'Biweekly':
        repayment = round(repayment / 2, 3)
        start = add_weeks(start, 2)
    else:
        start = add_months(start, 0)

    end = min(add_months(start, tenor + 1), last)
    amt = round(repayment)

    filtered = []
    for i, row in enumerate(all_rows):
        try:
            date = row[1]
            amount_str = row[4]
            cheque=row[7]
            if date >= start and date <= end and amount_str != '':
                amount = float(amount_str)
                if round(amount) in [amt, amt + 1, amt - 1]:
                    filtered.append((i, [date, round(amount)],cheque))
        except:
            continue

    if freq == 'Biweekly':
        payment_schedule = generate_due_dates(start, tenor, 14, last)
    else:
        payment_schedule = generate_due_dates(start, tenor, 30, last)

    return mapping(filtered, payment_schedule, used,before,after)

def final(loan_file, transaction_file,before,after):
    used = set()
    res = []
    loan_rows = edit_merge(loan_file)
    loan_rows = list(filter(lambda x: x[-1]==True, loan_rows))
    all_rows = read_csv(transaction_file)[1:]
    last_day = all_rows[-1][1]

    for elem in loan_rows:
        start = elem[0]
        tenor = int(elem[2])
        repayment = float(elem[4])
        freq = elem[5]
        principal = elem[1]
        schedule = default(all_rows, start, tenor, repayment, freq, last_day, used,before,after)
        res.append([start, tenor, principal, repayment, freq, schedule])

    return res
















def export_to_excel(loan_file, transaction_file, before, after, output_path):
    results = final(loan_file, transaction_file, before, after)

    wb = Workbook()
    ws = wb.active
    ws.title = "Repayment Schedule"

    headers = ["Loan Info", "Repayment Schedule", "Repayment Date", "Txn ID", "Cheque No.","Flag"]
    ws.append(headers)

    col_widths = [40, 20, 20, 15, 15]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border=thick_border
        cell.fill=fill_yellow

    current_row = 2
    for loan in results:
        start, tenor, principal, repayment, freq, schedule = loan

        loan_info_text = f"Start Date: {start}\nTenor: {tenor}\nPrinciple: {principal}\nRepayment: {repayment}\nFrequency: {freq}"
        repayment_count = len(schedule)
        end_row = current_row + repayment_count - 1
        merge_range = f"A{current_row}:A{end_row}"
        ws.merge_cells(merge_range)
        cell = ws[f"A{current_row}"]
        cell.value = loan_info_text
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.border=thick_border
        for entry in schedule:
            repayment_schedule, repayment_date, txn_id, cheque, flag = entry

            fill = green_fill if flag == "No Default" else red_fill

            
            ws.cell(row=current_row, column=2, value=repayment_schedule)
            ws.cell(row=current_row, column=3, value=repayment_date)
            ws.cell(row=current_row, column=4, value=txn_id)
            ws.cell(row=current_row, column=5, value=cheque)
            ws.cell(row=current_row, column=6, value=flag)

            
            for col in range(2, 7):
                cell = ws.cell(row=current_row, column=col)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = fill
            current_row += 1

    wb.save(output_path)
    print(f"Excel file saved as '{output_path}'")






def export_custom_results_to_excel(results, output_path):
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Repayment Schedule"

    headers = ["Loan Info", "Repayment Schedule", "Repayment Date", "Cheque No." ,"Txn ID", "Flag"]
    ws.append(headers)

    col_widths = [40, 20, 20, 15, 15]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill=fill_yellow
        cell.border=thick_border
    current_row = 2
    for loan in results:
        start, tenor, principal, repayment, freq, cheque,schedule = loan

        loan_info_text = f"Start Date: {start}\nTenor: {tenor}\nPrinciple: {principal}\nRepayment: {repayment}\nFrequency: {freq}"
        repayment_count = len(schedule)
        end_row = current_row + repayment_count - 1
        merge_range = f"A{current_row}:A{end_row}"
        ws.merge_cells(merge_range)
        cell = ws[f"A{current_row}"]
        cell.value = loan_info_text
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.border=thick_border
        for entry in schedule:
            repayment_schedule, repayment_date, txn_id, cheque, flag = entry

            fill = green_fill if flag == "No Default" else red_fill

            
            ws.cell(row=current_row, column=2, value=repayment_schedule)
            ws.cell(row=current_row, column=3, value=repayment_date)
            ws.cell(row=current_row, column=4, value=txn_id)
            ws.cell(row=current_row, column=5, value=cheque)
            ws.cell(row=current_row, column=6, value=flag)

            
            for col in range(2, 7):
                cell = ws.cell(row=current_row, column=col)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = fill
            current_row += 1

    wb.save(output_path)
    print(f"Excel file saved as '{output_path}'")

def manual(start,tenor,repayment,freq,transaction_file,before,after,output_path):
    used = set()
    all_rows=read_csv(transaction_file)[1:]
    last_day=all_rows[-1][1]
    schedule=default(all_rows,start,tenor,repayment,freq,last_day,used,before,after)

    wb=Workbook()
    ws=wb.active
    ws.title='Repayment Schedule'

    
    green_fill = PatternFill(fill_type="solid", fgColor="C6EFCE")
    red_fill = PatternFill(fill_type="solid", fgColor="FFC7CE")
    yellow_fill = PatternFill(fill_type="solid", fgColor="FFFACD")

    bold_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    wrap = Alignment(wrap_text=True, vertical="top")
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    headers=["Loan Info","Repayment Schedule","Repayment Date","Txn ID","Cheque No.","Flag"]
    ws.append(headers)
    for col, header in enumerate(headers,1):
        cell=ws.cell(row=1,column=col,value=header)
        cell.fill = yellow_fill
        cell.font = bold_font
        cell.alignment = center
        cell.border = border

    loan_info = f"Start Date: {start}\nTenor: {tenor}\nRepayment: {repayment}\nFrequency: {freq}"
    start_row = 2
    end_row = start_row + len(schedule) - 1
    ws.merge_cells(start_row=start_row, end_row=end_row, start_column=1, end_column=1)
    cell = ws.cell(row=start_row, column=1, value=loan_info)
    cell.alignment = wrap
    cell.border = border

    for i, row in enumerate(schedule, start=start_row):
        schedule_date, repay_date, txn_id, cheque, flag = row
        fill = green_fill if flag == "No Default" else red_fill
        row_values = [schedule_date, repay_date, txn_id, cheque, flag]

        for j, val in enumerate(row_values, start=2):
            cell = ws.cell(row=i, column=j, value=val)
            cell.fill = fill
            cell.alignment = center
            cell.border = border

    # Set column widths
    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 20

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    print(f"Saved manual preview to {output_path}")
    return output_path, schedule

    
    














