import pandas as pd
import xlsxwriter


def get_filename():
    file_name = raw_input('What would you like to name your xlsx?')
    if file_name is '':
        file_name = 'default'
    file_name = str(file_name)
    file_name = file_name + '.xlsx'
    return file_name
    
def initialize_page(writer, sheetName):
    initialize_worksheet_df = pd.DataFrame([])
    initialize_worksheet_df.to_excel(writer, sheet_name = sheetName)

def column_to_string(n):
    div = n
    string = ""
    temp = 0
    while div>0:
        module=(div)%26
        string = chr(65+module)+string
        div = (div-module)/26
    return string

def string_to_column(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

class Workbook(object):
    def __init__(self, writer):
        workbook = writer.book

        self.formats = self.set_formats(workbook)

    def set_formats(self, workbook):
        formats = {
            'h1':workbook.add_format({'font_name': 'Calibri', 'font_color': '#EC2551', 'font_size': 48, 'align': 'center', 'bg_color':'white'}),
            'h1_5': workbook.add_format({'font_name': 'Calibri', 'font_color': '#EC2551', 'font_size': 48, 'align': 'left', 'italic':True, 'bg_color':'white'}),
            'h1_5_subheader': workbook.add_format({'font_name': 'Calibri', 'font_color': '#EC2551', 'font_size': 22, 'bg_color':'white'}),
            'h2': workbook.add_format({'font_name': 'Calibri', 'font_color': 'white', 'bg_color': '#002241', 'font_size': 22, 'border':True, 'align': 'center'}),
            'h3': workbook.add_format({'font_name': 'Calibri', 'bg_color': '#D2EAFF', 'font_size': 20, 'border': True, 'align': 'center'}),
            'hr': workbook.add_format({'bg_color':'white', 'bottom':True}),
            'p': workbook.add_format({'font_name':'Calibri', 'font_size':18, 'align':'center'}),
            'th': workbook.add_format({'bg_color':'white', 'align': 'center', 'bottom':True, 'font_name':'Calibri', 'font_size':18}),
            'table_index': workbook.add_format({'bg_color':'white', 'font_name':'Calibri', 'right':True, 'font_size':18}),
            'table_value': workbook.add_format({'bg_color':'white', 'font_name':'Calibri', 'right':True, 'left':True, 'font_size':18, 'num_format': '#,##0_);(#,##0)' }),
            'month_header': workbook.add_format({'bg_color':'white', 'font_name':'Calibri', 'bottom':True, 'font_size':18, 'num_format':'[$-409]mmmm yyyy;@', 'align': 'center'}),
            'percentage': workbook.add_format({'num_format':'0.00%'}),
            'title_date': workbook.add_format({'align': 'center', 'font_size':22, 'bottom':True, 'num_format':'[$-409]mmmm yyyy;@'})
            }
        
        return formats
    

class Dashboard(object):

    def __init__(self, writer, workbook, track_export, client_name):
        worksheet = writer.sheets['Dashboard']
        worksheet.hide_gridlines(2)
        
        self.set_column_widths(worksheet)
        self.create_title_section(worksheet, workbook, track_export, start_row = 1,start_column = 1, client_name = 'A+E Network', merge_width=23)
   
    
    
    def set_column_widths(self, worksheet):
        worksheet.set_column(0, 0, 1.83)
        worksheet.set_column(1, 1, 3.83)
        worksheet.set_column(2, 2, 7.83)
        worksheet.set_column(3, 3, 38.17)
        worksheet.set_column(4, 9, 16.5)
        worksheet.set_column(10, 10, 28.6)
        worksheet.set_column(11, 11, 22.5)
        worksheet.set_column(12, 15, 19.5)
        worksheet.set_column(16, 16, 16.5)
        worksheet.set_column(17, 17, 21)
        worksheet.set_column(18, 18, 19.67)
        worksheet.set_column(19, 19, 13)
        worksheet.set_column(20, 20, 8)
        worksheet.set_column(21, 21, 8.7)
        worksheet.set_column(22, 22, 21)
        worksheet.set_column(23, 23, 18)
        worksheet.set_column(24, 24, 44)
        worksheet.set_column(25, 30, 10)
        worksheet.set_default_row(29)
        

    def create_title_section(self, worksheet, workbook, track_export, start_row, start_column, client_name,merge_width = 23):
        worksheet.merge_range(start_row,start_column,start_row+1,start_row+merge_width,client_name + ' - Competitive Report', workbook.formats['h1'])
        worksheet.merge_range(start_row+2,start_column,start_row+2,start_column+merge_width, track_export.month().columns[-2], workbook.formats['title_date'] )


    
