from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from dateutil import parser
from datetime import *
import pytz
import constants

def loadWorkbook(self):
    print "loading workbook..."
    wb = load_workbook('./pat.xlsx')
    sheet_names = wb.get_sheet_names()
    now = datetime.now(pytz.timezone(constants.TIME_ZONE))
    today_date_sheet_exists = False
    # Go through all of sheet names, set flag to true if one of the sheet names is today's month and year
    for name in sheet_names:
        if (name != "template"):
            sheet_date = parser.parse(str(name))
            if now.year == sheet_date.year and now.month == sheet_date.month:
                today_date_sheet_exists = True
    active_sheet = ""
    today_sheet_name = str(now.strftime('%Y')) + "-" + str(now.strftime('%m'))
    self.today_date = str(now.strftime('%-m')) + "/" + str(now.strftime('%-d')) + "/" + str(now.strftime('%y'))
    if (today_date_sheet_exists):
        print "editing existing sheet.."
        active_sheet = wb.get_sheet_by_name(today_sheet_name)
    else:
        print "creating new sheet.."
        template_sheet = wb.get_sheet_by_name("template")
        active_sheet = wb.copy_worksheet(template_sheet)
        active_sheet.title = str(now.strftime('%Y')) + "-" + str(now.strftime('%m'))
        # Create data validation for outage types
        dv = DataValidation(type="list", formula1='"No Outage,CaaS Unplanned,CaaS Planned,BSP Infrastructure Unplanned,BSP Infrastructure Planned,IAE Unplanned,IAE Planned,Other"', allow_blank=True)
        active_sheet.add_data_validation(dv)
        tmp_col = constants.init_outage_type_col
        while tmp_col < constants.last_index_col:
            tmp_row = constants.init_outage_type_row
            while tmp_row < constants.last_index_row:
                dv.add(active_sheet.cell(row=tmp_row,column=tmp_col))
                tmp_row += 1
            tmp_col += 7
        # Create dates on left side
        day = 1
        tmp_row = constants.init_dates_row
        while tmp_row <= constants.last_index_row:
            date_to_write = str(now.strftime('%-m')) + "/" + str(day) + "/" + str(now.strftime('%y'))
            active_sheet.cell(row=tmp_row,column=constants.dates_col).value = date_to_write
            day += 1
            tmp_row += 14
        print "finished creating new sheet"

    writeResolvedIncidentsToExcel(self,wb,active_sheet)
    saveWorkbook(self,wb)

def writeResolvedIncidentsToExcel(self,wb,active_sheet):
    wb.active = wb.index(active_sheet)
    tmp_col = constants.init_http_check_col
    while tmp_col <= constants.last_index_col:
        check_name = active_sheet.cell(row=constants.http_check_row,column=tmp_col).value
        for incident in self.resolved_incidents_array:
            titleExists = False
            if incident.title == "check-api-umbrella-status-non-prod" or incident.title == "check_docker dtr-":
                if check_name in incident.title:
                    titleExists = True
            elif check_name == incident.title:
                titleExists = True
            tmp_row = constants.init_dates_row
            while tmp_row <= constants.last_index_row:
                date = active_sheet.cell(row=tmp_row,column=constants.dates_col).value
                if date == self.today_date:
                    if titleExists:
                        row_increment = 1
                        while row_increment <= 11:
                            if not active_sheet.cell(row=tmp_row + row_increment,column=tmp_col + 1).value and not active_sheet.cell(row=tmp_row + row_increment,column=tmp_col + 2).value:
                                print "writing " + incident.title + " incident.."
                                # Write start time
                                active_sheet.cell(row=tmp_row + row_increment,column=tmp_col + 1).value = incident.created_at
                                # Write end time
                                active_sheet.cell(row=tmp_row + row_increment,column=tmp_col + 2).value = incident.resolved_time
                                break
                            row_increment += 1
                    if not active_sheet.cell(row=tmp_row + 2,column=tmp_col + 1).value:
                        active_sheet.cell(row=tmp_row + 2,column=tmp_col + 4).value = "No Outage"
                tmp_row += 14
        tmp_col += 7

def saveWorkbook(self,wb):
    print "saving workbook.."
    wb.save('./pat.xlsx')
    print "finished!"
