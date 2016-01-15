# Time card applicaiton for Team 1868 - for work with Opticon scanner
# in USB MSD mode (C04) and with fields programmed to be
# name - text field - bar code is a text string
# serial - numerical - scanner seried number
# time - time code in 24HR HH:MM:SS format
# date - date code in MM/DD/YYYY format.
# 9/12/2014 - Partha Srinivasan initial cut
import sys
import os.path
import csv
import datetime
import xlsxwriter

def parseDate(str):
  return datetime.datetime.strptime(str, '%m/%d/%Y').date()

# extract the date from a timestamp with times before 4am counting as previous
# day.
def adjustDate(ts):
  return (ts - datetime.timedelta(0, 4 * 3600)).date()

# default file names
scanfile = list()
outfile = 'timecard.xlsx'
startdate = parseDate("01/01/2016")
enddate = parseDate("12/31/2016")
minSeparation = 120

if len(sys.argv) > 1 :      # arguments passed
    i = 1
    while i < len(sys.argv):
        if sys.argv[i] == '-o':
            outfile = sys.argv[i+1]
            i = i+2
        elif sys.argv[i] == '-s':
            startdate = parseDate(sys.argv[i+1])
            i = i+2
        elif sys.argv[i] == '-e':
            enddate = parseDate(sys.argv[i+1])
            i = i+2
        else :
            scanfile.append(sys.argv[i])
            i = i+1 ;

# map(name, map(date, list(datetime)))
students = {}

# list(date)
dates = []

# list(tuple(string, date, string))
warnings = []

def calculateHours(times):
  result = 0.0
  i = 0
  while i < len(times) - 1:
    result += (times[i + 1] - times[i]).seconds / 3600.0
    i += 2
  return result

class DayReport:
   def __init__(self):
      self.scans = []
      self.ignored = []
      self.warn = False
   def fixUp(self, name, date):
      self.scans.sort()
      i = 0
      while i < len(self.scans) - 1:
         if (self.scans[i+1] - self.scans[i]).seconds < minSeparation:
            self.ignored.append(self.scans[i])
            del self.scans[i]
         else:
            i += 1
      if len(self.ignored) > 0:
         self.warn = True
         warnings.append((name, date, ("%d near events ignored" %
                                       len(self.ignored))))
      if len(self.scans) % 2 != 0:
         self.warn = True
         warnings.append((name, date, ("Odd number of events (%d)" %
                                       len(self.scans))))
   def append(self, time):
      self.scans.append(time)

   def hours(self):
     if len(self.scans) < 2:
       return 0.0
     elif len(self.scans) % 2 == 0:
       return calculateHours(self.scans)
     else:
       return max(calculateHours(self.scans),
                  calculateHours(self.scans[1:]))

#Now read in the scanner files - one file at a time
for file in scanfile:
    print ('Reading file', file)
    with open(file, 'rt') as inputfile :
        reader = csv.reader(inputfile, delimiter=',', quotechar='|')
        for row in reader:
            if len(row) > 0 and not row[0].startswith('#'):
                d = datetime.datetime.strptime(row[3], '%m/%d/%Y')
                t = datetime.datetime.strptime(row[2], '%H:%M:%S')
                dt = datetime.datetime.combine(d.date(), t.time())
                day = adjustDate(dt)
                if startdate <= day and day <= enddate:
                   times = students.setdefault(row[0], {})
                   times.setdefault(day, DayReport()).append(dt)
                   if day not in dates :
                      dates.append(day)

dates.sort()
for (name, entries) in students.items():
  for (date, report) in entries.items():
      report.fixUp(name, date)

warnings.sort()
names = sorted(students.keys())

print ("Total: ", len(students), 'names', 'with ', len(dates), 'days attended')
print ("Generating report from:", startdate, "to: ", enddate)

# Now prep the xlsx workbook
workbook  = xlsxwriter.Workbook(outfile)
sheet = workbook.add_worksheet('Hours')
format_date = workbook.add_format({'num_format': 'mm/dd/yy'})
green_num = workbook.add_format({'num_format':'0.000'})
green_num.set_bg_color('green')
yellow_num = workbook.add_format({'num_format':'0.000'})
yellow_num.set_bg_color('yellow')
format_num = workbook.add_format({'num_format':'0.000'})

row = 0
sheet.write(row, 0, 'Name')
sheet.write(row, 1, 'Total')
col = 1
for d in dates:
  col += 1
  sheet.write(row, col, d, format_date)

for name in names:
  total = 0.0
  row = row + 1
  col = 1
  sheet.write(row, 0, name)
  for d in dates:
    col += 1
    days = students[name]
    if d in days:
      hours = days[d].hours()
      warn = days[d].warn
    else:
      hours = 0.0
      warn = False
    total += hours
    sheet.write(row, col, hours, yellow_num if warn else format_num)
  sheet.write(row, 1, total, green_num if total >= 100.0 else format_num)

warn_sheet = workbook.add_worksheet('Warnings')
warn_sheet.write(0, 0, 'Name')
warn_sheet.write(0, 1, 'Date')
warn_sheet.write(0, 2, 'Warning')
row = 0
for (name, date, msg) in warnings:
   row += 1
   warn_sheet.write(row, 0, name)
   warn_sheet.write(row, 1, date, format_date)
   warn_sheet.write(row, 2, msg)

workbook.close()
