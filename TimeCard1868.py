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

def calculateHours(times):
  result = 0.0
  i = 0
  while i < len(times) - 1:
    result += (times[i + 1] - times[i]).seconds / 3600.0
    i += 2
  return result

# Stores the scans of a single student on the same day
class DayReport:
   def __init__(self):
      self.scans = []
      self.ignored = []
      self.warn = False

   # after the data is loaded, fix up the data
   def fixUp(self, name, date, track):
      self.scans.sort()
      i = 0
      while i < len(self.scans) - 1:
         if (self.scans[i+1] - self.scans[i]).seconds < min_separation:
            self.ignored.append(self.scans[i])
            del self.scans[i]
         else:
            i += 1
      if len(self.ignored) > 0:
         warnings.append((date, track, name, 
                          ("%d near duplicate events ignored" %
                           len(self.ignored))))
      if len(self.scans) % 2 != 0:
         self.warn = len(self.scans) == 1
         msg = ("Odd number of events: " + 
                ', '.join(map(lambda d: d.strftime('%H:%M'), self.scans)))
         warnings.append((date, track, name, msg))
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

class Track:
  def __init__(self, name, required_hours):
    self.name = name
    self.required_hours = required_hours
    # map(name, map(date, DayReport))
    self.times = {}
    # list(date)
    self.dates = []
    # map(name, hours)
    self.total = {}
    # map(week, hours)
    self.byWeek = {}

def buildTimesheet(workbook, track):
  sheet = workbook.add_worksheet(track.name)
  row = 0
  sheet.write(row, 0, 'Name')
  sheet.set_column(0, 0, 20)
  sheet.write(row, 1, 'Total')
  col = 1
  for d in track.dates:
    col += 1
    sheet.write(row, col, d, format_date)

  for name in names:
    total = 0.0
    row = row + 1
    col = 1
    sheet.write(row, 0, name)
    for d in track.dates:
      col += 1
      if name in track.times and d in track.times[name]:
        day = track.times[name][d]
        hours = day.hours()
        total += hours
        sheet.write(row, col, hours, yellow_num if day.warn else format_num)
    track.total[name] = total
    sheet.write(row, 1, total,
                green_total if total >= track.required_hours else black_total)

# default parameters
scanfile = list()
outfile = 'timecard.xlsx'
startdate = parseDate("01/01/2016")
enddate = parseDate("12/31/2016")
min_separation = 120 # ignore events less than 2 minutes apart
business_scanner = "105059"

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
            i = i+1

tech_track = Track("Technical", 90.0)
business_track = Track("Business", 10.0)

# list(tuple(date, track, name, message))
warnings = []

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
                  if row[1] == business_scanner:
                    track = business_track
                  else:
                    track = tech_track
                  times = track.times.setdefault(row[0], {})
                  times.setdefault(day, DayReport()).append(dt)
                  if day not in track.dates :
                    track.dates.append(day)

for track in [tech_track, business_track]:
  track.dates.sort(reverse=True)
  for (name, entries) in track.times.items():
    for (date, report) in entries.items():
      report.fixUp(name, date, track.name)
      week = int(date.strftime('%U'))
      track.byWeek[week] = track.byWeek.get(week, 0) + report.hours()

warnings.sort()
names = sorted(set(list(tech_track.times.keys()) + 
                   list(business_track.times.keys())))

print ("Total: ", len(names), 'names', 'with ', len(tech_track.dates),
       ' technical and ', len(business_track.dates), 'business days')
print ("Generating report from:", startdate, "to: ", enddate)

# Now prep the xlsx workbook
workbook  = xlsxwriter.Workbook(outfile)
format_date = workbook.add_format({'num_format': 'mm/dd/yy'})
green_total = workbook.add_format({'num_format':'0.00'})
green_total.set_bold()
green_total.set_bg_color('#00cc66')
black_total = workbook.add_format({'num_format':'0.00'})
black_total.set_bold()
yellow_num = workbook.add_format({'num_format':'0.00'})
yellow_num.set_bg_color('yellow')
format_num = workbook.add_format({'num_format':'0.00'})

buildTimesheet(workbook, tech_track)
buildTimesheet(workbook, business_track)

total_sheet = workbook.add_worksheet('Totals')
total_sheet.write(0, 0, 'Name')
total_sheet.set_column(0, 0, 20)
total_sheet.write(0, 1, 'Technical Hours')
total_sheet.set_column(1, 3, 15)
total_sheet.write(0, 2, 'Business Hours')
total_sheet.write(0, 3, 'Total Hours')
row = 0
for name in names:
  row += 1
  total_sheet.write(row, 0, name)
  tech_total = tech_track.total.get(name, 0.0)
  business_total = business_track.total.get(name, 0.0)
  overall_format = (green_total if tech_total + business_total >= 100
                    else black_total)
  total_sheet.write(row, 1, tech_total, overall_format)
  total_sheet.write(row, 2, business_total,
                    green_total if business_total >= 10 else black_total)
  total_sheet.write(row, 3, tech_total + business_total,
                    overall_format)
row += 5
for week in sorted(set(list(tech_track.byWeek.keys()) +
                       list(business_track.byWeek.keys()))):
  row += 1
  total_sheet.write(row, 0, 'Week %d' % week)
  tech = tech_track.byWeek.get(week, 0)
  business = business_track.byWeek.get(week, 0)
  total_sheet.write(row, 1, tech, black_total)
  total_sheet.write(row, 2, business, black_total)
  total_sheet.write(row, 3, tech + business, black_total)

warn_sheet = workbook.add_worksheet('Warnings')
warn_sheet.write(0, 0, 'Date')
warn_sheet.write(0, 1, 'Track')
warn_sheet.write(0, 2, 'Name')
warn_sheet.set_column(2, 2, 20)
warn_sheet.write(0, 3, 'Warning')
warn_sheet.set_column(3, 3, 60)
row = 0
for (date, track, name, msg) in warnings:
   row += 1
   warn_sheet.write(row, 0, date, format_date)
   warn_sheet.write(row, 1, track)
   warn_sheet.write(row, 2, name)
   warn_sheet.write(row, 3, msg)

workbook.close()
