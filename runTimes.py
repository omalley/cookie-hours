#!/usr/bin/env python3

# Time card applicaiton for FRC Team 1868

# Configure the settings for the season in config.yaml.
# I keep all of the data files in a directory structure
# like data/<year>/<dump-date>/*.TXT, but that is just so
# I can keep it straight. The script looks for all files
# named *.TXT under the dataRoot directory.

# When you run the script, it should look like:
#   owen@laptop> ./runTimes.py
#   Reading file data/2017/12-06/archi.TXT
#   Reading file data/2017/12-06/curie.TXT
#   Reading file data/2017/12-06/galileo.TXT
#   Reading file data/2017/12-06/newton.TXT
#   Total: 84 names with 8 technical, 5 business, and 0 post-bag days
#   Generating report timecard.xlsx from: 2016-09-01 to: 2017-05-01

# Upload the file to Google Sheets using "File/Import/Upload/Replace".
# After you upload, run the "Cookies/titles" macro to set the title bars.

# I've been running this script in Python 3.5 on a Mac using MacPorts.
# You'll need to install python35, py35-pip, and py35-readline.
# You'll need to pip install XlsxWriter, and PyYAML.

# 9/12/2014 - Partha Srinivasan initial cut

import os.path
import scanners
import xlsxwriter
import yaml

def buildTimesheet(workbook, names, track):
  sheet = workbook.add_worksheet(track.name)
  row = 0
  sheet.write(row, 0, 'Name')
  sheet.set_column(0, 0, 20)
  sheet.write(row, 1, 'Total')
  col = 1
  trainingNames = sorted(track.trainingNames())
  for eventName in trainingNames:
    col += 1
    sheet.write(row, col, eventName)
  for d in track.dates:
    col += 1
    sheet.write(row, col, d, format_date)

  for name in names:
    total = 0.0
    row = row + 1
    col = 1
    sheet.write(row, 0, name)
    for eventName in trainingNames:
      col += 1
      hours = track.trainingHours(name, eventName)
      if hours != 0:
        sheet.write(row, col, hours, time_formats["normal"])
        total += hours
    for d in track.dates:
      col += 1
      if name in track.people and d in track.people[name].times:
        day = track.people[name].times[d]
        hours = day.hours()
        total += hours
        sheet.write(row, col, hours, time_formats[day.state])
    track.total[name] = total
    sheet.write(row, 1, total,
                green_total if total >= track.required_hours else black_total)

# read configuration from config.yaml file
config = yaml.load(open("config.yaml", "r"))
outfile = config['output']

timecards = scanners.Timecards(config)
timecards.printSummary()
print ("Generating report", outfile)

# Now prep the xlsx workbook
workbook  = xlsxwriter.Workbook(outfile)

def makeColorFormat(color):
  result = workbook.add_format({'num_format':'0.00'})
  if color != "white":
    result.set_bg_color(color)
  return result

format_date = workbook.add_format({'num_format': 'mm/dd/yy'})
green_total = makeColorFormat('#00cc66')
green_total.set_bold()
black_total = makeColorFormat("white")
black_total.set_bold()

time_formats = {"normal": makeColorFormat("white"),
                "error": makeColorFormat("yellow"),
                "manual": makeColorFormat("#b7fcff")}

names = timecards.names()
total_sheet = workbook.add_worksheet('Totals')
buildTimesheet(workbook, names, timecards.tech_track)
buildTimesheet(workbook, names, timecards.business_track)
buildTimesheet(workbook, names, timecards.post_bag_track)
buildTimesheet(workbook, names, timecards.preseason_track)

total_sheet.write(0, 0, 'Name')
total_sheet.set_column(0, 0, 20)
total_sheet.write(0, 1, 'Technical Hours')
total_sheet.set_column(1, 6, 15)
total_sheet.write(0, 2, 'Business Hours')
total_sheet.write(0, 3, 'Total Pre-Bag')
total_sheet.write(0, 4, 'Post-Bag Hours')
total_sheet.write(0, 5, 'Total Hours')
total_sheet.write(0, 6, 'Pre-Season Hours')
row = 0
business_required = timecards.business_track.required_hours
prebag_required = (timecards.tech_track.required_hours +
                   timecards.business_track.required_hours)
postbag_required = timecards.post_bag_track.required_hours
for name in timecards.names():
  row += 1
  total_sheet.write(row, 0, name)
  tech_total = timecards.tech_track.total.get(name, 0.0)
  business_total = timecards.business_track.total.get(name, 0.0)
  overall_format = (green_total
                       if tech_total + business_total >= prebag_required
                    else black_total)
  total_sheet.write(row, 1, tech_total, overall_format)
  total_sheet.write(row, 2, business_total,
                    green_total if business_total >= business_required
                    else black_total)
  total_sheet.write(row, 3, tech_total + business_total, overall_format)
  post_bag_total = timecards.post_bag_track.total.get(name, 0.0)
  post_bag_style = (green_total if post_bag_total >= postbag_required
                    else black_total)
  total_sheet.write(row, 4, post_bag_total, post_bag_style)
  total_sheet.write(row, 5,
                    post_bag_total + business_total
                      + tech_total, post_bag_style)
  preseason_total = timecards.preseason_track.total.get(name, 0.0)
  total_sheet.write(row, 6, preseason_total, overall_format)

# print out the breakdown of hours per week
row += 5
weeks = sorted(set([week for track in timecards.tracks.values()
                         for week in track.byWeek.keys()]))
for week in weeks:
  row += 1
  total_sheet.write(row, 0, 'Week %d' % week)
  tech = timecards.tech_track.byWeek.get(week, 0)
  business = timecards.business_track.byWeek.get(week, 0)
  preseason = timecards.preseason_track.byWeek.get(week, 0)
  post_bag = timecards.post_bag_track.byWeek.get(week, 0)
  total_sheet.write(row, 1, tech, black_total)
  total_sheet.write(row, 2, business, black_total)
  total_sheet.write(row, 3, tech + business, black_total)
  total_sheet.write(row, 4, post_bag, black_total)
  total_sheet.write(row, 5, tech + business + post_bag, black_total)
  total_sheet.write(row, 6, preseason, black_total)
  
row += 1
total_sheet.write(row, 0, 'Total')
columnNames = "ABCDEFG"
for col in range(1, 7):
  total_sheet.write(row, col,
                    '=SUM(%s%d:%s%d)' % (columnNames[col],
                                         row - len(weeks) + 1,
                                         columnNames[col], row),
                    black_total)


warn_sheet = workbook.add_worksheet('Warnings')
warn_sheet.write(0, 0, 'Level')
warn_sheet.write(0, 1, 'Name')
warn_sheet.set_column(1, 1, 20)
warn_sheet.write(0, 2, 'Date')
warn_sheet.write(0, 3, 'Track')
warn_sheet.write(0, 4, 'Warning')
warn_sheet.set_column(4, 4, 60)
row = 0
for (level, name, date, track, msg) in timecards.warnings:
   row += 1
   warn_sheet.write(row, 0, level)
   warn_sheet.write(row, 1, name)
   warn_sheet.write(row, 2, date, format_date)
   warn_sheet.write(row, 3, track)
   warn_sheet.write(row, 4, msg)

workbook.close()
