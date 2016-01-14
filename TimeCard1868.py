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
from operator import itemgetter

# default file names
scanfile = list()
outfile = 'timecard.xlsx'
startdate = "01/01/2014"
enddate = "06/01/2015"
filecnt = 0

#print ("Arguments passed:", len(sys.argv), str(sys.argv))
if len(sys.argv) > 1 :      # arguments passed
    i = 1
    while i < len(sys.argv):
        if sys.argv[i] == '-o':
            outfile = sys.argv[i+1]
            i = i+2
        elif sys.argv[i] == '-s':
            startdate = sys.argv[i+1]
            i = i+2
        elif sys.argv[i] == '-e':
            enddate = sys.argv[i+1]          
            i = i+2
        else :
            scanfile.append(sys.argv[i])
            i = i+1 ;
            filecnt = filecnt+1 ;
print (filecnt, 'Input Files :', scanfile, 'Output: ', outfile)

s_date = datetime.datetime.strptime(startdate, '%m/%d/%Y').date()
e_date = datetime.datetime.strptime(enddate, '%m/%d/%Y').date()

names = list()
timecard = list()
times = list()
dates = list()

#Now read in the scanner files - one file at a time
for file in scanfile:
    print ('Reading file', file)
    with open(file, 'rt') as inputfile :
        reader = csv.reader(inputfile, delimiter=',', quotechar='|')
        for row in reader:
            if len(row) > 0 and row[0].startswith('#') is False:
                # print row
                if row[0] not in names:
                    names.append(row[0])
                    timecard.append(dict())
                    times.append(list())
                index = names.index(row[0])
                d = datetime.datetime.strptime(row[3], '%m/%d/%Y')
                t = datetime.datetime.strptime(row[2], '%H:%M:%S')
                dt = datetime.datetime.combine(d.date(), t.time())
                times[index].append(dt)
                if row[3] not in timecard[index] :
                    timecard[index][row[3]] = list()
                timecard[index][row[3]].append(dt)
                # put into master date list for the Col. headings list
                if row[3] not in dates :
                    dates.append(row[3])


print ("Total: ", len(names), 'names', 'with ', len(dates), 'days attended')
print ("Generating report from:", startdate, "to: ", enddate)
#print debug section
#for i in range(0, len(names)):
#    print 'name:', names[i], 'has', len(times[i]), 'time entries'
#    times[i].sort()
#    for j in range(0, len(times[i])):
#        print times[i][j].isoformat(' ')

# Now prep the xlsx workbook
workbook  = xlsxwriter.Workbook(outfile)
sheet = workbook.add_worksheet('This Week')
format_date = workbook.add_format({'num_format': 'mm/dd/yy'})
cell_format = workbook.add_format()
cell_format.set_bg_color('red')
format_num = workbook.add_format({'num_format':'0.000'})
dates.sort()
row = 0
col = 0
sheet.write(0, col, 'Name')
col = col+1
for j in range(len(dates)):
    sheet.write(0, col, dates[j])
    col = col+1

# Now all the scanned files have been read into a giant list of lists
# The main list - which is indexed by name has one list per name, and each item
# in the list containts a date/time pair. - times[i]
# For now - timcard[i][j] is a dictionary of times - which doesnt work to index over a day boundary
# so we use times - where for each name you can just index the entire timestamp list
# First sort the time data for each name and then
# process by Looping through each names time data
# for each day -
for i in range(len(names)):
    row = row + 1 # increment to next row in the sheet
    col = 0
    sheet.write(row, 0, names[i])
    #print ('name:', names[i], 'has', len(times[i]), 'time entries')
    #print times[i]
    times[i].sort()
    # take the sorted list of times and filter all stamps within 60 seconds of each other
    k = 0
    while k < (len(times[i])-1):
        tdelta = times[i][k+1] - times[i][k]
        if (tdelta.total_seconds() <= 60) :
            print('Deleting Duplicated entry for ', names[i], 'at', times[i][k])
            del times[i][k]
        else :
            k = k+1
    #print ('name:', names[i], 'has', len(times[i]), 'time entries')
    if len(times[i]) == 1 :
         print ('WARNING! - ONLY ONE ENTRY FOR: ', names[i], ' ON: ', times[i][0])
         print ('Manual processing needed')
    #print times[i]
    k = 0
    while k < len(times[i])-1 :
        curr_date = times[i][k].date() # current record date
        # only for days in the daterange
        j = curr_date.strftime('%m/%d/%Y')
        if (curr_date < s_date or curr_date  > e_date ) :
            #print 'Skipping date'
            k = k+1
            continue
        end_k = k
        while end_k < len(times[i]) and times[i][k].date() == times[i][end_k].date() :
            end_k = end_k + 1
        # check for one entry on next date with time before 4.00am
        if end_k < len(times[i]) and ( (times[i][end_k].date() - times[i][k].date()).days == 1 ) and (times[i][end_k].time() < datetime.time(4, 0, 0) ) :
            print ('WARNING OVERNIGHT TIMESTAMP ADJUSTMENT FOR ', names[i], 'ON:', times[i][end_k])
            end_k = end_k + 1 #  timestamp past midnight but before 4.00am belongs to previous day          
        # scan for new date - to determin end_k
        # - now we have the beginning and ending index for a given date.
        # Odd number of entries check here
        total = 0
        while k < end_k-1 :
            tdelta = times[i][k+1] - times[i][k]
            total += tdelta.total_seconds()
            k = k+2

        # now check for singletons, two entry timestamps and one exit or one entry and two exits
        # 30 minute is qualifier for entry exit dual stamps
        #
        if (end_k-k == 1 and total == 0) :
            print ('WARNING! - TRUE SINGLE ENTRY ONLY FOR: ', names[i], ' ON: ', times[i][k])
        elif (end_k-k == 1) :
            # true odd entry means k, k-1 and k-2 are all valid indices
            first_pair = times[i][k-1] - times[i][k-2]
            second_pair = times[i][k] - times[i][k-1]
            if first_pair.total_seconds() < 1800 or second_pair.total_seconds() < 1800 :
                # just add the singleton time
                print ('WARNING! - DUAL ENTRY TIMESTAMP FOR: ', names[i], ' ON: ', times[i][k-1])
                print ('WARNING! - ADDING TIME FOR: ', names[i], ' AT: ', times[i][k])
                total += second_pair.total_seconds()
               
        #print ('Time for : ', names[i], 'on :', j, 'is: ', (total/3600), 'hrs')

        # now increment to a new day
        k = end_k

        # Now format this for the spread sheet - we are on the right row, we have to
        # find the column for the date since not all girls work on all days
        if j not in dates :
            print ('FATAL ERROR: Time card entry not in master dates list!!!!!')
            exit()
        # we have a date entry - remember first column is the name
        # j is the date in iso string format '%m/%d/%Y'
        index = dates.index(j) + 1
        if total > 0 :
            sheet.write(row, index, (total/3600), format_num)
        else :
            sheet.write(row, index, 0, cell_format);
        #print timecard[i][j]

workbook.close()
