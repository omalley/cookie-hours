# Time card application for FRC Team 1868

This script reads the files from the timecard scanners that are used in
the lab, and produces a spreadsheet with each girl's hours. We upload the
sheet to Google Sheets for sharing with the girls.

The script breaks the time in to three tracks:
* Business - working on non-technical stuff
* Technical - working on the robot
* Post-Bag - working after bag and tag

To run for the next year, configure the settings for the season in
config.yaml. For 2017, config.yaml looked like:

    # Start and end of the season.
    # The script ignores any events outside of this range.
    startDate: 01/01/2017
    endDate: 05/01/2017
    
    # Date of bag and tag
    bagDate: 02/21/2017
    
    # The code for the business scanner
    businessScanner: "105059"
    
    # The root directory for this years data files
    dataRoot: data/2017

    # The output filename
    output: timecard.xlsx

    # The list of tracks and the minimum hours in each
    hours:
      Technical: 90.0
      Business: 10.0
      Post-Bag: 16.0

I keep all of the data files in a directory structure like
`data/<year>/<dump-date>/*.TXT`, but that is just a convention. The
script looks for all files named *.TXT under the `dataRoot`
directory. Don't check the data files into public github, since the
girls' names are there.

When you run the script, it should look like:

    owen@laptop> ./runTimes.py
    Reading file data/2017/12-06/archi.TXT
    Reading file data/2017/12-06/curie.TXT
    Reading file data/2017/12-06/galileo.TXT
    Reading file data/2017/12-06/newton.TXT
    Total: 84 names with 8 technical, 5 business, and 0 post-bag days
    Generating report timecard.xlsx from: 2016-09-01 to: 2017-05-01

Upload the file to Google Sheets using "File/Import/Upload/Replace".
After you upload, run the "Cookies/titles" macro to set the title bars.

When you need to make fixes to the data, I've added a file named
manual.yaml that is located in the `dataRoot` directory. It looks like:

    Technical:
      01/07/2017:
        M MOUSE: 7.5
        D DUCK: 5
      01/10/2017:
        S WHITE: 2.5
    Business:
    Post-Bag:

That would replace M Mouse's and D Duck's technical hours on 1/7 with
7.5 and 5, respectively. The next two lines set S White's technical
hours on 1/10 to 2.5. Indentation is important in this file and it
should always be:

    <track-name>:
      <date>:
        <name>: <hours>

I've been running this script in Python 3.5 on a Mac using MacPorts.
You'll need to install python35, py35-pip, and py35-readline.
You'll need to pip install XlsxWriter, and PyYAML.
