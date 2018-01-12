#!/usr/bin/env python3

# Time card applicaiton for FRC Team 1868

# This is the library that I use to parse and understand
# the file from the scanners and build the data structures
# that I need for the timecard reports.

# Works with Opticon scanner in USB MSD mode (C04)
# and with fields programmed to be:
#  * name - text field - bar code is a text string
#  * serial - numerical - scanner seried number
#  * time - time code in 24HR HH:MM:SS format
#  * date - date code in MM/DD/YYYY format.

import csv
import datetime
import functools
import glob
import operator
import os.path
import sys
import yaml

# ignore events less than 2 minutes apart
MIN_SEPARATION = 120

def parseDate(str):
  return datetime.datetime.strptime(str, '%m/%d/%Y').date()

# parse dates with long times (HH:MM:SS)
def parseDateTime(date, time):
  d = parseDate(date)
  t = datetime.datetime.strptime(time, '%H:%M:%S')
  return datetime.datetime.combine(d, t.time())

# parse dates with short times (HH:MM)
def parseDateShortTime(date, time):
  d = parseDate(date)
  t = datetime.datetime.strptime(time, '%H:%M')
  return datetime.datetime.combine(d, t.time())

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

# rearrange the name so that it sorts by last name
def mangleName(name):
  [first, last] = name.split(None, 1)
  return "%s, %s" % (last, first)

# training classes, which are tracked separately from their track
class Event:
  def __init__(self, name, times_str):
    self.name = name
    [date_str, start_str, stop_str] = times_str.split()
    self.start = parseDateShortTime(date_str, start_str)
    self.stop = parseDateShortTime(date_str, stop_str)

  def __repr__(self):
    return "%s from %s to %s" % (self.name, self.start, self.stop)

# Parse the list of trainings from the config file.
# map(name, list(times))
def parseEvents(event_list):
  result = []
  for name, times in event_list.items():
    for time in times:
      result.append(Event(name, time))
  return result

# Stores the scans of a single student on the same day
class DayReport:
   def __init__(self):
      self.scans = []
      self.ignored = []
      self.state = "normal"
      self.manual = 0
      # map(eventName, hours)
      self.training = {}

   # after the data is loaded, fix up the data
   def fixUp(self, name, date, track, trainings, warnings):
      self.scans.sort()
      i = 0
      while i < len(self.scans) - 1:
         if (self.scans[i+1] - self.scans[i]).seconds < MIN_SEPARATION:
            self.ignored.append(self.scans[i])
            del self.scans[i]
         else:
            i += 1
      if len(self.ignored) > 0:
         warnings.append(('info', name, date, track,
                          ("%d near duplicate events ignored" %
                           len(self.ignored))))
      if len(self.scans) % 2 != 0:
         if len(self.scans) == 1:
           self.state = "error"
         msg = ("Odd number of events: " +
                ', '.join(map(lambda d: d.strftime('%H:%M'), self.scans)))
         warnings.append(('ERR' if self.state == "error" else 'WARN',
                          name, date, track, msg))
         # In the case where the student has an odd number of events,
         # either drop the first or last event depending on what gives
         # the student more hours.
         if len(self.scans) == 1:
           self.scans = []
         elif calculateHours(self.scans) < calculateHours(self.scans[1:]):
           self.scans = self.scans[1:]
         else:
           self.scans = self.scans[:-1]
      for event in trainings:
        eventHours = self.eventHours(event)
        if eventHours > 0:
          self.training[event.name] = \
            self.training.get(event.name, 0) + eventHours

   def append(self, time):
      self.scans.append(time)

   # manually override the hours to the given value
   def manualUpdate(self, hours):
      self.state = "manual"
      self.manual = hours
      self.scans = []

   # was this student checked in at this time?
   def checkedIn(self, time):
     result = False
     if self.state != "manual":
       i = 0
       while i < len(self.scans) and self.scans[i] <= time:
         i += 1
         result = not result
     return result

   def hours(self):
     if self.state == "manual":
       return self.manual
     else:
       return calculateHours(self.scans) - self.trainingHours()


   # If the person was checked in during the event, calculate the hours
   # that they were at the event.
   def eventHours(self, event):
     result = 0
     if self.state != "manual":
       i = 0
       while i < len(self.scans) and self.scans[i] < event.stop:
         checkOut = self.scans[i+1]
         if checkOut > event.start:
           checkIn = self.scans[i]
           eventCheckIn = event.start if checkIn < event.start else checkIn
           eventCheckOut = event.stop if event.stop < checkOut else checkOut
           result += (eventCheckOut - eventCheckIn).seconds / 3600
         i += 2
     return result

   def trainingHours(self):
     return functools.reduce(operator.add, self.training.values(), 0)

# All of the information about a person with in a single track
class PersonInTrack:
  def __init__(self):
    # map(date, DayReport)
    self.times = {}

  def addDate(self, date, time):
    self.times.setdefault(date, DayReport()).append(time)

  def manualUpdate(self, date, hours):
    self.times.setdefault(date, DayReport()).manualUpdate(hours)

  def eventHours(self, eventName):
    result = 0
    for date, report in self.times.items():
      result += report.training.get(eventName, 0)
    return result

class Track:
  def __init__(self, name, required_hours, trainings):
    self.name = name
    self.required_hours = required_hours
    # map(name, PersonInTrack)
    self.people = {}
    # list(date)
    self.dates = []
    # map(name, hours)
    self.total = {}
    # map(week, hours)
    self.byWeek = {}
    # list(Event)
    self.training = parseEvents(trainings)

  # Generates the list of names that were checked in to this track at the
  # given time
  def namesAtTime(self, time):
    date = adjustDate(time)
    return sorted([name for (name, person) in self.people.items()
                        if date in person.times and
                           person.times[date].checkedIn(time)])

  def trainingNames(self):
    return { e.name for e in self.training }

  def trainingHours(self, personName, eventName):
    if personName in self.people:
      return self.people[personName].eventHours(eventName)
    return 0

class Timecards:
  def __init__(self, config):
    self.tracks = {}
    for name, trackConfig in config['tracks'].items():
      self.tracks[name] = Track(name, trackConfig['hours'],
                                trackConfig.get('training', {}))
    self.preseason_track = self.tracks['Pre-season']
    self.tech_track = self.tracks['Technical']
    self.business_track = self.tracks['Business']
    self.post_bag_track = self.tracks['Post-Bag']
    self.warnings = []
    self.start_date = parseDate(config['startDate'])
    self.end_date = parseDate(config['endDate'])
    self.kick_date = parseDate(config['kickOff'])
    self.bag_date = parseDate(config['bagDate'])
    self.data_root = config['dataRoot']
    self.mangle_names = config.get('mangleNames', True)
    self.business_scanner = config['businessScanner']
    self.readScanners(self.data_root)
    self.readOverrides(os.path.join(self.data_root, "manual.yaml"))
    self.fixup()
    self.warnings.sort()
    if len(self.post_bag_track.dates) > 0:
      self.post_bag_days = (self.post_bag_track.dates[0] - self.bag_date).days
    else:
      self.post_bag_days = 0

  def readScanners(self, data_root):
    for file in [y for x in os.walk(data_root)
                   for y in glob.glob(os.path.join(x[0], '*.TXT'))]:
      print ('Reading file', file)
      with open(file, 'rt') as inputfile :
        reader = csv.reader(inputfile, delimiter=',', quotechar='|')
        for row in reader:
          if len(row) > 0 and not row[0].startswith('#'):
            dt = parseDateTime(row[3], row[2])
            day = adjustDate(dt)
            if self.start_date <= day and day <= self.end_date:
              if day < self.kick_date:
                track = self.preseason_track
              elif day > self.bag_date:
                track = self.post_bag_track
              elif row[1] == self.business_scanner:
                track = self.business_track
              else:
                track = self.tech_track
              if self.mangle_names:
                name = mangleName(row[0])
              else:
                name = row[0]
              person = track.people.setdefault(name, PersonInTrack())
              person.addDate(day, dt)
              if day not in track.dates :
                track.dates.append(day)

  def names(self):
    return sorted(set([name for track in self.tracks.values()
                            for name in track.people.keys()]))

  # Read the manual updates from <dataRoot>/manual.yaml
  # It should look like:
  # <track name>:
  #   <date>:
  #     <name>: <hours>
  # For each entry, overrides any checkins on that date
  def readOverrides(self, filename):
    if os.path.isfile(filename):
      manualUpdates = yaml.load(open(filename, "r"))
      for (trackName, dateList) in manualUpdates.items():
        track = self.tracks[trackName]
        if dateList:
          for dateStr in dateList:
            day = parseDate(dateStr)
            for (rawName, hours) in dateList[dateStr].items():
              name = mangleName(rawName)
              person = track.people.setdefault(name, PersonInTrack())
              person.manualUpdate(day, hours)
              if day not in track.dates :
                track.dates.append(day)

  def fixup(self):
    for track in self.tracks.values():
      track.dates.sort(reverse=True)
      for (name, person) in track.people.items():
        for (date, report) in person.times.items():
          report.fixUp(name, date, track.name, track.training, self.warnings)
          week = (date - self.kick_date).days // 7
          track.byWeek[week] = track.byWeek.get(week, 0) + report.hours() + \
                               report.trainingHours()

  def printSummary(self):
    print ("Dates: start:", self.start_date, ", end:", self.end_date,
           ", bag:", self.bag_date)
    print ("Total of", len(self.names()), 'names with',
           len(self.preseason_track.dates), 'pre-season,',
           len(self.tech_track.dates), 'technical,',
           len(self.business_track.dates), 'business, and',
           self.post_bag_days, 'post-bag days')
    summary = {}
    for warn in self.warnings:
      kind = warn[0]
      summary[kind] = summary.get(kind, 0) + 1

    print("Warnings:", summary)
