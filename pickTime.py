#!/usr/bin/env python3

# This script is for when I need to make group edits
# where everyone transitions from technical to business
# at 2:30. It lets me find everyone who is checked in on
# a given track at a given time.

# ./pickTime Technical 01/07/2017 14:30:00

import scanners
import sys
import yaml

# read configuration from config.yaml file
config = yaml.load(open("config.yaml", "r"))
config['mangleNames'] = False
date = scanners.parseDateTime(sys.argv[2], sys.argv[3])

timecards = scanners.Timecards(config)
track = timecards.tracks[sys.argv[1]]
for name in track.namesAtTime(date):
  print(name)
