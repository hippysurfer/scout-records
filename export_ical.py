# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  export_ical.py [-d] <apiid> <token> <outdir> <section>... 
  export_ical.py (-h | --help)
  export_ical.py --version


Options:
  <section>      Section to export.
  <outdir>       Output directory for ical files.
  -d,--debug     Turn on debug output.
  -h,--help      Show this screen.
  --version      Show version.

"""

import os.path
import logging
from docopt import docopt
import datetime
import osm
from icalendar import Calendar, Event

NOW = datetime.datetime.now()

from pprint import pprint

from group import Group

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"


def orn(i):
    if i is None or i == 'None':
        return ""
    return i


def meeting2ical(section, event, i):
    e = Event()

    e['summary'] = "{}: {}".format(section, orn(event['title']))
    e['description'] = orn(event['notesforparents'])

    if event.start_time.time() != datetime.time(0, 0, 0):
        e.add('dtstart', event.start_time.astimezone())
    else:
        e.add('dtstart', event.start_time.date())

    if event.end_time.time() != datetime.time(0, 0, 0):
        e.add('dtend', event.end_time.astimezone())
    else:
        e.add('dtend', event.end_time.date())

    e.add('dtstamp', NOW)

    i.add_component(e)


def event2ical(section, event, i):
    e = Event()

    e['summary'] = "{}: {}".format(section, orn(event['name']))
    e['location'] = orn(event['location'])

    if event.start_time.time() != datetime.time(0, 0, 0):
        e.add('dtstart', event.start_time.astimezone())
    else:
        e.add('dtstart', event.start_time.date())

    if event.end_time.time() != datetime.time(0, 0, 0):
        e.add('dtend', event.end_time.astimezone())
    else:
        e.add('dtend', event.end_time.date())

    e.add('dtstamp', NOW)

    i.add_component(e)


def _main(osm, auth, sections, outdir):

    assert os.path.exists(outdir) and os.path.isdir(outdir)

    for section in sections:
        assert section in list(Group.SECTIONIDS.keys()) + ['Group'], \
            "section must be in {!r}.".format(Group.SECTIONIDS.keys())

    osm_sections = osm.OSM(auth, Group.SECTIONIDS.values())

    for section in sections:

        if section == "Group":
            i = Calendar()
            i['x-wr-calname'] = '7th Lichfield: Group Calendar'
            i['X-WR-CALDESC'] = 'Current Programme'
            i['calscale'] = 'GREGORIAN'
            i['X-WR-TIMEZONE'] = 'Europe/London'

            for s in Group.SECTIONIDS.keys():
                section_obj = osm_sections.sections[Group.SECTIONIDS[s]]

                [meeting2ical(s, event, i) for
                    event in section_obj.programme.events_by_date()]

                [event2ical(s, event, i) for
                    event in section_obj.events]

            open(os.path.join(outdir, section + ".ical"),
                 'w').write(i.to_ical().decode())

        else:
            section_obj = osm_sections.sections[Group.SECTIONIDS[section]]

            i = Calendar()
            i['x-wr-calname'] = '7th Lichfield: {}'.format(section)
            i['X-WR-CALDESC'] = '{} Programme'.format(section_obj.term['name'])
            i['calscale'] = 'GREGORIAN'
            i['X-WR-TIMEZONE'] = 'Europe/London'

            [meeting2ical(section, event, i) for
             event in section_obj.programme.events_by_date()]

            [event2ical(section, event, i) for
             event in section_obj.events]

            open(os.path.join(outdir, section + ".ical"),
                 'w').write(i.to_ical().decode())

if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.INFO

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    _main(osm, auth, args['<section>'], args['<outdir>'])























