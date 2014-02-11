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
import vobject as vo

from pprint import pprint

from group import Group

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"


def orn(i):
    if i is None or i == 'None':
        return ""
    return i


def event2ical(event, i):
    e = i.add('vevent')
    
    pprint(str(event))
    e.add('summary').value = orn(event['title'])
    e.add('description').value = orn(event['notesforparents'])

    if event.start_time != datetime.time(0, 0, 0):
        e.add('dtstart').value = datetime.datetime.combine(event.meeting_date,
                                                           event.start_time)

    if event.end_time != datetime.time(0, 0, 0):
        e.add('dtend').value = datetime.datetime.combine(event.meeting_date,
                                                         event.end_time)

    print("{!r}".format(event.end_time))


def _main(osm, auth, sections, outdir):

    assert os.path.exists(outdir) and os.path.isdir(outdir)

    for section in sections:
        assert section in Group.SECTIONIDS.keys(), \
            "section must be in {!r}.".format(Group.SECTIONIDS.keys())

    osm_sections = osm.OSM(auth, Group.SECTIONIDS.values())

    for section in sections:
        i = vo.iCalendar()
        i.add('calscale').value = "GREGORIAN"
        i.add('X-WR-TIMEZONE').value = "Europe/London"
        [event2ical(event, i) for
         event in osm_sections.sections[
             Group.SECTIONIDS[section]].programme.events_by_date()]

        open(os.path.join(outdir, section + ".ical"),
             'w').writelines(i.serialize())

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























