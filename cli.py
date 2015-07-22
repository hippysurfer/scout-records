# coding=utf-8
"""OSM Command Line

Usage:
   cli <apiid> <token> <section> contacts list
   cli <apiid> <token> <section> events list
   cli <apiid> <token> <section> events <event> attendees [--csv]
   cli <apiid> <token> <section> events <event> info

Options:
   --csv   Output in CSV format

"""

import logging
log = logging.getLogger(__name__)

from docopt import docopt
import osm
import tabulate
from csv import writer as csv_writer
import sys

from group import Group
from update import MAPPING

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"


def contacts_list(osm, auth, section, term=None):
    group = Group(osm, auth, MAPPING.keys(), term)

    for member in group.section_all_members(section):
        print("{} {}".format(member['first_name'], member['last_name']))


def events_list(osm, auth, section, term=None):
    group = Group(osm, auth, MAPPING.keys(), term)

    for event in group._sections.sections[Group.SECTIONIDS[section]].events:
        print(event['name'])


def events_info(osm, auth, section, event, term=None):
    group = Group(osm, auth, MAPPING.keys(), term)

    ev = group._sections.sections[
        Group.SECTIONIDS[section]].events.get_by_name(event)
    print(",".join([ev[_] for _ in ['name', 'startdate', 'enddate', 'location']]))


def events_attendees(osm, auth, section, event, term=None, csv=False):
    group = Group(osm, auth, MAPPING.keys(), term)

    ev = group._sections.sections[
        Group.SECTIONIDS[section]].events.get_by_name(event)
    attendees = ev.attendees
    mapping = ev.fieldmap
    output = ([str(attendee[_[1]]) for _ in mapping] for attendee in attendees)
    headers = (_[0] for _ in mapping)
    if csv:
        w = csv_writer(sys.stdout)
        w.writerow(list(headers))
        w.writerows(output)
    else:
        print(tabulate.tabulate(output, headers=headers))


if __name__ == '__main__':
    level = logging.INFO

    logging.basicConfig(level=level)

    args = docopt(__doc__, version='OSM 2.0')

    assert args['<section>'] in Group.SECTIONIDS.keys(), \
        "section must be in {!r}.".format(Group.SECTIONIDS.keys())

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    if args['events']:
        if args['list']:
            events_list(osm, auth, args['<section>'])
        elif args['attendees']:
            events_attendees(osm, auth, args['<section>'],
                             args['<event>'],
                             csv=args['--csv'])
        elif args['info']:
            events_info(osm, auth, args['<section>'], args['<event>'])
        else:
            log.error('unknown')
    elif args['contacts']:
        if args['list']:
            contacts_list(osm, auth, args['<section>'])
        else:
            log.error('unknown')
    else:
        log.error('unknown')
        
        
