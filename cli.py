# coding=utf-8
"""OSM Command Line

Usage:
   cli [options] <apiid> <token> <section> census list
   cli [options] <apiid> <token> <section> contacts list
   cli [options] <apiid> <token> <section> events list
   cli [options] <apiid> <token> <section> events <event> attendees
   cli [options] <apiid> <token> <section> events <event> info

Options:
   -a, --attending  Only list those that are attending.
   -c, --csv        Output in CSV format.
   --no_headers     Exclude headers from tables.

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


def census_list(osm, auth, section, term=None, csv=False, attending_only=False,
                     no_headers=False):

    group = Group(osm, auth, MAPPING.keys(), term)

    section_map = {'Garrick': 'Beavers',
                   'Paget': 'Beavers',
                   'Swinfen': 'Beavers',
                   'Maclean': 'Cubs',
                   'Somers': 'Cubs',
                   'Rowallan': 'Cubs',
                   'Erasmus': 'Scouts',
                   'Boswell': 'Scouts',
                   'Johnson': 'Scouts'}

    rows = []

    def add_row(section, member):
        rows.append([section_map[section], section, member['first_name'], member['last_name'],
                     member['date_of_birth'],
                     member['contact_primary_member.address1'],
                     member['contact_primary_1.address1'],
                     member['contact_primary_2.address1'],
                     member['floating.gender'].lower()])

    all_members_dict = group.all_yp_members_without_senior_duplicates_dict()

    for section in ('Swinfen', 'Paget', 'Garrick'):
        members = all_members_dict[section]
        for member in members:
            age = member.age().days / 365
            if (age > 5 and age < 9):
                add_row(section, member)
            else:
                log.info("Excluding: {} {} because not of Beaver age ({}).".format(
                    member['first_name'], member['last_name'], age
                ))

    for section in ('Maclean', 'Rowallan', 'Somers'):
        members = all_members_dict[section]
        for member in members:
            age = member.age().days / 365
            if (age > 7 and age < 11):
                add_row(section, member)
            else:
                log.info("Excluding: {} {} because not of Cub age ({}).".format(
                    member['first_name'], member['last_name'], age
                ))

    for section in ('Johnson', 'Boswell', 'Erasmus'):
        members = all_members_dict[section]
        for member in members:
            age = member.age().days / 365
            if (age > 10 and age < 16):
                add_row(section, member)
            else:
                log.info("Excluding: {} {} because not of Scout age ({}).".format(
                    member['first_name'], member['last_name'], age
                ))


    headers = ["Section", "Section Name", "First", "Last", "DOB", "Address1", "Address2", "Address3", "Gender"]

    if csv:
        w = csv_writer(sys.stdout)
        if not no_headers:
            w.writerow(list(headers))
        w.writerows(rows)
    else:
        if not no_headers:
            print(tabulate.tabulate(rows, headers=headers))
        else:
            print(tabulate.tabulate(rows, tablefmt="plain"))


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


def events_attendees(osm, auth, section, event,
                     term=None, csv=False, attending_only=False,
                     no_headers=False):
    group = Group(osm, auth, MAPPING.keys(), term)
    section_ = group._sections.sections[Group.SECTIONIDS[section]]
    ev = section_.events.get_by_name(event)
    attendees = ev.attendees
    mapping = ev.fieldmap
    if attending_only:
        attendees = [attendee for attendee in attendees
                     if attendee['attending'] == "Yes"]

    extra_fields = {
        'patrol': 'Six',
        'age': 'Age',
    }

    def fields(attendee):
        out = [str(attendee[_[1]]) for _ in mapping] + \
              [section_.members.get_by_event_attendee(attendee)[_] for _ in
               extra_fields.keys()]
        return out

    output = [fields(attendee)
              for attendee in attendees]
    headers = [_[0] for _ in mapping] + list(extra_fields.values())
    if csv:
        w = csv_writer(sys.stdout)
        if not no_headers:
            w.writerow(list(headers))
        w.writerows(output)
    else:
        if not no_headers:
            print(tabulate.tabulate(output, headers=headers))
        else:
            print(tabulate.tabulate(output, tablefmt="plain"))


if __name__ == '__main__':
    level = logging.INFO

    logging.basicConfig(level=level)

    args = docopt(__doc__, version='OSM 2.0')

    assert args['<section>'] in list(Group.SECTIONIDS.keys())+['Group'], \
        "section must be in {!r}.".format(list(Group.SECTIONIDS.keys())+['Group'])

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    if args['events']:
        if args['list']:
            events_list(osm, auth, args['<section>'])
        elif args['attendees']:
            events_attendees(osm, auth, args['<section>'],
                             args['<event>'],
                             csv=args['--csv'],
                             attending_only=args['--attending'],
                             no_headers=args['--no_headers'])
        elif args['info']:
            events_info(osm, auth, args['<section>'], args['<event>'])
        else:
            log.error('unknown')
    elif args['contacts']:
        if args['list']:
            contacts_list(osm, auth, args['<section>'])
        else:
            log.error('unknown')
    elif args['census']:
        if args['list']:
            census_list(osm, auth, args['<section>'],
                        csv=args['--csv'],
                        no_headers=args['--no_headers'])
        else:
            log.error('unknown')
    else:
        log.error('unknown')
        
        
