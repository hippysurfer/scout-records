# coding=utf-8
"""OSM Command Line

Usage:
   cli [options] <apiid> <token> census list
   cli [options] <apiid> <token> census yl list
   cli [options] <apiid> <token> census leavers
   cli [options] <apiid> <token> <section> movers list
   cli [options] <apiid> <token> <section> contacts list
   cli [options] <apiid> <token> <section> events list
   cli [options] <apiid> <token> <section> events <event> attendees
   cli [options] <apiid> <token> <section> events <event> info
   cli [options] <apiid> <token> <section> users list
   cli [options] <apiid> <token> <section> members badges
   cli [options] <apiid> <token> member badges <firstname> <lastname>
   cli [options] <apiid> <token> <section> payments <start> <end>

Options:
   -a, --attending       Only list those that are attending.
   -c, --csv             Output in CSV format.
   --no_headers          Exclude headers from tables.
   -t term, --term=term  Term to use
   -m age, --minage=age  Filter by age (decimal float).

"""

import logging

import datetime

log = logging.getLogger(__name__)

from dateutil import relativedelta
from docopt import docopt
import osm
import tabulate
from csv import writer as csv_writer
import sys

from group import Group
from update import MAPPING

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"


def census_list(osm, auth, term=None, csv=False, attending_only=False,
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


def census_yl_list(osm, auth, term=None, csv=False,
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

    for section in Group.YP_SECTIONS:
        yls = group.section_yl_members(section)
        for member in yls:
            add_row(section, member)

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


def census_leavers(osm, auth, term=None, csv=False,
                   no_headers=False):
    # Nasty hack - but I need a list of terms.
    somers_terms = Group(osm, auth, MAPPING.keys(), None)._sections.sections['20706'].get_terms()

    def find_term(name):
        return [_ for _ in somers_terms if _['name'] == name][0]

    terms = [find_term(_) for _ in
             ['Summer 2014',
              'Autumn 2014',
              'Spring 2015',
              'Summer 2015',
              'Autumn 2015',
              'Spring 2016']]

    pairs =[(terms[x],terms[x+1]) for x in range(len(terms)-1)]

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
    for old, new in pairs:
        old_term = Group(osm, auth, MAPPING.keys(), old['name'])
        new_term = Group(osm, auth, MAPPING.keys(), new['name'])

        old_members_raw = old_term.all_yp_members_without_senior_duplicates()
        new_members_raw = new_term.all_yp_members_without_senior_duplicates()

        old_members = [(_['first_name'], _['last_name'])
                       for _ in old_members_raw]

        new_members = [(_['first_name'], _['last_name'])
                       for _ in new_members_raw]

        missing = [_ for _ in old_members if not new_members.count(_)]

        for first,last in missing:
            sections = old_term.find_sections_by_name(first, last)
            member = old_members_raw[old_members.index((first, last))]
            age = member.age(ref_date=old.enddate).days // 365
            rows.append([old['name'],section_map[sections[0]],sections[0],first,last,age,member['date_of_birth'],member['floating.gender'].lower()])

    headers = ["Last Term", "Section", "Section Name", "First", "Last", "Age", "DOB", "Gender"]

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


def movers_list(osm, auth, section, age=None, term=None,
                csv=False, no_headers=False):
    group = Group(osm, auth, MAPPING.keys(), term)
    section_ = group._sections.sections[Group.SECTIONIDS[section]]

    headers = ['firstname', 'lastname', 'real_age', 'dob',
               "Date Parents Contacted", "Parents Preference",
               "Date Leaders Contacted", "Agreed Section",
               "Starting Date", "Leaving Date", "Notes", "Priority"]

    movers = section_.movers


    if age:
        threshold = (365*float(age))
        now = datetime.datetime.now()
        age_fn = lambda dob: (now - datetime.datetime.strptime(dob,'%Y-%m-%d')).days

        movers = [mover for mover in section_.movers
                  if age_fn(mover['dob']) > threshold]

    now = datetime.datetime.now()
    for mover in movers:
        real_dob = datetime.datetime.strptime(mover['dob'],'%Y-%m-%d')
        rel_age = relativedelta.relativedelta(now, real_dob)
        mover['real_age'] = "{}.{}".format(rel_age.years, rel_age.months)

    rows = [[section_['sectionname']] +[member[header] for header in headers]
            for member in movers]

    headers = ["Current Section"] + headers

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
    if not ev:
        log.error("No such event: {}".format(event))
        sys.exit(0)
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
              for attendee in attendees if section_.members.is_member(attendee['scoutid'])]
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


def users_list(osm, auth, section, csv=False, no_headers=False, term=None):
    group = Group(osm, auth, MAPPING.keys(), term)

    for user in group._sections.sections[Group.SECTIONIDS[section]].users:
        print(user['firstname'])

def members_badges(osm, auth, section, csv=False, no_headers=False, term=None):
    group = Group(osm, auth, MAPPING.keys(), term)

    #members = group._sections.sections[Group.SECTIONIDS[section]].members
    members = group.section_yp_members_without_leaders(section)
    rows = []
    for member in members:
        badges = member.get_badges(section_type=group.SECTION_TYPE[section])
        if badges:
            # If no badges - probably a leader
            challenge_new = len([badge for badge in badges
                                 if badge['awarded'] == '1' and badge['badge_group'] == '1'
                                 and not badge['badge'].endswith('(Pre 2015)')])
            challenge_old = len([badge for badge in badges
                                 if badge['awarded'] == '1' and badge['badge_group'] == '1'
                                 and badge['badge'].endswith('(Pre 2015)')])

            activity = len([badge for badge in badges if badge['awarded'] == '1' and badge['badge_group'] == '2'])
            staged = len([badge for badge in badges if badge['awarded'] == '1' and badge['badge_group'] == '3'])
            core = len([badge for badge in badges if badge['awarded'] == '1' and badge['badge_group'] == '4'])

            rows.append([member['date_of_birth'], member['last_name'], member['age'], section,
                         challenge_new, challenge_old, activity, staged, core])

    headers = ["DOB", "Last Name","Age", "Section Name", "Challenge", "Challenge_old", "Staged", "Activity", "Core"]

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


def member_badges(osm, auth, firstname, lastname, csv=False, no_headers=False, term=None):
    group = Group(osm, auth, MAPPING.keys(), term)

    members = group.find_by_name(firstname, lastname)
    #member = members[-1]
    rows = []
    for member in members:
        for section_type in ('beavers', 'cubs', 'scouts'):
            try:
                badges = member.get_badges(section_type=section_type)
                if badges is not None:
                    for badge in [_ for _ in badges if _['awarded'] == '1']:
                        rows.append([member['date_of_birth'], member['last_name'],
                                     member['age'], section_type, member._section['sectionname'],
                                     badge['badge'], datetime.date.fromtimestamp(int(badge['awarded_date'])).isoformat()])
            except:
                import traceback
                traceback.print_exc()
                pass

    headers = ["DOB", "Last Name", "Age", "Section Type", "Section Name", "Badge"]

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

def payments(osm, auth, section, start, end):
    group = Group(osm, auth, MAPPING.keys(), None)

    osm_section = group._sections.sections[Group.SECTIONIDS[section]]
    payments = osm_section.get_payments(start, end)

    print(payments.content.decode())

if __name__ == '__main__':
    level = logging.INFO

    logging.basicConfig(level=level)

    args = docopt(__doc__, version='OSM 2.0')

    if args['<section>']:
        assert args['<section>'] in list(Group.SECTIONIDS.keys()) + ['Group'], \
            "section must be in {!r}.".format(list(Group.SECTIONIDS.keys()) + ['Group'])

    term = args['--term'] if args['--term'] else None

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
    elif args['movers']:
        if args['list']:
            movers_list(osm, auth, args['<section>'],
                        age=args['--minage'],
                        csv=args['--csv'],
                        no_headers=args['--no_headers'])
        else:
            log.error('unknown')
    elif args['census']:
        if (args['yl'] and args['list']):
            census_yl_list(osm, auth,
                           csv=args['--csv'],
                           no_headers=args['--no_headers'])

        elif args['leavers']:
            census_leavers(osm, auth,
                        csv=args['--csv'],
                        no_headers=args['--no_headers'])


        elif args['list']:
            census_list(osm, auth,
                        term=args['--term'],
                        csv=args['--csv'],
                        no_headers=args['--no_headers'])

        else:
            log.error('unknown')
    elif args['users']:
        if args['list']:
            users_list(osm, auth, args['<section>'],
                       csv=args['--csv'],
                       no_headers=args['--no_headers'])
        else:
            log.error('unknown')
    elif args['members']:
        if args['badges']:
            members_badges(osm, auth, args['<section>'],
                       csv=args['--csv'],
                       no_headers=args['--no_headers'])
        else:
            log.error('unknown')
    elif args['member']:
        if args['badges']:
            member_badges(osm, auth, args['<firstname>'],
                           args['<lastname>'],
                           csv=args['--csv'],
                           no_headers=args['--no_headers'])
        else:
            log.error('unknown')

    elif args['payments']:
            payments(osm, auth, args['<section>'], args['<start>'], args['<end>'])
    else:
        log.error('unknown')
