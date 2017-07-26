# coding=utf-8
"""OSM Command Line

Usage:
   cli [options] <apiid> <token> census list
   cli [options] <apiid> <token> census yl list
   cli [options] <apiid> <token> census leavers
   cli [options] <apiid> <token> <section> movers list
   cli [options] <apiid> <token> <section> contacts list
   cli [options] <apiid> <token> <section> contacts details   
   cli [options] <apiid> <token> <section> events list
   cli [options] <apiid> <token> <section> events <event> attendees
   cli [options] <apiid> <token> <section> events <event> info
   cli [options] <apiid> <token> <section> users list
   cli [options] <apiid> <token> <section> members badges
   cli [options] <apiid> <token> member badges <firstname> <lastname>
   cli [options] <apiid> <token> <section> payments <start> <end>
   cli [options] <apiid> <token> group payments <outfile>


Options:
   -a, --attending       Only list those that are attending.
   -c, --csv             Output in CSV format.
   --no_headers          Exclude headers from tables.
   -t term, --term=term  Term to use
   -m age, --minage=age  Filter by age (decimal float).

"""

import logging

import datetime
import collections

from io import StringIO

log = logging.getLogger(__name__)

from dateutil import relativedelta
from docopt import docopt
import osm
import tabulate
from csv import writer as csv_writer
import sys
import pandas as pd

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
                     member['contact_primary_1.address2'],
                     member['contact_primary_1.address3'],
                     member['contact_primary_1.postcode'],
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

    headers = ["Section", "Section Name", "First", "Last", "DOB", "Address1", "Address1.1", "Address1.2", "Address1.3",
               "Address2", "Address3", "Gender"]

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
              'Spring 2016',
              'Summer 2016',
              'Autumn 2016',
              'Spring 2017',
              'Summer 2017',
              'Autumn 2017',
              ]]

    pairs = [(terms[x], terms[x + 1]) for x in range(len(terms) - 1)]

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

        for first, last in missing:
            sections = old_term.find_sections_by_name(first, last)
            member = old_members_raw[old_members.index((first, last))]
            age = member.age(ref_date=old.enddate).days // 365
            rows.append([old['name'], section_map[sections[0]], sections[0], first, last, age, member['date_of_birth'],
                         member['floating.gender'].lower()])

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


def contacts_list(osm, auth, sections, term=None):
    group = Group(osm, auth, MAPPING.keys(), term)

    for section in sections:
        for member in group.section_all_members(section):
            print("{} {}".format(member['first_name'], member['last_name']))


def contacts_detail(osm, auth, sections, csv=False, term=None, no_headers=False):
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
                     member['contact_primary_1.email1'],
                     member['contact_primary_1.address1'],
                     member['contact_primary_1.address2'],
                     member['contact_primary_1.address3'],
                     member['contact_primary_1.postcode'],
                     member['contact_primary_2.address1'],
                     member['floating.gender'].lower()])

    for section in sections:
        for member in group.section_all_members(section):
            add_row(section, member)

    headers = ["Section", "Section Name", "First", "Last", "DOB", "Email1", "Address1", "Address1.1", "Address1.2", "Address1.3",
               "Address2", "Address3", "Gender"]

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


def movers_list(osm, auth, sections, age=None, term=None,
                csv=False, no_headers=False):
    group = Group(osm, auth, MAPPING.keys(), term)

    rows = []

    for section in sections:
        section_ = group._sections.sections[Group.SECTIONIDS[section]]

        headers = ['firstname', 'lastname', 'real_age', 'dob',
                   "Date Parents Contacted", "Parents Preference",
                   "Date Leaders Contacted", "Agreed Section",
                   "Starting Date", "Leaving Date", "Notes", "Priority",
                   '8', '10 1/2', '14 1/2']

        movers = section_.movers

        if age:
            threshold = (365 * float(age))
            now = datetime.datetime.now()
            age_fn = lambda dob: (now - datetime.datetime.strptime(dob, '%Y-%m-%d')).days

            movers = [mover for mover in section_.movers
                      if age_fn(mover['dob']) > threshold]

        now = datetime.datetime.now()
        for mover in movers:
            real_dob = datetime.datetime.strptime(mover['dob'], '%Y-%m-%d')
            rel_age = relativedelta.relativedelta(now, real_dob)
            mover['real_age'] = "{0:02d}.{0:02d}".format(rel_age.years, rel_age.months)
            mover['8'] = (real_dob+relativedelta.relativedelta(years=8)).strftime("%b %y")
            mover['10 1/2'] = (real_dob + relativedelta.relativedelta(years=10, months=6)).strftime("%b %y")
            mover['14 1/2'] = (real_dob + relativedelta.relativedelta(years=14, months=6)).strftime("%b %y")

        rows += [[section_['sectionname']] +
                 [member[header] for header in headers]
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


def events_list(osm, auth, sections, term=None):
    group = Group(osm, auth, MAPPING.keys(), term)

    for section in sections:
        for event in group._sections.sections[Group.SECTIONIDS[section]].events:
            print(event['name'])


def events_info(osm, auth, sections, event, term=None):
    group = Group(osm, auth, MAPPING.keys(), term)

    for section in sections:
        ev = group._sections.sections[
            Group.SECTIONIDS[section]].events.get_by_name(event)
        print(",".join([ev[_] for _ in ['name', 'startdate', 'enddate', 'location']]))


def events_attendees(osm, auth, sections, event,
                     term=None, csv=False, attending_only=False,
                     no_headers=False):
    group = Group(osm, auth, MAPPING.keys(), term)

    for section in sections:
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


def users_list(osm, auth, sections, csv=False, no_headers=False, term=None):
    group = Group(osm, auth, MAPPING.keys(), term)

    for section in sections:
        for user in group._sections.sections[Group.SECTIONIDS[section]].users:
            print(user['firstname'])


def members_badges(osm, auth, sections, csv=False, no_headers=False, term=None):
    group = Group(osm, auth, MAPPING.keys(), term)

    for section in sections:
        # members = group._sections.sections[Group.SECTIONIDS[section]].members
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

        headers = ["DOB", "Last Name", "Age", "Section Name", "Challenge", "Challenge_old", "Staged", "Activity", "Core"]

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
    # member = members[-1]
    rows = []
    for member in members:
        for section_type in ('beavers', 'cubs', 'scouts'):
            try:
                badges = member.get_badges(section_type=section_type)
                if badges is not None:
                    for badge in [_ for _ in badges if _['awarded'] == '1']:
                        rows.append([member['date_of_birth'], member['last_name'],
                                     member['age'], section_type, member._section['sectionname'],
                                     badge['badge'],
                                     datetime.date.fromtimestamp(int(badge['awarded_date'])).isoformat()])
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


def payments(osm, auth, sections, start, end):
    group = Group(osm, auth, MAPPING.keys(), None)

    for section in sections:
        osm_section = group._sections.sections[Group.SECTIONIDS[section]]
        payments = osm_section.get_payments(start, end)

        print(payments.content.decode())


def group_payments(osm, auth, outfile):
    important_fields = ['first_name',
                        'last_name',
                        'joined',
                        'started',
                        'date_of_birth']

    # Define payment schedules that we are interested in.
    payment_schedules_list = [
        ('2016 - Spring Term - Part 1', datetime.date(2016, 1, 4)),
        ('2016 - Spring Term - Part 2', datetime.date(2016, 2, 22)),
        ('2016 - Summer Term - Part 1', datetime.date(2016, 4, 11)),
        ('2016 - Summer Term - Part 2', datetime.date(2016, 6, 6)),
        ('2016 - Autumn Term - Part 1', datetime.date(2016, 9, 5)),
        ('2016 - Autumn Term - Part 2', datetime.date(2016, 10, 31)),
        ('2017 - Spring Term - Part 1', datetime.date(2017, 1, 9)),
        ('2017 - Spring Term - Part 2', datetime.date(2017, 2, 20))
    ]

    payment_dates = collections.OrderedDict(payment_schedules_list)

    schedules = [_[0] for _ in payment_schedules_list]
    first_date = min([_[1] for _ in payment_schedules_list])
    last_date = max([_[1] for _ in payment_schedules_list])

    # Payment amounts. Assumes all schedules use the same quantity.
    general_amount = 17.95
    discount_amount = 12.13

    # Fetch all of the available data for each term on which a payment is due.
    group_by_date = {}
    for name, date in payment_dates.items():
        group_by_date[name] = Group(osm, auth, important_fields, on_date=date)

    # Get a list of all members in all terms across the whole group.
    all_yp_members = []
    for group in group_by_date.values():
        all_yp_members.extend(group.all_yp_members_without_leaders())

    all_yp_by_scout_id = {member['member_id']: member for member in all_yp_members}

    # Get the current group data for the current term.
    current = Group(osm, auth, important_fields)

    res = []
    for scoutid, member in all_yp_by_scout_id.items():
        current_member = current.find_by_scoutid_without_senior_duplicates(str(scoutid))
        section = current_member[0]._section['sectionname'] if len(current_member) else 'Unknown'
        d = collections.OrderedDict((('scoutid', scoutid), ('First name', member['first_name']),
                                     ('Last name', member['last_name']), ('joined', member['started']),
                                     ('left', member['end_date']),
                                     ('section', section)))

        joined = datetime.datetime.strptime(member['started'], "%Y-%m-%d").date()
        ended = datetime.datetime.strptime(member['end_date'], "%Y-%m-%d").date() if member['end_date'] else False

        amount = discount_amount if member['customisable_data.cf_subs_type_n_g_d_'] == 'D' else general_amount

        for schedule, date_ in payment_dates.items():
            d[schedule] = amount if ((joined < date_ and not ended) or
                                     (joined < date_ and ended > date_)) else 0
        res.append(d)

    tbl = pd.DataFrame(res)
    tbl.set_index(["Last name", "First name"], inplace=True)

    all_sections = list(Group.SECTIONIDS.keys())

    def fetch_section(section_name):
        section = group._sections.sections[Group.SECTIONIDS[section_name]]
        payments = section.get_payments(first_date.strftime('%Y-%m-%d'),
                                        last_date.strftime('%Y-%m-%d'))
        return pd.read_csv(StringIO(payments.content.decode())) if payments is not None else pd.DataFrame()

    all = pd.concat([fetch_section(name) for name in all_sections], ignore_index=True)

    all['Schedule'] = all['Schedule'].str.replace('^General Subscriptions.*$', 'General Subscriptions')
    all['Schedule'] = all['Schedule'].str.replace('^Discounted Subscriptions.*$', 'Discounted Subscriptions')
    subs = all[(all['Schedule'] == 'General Subscriptions') | (all['Schedule'] == 'Discounted Subscriptions')]

    pv = pd.pivot_table(subs, values='Net', index=['Last name', 'First name'], columns=['Schedule', 'Payment'])
    for schedule in schedules:
        pv['General Subscriptions'][schedule].fillna(
            pv['Discounted Subscriptions'][schedule], inplace=True)
    del pv['Discounted Subscriptions']
    pv.columns = pv.columns.droplevel()
    del pv['2015/Q3']

    combined = tbl.join(pv, lsuffix='_est', rsuffix='_act', how='outer')

    for schedule in schedules:
        for suffix in ['_est', '_act']:
            combined[schedule + suffix].fillna(0, inplace=True)

    for schedule in schedules:
        combined[schedule + '_var'] = combined.apply(lambda row: row[schedule + '_est'] - row[schedule + '_act'],
                                                     axis=1)

    combined.to_excel(outfile, sheet_name="Data", merge_cells=False)


if __name__ == '__main__':
    level = logging.INFO

    logging.basicConfig(level=level)

    args = docopt(__doc__, version='OSM 2.0')

    sections = None
    if args['<section>']:
        section = args['<section>']
        assert section in list(Group.SECTIONIDS.keys()) + ['Group'] + list(Group.SECTIONS_BY_TYPE.keys()), \
            "section must be in {!r}.".format(list(Group.SECTIONIDS.keys()) + ['Group'])

        sections = Group.SECTIONS_BY_TYPE[section] if section in Group.SECTIONS_BY_TYPE.keys() else [section,]

    term = args['--term'] if args['--term'] else None

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    if args['events']:
        if args['list']:
            events_list(osm, auth, sections)
        elif args['attendees']:
            events_attendees(osm, auth, sections,
                             args['<event>'],
                             csv=args['--csv'],
                             attending_only=args['--attending'],
                             no_headers=args['--no_headers'])
        elif args['info']:
            events_info(osm, auth, sections, args['<event>'])
        else:
            log.error('unknown')
    elif args['contacts']:
        if args['list']:
            contacts_list(osm, auth, sections)
        elif args['details']:
            contacts_detail(osm, auth, sections, csv=args['--csv'],
                          no_headers=args['--no_headers'])
        else:
            log.error('unknown')
    elif args['movers']:
        if args['list']:
            movers_list(osm, auth, sections,
                        age=args['--minage'],
                        csv=args['--csv'],
                        no_headers=args['--no_headers'])
        else:
            log.error('unknown')
    elif args['census']:
        if args['yl'] and args['list']:
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
            users_list(osm, auth, sections,
                       csv=args['--csv'],
                       no_headers=args['--no_headers'])
        else:
            log.error('unknown')
    elif args['members']:
        if args['badges']:
            members_badges(osm, auth, sections,
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

    elif args['group']:
        if args['payments']:
            group_payments(osm, auth, args['<outfile>'])
        else:
            log.error('unknown')

    elif args['payments']:
        payments(osm, auth, sections, args['<start>'], args['<end>'])
    else:
        log.error('unknown')
