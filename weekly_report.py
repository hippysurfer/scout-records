# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  weekly_report.py [-d | --debug] [-n | --no_email] [--email=<email>] [-w | --web]
                   [--quarter=<quarter>] [--term=<term>] <apiid> <token> <section>...
  weekly_report.py (-h | --help)
  weekly_report.py --version


Options:
  <section>      Section to export.
  -d,--debug     Turn on debug output.
  -n,--no_email  Do not send email.
  -w,--web       Serve report on local web server.
  --email=<email> Send to only this email address.
  --quarter=<quarter> Which quarter to use [default: current].
  --term=<term>  Which OSM term to use [default: current].
  -h,--help      Show this screen.
  --version      Show version.

"""

from collections import OrderedDict
import datetime
import logging
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import osm
from docopt import docopt
from group import Group
from group import OSM_REF_FIELD

MIN_AGE = Group.MIN_AGE
MAX_AGE = Group.MAX_AGE

# import compass

FROM = "Richard Taylor <r.taylor@bcs.org.uk>"
TO = {'Group': ['hippysurfer@gmail.com',
                'mike.armstrong@7thlichfield.org.uk',
                'adrian.grew@tesco.net'],
      'Maclean': ['maclean@7thlichfield.org.uk'],
      'Somers': ['somers@7thlichfield.org.uk'],
      'Swinfen': ['pten2106@yahoo.co.uk'],
      'Garrick': ['caroline_fellows@hotmail.com'],
      'Paget': ['riddleshome@gmail.com'],
      'Rowallan': ['markjoint@hotmail.co.uk'],
      'Johnson': ['simon@scouting.me.uk'],
      'Boswell': ['boswell@7thlichfield.org.uk'],
      'Erasmus': ['paul@scouting.me.uk'],
      'Adult': []}

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"

log = logging.getLogger(__name__)


class Reporter(object):
    STYLE = "\n".join(["hr {color:sienna;}",
                       "h1 {border-bottom-style:solid; border-color:red;}",
                       "p {margin-left:20px;}",
                       "table {border-collapse:collapse;}",
                       "table, th, td {border: 1px solid black;}"
                       ])

    def __init__(self):
        self.t = ""

    def send(self, to, subject,
             fro=FROM):
        for dest in to:
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = fro
            msg['To'] = dest

            body = MIMEText(self.t, 'html')

            msg.attach(body)

            # hostname = 'www.thegrindstone.me.uk' \
            #           if not socket.gethostname() == 'rat' \
            #           else 'localhost'
            hostname = 'localhost'

            s = smtplib.SMTP(hostname)
            s.sendmail(fro, dest, msg.as_string())
            s.quit()

    def report(self):
        return "<html><head><style>{}</style></head>" \
               "<body>\n{}\n</body></html>".format(self.STYLE, self.t)

    def title(self, title):
        self.t += '<h1 style="border-bottom-style:solid; border-color:red;"' \
                  '>{}</h1>\n'.format(title)

    def sub_title(self, sub_title):
        self.t += "<h2>{}</h2>\n".format(sub_title)

    def p(self, line):
        self.t += '<p style="margin-left:20px;">{}</p>\n'.format(line)

    def t_start(self, headings):
        self.t += '<table style="border-collapse:collapse;">'
        self.t += '<tr style="border: 1px solid black; padding: 10px;">'
        self.t += "".join(['<th style="border: 1px solid black;'
                           'padding: 10px;">{}</th>'.format(cell) for cell
                           in headings])
        self.t += "</tr>\n"

    def t_row(self, cells):
        self.t += '<tr style="border: 1px solid black; padding: 10px;">'
        self.t += "".join(['<td style="border: 1px solid black;'
                           'padding: 10px;">{}</td>'.format(cell) for cell
                           in cells])
        self.t += "</tr>\n"

    def t_end(self):
        self.t += "</table>\n"

    def ul(self, lines):
        self.t += "<ul>\n"
        self.t += "\n".join(["<li>{}</li>".format(line) for line in lines])
        self.t += "</ul>\n"


# noinspection PyUnusedLocal
def intro(r, group, section):
    r.title("Section report for {}".format(section))

    r.p("This is an automatic report of the OSM data for '{}'. "
        "It is sent on the first dat of each month.".format(section))
    r.p("Below is listed any potential problems with the data held "
        "in this section.")
    r.p("Please update the records to correct the identified issues.")
    r.p("You are receiving this email because you are listed as the OSM "
        "coordinator for this section. If this is incorrect please let "
        "me know so that I can correct the address.")
    r.p("If there is nothing below this list it means that there are no "
        "identified problems.")


def check_bad_data(r, group, section):

    subs_members_ids = [member[OSM_REF_FIELD] for member in
                        group.section_all_members(group.SUBS_SECTION)]
    all_yp_members_without_senior_duplicates = group.all_yp_members_without_senior_duplicates()

    members = group.section_yp_members_without_leaders(section) if \
        section != 'Adult' else group.section_all_members(section)

    reports = []
    for member in members:
        report = []

        if member['floating.gender'].lower() not in ['m', 'f', 'male', 'female']:
            report.append("Sex ({}) not in 'M', 'F', 'Male', 'Female'".format(
                member['floating.gender']))

        if section in group.YP_SECTIONS:
            if member['contact_primary_1.address1'].strip() == '':
                report.append("Primary Address missing")

            if (member['contact_primary_1.phone1'].strip() == '' and
                    member['contact_primary_1.phone2'].strip() == '' and
                    member['contact_primary_2.phone1'].strip() == '' and
                    member['contact_primary_2.phone2'].strip() == ''):
                report.append("No telephone number in primary contact 1 or 2")

            if (int(member.age().days / 365) < MIN_AGE[section] or
                    int(member.age().days / 365) > MAX_AGE[section]):
                report.append("Age ({}) is out of range ({} - {})".format(
                    int(member.age().days / 365),
                    MIN_AGE[section], MAX_AGE[section]))

            if member[OSM_REF_FIELD] not in subs_members_ids:
                report.append("Not in Subs Section")

        elif section == group.SUBS_SECTION:
            if member not in all_yp_members_without_senior_duplicates:
                report.append("Not in any YP section.")

        else:

            if member['contact_primary_member.address1'].strip() == '':
                report.append("Member Address missing")

            if (member['contact_primary_member.phone1'].strip() == '' and
                    member['contact_primary_member.phone2'].strip() == ''):
                report.append("No telephone number for member.")

        if report:
            reports.append((member, report))

    if reports:
        r.sub_title("Records with bad or missing data.")

        for report in reports:
            r.p("{} {}:".format(report[0]['first_name'],
                                report[0]['last_name']))
            r.ul(report[1])

# def section_compass_check(r, group, section):
#     """Check the content of Compass for discrepancies with OSM
#     for a specific section."""

#     c = compass.Compass(outdir=os.path.abspath('compass_exports'))
#     c.load_from_dir()

#     group.set_yl_as_yp(True)
#     osm_members = group.section_yp_members_without_leaders(section)

#     r.sub_title('Compass')
#     r.p('The following records appear in OSM but do not appear in Compass. '
#         'This may be because their entry in Compass has a different Firstname '
#         'or Lastname.')

#     if section not in c.sections():
#         r.p('No Compass data available for Section: {}'.format(section))
#         return

#     members_missing_in_compass = [member for member in osm_members
#                                   if c.find_by_name(
#                                       member['firstname'],
#                                       member['lastname'],
#                                       section_wanted=section,
#                                       ignore_second_name=True).empty]

#     if len(members_missing_in_compass):
#         r.t_start(compass.required_headings)

#         # Find YP missing from Compass
#         for member in members_missing_in_compass:
#                 compass_record = compass.member2compass(member, section)
#                 r.t_row([compass_record[k] for k in compass.required_headings])
#         r.t_end()

#     r.p('The following records appear in Compass but do not appear in OSM. '
#         'This may be because the Firstname or Lastname is different.')

#     members_missing_in_osm = [
#         member for member in c.section_yp_members_without_leaders(section)
#         if not group.find_by_name(member['forenames'],
#                                   member['surname'],
#                                   section_wanted=section,
#                                   ignore_second_name=True)]

#     if len(members_missing_in_osm):
#         keys = members_missing_in_osm[0].keys()
#         r.t_start(keys)

#         for member in members_missing_in_osm:
#             r.t_row([member[k] for k in keys])

#         r.t_end()


# def process_compass(r, group):
#     "Check the content of Compass for discrepancies with OSM"

#     c = compass.Compass(outdir=os.path.abspath('compass_exports'))
#     c.load_from_dir()

#     group.set_yl_as_yp(True)
#     all_yp = group.all_yp_members_without_senior_duplicates_dict()

#     r.sub_title('Compass')
#     r.p('The following records appear in OSM but do not appear in Compass '
#         'This may be because their entry in Compass has a different Firstname '
#         'or Lastname.  Only the first part of the Firstname is taken in to account.')

#     # Generate a dict of the sections with missing members.
#     members_missing_in_compass = {
#         s: [member for member in all_yp[s]
#             if c.find_by_name(
#                 member['firstname'],
#                 member['lastname'],
#                 ignore_second_name=True).empty]
#         for s in all_yp.keys()}

#     # If the dict is not empty.
#     if list(itertools.chain(*members_missing_in_compass.values())):
#         r.t_start(['OSM section'] + list(compass.required_headings))

#         for section, members in members_missing_in_compass.items():
#             for member in members:
#                 compass_record = compass.member2compass(member, section)
#                 r.t_row([section] + [compass_record[k]
#                                        for k in compass.required_headings])
#         r.t_end()

#     r.p('The following records appear in Compass but do not appear in OSM '
#         'This may be because the Firstname or Lastname is different. Only '
#         'the first part of the Firstname is taken in to account.')

#     r.t_start(['Compass Section', 'Membership', 'Firstname', 'Surname'])

#     compass_sections = c.all_yp_members_dict()
#     for section in compass_sections.keys():
#         for i, member in compass_sections[section].iterrows():
#             if not group.find_by_name(member['forenames'],
#                                       member['surname'],
#                                       ignore_second_name=True):
#                 r.t_row([section,
#                          member['membership_number'],
#                          member['forenames'],
#                          member['surname']])

#     r.t_end()

#     r.p('The following records appear more than once in Compass with different Membership numbers')

#     r.t_start(['Membership', 'Firstname', 'Surname', 'Location'])

#     for member in c.members_with_multiple_membership_numbers():
#         r.t_row([member['membership_number'].values[0],
#                  member['forenames'].values[0],
#                  member['surname'].values[0],
#                  member['location'].values[0]])
#     r.t_end()


def census(r, group):
    census_ = group.census()

    r.sub_title('Census')

    for i in ['Beavers', 'Cubs', 'Scouts']:
        r.t_start(['Section', 'Sex'] +
                  ["age - {} yrs".format(str(age))
                   for age in sorted(census_[i]['M'].keys())] +
                  ['Total'])
        male_counts = [census_[i]['M'][j] for j in sorted(census_[i]['M'].keys())]
        r.t_row([i, 'Male'] + male_counts + [sum(male_counts),])

        female_counts = [census_[i]['F'][j] for j in sorted(census_[i]['F'].keys())]
        r.t_row([i, 'Female'] + female_counts + [sum(female_counts),])

        other_counts = [census_[i]['O'][j] for j in sorted(census_[i]['O'].keys())]
        r.t_row([i, 'Other'] + other_counts + [sum(other_counts), ])

        r.t_end()
        r.p("Total {} = {}".format(i, sum(male_counts) + sum(female_counts) + sum(other_counts)))
        r.p("")

    total_male = 0
    for i in ['Beavers', 'Cubs', 'Scouts']:
        total_male += sum(census_[i]['M'].values())

    total_female = 0
    for i in ['Beavers', 'Cubs', 'Scouts']:
        total_female += sum(census_[i]['F'].values())

    total_other = 0
    for i in ['Beavers', 'Cubs', 'Scouts']:
        total_other += sum(census_[i]['O'].values())

    r.p("Total Male YP = {}".format(total_male))
    r.p("Total Female YP = {}".format(total_female))
    r.p("Total Other YP = {}".format(total_other))
    r.p("Census Total YP = {}".format(total_male + total_female + total_other))


COMMON = [intro,
          check_bad_data]

# NOT_ADULT = [section_compass_check, ]
NOT_ADULT = []

elements = OrderedDict((
    ('Garrick', COMMON + NOT_ADULT),
    ('Paget', COMMON + NOT_ADULT),
    ('Swinfen', COMMON + NOT_ADULT),
    ('Maclean', COMMON + NOT_ADULT),
    ('Rowallan', COMMON + NOT_ADULT),
    ('Somers', COMMON + NOT_ADULT),
    ('Erasmus', COMMON + NOT_ADULT),
    ('Boswell', COMMON + NOT_ADULT),
    ('Johnson', COMMON + NOT_ADULT),
    ('Adult', COMMON)))


def group_report(r, group, quarter, term):
    r.title("Group Report (Quarter: {} Term: {})".format(quarter, term))

    census(r, group)

    all_yp_members_without_leaders = group.all_yp_members_without_leaders_dict()
    all_yp_members_without_senior_duplicates = group.all_yp_members_without_senior_duplicates_dict()
    all_yl_members = group.all_yl_members_dict()

    r.sub_title("Section Totals")
    r.t_start(["Section", "Total YP (including Scubbers)", "Total YP (excluding Scubbers)", "Young Leaders"])

    total_yp_members_without_leaders = 0
    total_yp_members_without_senior_duplicates = 0
    total_yl_members = 0
    for section in group.YP_SECTIONS:
        yp_members_without_leaders = len(all_yp_members_without_leaders[section])
        yp_members_without_senior_duplicates = len(all_yp_members_without_senior_duplicates[section])
        yl_members = len(all_yl_members[section])
        r.t_row([section, yp_members_without_leaders, yp_members_without_senior_duplicates, yl_members])
        total_yp_members_without_leaders += yp_members_without_leaders
        total_yp_members_without_senior_duplicates += yp_members_without_senior_duplicates
        total_yl_members += yl_members

    r.t_row(["Total", total_yp_members_without_leaders, total_yp_members_without_senior_duplicates, total_yl_members])
    r.t_end()
    r.p("")

    r.sub_title("Section Type Totals")
    r.t_start(["Section Type", "Total YP (including Scubbers)", "Total YP (excluding Scubbers)", "Young Leaders"])

    total_yp_members_without_leaders = 0
    total_yp_members_without_senior_duplicates = 0
    total_yl_members = 0
    for section_type in group.SECTIONS_BY_TYPE.keys():
        section_type_without_leaders = 0
        section_type_without_senior_duplicates = 0
        section_type_yl_members = 0
        for section in group.SECTIONS_BY_TYPE[section_type]:
            section_type_without_leaders += len(all_yp_members_without_leaders[section])
            section_type_without_senior_duplicates += len(all_yp_members_without_senior_duplicates[section])
            section_type_yl_members += len(all_yl_members[section])

        r.t_row([section_type, section_type_without_leaders, section_type_without_senior_duplicates, section_type_yl_members])

        total_yp_members_without_leaders += section_type_without_leaders
        total_yp_members_without_senior_duplicates += section_type_without_senior_duplicates
        total_yl_members += section_type_yl_members

    r.t_row(["Total", total_yp_members_without_leaders, total_yp_members_without_senior_duplicates, total_yl_members])
    r.t_end()
    r.p("")

    # Get a list of all YL in all of the YP sections
    yls = group.all_yl_members()
    # Create a list of their names.
    refs = ["{} {}".format(_['first_name'], _['last_name']) for _ in yls]
    # Get a list of any that appear more than once.
    duplicates = set(["{} {}".format(_['first_name'], _['last_name']) for _ in yls
                      if refs.count("{} {}".format(_['first_name'], _['last_name'])) > 1])
    # Get the total number of duplications.
    dup_count = sum(set([refs.count("{} {}".format(_['first_name'], _['last_name'])) for _ in yls
                         if refs.count("{} {}".format(_['first_name'], _['last_name'])) > 1]))

    r.sub_title("Young Leaders that are in more than 1 Section")

    r.t_start(["Name", "Number of Sections"])

    for yl in duplicates:
        r.t_row([yl, refs.count(yl)])

    r.t_end()

    r.p("Total Young Leaders in Sections (duplicates removed) = {}".format(total_yl_members - dup_count))

    # process_compass(r, group)

    # process_finance_spreadsheet(r, group, quarter)

    for section in elements.keys():
        for element in elements[section]:
            element(r, group, section)


def _main(osm_, auth_, sections, no_email, email, quarter, term, http):
    if isinstance(sections, str):
        sections = [sections, ]

    important_fields = ['first_name',
                        'last_name',
                        'joined',
                        'started',
                        'date_of_birth']

    group = Group(osm_, auth_, important_fields, term)

    for section in sections:
        assert section in list(group.SECTIONIDS.keys()) + ['Group', ], \
            "section must be in {!r}.".format(group.SECTIONIDS.keys())

    for section in sections:
        r = Reporter()

        if section == 'Group':
            group_report(r, group, quarter,
                         term if term is not None else "Active")
        else:
            for element in elements[section]:
                element(r, group, section)

        if no_email:
            print(r.report())
        elif email:
            print("Sending to {}".format(email))
            r.send([email, ],
                   'OSM Data Integrity Report for {}'.format(section))
        else:
            r.send(TO[section],
                   'OSM Data Integrity Report for {}'.format(section))

        if http:
            print("Serving {}".format(section))
            serve(r)


def serve(report):
    import http.server

    # noinspection PyClassHasNoInit
    class Handler(http.server.BaseHTTPRequestHandler):

        def do_GET(self):
            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.end_headers()
            self.wfile.write(bytes(report.report(), 'UTF-8'))

    server_address = ('', 8000)
    httpd = http.server.HTTPServer(server_address, Handler)

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass

    httpd.server_close()


def get_quarter():
    """Return the current quarter from today's date."""
    month = datetime.datetime.today().month
    if month in [4, 5, 6]:
        return "Q1"
    if month in [7, 8, 9]:
        return "Q2"
    if month in [10, 11, 12]:
        return "Q3"
    if month in [1, 2, 3]:
        return "Q4"


if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.WARN

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    if args['--term'] in [None, 'current']:
        args['--term'] = None

    if args['--quarter'] in [None, 'current']:
        args['--quarter'] = get_quarter()

    if args['--quarter'] not in ['Q1', 'Q2', 'Q3', 'Q4']:
        log.error("Invalid quarter ({}): quarter must be in "
                  "['Q1', 'Q2', 'Q3', 'Q4']".format(args['--quarter']))
        sys.exit(1)

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    _main(osm, auth,
          args['<section>'],
          args['--no_email'],
          args['--email'],
          args['--quarter'],
          args['--term'],
          args['--web'])
