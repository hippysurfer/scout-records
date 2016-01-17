# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  weekly_report.py [-d | --debug] [-n | --no_email] [--email=<email>] [-w | --web] [--quarter=<quarter>] [--term=<term>] <apiid> <token> <section>...
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

import logging
from docopt import docopt
import osm
import smtplib
import socket
import re
import sys
import datetime
import os.path
import itertools

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from group import Group
from update import MAPPING
from group import OSM_REF_FIELD
import finance
import google
#import compass

PERSONAL_REFERENCE_RE = re.compile('^[A-Z0-9]{4}-[A-Z]{2}-\d{6}$')

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
      'Boswell': ['marc.henson@7thlichfield.org.uk'],
      'Erasmus': ['paul@scouting.me.uk'],
      'Adult': ['susanjowen@btinternet.com']}

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

            #hostname = 'www.thegrindstone.me.uk' \
            #           if not socket.gethostname() == 'rat' \
            #           else 'localhost'
            hostname = 'localhost'

            s = smtplib.SMTP(hostname)
            s.sendmail(fro, dest, msg.as_string())
            s.quit()

    def report(self):
        return "<html><head><style>{}</style></head>"\
            "<body>\n{}\n</body></html>".format(self.STYLE, self.t)

    def title(self, title):
        self.t += '<h1 style="border-bottom-style:solid; border-color:red;"'\
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
    MIN_AGE = {
        'Swinfen': 5,
        'Paget': 5,
        'Garrick': 5,
        'Maclean': 7,
        'Rowallan': 7,
        'Somers': 7,
        'Johnson': 10,
        'Boswell': 10,
        'Erasmus': 10,
    }

    MAX_AGE = {
        'Swinfen': 8,
        'Paget': 8,
        'Garrick': 8,
        'Maclean': 10,
        'Rowallan': 10,
        'Somers': 10,
        'Johnson': 15,
        'Boswell': 15,
        'Erasmus': 15
    }

    members = group.section_yp_members_without_leaders(section) if \
        section != 'Adult' else group.section_all_members(section)

    reports = []
    for member in members:
        report = []

        if member['floating.gender'].lower() not in ['m', 'f', 'male', 'female']:
            report.append("Sex ({}) not in 'M', 'F', 'Male', 'Female'".format(
                member['floating.gender']))

        if member['customisable_data.PersonalReference'].strip() == '':
            report.append("<b>Missing Personal Reference</b>")

        elif not PERSONAL_REFERENCE_RE.match(
                member['customisable_data.PersonalReference'].strip()):
            report.append("<b>Bad Personal Reference ('{}')"
                          " must match pattern: SSSS-FF-DDMMYY</b>".format(
                              member['customisable_data.PersonalReference'].strip()))

        if member['contact_primary_1.address1'].strip() == '':
            report.append("Primary Address missing")

        if member['contact_primary_1.phone1'].strip() == '':
            report.append("Home Tel missing")

        if section != 'Adult':

            if int(member.age().days / 365) < MIN_AGE[section] or \
               int(member.age().days / 365) > MAX_AGE[section]:

                report.append("Age ({}) is out of range ({} - {})".format(
                    int(member.age().days / 365),
                    MIN_AGE[section], MAX_AGE[section]))

        if report:
            reports.append((member, report))

    if reports:
        r.sub_title("Records with bad or missing data.")

        for report in reports:
            r.p("{} {}:".format(report[0]['first_name'],
                                report[0]['last_name']))
            r.ul(report[1])


def process_finance_spreadsheet(r, group, quarter):
    log.info("Processing finance spreadsheet...")

    gc = google.conn()

    fin = gc.open(finance.FINANCE_SPREADSHEET_NAME)
    #fin = gc.open_by_key('0AobgMqwG6nlpdHdocFkwVVhNd2pGbTBRX1pCanBVdVE')

    # Fetch list of personal references from the finance spreadsheet
    # along with the current "Q4" section.
    wks = fin.worksheet(finance.DETAIL_WKS)
    headings = wks.row_values(finance.FIN_HEADER_ROW)
    fin_references = wks.col_values(
        1 + headings.index(
            finance.FIN_MAPPING_DETAILS[OSM_REF_FIELD])
    )[finance.FIN_HEADER_ROW:]

    # TODO: Parameterise the selection of the current quarter.
    #q4_section = wks.col_values(
    #    1 + headings.index('{} Sec'.format(quarter)))[finance.FIN_HEADER_ROW:]

    group.set_yl_as_yp(False)
    all_yp = group.all_yp_members_without_senior_duplicates_dict()

    # create a map from refs to members for later lookup
    all_members = {}
    for name, section_members in all_yp.items():
        for member in section_members:
            all_members[member[OSM_REF_FIELD]] = member

    # Create a list of all YP that are not on the finance list.
    # get list of all references from sections
    new_members = []
    for name, section_members in all_yp.items():
        for member in section_members:
            if member[OSM_REF_FIELD] not in fin_references:
                member['SeniorSection'] = name
                new_members.append(member)

    r.sub_title("Finance Spreadsheet")
    r.p("The following members appear in the sections in OSM but do not appear"
        " on the Finance Spreadsheet. (New Members)")
    headings = [('Patrol', 'patrol'),
                ('SeniorSection', 'SeniorSection'),
                ('PersonalReference', 'customisable_data.PersonalReference'),
                ('Membership', 'customisable_data.membershipno'),
                ('Firstname', 'first_name'),
                ('Lastname', 'last_name'),
                ('PersonalEmail', 'contact_primary_member.email1'),
                ('DadEmail', 'contact_primary_2.email1'),
                ('MumEmail', 'contact_primary_1.email1'),
                ('dob', 'date_of_birth'),
                ('Joined', 'joined'),
                ('Started', 'started')]

    r.t_start([h[0] for h in headings])
    for member in new_members:
        r.t_row([member[k[1]] for k in headings])
    r.t_end()

    # Create a list of all YP that are on the finance list but are not
    # in OSM.
    all_osm_references = []
    for name, section_members in all_yp.items():
        all_osm_references.extend([member[OSM_REF_FIELD].strip()
                                   for member in section_members])

    missing_references = []
    for ref in fin_references:
        if ref and (ref.strip() not in all_osm_references):
            missing_references.append(ref)

    r.p("The following members appear in the Finance Spreadsheet but "
        "do not appear in the sections on OSM. (Old Member)")

    headings = [('Membership', 'customisable_data.membershipno'),
                ('Firstname', 'first_name'),
                ('Lastname', 'last_name'),
                ('Patrol', 'patrol')]

    r.t_start([h[0] for h in headings])

    for ref in missing_references:
        if not ref:
            log.warn("Ignoring null reference.")
            continue
        member = group.find_by_ref(ref)
        if len(member) > 0:
            r.t_row([member[0][k[1]] for k in headings])
        else:
            r.t_row([ref,])

    r.t_end()

    # # Create a list of all YP who are on the finanace list but are not
    # # in the same section in OSM.
    # section_map = {'Maclean': 'MP',
    #                'Rowallan': 'RP',
    #                'Somers': 'SP',
    #                'Swinfen': 'BC',
    #                'Garrick': 'GC',
    #                'Boswell': 'BT',
    #                'Johnson': 'JT',
    #                'Erasmus': 'ET',
    #                'Paget': 'PC'}

    # log.debug("fin_references - {}".format(fin_references))
    # log.debug("q_section - {}".format(q4_section))

    # changed_members = []
    # for name, section_members in all_yp.items():
    #     for member in section_members:
    #         log.debug("member = {}".format(member))
    #         if member[OSM_REF_FIELD] in fin_references:
    #             try:
    #                 previous_section = q4_section[fin_references.index(
    #                     member[OSM_REF_FIELD])]
    #             except IndexError:
    #                 # If the spreadsheet does not have enough columns we assume that
    #                 # the previous section was None: i.e. this is a new YP.
    #                 previous_section = ""

    #             if section_map[name] != previous_section:
    #                 changed_members.append((member,
    #                                         previous_section,
    #                                         section_map[name]))

    # r.p("The following have moved sections on OSM but are "
    #     "still recorded in their old section in the Finance "
    #     "Spreadsheet (Changed members)")
    # r.t_start(["Membership", "Old", "New", "First", "Last"])
    # for member in changed_members:
    #     r.t_row([
    #         member[0][OSM_REF_FIELD],
    #         member[1],
    #         member[2],
    #         all_members[member[0][OSM_REF_FIELD]]['first_name'],
    #         all_members[member[0][OSM_REF_FIELD]]['last_name']])

    # r.t_end()


# def section_compass_check(r, group, section):
#     """Check the content of Compass for descrepencies with OSM
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
#     "Check the content of Compass for descrepencies with OSM"

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
    census = group.census()

    r.sub_title('Census')

    for i in ['Beavers', 'Cubs', 'Scouts']:
        r.t_start(['Section', 'Sex'] +
                  ["age - {} yrs".format(str(age))
                   for age in sorted(census[i]['M'].keys())])
        r.t_row([i, 'Male'] + [census[i]['M'][j] for j
                               in sorted(census[i]['M'].keys())])
        r.t_row([i, 'Female'] + [census[i]['F'][j] for j
                                 in sorted(census[i]['F'].keys())])
        r.t_end()

    total_male = 0
    for i in ['Beavers', 'Cubs', 'Scouts']:
        total_male += sum(census[i]['M'].values())

    total_female = 0
    for i in ['Beavers', 'Cubs', 'Scouts']:
        total_female += sum(census[i]['F'].values())

    r.p("Male = {}".format(total_male))
    r.p("Female = {}".format(total_female))
    r.p("Census total = {}".format(total_male + total_female))

COMMON = [intro,
          check_bad_data]

#NOT_ADULT = [section_compass_check, ]
NOT_ADULT = []

elements = {'Maclean': COMMON + NOT_ADULT,
            'Rowallan': COMMON + NOT_ADULT,
            'Garrick': COMMON + NOT_ADULT,
            'Somers': COMMON + NOT_ADULT,
            'Erasmus': COMMON + NOT_ADULT,
            'Swinfen': COMMON + NOT_ADULT,
            'Boswell': COMMON + NOT_ADULT,
            'Johnson': COMMON + NOT_ADULT,
            'Paget': COMMON + NOT_ADULT,
            'Adult': COMMON}


def group_report(r, group, quarter, term):
    r.title("Group Report (Quarter: {} Term: {})".format(quarter, term))

    census(r, group)

    r.sub_title("Section Totals (including Scubbers)")
    r.t_start(["Section", "Total YP"])

    for section, members in group.all_yp_members_without_leaders_dict().items():
        r.t_row([section, len(members)])

    r.t_end()

    r.sub_title("Section Totals (excluding Scubbers)")

    r.t_start(["Section", "Total YP"])

    for section, members in group.all_yp_members_without_senior_duplicates_dict().items():
        r.t_row([section, len(members)])

    r.t_end()

    #process_compass(r, group)

    #process_finance_spreadsheet(r, group, quarter)

    for section in elements.keys():
        for element in elements[section]:
            element(r, group, section)


def _main(osm, auth, sections, no_email, email, quarter, term, http):

    if isinstance(sections, str):
        sections = [sections, ]

    group = Group(osm, auth, MAPPING.keys(), term)

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

    class Handler(http.server.BaseHTTPRequestHandler):

        def do_GET(s):
            s.send_response(200)
            s.send_header("Content-type", "text/html")
            s.end_headers()
            s.wfile.write(bytes(report.report(), 'UTF-8'))

    server_address = ('', 8000)
    httpd = http.server.HTTPServer(server_address, Handler)

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass

    httpd.server_close()


def get_quarter():
    """Return the currect quarter from todays date."""
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
