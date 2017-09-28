"""Online Scout Manager Interface.

Usage:
  google_audit_report.py [-d | --debug] [-n | --no_email] [--email=<email>] [-w | --web] <apiid> <token> <section>
  google_audit_report.py (-h | --help)
  google_audit_report.py --version


Options:
  -d,--debug     Turn on debug output.
  -n,--no_email  Do not send email.
  -w,--web       Serve report on local web server.
  --email=<email> Send to only this email address.
  -h,--help      Show this screen.
  --version      Show version.
"""

import os
import csv
import string
import subprocess
import logging
import dateutil.parser
import datetime
import functools

from docopt import docopt
from weekly_report import Reporter


import osm
from group import Group
from update import MAPPING

log = logging.getLogger(__name__)

GAMX = os.environ.get('GAMX_HOME', '/home/rjt/bin/gamx/gam')
GAM = os.environ.get('GAM_HOME', '/home/rjt/bin/gam/gam')

NOW = datetime.datetime.now(datetime.timezone.utc)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"

def fetch_osm_adults(osm, auth, sections):
    group = Group(osm, auth, MAPPING.keys(), term=None)

    contacts = []

    def get(field, member_, section):
        return member_["{}.{}".format(section, field)]

    for section_ in sections:
        for member in group.section_all_members(section_):
            f1 = functools.partial(get, member_=member, section='contact_primary_member')

            first = string.capwords(member['first_name'].strip())
            last = string.capwords(member['last_name'].strip())
            full_name = f'{first} {last}'

            def is_valid_email(func, tag):
                email_ = func(tag).strip()
                return (len(email_) != 0
                        and (not email_.startswith('x ') and
                             func("{}_leaders".format(tag)) == "yes"))

            # Add member data
            email = f1('email1').strip()
            if group.is_leader(member) and not is_valid_email(f1, 'email1'):
                log.warning(f'{full_name} is a leader but does not have a member email address.')

            if group.is_leader(member) and is_valid_email(f1, 'email1'):
                key = '{}-{}-{}'.format(first.lower(), last.lower(), email.lower())
                if key not in contacts:
                    contacts.append(dict(first=first, last=last, email=email.lower()))
    return contacts


def fetch_user_report():
    out = subprocess.run(args=[GAMX, 'print', 'users', 'allfields'], stdout=subprocess.PIPE, check=True,
                         universal_newlines=True)
    return list(csv.DictReader(out.stdout.splitlines()))


def fetch_email_forwarding():
    out = subprocess.run(args=[GAMX, 'all', 'users', 'print', 'forward'], stdout=subprocess.PIPE, check=True,
                         universal_newlines=True)
    return list(csv.DictReader(out.stdout.splitlines()))


def fetch_email_forwarding_address_status():
    out = subprocess.run(args=[GAMX, 'all', 'users', 'print', 'forwardingaddresses'], stdout=subprocess.PIPE, check=True,
                         universal_newlines=True)
    return list(csv.DictReader(out.stdout.splitlines()))


def fetch_group_members():
    out = subprocess.run(args=[GAMX, 'print', 'group-members'], stdout=subprocess.PIPE, check=True,
                         universal_newlines=True)
    return list(csv.DictReader(out.stdout.splitlines()))


def fetch_admin_team_drives():
    out = subprocess.run(args=[GAM, 'user', 'admin', 'print', 'teamdrives'], stdout=subprocess.PIPE, check=True,
                         universal_newlines=True)
    return list(csv.DictReader(out.stdout.splitlines()))


def fetch_team_drive_acl(drive_id):
    out = subprocess.run(args=[GAM, 'user', 'admin', 'show', 'drivefileacl', drive_id], stdout=subprocess.PIPE, check=True,
                         universal_newlines=True)

    ret = []
    current = dict()
    for line in out.stdout.splitlines():
        if line == '':
            ret.append(current)
            current = dict()
            continue
        try:
            k, v = line.split(':')
            current[k.strip()] = v.strip()
        except ValueError:
            continue

    return ret


def intro(r):
    r.title("Google Audit Report")

    r.p("This is an automatic report ... ")


def _main(osm, auth, sections, no_email, email, http):
    osm_adults = fetch_osm_adults(osm, auth, sections)
    users = fetch_user_report()
    osm_users_first_last = set([(_['first'], _['last']) for _ in osm_adults])
    g_users_first_last = set([(_['name.givenName'], _['name.familyName']) for _ in users])

    missing_in_g = osm_users_first_last - g_users_first_last
    missing_in_g_with_email_addresses = [_ for _ in osm_adults if (_['first'], _['last']) in missing_in_g]

    # remove duplicates
    missing_in_g_with_email_addresses = [dict(t) for t in set([tuple(d.items()) for d in
                                                               missing_in_g_with_email_addresses])]
    extras_in_g = g_users_first_last - osm_users_first_last

    r = Reporter()

    admin_team_drives = fetch_admin_team_drives()

    forwarding = fetch_email_forwarding()
    forwardingaddresses = fetch_email_forwarding_address_status()
    group_members = fetch_group_members()

    permission_group_members = [member for member in group_members
                                if member['group'].startswith('7th-lichfield-')]

    role_group_members = [member for member in group_members
                          if not member['group'].startswith('7th-lichfield-')]

    r.sub_title("All users")
    r.p('This is a list of all of the user accounts in the domain. It should only list real people.')
    r.t_start(["Name", "Email"])

    for user in users:
        r.t_row([user['name.fullName'], user['primaryEmail']])

    r.t_end()
    r.p("")

    r.sub_title("Users that are forwarding their email")
    r.p('This is a list of people that have forwarded their @7thlichfield.org.uk email address to '
        'another account somewhere. These people will get their email but will not be able to access '
        'team drives, unless they login to their @7thlichfield.org.uk account.')

    r.t_start(["User", "Forwarding Address"])

    forwarding_users = [user for user in forwarding if user['forwardEnabled'] == 'True']
    for user in forwarding_users:
        r.t_row([user['User'], user['forwardTo']])

    r.t_end()
    r.p("")

    r.sub_title("Users that do not have forwarding and have not logged in for a month")
    r.p('These are people that have not logged in for more than 30 days but do not have an active '
        'email forwarder set up. These people will not be getting any email that is sent to their '
        '@7thlichfield.org.uk account at all.')
    r.p('If the Status == pending (and there is a Forwarding Address listed) it means that they have '
        'been sent email requesting them to authorise forwarding but they have not responded to it.')

    r.t_start(["User", "Email", "Last Login", "Forwarding Address", "Status"])

    forwarding_users_email_addresses = [user['User'] for user in forwarding_users]
    last_login_users = [user for user in users if
                        (user['lastLoginTime'] == 'Never' or
                         (NOW - dateutil.parser.parse(user['lastLoginTime'])).days > 30) and
                        user['primaryEmail'] not in forwarding_users_email_addresses]

    forwardingaddresses_by_email = dict([(_['User'], _) for _ in forwardingaddresses])
    for user in last_login_users:
        forwarding_email = (forwardingaddresses_by_email[user['primaryEmail']]['forwardingEmail']
                            if forwardingaddresses_by_email.get(user['primaryEmail']) else "Not set")
        verification_status = (forwardingaddresses_by_email[user['primaryEmail']]['verificationStatus']
                               if forwardingaddresses_by_email.get(user['primaryEmail']) else "Not set")

        r.t_row([user['name.fullName'],
                 user['primaryEmail'],
                 user['lastLoginTime'],
                 forwarding_email,
                 verification_status])

    r.t_end()
    r.p("")

    r.sub_title("Adults that are in OSM but do not have Google accounts")
    r.p('This is a list of people that are listed as adults in OSN but do not appear to have '
        '@7thlichfield.org.uk accounts.')

    r.t_start(["First", "Last", "OSM Member Email"])

    for user in missing_in_g_with_email_addresses:
        r.t_row([user['first'], user['last'], user['email']])

    r.t_end()
    r.p("")

    r.sub_title("Users with @7thlichfield.org.uk accounts that are not in OSM")
    r.p('This is a list of people that have @7thlichfield.org.uk but do not appear to be in OSM.')

    r.t_start(["First", "Last"])

    for user in extras_in_g:
        r.t_row([user[0], user[1]])

    r.t_end()
    r.p("")

    r.sub_title("Groups used to apply permissions to Team Drives")
    r.p('These groups are used to control who can access the Team Drives.')

    r.t_start(["Group", "Member", "Role", "Status"])

    last = ""
    for member in permission_group_members:
        r.t_row([member['group'] if member['group'] != last else "",
                 member['email'],
                 member['status'],
                 member['role']])
        last = member['group']

    r.t_end()
    r.p("")

    r.sub_title("Groups used as collaborative inboxes for roles")
    r.p('These groups are the addresses of the inboxes used for roles within the group.')

    r.t_start(["Group", "Member", "Role", "Status"])

    last = ""
    for member in role_group_members:
        r.t_row([member['group'] if member['group'] != last else "",
                 member['email'],
                 member['status'],
                 member['role']])
        last = member['group']

    r.t_end()
    r.p("")

    r.sub_title("Team Drive access permissions")
    r.p('This list shows the permissions set for the top level of the Team Drives. All files within the '
        'Team Drive will have at least these permissions.')

    r.t_start(["Team Drive", "Name", "Email", "Role", "Team Drive Role", "Deleted"])

    role_map = {'organizer': 'full',
                'writer': 'edit',
                'reader': 'view',
                'commenter': 'comment'}

    for drive in admin_team_drives:
        last = ""
        for perm in fetch_team_drive_acl(drive['id']):
            r.t_row([drive['name'] if drive['name'] != last else "",
                     perm['displayName'],
                     perm['emailAddress'],
                     perm['role'],
                     role_map.get(perm['role'], ''),
                     perm['deleted']])
            last = drive['name']

    r.t_end()
    r.p("")

    if no_email:
        print(r.report())
    elif email:
        print("Sending to {}".format(email))
        r.send([email, ],
               'Google Audit Report')

    if http:
        print("Serving on http://localhost:8000/")
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


if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.WARN

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    sections = None
    if args['<section>']:
        section = args['<section>']

        if section == 'Group':
            sections = Group.YP_SECTIONS + [Group.ADULT_SECTION, ]
        else:
            assert section in list(Group.SECTIONIDS.keys()) + ['Group'] + list(Group.SECTIONS_BY_TYPE.keys()), \
                "section must be in {!r}.".format(list(Group.SECTIONIDS.keys()) + ['Group'])

            sections = Group.SECTIONS_BY_TYPE[section] if section in Group.SECTIONS_BY_TYPE.keys() else [section, ]

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    _main(osm, auth, sections,
          args['--no_email'],
          args['--email'],
          args['--web'])