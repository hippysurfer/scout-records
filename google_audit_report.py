"""Online Scout Manager Interface.

Usage:
  google_audit_report.py [-d | --debug] [-n | --no_email] [--email=<email>] [-w | --web]
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

import csv
import subprocess
import logging
import dateutil.parser
import datetime

from docopt import docopt
from weekly_report import Reporter

log = logging.getLogger(__name__)

GAMX = '/home/rjt/bin/gamx/gam'
GAM = '/home/rjt/bin/gam/gam'

NOW = datetime.datetime.now(datetime.timezone.utc)


def fetch_user_report():
    out = subprocess.run(args=[GAMX, 'report', 'users'], stdout=subprocess.PIPE, check=True,
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


def _main(no_email, email, http):
    r = Reporter()

    admin_team_drives = fetch_admin_team_drives()

    users = fetch_user_report()
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
        r.t_row([user['accounts:admin_set_name'], user['email']])

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
                        (user['accounts:last_login_time'] == 'Never' or
                         (NOW - dateutil.parser.parse(user['accounts:last_login_time'])).days > 30) and
                        user['email'] not in forwarding_users_email_addresses]

    forwardingaddresses_by_email = dict([(_['User'], _) for _ in forwardingaddresses])
    for user in last_login_users:
        forwarding_email = (forwardingaddresses_by_email[user['email']]['forwardingEmail']
                            if forwardingaddresses_by_email.get(user['email']) else "Not set")
        verification_status = (forwardingaddresses_by_email[user['email']]['verificationStatus']
                               if forwardingaddresses_by_email.get(user['email']) else "Not set")

        r.t_row([user['accounts:admin_set_name'],
                 user['email'],
                 user['accounts:last_login_time'],
                 forwarding_email,
                 verification_status])

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

    _main(args['--no_email'],
          args['--email'],
          args['--web'])
