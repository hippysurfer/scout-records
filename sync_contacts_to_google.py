# coding=utf-8
"""OSM Command Line

Usage:
   sync_contact_details_to_google [options] <apiid> <token> <section> [<google_account> ...]   

If you provide one or more google_accounts the contacts will be added to those accounts as
contacts as well as syncing them the google global directory.

Options:
   -t term, --term=term  Term to use
   --delete-groups       Delete and recreate all the google groups.
   
   
"""
import csv
import io
import logging
import functools
import string
import subprocess
import re

from docopt import docopt
import osm
import datetime

from group import Group
from update import MAPPING

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"

GAM = '/home/rjt/bin/gamx/gam'
GAM_SUBPROCESS_OPTS = {'shell': True, 'universal_newlines': True, 'stderr': subprocess.PIPE}
# ADMIN_USER = 'admin@7thlichfield.org.uk'

DATETIME = datetime.datetime.now().strftime('%d, %b %Y %H:%M')

# GROUP_MANAGER_MAP = {
#     # Exoficio and Admin can post to all.
#     'osm-.*@.*': ['7th-lichfield-admin@7thlichfield.org.uk',
#                   '7th-lichfield-ex-officio@7thlichfield.org.uk'],
#
#     # Only ex-officio and volunteering should be able to post to all.
#     # Need to add volunteering TODO
#     'osm-all@.*': [],
#
#     # Can't remember what this is? TODO
#     'osm-adult-young-leaders@.*': [],
#
#     # All Leaders should be able to post to all leaders
#     'osm-.*-leaders@.*': ['7th-lichfield-leaders@7thlichfield.org.uk'],
#
#     # We may have a YL coordinator in the future that needs to be able to do this.
#     # 'osm-all-parents@.*': [],
#     # 'osm-all-young-leaders@.*': [],
#
#     # Section leaders can post to the parents in their section.
#     'osm-(beavers|swinfen|paget|garrick)-parents@.*': ['7th-lichfield-beaver-leaders@7thlichfield.org.uk'],
#
#     'osm-(cubs|maclean|rowallan|maclean)-parents@.*': ['7th-lichfield-cub-leaders@7thlichfield.org.uk'],
#
#     'osm-(scouts|boswell|johnson|erasmus)-parents@.*': ['7th-lichfield-scout-leaders@7thlichfield.org.uk'],
#
#     # Section leaders can post to young leaders in their section.
#     # AGSL should be it in the leader groups.
#     'osm-(beavers|swinfen|paget|garrick)-young-leaders@.*': ['7th-lichfield-beaver-leaders@7thlichfield.org.uk'],
#
#     'osm-(cubs|maclean|rowallan|maclean)-young-leaders@.*': ['7th-lichfield-cub-leaders@7thlichfield.org.uk'],
#
#     'osm-(scouts|boswell|johnson|erasmus)-young-leaders@.*': ['7th-lichfield-scout-leaders@7thlichfield.org.uk'],
#
# }

GROUP_MODERATOR_MAP = {
    # Exoficio and Admin can post to all.
    'osm-.*@.*': ['7th-lichfield-admin@7thlichfield.org.uk'],
}


def run_with_input(args, inp, ignore_exc=False):
    p = subprocess.Popen(
        args=args,
        stdin=subprocess.PIPE, stderr=subprocess.PIPE,
        universal_newlines=True)

    std_out, std_err = p.communicate(input=inp)
    p.wait()
    if p.returncode != 0:
        # an error happened!
        err_msg = "cmd: {}, \nError: {}. \nCode: {} \nInp: {}".format(
            repr(args), std_err.strip(), p.returncode, repr(inp)[0:100])
        log.warning(err_msg)
        if not ignore_exc:
            raise Exception(err_msg)

    return std_out


def run_with_output(args, default=None):
    try:
        out = subprocess.run(args=args,
                             stdout=subprocess.PIPE,
                             stderr=subprocess.PIPE,
                             check=True,
                             universal_newlines=True)
        return out.stdout
    except subprocess.CalledProcessError as e:
        print_gam_error(e)

    return default


def print_gam_error(exc):
    print('Error executing GAM command: {} \n'
          'exit code: {}\n'
          'stdout: {}\n'
          'stderr: {}'.format(repr(exc.cmd), exc.returncode, exc.output, exc.stderr))


def sync_contacts(osm, auth, sections, google_accounts,
                  csv=False, term=None, no_headers=False,
                  delete_google_groups=False):
    group = Group(osm, auth, MAPPING.keys(), term)
    section_map = Group.SECTION_TYPE
    contacts = {}

    def get(field, member, section):
        return member["{}.{}".format(section, field)]

    def get_address(func):
        return ",".join([func('address1'),
                         func('address2'),
                         func('address3'),
                         func('postcode')])

    def add_member_contacts(section, section_type, member):
        f1 = functools.partial(get, member=member, section='contact_primary_member')
        custom_field = functools.partial(get, member=member, section='customisable_data')

        first = string.capwords(member['first_name'].strip())
        last_ = string.capwords(member['last_name'].strip())
        last = f'{last_} (OSM)'
        full_name_osm = f'{first} {last}'
        full_name = f'{first} {last}'
        section_capped = string.capwords(section)
        yp_name = string.capwords(f"{first} {last_} ({section_capped} {section_type})")
        parent_groups = [string.capwords(_) for _ in [f'{section} Parents', f'{section_type} Parents',
                                                      f'All Parents', 'All']]
        yl_groups = [string.capwords(_) for _ in
                     [f'{section} Young Leaders', f'{section_type} Young Leaders',
                      f'All Young Leaders', 'All Adults and Young Leaders', 'All']]

        leader_groups = [string.capwords(_) for _ in
                         [f'{section} Leaders', f'{section_type} Leaders',
                          'All Leaders', 'All Adults and Young Leaders',
                          'All']]

        all_adults_and_young_leaders_groups = ['All Adults And Young Leaders', 'All']

        exec_groups = ['Exec Members', 'All']

        def is_valid_email(func, tag):
            email_ = func(tag).strip()
            return (len(email_) != 0
                    and (not email_.startswith('x ')))

        def parse_tel(number_field, default_name):
            index = 0
            for i in range(len(number_field)):
                if number_field[i] not in ['0', '1', '2', '3', '4', '5',
                                           '6', '7', '8', '9', ' ', '\t']:
                    index = i
                    break

            number_ = number_field[:index].strip() if index != 0 else number_field
            name_ = number_field[index:].strip() if index != 0 else default_name

            # print("input = {}, index = {}, number = {}, name = {}".format(
            #    number_field, index, number, name))

            return number_, name_

        # Add member data
        email = f1('email1').strip()
        if group.is_leader(member) and not is_valid_email(f1, 'email1'):
            log.warning(f'{full_name} ({section}) is a leader but does not have a member email address.')

        if is_valid_email(f1, 'email1'):
            key = '{}-{}-{}'.format(first.lower(), last.lower(), email.lower())
            address = get_address(f1)
            if key not in contacts:
                contacts[key] = dict(first=first,
                                     last=last,
                                     email=email.lower(),
                                     addresses=[],
                                     yp=[],
                                     sections=[],
                                     groups=[],
                                     phones=[])
            contacts[key]['sections'].append(section)
            if not group.is_leader(member):
                contacts[key]['yp'].append(yp_name)

            if section.lower() == 'adult':
                # Everyone in the adult group is either a leader, a non-leader adult or a YL.
                contacts[key]['groups'].extend(all_adults_and_young_leaders_groups)

                if custom_field('cf_exec').lower() in ['y', 'yes']:
                    contacts[key]['groups'].extend(exec_groups)
            else:
                # If we are not in the adult group, we must be in a normal section group
                # so, if they are an adult they must be a leader.
                if group.is_leader(member):
                    contacts[key]['groups'].extend(leader_groups)

            if group.is_yl(member):
                contacts[key]['groups'].extend(yl_groups)
            if len(address):
                contacts[key]['addresses'].append(('member', address))

            for _ in ['phone1', 'phone2']:
                number, name = parse_tel(f1(_), _)
                if number != "":
                    contacts[key]['phones'].append((number, name))

        # For all we add the contact details if they say they want to be contacted
        if not (section.lower() == 'adult' or group.is_leader(member)):
            for _ in ('contact_primary_1', 'contact_primary_2'):
                f = functools.partial(get, member=member, section=_)
                _email = f('email1').strip()
                if is_valid_email(f, 'email1'):

                    _first = f('firstname').strip()
                    _last = f('lastname').strip() if len(f('lastname').strip()) > 0 else last_
                    _last_osm = f'{_last} (OSM)'
                    key = '{}-{}-{}'.format(_first.lower(), _last_osm.lower(), _email.lower())
                    address = get_address(f)
                    if key not in contacts:
                        contacts[key] = dict(first=_first,
                                             last=_last_osm,
                                             email=_email.lower(),
                                             addresses=[],
                                             yp=[],
                                             sections=[],
                                             groups=[],
                                             phones=[])

                    contacts[key]['sections'].append(section)
                    contacts[key]['yp'].append(yp_name)
                    contacts[key]['groups'].extend(parent_groups)
                    if len(address):
                        contacts[key]['addresses'].append((f'{_first} {_last}', address))

                    if group.is_yl(member):
                        contacts[key]['groups'].extend(yl_groups)

                    for _ in ['phone1', 'phone2']:
                        number, name = parse_tel(f(_), _)
                        if number != "":
                            contacts[key]['phones'].append((number, name))

    # For every member in the group we want to extract each unique contact address. We ignore all those that
    # not marked for wanting contact.

    # We then add all of the these unique contacts to google.

    # Now we add each of the contacts to the groups that they are associated with.

    log.info("Fetch members from OSM")
    for section_ in sections:
        for member in group.section_all_members(section_):
            add_member_contacts(section_, Group.SECTION_TYPE[section_], member)

    # for c in contacts.values():
    #    print(f'{c}')

    #  remove duplicates
    groups = []
    for key, contact in contacts.items():
        contact['groups'] = set(contact['groups'])
        contact['sections'] = set(contact['sections'])
        contact['yp'] = set(contact['yp'])
        contact['addresses'] = set(contact['addresses'])
        contact['phones'] = set(contact['phones'])
        groups.extend(contact['groups'])

    #  Gather all the groups
    groups = set(groups)
    group_names = [f"{_} (OSM)" for _ in groups]

    contacts = [contact for key, contact in contacts.items()]

    # Sync up the google groups.

    log.info("Fetch list of groups from google")
    existing_osm_groups = fetch_list_of_groups()

    # Fetch the list of group-members - this is used to find who should
    # be managers of the groups.
    existing_role_group_members = fetch_group_members(
        fetch_list_of_groups(prefix='7th-'))

    if delete_google_groups:
        for group in existing_osm_groups:
            log.info(f"Deleting group: {group}")
            delete_google_group(group)
        existing_osm_groups = fetch_list_of_groups()  # should return empty

    missing_groups = [_ for _ in groups if
                      convert_group_name_to_email_address(_) not in existing_osm_groups]
    if len(missing_groups) > 0:
        create_groups(missing_groups)

    for group in groups:
        group_email_address = convert_group_name_to_email_address(group)
        group_moderators = get_group_moderaters(group_email_address,
                                                existing_role_group_members)
        group_members = [contact for contact in contacts
                         if (group in contact['groups'] and
                             contact['email'] not in group_moderators)]
        if len(group_moderators) == 0:
            log.warning(f'No managers for group: {group_email_address}')

        if len(group_members) == 0:
            log.warning(f'No members in group: {group}')

        log.info(f"Syncing contacts in google group: {group}")
        sync_contacts_in_group(
            group_email_address,
            [_['email'] for _ in group_members])

        sync_contacts_in_group(
            group_email_address, group_moderators, role='manager')

    # Sync all contacts to the global directory.
    log.info("delete OSM contacts in directory")
    delete_osm_contacts_already_in_gam()

    log.info("Create all contacts in directory")
    for contact in contacts:
        create_osm_contact_in_gam(contact)

    for google_account in google_accounts:
        if True:  # Flip this to false to manually remove all contacts from the users.

            log.info(f"Syncing OSM contacts into google account for: {google_account}")
            # Setup and sync all contacts and contact groups to a user.
            existing_groups = fetch_list_of_contact_groups(google_account)
            create_contact_groups([_ for _ in group_names if _ not in existing_groups], google_account)

            delete_osm_contacts_already_in_google(google_account)

            for contact in contacts:
                create_osm_contact_in_google(contact, google_account)

        else:
            # To remove all contacts and contact groups from a user.
            existing_groups = fetch_list_of_contact_groups(google_account, field='ContactGroupID')
            delete_osm_contacts_already_in_google(google_account)
            delete_contact_groups(existing_groups, google_account)

    log.info("Finished.")

# $ ~/bin/gamx/gam user admin@7thlichfield.org.uk create contactgroup name 'OSM'
# ~/bin/gamx/gam user admin@7thlichfield.org.uk create contact familyname "Taylor (OSM)" givenname "Richard" email home "r.taylor@bcs.org.uk" primary contactgroup 'OSM' userdefinedfield 'Section' 'Somers Cubs'


def get_group_moderaters(group_email_address, existing_group_members):
    managers = []
    for key, val in GROUP_MODERATOR_MAP.items():
        if re.match(key, group_email_address) is not None:
            for group in val:
                m = [_['email'] for _ in existing_group_members
                     if _['group'] == group]
                managers.extend(m)
    return managers


def convert_group_name_to_email_address(group_name):
    name = group_name.lower()
    name = name.replace('(', '')
    name = name.replace(')', '')
    name = name.replace(' ', '-')
    return f'osm-{name}@7thlichfield.org.uk'


def to_csv(data, headers):
    out = io.StringIO()
    writer = csv.writer(out)
    writer.writerow(headers)
    for row in data:
        writer.writerow(row)
    return out.getvalue()


def to_csv_from_dict(data):
    out = io.StringIO()
    headers = data[0].keys()
    writer = csv.DictWriter(out, fieldnames=headers)
    writer.writeheader()
    for row in data:
        writer.writerow(row)
    return out.getvalue()


def create_osm_contact_in_google(contact, google_account):
    #   ~/bin/gamx/gam user admin@7thlichfield.org.uk create contact familyname "Taylor (OSM)" givenname "Richard" email home "r.taylor@bcs.org.uk" primary contactgroup 'OSM' userdefinedfield 'Section' 'Somers Cubs'
    groups = []
    for group in contact['groups']:
        groups.extend(['contactgroup', f'{group} (OSM)'])
    note = "\n".join(contact['yp'])
    note = note + '\n\n' + f'Last synchronized from OSM: {DATETIME}'
    phones = []
    for phone in contact['phones']:
        phones.extend(['phone', phone[1], phone[0], 'notprimary'])

    output = run_with_output(
        args=[GAM, 'user', google_account, 'create', 'contact',
              'familyname', contact['last'],
              'givenname', contact['first'],
              'note', note,
              'email', 'home', contact['email'], 'primary',
              ] + groups + phones)

    return output


def create_osm_contact_in_gam(contact):
    # groups = []
    # for group in contact['groups']:
    #    groups.extend(['contactgroup', f'{group} (OSM)'])
    note = "\n".join(contact['yp'])
    note = note + '\n\n' + f'Last synchronized from OSM: {DATETIME}'
    phones = []
    for phone in contact['phones']:
        phones.extend(['phone', phone[1], phone[0], 'notprimary'])

    run_with_output(args=[GAM, 'create', 'contact',
                          'familyname', contact['last'],
                          'givenname', contact['first'],
                          'note', note,
                          'email', 'home', contact['email'], 'primary',
                          ] + phones)


def fetch_list_of_contact_groups(google_account, field='ContactGroupName'):
    # ~/bin/gamx/gam user admin@7thlichfield.org.uk print contactgroup
    out = run_with_output(args=[GAM, 'user', google_account, 'print', 'contactgroup'])
    groups = list(csv.DictReader(out.splitlines()))
    osm_groups = []
    if len(groups):
        osm_groups = [_[field] for _ in groups if _['ContactGroupName'].endswith('(OSM)')]
    else:
        print("No contact groups found!")
    return osm_groups


def fetch_list_of_groups(prefix='osm-'):
    out = run_with_output(args=[GAM, 'print', 'groups'])
    groups = list(csv.DictReader(out.splitlines()))
    osm_groups = []
    if len(groups):
        osm_groups = [_['Email'] for _ in groups if _['Email'].startswith(prefix)]
    else:
        print("No groups found!")
    return osm_groups


def fetch_group_members(groups):
    all_members = []
    for group in groups:
        out = run_with_output(args=[GAM, 'print', 'group-members',
                                    'group', group,
                                    'recursive', 'noduplicates'],
                              default="")

        group_members = list(csv.DictReader(out.splitlines()))
        all_members.extend(group_members)

    return all_members


def create_contact_groups(names, google_account):
    #  $ ~/bin/gamx/gam user admin@7thlichfield.org.uk create contactgroup name 'All (OSM)'
    csv_text = to_csv([[_] for _ in names], ['name'])
    run_with_input(args=[GAM, 'csv', '-', 'gam', 'user',
                         google_account, 'create', 'contactgroup',
                         'name', '~name'],
                   inp=csv_text)


def delete_contact_groups(group_ids, google_account):
    csv_text = to_csv([[_] for _ in group_ids], ['ContactGroupID'])
    run_with_input(args=[GAM, 'csv', '-', 'gam', 'user', google_account,
                         'delete', 'contactgroups', '~ContactGroupID'],
                   inp=csv_text)


def delete_google_group(group_name):
    run_with_output(args=[GAM, 'delete', 'group', group_name])


def sync_contacts_in_group(group, group_members_emails, role='member'):
    run_with_input(args=[GAM, 'update', 'group', group, 'sync',
                         role, 'file', '-'],
                   inp="\n".join(group_members_emails),
                   ignore_exc=True)


def create_groups(names):
    defaults = {
        'allowExternalMembers': 'true',
        'whoCanJoin': 'INVITED_CAN_JOIN',
        'whoCanViewMembership': 'ALL_MANAGERS_CAN_VIEW',
        'includeCustomFooter': 'false',
        'defaultMessageDenyNotificationText': 'You do not have permission to post to this group',
        'includeInGlobalAddressList': 'true',
        'archiveOnly': 'false',
        'isArchived': 'true',
        'membersCanPostAsTheGroup': 'false',
        'defaultMessageDenyNotificationText': "Message is being moderated.",
        'allowWebPosting': 'false',
        'messageModerationLevel': 'MODERATE_NONE',
        'replyTo': 'REPLY_TO_IGNORE',
        'sendMessageDenyNotification': 'true',
        'messageModerationLevel': 'MODERATE_ALL_MESSAGES',
        'whoCanContactOwner': 'ALL_IN_DOMAIN_CAN_CONTACT',
        'messageDisplayFont': 'DEFAULT_FONT',
        'whoCanLeaveGroup': 'ALL_MEMBERS_CAN_LEAVE',
        'whoCanAdd': 'ALL_MANAGERS_CAN_ADD',
        'whoCanPostMessage': 'ALL_IN_DOMAIN_CAN_POST',
        'whoCanInvite': 'ALL_MANAGERS_CAN_INVITE',
        'spamModerationLevel': 'MODERATE',
        'whoCanViewGroup': 'ALL_MANAGERS_CAN_VIEW',
        'showInGroupDirectory': 'true',
        'maxMessageBytes': '25000000',
        'customFooterText': (
            'You are receiving this message because you (or your child) is a member '
            'of the 7th Lichfield Scout Group. If you do not want to receive these '
            'messages you can mark your email address in My Scout so that Leader\'s '
            'emails will not be sent. However, if you do so you will not receive '
            'any emails from the Group. Please speak to your section Leader or email '
            'admin@7thlichfield.org.uk if you have any queries.'),
        'allowGoogleCommunication': 'false'
    }
    details = []
    for group_name in names:
        email = convert_group_name_to_email_address(group_name)
        group_entry = {
            'email': f'{email}',
            'name': f'{group_name} (OSM)',
            'description': (f'{group_name} (OSM) is auto-populated from OSM\n'
                            'Any changes that you make will be overwritten next time it is updated.')

        }
        group_entry.update(defaults)
        details.append(group_entry)
    csv_text = to_csv_from_dict(details)
    keys = [(k, f'~{k}') for k in details[0].keys()]
    fields = [_ for sub_list in keys for _ in sub_list]
    del fields[fields.index('email')]
    run_with_input(args=[GAM, 'csv', '-', 'gam', 'create',
                         'group'] + fields,
                   inp=csv_text)


def fetch_osm_contacts_already_in_google(google_account):
    # ~/bin/gamx/gam user admin@7thlichfield.org.uk print contacts
    out = run_with_output(args=[GAM, 'user', google_account, 'print', 'contacts'])
    contacts = list(csv.DictReader(out.splitlines()))
    osm_contacts = [_ for _ in contacts if _['Family Name'].endswith('(OSM)')]
    return osm_contacts


def delete_osm_contacts_already_in_google(google_account):
    # $ ~/bin/gamx/gam user admin@7thlichfield.org.uk delete contacts 2253d1de88ebd5f2
    osm_contacts = fetch_osm_contacts_already_in_google(google_account)
    csv_text = to_csv([[contact['ContactID'], ] for contact in osm_contacts], ['contact'])
    run_with_input(args=[GAM, 'csv', '-', 'gam', 'user', google_account,
                         'delete', 'contacts', '~contact'],
                   inp=csv_text)


def fetch_osm_contacts_already_in_gam():
    out = run_with_output(args=[GAM, 'print', 'contacts'],
                          default="")
    contacts = list(csv.DictReader(out.splitlines()))
    osm_contacts = [_ for _ in contacts if _['Family Name'].endswith('(OSM)')]
    return osm_contacts


def delete_osm_contacts_already_in_gam():
    osm_contacts = fetch_osm_contacts_already_in_gam()
    csv_text = to_csv([[contact['ContactID'], ] for contact in osm_contacts],
                      ['contact'])
    run_with_input(
        args=[GAM, 'csv', '-', 'gam', 'delete', 'contacts', '~contact'],
        inp=csv_text)

if __name__ == '__main__':
    level = logging.INFO

    logging.basicConfig(level=level)

    # groups = [f"{_} (OSM)" for _ in ['test1', 'test2']]
    # existing_groups = fetch_list_of_contact_groups()
    # for group in [_ for _ in groups if _ not in existing_groups]:
    #     create_contact_group(group)

    args = docopt(__doc__, version='OSM 2.0')

    sections = None
    if args['<section>']:
        section = args['<section>']

        if section == 'Group':
            sections = Group.YP_SECTIONS + [Group.ADULT_SECTION, ]
        else:
            assert section in list(Group.SECTIONIDS.keys()) + ['Group'] + list(Group.SECTIONS_BY_TYPE.keys()), \
                "section must be in {!r}.".format(list(Group.SECTIONIDS.keys()) + ['Group'])

            sections = Group.SECTIONS_BY_TYPE[section] if section in Group.SECTIONS_BY_TYPE.keys() else [section, ]

    term = args['--term'] if args['--term'] else None

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    sync_contacts(osm, auth, sections, args['<google_account>'],
                  delete_google_groups=args['--delete-groups'])
