# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  export_group_vcard.py [-d] [--term=<term>] [--email=<address>]
         <apiid> <token> <section>...
  export_group_vcard.py (-h | --help)
  export_group_vcard.py --version


Options:
  <section>      Section to export.
  <outdir>       Output directory for vcard files.
  -d,--debug     Turn on debug output.
  --email=<email> Send to only this email address.
  --term=<term>  Which OSM term to use [default: current].
  -h,--help      Show this screen.
  --version      Show version.

"""

import os.path
import sys
import logging
import itertools
import functools
from csv import writer as csv_writer
from docopt import docopt
import osm

from group import Group, OSM_REF_FIELD
from update import MAPPING

from export_vcards import (
    parse_tel, next_f, get )
#    member2vcard)

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"


def member2contacts(member, section):
    """
    Return up to three records for each member entry.

    1. YP member
    2. Dad
    3. Mum
    """
    # If the email is marked as private do not include it.
    def get_email(email_field):
        return member[email_field].strip().lower() if not (member[email_field].startswith('x ') or
            (member["{}_leaders".format(email_field)] != "yes")) else ""

    ret = [(member['last_name'], member['first_name'], get_email('contact_primary_member.email1')),
           (member['last_name'], member['first_name'], get_email('contact_primary_member.email2'))]

    # No need to continue if we are doing an adult.

    if section == 'Adult' or (
            member.get('Patrol', '') == 'Leaders' and
            int(member.age().days / 365) > 18):
        return ret

    for parent in ['contact_primary_1', 'contact_primary_2']:
        f = functools.partial(get, member=member, section=parent)

        for email in ("email1", "email2"):
            ret.append((f('lastname') if f('lastname').strip() else member['last_name'],
                        f('firstname'),
                        get_email("{}.{}".format(parent,email))))

    return ret


def _main(osm, auth, sections):

    group = Group(osm, auth, MAPPING.keys(), None)

    for section in sections:
        assert section in group.SECTIONIDS.keys(), \
            "section must be in {!r}.".format(group.SECTIONIDS.keys())

    contacts = []

    for section in sections:
        section_contacts = [member2contacts(member, section) for
                          member in group.section_all_members(section)]

        #  flatten list of lists.
        contacts += list(itertools.chain(*section_contacts))

    # Remove blank emails
    contacts = [contact for contact in contacts if contact[2].strip() != "" ]

    # remove duplicates
    by_email = {contact[2]: contact for contact in contacts}
    contacts = list(by_email.values())

    w = csv_writer(sys.stdout)
    w.writerows(contacts)

if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.WARN

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    _main(osm, auth, args['<section>'])
