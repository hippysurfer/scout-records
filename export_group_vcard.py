# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  export_group_vcard.py [-d] [--term=<term>] [--email=<address>] <apiid> <token> <outdir> <section>... 
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
import logging
import itertools
import socket
import functools
import smtplib
from docopt import docopt
import osm
import vobject as vo

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from group import Group, OSM_REF_FIELD
from update import MAPPING

from export_vcards import (
    parse_tel, next_f, get,
    member2vcard)

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"

FROM = "Richard Taylor <r.taylor@bcs.org.uk>"

def send(to, subject, vcards, fro=FROM):

    for dest in to:
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = fro
        msg['To'] = dest

        body = MIMEText(vcards, 'vcard')
        body.add_header('Content-Disposition', 'attachment',
                        filename="group.vcf")
        msg.attach(body)

        hostname = 'www.thegrindstone.me.uk' \
            if not socket.gethostname() == 'rat' \
            else 'localhost'

        s = smtplib.SMTP(hostname)

        try:
            s.sendmail(fro, dest, msg.as_string())
        except:
            log.error(msg.as_string(),
                      exc_info=True)

        s.quit()


def member2vcards(member, section):
    """
    Create up to three vcards for each member entry.

    1. YP member
    2. Dad
    3. Mum
    """

    ret = [member2vcard(member, section)]

    # No need to continue if we are doing an adult.

    if section == 'Adult' or (
            member.get('Patrol', '') == 'Leaders' and
            int(member.age().days / 365) > 18):
        return ret

    for parent in ['contact_primary_1', 'contact_primary_2']:
        f = functools.partial(get, member=member, section=parent)

        full_name = "{} {}".format(f('firstname'),
                                   f('lastname'))

        if full_name.strip() != '':
            ret += [process_parent(full_name, member, section, f), ]

    return ret


def process_parent(name, member, section, f):

    # If the name does not appear to have lastname part
    # add it from the member name.
    if len(name.strip().split(' ')) < 2:
        name = "{} {}".format(name.strip(), member['last_name'])

    j = vo.vCard()

    uid = j.add('UID')
    uid.value = "{}{}.OSM@thegrindstone.me.uk".format(
        name.replace(" ", ""),
        member[OSM_REF_FIELD])

    j.add('n')
    j.n.value = vo.vcard.Name(
        family=f('lastname') if f('lastname').strip() else member['last_name'],
        given=f('firstname'))
    j.add('fn')
    j.fn.value = name

    next_ = next_f(j, 0).next_f

    for _ in ['phone1', 'phone2']:
        number, name = parse_tel(f(_), _)
        next_('tel', name, number)

    for _ in ['email1', 'email2']:
        next_('email', _, f(_))

    next_('adr', 'Primary',
          vo.vcard.Address(
              street=f('address1'),
              city=f('address2'),
              region=f('address3'),
              country=f('address4'),
              code=f('postcode')))

    org = j.add('org')
    org.value = [section, ]

    note = j.add('note')
    note.value = "Child: {} {} ({})\n".format(
        member['first_name'], member['last_name'], section)

    cat = j.add('CATEGORIES')
    cat.value = ("7th", "7th Lichfield Parent")

    return j.serialize()


def _main(osm, auth, sections, outdir, email, term):

    assert os.path.exists(outdir) and os.path.isdir(outdir)

    group = Group(osm, auth, MAPPING.keys(), term)

    for section in sections:
        assert section in group.SECTIONIDS.keys(), \
            "section must be in {!r}.".format(group.SECTIONIDS.keys())

    vcards = []

    for section in sections:
        section_vcards = [member2vcards(member, section) for
                          member in group.section_all_members(section)]

        #  flatten list of lists.
        vcards += list(itertools.chain(*section_vcards))

    open(os.path.join(outdir, "group.vcf"), 'w').writelines(vcards)

    if email:
        send([email, ], "OSM Group vcards", "".join(vcards))

if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.INFO

    if args['--term'] in [None, 'current']:
        args['--term'] = None

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    _main(osm, auth, args['<section>'],
          args['<outdir>'], args['--email'], args['--term'])























