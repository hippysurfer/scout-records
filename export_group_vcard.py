# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  export_group_vcard.py [-d] [--email=<address>] <apiid> <token> <outdir> <section>... 
  export_group_vcard.py (-h | --help)
  export_group_vcard.py --version


Options:
  <section>      Section to export.
  <outdir>       Output directory for vcard files.
  -d,--debug     Turn on debug output.
  --email=<email> Send to only this email address.
  -h,--help      Show this screen.
  --version      Show version.

"""

import os.path
import logging
import datetime
import itertools
import socket
import smtplib
from docopt import docopt
import osm
import vobject as vo

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from group import Group, OSM_REF_FIELD
from update import MAPPING

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


def parse_tel(number_field, default_name):
    index = 0
    for i in range(len(number_field)):
        if number_field[i] not in ['0', '1', '2', '3', '4', '5',
                                   '6', '7', '8', '9', ' ', '\t']:
            index = i
            break

    number = number_field[:index].strip() if index != 0 else number_field
    name = number_field[index:].strip() if index != 0 else default_name

    #print("input = {}, index = {}, number = {}, name = {}".format(
    #    number_field, index, number, name))

    return number, name

def member2vcards(member, section):
    """
    Create up to three vcards for each member entry.

    1. YP member
    2. Dad
    3. Mum
    """

            
    # Produce primary vcard
    j = vo.vCard()
    
    uid = j.add('UID')
    uid.value = "{}.OSM@thegrindstone.me.uk".format(
        member['PersonalReference'])

    j.add('n')
    j.n.value = vo.vcard.Name(family=member['lastname'],
                              given=member['firstname'])
    j.add('fn')
    j.fn.value = '{} {}'.format(member['firstname'],
                                member['lastname'])

    item1 = j.add('X-ABLabel', 'item1')
    item1.value = 'Personal'
    pe = j.add('email', 'item1')
    pe.value = member['PersonalEmail']
    pe.type_paramlist = ['INTERNET', 'pref']

    org = j.add('org')
    org.value = [section, ]

    number, name = parse_tel(member['PersonalMob'],
                             'Personal Mob')
    item4 = j.add('X-ABLabel', 'item4')
    item4.value = name
    ptel = j.add('tel', 'item4')
    ptel.value = number

    number, name = parse_tel(member['MumMob' if section != 'Adult'
                                    else 'NOKMob1'],
                             'Mum Mob' if section != 'Adult' else 'NOK Mob1')
    item5 = j.add('X-ABLabel', 'item5')
    item5.value = name
    mtel = j.add('tel', 'item5')
    mtel.value = number

    number, name = parse_tel(member['DadMob' if section != 'Adult'
                                    else 'NOKMob2'],
                             'Dad Mob' if section != 'Adult' else 'NOK Mob2')
    item6 = j.add('X-ABLabel', 'item6')
    item6.value = name
    dtel = j.add('tel', 'item6')
    dtel.value = number

    item7 = j.add('X-ABLabel', 'item7')
    item7.value = 'Primary Address'
    addr = j.add('adr', 'item7')
    addr.value = vo.vcard.Address(street=member['PrimaryAddress'])

    item8 = j.add('X-ABLabel', 'item8')
    item8.value = 'Secondary Address' \
                  if section != 'Adult' else 'NOK Address 1'
    addr = j.add('adr', 'item8')
    addr.value = vo.vcard.Address(
        street=member['SecondaryAddress'
                      if section != 'Adult' else 'NOKAddress1'])

    number, name = parse_tel(member['HomeTel'],
                             'Home')
    item9 = j.add('X-ABLabel', 'item9')
    item9.value = name
    htel = j.add('tel', 'item9')
    htel.value = number


    bday = j.add('bday')
    bday.value = datetime.datetime.strptime(
        member['dob'], '%d/%m/%Y').strftime('%Y-%m-%d')

    cat = j.add('CATEGORIES')

    note = j.add('note')
    if section == 'Adult':
        note.value = "NOKs: {}\nMedical: {}\nNotes: {}\nSection: {}\n".format(
            member['NextofKinNames'], member['Medical'], member['Notes'], section)
        cat.value = ("7th", "7th Adult")
    elif member.get('Patrol','') == 'Leaders':
        note.value = "NOKs: {}\nMedical: {}\nNotes: {}\nRole:{}\nSection: {}\n".format(
            member['Parents'], member['Medical'], member['Notes'], member.get('Patrol',''), section)
        cat.value = ("7th", "7th Section Leader")
    else:
        cat.value = ("7th", "7th YP")
        note.value = "Parents: {}\nMedical: {}\nNotes: {}\nPatrol:{}\nSection: {}\n".format(
            member['Parents'], member['Medical'], member['Notes'], member.get('Patrol',''), section)


    ret = [j.serialize(), ]

    # No need to continue if we are doing an adult.

    if section == 'Adult' or member.get('Patrol','') == 'Leaders':
        return ret


    # Try to work out the parents entries.

    try:
        if member['parents'].count('&') != 0:
            sep = "&"
        else:
            sep = "/"
        (dad, mum) = member['parents'].split(sep, 2)
        dad.strip()
        mum.strip()
    except ValueError:
        # This means the parents field does not have a seperator.
        # If it is not empty and there is an email address for Dad 
        # we assume that it is dad, otherwise we assume it is Mum
        if ( member['parents'].strip() != '' and
             member['DadEmail'].strip() != '' ):
            dad = member['parents'].strip()
            mum = None
        elif ( member['parents'].strip() != '' and
               member['MumEmail'].strip() != '' ):
            mum = member['parents'].strip()
            dad = None
        else:
            mum = None
            dad = None

    if dad:
        ret += [process_parent(member, dad,
                               member['DadEmail'],
                               member['DadMob'],
                               section),]

    if mum:
        ret += [process_parent(member, mum,
                               member['MumEmail'],
                               member['MumMob'],
                               section),]
    
    return ret

def process_parent(member, parent, email, mob, section):

    # Try to add surname if is missing.
    if not len(parent.strip().split(" ")) > 1:
        parent = "{} {}".format(parent, member['lastname'].strip())
        
    j = vo.vCard()
    
    uid = j.add('UID')
    uid.value = "{}{}.OSM@thegrindstone.me.uk".format(
        parent.replace(" ",""),
        member['PersonalReference'])

    j.add('n')
    j.n.value = vo.vcard.Name(family=parent.split(' ', 1)[1].strip(),
                              given=parent.split(' ')[0].strip())
    j.add('fn')
    j.fn.value = '{} {}'.format(parent.split(' ')[0].strip(),
                                parent.split(' ', 1)[1].strip())

    item1 = j.add('X-ABLabel', 'item1')
    item1.value = 'Personal'
    pe = j.add('email', 'item1')
    pe.value = email
    pe.type_paramlist = ['INTERNET', 'pref']

    org = j.add('org')
    org.value = [section, ]

    number, name = parse_tel(mob,
                             'Personal Mob')
    item4 = j.add('X-ABLabel', 'item4')
    item4.value = name
    ptel = j.add('tel', 'item4')
    ptel.value = number

    item7 = j.add('X-ABLabel', 'item7')
    item7.value = 'Primary Address'
    addr = j.add('adr', 'item7')
    addr.value = vo.vcard.Address(street=member['PrimaryAddress'])

    item8 = j.add('X-ABLabel', 'item8')
    item8.value = 'Secondary Address' \
                  if section != 'Adult' else 'NOK Address 1'
    addr = j.add('adr', 'item8')
    addr.value = vo.vcard.Address(
        street=member['SecondaryAddress'
                      if section != 'Adult' else 'NOKAddress1'])

    number, name = parse_tel(member['HomeTel'],
                             'Home')
    item9 = j.add('X-ABLabel', 'item9')
    item9.value = name
    htel = j.add('tel', 'item9')
    htel.value = number


    note = j.add('note')
    note.value = "Child: {} {} ({})\n".format(
        member['firstname'], member['lastname'], section)

    cat = j.add('CATEGORIES')
    cat.value = ("7th", "7th Lichfield Parent")

    return j.serialize()


def _main(osm, auth, sections, outdir, email):

    assert os.path.exists(outdir) and os.path.isdir(outdir)

    group = Group(osm, auth, MAPPING.keys())

    for section in sections:
        assert section in group.SECTIONIDS.keys(), \
            "section must be in {!r}.".format(group.SECTIONIDS.keys())

    vcards = []

    for section in sections:
        section_vcards = [member2vcards(member, section) for
                          member in group.section_all_members(section)]


        # flatten list of lists.
        vcards += list(itertools.chain(*section_vcards))

    open(os.path.join(outdir, "group.vcf"), 'w').writelines(vcards)

    if email:
        send([email,], "OSM Group vcards", "".join(vcards))

if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.INFO

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    _main(osm, auth, args['<section>'],
          args['<outdir>'], args['--email'])























