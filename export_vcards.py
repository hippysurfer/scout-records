# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  export_vcards.py [-d] <apiid> <token> <outdir> <section>... 
  export_vcards.py (-h | --help)
  export_vcards.py --version


Options:
  <section>      Section to export.
  <outdir>       Output directory for vcard files.
  -d,--debug     Turn on debug output.
  -h,--help      Show this screen.
  --version      Show version.

"""

import os.path
import logging
import datetime
from docopt import docopt
import osm
import vobject as vo

from group import Group, OSM_REF_FIELD
from update import MAPPING

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"


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

def member2vcard(member, section):
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

    item2 = j.add('X-ABLabel', 'item2')
    item2.value = 'Mum' if section != 'Adult' else "NOK"
    me = j.add('email', 'item2')
    me.value = member['MumEmail' if section != 'Adult' else 'NOKEmail1']
    me.type_param = 'INTERNET'

    item3 = j.add('X-ABLabel', 'item3')
    item3.value = 'Dad' if section != 'Adult' else 'NOK2'
    de = j.add('email', 'item3')
    de.value = member['DadEmail' if section != 'Adult' else 'NOKEmail2']
    de.type_param = 'INTERNET'

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

    if section == 'Adult':
        item10 = j.add('X-ABLabel', 'item10')
        item10.value = 'NOK Address 2'
        addr = j.add('adr', 'item10')
        addr.value = vo.vcard.Address(
            street=member['NOKAddress2'])

    bday = j.add('bday')
    bday.value = datetime.datetime.strptime(
        member['dob'], '%d/%m/%Y').strftime('%Y-%m-%d')

    note = j.add('note')
    if section == 'Adult':
        note.value = "NOKs: {}\nMedical: {}\nNotes: {}\n".format(
            member['NextofKinNames'], member['Medical'], member['Notes'])
    else:
        note.value = "Parents: Dad - {} Mum - {}\nMedical: {}\nNotes: {}\n".format(
            member['DadsName'], member['MumsName'], member['Medical'], member['Notes'])

    return j.serialize()


def _main(osm, auth, sections, outdir):

    assert os.path.exists(outdir) and os.path.isdir(outdir)

    group = Group(osm, auth, MAPPING.keys())

    for section in sections:
        assert section in group.SECTIONIDS.keys(), \
            "section must be in {!r}.".format(group.SECTIONIDS.keys())

    for section in sections:
        vcards = [member2vcard(member, section) for
                  member in group.section_all_members(section)]

        open(os.path.join(outdir, section + ".vcf"), 'w').writelines(vcards)

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

    _main(osm, auth, args['<section>'], args['<outdir>'])























