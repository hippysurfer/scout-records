# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  export_compass.py [-d] [--term=<term>] <apiid> <token> <outdir> <section>... 
  export_compass.py (-h | --help)
  export_compass.py --version


Options:
  <section>      Section to export.
  <outdir>       Output directory for vcard files.
  --term=<term>  Which OSM term to use [default: current].
  -d,--debug     Turn on debug output.
  -h,--help      Show this screen.
  --version      Show version.

"""

import os.path
import logging
import datetime
from docopt import docopt
import osm
import csv
import re

from group import Group, OSM_REF_FIELD
from update import MAPPING

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"

postcode = re.compile("(GIR ?0AA|[A-PR-UWYZ]([0-9]{1,2}|([A-HK-Y][0-9]([0-9ABEHMNPRV-Y])?)|[0-9][A-HJKPS-UW]) ?[0-9][ABD-HJLNP-UW-Z]{2})")

date_format = '%Y-%m-%d'

compass_headings = ("Membership Number",
                    "Look-up Title **",
                    "Forename(s)**",
                    "Surnames **",
                    "Look-up Role **",
                    "Look-up Youth  Leadership",
                    "Date of Birth **",
                    "Address Line 1 **",
                    "Address Line 2",
                    "Address Line 3",
                    "Address Line 4",
                    "Town **",
                    "Postal County",
                    "Postcode **",
                    "Look-up Postal Country **",
                    "Telephone",
                    "Email *",
                    "Start Date for this role**",
                    "Intentionally Blank Column",
                    "Name of Lodge/Six/Patrol",
                    "Date they joined Scouts*",
                    "Look-up Nationality *",
                    "Look-up Ethnicity *",
                    "Look-up Faith/Religion *",
                    "Look-up Parent 1 Title *",
                    "Parent 1 Forename *",
                    "Parent 1 Surname *",
                    "Parent 1 Known As",
                    "Parent 1 Date of Birth",
                    "Look-up Parent 1 Gender *",
                    "Look-up Parent 1 Relationship *",
                    "Parent 1 Email 1 *",
                    "Parent 1 Email 2",
                    "Parent 1 Telephone 1 *",
                    "Parent 1 Telephone 2",
                    "Look-up Parent 1 Gift Aid",
                    "Parent 1 Other Information",
                    "Look-up Parent 2 Title*",
                    "Parent 2 Forename *",
                    "Parent 2 Surname*",
                    "Parent 2 Known As",
                    "Parent 2 Date of Birth",
                    "Look-up Parent 2 Gender*",
                    "Look-up Parent 2 Relationship*",
                    "Parent 2 Email 1 *",
                    "Parent 2 Email 2",
                    "Parent 2 Telephone 1*",
                    "Parent 2 Telephone 2",
                    "Parent 2 Gift Aid",
                    "Parent 2 Other Information",
                    "Emergency Contact Forename *",
                    "Emergency Contact Surname *",
                    "Emergency Contact Known As",
                    "Emergency Contact Relationship *",
                    "Emergency Contact Telephone Number 1 *",
                    "Emergency Contact Telephone Number 2",
                    "Emergency Contact Telephone Number 3",
                    "Doctor / Surgery*",
                    "Surgery Address 1",
                    "Surgery Address 2",
                    "Surgery Address 3",
                    "Surgery Town",
                    "Surgery County",
                    "Surgery Postcode",
                    "Look-up Surgery Country",
                    "Surgery Telephone 1 *",
                    "Surgery Telephone 2",
                    "NHS Number",
                    "Dietary Needs",
                    "Medical Information")

required_headings = ("Look-up Title **",
                    "Forename(s)**",
                    "Surnames **",
                    "Look-up Role **",
                    "Date of Birth **",
                    "Address Line 1 **",
                    "Town **",
                    "Postcode **",
                    "Look-up Postal Country **",
                    "Email *",
                    "Start Date for this role**",
                    "Date they joined Scouts*",
                    "Look-up Nationality *",
                    "Look-up Ethnicity *",
                    "Look-up Faith/Religion *",
                    "Look-up Parent 1 Title *",
                    "Parent 1 Forename *",
                    "Parent 1 Surname *",
                    "Look-up Parent 1 Gender *",
                    "Look-up Parent 1 Relationship *",
                    "Parent 1 Email 1 *",
                    "Parent 1 Telephone 1 *",
                    "Emergency Contact Forename *",
                    "Emergency Contact Surname *",
                    "Emergency Contact Relationship *",
                    "Emergency Contact Telephone Number 1 *",
                    "Doctor / Surgery*",
                    "Surgery Telephone 1 *")


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


def parse_addr(addr):
    a = {}
    # try to split on comma
    l = addr.split(',')

    if len(l) == 1:
        words = l[0].split(" ")
        post_found = False
        if postcode.search(" ".join(words[-2:-1])):
            # Are the last 2 words a postcode?
            a['postcode'] = " ".join(words[-2:-1])
            post_found = True
        else:
            a['postcode'] = "WS13 6ET"

        # Does it contain "lichfield"
        inx = l[0].lower().find("lichfield")
        if inx != -1:
            # Street is all the characters up to "lichfield"
            a['street'] = l[0][0:inx]
            a['town'] = "Lichfield"
        else:
            if post_found:
                # Street is all characters up to postcode
                a['street'] = l[0][0:l[0].find(a['postcode'])]
            else:
                a['street'] = l[0]
                a['town'] = "Lichfield"

        # Just the street?
        a['street'] = l[0]
        a['town'] = "Lichfield"
        a['postcode'] = "WS13 6ET"
    elif len(l) == 2:
        # Is the last one a postcode?
        if postcode.search(l[1]):
            a['street'] = l[0]
            a['town'] = "Lichfield"
            a['postcode'] = l[1]
        elif l[1].lower() == 'lichfield':
            a['street'] = l[0]
            a['town'] = "Lichfield"
            a['postcode'] = "WS13 6ET"
        else:
            a['street'] = " ".join(l)
            a['town'] = "Lichfield"
            a['postcode'] = "WS13 6ET"
    elif len(l) == 3:
        # Probably got all 3
        a['street'] = l[0]
        a['town'] = l[1]
        a['postcode'] = l[2]
    elif len(l) > 3:
        # Is the last one a postcode?
        if postcode.search(l[-1]):
            a['street'] = ",".join(l[:-2])
            a['town'] = l[-2]
            a['postcode'] = l[-1]
        else:
            a['street'] = ",".join(l[:-1])
            a['town'] = l[-1]
            a['postcode'] = "WS13 6ET"
    return a


def get_tel(member, priority_list):
    for i in priority_list:
        if member[i].strip() != '':
            tel = parse_tel(member[i], '')[0]
            if not tel.startswith('0'):
                tel = "01543 {}".format(tel)
            return tel
    return "01543 123456"

def get_parent(member):
    p = {}
    if (member['MumEmail'].strip() != ''
        and not member['MumEmail'].startswith('x ')):
        p['title'] = "Ms"
        p['forename'] = member['MumsName'] \
                        if member['MumsName'].strip() != ''\
                           else member['DadsName']
        p['surname'] = member['lastname']
        p['gender'] = 'Unknown'
        p['relationship'] = 'Mother'
        p['email'] = member['MumEmail']
        p['tel'] = get_tel(member, ('HomeTel', 'MumMob', 'DadMob'))
    else:
        p['title'] = "Mr"
        p['forename'] = member['DadsName'] \
                        if member['DadsName'].strip() != ''\
                           else member['MumsName']
        p['surname'] = member['lastname']
        p['gender'] = 'Unknown'
        p['relationship'] = 'Father'
        p['email'] = member['DadEmail']
        p['tel'] = get_tel(member, ('HomeTel', 'DadMob', 'MumMob'))

    parts = p['forename'].split(' ')
    if len(parts) > 1:
        if parts[-1].strip() == p['surname']:
            p['forename'] = " ".join(parts[:-1])

    return p


def member2compass(member, section):
    j = {}

    j["Look-up Title **"] = 'Mr' \
                            if (member['Sex'].lower() == 'm' or
                                member['Sex'].lower() == 'male') else 'Miss'
    j["Forename(s)**"] = member['firstname'].strip()
    j["Surnames **"] = member['lastname'].strip()

    if section in ['Paget', 'Garrick', 'Brown']:
        section = 'Beaver Scout'
    elif section in ['Maclean', 'Rowallan', 'Somers']:
        section = 'Cub Scout'
    elif section in ['Boswell', 'Johnson', 'Erasmus']:
        section = 'Scout'
    else:
        log.warn("unkown section type: {}".format(section))
        section = 'UNKNOWN'

    j["Look-up Role **"] = section

    j["Date of Birth **"] = datetime.datetime.strptime(
        member['dob'], '%d/%m/%Y').strftime(date_format)

    addr = parse_addr(member['PrimaryAddress'])
    j["Address Line 1 **"] = addr['street']
    j["Town **"] = addr['town']
    j["Postcode **"] = addr['postcode']
    j["Look-up Postal Country **"] = 'United Kingdom'

    j["Email *"] = member['MumEmail'] \
        if member['MumEmail'].strip() != '' \
        else member['DadEmail']

    j["Start Date for this role**"] = datetime.datetime.strptime(
        member['started'], '%d/%m/%Y').strftime(date_format)

    j["Date they joined Scouts*"] = datetime.datetime.strptime(
        member['joined'], '%d/%m/%Y').strftime(date_format)

    j["Look-up Nationality *"] = "British"
    j["Look-up Ethnicity *"] = "English/Welsh/Scottish/Northern Irish/British"
    j["Look-up Faith/Religion *"] = "No religion"

    p = get_parent(member)
    j["Look-up Parent 1 Title *"] = p['title']
    j["Parent 1 Forename *"] = p['forename']
    j["Parent 1 Surname *"] = p['surname']
    j["Look-up Parent 1 Gender *"] = p['gender']
    j["Look-up Parent 1 Relationship *"] = p['relationship']
    j["Parent 1 Email 1 *"] = p['email']
    j["Parent 1 Telephone 1 *"] = p['tel']

    j["Emergency Contact Forename *"] = j["Parent 1 Forename *"]
    j["Emergency Contact Surname *"] = j["Parent 1 Surname *"]
    j["Emergency Contact Relationship *"] = \
                                j["Look-up Parent 1 Relationship *"]
    j["Emergency Contact Telephone Number 1 *"] = j["Parent 1 Telephone 1 *"]
    j["Doctor / Surgery*"] = 'Unknown'
    j["Surgery Telephone 1 *"] = '01543 123456'

    return j


def check(entry, section):
    # Check that all required fields have something in them.
    for k in required_headings:
        if entry[k].strip() == '':
            log.warn("{}: {} {} - emtpy field: {}".format(
                section,
                entry["Forename(s)**"],
                entry["Surnames **"],
                k))


def _main(osm, auth, sections, outdir, term):

    assert os.path.exists(outdir) and os.path.isdir(outdir)

    group = Group(osm, auth, MAPPING.keys(), term)

    for section in sections:
        assert section in group.SECTIONIDS.keys(), \
            "section must be in {!r}.".format(group.SECTIONIDS.keys())

    for section in sections:
        entries = [member2compass(member, section) for
                   member in group.section_yp_members_without_leaders(section)]

        [check(entry, section) for entry in entries]

        with open(os.path.join(outdir, section + ".csv"), "w") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=compass_headings)
            writer.writeheader()
            [writer.writerow(entry) for entry in entries]

if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.INFO

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    if args['--term'] in [None, 'current']:
        args['--term'] = None

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    _main(osm, auth, args['<section>'], args['<outdir>'],
          args['--term'])























