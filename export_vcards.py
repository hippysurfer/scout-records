# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  export_vcards.py [-d] [--term=<term>] <apiid> <token> <outdir> <section>... 
  export_vcards.py (-h | --help)
  export_vcards.py --version


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
import functools
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


class next_f:
    def __init__(self, card, n=0):
        self._card = card
        self._last_f = n

    def next_f(self, type_, label, value, label_type='X-ABLabel'):
        if not value:
            return
        self._last_f += 1
        item = "item{}".format(self._last_f)
        i = self._card.add(label_type, item)
        i.value = label
        v = self._card.add(type_, item)
        if type_ == 'email':
            v.type_paramlist = ['INTERNET']
        v.encoded = True
        v.value = value


def get(field, member, section):
    return member["{}.{}".format(section, field)]


def add_base(next_, member, label, section, field_func=None):
    f = field_func if field_func else functools.partial(get, 
                                                        member=member, 
                                                        section=section)

    for _ in ['phone1', 'phone2']:
        number, name = parse_tel(f(_), label)
        next_('tel', name, number)

    for _ in ['email1', 'email2']:
        next_('email', label, f(_))

    next_('adr', label,
          vo.vcard.Address(
              street=f('address1'),
              city=f('address2'),
              region=f('address3'),
              country=f('address4'),
              code=f('postcode')))


def add_contact(next_, member, label, section):
    f = functools.partial(get, member=member, section=section)

    full_name = "{} {} ({})".format(f('firstname'),
                                    f('lastname'),
                                    label)

    add_base(next_, member, full_name, section, f)


def member2vcard(member, section):
    j = vo.vCard()

    uid = j.add('UID')
    uid.value = "{}.OSM@thegrindstone.me.uk".format(
        member[OSM_REF_FIELD])

    j.add('n')
    j.n.value = vo.vcard.Name(family=member['last_name'],
                              given=member['first_name'])
    j.add('fn')
    j.fn.value = '{} {}'.format(member['first_name'],
                                member['last_name'])

    next_ = next_f(j, 0).next_f

    add_base(next_, member,
                'Member' if section != 'Adult' else 'NOK1',
                'contact_primary_member')
    add_contact(next_, member,
                'Parent1' if section != 'Adult' else 'NOK1',
                'contact_primary_1')
    add_contact(next_, member,
                'Parent2' if section != 'Adult' else 'NOK2',
                'contact_primary_2')
    add_contact(next_, member, 'Emergency', 'emergency')

    org = j.add('org')
    org.value = [section, ]

    bday = j.add('bday')
    bday.value = datetime.datetime.strptime(
        member['date_of_birth'], '%Y-%m-%d').strftime('%Y-%m-%d')

    f = functools.partial(get, member=member, section='contact_primary_1')
    parent1 = "{} {}".format(f('firstname'),
                             f('lastname'))

    f = functools.partial(get, member=member, section='contact_primary_2')
    parent2 = "{} {}".format(f('firstname'),
                             f('lastname'))

    f = functools.partial(get, member=member, section='customisable_data')
    medical = f('medical')
    notes = f('notes')

    note = j.add('note')

    cat = j.add('CATEGORIES')

    if section == 'Adult':
        note.value = "NOKs: {} / {}\nMedical: {}\nNotes: {}\nSection: {}\n".format(
            parent1, parent2, medical, notes, section
        )
        cat.value = ('7th', '7th Adult', '7th {}'.format(section))
    elif member.get('patrol', '') == 'Leaders':
        if int(member.age().days / 365) > 18:
            note.value = "NOKs: {} / {}\nMedical: {}\nNotes: {}\nRole: {}\nSection: {}\n".format(
                parent1, parent2, medical, notes, member.get('patrol', ''), section
            )
            cat.value = ('7th', '7th Section Leader', '7th {} Leader'.format(section))
        else:
            note.value = "Parents: {} / {}\nMedical: {}\nNotes: {}\nPatrol: {}\nSection: {}\n".format(
                parent1, parent2, medical, notes, member.get('patrol', ''), section
            )
            cat.value = ('7th', '7th YL', '7th {} YL'.format(section))
            
    else:
        note.value = "Parents: {} / {}\nMedical: {}\nNotes: {}\nPatrol: {}\nSection: {}\n".format(
            parent1, parent2, medical, notes, member.get('patrol', ''), section
        )
        cat.value = ('7th', '7th YP', '7th {} YP'.format(section))

    return j.serialize()


def _main(osm, auth, sections, outdir, term):

    assert os.path.exists(outdir) and os.path.isdir(outdir)

    group = Group(osm, auth, MAPPING.keys(), term)

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
        level = logging.WARN

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    if args['--term'] in [None, 'current']:
        args['--term'] = None

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    _main(osm, auth, args['<section>'], args['<outdir>'],
          args['--term'])























