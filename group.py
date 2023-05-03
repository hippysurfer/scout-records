from datetime import datetime
from dateutil import relativedelta
from collections import OrderedDict
import logging
import osm

log = logging.getLogger(__name__)

# OSM_REF_FIELD = 'customisable_data.membershipno'
# OSM_REF_FIELD = 'customisable_data.PersonalReference'
OSM_REF_FIELD = 'member_id'


class Member(osm.Member):

    def age(self, ref_date=datetime.now()):
        try:
            dob = datetime.strptime(
                self['date_of_birth'], '%Y-%m-%d')
        except ValueError:
            log.error(f"Invalid date_of_birth {self['date_of_birth']} in record {self}")
            raise
        return ref_date - dob

    def age_in_years_and_months(self):
        now = datetime.now()
        real_dob = datetime.strptime(self['date_of_birth'], '%Y-%m-%d')
        rel_age = relativedelta.relativedelta(now, real_dob)
        return "{}.{}".format(rel_age.years, rel_age.months)

osm.MemberClass = Member


class Group(object):
    SECTIONIDS = OrderedDict((
        ('Saturn', '69414'),
        ('Paget', '9960'),
        ('Swinfen', '17326'),
        ('Garrick', '20711'),
        ('Maclean', '14324'),
        ('Rowallan', '12700'),
        ('Somers', '20706'),
        ('Boswell', '10363'),
        ('Johnson', '5882'),
        ('Erasmus', '20707'),
        ('Adult', '18305'),
        ('Subs', '33593')))
    # 'Waiting List': ""}

    ADULT_SECTION = 'Adult'
    SUBS_SECTION = "Subs"

    YP_SECTIONS = ['Saturn',
                   'Paget',
                   'Swinfen',
                   'Garrick',
                   'Maclean',
                   'Rowallan',
                   'Somers',
                   'Boswell',
                   'Johnson',
                   'Erasmus']

    SECTION_TYPE = {
        'Adult': 'adult',
        'Subs': 'adult',
        'Saturn': 'squirrels',
        'Paget': 'beavers',
        'Swinfen': 'beavers',
        'Maclean': 'cubs',
        'Rowallan': 'cubs',
        'Boswell': 'scouts',
        'Johnson': 'scouts',
        'Garrick': 'beavers',
        'Erasmus': 'scouts',
        'Somers': 'cubs'
    }

    # Create a reverse mapping of the Sections to their types.
    SECTIONS_BY_TYPE = OrderedDict((
        ('squirrels', ['Saturn']),
        ('beavers', ['Paget', 'Swinfen', 'Garrick']),
        ('cubs', ['Maclean', 'Rowallan', 'Somers']),
        ('scouts', ['Boswell', 'Johnson', 'Erasmus'])))

    MIN_AGE = {
        'Saturn': 4,
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
        'Saturn': 6,
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

    def __init__(self, osm, auth, important_fields, term=None, on_date=None,
                 include_yl_as_yp=True, object_types=osm.ALL_OBJECTS):
        self._osm = osm
        self._important_fields = important_fields
        self._sections = self._osm.OSM(auth, self.SECTIONIDS.values(),
                                       term, on_date, object_types=object_types)
        self.include_yl_as_yp = include_yl_as_yp

    def section_all_members(self, section):
        # If there a no members the 'members' will be an empty list
        # rather than an empty dict so we trap this an return an empty
        # list.
        try:
            return self._sections.sections[
                self.SECTIONIDS[section]].members.values()
        except AttributeError:
            return []

    def all_adult_members(self):
        return self.section_all_members('Adult')

    def all_adult_references(self):
        return [member[OSM_REF_FIELD] for member in self.all_adult_members()]

    def all_section_references(self, section):
        return [member[OSM_REF_FIELD]
                for member in
                self.section_all_members(section)]

    def _section_missing_references(self, section):
        return [member for member in
                self.section_all_members(section)
                if member[OSM_REF_FIELD].strip() == ""]

    def missing_adult_references(self):
        return self._section_missing_references('Adult')

    def all_yp_members_dict(self):
        return {s: self.section_all_members(s) for
                s in self.YP_SECTIONS}

    def all_yl_members_dict(self):
        return {s: self.section_yl_members(s) for
                s in self.YP_SECTIONS}

    def all_yp_members_without_leaders_dict(self):
        return {s: self.section_yp_members_without_leaders(s) for
                s in self.YP_SECTIONS}

    def all_yp_members_without_senior_duplicates_dict(self):
        return {'Paget': self.remove_senior_duplicates('Paget',
                                                       self.all_cubs()),
                'Swinfen': self.remove_senior_duplicates('Swinfen',
                                                         self.all_cubs()),
                'Maclean': self.remove_senior_duplicates('Maclean',
                                                         self.all_scouts()),
                'Rowallan': self.remove_senior_duplicates('Rowallan',
                                                          self.all_scouts()),
                'Garrick': self.remove_senior_duplicates('Garrick',
                                                         self.all_scouts()),
                'Somers': self.remove_senior_duplicates('Somers',
                                                        self.all_scouts()),
                'Erasmus': self.section_yp_members_without_leaders('Erasmus'),
                'Boswell': self.section_yp_members_without_leaders('Boswell'),
                'Johnson': self.section_yp_members_without_leaders('Johnson')}

    def all_yp_members_without_senior_duplicates(self):
        all_section = self.all_yp_members_without_senior_duplicates_dict()
        all = []
        for s in self.YP_SECTIONS:
            all.extend(all_section[s])
        return all

    def section_missing_references(self, section_name):
        return self._section_missing_references(section_name)

    def set_yl_as_yp(self, yes):
        """Include Young Leaders as Young People"""
        self.include_yl_as_yp = yes

    def get_yp_patrol_exclude_list(self):
        """
        :return: The list of patrol names that should be considered as not YP.
        :rtype:
        """
        l = ['leaders', 'winter adv.']
        if self.include_yl_as_yp:
            l += ['young leaders', ]
        return l

    def is_yl(self, member):
        """
        Try to work out if the member is a Young Leader or not.

        :param member:
        :return: True or False
        """
        return ((member['patrol'].lower() in
                 self.get_yp_patrol_exclude_list()) and
                ((member.age().days / 365 > 15) and
                 (member.age().days / 365 < 18)) or
                ((member.age().days / 365 > 15.8) and
                 (member.age().days / 365 < 18)))

    def is_scout_helper(self, member):
        """
                Try to work out if the member is a Scout helper in a junior section.

                :param member:
                :return: True or False
                """
        return ((member['patrol'].lower() in
                 self.get_yp_patrol_exclude_list()) and
                (member.age().days / 365 <= 15))

    def is_leader(self, member):
        """
        Try to work out if the member is a Leader or not.

        :param member:
        :return: True or False
        """
        return member.age().days / 365 > 18

    def is_yp(self, member):
        return not (self.is_yl(member) or self.is_leader(member) or self.is_scout_helper(member))

    def section_yp_members_without_leaders(self, section):
        return [member for member in
                self.section_all_members(section)
                if self.is_yp(member)]

    def section_yl_members(self, section):
        return [member for member in
                self.section_all_members(section)
                if self.is_yl(member)]

    def section_leaders_in_yp_section(self, section):
        return [member for member in
                self.section_all_members(
                    section)
                if self.is_leader(member)]

    def all_leaders_in_yp_sections(self):
        # Make a list of all the leaders
        all_leaders = []
        for section in self.YP_SECTIONS:
            all_leaders.extend(
                self.section_leaders_in_yp_section(section))
        return all_leaders

    def all_yp_members_without_leaders(self):
        all_yps = []
        for section in self.YP_SECTIONS:
            all_yps.extend(
                self.section_yp_members_without_leaders(section))
        return all_yps

    def all_yl_members(self):
        all_yls = []
        for section in self.YP_SECTIONS:
            all_yls.extend(
                self.section_yl_members(section))
        return all_yls

    def all_beavers(self):
        return self.section_yp_members_without_leaders('Paget') +\
            self.section_yp_members_without_leaders('Swinfen') +\
            self.section_yp_members_without_leaders('Garrick')

    def all_cubs(self):
        return self.section_yp_members_without_leaders('Rowallan') +\
            self.section_yp_members_without_leaders('Maclean') + \
            self.section_yp_members_without_leaders('Somers')

    def all_scouts(self):
        return self.section_yp_members_without_leaders('Johnson') +\
            self.section_yp_members_without_leaders('Boswell') +\
            self.section_yp_members_without_leaders('Erasmus')

    def find_ref_in_sections(self, reference, exclude_sections=None):
        """Search for a reference in all of the sections.

        return a list of section names."""

        matching_sections = []

        for section in [_ for _ in self.SECTIONIDS.keys() if
                _ not in (exclude_sections if exclude_sections else [])]:
            if reference in self.all_section_references(
                    section):
                matching_sections.append(section)

        return matching_sections

    def find_sections_by_name(self, firstname, lastname):
        """
        Return a list of sections that contain records with matching names.
        """
        l = []
        sections = self.all_yp_members_dict()
        for section in sections.keys():
            for member in sections[section]:
                osm_firstname = member['first_name'].lower().strip()
                if (osm_firstname.lower().strip() ==
                        firstname.lower().strip() and
                    member['last_name'].lower().strip() ==
                        lastname.lower().strip()):
                    l.append(section)
        return l

    def find_by_name(self, firstname, lastname, section_wanted=None,
                     ignore_second_name=False):
        """Return a list of records with matching names"""
        l = []
        sections = self.all_yp_members_dict()
        for section in sections.keys():
            if (section_wanted and section_wanted != section):
                continue
            for member in sections[section]:
                osm_firstname = member['first_name'].lower().strip()
                if ignore_second_name:
                    osm_firstname = osm_firstname.split(' ')[0]
                    firstname = firstname.split(' ')[0]
                if (osm_firstname.lower().strip() ==
                        firstname.lower().strip() and
                    member['last_name'].lower().strip() ==
                        lastname.lower().strip()):
                    l.append(member)
        return l

    #def find_by_ref(self, ref, section_wanted=None):
    #    """Return a list of records with matching refs"""
    #    l = []
    #    sections = self.all_yp_members_dict()
    #    for section in sections.keys():
    #        if (section_wanted and section_wanted != section):
    #            continue
    #        for member in sections[section]:
    #            if (member[OSM_REF_FIELD].lower().strip() ==
    #                    ref.lower().strip()):
    #                l.append(member)
    #    return l

    def find_by_scoutid(self, scoutid, section_wanted=None):
        """Return a list of records with matching scoutids"""
        l = []
        sections = self.all_yp_members_dict()
        for section in sections.keys():
            if (section_wanted and section_wanted != section):
                continue
            for member in sections[section]:
                try:
                    if (str(member['member_id']) == scoutid.strip()):
                        l.append(member)
                except:
                    print(str(member))
                    raise
        return l

    def find_by_scoutid_without_senior_duplicates(self, scoutid, section_wanted=None):
        """Return a list of records with matching scoutids excluding senior duplicates"""
        l = []
        sections = self.all_yp_members_without_senior_duplicates_dict()
        for section in sections.keys():
            if (section_wanted and section_wanted != section):
                continue
            for member in sections[section]:
                try:
                    if (str(member['member_id']) == scoutid.strip()):
                        l.append(member)
                except:
                    print(str(member))
                    raise
        return l

    # For each section we need to look at whether a member appears in a
    # senior section too (they will if they are in the process of
    # moving). If they are in a senior section we want to favour the
    # senior records (but warn if it is different).
    def remove_senior_duplicates(self, section, senior_members):
        kept_members = []
        for member in self.section_yp_members_without_leaders(section):
            if member[OSM_REF_FIELD] == "":
                # If no Personal Reference is set, do not look for
                # duplicates
                kept_members.append(member)
                continue

            matching_senior_members = [senior_member for senior_member
                                       in senior_members
                                       if senior_member[OSM_REF_FIELD] ==
                                       member[OSM_REF_FIELD]]
            if matching_senior_members:
                log.info("{} section: {} is in senior section - "
                         "favouring senior record".format(
                             section, member[OSM_REF_FIELD]))
                # check whether all of the fields are the same.
                for senior_member in matching_senior_members:
                    for field in self._important_fields:
                        if field == 'joined':
                            # We expect the joined field to be different.
                            continue
                        try:
                            if member[field] != senior_member[field]:
                                log.warn('{} section: {} senior record'
                                         'field mismatch ({}) "{}" != "{}"'
                                         '\n {}\n\n'.format(
                                             section, member[OSM_REF_FIELD],
                                             field, member[field],
                                             senior_member[field],
                                             str(member)))
                        except:
                            log.warn("Failed to compare field {}".format(
                                field
                            ), exc_info=True)
            else:
                kept_members.append(member)
        return kept_members

    def girls_in_section(self, section):
        return [m for m in self.all_yp_members_without_senior_duplicates_dict()[section]
                if (m['floating.gender'].lower() == 'f' or
                    m['floating.gender'].lower() == 'female')]

    def boys_in_section(self, section):
        return [m for m in self.all_yp_members_without_senior_duplicates_dict()[section]
                if (m['floating.gender'].lower() == 'm' or
                    m['floating.gender'].lower() == 'male')]

    def others_in_section(self, section):
        return [m for m in self.all_yp_members_without_senior_duplicates_dict()[section]
                if (not ((m['floating.gender'].lower() == 'm' or m['floating.gender'].lower() == 'male') or
                    (m['floating.gender'].lower() == 'f' or m['floating.gender'].lower() == 'female')))]

    def census(self):
        """Return the information required for the annual census."""

        r = {'Beavers': {'M': {5: 0, 6: 0, 7: 0, 8: 0},
                         'F': {5: 0, 6: 0, 7: 0, 8: 0},
                         'O': {5: 0, 6: 0, 7: 0, 8: 0}
                         },
             'Swinfen': {'M': {5: 0, 6: 0, 7: 0, 8: 0},
                         'F': {5: 0, 6: 0, 7: 0, 8: 0},
                         'O': {5: 0, 6: 0, 7: 0, 8: 0}
                         },
             'Paget': {'M': {5: 0, 6: 0, 7: 0, 8: 0},
                         'F': {5: 0, 6: 0, 7: 0, 8: 0},
                         'O': {5: 0, 6: 0, 7: 0, 8: 0}
                         },
             'Garrick': {'M': {5: 0, 6: 0, 7: 0, 8: 0},
                         'F': {5: 0, 6: 0, 7: 0, 8: 0},
                         'O': {5: 0, 6: 0, 7: 0, 8: 0}
                         },
             'Cubs': {'M': {7: 0, 8: 0, 9: 0, 10: 0},
                      'F': {7: 0, 8: 0, 9: 0, 10: 0},
                      'O': {7: 0, 8: 0, 9: 0, 10: 0}},
             'Maclean': {'M': {7: 0, 8: 0, 9: 0, 10: 0},
                      'F': {7: 0, 8: 0, 9: 0, 10: 0},
                      'O': {7: 0, 8: 0, 9: 0, 10: 0}},
             'Rowallan': {'M': {7: 0, 8: 0, 9: 0, 10: 0},
                      'F': {7: 0, 8: 0, 9: 0, 10: 0},
                      'O': {7: 0, 8: 0, 9: 0, 10: 0}},
             'Somers': {'M': {7: 0, 8: 0, 9: 0, 10: 0},
                      'F': {7: 0, 8: 0, 9: 0, 10: 0},
                      'O': {7: 0, 8: 0, 9: 0, 10: 0}},
             'Scouts': {'M': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0},
                        'F': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0},
                        'O': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0}},
             'Boswell': {'M': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0},
                        'F': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0},
                        'O': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0}},
             'Johnson': {'M': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0},
                        'F': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0},
                        'O': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0}},
             'Erasmus': {'M': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0},
                        'F': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0},
                        'O': {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0}}}

        for i in range(5, 9):
            r['Beavers']['F'][i] = len(
                [m for m in
                 self.girls_in_section('Swinfen') +
                 self.girls_in_section('Paget') +
                 self.girls_in_section('Garrick')
                 if int(m.age().days / 365) == i])
            r['Beavers']['M'][i] = len(
                [m for m in
                 self.boys_in_section('Swinfen') +
                 self.boys_in_section('Paget') +
                 self.boys_in_section('Garrick')
                 if int(m.age().days / 365) == i])
            r['Beavers']['O'][i] = len(
                [m for m in
                 self.others_in_section('Swinfen') +
                 self.others_in_section('Paget') +
                 self.others_in_section('Garrick')
                 if int(m.age().days / 365) == i])
            for section in ['Swinfen', 'Paget', 'Garrick']:
                r[section]['F'][i] = len(
                    [m for m in
                     self.girls_in_section(section)
                     if int(m.age().days / 365) == i])
                r[section]['M'][i] = len(
                    [m for m in
                     self.boys_in_section(section)
                     if int(m.age().days / 365) == i])
                r[section]['O'][i] = len(
                    [m for m in
                     self.others_in_section(section)
                     if int(m.age().days / 365) == i])

        for i in range(7, 11):
            r['Cubs']['F'][i] = len(
                [m for m in
                 self.girls_in_section('Maclean') +
                 self.girls_in_section('Rowallan') +
                 self.girls_in_section('Somers')
                 if int(m.age().days / 365) == i])
            r['Cubs']['M'][i] = len(
                [m for m in
                 self.boys_in_section('Maclean') +
                 self.boys_in_section('Rowallan') +
                 self.boys_in_section('Somers')
                 if int(m.age().days / 365) == i])
            r['Cubs']['O'][i] = len(
                [m for m in
                 self.others_in_section('Maclean') +
                 self.others_in_section('Rowallan') +
                 self.others_in_section('Somers')
                 if int(m.age().days / 365) == i])
            for section in ['Maclean', 'Rowallan', 'Somers']:
                r[section]['F'][i] = len(
                    [m for m in
                     self.girls_in_section(section)
                     if int(m.age().days / 365) == i])
                r[section]['M'][i] = len(
                    [m for m in
                     self.boys_in_section(section)
                     if int(m.age().days / 365) == i])
                r[section]['O'][i] = len(
                    [m for m in
                     self.others_in_section(section)
                     if int(m.age().days / 365) == i])

        for i in range(10, 16):
            r['Scouts']['F'][i] = len(
                [m for m in
                 self.girls_in_section('Johnson') +
                 self.girls_in_section('Boswell') +
                 self.girls_in_section('Erasmus')
                 if int(m.age().days / 365) == i])
            r['Scouts']['M'][i] = len(
                [m for m in
                 self.boys_in_section('Johnson') +
                 self.boys_in_section('Boswell') +
                 self.boys_in_section('Erasmus')
                 if int(m.age().days / 365) == i])
            r['Scouts']['O'][i] = len(
                [m for m in
                 self.others_in_section('Johnson') +
                 self.others_in_section('Boswell') +
                 self.others_in_section('Erasmus')
                 if int(m.age().days / 365) == i])
            for section in ['Boswell', 'Johnson', 'Erasmus']:
                r[section]['F'][i] = len(
                    [m for m in
                     self.girls_in_section(section)
                     if int(m.age().days / 365) == i])
                r[section]['M'][i] = len(
                    [m for m in
                     self.boys_in_section(section)
                     if int(m.age().days / 365) == i])
                r[section]['O'][i] = len(
                    [m for m in
                     self.others_in_section(section)
                     if int(m.age().days / 365) == i])
        return r
