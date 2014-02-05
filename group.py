from datetime import datetime
import logging
import osm

log = logging.getLogger(__name__)

OSM_REF_FIELD = 'PersonalReference'


class Member(osm.Member):

    def age(self, ref_date=datetime.now()):
        dob = datetime.strptime(
            self['dob'], '%d/%m/%Y')
        return ref_date - dob

osm.MemberClass = Member

class Group(object):
    
    SECTIONIDS = {'Adult': '18305',
                  'Paget': '9960',
                  'Brown': '17326',
                  'Maclean': '14324',
                  'Rowallan': '12700',
                  'Boswell': '10363',
                  'Johnson': '5882'}
                   #'Waiting List': ""}

    ADULT_SECTION = 'Adult'

    YP_SECTIONS = ['Paget',
                   'Brown',
                   'Maclean',
                   'Rowallan',
                   'Boswell',
                   'Johnson']

    def __init__(self, osm, auth, important_fields):
        self._osm = osm
        self._important_fields = important_fields
        self._sections = self._osm.OSM(auth, self.SECTIONIDS.values())

    def section_all_members(self, section):
        return self._sections.sections[self.SECTIONIDS[section]].members.values()

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

    def all_yp_members_without_senior_duplicates_dict(self):
        return {'Paget': self.remove_senior_duplicates('Paget',
                                                       self.all_cubs()),
                'Brown': self.remove_senior_duplicates('Brown',
                                                       self.all_cubs()),
                'Maclean': self.remove_senior_duplicates('Maclean',
                                                         self.all_scouts()),
                'Rowallan': self.remove_senior_duplicates('Rowallan',
                                                          self.all_scouts()),
                'Boswell': self.section_yp_members_without_leaders('Boswell'),
                'Johnson': self.section_yp_members_without_leaders('Johnson')}

        return {s: self.section_all_members(s) for
                s in self.YP_SECTIONS}
    
    def section_missing_references(self, section_name):
        return self._section_missing_references(section_name)

    def section_yp_members_without_leaders(self, section):
        return [member for member in
                self.section_all_members(section)
                if not member['patrol'].lower() in
                ['leaders', 'young leaders']]

    def section_leaders_in_yp_section(self, section):
        return [member for member in
                self.section_all_members(
                    section)
                if member['patrol'].lower() in
                ['leaders', 'young leaders']]

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

    def all_beavers(self):
        return self.section_yp_members_without_leaders('Paget') +\
            self.section_yp_members_without_leaders('Brown')

    def all_cubs(self):
        return self.section_yp_members_without_leaders('Rowallan') +\
            self.section_yp_members_without_leaders('Maclean')

    def all_scouts(self):
        return self.section_yp_members_without_leaders('Johnson') +\
            self.section_yp_members_without_leaders('Boswell')
        
    def find_ref_in_sections(self, reference):
        """Search for a reference in all of the sections.
        
        return a list of section names."""

        matching_sections = []

        for section in self.SECTIONIDS.keys():
            if reference in self.all_section_references(
                    section):
                matching_sections.append(section)

        return matching_sections

    # For each section we need to look at whether a member appears in a
    # senior section too (they will if they are in the process of
    # moving). If they are in a senior section we want to favour the
    # senior records (but warn if it is different).
    def remove_senior_duplicates(self, section, senior_members):
        kept_members = []
        for member in self.section_yp_members_without_leaders(section):
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
                        if member[field] != senior_member[field]:
                            log.warn('{} section: {} senior record'
                                     'field mismatch ({}) "{}" != "{}"'
                                     '\n {}\n\n'.format(
                                         section, member[OSM_REF_FIELD],
                                         field, member[field],
                                         senior_member[field],
                                         str(member)))
            else:
                kept_members.append(member)
        return kept_members

    def girls_in_section(self, section):
        return [m for m in self.section_yp_members_without_leaders(section)
                if m['Sex'].lower() == 'f' or m['Sex'].lower() == 'female']

    def boys_in_section(self, section):
        return [m for m in self.section_yp_members_without_leaders(section)
                if m['Sex'].lower() == 'm' or m['Sex'].lower() == 'male']

    def census(self):
        """Return the information required for the anual census."""

        r = {'Beavers': {'M':
                         {5: 0, 6: 0, 7: 0, 8: 0},
                         'F':
                         {5: 0, 6: 0, 7: 0, 8: 0}},
             'Cubs': {'M':
                      {7: 0, 8: 0, 9: 0, 10: 0},
                      'F':
                      {7: 0, 8: 0, 9: 0, 10: 0}},
             'Scouts': {'M':
                        {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0},
                        'F':
                        {10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0}}}

        for i in range(5, 9):
            r['Beavers']['F'][i] = len(
                [m for m in
                 self.girls_in_section('Brown') +
                 self.girls_in_section('Paget')
                 if int(m.age().days / 365) == i])
            r['Beavers']['M'][i] = len(
                [m for m in
                 self.boys_in_section('Brown') +
                 self.boys_in_section('Paget')
                 if int(m.age().days / 365) == i])

        for i in range(7, 11):
            r['Cubs']['F'][i] = len(
                [m for m in
                 self.girls_in_section('Maclean') +
                 self.girls_in_section('Rowallan')
                 if int(m.age().days / 365) == i])
            r['Cubs']['M'][i] = len(
                [m for m in
                 self.boys_in_section('Maclean') +
                 self.boys_in_section('Rowallan')
                 if int(m.age().days / 365) == i])

        for i in range(10, 16):
            r['Scouts']['F'][i] = len(
                [m for m in
                 self.girls_in_section('Johnson') +
                 self.girls_in_section('Boswell')
                 if int(m.age().days / 365) == i])
            r['Scouts']['M'][i] = len(
                [m for m in
                 self.boys_in_section('Johnson') +
                 self.boys_in_section('Boswell')
                 if int(m.age().days / 365) == i])

        return r
