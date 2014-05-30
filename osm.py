# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  osm.py <apiid> <token>
  osm.py <apiid> <token> run <query>
  osm.py <apiid> <token> -a <email> <password>
  osm.py (-h | --help)
  osm.py --version


Options:
  -h --help      Show this screen.
  --version      Show version.
  -a             Request authorisation credentials.

"""

from docopt import docopt

import sys
import urllib.request, urllib.parse, urllib.error
import urllib.request, urllib.error, urllib.parse
import json
import pickle
import logging
import time
import datetime
import dateutil.tz
import dateutil.parser
import pytz
import pprint
import io
import collections
from operator import attrgetter

log = logging.getLogger(__name__)
pp = pprint.PrettyPrinter(indent=4)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"

TZ = dateutil.tz.gettz("Europe/London")
pyTZ = pytz.timezone('Europe/London')
FMT = '%Y-%m-%d %H:%M:%S %Z%z'


class OSMException(Exception):
    def __init__(self, url, values, error):
        self._url = url
        self._values = values
        self._error = error

    def __str__(self):
        return "OSM API Error from {0}:\n" \
               "values = {1}\n" \
               "result = {2}".format(self._url,
                                     self._values,
                                     self._error)


class OSMObject(collections.MutableMapping):
    def __init__(self, osm, accessor, record):
        self._osm = osm
        self._accessor = accessor
        self._record = record

    # def __getattr__(self, key):
    #     try:
    #         return self._record[key]
    #     except:
    #         raise AttributeError("%r object has no attribute %r" %
    #                              (type(self).__name__, key))

    def __getitem__(self, key):
        try:
            return self._record[key]
        except:
            raise KeyError("%r object has no attribute %r" %
                           (type(self).__name__, key))

    def __setitem__(self, key, value):
        try:
            self._record[key] = value
        except:
            raise KeyError("%r object has no attribute %r" %
                           (type(self).__name__, key))

    def __delitem__(self, key):
        try:
            del (self._record[key])
        except:
            raise KeyError

    def __len__(self):
        return len(self._record)

    def __iter__(self):
        return self._record.__iter__()

    def __str__(self):
        return pprint.pformat(self._record)


class Accessor(object):
    __cache__ = {}
    _cache_hits = 0
    _cache_misses = 0

    BASE_URL = "https://www.onlinescoutmanager.co.uk/"

    def __init__(self, authorisor):
        self._auth = authorisor

    @classmethod
    def clear_cache(cls):
        cls.__cache__ = {}

    @classmethod
    def __cache_save__(cls, cache_file):
      log.info("Saving cache: (hits = {}, misses = {})".format(cls._cache_hits, cls._cache_misses))
      pickle.dump(cls.__cache__, cache_file)

    @classmethod
    def __cache_load__(cls, cache_file):
        cls.__cache__ = pickle.load(cache_file)

    @classmethod
    def __cache_lookup__(cls, url, data):
        k = url + repr(data)
        if k in cls.__cache__:
            log.debug('Cache hit')
            cls._cache_hits += 1
            #log.debug("Cache hit: ({0}) = {1}\n".format(k,
            #                                            cls.__cache__[k]))
            return cls.__cache__[k]

        cls._cache_misses += 1
        #log.debug("Cache miss: ({0})\n"\
        #          "Keys: {1}\n".format(k,
        #                               cls.__cache__.keys()))

        return None

    @classmethod
    def __cache_set__(cls, url, data, value):
        cls.__cache__[url + repr(data)] = value

    def __call__(self, query, fields=None, authorising=False, clear_cache=False, debug=False):

        if clear_cache:
            self.clear_cache()

        url = self.BASE_URL + query

        values = {'apiid': self._auth.apiid,
                  'token': self._auth.token}

        if not authorising:
            values.update({'userid': self._auth.userid,
                           'secret': self._auth.secret})

        if fields:
            values.update(fields)

        data = urllib.parse.urlencode(values)

        req = urllib.request.Request(url, data.encode('utf-8'))

        obj = self.__class__.__cache_lookup__(url, data)

        if not obj:

          log.debug("urlopen: {0}, {1}".format(url, data.encode('utf-8')))
        
          try:
            result = urllib.request.urlopen(req).readall().decode('utf-8')
          except:
            log.error("urlopen failed: {0}, {1}".format(url, data.encode('utf-8')))
            raise


          # Crude test to see if the response is JSON
          # OSM returns a string as an error case.
          try:
              if result[0] not in ('[', '{'):
                  log.warn("Result not JSON because: {0} not in  ('[', '{{')".format(result[0]))
                  raise OSMException(url, values, result)
          except IndexError:
              # This means that result is not a list
              log.warn("Result not a list: {0} {1}".format(url, values))
              log.error(repr(result))
              raise

          obj = json.loads(result)

          if 'error' in obj:
              log.warn("Error in JSON obj: {0} {1}".format(url, values))
              raise OSMException(url, values, obj['error'])
          if 'err' in obj:
              log.warn("Err in JSON obj: {0} {1}".format(url, values))
              raise OSMException(url, values, obj['err'])

          self.__class__.__cache_set__(url, data, obj)

        if debug:
            log.debug(pp.pformat(obj))
        return obj


class Authorisor(object):
    def __init__(self, apiid, token):
        self.apiid = apiid
        self.token = token

        self.userid = None
        self.secret = None

    def authorise(self, email, password):
        fields = {'email': email,
                  'password': password}

        accessor = Accessor(self)
        creds = accessor("users.php?action=authorise", fields, authorising=True)

        self.userid = creds['userid']
        self.secret = creds['secret']

    def save_to_file(self, dest):
        dest.write(self.userid + '\n')
        dest.write(self.secret + '\n')

    def load_from_file(self, src):
        self.userid = src.readline()[:-1]
        self.secret = src.readline()[:-1]


class Term(OSMObject):
    def __init__(self, osm, accessor, record):
        OSMObject.__init__(self, osm, accessor, record)

        self.startdate = datetime.datetime.strptime(record['startdate'],
                                                    '%Y-%m-%d')

        self.enddate = datetime.datetime.strptime(record['enddate'],
                                                  '%Y-%m-%d')

    def is_active(self):
        now = datetime.datetime.now()
        return (self.startdate < now) and (self.enddate > now)


class Badge(OSMObject):
    def __init__(self, osm, accessor, section, badge_type, details, structure):
        self._section = section
        self._badge_type = badge_type
        self.name = details['name']
        self.table = details['table']

        activities = {}
        if len(structure) > 1:
            for activity in [row['name'] for row in structure[1]['rows']]:
                activities[activity] = ''

        OSMObject.__init__(self, osm, accessor, activities)

    def get_members(self):
        url = "challenges.php?"\
            "&termid={0}" \
            "&type={1}" \
            "&sectionid={2}" \
            "&section={3}" \
            "&c={4}".format(self._section.term['termid'],
                            self._badge_type,
                            self._section['sectionid'],
                            self._section['section'],
                            self.name.lower())
        
        return [ OSMObject(self._osm,
                           self._accessor,
                           record) for record in \
                           self._accessor(url)['items'] ]

class Badges(OSMObject):
    def __init__(self, osm, accessor, record, section, badge_type):
        self._section = section
        self._badge_type = badge_type
        self._order = record['badgeOrder']
        self._details = record['details']
        self._stock = record['stock']
        self._structure = record['structure']

        badges = {}
        if self._details:
          for badge in list(self._details.keys()):
              badges[badge] = Badge(osm, accessor,
                                    self._section,
                                    self._badge_type,
                                    self._details[badge],
                                    self._structure[badge])

        OSMObject.__init__(self, osm, accessor, badges)


class Member(OSMObject):

    def __init__(self, osm, section, accessor, column_map, record):
        OSMObject.__init__(self, osm, accessor, record)

        self._section = section
        self._column_map = column_map
        for k, v in list(self._column_map.items()):
            self._column_map[k] = v.replace(' ', '')

        self._reverse_column_map = dict((reversed(list(i)) for
                                         i in list(column_map.items())))
        self._changed_keys = []

    def __str__(self):
        return "Member: \n Section = {!r}\n column_map = {!r} \n" \
            "reverse_column_map = {!r} \n record = {}".format(
                self._section, self._column_map,
                self._reverse_column_map, OSMObject.__str__(self))

    def __getattr__(self, key):
        try:
            return self._record[key]
        except:

            try:
                return self._record[self._reverse_column_map[key]]
            except:
                raise KeyError("{!r} object has no attribute {!r}: "
                               "primary keys {!r}, "
                               "reverse keys: {!r}".format(
                                   type(self).__name__, key,
                                   self._record.keys(),
                                   self._reverse_column_map.keys()))

    def __getitem__(self, key):
        try:
            return self._record[key]
        except:
            try:
                return self._record[self._reverse_column_map[key]]
            except:
                raise KeyError("{!r} object has no attribute {!r}:"
                               "primary keys {!r}, reverse keys: {!r}".format(
                                   type(self).__name__, key,
                                   self._record.keys(),
                                   self._reverse_column_map.keys()))
         
    def __setitem__(self, key, value):
        try:
            self._record[key] = value
            if key not in self._changed_keys:
                self._changed_keys.append(key)
        except:
            try:
                self._record[self._reverse_column_map[key]] = value
                if self._reverse_column_map[key] not in self._changed_keys:
                    self._changed_keys.append(self._reverse_column_map[key])
 
            except:
                raise KeyError("{!r} object has no attribute {!r}:"
                               "primary keys {!r}, reverse keys: {!r}".format(
                                   type(self).__name__, key,
                                   self._record.keys(),
                                   self._reverse_column_map.keys()))

            raise KeyError("{!r} object has no attribute {!r}: "
                           "avaliable keys: {!r}".format(
                               (type(self).__name__, key, 
                                self._record.keys())))

    # def remove(self, last_date):
    #     """Remove the member record."""
    #     delete_url='users.php?action=deleteMember&type=leaveremove&section={0}'
    #     delete_url = delete_url.format(self._section.section)
    #     fields={ 'scouts': ["{0}".format(self.scoutid),],
    #              'sectionid': self._section.sectionid,
    #              'date': last_date }
        
    #     self._accessor(delete_url, fields, clear_cache=True, debug=True)

    def save(self):
        """Write the member to the section."""
        update_url='users.php?action=updateMember&dateFormat=generic'
        patrol_url='users.php?action=updateMemberPatrol'
        create_url='users.php?action=newMember'

        if self['scoutid'] == '':
            # create
            fields = {}
            for key in self._changed_keys:
                fields[key] = self._record[key]
            fields['sectionid'] = self._section['sectionid']
            record = self._accessor(create_url, fields, clear_cache=True, debug=True)
            self['scoutid'] = record['scoutid']
        else:
            # update
            fields = {}
            for key in self._changed_keys:
                fields[key] = self._record[key]

            result = True
            for key in fields:
                record = self._accessor(update_url, 
                                        { 'scoutid': self['scoutid'],
                                          'column': self._reverse_column_map[key],
                                          'value': fields[key],
                                          'sectionid': self._section['sectionid'] }, 
                                        clear_cache=True, debug=True)
                if record[self._reverse_column_map[key]] != fields[key]:
                    result = False

            # TODO handle change to grouping.

            return result

    def get_badges(self):
        "Return a list of badges objects for this member."

        ret = []
        for i in list(self._section.challenge.values()):
            ret.extend( [ badge for badge in badge.get_members() \
                          if badge['scoutid'] == self['scoutid'] ] )
        return ret
        
        
class Members(OSMObject):
    DEFAULT_DICT = {  'address': '',
                      'address2': '',
                      'age': '',
                      'custom1': '',
                      'custom2': '',
                      'custom3': '',
                      'custom4': '',
                      'custom5': '',
                      'custom6': '',
                      'custom7': '',
                      'custom8': '',
                      'custom9': '',
                      'dob': '',
                      'email1': '',
                      'email2': '',
                      'email3': '',
                      'email4': '',
                      'ethnicity': '',
                      'firstname': '',
                      'joined': '',
                      'joining_in_yrs': '',
                      'lastname': '',
                      'medical': '',
                      'notes': '',
                      'parents': '',
                      'patrol': '',
                      'patrolid': '',
                      'patrolleader': '',
                      'phone1': '',
                      'phone2': '',
                      'phone3': '',
                      'phone4': '',
                      'religion': '',
                      'school': '',
                      'scoutid': '',
                      'started': '',
                      'subs': '',
                      'type': '',
                      'yrs': 0}

    def __init__(self, osm, section, accessor, column_map, record):
        self._osm = osm,
        self._section = section
        self._accessor = accessor,
        self._column_map = column_map
        self._identifier = record['identifier']

        members = {}
        for member in record['items']:
            members[member[self._identifier]] = MemberClass(osm, section, accessor, column_map, member)

        OSMObject.__init__(self, osm, accessor, members)

    def new_member(self, firstname, lastname, dob, startedsection, started):
        new_member = MemberClass(self._osm, self._section,
                                 self._accessor,self._column_map,self.DEFAULT_DICT)
        new_member['firstname'] = firstname
        new_member['lastname'] = lastname
        new_member['dob'] = dob
        new_member['startedsection'] = startedsection
        new_member['started'] = started
        new_member['patrolid'] = '-1'
        new_member['patrolleader'] = '0'
        new_member['phone1'] = ''
        new_member['email1'] = ''
        return new_member


class Event(OSMObject):

    def __init__(self, osm, section, accessor, record):
        OSMObject.__init__(self, osm, accessor, record)

        self.meeting_date = datetime.datetime.strptime(
            self._record['meetingdate'], '%Y-%m-%d')

        h, m, s = (int(i) for i in self._record['starttime'].split(':'))
        self.start_time = datetime.datetime.combine(
            self.meeting_date,
            datetime.time(h, m, s))
        self.start_time = dateutil.parser.parse(
            pyTZ.localize(self.start_time).strftime(FMT))

        h, m, s = (int(i) for i in self._record['endtime'].split(':'))
        self.end_time = datetime.datetime.combine(
            self.meeting_date,
            datetime.time(h, m, s))
        self.end_time = dateutil.parser.parse(
            pyTZ.localize(self.end_time).strftime(FMT))

    def __str__(self):
        return "{} - {} - {} - {} - {}".format(
            self['title'], self['notesforparents'],
            self['meetingdate'],
            self['starttime'], self['endtime'])


class Programme(OSMObject):
    def __init__(self, osm, section, accessor, record):
        self._osm = osm,
        self._section = section
        self._accessor = accessor

        events = {}
        for event in record['items']:
            events["{} - {}".format(
                event['title'],
                event['meetingdate'])] = Event(osm, section, accessor, event)

        OSMObject.__init__(self, osm, accessor, events)

    def events_by_date(self):
        """Return a list of events sorted by start date (and time)"""

        return sorted(self._record.values(),
                      key=attrgetter('meeting_date'))
    

class Section(OSMObject):
    def __init__(self, osm, accessor, record, init=True):
        OSMObject.__init__(self, osm, accessor, record)

        try:
            self._member_column_map = record['sectionConfig']['columnNames']
        except KeyError:
            log.debug("No extra member columns.")
            self._member_column_map = {}

        if init:
          self.init()

    def init(self):
        self.terms = [term for term in self._osm.terms(self['sectionid'])
                      if term.is_active()]

        if len(self.terms) > 1:
            log.warn("{!r}: More than 1 term is active, picking "
                     "last in list {!r}".format(
                         self['sectionname'],
                         [(term['name'], term['past']) for
                          term in self.terms]))
            sys.exit(0)

        if len(self.terms) == 0:
            # If there is no active term it does make sense to gather
            # badge info.
            self.term = None
            self.challenge = None
            self.activity = None
            self.staged = None
            self.core = None
        else:
            self.term = self.terms[-1]
            self.challenge = self._get_badges('challenge')
            self.activity = self._get_badges('activity')
            self.staged = self._get_badges('staged')
            self.core = self._get_badges('core')
        
        try:
            self.members = self._get_members()
        except:
            log.warn("Failed to get members for section {0}"
                     .format(self['sectionname']))
            self.members = []

        try:
            self.programme = self._get_programme()
        except urllib.error.HTTPError as err:
            log.warn("Failed to get programme for section {0}: {1}"
                     .format(self['sectionname'],
                             err))
            self.programme = []
        except:
            log.warn("Failed to get programme for section {0}"
                     .format(self['sectionname']),
                     exc_info=True)
            self.programme = []
          
    def __repr__(self):
        return 'Section({0}, "{1}", "{2}")'.format(
            self['sectionid'],
            self['sectionname'],
            self['section'])

    def _get_badges(self, badge_type):
        url = "challenges.php?action=getInitialBadges" \
              "&type={0}" \
              "&sectionid={1}" \
              "&section={2}" \
              "&termid={3}" \
            .format(badge_type, 
                    self['sectionid'],
                    self['section'],
                    self.term['termid'])

        return Badges(self._osm, self._accessor,
                      self._accessor(url), self, badge_type)

    def events(self):
        pass

    def _get_members(self):
        url = "users.php?&action=getUserDetails" \
              "&sectionid={0}" \
              "&termid={1}" \
              "&dateFormat=uk" \
              "&section={2}" \
            .format(self['sectionid'],
                    self.term['termid'],
                    self['section'])

        return Members(self._osm, self, self._accessor,
                       self._member_column_map, self._accessor(url))

    def _get_programme(self):
        url = "programme.php?action=getProgrammeSummary"\
              "&sectionid={0}&termid={1}".format(self['sectionid'],
                                                 self.term['termid'])

        return Programme(self._osm, self, self._accessor,
                         self._accessor(url))


class OSM(object):
    def __init__(self, authorisor, sectionid_list=False):
        self._accessor = Accessor(authorisor)

        self.sections = {}
        self.section = None

        self.init(sectionid_list)

    def init(self, sectionid_list=False):
        roles = self._accessor('api.php?action=getUserRoles')

        self.sections = {}

        for section in [Section(self, self._accessor, role, init=False)
                        for role in roles
                        if 'section' in role]:
            if sectionid_list is False or \
               section['sectionid'] in sectionid_list:
                section.init()
                self.sections[section['sectionid']] = section
            
                if section['isDefault'] == '1':
                    self.section = section
                    log.info("Default section = {0}, term = {1}".format(
                        self.section['sectionname'],
                        self.section.term['name']))

        if self.section is None:
            self.section = self.sections[-1]

    def terms(self, sectionid):
        terms = self._accessor('api.php?action=getTerms')
        if sectionid in terms:
            return [Term(self, self._accessor, term) for term
                    in terms[sectionid]]
        return []

MemberClass = Member

if __name__ == '__main__':

    logging.basicConfig(level=logging.DEBUG)
    log.debug("Debug On\n")

    try:
        Accessor.__cache_load__(open(DEF_CACHE, 'r'))
    except:
        log.debug("Failed to load cache file\n")

    args = docopt(__doc__, version='OSM 2.0')
    print (args)
    if args['-a']:
        auth = Authorisor(args['<apiid>'], args['<token>'])
        auth.authorise(args['<email>'],
                       args['<password>'])
        auth.save_to_file(open(DEF_CREDS, 'w'))
        sys.exit(0)

    auth = Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    if args['run']:
        accessor = Accessor(auth)

        pp.pprint(accessor(args['<query>']))


    osm = OSM(auth)

    log.debug('Sections - {0}\n'.format(osm.sections))


    test_section = '15797'
    for badge in list(osm.sections[test_section].challenge.values()):
        log.debug('{0}'.format(badge._record))
              
    #members = osm.sections[test_section].members

    #    member = members[members.keys()[0]]
    #member['special'] = 'changed'
    #member.save()
    
    #new_member = members.new_member('New First 2','New Last 2','02/09/2004','02/12/2012','02/11/2012')
    
    #log.debug("New member = {0}: {1}".format(new_member.firstname,new_member.lastname))
    #new_member.save()

    
        
    #for k,v in osm.sections['14324'].members.items():
    #    log.debug("{0}: {1} {2} {3}".format(k,v.firstname,v.lastname,v.TermtoScouts))

    #for k,v in osm.sections['14324'].activity.items():
    #    log.debug("{0}: {1}".format(k,v.keys()))


    #pp.pprint(osm.sections['14324'].members())
    #for k,v in osm.sections['14324'].challenge.items():
    #    log.debug("{0}: {1}".format(k,v.keys()))


    Accessor.__cache_save__(open(DEF_CACHE, 'w'))
