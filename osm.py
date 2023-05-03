# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  osm.py [-c=<client_id>] [-p=<client_secret>] [-d] [-s=<sectionid>]
  osm.py [-c=<client_id>] [-p=<client_secret>] [-d] [-s=<sectionid>] run <query>
  osm.py (-h | --help)
  osm.py --version


Options:
  -h --help      Show this screen.
  --version      Show version.
  -a             Request authorisation credentials.
  -d             Enable debug
  -s=<sectionid> Section ID to query [default: all].

"""
import enum
import time
import traceback

from docopt import docopt

import sys
import requests_cache
import requests_oauth2client as oauth2
import logging
import datetime
import dateutil.tz
import dateutil.parser
import pytz
import pprint

if sys.version_info.major == 3 and sys.version_info.minor >= 10:
    from collections.abc import MutableMapping, Sequence
else:
    from collections import MutableMapping, Sequence

import os.path
from os.path import expanduser
from operator import attrgetter

log = logging.getLogger(__name__)
pp = pprint.PrettyPrinter(indent=4)

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


class OSMObject(MutableMapping):

    def __init__(self, osm, accessor, record):
        self._osm = osm
        self._accessor = accessor
        self._record = record

    def __getitem__(self, key):
        try:
            return self._record[key]
        except:
            raise KeyError("%r object has no attribute %r. %r" %
                           (type(self).__name__, key, self._record))

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
    BASE_URL = "https://www.onlinescoutmanager.co.uk/"

    def __init__(self, authorisor):
        self._auth = authorisor

        # self.session = oauth2.requests.Session()
        self._session = requests_cache.CachedSession(
            os.path.join(expanduser("~"), '.osm_request_cache'),
            allowable_methods=('GET', 'POST'),
            include_get_headers=True,
            expire_after=60 * 60
        )
        self._session.auth = self._auth.auth

    def __call__(self, query, fields=None, authorising=False,
                 clear_cache=False, debug=False, result_type='json'):
        #if clear_cache:
        #    self.clear_cache()

        url = self.BASE_URL + query
        values = {}
        # values = {
        #     'apiid': self._auth.apiid,
        #     'token': self._auth.token}
        #
        # if not authorising:
        #     values.update({
        #         'userid': self._auth.userid,
        #         'secret': self._auth.secret})

        if fields:
            values.update(fields)

        log.debug("posting: {} {}".format(url, values))

        retry_limit = 2
        while retry_limit > 0:
            try:
                result = self._session.post(url, data=values)
                retry_limit -= 1
            except:
                log.error("urlopen failed: {0}, {1}".format(
                    url, repr(values)))
                raise

            if result.status_code == 429:
                retry_after = int(result.headers['Retry-After'])
                log.warning(f"Exceeded OSM API limited, waiting {retry_after}s before retry ...")
                time.sleep(retry_after)
                log.warning("Retrying")
            else:
                break

        if result.status_code != 200:
            log.error("urlopen failed with status code {}: {}, {}".format(
                result.status_code, url, repr(values)))
            return None

        # Trap if we are exceeding our ratelimit

        ratelimit = int(result.headers['x-ratelimit-limit'])
        ratelimit_remaining = int(result.headers['x-ratelimit-remaining'])
        ratelimit_reset_period = int(result.headers['x-ratelimit-reset'])

        log.warning(f"Ratelimit remaining: (limit={ratelimit}, remaining={ratelimit_remaining})")

        if ratelimit_remaining < (ratelimit / 100) * 50:
            log.warning(f"Reached 50% of OSM ratelimit. "
                        f"(limit={ratelimit}, remaining={ratelimit_remaining}, "
                        f"reset_period={ratelimit_reset_period}")
        elif ratelimit_remaining < (ratelimit / 100) * 20:
            log.warning(f"Reached 80% of OSM ratelimit. "
                        f"(limit={ratelimit}, remaining={ratelimit_remaining}, "
                        f"reset_period={ratelimit_reset_period}")
        elif ratelimit_remaining < (ratelimit / 100) * 5:
            log.warning(f"Reached 95% of OSM ratelimit. "
                        f"(limit={ratelimit}, remaining={ratelimit_remaining}, "
                        f"reset_period={ratelimit_reset_period}")
            log.warning(f"Waiting for reset period of {ratelimit_reset_period}s...")
            time.sleep(ratelimit_reset_period + 30)
            log.warning(f"Finished waiting, carrying on.")

        # Crude test to see if the response is JSON
        # OSM returns a string as an error case.
        if result_type == 'json':
            try:
                obj = result.json()
            except:
                log.warning("Result not JSON because: {0} not in "
                         "('[', '{{')".format(result.text))
                raise OSMException(url, values, result)
        else:
            obj = result

        if debug:
            log.debug(pp.pformat(obj))

        log.debug(f"(limit={ratelimit}, remaining={ratelimit_remaining}, "
                  f"reset_period={ratelimit_reset_period}")
        return obj


class Authorisor(object):

    def __init__(self, creds_file=None, client_id=None, client_secret=None):
        if creds_file is not None:
            self.load_from_file(creds_file)
        else:
            self.client_id = client_id
            self.client_secret = client_secret
        self.auth = None
        self.authorise()

    def authorise(self):
        oauth2client = oauth2.OAuth2Client(
            token_endpoint="https://www.onlinescoutmanager.co.uk/oauth/token",
            auth=(self.client_id, self.client_secret),
            )

        self.auth = oauth2.OAuth2ClientCredentialsAuth(
            oauth2client,
            scope="section:member:read",
            resource="https://www.onlinescoutmanager.co.uk/oauth/resource"
        )

    def save_to_file(self, dest):
        dest.write(self.client_id + '\n')
        dest.write(self.client_secret + '\n')

    def load_from_file(self, src):
        self.client_id = src.readline()[:-1]
        self.client_secret = src.readline()[:-1]


class Term(OSMObject):

    def __init__(self, osm, accessor, record):
        OSMObject.__init__(self, osm, accessor, record)

        self.startdate = datetime.datetime.strptime(record['startdate'],
                                                    '%Y-%m-%d')

        self.enddate = datetime.datetime.strptime(record['enddate'],
                                                  '%Y-%m-%d')

    def is_active(self, date=None):
        now = datetime.datetime.now().date() if date is None else date
        return (self.startdate.date() <= now) and (self.enddate.date() >= now)

    def __repr__(self):
        return "{}: {} - {} {}".format(self['name'],
                                       self.startdate,
                                       self.enddate,
                                       datetime.datetime.now())


class Badge(OSMObject):

    def __init__(self, osm, accessor, section, badge_type, details, structure):
        self._section = section
        self._badge_type = badge_type
        self.name = details['name']
        # self.table = details['table']

        activities = {}
        if len(structure) > 1:
            for activity in [row['name'] for row in structure[1]['rows']]:
                activities[activity] = ''

        OSMObject.__init__(self, osm, accessor, activities)

    def get_members(self):
        url = "challenges.php?" \
              "&termid={0}" \
              "&type={1}" \
              "&sectionid={2}" \
              "&section={3}" \
              "&c={4}".format(self._section.term['termid'],
                              self._badge_type,
                              self._section['section_id'],
                              self._section['section'],
                              self.name.lower())

        return [OSMObject(self._osm,
                          self._accessor,
                          record) for record in
                self._accessor(url)['items']]


class Badges(OSMObject):

    def __init__(self, osm, accessor, record, section, badge_type):
        self._section = section
        self._badge_type = badge_type
        self._order = record['badgeOrder']
        self._details = record['details']
        # self._stock = record['stock']
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
    # List of custom groups that we expect to find.
    EXPECTED_CUSTOM = ('customisable_data', 'contact_primary_1',
                       'contact_primary_2', 'contact_primary_member')

    def __init__(self, osm, section, accessor, record, custom):
        OSMObject.__init__(self, osm, accessor, record)

        self._section = section
        self._custom = custom

    def check_custom_group(self):
        '''Report any errors missing custom groups.'''

        custom_groups = [group['identifier'] for group in
                         self._custom]

        missing = [group for group in self.__class__.EXPECTED_CUSTOM
                   if group not in custom_groups]

        if len(missing) > 0:
            log.warning("Member record ({}) missing custom groups: {}".fomat(
                " ".join(self['first_name'], self['last_name']),
                ",".join(missing)))

    def __str__(self):
        groups = [(group['identifier'], group['columns'])
                  for group in self._custom]

        out = ""
        for group in groups:
            group_id, columns = group

            for column in columns:
                out += "{}.{} ({}) = {}\n".format(
                    group_id, column['varname'],
                    column['label'], column['value'])

        return "Member: \n Section = {!r}\n record = {} \n custom = {}".format(
            self._section, OSMObject.__str__(self), out)

    def lookup(self, key):
        """Attempt to find a key in the member record."""

        head, tail = key.split('.')

        group_id, columns = [(group['group_id'], group['columns'])
                             for group in self._custom
                             if group['identifier'] == head][0]

        try:
            value = [column['value'] for column
                     in columns
                     if column['varname'] == tail][0]
        except:
            # print("\n".join([column['varname'] for column in columns]))
            value = [
                column['value'] for column
                in columns
                if column['label'].lower().replace(" ", "") == tail.lower()][0]

        return value

    def get_badges(self, section_type):
        """Return a list of awarded badges"""
        url = "ext/badges/badgesbyperson/?action=loadBadgesByMember&" \
              "section={}" \
              "&sectionid={}&term_id={}".format(
            section_type,
            int(self._section['section_id']),
            self._section.term['termid'])

        badge_data = self._accessor(url)

        my_badges = [member for member in badge_data['data']
                     if member['scout_id'] == self['member_id']]

        return my_badges[0]['badges'] if my_badges else None

    def __getattr__(self, key):
        try:
            return self.__dict__['_record'][key]
        except:
            try:
                return self.lookup(key)
            except:
                raise KeyError("{!r} object has no attribute {!r}\n"
                               "  Record was {}\n"
                               "".format(type(self).__name__, key,
                                         str(self)))

    def __getitem__(self, key):
        try:
            return self.__dict__['_record'][key]
        except:
            try:
                return self.lookup(key)
            except:
                raise KeyError("{!r} object has no attribute {!r}\n"
                               "  Record was {}:\n"
                               "".format(
                    type(self).__name__, key,
                    str(self)
                ))

    # def get_badges(self):
    #     "Return a list of badges objects for this member."
    #
    #     ret = []
    #     for i in list(self._section.challenge.values()):
    #         ret.extend([badge for badge in badge.get_members()
    #                     if badge['scoutid'] == self['scoutid']])
    #     return ret


class Users(Sequence):

    def __init__(self, osm, section, accessor, record):
        if record:
            self._users = record
        else:
            self._users = []

    def __getitem__(self, indx):
        return self._users[indx]

    def __len__(self):
        return len(self._users)

    def get_by_name(self, name):
        return [_ for _ in self._users if _['firstname'] == name][0]


class Members(OSMObject):

    def __init__(self, osm, section, accessor, record):
        self._osm = osm,
        self._section = section
        self._accessor = accessor
        self._column_map = record['meta']['structure']

        members = {}
        for key, member in record['data'].items():
            # This is the old method of getting the full member data.
            # # Fetch detailed info for member.
            # url = "ext/customdata/?action=getData&section_id={}".format(
            #     int(self._section['section_id']))
            # # "&section_id={}".format(self._section['section_id'])
            # fields = {
            #     'associated_id': key,
            #     'associated_type': 'member',
            #     'context': 'members'}
            #
            # custom_data = self._accessor(url, fields=fields)

            # This is just a dummy as we are not doing it this way any more.
            custom_data = {'data': {}}

            # We need to extra all of the custom data in sensibly names keys on the member.
            # The keynames are constructed from the column map dictionary. This is a two
            # level dict. We construct a key as identifier.varname and then extract its value
            # from the main member dict using the group_id and column_id.
            for column in self._column_map:
                for field in column['columns']:
                    value = member['custom_data'][str(column['group_id'])][str(field['column_id'])]
                    new_key = "{}.{}".format(column['identifier'], field['varname'])
                    member[new_key] = value

            members[key] = MemberClass(
                osm, section, accessor, member, custom_data['data'])

        OSMObject.__init__(self, osm, accessor, members)

    def get_by_event_attendee(self, attendee):
        return self[attendee['scoutid']]

    def is_member(self, scoutid):
        """Return true is the scoutid is in this Members list."""
        try:
            if self[scoutid]:
                return True
            else:
                return False
        except KeyError:
            return False


class Movers(Sequence):

    def __init__(self, osm, section, accessor, headers, data):
        self._osm = osm,
        self._section = section
        self._accessor = accessor

        self._movers = []

        fields = {row['field']: row['name'] for row in headers['structure'][1]['rows']}

        self._movers = data['items']

        # Merge the field names to make look up easier.
        for mover in self._movers:
            for field, name in fields.items():
                mover[name] = mover[field]
                # Remove the 'F_1' style name as it is just confusing.
                del mover[field]

        # Build a list of all available headings.
        self.headers = list(self._movers[0].keys())

    def __getitem__(self, indx):
        return self._movers[indx]

    def __len__(self):
        return len(self._movers)


class Event(OSMObject):

    def __init__(self, osm, section, accessor, record):
        self._osm = osm,
        self._section = section
        self._accessor = accessor
        OSMObject.__init__(self, osm, accessor, record)

        self._attendees = None
        self._fieldmap = None

        try:
            self.start_date = datetime.datetime.strptime(
                self._record['startdate'], '%d/%m/%Y')

            if self._record['enddate'] not in ('//', '00/00/0000'):
                self.end_date = datetime.datetime.strptime(
                    self._record['enddate'], '%d/%m/%Y')
            else:
                self.end_date = self.start_date

            if self._record['starttime']:
                h, m, s = (int(i) for i in self._record['starttime'].split(':'))
                self.start_time = datetime.datetime.combine(
                    self.start_date,
                    datetime.time(h, m, s))
                self.start_time = dateutil.parser.parse(
                    pyTZ.localize(self.start_time).strftime(FMT))
            else:
                self.start_time = self.start_date

            if self._record['endtime']:
                h, m, s = (int(i) for i in self._record['endtime'].split(':'))
                self.end_time = datetime.datetime.combine(
                    self.end_date,
                    datetime.time(h, m, s))
                self.end_time = dateutil.parser.parse(
                    pyTZ.localize(self.end_time).strftime(FMT))
            else:
                self.end_time = self.end_date

        except:
            log.warning(f"Failed to setup end and start date for: {self._record}", exc_info=True)
            raise

    @property
    def fieldmap(self):
        if not self._fieldmap:
            raw = self._get_fieldmap(self._osm,
                                     self._section,
                                     self._accessor)

            rows = [x['rows'] for x in raw]
            flat = [i for j in rows for i in j]
            self._fieldmap = [(d['name'], d['field']) for d in flat]
            self._fieldmap = [_ for _ in self._fieldmap if _[0].strip() != ""]
        return self._fieldmap

    @property
    def attendees(self):
        if not self._attendees:
            self._attendees = self._get_attendees(self._osm,
                                                  self._section,
                                                  self._accessor)
        return self._attendees

    def _get_fieldmap(self, osm, section, accessor):
        url = "ext/events/event/?action=getStructureForEvent" \
              "&sectionid={0}" \
              "&termid={1}" \
              "&eventid={2}" \
            .format(section['section_id'],
                    section.term['termid'],
                    self['eventid'])

        return accessor(url)['structure']

    def _get_attendees(self, osm, section, accessor):
        url = "ext/events/event/?action=getAttendance" \
              "&sectionid={0}" \
              "&termid={1}" \
              "&eventid={2}" \
            .format(section['section_id'],
                    section.term['termid'],
                    self['eventid'])

        return accessor(url)['items']

    def __str__(self):
        return "{} - {} - {}".format(
            self['name'],
            self['date'],
            self['location']
        )


class Events(Sequence):

    def __init__(self, osm, section, accessor, record):
        if record:
            self._events = [Event(osm, section, accessor, _)
                            for _ in record['items']]
        else:
            self._events = []

    def __getitem__(self, indx):
        return self._events[indx]

    def __len__(self):
        return len(self._events)

    def get_by_name(self, name):
        try:
            return [_ for _ in self._events if _['name'] == name][0]
        except IndexError:
            return None


class Meeting(OSMObject):

    def __init__(self, osm, section, accessor, record):
        OSMObject.__init__(self, osm, accessor, record)

        self.meeting_date = datetime.datetime.strptime(
            self._record['meetingdate'], '%Y-%m-%d')

        try:
            h, m, s = (int(i) for i in self._record['starttime'].split(':'))
            self.start_time = datetime.datetime.combine(
                self.meeting_date,
                datetime.time(h, m, s))
            self.start_time = dateutil.parser.parse(
                pyTZ.localize(self.start_time).strftime(FMT))
        except:
            # log.warning("Failed to interpret starttime ({}). Using 12:00:00. "
            #          "section: {}\n"
            #          "record = {}\n error = {}".format(
            #                 self._record['starttime'],
            #                 section['section_name'],
            #                 repr(record),
            #                 traceback.print_exc()))
            self.start_time = datetime.datetime.combine(
                self.meeting_date,
                datetime.time(12, 0, 0))
            self.start_time = dateutil.parser.parse(
                pyTZ.localize(self.start_time).strftime(FMT))

        try:
            h, m, s = (int(i) for i in self._record['endtime'].split(':'))
            self.end_time = datetime.datetime.combine(
                self.meeting_date,
                datetime.time(h, m, s))
            self.end_time = dateutil.parser.parse(
                pyTZ.localize(self.end_time).strftime(FMT))
        except:
            # log.warning("Failed to interpret endtime ({}). Using starttime. "
            #          "section: {}\n"
            #          "record = {}\n error = {}".format(
            #                 self._record['endtime'],
            #                 section['section_name'],
            #                 repr(record),
            #                 traceback.print_exc()))
            self.end_time = self.start_time

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

        if record != []:
            # If the record is an empty list it means that the programme
            # is empty.
            try:
                for event in record['items']:
                    events["{} - {}".format(
                        event['title'],
                        event['meetingdate'])] = Meeting(osm, section,
                                                         accessor, event)
            except:
                log.warning("Failed to process events in programme for "
                         "section: {0}\n"
                         "record = {1}\n error = {2}".format(
                    section['section_name'],
                    repr(record),
                    traceback.print_exc()))

        OSMObject.__init__(self, osm, accessor, events)

    def events_by_date(self):
        """Return a list of events sorted by start date (and time)"""

        return sorted(self._record.values(),
                      key=attrgetter('meeting_date'))


class ObjectTypes(enum.Enum):
    MEMBERS = 1
    PROGRAMME = 2
    EVENTS = 3
    USERS = 4
    MOVERS = 5


ALL_OBJECTS = (ObjectTypes.MEMBERS, ObjectTypes.PROGRAMME, ObjectTypes.EVENTS, ObjectTypes.USERS, ObjectTypes.MOVERS)


class Section(OSMObject):

    def __init__(self, osm, accessor, record, init=True, term=None, on_date=None, object_types=ALL_OBJECTS):
        OSMObject.__init__(self, osm, accessor, record)

        self.requested_term = term
        self.requested_date = on_date
        self.object_types = object_types

        self.members = []
        self.programme = []
        self.events = []
        self.users = []
        self.movers = []
        if init:
            self.init()

    def init(self):
        log.debug("Requested term = {}".format(self.requested_term))

        self.all_terms = [term for term in self._osm.terms(self['section_id'])]
        log.debug("All terms = {!r}".format(self.all_terms))

        if self.requested_term is not None:
            # We have requested a specific term.
            self.terms = [term for term in self._osm.terms(self['section_id'])
                          if term['name'].strip() == self.requested_term.strip()]

            if len(self.terms) != 1:
                log.warning("Requested term ({}) for section {} is not in available "
                         "terms ({})".format(
                    self.requested_term,
                    ",".join([term['name'] for term in self.terms]),
                    self['section_name']))
                # sys.exit(1)
        elif self.requested_date is not None:
            # We have requested a specific date. Need to find the term that encloses that date.
            self.terms = [term for term in self._osm.terms(self['section_id'])
                          if term.is_active(self.requested_date)]

            if len(self.terms) != 1:
                log.warning("Cannot find a term that encloses the requested date ({})"
                         "for section {} terms ({})".format(
                    self.requested_date,
                    self['section_name'], ",".join([term['name'] for term in self.terms]), ))
        else:
            self.terms = [term for term in self._osm.terms(self['section_id'])
                          if term.is_active()]

            log.debug("Active terms = {!r}".format(self.terms))

            if len(self.terms) > 1:
                log.error("{!r}: More than 1 term is active, picking "
                          "last in list {!r}".format(
                    self['section_name'],
                    [(term['name'], term['past']) for
                     term in self.terms]))

        if len(self.terms) == 0 or self.terms[-1]['name'] == 'All':
            # If there is no active term it does make sense to gather
            # badge info.
            self.term = None
            self.challenge = None
            self.activity = None
            self.staged = None
            self.core = None
        else:
            self.term = self.terms[-1]
            # self.challenge = self._get_badges('challenge')
            # self.activity = self._get_badges('activity')
            # self.staged = self._get_badges('staged')
            # self.core = self._get_badges('core')

        log.debug("Configured term = {}".format(self.term))

        if ObjectTypes.MEMBERS in self.object_types:
            try:
                self.members = self._get_members()
            except:
                log.warning("Failed to get members for section {0}"
                            .format(self['section_name']), exc_info=True)
                self.members = []
                raise

        if ObjectTypes.PROGRAMME in self.object_types:
            if self.term:
                try:
                    self.programme = self._get_programme()
                except:
                    log.warning("Failed to get programme for section {0}"
                                .format(self['section_name']),
                                 exc_info=True)
                    raise

        if ObjectTypes.EVENTS in self.object_types:
            try:
                self.events = self._get_events()
            except:
                log.warning("Failed to get events for section {0}"
                             .format(self['section_name']),
                             exc_info=True)
                raise

        if ObjectTypes.USERS in self.object_types:
            try:
                self.users = self._get_users()
            except:
                log.warning("Failed to get users for section {0}"
                             .format(self['section_name']),
                              exc_info=True)
                raise

        if ObjectTypes.MOVERS in self.object_types:
            try:
                self.movers = self._get_movers()
            except:
                log.warning("Failed to get movers for section {0}"
                             .format(self['section_name']),
                              exc_info=True)
                raise

    def __repr__(self):
        return 'Section({0}, "{1}", "{2}")'.format(
            self['section_id'],
            self['section_name'],
            self['section_type'])

    def _get_badges(self, badge_type):
        url = "challenges.php?action=getInitialBadges" \
              "&type={0}" \
              "&sectionid={1}" \
              "&section={2}" \
              "&termid={3}" \
            .format(badge_type,
                    self['section_id'],
                    self['section_type'],
                    self.term['termid'])

        return Badges(self._osm, self._accessor,
                      self._accessor(url), self, badge_type)

    # def events(self):
    #    pass

    def _get_events(self):
        url = "ext/events/summary/?action=get" \
              "&sectionid={0}" \
              "&termid={1}" \
            .format(self['section_id'],
                    self.term['termid'])

        return Events(self._osm, self, self._accessor,
                      self._accessor(url))

    def _get_users(self):
        url = "ext/settings/access/?action=getUsersForSection" \
              "&sectionid={0}".format(self['section_id'])

        return Users(self._osm, self, self._accessor,
                     self._accessor(url))

    def _get_movers(self):
        # Get the list of flexi record tables and look for "Moving On"
        flexi_url = ('ext/members/flexirecords/?action=getFlexiRecords'
                     '&sectionid={}&archived=n'.format(self['section_id']))
        flexi_records = self._accessor(flexi_url)

        if not flexi_records:
            log.warning("Can't access flexi records for {}"
                     .format(self['section_name']))
            return None

        moving_on_table = [item for item in flexi_records['items']
                           if item['name'] == 'Moving On']

        if not moving_on_table:
            log.warning("No movers table for section {} ({})"
                     .format(self['section_name'], self['section_id']))
            return []

        headers = self._accessor(('ext/members/flexirecords/?action=getStructure'
                                  '&sectionid={}&extraid={}'.format(self['section_id'],
                                                                    moving_on_table[0]['extraid'])))

        data = self._accessor(('ext/members/flexirecords/?action=getData&extraid={}&'
                               'sectionid={}&termid={}'
                               '&section=cubs'.format(moving_on_table[0]['extraid'],
                                                      self['section_id'],
                                                      self.term['termid'])))

        if headers and data:
            return Movers(self._osm, self, self._accessor,
                          headers, data)
        return None

    def _get_members(self):
        url = "ext/members/contact/grid/?action=getMembers" \
              "&section_id={0}" \
              "&term_id={1}" \
              "&dateFormat=uk" \
              "&section={2}" \
            .format(self['section_id'],
                    self.term['termid'],
                    self['section_type'])

        return Members(self._osm, self, self._accessor,
                       self._accessor(url))

    def _get_programme(self):
        url = "programme.php?action=getProgrammeSummary" \
              "&sectionid={0}&termid={1}".format(self['section_id'],
                                                 self.term['termid'])

        return Programme(self._osm, self, self._accessor,
                         self._accessor(url))

    def get_terms(self):
        return [term for term in self._osm.terms(self['section_id'])]

    def get_payments(self, start, end):
        url = "/ext/finances/onlinepayments/?action=exportWithdrawals" \
              "&sectionid={0}" \
              "&start={1}" \
              "&end={2}" \
            .format(self['section_id'],
                    start, end)

        return self._accessor(url, result_type='csv')


class OSM(object):

    def __init__(self, authorisor, sectionid_list=False, term=None,
                 on_date=None, object_types=ALL_OBJECTS):
        self._accessor = Accessor(authorisor)

        self.sections = {}
        self.section = None
        self.object_types = object_types

        self.init(sectionid_list, term, on_date)

    def init(self, sectionid_list=False, term=None, on_date=None):
        #roles = self._accessor('api.php?action=getUserRoles')
        roles = self._accessor._session.get("https://www.onlinescoutmanager.co.uk/oauth/resource").json()
        roles = roles['data']['sections']

        self.sections = {}

        for section in [Section(self, self._accessor, role,
                                init=False, term=term, on_date=on_date,
                                object_types=self.object_types)
                        for role in roles
                        #if 'section' in role
                        ]:
            if sectionid_list is False or \
                    str(section['section_id']) in sectionid_list:
                section.init()
                self.sections[str(section['section_id'])] = section

                # if section['isDefault'] == '1':
                #     self.section = section
                #     log.info("Default section = {0}, term = {1}".format(
                #         self.section['section_name'],
                #         self.section.term['name'] if self.section.term
                #         else "None"))

        if self.section is None:
            self.section = list(self.sections.values())[-1]

        # Warn if the active term is different for any of the
        # selected selections.

        if any([section.term is None for section in self.sections.values()]):
            log.warning("Term not set for at least one section")
        elif len(set([section.term['name'].strip() for section in
                      self.sections.values()])) != 1:
            log.warning("Not all sections have the same active term: \n "
                     "{}".format(
                "\n".join(
                    ["{} - {}".format(
                        section['section_name'], section.term['name'])
                        for section in self.sections.values()])))

    def terms(self, sectionid):
        sectionid = sectionid if type(sectionid) is str else str(sectionid)
        terms = self._accessor('api.php?action=getTerms')
        if sectionid in terms:
            return [Term(self, self._accessor, term) for term
                    in terms[sectionid]]
        return []


MemberClass = Member

if __name__ == '__main__':
    args = docopt(__doc__, version='OSM 2.0')

    loglevel = logging.DEBUG if args['-d'] else logging.WARN
    logging.basicConfig(level=loglevel)
    log.debug("Debug On\n")
    log.debug(args)

    sectionid_list = [args['-s'], ] if args['-s'] else [20706]

    #client_id = ''
    #client_secret = ''

    auth = Authorisor(open(DEF_CREDS, 'r'))

    if args['run']:
        accessor = Accessor(auth)

        pp.pprint(accessor(args['<query>']))

    osm = OSM(auth, sectionid_list, None, None, object_types=(ObjectTypes.MEMBERS, ))
    test_section = args['-s'] if args['-s'] else 20706
    members = osm.sections[test_section].members

    # import pdb
    # pdb.set_trace()
    for v in members.values():
        print(str(v))

    # log.debug('Sections - {0}\n'.format(osm.sections))

    # for badge in list(osm.sections[test_section].challenge.values()):
    #    log.debug('{0}'.format(badge._record))

    #    member = members[members.keys()[0]]
    # member['special'] = 'changed'
    # member.save()

    # new_member = members.new_member('New First 2','New Last 2','02/09/2004',
    #    '02/12/2012','02/11/2012')
    # log.debug("New member = {0}: {1}".format(new_member.firstname,
    #    new_member.lastname))
    # new_member.save()

    # for k,v in osm.sections['14324'].members.items():
    #    log.debug("{0}: {1} {2} {3}".format(k,v.firstname,
    #    v.lastname,v.TermtoScouts))

    # for k,v in osm.sections['14324'].activity.items():
    #    log.debug("{0}: {1}".format(k,v.keys()))

    # pp.pprint(osm.sections['14324'].members())
    # for k,v in osm.sections['14324'].challenge.items():
    #    log.debug("{0}: {1}".format(k,v.keys()))

    #Accessor.__cache_save__(open(DEF_CACHE, 'wb'))
