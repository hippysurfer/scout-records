#!/usr/bin/env python
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU Library General Public License as published by
# the Free Software Foundation; version 2 only
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Library General Public License for more details.
#
# You should have received a copy of the GNU Library General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.
#
# Copyright (C) 2007 Andrea Righi <righiandr@users.sf.net>
#
# Description:
#  - synchronize a local iCal (.ics) file with Google Calendar.
#
# Requirements:
#  - python-vobject, python-gdata, python-httplib2

"""
googlecalendarsync by Andrea Righi <righiandr@users.sf.net>
"""

__version__ = '0.4'

import sys, os, re, getopt, string, time, shutil
import vobject, httplib2, ConfigParser, md5

try:
        from xml.etree import ElementTree
except ImportError:
        from elementtree import ElementTree
import gdata.calendar.service
import gdata.service
import atom.service
import gdata.calendar
import atom
import traceback

# Google Calendar class.
class GoogleCalendar:
        def __init__(self, login, password, url, calendar='default'):
                self.private_url = url
                self.calendar_service = gdata.calendar.service.CalendarService()
                self.calendar_service.email = login
                self.calendar_service.password = password
                self.calendar_service.source = 'googlecalendarsync'
                self.calendar_service.ProgrammaticLogin()
                self._calendar = calendar

        # Properly encode unicode characters.
        def encode_element(self, el):
                return unicode(el).encode('ascii', 'replace')

        # Use the Google-compliant datetime format for single events.
        def format_datetime(self, date):
                try:
                        if re.match(r'^\d{4}-\d{2}-\d{2}$', str(date)):
                                return str(date)
                        else:
                                return str(time.strftime("%Y-%m-%dT%H:%M:%S.000Z", date.utctimetuple()))
                except Exception, e:
                        print type(e), e.args, e
                        return str(date)

        # Use the Google-compliant datetime format for recurring events.
        def format_datetime_recurring(self, date):
                try:
                        if re.match(r'^\d{4}-\d{2}-\d{2}$', str(date)):
                                return str(date).replace('-', '')
                        else:
                                return str(time.strftime("%Y%m%dT%H%M%SZ", date.utctimetuple()))
                except Exception, e:
                        print type(e), e.args, e
                        return str(date)

        # Use the Google-compliant alarm format.
        def format_alarm(self, alarm):
                google_minutes = [5, 10, 15, 20, 25, 30, 45, 60, 120, 180, 1440, 2880, 10080]
                m = re.match('-(\d+)( day[s]?, )?(\d+):(\d{2}):(\d{2})', alarm)
                try:
                        time = m.groups()
                        t = 60 * ((int(time[0]) - 1) * 24 + (23 - int(time[2]))) + (60 - int(time[3]))
                        # Find the closest minutes value valid for Google.
                        closest_min = google_minutes[0]
                        closest_diff = sys.maxint
                        for m in google_minutes:
                                diff = abs(t - m)
                                if diff == 0:
                                        return m
                                if diff < closest_diff:
                                        closest_min = m
                                        closest_diff = diff
                        return closest_min
                except:
                        return 0

        # Convert a iCal event to a Google Calendar event.
        def ical2gcal(self, e, dt):
                # Parse iCal event.
                event = {}
                event['uid'] = self.encode_element(dt.uid.value)
                event['subject'] = self.encode_element(dt.summary.value)
                if hasattr(dt, 'description') and (dt.description is not None):
                        event['description'] = self.encode_element(dt.description.value)
                else:
                        event['description'] = ''
                if hasattr(dt, 'location'):
                        event['where'] = self.encode_element(dt.location.value)
                else:
                        event['where'] = ''
                if hasattr(dt, 'status'):
                        event['status'] = self.encode_element(dt.status.value)
                else:
                        event['status'] = 'CONFIRMED'
                if hasattr(dt, 'organizer'):
                        event['organizer'] = self.encode_element(dt.organizer.params['CN'][0])
                        event['mailto'] = self.encode_element(dt.organizer.value)
                        event['mailto'] = re.search('(?<=MAILTO:).+', event['mailto']).group(0)
                if hasattr(dt, 'rrule'):
                        event['rrule'] = self.encode_element(dt.rrule.value)
                if hasattr(dt, 'dtstart'):
                        event['start'] = dt.dtstart.value
                if hasattr(dt, 'dtend'):
                        event['end'] = dt.dtend.value
                if hasattr(dt, 'valarm'):
                        event['alarm'] = self.format_alarm(self.encode_element(dt.valarm.trigger.value))

                # Convert into a Google Calendar event.
                try:
                        e.title = atom.Title(text=event['subject'])
                        e.extended_property.append(gdata.calendar.ExtendedProperty(name='local_uid', value=event['uid']))
                        e.content = atom.Content(text=event['description'])
                        e.where.append(gdata.calendar.Where(value_string=event['where']))
                        e.event_status = gdata.calendar.EventStatus()
                        e.event_status.value = event['status']
                        if event.has_key('organizer'):
                                attendee = gdata.calendar.Who()
                                attendee.rel = 'ORGANIZER'
                                attendee.name = event['organizer']
                                attendee.email = event['mailto']
                                attendee.attendee_status = gdata.calendar.AttendeeStatus()
                                attendee.attendee_status.value = 'ACCEPTED'
                                if len(e.who) > 0:
                                        e.who[0] = attendee
                                else:
                                        e.who.append(attendee)
                        # TODO: handle list of attendees.
                        if event.has_key('rrule'):
                                # Recurring event.
                                recurrence_data = ('DTSTART;VALUE=DATE:%s\r\n'
                                        + 'DTEND;VALUE=DATE:%s\r\n'
                                        + 'RRULE:%s\r\n') % ( \
                                        self.format_datetime_recurring(event['start']), \
                                        self.format_datetime_recurring(event['end']), \
                                        event['rrule'])
                                e.recurrence = gdata.calendar.Recurrence(text=recurrence_data)
                        else:
                                # Single-occurrence event.
                                if len(e.when) > 0:
                                        e.when[0] = gdata.calendar.When(start_time=self.format_datetime(event['start']), \
                                                                        end_time=self.format_datetime(event['end']))
                                else:
                                        e.when.append(gdata.calendar.When(start_time=self.format_datetime(event['start']), \
                                                                          end_time=self.format_datetime(event['end'])))
                                if event.has_key('alarm'):
                                        # Set reminder.
                                        for a_when in e.when:
                                                if len(a_when.reminder) > 0:
                                                        a_when.reminder[0].minutes = event['alarm']
                                                else:
                                                        a_when.reminder.append(gdata.calendar.Reminder(minutes=event['alarm']))
                except Exception, e:
                        print >> sys.stderr, 'ERROR: couldn\'t create gdata event object: ', event['subject']
                        print type(e), e.args, e
                        sys.exit(1)

        # Return the list of events in the Google Calendar.
        def elements(self):
                try:
                        feed = self.calendar_service.GetCalendarEventFeed(
                                uri='/calendar/feeds/%s/private/full' % (self._calendar,))
                except:
                        print >> sys.stderr, 'ERROR: couldn\'t retrieve Google Calendar event list'
                        sys.exit(1)
                ret = []
                for i, event in enumerate(feed.entry):
                        ret.append(event)
                return ret

        # Fix all the Google Calendar events adding the extended property
        # "local_uid" used to properly the single events.
        def fix_remote_uids(self):
                for e in self.elements():
                        found = False
                        if e.extended_property:
                                for num, p in enumerate(e.extended_property):
                                        if (p.name == 'local_uid'):
                                                found = True
                        if not found:
                                id = os.path.basename(e.id.text) + '@google.com'
                                print 'fixing', id, 'for event', e.id.text
                                e.extended_property.append(gdata.calendar.ExtendedProperty(name='local_uid', value=id))
                                try:
                                        new_event = self.calendar_service.UpdateEvent(e.GetEditLink().href, e)
                                        print 'Fixed event (%s): %s' % (self.private_url, new_event.id.text,)
                                except:
                                        print >> sys.stderr, 'WARNING: couldn\'t update entry %s to %s!' % (id, self.private_url)

        # Translate a remote uid into the local uid.
        def get_local_uid(self, uid):
                for e in self.elements():
                        local_uid = os.path.basename(e.id.text)
                        if not re.match(r'@google\.com$', local_uid):
                                local_uid = local_uid + '@google.com'
                        if (local_uid == uid):
                                if e.extended_property:
                                        for num, p in enumerate(e.extended_property):
                                                if (p.name == 'local_uid'):
                                                        return p.value
                return None

        # Retrieve an event from Google Calendar by local UID.
        def get_event_by_uid(self, uid):
                for e in self.elements():
                        if e.extended_property:
                                for num, p in enumerate(e.extended_property):
                                        if (p.name == 'local_uid'):
                                                if (p.value == self.get_local_uid(uid)):
                                                        return e
                return None

        # Insert a new Google Calendar event.
        def insert(self, event):
                try:
                        e = gdata.calendar.CalendarEventEntry()
                except:
                        print >> sys.stderr, 'ERROR: couldn\'t create gdata calendar object'
                        sys.exit(1)
                self.ical2gcal(e, event)
                try:
                        # Get calendar private feed URL.
                        new_event = self.calendar_service.InsertEvent(e, '/calendar/feeds/%s/private/full' % (self._calendar,))
                        print 'New event inserted (%s): %s' % (self.private_url, new_event.id.text,)
                except Exception, e:
                        print >> sys.stderr, 'WARNING: couldn\'t insert entry %s to %s (%s)!' \
                                % (event.uid.value, self.private_url,
                                   '/calendar/feeds/%s/private/full' % (self._calendar,))
                        print type(e), e.args, e
                return new_event

        # Update a Google Calendar event.
        def update(self, event):
                e = self.get_event_by_uid(event.uid.value)
                if e is None:
                        print >> sys.stderr, 'WARNING: event %s not found in %s' % (event.uid.value, self.private_url)
                        return
                self.ical2gcal(e, event)
                try:
                        new_event = self.calendar_service.UpdateEvent(e.GetEditLink().href, e)
                        print 'Updated event (%s): %s' % (self.private_url, new_event.id.text,)
                except Exception, e:
                        print >> sys.stderr, 'WARNING: couldn\'t update entry %s to %s!' % (event.uid.value, self.private_url)
                        print type(e), e.args, e
                return new_event

        # Delete a Google Calendar event.
        def delete(self, event):
                e = self.get_event_by_uid(event.uid.value)
                if e is None:
                        print >> sys.stderr, 'WARNING: event %s not found in %s!' % (event.uid.value, self.private_url)
                        return
                try:
                        ret = self.calendar_service.DeleteEvent(e.GetEditLink().href)
                        #ret = self.calendar_service.DeleteEvent(e.GetEditLink().href)
                        print 'Deleted event (%s): %s: (%s)' % (self.private_url, e.id.text,
                                                                repr(ret))
                except Exception, e:
                        print >> sys.stderr, 'WARNING: couldn\'t delete entry %s in %s!' % (event.uid.value, self.private_url)
                        print type(e), e.args, e

        # List all the Google Calendar events.
        def list(self):
                for e in self.elements():
                        print e.title.text, '-->', e.id.text

        # Commit changes to Google Calendar.
        def sync(self):
                print 'Synchronized ', self.private_url
                pass

class iCalCalendar:
        def __init__(self, url, login=None, password=None):
                self.url = url
                m = re.match('^http', self.url)
                try:
                        if m:
                                # Remote calendar.
                                h = httplib2.Http()
                                h.add_credentials(login, password)
                                h.follow_all_redirects = True
                                resp, content = h.request(self.url, "GET")
                                assert(resp['status'] == '200')
                        else:
                                # Local calendar.
                                stream = file(self.url)
                                content = stream.read()
                                stream.close()
                        self.cal = vobject.readOne(content, findBegin='false')
                except:
                        # Create an empty calendar object.
                        self.cal = vobject.iCalendar()

        # Return the list of events in the iCal Calendar.
        def elements(self):
                ret = []
                for event in self.cal.components():
                        if (event.name == 'VEVENT') and hasattr(event, 'summary') and hasattr(event, 'uid'):
                                ret.append(event)
                return ret

        # Retrieve an event from Google Calendar by local UID.
        def get_event_by_uid(self, uid):
                for e in self.elements():
                        if e.uid.value == uid:
                                return e
                return None

        # Insert a new iCal event.
        def insert(self, event):
                self.cal.add(event)
                print 'New event inserted (%s): %s' % (self.url, event.uid.value)
                return event

        # Update a Google Calendar event.
        def update(self, event):
                e = self.get_event_by_uid(event.uid.value)
                if e is None:
                        print >> sys.stderr, 'WARNING: event %s not found in %s!' % (event.uid.value, self.url)
                        return
                e.copy(event)
                print 'Updated event (%s): %s' % (self.url, e.uid.value,)
                return event

        # Delete a iCal Calendar event.
        def delete(self, event):
                e = self.get_event_by_uid(event.uid.value)
                self.cal.remove(e)
                print 'Deleted event (%s): %s' % (self.url, e.uid.value,)

        # List all the iCal events.
        def list(self):
                for event in self.elements():
                        print event.summary.value, '-->', event.uid.value

        # Commit changes to iCal Calendar.
        def sync(self):
                print 'Synchronized ', self.url
                m = re.match('^http', self.url)
                if m:
                        print >> sys.stderr, 'ERROR: couldn\'t sync a remote calendar directly: ', self.url
                        sys.exit(1)
                try:
                        f = open(self.url, 'w')
                        f.write(unicode(self.cal.serialize()).encode('ascii', 'replace'))
                        f.close()
                except Exception, e:
                        print >> sys.stderr, 'ERROR: couldn\'t write to local calendar: ', self.url
                        print type(e), e.args, e
                        sys.exit(1)

# Class used for logging stuff.
class Logger:
        def __init__(self, f):
                self.f = f

        def write(self, s):
                self.f.write(s)
                self.f.flush()

#config_file = os.getenv('HOME') + '/.googlecalsync/config'

config_file = 'config'

default_configuration = """\
[google]
username = <GOOGLE ACCOUNT USERNAME>
password = <GOOGLE ACCOUNT PASSWORD>

[local]
ical_file = <PATH OF THE LOCAL iCal FILE>
workdir = work"""

def version():
        print 'GoogleCalSync version', __version__
        print """
Copyright (C) 2007 Andrea Righi <righiandr@users.sf.net>
This is free software; see the source for copying conditions.  There is NO
warranty; not even for MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
"""

def usage():
        print """\
Description: synchronize a local iCal (.ics) file with Google Calendar.

Create a configuration file (""" + config_file + """)

Use the following template:
##########
""" + default_configuration + """
##########

Then run:
  googlecalsync.py [-n | --dry-run]
"""

# Main.
if __name__ == '__main__':

### Get command line options ###

        try:
                opts, args = getopt.getopt(sys.argv[1:], "hvd", ["help", "version", "dry-run"])
        except getopt.GetoptError:
                usage()
                sys.exit(1)

        dry_run = False
        for o, a in opts:
                if o in ('-h', '--help'):
                        version()
                        usage()
                        sys.exit(0)
                if o in ('-v', '--version'):
                        version()
                        sys.exit(0)
                if o in ('d', '--dry-run'):
                        dry_run = True

### Parse configuration file ###

        if not os.path.isfile(config_file):
                print >> sys.stderr, 'ERROR: configuration file not found:', config_file
                sys.exit(1)

        config = ConfigParser.ConfigParser()
        # Get mandatory parameters.
        try:
                config.read(config_file)
                login = config.get('google', 'username')
                password = config.get('google', 'password')
                calendar = 'ork75rqndt9m073rmthjebg168@group.calendar.google.com'
                private_url = 'https://www.google.com/calendar/ical/' + calendar + '/private/basic.ics'
                #private_url = 'https://www.google.com/calendar/ical/' + login + '/private/basic.ics'
                local_cal_file = os.path.expandvars(config.get('local', 'ical_file'))
                workdir = os.path.expandvars(config.get('local', 'workdir'))
        except Exception, e:
                print 'ERROR: not a valid configuration file, check', config_file
                print type(e), e.args, e
                sys.exit(1)
        # Get optional parameters.
        try:
                logfile = os.path.expandvars(config.get('local', 'logfile'))
        except:
                logfile = None

### Initialization ###

        # Create working directory.
        if not os.path.isdir(workdir):
                try:
                        os.makedirs(workdir)
                except:
                        print >> sys.stderr, 'ERROR: couldn\'t make working directory:', workdir
                        sys.exit(1)

        # Open the log file.
        if (logfile is not None):
                try:
                        pass
                except:
                        print >> sys.stderr, 'ERROR: couldn\'t initialize log file:', logfile
                        sys.exit(1)

        # Initialize local calendar object.
        try:
                ical = iCalCalendar(local_cal_file)
        except:
                print >> sys.stderr, 'ERROR: couldn\'t initialize local calendar object with:', local_cal_file
                sys.exit(1)

        # Initialize remote calendar object.
        try:
                gcal = GoogleCalendar(login, password, private_url, calendar)
        except:
                print >> sys.stderr, 'ERROR: couldn\'t initialize Google Calendar object with:', private_url
                sys.exit(1)

        # Fix remote calendar events.
        print 'fixing remote events ID for:', private_url
        if not dry_run:
                gcal.fix_remote_uids()

### Delete all remote events ###

        # Open remote calendar object in iCal format.
        #try:
        #        ical_remote = iCalCalendar(private_url, login, password)
        #except:
        #        print >> sys.stderr, 'ERROR: couldn\'t open remote iCal object:', private_url
        #        sys.exit(1)

        # Sync new and updated items.

        e = gcal.elements()
        print "Deleting %d events" % (len(e),)
        for event in e:
                print 'Deleting event: '
                if not dry_run:
                        gcal.delete(event)


        #sys.exit(0)

### Add all local events ###

        # Synchronize new and updated items.
        for event in ical.elements():

                ### Insert ###
                print 'inserting new event', event.uid.value, 'to Google Calendar'
                if not dry_run:
                        try:
                                gcal.insert(event)
                        except Exception as e:
                                print sys.stderr, 'WARNING: couldn\'t insert entry in Google Calendar:', event.uid.value
                                traceback.print_exc()

        # Commit changes.
        print 'committing changes to Google Calendar'
        if not dry_run:
                ### Sync ###
                gcal.sync()


