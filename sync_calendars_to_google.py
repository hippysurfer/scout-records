import sys
import yaml
import csv
import subprocess
import datetime
import dateutil.tz
from dateutil.parser import parse
import logging

import requests
from icalendar import Calendar, Event

from sync_contacts_to_google import to_csv_from_dict

log = logging.getLogger(__name__)

GAM = '/home/rjt/bin/gamx/gam'

# CALENDAR_ID = "7thlichfield.org.uk_gehutke2qdqmkh14cqj96cjsro@group.calendar.google.com"

in_cals = {}


def fetch_existing_events(calendar_id):
    try:
        out = subprocess.run(args=[GAM, 'calendar', calendar_id,
                                   'print', 'events'], stdout=subprocess.PIPE,
                             check=True,
                             universal_newlines=True)
        events = list(csv.DictReader(out.stdout.splitlines()))
    except subprocess.CalledProcessError:
        log.warning('Error in fetch_existing_events. Probably means no available events',
                    exc_info=True)
        events = []
    return events


def add_calendar_events(calendar_id, events):
    if len(events) > 0:
        csv_text = to_csv_from_dict(events)
        keys = [(k, f'~{k}') for k in events[0].keys()]
        fields = [_ for sub_list in keys for _ in sub_list]
        p = subprocess.Popen(stdin=subprocess.PIPE,
                             args=[GAM, 'csv', '-', 'gam', 'calendar', calendar_id,
                                   'add', 'event'] + fields,
                             universal_newlines=True)
        p.communicate(input=csv_text)
        p.wait()


def update_calendar_events(calendar_id, events):
    if len(events) > 0:
        csv_text = to_csv_from_dict(events)
        keys = [(k, f'~{k}') for k in events[0].keys() if k != 'calendarId']
        fields = [_ for sub_list in keys for _ in sub_list]
        del fields[fields.index('id')]
        p = subprocess.Popen(stdin=subprocess.PIPE,
                             args=[GAM, 'csv', '-', 'gam', 'calendar', calendar_id,
                                   'update', 'event', 'events'] + fields,
                             universal_newlines=True)
        p.communicate(input=csv_text)
        p.wait()


def delete_calendar_events(calendar_id, events):
    if len(events) > 0:
        csv_text = to_csv_from_dict(events)
        p = subprocess.Popen(stdin=subprocess.PIPE,
                             args=[GAM, 'csv', '-', 'gam', 'calendar', calendar_id,
                                   'delete', 'event', 'events', '~id', 'doit'],
                             universal_newlines=True)
        p.communicate(input=csv_text)
        p.wait()


def convert_osm_uid_to_valid_google_id(osm_uid):
    name = osm_uid.lower()
    name = name.replace('-', '')
    name = name.replace('/', '')
    name = name.replace(' ', '')
    return name


def sync_google_calendar(calendar_id, current_events):
    existing_events = fetch_existing_events(calendar_id)

    current_event_ids = set([event['id'] for event in current_events])
    existing_event_ids = set([event['id'] for event in existing_events])

    update_events_ids = current_event_ids & existing_event_ids
    add_event_ids = current_event_ids - existing_event_ids
    delete_event_ids = existing_event_ids - current_event_ids

    update_events = [_ for _ in current_events if _['id'] in update_events_ids]
    add_events = [_ for _ in current_events if _['id'] in add_event_ids]
    delete_events = [_ for _ in existing_events if _['id'] in delete_event_ids]

    # Update the update events.
    changed_events = []
    for event in update_events:
        current_event = [_ for _ in existing_events if _['id'] == event['id']][0]
        # event['start'] = parse(event['start'])
        # event['end'] = parse(event['end'])
        # log.warning((event['start'] == current_event['start.dateTime']))
        # log.warning(f"changed: '{event['start']}' != '{current_event['start.dateTime']}'")  
        description = event['description'].replace('\n','')
        res = [event['start'] == parse(current_event['start.dateTime']),
                    event['end'] == parse(current_event['end.dateTime']),
                    event['summary'] == current_event['summary'],
                    description == current_event['description'],
                    event['location'] == current_event['location']]
        if not all(res):
            log.warning(f"changed: '{event['start']}' != '{current_event['start.dateTime']}'")  
            log.warning(f"changed: '{event['end']}' != '{current_event['end.dateTime']}'")  
            log.warning(f"changed: '{event['summary']}' != '{current_event['summary']}'")  
            log.warning(f"changed: '{description}' != '{current_event['description']}'")  
            log.warning(f"changed: '{event['location']}' != '{current_event['location']}'")  
            # event['start'] = current_event['start.dateTime']
            # event['end'] = current_event['end.dateTime']
            # event['summary'] = current_event['summary']
            # event['description'] = current_event['description']
            # event['location'] = current_event['location']
            changed_events.append(event)
            log.warning(f"{res}")
            # for char in range(len(current_event['start.dateTime'])):
            #     if event['start'][char] == current_event['start.dateTime'][char]:
            #         log.warning(f"{event['start'][char]} != {current_event['start.dateTime'][char]}")
            # log.warning(type(event['start']))
            # log.warning(type(current_event['start.dateTime']))
        # else:
        #     log.warning(f"unchanged: {event['summary']} == {current_event['summary']}")  

    update_calendar_events(calendar_id, changed_events)
    add_calendar_events(calendar_id, add_events)
    delete_calendar_events(calendar_id, delete_events)


def read_config(filename):
    """Read YAML config file"""
    with open(filename, 'r') as f:
        return yaml.load(f)


def download_calendar(url):
    """Download and parse ics file"""
    req = requests.get(url)
    cal = Calendar.from_ical(req.text)
    return cal


def read_calendar(filename):
    """Parse local ics file"""
    with open(filename, 'r') as f:
        cal = Calendar.from_ical(f)
    return cal


def write_calendar(options, sources, google_calendar):
    """Create and write ics file"""
    cal = Calendar()
    google_events = []
    timezones_cache = []
    for key, value in options.items():
        cal.add(key, value)

    for source_id, category in sources.items():
        for timezone in in_cals[source_id].walk('VTIMEZONE'):
            if timezone['tzid'] not in timezones_cache:
                timezones_cache.append(timezone['tzid'])
                cal.add_component(timezone)
        for event in in_cals[source_id].walk('VEVENT'):
            event_copy = Event(event)
            event_copy['SUMMARY'] = f'{category}: ' + event['SUMMARY']
            event_copy.add('categories', category)
            cal.add_component(event_copy)
            # dt_utc = event_copy['DTSTART'].dt.astimezone(pytz.utc)

            start = event_copy['DTSTART'].dt
            start = (start if isinstance(start, datetime.datetime)
                     else datetime.datetime.combine(
                start, datetime.time(
                    hour=0, minute=0, second=0, tzinfo=dateutil.tz.tzutc())))

            end = event_copy['DTEND'].dt if event_copy.get('DTEND') else start

            end = (end if isinstance(end, datetime.datetime)
                   else datetime.datetime.combine(
                end, datetime.time(
                    hour=0, minute=0, second=0, tzinfo=dateutil.tz.tzutc())))

            google_event = {
                'id': convert_osm_uid_to_valid_google_id(
                    str(event_copy['UID'])),
                'start': start,
                'end': end,
                'summary': str(event_copy['SUMMARY']),
                'description': str(event_copy.get('DESCRIPTION', '')),
                'location': str(event_copy.get('LOCATION', ''))}
            google_events.append(google_event)

    sync_google_calendar(google_calendar, google_events)


def main():
    # Get the name of the config file
    if len(sys.argv) != 2:
        print('Usage:')
        print('   merge-ics <config_file>')
        return 1
    config_file = sys.argv[1]

    # Read config
    try:
        config = read_config(config_file)
    except IOError:
        print('Unable to open ' + config_file)
        return 1
    except yaml.YAMLError:
        print('Unable to parse ' + config_file)
        return 1

    # Read/download and parse input calendars
    for source_id, source in config['sources'].items():
        if 'filename' in source:
            try:
                in_cals[source_id] = read_calendar(source['filename'])
            except IOError:
                print('Unable to open ' + source['filename'])
            except ValueError:
                print('Unable to parse ' + source['filename'])
        if (source_id not in in_cals) and ('url' in source):
            try:
                in_cals[source_id] = download_calendar(source['url'])
            except requests.exceptions.RequestException:
                print('Unable to download ' + source['url'])
            except ValueError:
                print('Unable to parse ' + source['url'])

    # Create and write output calendars
    for sink in config['sinks']:
        try:
            write_calendar(sink['options'], sink['sources'], sink['google_calendar'])
        except IOError:
            print('Unable to write ' + sink['filename'])
        except ValueError:
            print('Unable to create calendar ' + sink['filename'])

if __name__ == '__main__':
    level = logging.INFO

    logging.basicConfig(level=level)
    
    main()
