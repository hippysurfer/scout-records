# coding=utf-8
"""OSM Command Line

Usage:
   cli <apiid> <token> <section> contacts list
   cli <apiid> <token> <section> events list
   cli <apiid> <token> <section> events <event> attendees
   cli <apiid> <token> <section> events <event> info

Options:
   

"""

import logging
log = logging.getLogger(__name__)

from docopt import docopt


if __name__ == '__main__':
    level = logging.INFO

    logging.basicConfig(level=level)

    args = docopt(__doc__, version='OSM 2.0')

    if args['events']:
        if args['list']:
            log.info('events list')
        elif args['attendees']:
            log.info('events attendees')
        elif args['info']:
            log.info('events info')
        else:
            log.error('unknown')
    elif args['contacts']:
        log.info("contacts")
    else:
        log.error('unknown')
        
        
