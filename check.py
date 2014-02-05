# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  check.py [-d] <apiid> <token>
  check.py (-h | --help)
  check.py --version


Options:
  -d,--debug     Turn on debug output.
  -h,--help      Show this screen.
  --version      Show version.

"""

import logging
from docopt import docopt
import osm

from group import Group, OSM_REF_FIELD
from update import MAPPING

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"


def _main(osm, auth):
    #test_section = '15797'

    group = Group(osm, auth, MAPPING.keys())

    if group.missing_adult_references():
        log.warn("Missing adult references {!r}".format(
            group.missing_adult_references()))

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

    # creds needs to contain a tuple of the following form
    #     creds = ('username','password')

    _main(osm, auth)
