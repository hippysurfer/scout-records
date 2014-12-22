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
from docopt import docopt
import osm
import csv

from group import Group
from update import MAPPING

from compass import member2compass
from compass import check
from compass import compass_headings

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"


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























