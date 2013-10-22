# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  update.py <apiid> <token>
  update.py <apiid> <token> -a <email> <password>
  update.py (-h | --help)
  update.py --version


Options:
  -h --help      Show this screen.
  --version      Show version.
  -a             Request authorisation credentials.

"""

import sys
import logging
from docopt import docopt
import osm

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"


if __name__ == '__main__':

    logging.basicConfig(level=logging.INFO)
    log.debug("Debug On\n")

    try:
        osm.Accessor.__cache_load__(open(DEF_CACHE, 'r'))
    except:
        log.debug("Failed to load cache file\n")

    args = docopt(__doc__, version='OSM 2.0')

    if args['-a']:
        auth = osm.Authorisor(args['<apiid>'], args['<token>'])
        auth.authorise(args['<email>'],
                       args['<password>'])
        auth.save_to_file(open(DEF_CREDS, 'w'))
        sys.exit(0)

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    # read in all of the personal details from each of the OSM sections

    sections = osm.OSM(auth)

    test_section = '15797'

    print(sections.sections)

    members = sections.sections[test_section].members

    for k,v in list(members.items()):
      log.debug("{0}: {1} {2} {3}".format(k,v.firstname,v.lastname,v.Email1))


    #   Search for duplicate YP references

    #   Search for multiple references to the same Family Reference
    #     warn if the details are not the same

    # connect to spreadsheet

    import gspread
    import creds

    # creds needs to contain a tuple of the following form
    #   creds = ('username','password')


    gc = gspread.login(*creds.creds)
    wks = gc.open("TestSpread").sheet1

    row = 2
    for k,v in list(members.items()):
      wks.update_acell("B%d" % row, v.firstname)
      wks.update_acell("C%d" % row, v.lastname)
      wks.update_acell("D%d" % row, v.Email1)
      row += 1

    
    wks.range('A1:B4')

#   for each YP reference
#      find entry to details tab and update field or add new row

#   for each Parent reference
#      find entry in parent tab and update field or add new row

#   for each YP without a reference
#      add row to 'errors' tab if name not already listed
