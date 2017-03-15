# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  sync_contacts.py [-d] <vcard_dir> <section>...
  sync_contacts.py (-h | --help)
  sync_contacts.py --version


Options:
  <section>      Sections to sync
  <vcard_dir>    Directory of vcards
  -d,--debug     Turn on debug output.
  -h,--help      Show this screen.
  --version      Show version.
  -a             Request authorisation credentials.

"""

import os.path
import logging
from docopt import docopt
import socket

log = logging.getLogger(__name__)

from export_explorer_vcards import ExplorerGroup
from carddav_util.carddav import PyCardDAV
import vobject
import uuid

from owncloud_accounts import ACCOUNTS

import requests.packages.urllib3
requests.packages.urllib3.disable_warnings()


def _main(sections, vcard_dir):
    assert os.path.exists(vcard_dir) and os.path.isdir(vcard_dir)

    if isinstance(sections, str):
        sections = [sections, ]

    for section in sections:
        assert section in ExplorerGroup.SECTIONIDS.keys(), \
            "section must be in {!r}.".format(ExplorerGroup.SECTIONIDS.keys())

    for section in sections:
        vcard_file = "{}.vcf".format(section)
        user, passwd = ACCOUNTS[section]

        hostname = 'www.thegrindstone.me.uk' \
                   if not socket.gethostname() == 'rat' \
                   else 'localhost'

        url = 'https://{}/owncloud/'\
              'remote.php/carddav/addressbooks/{}/contacts/'.format(hostname,
                                                                    user)

        log.debug('Connecting to: {}'.format(url))
        dav = PyCardDAV(url, user=user, passwd=passwd,
                        write_support=True, auth='basic',
                        verify=False)


        abook = dav.get_abook()
        nCards = len(abook.keys())

        # for each card, delete the card
        curr = 1
        for href, etag in list(abook.items()):
            log.debug("Deleting {} of {}.".format(curr, nCards))
            curr += 1
            card = dav.delete_vcard(href, etag)

        # now read in the new cards and upload.
        with open(os.path.join(vcard_dir, vcard_file), 'r') as f:
            cards = []
            for card in vobject.readComponents(f, validate=True):
                cards.append(card)
            nCards = len(cards)

            log.info("Uploading {} cards.".format(nCards))

            curr = 1
            for card in cards:
                log.debug("Uploading {} of {}.".format(curr, nCards))

                if hasattr(card, 'prodid'):
                    del card.prodid

                if not hasattr(card, 'uid'):
                    card.add('uid')
                card.uid.value = str(uuid.uuid4())

                log.debug(type(card.serialize()))
                log.debug(card.serialize())

                dav.upload_new_card(card.serialize())

                curr += 1


if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.INFO

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    _main(args['<section>'], args['<vcard_dir>'])
