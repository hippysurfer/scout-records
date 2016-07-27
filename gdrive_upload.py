#!/usr/bin/env python
# coding: utf-8
"""Upload file to Google Drive

Usage:
  gdrive_upload.py [-d] <file> [<folder>] [<new_name>]
  gdrive_upload.py (-h | --help)
  gdrive_upload.py --version


Options:
  <file>                Path to file to upload
  <folder>              ID of Drive folder
  <new_name>            New name of file on Drive
  -d,--debug            Turn on debug output.
  -h,--help             Show this screen.
  --version             Show version.
"""

import logging

from docopt import docopt
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

log = logging.getLogger(__name__)

flags = None

SCOPES = 'https://www.googleapis.com/auth/drive.file'
store = file.Storage('storage.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
    creds = tools.run_flow(flow, store, flags) \
        if flags else tools.run(flow, store)
DRIVE = build('drive', 'v3', http=creds.authorize(Http()))


def upload(path, folder=None, filename=None, mimetype='application/vnd.google-apps.sheet'):
    filename = path if not filename else filename

    metadata = {'name': filename,
                'mimeType': mimetype}
    if folder:
        metadata.update({'parents': [folder,]})

    res = DRIVE.files().create(body=metadata, media_body=path).execute()
    if res:
        log.debug('Uploaded "%s" (%s)' % (filename, res['mimeType']))
    else:
        log.debug('Upload failed')


if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.WARN

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    file = args['<file>']
    folder = args['<folder>'] if args['<folder>'] else None
    new_name = args['<new_name>'] if args['<new_name>'] else None

    upload(file, folder, new_name)