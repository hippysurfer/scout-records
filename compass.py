# coding: utf-8
"""Online Scout Manager Interface.

Usage:
  export_compass.py [-d] <user> <password> <outdir> <section>... 
  export_compass.py (-h | --help)
  export_compass.py --version


Options:
  <user>         Username
  <password>     Password
  <section>      Section to export.
  <outdir>       Output directory for cvs files.
  -d,--debug     Turn on debug output.
  -h,--help      Show this screen.
  --version      Show version.

"""

from docopt import docopt
import os.path
import logging
import time
import csv

from pyvirtualdisplay import Display
from splinter import Browser

log = logging.getLogger(__name__)


class Compass:

    def __init__(self, username='', password='', outdir=''):
        self._username = username
        self._password = password
        self._outdir = outdir
        self._browser = None

    def quit(self):
        if self._browser:
            self._browser.quit()
            self._browser = None

    def loggin(self):
        prefs = {
            "browser.download.folderList": 2,
            "browser.download.manager.showWhenStarting": False,
            "browser.download.dir": self._outdir,
            "browser.helperApps.neverAsk.saveToDisk": "application/octet-stream,application/msexcel,application/csv"}

        self._browser = Browser('firefox', profile_preferences=prefs)

        self._browser.visit('https://compass.scouts.org.uk/login/User/Login')

        self._browser.fill('EM', self._username)
        self._browser.fill('PW', self._password)
        time.sleep(2)
        self._browser.find_by_value('Submit').first.click()

        # Look for the Role selection menu and select my Group Admin role.
        self._browser.is_element_present_by_name(
            'ctl00$UserTitleMenu$cboUCRoles',
            wait_time=30)
        self._browser.select('ctl00$UserTitleMenu$cboUCRoles', '1253644')

    def export(self, section):
        # Select the My Scouting link.
        self._browser.is_text_present('My Scouting', wait_time=30)
        self._browser.click_link_by_text('My Scouting')

        def wait_then_click_xpath(xpath, wait_time=30):
            self._browser.is_element_present_by_xpath(
                xpath, wait_time=wait_time)
            self._browser.find_by_xpath(xpath).click()

        # Click the "Group Sections" hotspot.
        wait_then_click_xpath('//*[@id="TR_HIER7"]/h2')

        # Clink the link that shows the number of members in the section.
        # This is the one bit that is section specific.
        # We might be able to match on the Section name in the list,
        # which would make it more robust but at present we just hard
        # the location in the list.
        section_map = {
            'garrick': 2,
            'paget': 3,
            'swinfen': 4,
            'brown': 4,
            'maclean': 5,
            'rowallan': 6,
            'somers': 7,
            'boswell': 8,
            'erasmus': 9,
            'johnson': 10
        }
        wait_then_click_xpath(
            '//*[@id="TR_HIER7_TBL"]/tbody/tr[{}]/td[4]/a'.format(
                section_map[section.lower()]
            ))

        # Click on the Export button.
        wait_then_click_xpath('//*[@id="bnExport"]')

        # Click to say that we want a CSV output.
        wait_then_click_xpath(
            '//*[@id="tbl_hdv"]/div/table/tbody/tr[2]/td[2]/input')
        time.sleep(2)

        # Click to say that we want all fields.
        wait_then_click_xpath('//*[@id="bnOK"]')

        download_path = os.path.join(self._outdir, 'CompassExport.csv')

        if os.path.exists(download_path):
            log.warn("Removing stale download file.")
            os.remove(download_path)

        # Click the warning.
        wait_then_click_xpath('//*[@id="bnAlertOK"]')

        # Browser will now download the csv file into outdir. It will be called
        # CompassExport.

        # Wait for file.
        timeout = 30
        while not os.path.exists(download_path):
            time.sleep(1)
            timeout -= 1
            if timeout <= 0:
                log.warn("Timeout waiting for {} export to download.".fomat(
                    section
                ))
                break

        # rename download file.
        os.rename(download_path,
                  os.path.join(self._outdir, '{}.csv'.format(section)))

        log.info("Completed download for {}.".format(section))

        # Draw breath
        time.sleep(1)

    def load_from_dir(self):
        # Load the records form the set of files in self._outdir.

        log.debug('Loading from {}'.format(self._outdir))

        self._records_by_section = {}
        for section in os.listdir(self._outdir):
            section_name = os.path.splitext(section)[0]
            
            log.debug('Loading Compass data for {}'.format(section_name))
            reader = csv.DictReader(open(
                os.path.join(self._outdir, section)))

            self._records_by_section[section_name] = list(reader)

    def find_by_name(self, firstname, lastname, section_wanted=None):
        """Return list of matching records."""
        l = []
        for section in self._records_by_section.keys():
            if (section_wanted and section_wanted != section):
                continue
            for r in self._records_by_section[section]:
                if (r['forenames'].strip().lower() == firstname.strip().lower() and
                        r['surname'].strip().lower() == lastname.strip().lower()):
                    l.append(r)

        return l

    def sections(self):
        "Return a list of the sections for which we have data."
        return self._records_by_section.keys()

    def all_yp_members_dict(self):
        return {s: self.section_all_members(s) for
                s in self.sections()}

    def section_all_members(self, section):
        return self._records_by_section[section]

    def section_yp_members_without_leaders(self, section):
        return [member for member in
                self.section_all_members(section)
                if member['role'].lower() in
                ['Beaver Scout', 'Cub Scout', 'Scout']]


def _main(username, password, sections, outdir):

    display = Display(visible=0, size=(1920, 1080))

    try:
        display.start()

        compass = Compass(username, password, outdir)
        compass.loggin()

        try:
            for section in sections:
                compass.export(section)
        finally:
            compass.quit()

    finally:
        display.stop()

if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.INFO

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    log.debug(args['<section>'])

    _main(args['<user>'], args['<password>'],
          args['<section>'], os.path.abspath(args['<outdir>']))
