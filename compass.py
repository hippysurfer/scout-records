# coding: utf-8
"""Online Scout Manager Interface.

Usage:
  compass.py [-d] <user> <password> <outdir> <section>...
  compass.py (-h | --help)
  compass.py --version


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
import datetime
import re
import pandas as pd
from io import StringIO
from lxml import etree
from contextlib import contextmanager
from functools import lru_cache

from pyvirtualdisplay import Display
from splinter import Browser

log = logging.getLogger(__name__)

postcode = re.compile(
    "(GIR ?0AA|[A-PR-UWYZ]([0-9]{1,2}|([A-HK-Y][0-9]([0-9ABEHMNPRV-Y])?)|[0-9][A-HJKPS-UW]) ?[0-9][ABD-HJLNP-UW-Z]{2})")

date_format = '%Y-%m-%d'

compass_headings = ("Membership Number",
                    "Look-up Title **",
                    "Forename(s)**",
                    "Surnames **",
                    "Look-up Role **",
                    "Look-up Youth  Leadership",
                    "Date of Birth **",
                    "Address Line 1 **",
                    "Address Line 2",
                    "Address Line 3",
                    "Address Line 4",
                    "Town **",
                    "Postal County",
                    "Postcode **",
                    "Look-up Postal Country **",
                    "Telephone",
                    "Email *",
                    "Start Date for this role**",
                    "Intentionally Blank Column",
                    "Name of Lodge/Six/Patrol",
                    "Date they joined Scouts*",
                    "Look-up Nationality *",
                    "Look-up Ethnicity *",
                    "Look-up Faith/Religion *",
                    "Look-up Parent 1 Title *",
                    "Parent 1 Forename *",
                    "Parent 1 Surname *",
                    "Parent 1 Known As",
                    "Parent 1 Date of Birth",
                    "Look-up Parent 1 Gender *",
                    "Look-up Parent 1 Relationship *",
                    "Parent 1 Email 1 *",
                    "Parent 1 Email 2",
                    "Parent 1 Telephone 1 *",
                    "Parent 1 Telephone 2",
                    "Look-up Parent 1 Gift Aid",
                    "Parent 1 Other Information",
                    "Look-up Parent 2 Title*",
                    "Parent 2 Forename *",
                    "Parent 2 Surname*",
                    "Parent 2 Known As",
                    "Parent 2 Date of Birth",
                    "Look-up Parent 2 Gender*",
                    "Look-up Parent 2 Relationship*",
                    "Parent 2 Email 1 *",
                    "Parent 2 Email 2",
                    "Parent 2 Telephone 1*",
                    "Parent 2 Telephone 2",
                    "Parent 2 Gift Aid",
                    "Parent 2 Other Information",
                    "Emergency Contact Forename *",
                    "Emergency Contact Surname *",
                    "Emergency Contact Known As",
                    "Emergency Contact Relationship *",
                    "Emergency Contact Telephone Number 1 *",
                    "Emergency Contact Telephone Number 2",
                    "Emergency Contact Telephone Number 3",
                    "Doctor / Surgery*",
                    "Surgery Address 1",
                    "Surgery Address 2",
                    "Surgery Address 3",
                    "Surgery Town",
                    "Surgery County",
                    "Surgery Postcode",
                    "Look-up Surgery Country",
                    "Surgery Telephone 1 *",
                    "Surgery Telephone 2",
                    "NHS Number",
                    "Dietary Needs",
                    "Medical Information")

required_headings = ("Membership Number",
                     "Look-up Title **",
                     "Forename(s)**",
                     "Surnames **",
                     "Look-up Role **",
                     "Date of Birth **",
                     "Address Line 1 **",
                     "Town **",
                     "Postcode **",
                     "Look-up Postal Country **",
                     "Email *",
                     "Start Date for this role**",
                     "Date they joined Scouts*",
                     "Look-up Nationality *",
                     "Look-up Ethnicity *",
                     "Look-up Faith/Religion *",
                     "Look-up Parent 1 Title *",
                     "Parent 1 Forename *",
                     "Parent 1 Surname *",
                     "Look-up Parent 1 Gender *",
                     "Look-up Parent 1 Relationship *",
                     "Parent 1 Email 1 *",
                     "Parent 1 Telephone 1 *",
                     "Emergency Contact Forename *",
                     "Emergency Contact Surname *",
                     "Emergency Contact Relationship *",
                     "Emergency Contact Telephone Number 1 *",
                     "Doctor / Surgery*",
                     "Surgery Telephone 1 *")


def parse_tel(number_field, default_name):
    index = 0
    for i in range(len(number_field)):
        if number_field[i] not in ['0', '1', '2', '3', '4', '5',
                                   '6', '7', '8', '9', ' ', '\t']:
            index = i
            break

    number = number_field[:index].strip() if index != 0 else number_field
    name = number_field[index:].strip() if index != 0 else default_name

    return number, name


def parse_addr(addr):
    a = {}
    # try to split on comma
    l = addr.split(',')

    if len(l) == 1:
        words = l[0].split(" ")
        post_found = False
        if postcode.search(" ".join(words[-2:-1])):
            # Are the last 2 words a postcode?
            a['postcode'] = " ".join(words[-2:-1])
            post_found = True
        else:
            a['postcode'] = "WS13 6ET"

        # Does it contain "lichfield"
        inx = l[0].lower().find("lichfield")
        if inx != -1:
            # Street is all the characters up to "lichfield"
            a['street'] = l[0][0:inx]
            a['town'] = "Lichfield"
        else:
            if post_found:
                # Street is all characters up to postcode
                a['street'] = l[0][0:l[0].find(a['postcode'])]
            else:
                a['street'] = l[0]
                a['town'] = "Lichfield"

        # Just the street?
        a['street'] = l[0]
        a['town'] = "Lichfield"
        a['postcode'] = "WS13 6ET"
    elif len(l) == 2:
        # Is the last one a postcode?
        if postcode.search(l[1]):
            a['street'] = l[0]
            a['town'] = "Lichfield"
            a['postcode'] = l[1]
        elif l[1].lower() == 'lichfield':
            a['street'] = l[0]
            a['town'] = "Lichfield"
            a['postcode'] = "WS13 6ET"
        else:
            a['street'] = " ".join(l)
            a['town'] = "Lichfield"
            a['postcode'] = "WS13 6ET"
    elif len(l) == 3:
        # Probably got all 3
        a['street'] = l[0]
        a['town'] = l[1]
        a['postcode'] = l[2]
    elif len(l) > 3:
        # Is the last one a postcode?
        if postcode.search(l[-1]):
            a['street'] = ",".join(l[:-2])
            a['town'] = l[-2]
            a['postcode'] = l[-1]
        else:
            a['street'] = ",".join(l[:-1])
            a['town'] = l[-1]
            a['postcode'] = "WS13 6ET"
    return a


def get_tel(member, priority_list):
    for i in priority_list:
        if member[i].strip() != '':
            tel = parse_tel(member[i], '')[0]
            if not tel.startswith('0'):
                tel = "01543 {}".format(tel)
            return tel
    return "01543 123456"


def get_parent(member):
    p = {}
    if (member['MumEmail'].strip() != ''
            and not member['MumEmail'].startswith('x ')):
        p['title'] = "Ms"
        p['forename'] = member['MumsName'] \
            if member['MumsName'].strip() != ''\
            else member['DadsName']
        p['surname'] = member['lastname']
        p['gender'] = 'Unknown'
        p['relationship'] = 'Mother'
        p['email'] = member['MumEmail']
        p['tel'] = get_tel(member, ('HomeTel', 'MumMob', 'DadMob'))
    else:
        p['title'] = "Mr"
        p['forename'] = member['DadsName'] \
            if member['DadsName'].strip() != ''\
            else member['MumsName']
        p['surname'] = member['lastname']
        p['gender'] = 'Unknown'
        p['relationship'] = 'Father'
        p['email'] = member['DadEmail']
        p['tel'] = get_tel(member, ('HomeTel', 'DadMob', 'MumMob'))

    parts = p['forename'].split(' ')
    if len(parts) > 1:
        if parts[-1].strip() == p['surname']:
            p['forename'] = " ".join(parts[:-1])

    return p


def member2compass(member, section):
    j = {}

    j['Membership Number'] = member['Membership']
    j["Look-up Title **"] = 'Mr' \
                            if (member['Sex'].lower() == 'm' or
                                member['Sex'].lower() == 'male') else 'Miss'
    j["Forename(s)**"] = member['firstname'].strip()
    j["Surnames **"] = member['lastname'].strip()

    if section in ['Paget', 'Garrick', 'Swinfen']:
        section = 'Beaver Scout'
    elif section in ['Maclean', 'Rowallan', 'Somers']:
        section = 'Cub Scout'
    elif section in ['Boswell', 'Johnson', 'Erasmus']:
        section = 'Scout'
    else:
        log.warn("unkown section type: {}".format(section))
        section = 'UNKNOWN'

    j["Look-up Role **"] = section

    j["Date of Birth **"] = datetime.datetime.strptime(
        member['dob'], '%d/%m/%Y').strftime(date_format)

    addr = parse_addr(member['PrimaryAddress'])
    j["Address Line 1 **"] = addr['street']
    j["Town **"] = addr['town']
    j["Postcode **"] = addr['postcode']
    j["Look-up Postal Country **"] = 'United Kingdom'

    j["Email *"] = member['MumEmail'] \
        if member['MumEmail'].strip() != '' \
        else member['DadEmail']

    j["Start Date for this role**"] = datetime.datetime.strptime(
        member['started'], '%d/%m/%Y').strftime(date_format)

    j["Date they joined Scouts*"] = datetime.datetime.strptime(
        member['joined'], '%d/%m/%Y').strftime(date_format)

    j["Look-up Nationality *"] = "British"
    j["Look-up Ethnicity *"] = "English/Welsh/Scottish/Northern Irish/British"
    j["Look-up Faith/Religion *"] = "No religion"

    p = get_parent(member)
    j["Look-up Parent 1 Title *"] = p['title']
    j["Parent 1 Forename *"] = p['forename']
    j["Parent 1 Surname *"] = p['surname']
    j["Look-up Parent 1 Gender *"] = p['gender']
    j["Look-up Parent 1 Relationship *"] = p['relationship']
    j["Parent 1 Email 1 *"] = p['email']
    j["Parent 1 Telephone 1 *"] = p['tel']

    j["Emergency Contact Forename *"] = j["Parent 1 Forename *"]
    j["Emergency Contact Surname *"] = j["Parent 1 Surname *"]
    j["Emergency Contact Relationship *"] = \
        j["Look-up Parent 1 Relationship *"]
    j["Emergency Contact Telephone Number 1 *"] = j["Parent 1 Telephone 1 *"]
    j["Doctor / Surgery*"] = 'Unknown'
    j["Surgery Telephone 1 *"] = '01543 123456'

    return j


def check(entry, section):
    # Check that all required fields have something in them.
    for k in required_headings:
        if entry[k].strip() == '':
            log.warn("{}: {} {} - emtpy field: {}".format(
                section,
                entry["Forename(s)**"],
                entry["Surnames **"],
                k))


@contextmanager
def display():
    d = Display(visible=1, size=(1000, 1024))

    try:
        d.start()

        yield d
    finally:

        d.stop()
        del d


class Compass:

    def __init__(self, username='', password='', outdir=''):
        self._username = username
        self._password = password
        self._outdir = outdir

        self._browser = None
        self._record = None


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

        self._browser = Browser('chrome') #, profile_preferences=prefs)

        self._browser.visit('https://compass.scouts.org.uk/login/User/Login')

        self._browser.fill('EM', self._username)
        self._browser.fill('PW', self._password)
        time.sleep(1)
        self._browser.find_by_text('Log in').first.click()

        # Look for the Role selection menu and select my Group Admin role.
        self._browser.is_element_present_by_name(
            'ctl00$UserTitleMenu$cboUCRoles',
            wait_time=30)
        self._browser.select('ctl00$UserTitleMenu$cboUCRoles', '1253644')
        time.sleep(1)

    def wait_then_click_xpath(self, xpath, wait_time=30, frame=None):
        frame = self._browser if frame is None else frame
        while True:
            try:
                if frame.is_element_present_by_xpath(xpath, wait_time=wait_time):
                    frame.find_by_xpath(xpath).click()
                    break
                else:
                    log.warning("Timeout expired waiting for {}".format(xpath))
                    time.sleep(1)
            except:
                log.warning("Caught exception: ", exc_info=True)

    def wait_then_click_text(self, text, wait_time=30, frame=None):
        frame = self._browser if frame is None else frame
        while True:
            if frame.is_text_present(text, wait_time=wait_time):
                frame.click_link_by_text(text)
                break
            else:
                log.warning("Timeout expired waiting for {}".format(text))

    def adult_training(self):
        self.home()

        # Navigate to training page a show all records.
        self.wait_then_click_text('Training')
        time.sleep(1)
        self.wait_then_click_text('Adult Training')
        time.sleep(1)
        self.wait_then_click_xpath('//*[@id="bn_p1_search"]')

    def home(self):
        # Click the logo to take us to the top
        self.wait_then_click_xpath('//*[@alt="Compass Logo"]')
        time.sleep(1)

    def search(self):
        self.home()

        # Click search button
        self.wait_then_click_xpath('//*[@id="mn_SB"]')
        time.sleep(1)

        # Click "Find Member(s)"
        self.wait_then_click_xpath('//*[@id="mn_MS"]')
        time.sleep(1)

        # Navigate to training page a show all records.
        with self._browser.get_iframe('popup_iframe') as i:
            self.wait_then_click_xpath('//*[@id="LBTN2"]', frame=i)
            time.sleep(1)
            self.wait_then_click_xpath('//*[@class="popup_footer_right_div"]/a', frame=i)
            time.sleep(1)

    def lookup_member(self, member_number):
        self.home()

        # Click search button
        self.wait_then_click_xpath('//*[@id="mn_SB"]')
        time.sleep(1)

        xpath = '//*[@id="CNLookup2"]'
        while True:
            try:
                if self._browser.is_element_present_by_xpath(xpath, wait_time=30):
                    self._browser.find_by_xpath(xpath).fill(member_number)
                    break
                else:
                    log.warning("Timeout expired waiting for {}".format(xpath))
                    time.sleep(1)
            except:
                log.warning("Caught exception: ", exc_info=True)

        self.wait_then_click_xpath('//*[@id="mn_QS"]')

    def fetch_table(self, table_id):
        parser = etree.HTMLParser()

        def columns(row):
            return ["".join(_.itertext()) for _ in
                    etree.parse(StringIO(row.html), parser).findall('/*/td')]

        def headers(row):
            return ["".join(_.itertext()) for _ in
                    etree.parse(StringIO(row.html), parser).findall('/*/td')]

        headers_xpath = '//*[@id ="{}"]/thead/*'.format(table_id)
        table_xpath = '//*[@id ="{}"]/tbody/tr[not(@style="display: none;")]'.format(table_id)

        if self._browser.is_element_present_by_xpath(table_xpath, wait_time=5):

            headings = [headers(row) for row
                        in self._browser.find_by_xpath(headers_xpath)][0]

            records = [columns(row) for row
                       in self._browser.find_by_xpath(table_xpath)]

            # Extend the length of each row to the same length as the columns
            records = [row+([None] * (len(headings)-len(row))) for row in records]

            # And add dummy columns if we do not have enough headings
            headings = headings + ["dummy{}".format(_) for _ in range(0,len(records[0]) - len(headings))]

            return pd.DataFrame.from_records(records, columns=headings)

        log.warning("Failed to find table {}".format(table_id))
        return None

    def member_training_record(self, member_number, member_name):
        self.lookup_member(member_number)

        # Select Training record
        self.wait_then_click_xpath('//*[@id="LBTN5"]')

        personal_learning_plans = self.fetch_table('tbl_p5_TrainModules')
        personal_learning_plans['member'] = member_number
        personal_learning_plans['name'] = member_name

        training_record = self.fetch_table('tbl_p5_AllTrainModules')
        training_record['member'] = member_number
        training_record['name'] = member_name

        mandatory_learning = self.fetch_table('tbl_p5_TrainOGL')
        mandatory_learning['member'] = member_number
        mandatory_learning['name'] = member_name

        return personal_learning_plans, personal_learning_plans, mandatory_learning

    def member_permits(self, member_number, member_name):
        self.lookup_member(member_number)

        # Select Permits
        self.wait_then_click_xpath('//*[@id="LBTN4"]')

        permits = self.fetch_table('tbl_p4_permits')
        if permits is not None:
            permits['member'] = member_number
            permits['name'] = member_name

        return permits

    @lru_cache()
    def get_all_adult_trainers(self):

        self.adult_training()

        return self.fetch_table('tbl_p1_results')

    @lru_cache()
    def get_all_group_members(self):

        self.search()

        self._browser.is_element_present_by_xpath('//*[@id = "MemberSearch"]/tbody', wait_time=10)
        time.sleep(1)

        # Hack to ensure that all of the search results loaded.
        for i in range(0, 5):
            self._browser.execute_script(
                'document.getElementById("ctl00_main_working_panel_scrollarea").scrollTop = 100000')
            time.sleep(1)

        return self.fetch_table('MemberSearch')

    def export(self, section):
        # Select the My Scouting link.
        self._browser.is_text_present('My Scouting', wait_time=30)
        self._browser.click_link_by_text('My Scouting')

        # Click the "Group Sections" hotspot.
        self.wait_then_click_xpath('//*[@id="TR_HIER7"]/h2')

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
        self.wait_then_click_xpath(
            '//*[@id="TR_HIER7_TBL"]/tbody/tr[{}]/td[4]/a'.format(
                section_map[section.lower()]
            ))

        # Click on the Export button.
        self.wait_then_click_xpath('//*[@id="bnExport"]')

        # Click to say that we want a CSV output.
        self.wait_then_click_xpath(
            '//*[@id="tbl_hdv"]/div/table/tbody/tr[2]/td[2]/input')
        time.sleep(2)

        # Click to say that we want all fields.
        self.wait_then_click_xpath('//*[@id="bnOK"]')

        download_path = os.path.join(self._outdir, 'CompassExport.csv')

        if os.path.exists(download_path):
            log.warn("Removing stale download file.")
            os.remove(download_path)

        # Click the warning.
        self.wait_then_click_xpath('//*[@id="bnAlertOK"]')

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

        def get_section(path, section):
            df = pd.read_csv(path, dtype=object, sep=',')
            df['section'] = section
            df['forenames_l'] = [_.lower().strip() for _ in df['forenames']]
            df['surname_l'] = [_.lower().strip() for _ in df['surname']]
            return df

        self._records = pd.DataFrame().append(
            [get_section(os.path.join(self._outdir, section),
                         os.path.splitext(section)[0])
             for section in os.listdir(self._outdir)], ignore_index=True)

    def find_by_name(self, firstname, lastname, section_wanted=None,
                     ignore_second_name=True):
        """Return list of matching records."""

        recs = self._records
        
        if ignore_second_name:
            df = recs[
                (recs.forenames_l.str.lower().str.match(
                        '^{}.*$'.format(firstname.strip(' ')[0].lower().strip()))) &
                  (recs.surname_l == lastname.lower().strip())]
            
        else:
            df = recs[(recs.forenames_l == firstname.lower().strip()) &
                      (recs.surname_l == lastname.lower().strip())]

        if section_wanted is not None:
            df = df[(df['section'] == section_wanted)]

        return df

    def sections(self):
        "Return a list of the sections for which we have data."
        return self._records['section'].unique()

    def all_yp_members_dict(self):
        return {s: members for s, members in self._records.groupby('section')}

    def section_all_members(self, section):
        return [m for i, m in self._records[
            self._records['section'] == section].iterrows()]

    def section_yp_members_without_leaders(self, section):
        return [m for i, m in self._records[
            (self._records['section'] == section) &
            (self._records['role'].isin(
                ['Beaver Scout', 'Cub Scout', 'Scout']))].iterrows()]

    def members_with_multiple_membership_numbers(self):
        return [member for s, member in self._records.groupby(
            ['forenames', 'surname']).filter(
                lambda x: len(x['membership_number'].unique()) > 1).groupby(
                    ['forenames', 'surname', 'membership_number'])]


def _main(username, password, sections, outdir):

    with display():

        compass = Compass(username, password, outdir)
        compass.loggin()

        try:
            all = compass.get_all_group_members()

            all[['No', 'Name']][0:2].values.tolist()

            # print(compass.member_permits('00857285'))
            # print(compass.member_training_record('00857285'))


            time.sleep(1)
            # for section in sections:
            #    compass.export(section)
        finally:
            compass.quit()

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
