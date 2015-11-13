
# coding: utf-8
"""Online Scout Manager Interface - generate report from subs.

Usage:
  subs_report.py [-d] [--term=<term>] [--email=<address>]
         <apiid> <token> <outdir>
  subs_report.py (-h | --help)
  subs_report.py --version


Options:
  <outdir>       Output directory for vcard files.
  -d,--debug     Turn on debug output.
  --email=<email> Send to only this email address.
  --term=<term>  Which OSM term to use [default: current].
  -h,--help      Show this screen.
  --version      Show version.
"""

# Setup the OSM access

# In[1]:

import os.path
import osm
from group import Group
import update
import json
import traceback
import logging
import itertools
import smtplib
from docopt import docopt

from pandas.io.json import json_normalize
import pandas as pd

from email.encoders import encode_base64
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"

FROM = "Richard Taylor <r.taylor@bcs.org.uk>"


def send(to, subject, report_path, fro=FROM):

    for dest in to:
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = fro
        msg['To'] = dest

        fp = open(report_path, 'rb')
        file1 = MIMEBase('application', 'vnd.ms-excel')
        file1.set_payload(fp.read())
        fp.close()
        encode_base64(file1)
        file1.add_header('Content-Disposition',
                         'attachment;filename=output.xlsx')

        msg.attach(file1)

        hostname = 'localhost'

        s = smtplib.SMTP(hostname)

        try:
            s.sendmail(fro, dest, msg.as_string())
        except:
            log.error(msg.as_string(),
                      exc_info=True)

        s.quit()


def get_status(d):
    if not d:
        return "Payment Required?"
    detail = [_ for _ in d if _['latest'] == '1']
    return detail[0]['status']


def fetch_scheme(group, acc, section, scheme, term):

    def set_subs_type(d, group=group):
        try:
            return group.find_by_scoutid(
                d['scoutid'])[0]['customisable_data.cf_subs_type_n_g_d_']
        except:
            print("failed to find sub type for: {} {}".format(
                d['scoutid'],
                traceback.format_exc()))
            return "Unknown"

    schedules = acc("ext/finances/onlinepayments/?action=getPaymentSchedule"
                    "&sectionid={}&schemeid={}&termid={}".format(
                        section['id'], scheme['schemeid'], term))
    status = acc("ext/finances/onlinepayments/?action="
                 "getPaymentStatus&sectionid={}&schemeid={}&termid={}".format(
                     section['id'], scheme['schemeid'], term))

    schedules = [_ for _ in schedules['payments'] if _['archived'] == '0']

    try:
        data = json_normalize(status['items'])
    except:
        return pd.DataFrame()

    for schedule in schedules:
        data[schedule['paymentid']] = data[schedule['paymentid']].apply(
            lambda x: get_status(json.loads(x)['status']))

    data['subs_type'] = data.apply(set_subs_type, axis=1)

    data['section'] = section['name']
    data['scheme'] = (
        "General Subscriptions"
        if scheme['name'].startswith("General Subscriptions")
        else "Discounted Subscriptions")

    for schedule in schedules:
        data.rename(columns={schedule['paymentid']: schedule['name']},
                    inplace=True)

    return data


# In[3]:

def fetch_section(group, acc, section, term):
    schemes = acc(
        "ext/finances/onlinepayments/?action=getSchemes&sectionid={}".format(
            section['id']))

    # filter only General and Discounted Subscriptions
    schemes = [_ for _ in schemes['items'] if (
        _['name'].startswith("General Subscriptions") or
        _['name'].startswith("Discounted Subscriptions"))]

    # Check that we only have two subscriptions remaining. If there is
    # more, the rest of the report is going to barf.
    if len(schemes) > 2:
        log.error("Found more than 2 matching schemes in {}."
                  "Matching schemes were: {}".format(section['name'],
                                                     ",".join(schemes)))

    c = pd.concat([fetch_scheme(group, acc, section, scheme, term)
                   for scheme in schemes
                   if scheme['name'] != 'Camps and Events'],
                  ignore_index=True)
    return c


def _main(osm, auth, outdir, email, term):

    assert os.path.exists(outdir) and os.path.isdir(outdir)

    group = Group(osm, auth, update.MAPPING.keys(), term)

    # Nasty hack to pick up the current term if the user did not
    # pass in a specific term.
    actual_term = list(group._sections.sections.values())[0].term['termid']
    acc = group._sections._accessor

    sections = [
        {'name': 'Paget', 'id': '9960'},
        {'name': 'Swinfen', 'id': '17326'},
        {'name': 'Maclean', 'id': '14324'},
        {'name': 'Rowallan', 'id': '12700'},
        {'name': 'Johnson', 'id': '5882'},
        {'name': 'Garrick', 'id': '20711'},
        {'name': 'Erasmus', 'id': '20707'},
        {'name': 'Somers', 'id': '20706'},
        {'name': 'Boswell', 'id': '10363'}
    ]

    subs_names = ['General Subscriptions', 'Discounted Subscriptions']
    subs_types = ['G', 'D']
    subs_names_and_types = list(zip(subs_names, subs_types))
    all_types = subs_types + ['N', ]

    al = pd.concat([fetch_section(group, acc, section, actual_term)
                    for section in sections], ignore_index=True)

    # al[(al['scheme'] == 'Discounted Subscriptions') & (
    #    al['subs_type'] == 'D')].dropna(axis=1, how='all')

    # find all members that do not have at least one subscription to either
    # 'Discounted Subscriptions' or 'General Subscriptions'
    # filtered by those that have a 'N' in their subscription type.
    #

    # all_yp_members = group.all_yp_members_without_leaders()

    all = [[[_['member_id'],
             _['first_name'],
             _['last_name'],
             _['customisable_data.cf_subs_type_n_g_d_'],
             section] for _ in
            group.section_yp_members_without_leaders(section)]
           for section in group.YP_SECTIONS]

    all_in_one = list(itertools.chain.from_iterable(all))

    all_members_df = pd.DataFrame(all_in_one, columns=(
        'scoutid', 'firstname', 'lastname', 'subs_type', 'section'))

    al_only_subs = al[al['scheme'].isin(subs_names)]
    # only those that are paying more than one subscription.
    members_paying_multiple_subs = al_only_subs[
        al_only_subs.duplicated('scoutid', take_last=True) |
        al_only_subs.duplicated('scoutid')]

    # In[11]:

    # Not used ?
    # gen = al[al['scheme'] == 'General Subscriptions'].dropna(axis=1, how='all')
    # all_gen_members = all_members_df[all_members_df['subs_type'] == 'G']
    # all_gen_members['scoutid'] = all_gen_members['scoutid'].astype(str)
    # all_gen_members[~all_gen_members['scoutid'].isin(gen['scoutid'].values)]

    # In[12]:

    out_path = os.path.join(outdir, 'output.xlsx')

    with pd.ExcelWriter(out_path,
                        engine='xlsxwriter') as writer:

        # Status of all subs.
        for scheme in subs_names:
            al[al['scheme'] == scheme].dropna(
                axis=1, how='all').to_excel(writer, scheme)

        # All subs with the correct subs_type
        for scheme in subs_names_and_types:
            al[(al['scheme'] == scheme[0]) &
                (al['subs_type'] == scheme[1])].dropna(
                axis=1, how='all').to_excel(writer, scheme[0] + "_OK")

        # All subs with the wrong subs type
        for scheme in subs_names_and_types:
            al[(al['scheme'] == scheme[0]) &
                (al['subs_type'] != scheme[1])].dropna(
                axis=1, how='all').to_excel(writer, scheme[0] + "_BAD")

        # Members not in the subs that their sub_type says they should be.
        for scheme in subs_names_and_types:
            gen = al[al['scheme'] == scheme[0]].dropna(axis=1, how='all')
            all_gen_members = all_members_df[
                all_members_df['subs_type'] == scheme[1]]
            all_gen_members['scoutid'] = all_gen_members['scoutid'].astype(str)
            all_gen_members[~all_gen_members['scoutid'].isin(
                gen['scoutid'].values)].to_excel(writer, "Not in " + scheme[0])

        # All YP members without their subs_type set to anything.
        all_members_df[~all_members_df['subs_type'].isin(
            all_types)].to_excel(writer, "Unknown Subs Type")

        # Members paying multiple subs
        members_paying_multiple_subs.dropna(
            axis=1, how='all').to_excel(writer, "Multiple  payers")

    if email:
        send([email, ], "OSM Subs Report", out_path)


if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.WARN

    if args['--term'] in [None, 'current']:
        args['--term'] = None

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    _main(osm, auth,
          args['<outdir>'], args['--email'], args['--term'])
