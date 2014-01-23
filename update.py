# coding=utf-8
"""Online Scout Manager Interface.

Usage:
  update.py [-d] <apiid> <token> 
  update.py [-d] <apiid> <token> -a <email> <password>
  update.py (-h | --help)
  update.py --version


Options:
  -d,--debug     Turn on debug output.
  -h,--help      Show this screen.
  --version      Show version.
  -a             Request authorisation credentials.

"""

import sys
import logging
import datetime
from docopt import docopt
import osm

import gspread
import creds

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"

MEMBER_SPREADSHEET_NAME = "TestSpread"
FINANCE_SPREADSHEET_NAME = "Copy of Simplified Membership Master List (Apr 2013 - Mar 2014)"

SECTIONIDS = {'Adult': '18305',
              'Paget': '9960',
              'Brown': '17326',
              'Maclean': '14324',
              'Rowallan': '12700',
              'Boswell': '10363',
              'Johnson': '5882'}
              #'Waiting List': ""}

ADULT_MAPPING = {'firstname': 'Firstname',
                 'lastname': 'Lastname',
                 'joined': 'Joined',
                 'started': 'Started section',
                 'HomeTel': 'Home Tel',
                 'PersonalMob': 'Personal Mob',
                 'NOKMob1': 'NOK Mob1',
                 'NOKMob2': 'NOK Mob2',
                 'PersonalEmail': 'Personal Email',
                 'NOKEmail1': 'NOK Email1',
                 'NOKEmail2': 'NOK Email2',
                 'dob': 'Date of birth',
                 'PrimaryAddress': 'Primary Address',
                 'NOKAddress': 'NOK Address',
                 #'Parents': 'Parents',
                 'Notes': 'Notes',
                 'Medical': 'Medical',
                 'PlaceofWork': 'Place of Work',
                 'Hobbies': 'Hobbies',
                 'GiftAid': 'Gift Aid',
                 #'FathersOccupation':'Fathers Occupation',
                 #'MothersOccupation':'Mothers Occupation',
                 #'Datetonextsection':'Date to next section',
                 #'BeaverColony':'Beaver Colony',
                 #'CubPack':'Cub Pack',
                 #'ScoutTroop':'Scout Troop',
                 'FamilyReference': 'Family Reference',
                 'PersonalReference': 'Personal Reference',
                 'Ethnicity':'Ethnicity'}

# Finance spreadsheet
FIN_HEADER_ROW = 1
FIN_MAPPING_DETAILS = {'PersonalReference': 'SO UID'}

MAPPING = {'firstname': 'Firstname',
           'lastname': 'Lastname',
           'joined': 'Joined',
           'started': 'Started section',
           'HomeTel': 'Home Tel',
           'PersonalMob': 'Personal Mob',
           'DadMob': 'Dad Mob',
           'MumMob': 'Mum Mob',
           'PersonalEmail': 'Personal Email',
           'DadEmail': 'Dad Email',
           'MumEmail': 'Mum Email',
           'dob': 'Date of birth',
           'PrimaryAddress': 'Primary Address',
           'SecondaryAddress': 'Secondary Address',
           'Parents': 'Parents',
           'Notes': 'Notes',
           'Medical': 'Medical',
           'School': 'School',
           'Hobbies': 'Hobbies',
           'GiftAid': 'Gift Aid',
           'Sex': 'Sex',
           'FathersOccupation': 'Fathers Occupation',
           'MothersOccupation': 'Mothers Occupation',
           'Datetonextsection': 'Date to next section',
           'BeaverColony': 'Beaver Colony',
           'CubPack': 'Cub Pack',
           'ScoutTroop': 'Scout Troop',
           'FamilyReference': 'Family Reference',
           'PersonalReference': 'Personal Reference',
           'Ethnicity': 'Ethnicity'}

TOP_OFFSET = 3  # Number of rows above the heading row in the gs
HEADER_ROW = TOP_OFFSET + 1
YP_WKS = 'Master'
ADULT_WKS = 'Leaders'

DETAIL_WKS = "Sections In Detail"

OSM_REF_FIELD = 'PersonalReference'

GS_SOURCE_FIELD = 'Source'  # Field used to hold the source of a record.
# Field used to hold the date of the last update.
GS_LAST_UPDATE_FIELD = 'Last Updated'


def format_date(field):
    """GS assumes all dates are in US format. So we have to
    guess whether a field looks like a date and reverse the format.
    """
    try:
        new_field = datetime.datetime.strptime(
            field, '%d/%m/%Y').strftime('%m/%d/%Y')
    except ValueError:
        new_field = field

    return new_field


def find_ref_in_sections(reference, section_list):
    """Return a list of the section names for each section that
    contains reference.

    Call with a reference to search for and a list of (section_name,
    section) pairs."""

    if len(section_list) == 0:
        return []
    section = section_list[0]
    section_name = section[0]
    references = [member[OSM_REF_FIELD] for member in section[1]]
    if reference in references:
        return [section_name] + find_ref_in_sections(
            reference, section_list[1:])
    else:
        return find_ref_in_sections(reference, section_list[1:])


def process_section(name, section_members, spread, mapping, target_wks=YP_WKS):
    wks = spread.worksheet(target_wks)

    members = section_members

    # connect to spreadsheet
    headings = wks.row_values(HEADER_ROW)

    # Fetch the list of references that are already in the gs
    references = wks.col_values(
        1 + headings.index(mapping[OSM_REF_FIELD]))[1 + TOP_OFFSET:]

    def update_source(reference, row=None):
        """Update the source field in GS to record which section was the
        source the last update.

        """

        if row is None:
            row = 2 + TOP_OFFSET + references.index(reference)

        wks.update_cell(row, 1 + headings.index(GS_SOURCE_FIELD), name)
        wks.update_cell(row, 1 + headings.index(GS_LAST_UPDATE_FIELD),
                        datetime.datetime.today().strftime('%m/%d/%Y'))

    def update_record(osm_values, reference, member):
        updated = False
        for osm_field, gs_field in mapping.items():
            gs_value = osm_values[
                references.index(reference)][headings.index(gs_field)]
            osm_value = member[osm_field]
            # if osm_value == '':
            # osm_value = None # gs returns None for '' so this make the
            # comparison work.
            if gs_value != osm_value:
                log.info("Updating (from {}) - [{}, {}, {}] gs value ({!r}) != osm value ({!r}) setting to ({})  gs: {!r}\n".format(
                    name, reference, osm_field, gs_field,
                    gs_value,
                    osm_value, format_date(osm_value),
                    osm_values[references.index(reference)]))
                wks.update_cell(
                    2 + TOP_OFFSET +
                    references.index(reference), 1 +
                    headings.index(gs_field),
                    format_date(osm_value))
                updated = True

        if updated:
            # If any of the field have been updated we want to change the
            # Source column to note where from and when.
            update_source(reference)

    updated_members = []
    updated_references = []

    # Update any records that are in both the gs and osm.
    osm_values = wks.get_all_values()[TOP_OFFSET + 1:]
    for member in members:
        if member[OSM_REF_FIELD] in references:
            update_record(osm_values, member[OSM_REF_FIELD], member)
            updated_members.append(member)
            updated_references.append(member[OSM_REF_FIELD])

    new_members = [
        member for member in members if member not in updated_members]
    #deleted_references = [ ref for ref in references if (ref not in updated_references) and ref != None ]

    # new_members now contains only new records and
    # deleted_references contains only deleted records

    # append new records
    # find first empty row where there are no rows below it that have
    # any content
    if len(new_members) > 0:
        empty_row = False
        for row in range(1, wks.row_count + 1):
            is_empty = (
                len([cell for cell in wks.row_values(row) if cell != None]) == 0)
            if is_empty and empty_row == False:
                empty_row = row
            elif not is_empty and empty_row != False:
                empty_row = False

        # If there are not enough spare row in spreadsheet add extra rows
        start_row = empty_row
        if empty_row == False:
            start_row = wks.row_count + 1
            wks.add_rows(len(members))
        elif (wks.row_count - empty_row) < len(new_members):
            wks.add_rows(len(members) - (wks.row_count - empty_row))

        # Insert the new records
        row = start_row
        for member in new_members:
            for osm_field, gs_field in mapping.items():
                wks.update_cell(row, 1 + headings.index(gs_field),
                                format_date(member[osm_field]))

            update_source(member[OSM_REF_FIELD], row=row)

            row += 1


def delete_members(spread, references, wks=YP_WKS):
    wks = spread.worksheet(wks)
    headings = wks.row_values(HEADER_ROW)

    # Handle deleted records.
    # These are moved to a special 'Deleted' worksheet
    try:
        del_wks = spread.worksheet('Deleted')
    except gspread.WorksheetNotFound:
        del_wks = spread.add_worksheet('Deleted', 1, wks.col_count)
        header = wks.row_values(HEADER_ROW)
        for col in range(1, len(header) + 1):
            del_wks.update_cell(TOP_OFFSET, col, header[col - 1])

    def move_row(reference, from_wks, to_wks):
        # find reference in from_wks
        current_references = from_wks.col_values(
            1 + headings.index(MAPPING[OSM_REF_FIELD]))[1 + TOP_OFFSET:]
        from_row = current_references.index(reference) + 1 + TOP_OFFSET + 1

        # add row to to_wks
        from_row_values = from_wks.row_values(from_row)

        # Tidy up the nasty habbit of gs putting 'None' as a string
        values = []
        for item in from_row_values:
            if item is None or item == 'None':
                item = ''
            values.append(item)

        to_wks.append_row(values)

        # remove row from from_wks
        for col in range(1, len(from_row_values) + 1):
            from_wks.update_cell(from_row, col, '')

    log.info("Going to delete {!r}\n".format(references))

    for reference in references:
        move_row(reference, wks, del_wks)

def process_finance_spreadsheet(gc, all_yp):
    import pprint

    fin = gc.open(FINANCE_SPREADSHEET_NAME)

    # Fetch list of personal references from the finance spreadsheet
    # along with the current "Q4" section.
    wks = fin.worksheet(DETAIL_WKS)
    headings = wks.row_values(FIN_HEADER_ROW)
    fin_references = wks.col_values(
        1 + headings.index(
            FIN_MAPPING_DETAILS[OSM_REF_FIELD]))[FIN_HEADER_ROW:]

    # TODO: Parameterise the selection of the current quarter.
    q4_section = wks.col_values(
        1 + headings.index('Q4 Sec'))[FIN_HEADER_ROW:]

    # create a map from refs to members for later lookup
    all_members = {}
    for name, section_members in all_yp.items():
        for member in section_members:
            all_members[member[OSM_REF_FIELD]] = member

    # Create a list of all YP that are not on the finance list.
    # get list of all references from sections
    new_members = []
    for name, section_members in all_yp.items():
        for member in section_members:
            if member[OSM_REF_FIELD] not in fin_references:
                member['SeniorSection'] = name
                new_members.append(member)

    log.info("New members")
    headings = ['patrol', 'SeniorSection', 'PersonalReference', 'firstname', 'lastname',
                'PersonalEmail', 'DadEmail', 'MumEmail',
                'dob', 'joined', 'started']
    with open("new.csv", 'w') as f:
        f.write("\t".join(headings))
        f.write("\n")
        for member in new_members:
            for k in headings:
                f.write(str(member[k])+'\t')
            f.write("\n")


    # Create a list of all YP that are on the finance list but are not
    # in OSM.
    all_osm_references = []
    for name, section_members in all_yp.items():
        all_osm_references.extend([member[OSM_REF_FIELD]
                                   for member in section_members])

    missing_references = []
    for ref in fin_references:
        if ref not in all_osm_references:
            missing_references.append(ref)

    log.info("Old members")
    open("missing.csv", 'w').writelines([ l+'\n' for l in missing_references ])

    # Create a list of all YP who are on the finanace list but are not
    # in the same section in OSM.
    section_map = {'Maclean': 'MP',
                   'Rowallan': 'RP',
                   'Brown': 'BC',
                   'Boswell': 'BT',
                   'Johnson': 'JT',
                   'Paget': 'PC'}

    changed_members = []
    for name, section_members in all_yp.items():
        for member in section_members:
            if member[OSM_REF_FIELD] in fin_references \
               and section_map[name] != q4_section[fin_references.index(
                   member[OSM_REF_FIELD])]:
                changed_members.append((member,
                                        q4_section[fin_references.index(
                                            member[OSM_REF_FIELD])],
                                        section_map[name]))

    with open("changed.csv", 'w') as f:
        f.write("Personal Reference\tOld\tNew\tFirst\tLast\n")
        for member in changed_members:
            f.write("\t".join([
                member[0][OSM_REF_FIELD],
                member[1],
                member[2],
                all_members[member[0][OSM_REF_FIELD]]['firstname'],
                all_members[member[0][OSM_REF_FIELD]]['lastname']]))
            f.write("\n")

def _main(osm, gc):
    spread = gc.open(MEMBER_SPREADSHEET_NAME)

    sections = osm.OSM(auth, SECTIONIDS.values())

    #test_section = '15797'

    def all_members(section_id):
        return sections.sections[section_id].members.values()

    log.info("Processing adults...")
    # Process Adults ##################
    adult_members = all_members(SECTIONIDS['Adult'])

    process_section('Adult', adult_members, spread,
                    ADULT_MAPPING, target_wks=ADULT_WKS)

    # Get list of all reference is OSM
    osm_references = [member[OSM_REF_FIELD] for member in adult_members]

    # get list of references on Leader wks
    wks = spread.worksheet(ADULT_WKS)
    headings = wks.row_values(HEADER_ROW)
    gs_references = wks.col_values(
        1 + headings.index(ADULT_MAPPING[OSM_REF_FIELD]))[TOP_OFFSET + 1:]

    log.info("Deleting old leaders...")
    # remove references that appear on Leaders wks but not in any
    # section
    delete_members(spread,
                   [reference for reference in gs_references
                    if (reference not in osm_references) and reference != None],
                   wks=ADULT_WKS)

    # Process YP Sections #################
    log.info("Processing YP...")

    # read in all of the personal details from each of the OSM sections
    all_members = {'Paget': all_members(SECTIONIDS['Paget']),
                   'Brown': all_members(SECTIONIDS['Brown']),
                   'Maclean': all_members(SECTIONIDS['Maclean']),
                   'Rowallan': all_members(SECTIONIDS['Rowallan']),
                   'Boswell': all_members(SECTIONIDS['Boswell']),
                   'Johnson': all_members(SECTIONIDS['Johnson'])}

    # Make a list of all the leaders
    all_leaders = []
    for section in all_members.keys():
        all_leaders.extend([member for member in all_members[section]
                            if member['patrol'].lower() in
                            ['leaders', 'young leaders']])

    log.info("Check adults in YP sections...")
    # check that all leaders are copied into the Adult section.
    for leader in all_leaders:
        if leader[OSM_REF_FIELD] not in osm_references:
            log.warn("Leader {} is in {!r} section but not in the Adult section.".format(
                leader[OSM_REF_FIELD],
                find_ref_in_sections(
                    leader[OSM_REF_FIELD],
                    [(key, value) for (key, value) in all_members.items()] + [('Adult', adult_members)])))

    # make lists of sections with the Leaders removed
    all_yp_members = {}
    for section in all_members.keys():
        all_yp_members[section] = [member for member in all_members[section]
                                   if not member['patrol'].lower() in
                                   ['leaders', 'young leaders']]

    beaver_members = all_yp_members['Paget'] + all_yp_members['Brown']
    cub_members = all_yp_members['Rowallan'] + all_yp_members['Maclean']
    scout_members = all_yp_members['Johnson'] + all_yp_members['Boswell']

    # TODO: Search for multiple references to the same Family Reference
    #             warn if the details are not the same
    # For each section we need to look at whether a member appears in a
    # senior section too (they will if they are in the process of
    # moving). If they are in a senior section we want to favour the
    # senior records (but warn if it is different).
    def remove_senior_duplicates(section, senior_members):
        kept_members = []
        for member in all_yp_members[section]:
            matching_senior_members = [senior_member for senior_member in senior_members
                                       if senior_member[OSM_REF_FIELD] == member[OSM_REF_FIELD]]
            if len(matching_senior_members) > 0:
                log.info("{} section: {} is in senior section - favouring senior record".format(
                    section, member[OSM_REF_FIELD]))
                # check whether all of the field are the same.
                for senior_member in matching_senior_members:
                    for field in MAPPING.keys():
                        if field == 'joined':
                            # We expect the joined field to be different.
                            continue
                        if member[field] != senior_member[field]:
                            log.warn('{} section: {} senior record field mismatch ({}) "{}" != "{}"'.format(
                                section, member[OSM_REF_FIELD],
                                field, member[field], senior_member[field]))
            else:
                kept_members.append(member)
        return kept_members

    log.info("Remove duplicate records at that are in senior sections ...")
    for beaver_section in ['Paget', 'Brown']:
        all_yp_members[beaver_section] = remove_senior_duplicates(
            beaver_section, cub_members)

    for cub_section in ['Maclean', 'Rowallan']:
        all_yp_members[cub_section] = remove_senior_duplicates(
            cub_section, scout_members)

    # # Process the remaining members
    # for name, section_members in all_yp_members.items():
    #     log.info("Processing section {}".format(name))
    #     process_section(name, section_members, spread, MAPPING)

    log.info("Removing old members...")
    # # remove deleted members
    # # get list of all references from sections
    # all_references = []
    # for name, section_members in all_yp_members.items():
    #     all_references.extend([member[OSM_REF_FIELD]
    #                           for member in section_members])

    # # get list of references on Master wks
    # wks = spread.worksheet(YP_WKS)
    # headings = wks.row_values(HEADER_ROW)
    # gs_references = wks.col_values(
    #     1 + headings.index(MAPPING[OSM_REF_FIELD]))[1 + TOP_OFFSET:]

    # # remove references that appear on Master wks but not in any
    # # section
    # delete_members(spread,
    #                [reference for reference in gs_references
    #                 if (reference not in all_references) and reference != None])

    log.info("Processing finance spreadsheet...")
    process_finance_spreadsheet(gc, all_yp_members)

if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.INFO

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    try:
        osm.Accessor.__cache_load__(open(DEF_CACHE, 'rb'))
    except:
        log.warn("Failed to load cache file\n")

    if args['-a']:
        auth = osm.Authorisor(args['<apiid>'], args['<token>'])
        auth.authorise(args['<email>'],
                       args['<password>'])
        auth.save_to_file(open(DEF_CREDS, 'w'))
        sys.exit(0)

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    # creds needs to contain a tuple of the following form
    #     creds = ('username','password')
    gc = gspread.login(*creds.creds)

    try:
        _main(osm, gc)
    finally:
        osm.Accessor.__cache_save__(open(DEF_CACHE, 'wb'))
