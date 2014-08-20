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
import pprint


import gspread
import creds

from group import Group, OSM_REF_FIELD

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"

MEMBER_SPREADSHEET_NAME = "TestSpread"


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
                 'NOKAddress1': 'NOK Address',
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
                 #'Ethnicity':'Ethnicity'
                 }


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
           'DadsName': 'Dads Name',
           'MumsName': 'Dads Name',
           'Notes': 'Notes',
           'Medical': 'Medical',
           'School': 'School',
           'Hobbies': 'Hobbies',
           'DadDBS': 'Dad DBS',
           'MumDBS': 'Mum DBS',
           'Sex': 'Sex',
           'FathersOccupation': 'Fathers Occupation',
           'MothersOccupation': 'Mothers Occupation',
           #'Datetonextsection': 'Date to next section',
           #'BeaverColony': 'Beaver Colony',
           #'CubPack': 'Cub Pack',
           #'ScoutTroop': 'Scout Troop',
           'FamilyReference': 'Family Reference',
           'PersonalReference': 'Personal Reference',
           #'Ethnicity': 'Ethnicity'
           }

TOP_OFFSET = 3  # Number of rows above the heading row in the gs
HEADER_ROW = TOP_OFFSET + 1
YP_WKS = 'Master'
ADULT_WKS = 'Leaders'


GS_SOURCE_FIELD = 'Source'  # Field used to hold the source of a record.
# Field used to hold the date of the last update.
GS_LAST_UPDATE_FIELD = 'Last Updated'


def format_date(field):
    """GS assumes all dates are in US format. So we have to
    guess whether a field looks like a date and reverse the format.
    """
    try:
        new_field = datetime.datetime.strptime(
            field, '%d/%m/%Y').strftime('%d %B %Y')
    except ValueError:
        new_field = field

    return new_field


# def find_ref_in_sections(reference, section_list):
#     """Return a list of the section names for each section that
#     contains reference.

#     Call with a reference to search for and a list of (section_name,
#     section) pairs."""

#     if len(section_list) == 0:
#         return []
#     section = section_list[0]
#     section_name = section[0]
#     references = [member[OSM_REF_FIELD] for member in section[1]]
#     if reference in references:
#         return [section_name] + find_ref_in_sections(
#             reference, section_list[1:])
#     else:
#         return find_ref_in_sections(reference, section_list[1:])


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
            if format_date(gs_value) != format_date(osm_value):
                log.info("Updating (from {}) - [{}, {}, {}]"
                         "gs value ({!r})[{}] != osm value "
                         "({!r}) setting to ({})  gs: {!r}\n".format(
                             name, reference, osm_field, gs_field,
                             format_date(gs_value), headings.index(gs_field),
                             format_date(osm_value), format_date(osm_value),
                             osm_values[references.index(reference)]))
                wks.update_cell(
                    2 + TOP_OFFSET +
                    references.index(reference), 1 +
                    headings.index(gs_field),
                    format_date(osm_value))
                updated = True

        if updated:
            log.debug("Updated member information OSM record = \n {}".format(
                str(member)))
            log.debug("Updated member information GS record = \n {}".format(
                pprint.pformat(
                    osm_values[references.index(reference)])))

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

    # new_members now contains only new records and
    # deleted_references contains only deleted records

    # append new records
    # find first empty row where there are no rows below it that have
    # any content
    if len(new_members) > 0:
        empty_row = False
        for row in range(1, wks.row_count + 1):
            is_empty = (
                len([cell for cell in wks.row_values(row)
                     if cell is not None]) == 0)
            if is_empty and empty_row is False:
                empty_row = row
            elif not is_empty and empty_row is not False:
                empty_row = False

        # If there are not enough spare row in spreadsheet add extra rows
        start_row = empty_row
        if empty_row is False:
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




def process_adults(group, spread):

    log.info("Processing adults...")
    # Process Adults ##################
    adult_members = group.all_adult_members()

    process_section('Adult', adult_members, spread,
                    ADULT_MAPPING, target_wks=ADULT_WKS)

    # Warn about missing references
    missing_ref_names = ["{} {}".format(member['firstname'],
                                        member['lastname'])
                         for member in group.missing_adult_references()]

    if missing_ref_names:
        log.warn("The following adults do not have references: \n {}".format(
            "\n\t".join(missing_ref_names)))

    # Get list of all reference is OSM
    adult_osm_references = group.all_adult_references()

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
                    if (reference not in adult_osm_references) and
                    reference is not None],
                   wks=ADULT_WKS)


def process_yp(group, spread):
    # Process YP Sections #################
    log.info("Processing YP...")

    # Warn about missing references
    for section in group.YP_SECTIONS:
        # Warn about missing references
        missing_ref_names = ["{} {}".format(member['firstname'],
                                            member['lastname'])
                             for member in
                             group.section_missing_references(section)]

        if missing_ref_names:
            log.warn("The following members in {} "
                     "do not have references: \n {}".format(
                         section,
                         "\n\t".join(missing_ref_names)))

    all_yp_section_leaders = group.all_leaders_in_yp_sections()

    log.info("Check adults in YP sections...")
    # check that all leaders are copied into the Adult section.
    for leader in all_yp_section_leaders:
        if leader[OSM_REF_FIELD] not in group.all_adult_references():
            log.warn("Leader {} is in {!r} section "
                     "but not in the Adult section.".format(
                         leader[OSM_REF_FIELD],
                         group.find_ref_in_sections(
                             leader[OSM_REF_FIELD])))

    # make lists of sections with the Leaders removed
    # TODO: Search for multiple references to the same Family Reference
    #             warn if the details are not the same

    log.info("Remove duplicate records at that are in senior sections ...")
    all_yp_members = group.all_yp_members_without_senior_duplicates_dict()

    # Process the remaining members
    for name, section_members in all_yp_members.items():
        log.info("Processing section {}".format(name))
        process_section(name, section_members, spread, MAPPING)

    log.info("Removing old members...")
    # remove deleted members
    # get list of all references from sections
    all_references = []
    for name, section_members in all_yp_members.items():
        all_references.extend([member[OSM_REF_FIELD]
                              for member in section_members])

    # get list of references on Master wks
    wks = spread.worksheet(YP_WKS)
    headings = wks.row_values(HEADER_ROW)
    gs_references = wks.col_values(
        1 + headings.index(MAPPING[OSM_REF_FIELD]))[1 + TOP_OFFSET:]

    # remove references that appear on Master wks but not in any
    # section
    delete_members(spread,
                   [reference for reference in gs_references
                    if (reference not in all_references) and
                    reference is not None])


def _main(osm, gc, auth):
    #test_section = '15797'

    spread = gc.open(MEMBER_SPREADSHEET_NAME)

    group = Group(osm, auth, MAPPING.keys())

    process_adults(group, spread)

    process_yp(group, spread)

    #process_finance_spreadsheet(gc, group)


if __name__ == '__main__':

    args = docopt(__doc__, version='OSM 2.0')

    if args['--debug']:
        level = logging.DEBUG
    else:
        level = logging.INFO

    logging.basicConfig(level=level)
    log.debug("Debug On\n")

    #try:
    #    osm.Accessor.__cache_load__(open(DEF_CACHE, 'rb'))
    #except:
    #    log.warn("Failed to load cache file\n")

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

    #try:
    #    _main(osm, gc)
    #finally:
    #    osm.Accessor.__cache_save__(open(DEF_CACHE, 'wb'))
    _main(osm, gc, auth)
