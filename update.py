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
import datetime
from docopt import docopt
import osm

import gspread
import creds

log = logging.getLogger(__name__)

DEF_CACHE = "osm.cache"
DEF_CREDS = "osm.creds"

adult_mapping = {'firstname':'Firstname',
                 'lastname':'Lastname',
                 'joined':'Joined',
                 'started':'Started section',
                 'HomeTel':'Home Tel',
                 'PersonalMob':'Personal Mob',
                 'NOKMob1':'NOK Mob1',
                 'NOKMob2':'NOK Mob2',
                 'PersonalEmail':'Personal Email',
                 'NOKEmail1':'NOK Email1',
                 'NOKEmail2':'NOK Email2',
                 'dob':'Date of birth',
                 'PrimaryAddress':'Primary Address',
                 'NOKAddress':'NOK Address',
                 #'Parents':'Parents',
                 'Notes':'Notes',
                 'Medical':'Medical',
                 'PlaceofWork':'Place of Work',
                 'Hobbies':'Hobbies',
                 'GiftAid':'Gift Aid',
                 #'FathersOccupation':'Fathers Occupation',
                 #'MothersOccupation':'Mothers Occupation',
                 #'Datetonextsection':'Date to next section',
                 #'BeaverColony':'Beaver Colony',
                 #'CubPack':'Cub Pack',
                 #'ScoutTroop':'Scout Troop',
                 'FamilyReference':'Family Reference',
                 'PersonalReference':'Personal Reference'}


mapping = {'firstname':'Firstname',
           'lastname':'Lastname',
           'joined':'Joined',
           'started':'Started section',
           'HomeTel':'Home Tel',
           'PersonalMob':'Personal Mob',
           'DadMob':'Dad Mob',
           'MumMob':'Mum Mob',
           'PersonalEmail':'Personal Email',
           'DadEmail':'Dad Email',
           'MumEmail':'Mum Email',
           'dob':'Date of birth',
           'PrimaryAddress':'Primary Address',
           'SecondaryAddress':'Secondary Address',
           'Parents':'Parents',
           'Notes':'Notes',
           'Medical':'Medical',
           'School':'School',
           'Hobbies':'Hobbies',
           'GiftAid':'Gift Aid',
           'FathersOccupation':'Fathers Occupation',
           'MothersOccupation':'Mothers Occupation',
           'Datetonextsection':'Date to next section',
           'BeaverColony':'Beaver Colony',
           'CubPack':'Cub Pack',
           'ScoutTroop':'Scout Troop',
           'FamilyReference':'Family Reference',
           'PersonalReference':'Personal Reference'}

yp_wks = 'Master'
adult_wks = 'Leaders'

ref_field='PersonalReference'

def format_date(field):
  """GS assumes all dates are in US format. So we have to 
  guess whether a field looks like a date and reverse the format.
  """
  try:
    new_field = datetime.datetime.strptime(field,'%d/%m/%Y').strftime('%m/%d/%Y')
  except ValueError:
    new_field = field

  return new_field

def process_section(name, section, spread, mapping, filter_func=lambda x:  True, target_wks=yp_wks):
  wks = spread.worksheet(target_wks)

  members = [ member for member in list(section.members.values()) \
              if filter_func(member) ]

  #   Search for duplicate YP references

  #   Search for multiple references to the same Family Reference
  #     warn if the details are not the same

  # connect to spreadsheet

  headings = wks.row_values(1)
  print(headings)

  # Fetch the list of references that are already in the gs
  references = wks.col_values(1+headings.index(mapping[ref_field]))[1:]

  def update_record(reference, member):
      for osm_field,gs_field in mapping.items():
        gs_value = wks.cell(2+references.index(reference), 1+headings.index(gs_field)).value
        osm_value = member[osm_field]
        if osm_value == '':
           osm_value = None # gs returns None for '' so this make the comparison work.
        if gs_value != osm_value:
          print("Updating [{}, {}, {}] gs value ({}) != osm value ({}) setting to ({})\n".format(
            reference, osm_field, gs_field,
            wks.cell(2+references.index(reference), 1+headings.index(gs_field)).value, 
            member[osm_field],format_date(member[osm_field])))
          wks.update_cell(2+references.index(reference), 1+headings.index(gs_field), format_date(member[osm_field]))

  updated_members = []
  updated_references = []

  # Update any records that are in both the gs and osm.
  for member in members:
    if member[ref_field] in references:
      update_record(member[ref_field], member)
      updated_members.append(member)
      updated_references.append(member[ref_field])

  new_members = [ member for member in members if member not in updated_members ]
  #deleted_references = [ ref for ref in references if (ref not in updated_references) and ref != None ]

  # new_members now contains only new records and 
  # deleted_references contains only deleted records

  # append new records
  # find first empty row where there are no rows below it that have
  # any content
  empty_row = False
  for row in range(1,wks.row_count+1):
    is_empty = (len([ cell for cell in wks.row_values(row) if cell != None ]) == 0)
    if is_empty and empty_row == False:
      empty_row = row
    elif not is_empty and empty_row != False:
      empty_row = False

  # If there are not enough spare row in spreadsheet add extra rows
  start_row = empty_row
  if empty_row == False:
    start_row = wks.row_count+1
    wks.add_rows(len(members))
  elif (wks.row_count - empty_row) < len(new_members):
    wks.add_rows(len(members) - (wks.row_count - empty_row))


  # Insert the new records
  row = start_row
  for member in new_members:
    for osm_field,gs_field in mapping.items():
      wks.update_cell(row, 1+headings.index(gs_field), format_date(member[osm_field]))
    row += 1

def delete_members(spread, references, wks=yp_wks):
  wks = spread.worksheet(wks)
  headings = wks.row_values(1)

  # Handle deleted records.
  # These are moved to a special 'Deleted' worksheet
  try:
    del_wks = spread.worksheet('Deleted')
  except gspread.WorksheetNotFound:
    del_wks = spread.add_worksheet('Deleted',1, wks.col_count)
    header=wks.row_values(1)
    for col in range(1,len(header)+1):
      del_wks.update_cell(1,col,header[col-1])

  def move_row(reference, from_wks, to_wks):
    # find reference in from_wks
    current_references = from_wks.col_values(1+headings.index(mapping[ref_field]))[1:]
    from_row = current_references.index(reference)+2 

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
    for col in range(1,len(from_row_values)+1):
      from_wks.update_cell(from_row,col,'')


  print ("Going to delete {!r}\n".format(references))

  for reference in references:
    move_row(reference, wks, del_wks)

if __name__ == '__main__':

    logging.basicConfig(level=logging.INFO)
    log.debug("Debug On\n")

    try:
        osm.Accessor.__cache_load__(open(DEF_CACHE, 'rb'))
    except:
        log.warn("Failed to load cache file\n")

    args = docopt(__doc__, version='OSM 2.0')

    if args['-a']:
        auth = osm.Authorisor(args['<apiid>'], args['<token>'])
        auth.authorise(args['<email>'],
                       args['<password>'])
        auth.save_to_file(open(DEF_CREDS, 'w'))
        sys.exit(0)

    auth = osm.Authorisor(args['<apiid>'], args['<token>'])
    auth.load_from_file(open(DEF_CREDS, 'r'))

    # creds needs to contain a tuple of the following form
    #   creds = ('username','password')
    gc = gspread.login(*creds.creds)
    spread = gc.open("TestSpread")
  
    # read in all of the personal details from each of the OSM sections

    sections = osm.OSM(auth)

    #test_section = '15797'
    adult = '18305'
    yp_sections = {'maclean': '14324',
                   'boswell': '10363'}

    print(sections.sections)

    process_section('Adult',sections.sections[adult],spread,adult_mapping,target_wks=adult_wks)

    ######  Process YP Sections #################

    # used to filter Leaders from the YP sections
    def filter_func(member):
      return not member['patrol'] == 'Leaders'

    for name, section in yp_sections.items():
      process_section(name,sections.sections[section],spread,mapping,filter_func)

    # remove deleted members
    # get list of all references from sections
    all_references = []
    for name, section in yp_sections.items():
      all_references.extend([ member[ref_field] for member in sections.sections[section].members.values() \
                              if filter_func(member) ])

    # get list of references on Master wks
    wks = spread.worksheet(yp_wks)
    headings = wks.row_values(1)
    gs_references = wks.col_values(1+headings.index(mapping[ref_field]))[1:]

    # remove references that appear on Master wks but not in any
    # section
    delete_members(spread,
                   [ reference for reference in gs_references if reference not in all_references ] )
    
    osm.Accessor.__cache_save__(open(DEF_CACHE, 'wb'))
