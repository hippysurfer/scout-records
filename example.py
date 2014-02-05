# coding: utf-8
import gspread
import creds

# creds needs to contain a tuple of the following form
#   creds = ('username','password')


gc = gspread.login(*creds.creds)
wks = gc.open("TestSpread").sheet1
wks.update_acell("B2","wow")
wks.range('A1:B4')

