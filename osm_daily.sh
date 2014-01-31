#!/bin/bash


cd /home/rjt/scout-records

source py/bin/activate

ipython3 ./export_vcards.py 95 b25a4455e2171fec9b34748cc6cb4707 vcards Boswell Maclean

ipython3 ./sync_contacts.py vcards Maclean Boswell

ipython3 ./update.py 95 b25a4455e2171fec9b34748cc6cb4707

