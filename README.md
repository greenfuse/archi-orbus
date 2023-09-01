# archi-orbus
Python script to create Orbus uploadable xlsx from archimate Model Exchange file. To enable modelled objects and relationships to be uploaded into OrbusInfinity.
 
Requires the python openpyxl module which can be installed via pip.

Run modelexchange_orbus.py with python. Raises a prompt for the selection of an xml file. This needs to be a Model Exchange file such as an export from Archi. Creates an excel file in the same folder as the xml file. This new file has the same name but with the .xlsx extension. The excel file can be selected for upload from OrbusInfinity.

All efforts have been made to ensure that this works for me but if you want to try it, use this at your own risk.
