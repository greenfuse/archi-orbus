# archi-orbus
Python script to create Orbus uploadable xlsx from archimate Model Exchange file. To enable modelled objects and relationships to be uploaded into OrbusInfinity.
 
Requires the python openpyxl module which can be installed via pip.

Run with python. Raises a prompt for the selection of an xml file. This needs to be a Model Exchange file such as an export from Archi. Creates an excel file in the same folder as the xml file. This new file has the same name but with the .xlsx extension. The excel file can be selected for upload from OrbusInfinity.
