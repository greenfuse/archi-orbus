# archi-orbus
Python script to create Orbus uploadable xlsx from archimate Model Exchange file (such as xml exported from Archi). This is to enable modelled objects and relationships to be uploaded into OrbusInfinity.
 
Requires the python openpyxl module which can be installed via pip.

For OrbusInfinity with the 'Core EA & Archimate' 3.1 solution

Run modelexchange_orbus.py with python. Raises a prompt for the selection of an xml file. This needs to be a Model Exchange file such as an xml export from Archi. Creates an excel file in the same folder as the xml file. This new file has the same name but with the .xlsx extension. The excel file can be selected for upload from OrbusInfinity.

 The new spreadsheet includes: 
  - Worksheets for Objects and Relationships.
  - Object sheet has columns with element properties. If the column header matches an object attribute, it will be uploaded in Orbus.
  - Hidden elements and relationships that are not in a view (soft deleted). If you don't want these, delete from Archi before exporting the xml.
 
Not supported:
  - Relationship junctions (And/Or), these are not used in Orbus out of the box.
  - Container elements. These are not used in Orbus out of the box.

All efforts have been made to ensure that this works for me but if you want to try it, use this at your own risk.
