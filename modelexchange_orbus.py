import xml
import re
import os
from tkinter import filedialog

import openpyxl

homedir = os.path.expanduser('~')
filetypes = (
    ('xml files', '*.xml'),
    )
selectfile = filedialog.askopenfile(
    title="Select the Model Exchange file",
    initialdir=homedir,
    filetypes=filetypes
    )

if not selectfile:
    print("Nothing selected. Bye bye")
    exit()
    
filepath = selectfile.name
dirname, basename = os.path.split(filepath)
filename, filetype = os.path.splitext(basename)
destination_filepath = os.path.join(dirname, (filename + '.xlsx'))

# get the namespace details
iterparse =  xml.etree.ElementTree.iterparse(filepath, ('start', 'start-ns'))
events = "start", "start-ns"
ns = {}
for event, elem in xml.etree.ElementTree.iterparse(filepath, events):
    if event == "start-ns":
        if elem[0] == "":
            ns['xmlns'] = elem[1]
        else:
            ns[elem[0]] = elem[1]
        
xmlns = ns['xmlns']
xsi = '{' +  ns['xsi'] + "}"

# ensure the xml has the correct Model Exchange namespace
if not 'www.opengroup.org/xsd/archimate' in xmlns:
    print("This is not a Model Exchange file")
    exit()

# OrbusInfinity does not use junctions
lst_exceptions = ["AndJunction", "OrJunction"]
# Create lists with headers 
lst_elements = [("Type", "Name", "Identifier")]
lst_relationships = [
    (
    "Lead: Name", 
    "Lead: Type", 
    "Relationship: Type", 
    "Member: Name",
    "Member: Type"
    )
    ]

# extract the objects and relationships from the xml
tree = xml.etree.ElementTree.parse(filepath)
root = tree.getroot()
relationships = root.findall('xmlns:relationships/xmlns:relationship', ns)
elements = root.findall('xmlns:elements/xmlns:element', ns)

#add the objects to the list
for element in elements:
    identifier = element.attrib['identifier']
    element_type = element.attrib[xsi + "type"]
    orbus_element_type = re.sub( r"([A-Z])", r" \1", element_type).lstrip()
    orbus_element_type = orbus_element_type.capitalize()
    name_element = element.find('xmlns:name', ns)
    name = name_element.text
    if name not in lst_exceptions:
        element_details = (orbus_element_type, name, identifier)
        lst_elements.append(element_details)

# add the relationships to the list
for relationship in relationships:
    source_id = relationship.attrib['source']
    target_id = relationship.attrib['target']
    for item in lst_elements:
        if source_id in item:
            source_type = item[0]
            source_name = item[1]
        if target_id in item:
            target_type = item[0]
            target_name = item[1]
    if source_name and target_name:
        relationship_type = relationship.attrib[xsi + "type"]
        orbus_relationship_type = "Archimate: " + relationship_type
        relationship_details = (source_name, source_type, orbus_relationship_type, target_name, target_type)
        lst_relationships.append(relationship_details)

# generate the excel spreadsheet and workbooks
wb = openpyxl.Workbook()
ws_objects = wb.create_sheet("Objects", 0)

# avoid duplicate objects
unique_list = []
for orbus_object in lst_elements:
    orbus_type, name, identifier = orbus_object
    detail = (orbus_type, name)
    if detail not in unique_list:
        unique_list.append(detail)
        ws_objects.append(detail)
    
ws_relationships = wb.create_sheet("Relationships", 1)

for relationship in lst_relationships:
    ws_relationships.append(relationship)
    
wb.save(destination_filepath)
print("Saved file " + destination_filepath)    
