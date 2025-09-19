import xml.etree.ElementTree as ET
from openpyxl import Workbook

xml_data = """
<root>
    <employee>
        <id>1</id>
        <name>John Doe</name>
        <department>IT</department>
    </employee>
    <employee>
        <id>2</id>
        <name>Jane Smith</name>
        <department>HR</department>
    </employee>
</root>
"""

# Parse XML
root = ET.fromstring(xml_data)

# First im creating excel work book
wb = Workbook()
ws = wb.active
ws.title = "Employees"

#  just mapping
mapping = {
    "id": "Employee ID",
    "name": "Full Name",
    "department": "Department"
}

# here im writing headerss
headers = list(mapping.values())
ws.append(headers)

# here im wiriging rows
for emp in root.findall("employee"):
    row = []
    for xml_tag in mapping.keys():
        element = emp.find(xml_tag)
        row.append(element.text if element is not None else "")
    ws.append(row)

# here Save Excel file
wb.save("employees.xlsx")
print("Excel file created: employees.xlsx")
