import gzip
from lxml import etree

ns = {'gnm': "http://www.gnumeric.org/v10.dtd",
      'xsi': "http://www.w3.org/2001/XMLSchema-instance",
      'office': "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
      'xlink': "http://www.w3.org/1999/xlink",
      'dc': "http://purl.org/dc/elements/1.1/",
      'meta': "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
      'ooo': "http://openoffice.org/2004/office"}

with gzip.open('Investments.gnumeric', mode='rb') as fin:
    contents = fin.read()

root = etree.fromstring(contents)

sheet = root.find('gnm:Sheets', ns)[1]
cells = sheet.find('gnm:Cells', ns)
