import xml.etree.ElementTree as ET
import os.path
import tempfile
import csv
import uuid

import docX2csv_lib


def updcsv(csvList, style, style_text,  page):
    csvList.append(
    {
        'Style' : style,
        'Style Text' : style_text,
        'Page': page,
    })



# docX_file = 'WorksOK.docx'
docX_file = 'Problem.docx'




tmp_dir = tempfile.TemporaryDirectory()
xmlfile = os.path.join(docX2csv_lib.extract_document_xml(docX_file, tmp_dir.name),docX2csv_lib.XML_DOC_PATH.replace('/', '\\'))
# generate the path and name of the csv files. This is identical to the source document except for a different extension
csv_fl = os.path.splitext(docX_file)[0] + '.csv'
        
crossref_items = ['RACIResp', 'RACIAccountable', 'RACIInf']
crossref_style_dict = {}
    
# Process the file
parser = ET.XMLParser(encoding="utf-8")
tree = ET.parse(xmlfile, parser=parser)

root = tree.getroot()
page = 1

# Because of a quirk in the docx xml format there can be two page breaks on two adjacent
# './/w:p' nodes one related to a style and the other being a <w:lastRenderedPageBreak/>
# in this scenario only  count as one page.
pagebreak_prior = False

ET.register_namespace("w", docX2csv_lib.NS_URI)
ns = {"w": docX2csv_lib.NS_URI}
# ET.dump(tree)

for x in root.findall('.//w:p', ns):
    # print (x)
    style_text = ''
    style = None
    if docX2csv_lib.page_break(x):
        if not pagebreak_prior:
            page += 1
        pagebreak_prior = True
    else:
        pagebreak_prior = False      
    
    for y in x:
        if y.tag == docX2csv_lib.NW_URI_TAG + 'pPr':
            # Process Cross Reference Styles
            style, styletag_found = docX2csv_lib.proc_pPr_pStyle(y, crossref_items) or (None, False)
            if style is None:
                break
            else:
                crossref_style_dict[uuid.uuid4().node] = (style, docX2csv_lib.proc_r_t(x), page)
        

csvList = []
            
for x in crossref_style_dict:
    style = crossref_style_dict[x][0]
    style_text = crossref_style_dict[x][1]
    page = crossref_style_dict[x][2]
    
    updcsv(csvList, style, style_text, page)
    
csvColumns = ['Style','Style Text','Page']
try:
    with open(csv_fl, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csvColumns)
        writer.writeheader()
        for data in csvList:
            writer.writerow(data)
except IOError:
    print("I/O error")


