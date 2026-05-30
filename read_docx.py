import zipfile
import xml.etree.ElementTree as ET
import os

docx_file = 'filedepeitonovo.docx'
with zipfile.ZipFile(docx_file) as z:
    media = [f for f in z.namelist() if f.startswith('word/media/')]
    print("Media files:", media)
    for m in media:
        z.extract(m, '.')
        print(f"Extracted {m}")
        
    xml_content = z.read('word/document.xml')
    tree = ET.fromstring(xml_content)
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    texts = tree.findall('.//w:t', namespaces)
    print("Extracted text:")
    for t in texts:
        if t.text:
            print(t.text)
