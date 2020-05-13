import os
import re
from urllib.parse import quote, unquote
import zipfile

import mammoth

from .omml import omml2tex

_omath_pattern = re.compile(r'<m:oMath[^<>]*>.+?</m:oMath>', flags=re.S)
_omath_para_pattern = re.compile(r'<m:oMathPara\s*>(.+?)</m:oMathPara>',
                                 flags=re.S)


def quote_omath(xml_content):
    def replace(match):
        quoted_omath = quote(match.group(0))
        return '<w:t>$omml$ {} $/omml$</w:t>'.format(quoted_omath)

    xml_content = _omath_pattern.sub(replace, xml_content)
    xml_content = _omath_para_pattern.sub(lambda m: m.group(1), xml_content)
    return xml_content


_omml_pattern = re.compile(r'\$omml\$(.+?)\$/omml\$')


def convert_quoted_omath_to_tex(html):
    def replace(match):
        omml_content = unquote(match.group(1))
        return omml2tex(omml_content)

    return _omml_pattern.sub(replace, html)


def pre_process_docx(docx_filename):
    name_ext = list(os.path.splitext(docx_filename))
    name_ext.insert(1, '_copy')
    new_docx_filename = ''.join(name_ext)

    document_filename = 'word/document.xml'
    with zipfile.ZipFile(docx_filename) as z_in:
        with zipfile.ZipFile(new_docx_filename, 'w') as z_out:
            z_out.comment = z_in.comment
            for item in z_in.infolist():
                if item.filename != document_filename:
                    z_out.writestr(item, z_in.read(item.filename))
            document_xml = z_in.read(document_filename).decode('utf8')
            document_xml = quote_omath(document_xml).encode('utf8')
            z_out.writestr(document_filename, document_xml)
    return new_docx_filename


def convert_to_html(docx_filename):
    new_docx_filename = pre_process_docx(docx_filename)
    res = mammoth.convert_to_html(new_docx_filename)

    name_ext = list(os.path.splitext(docx_filename))
    name_ext[1] = '.html'
    html_filename = ''.join(name_ext)
    html = convert_quoted_omath_to_tex(res.value)
    with open(html_filename, 'w') as f:
        f.write(html)
