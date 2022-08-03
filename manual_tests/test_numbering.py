import pytest
import os

import docx
from docx.oxml.ns import qn

CURDIR = os.path.abspath(os.path.dirname(__file__))
OUTPUT_DIR = os.path.join(CURDIR, "output")


def save_document(doc, filename):
    if not os.path.isdir(OUTPUT_DIR):
        os.mkdir(OUTPUT_DIR)
    doc.save(os.path.join(OUTPUT_DIR, filename))


class ManuallyCheckNumbering:

    @pytest.mark.manual
    def it_can_create_bullets(self):
        doc = docx.Document()
        for child in doc._part.numbering_part._element[:]:
            if child.tag == qn("w:abstractNum") or child.tag == qn("w:num"):
                doc._part.numbering_part._element.remove(child)

        abnum = doc.create_new_bullet_definition()
        numist = doc.create_new_numbering_instance(abnum)
        numist.add_paragraph(0, "b1")
        numist.add_paragraph(0, "b2")
        numist.add_paragraph(1, "b3")
        save_document(doc, "bullet-list.docx")
