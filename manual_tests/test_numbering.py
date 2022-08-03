import pytest
import os

import docx

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
        abnum = doc.create_new_bullet_definition()
        numist = doc.create_new_numbering_instance(abnum)
        numist.add_paragraph(0, "b1")
        numist.add_paragraph(0, "b2")
        numist.add_paragraph(1, "b3")
        save_document(doc, "bullet-list.docx")

    @pytest.mark.manual
    def it_can_create_decimal(self):
        doc = docx.Document()
        abnum = doc.create_new_simple_decimal_definition()
        numist = doc.create_new_numbering_instance(abnum)
        numist.add_paragraph(0, "b1")
        numist.add_paragraph(0, "b2")
        numist.add_paragraph(1, "b3")
        save_document(doc, "numbering-decimal-list.docx")
