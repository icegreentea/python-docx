import pytest

import docx

class ManuallyCheckNumbering:

    @pytest.mark.manual
    def it_can_create_bullets(self):
        doc = docx.Document()
        abnum = doc.create_new_bullet_definition()
        numist = doc.create_new_numbering_instance(abnum)
        numist.add_paragraph(1, "b1")
        numist.add_paragraph(1, "b2")
        numist.add_paragraph(2, "b3")
        doc.save("blah.docx")