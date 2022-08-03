from collections.abc import Sequence
from copy import deepcopy
from docx.shared import ElementProxy, Inches


class NumberingInstance:
    def __init__(self, numbering_element, numbering_part, document_part):
        from docx.document import Document
        self._numbering_part = numbering_part
        self._doc_part = document_part
        self._element = numbering_element
        self._doc = Document(self._doc_part.element, self._doc_part)

    def add_paragraph(self, indent_level, text):
        para = self._doc.add_paragraph(text)
        p_elm = para._element
        pPr = p_elm.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        ilvl = numPr.get_or_add_ilvl()
        ilvl.val = indent_level

        numId = numPr.get_or_add_numId()
        numId.val = self._element.numId

        return para


class AbstractNumberingDefinition(ElementProxy, Sequence):
    @property
    def name(self):
        return self._element.name.val

    @name.setter
    def name(self, value):
        self._element.name.val = value

    @property
    def abstract_num_id(self):
        return self._element.abstractNumId

    def __getitem__(self, key):
        if isinstance(key, slice):
            return [
                NumberingLevelDefinition(lvl, self._element)
                for lvl in self._element.lvl_lst[key]
            ]
        return NumberingLevelDefinition(self._element.lvl_lst[key], self._element)

    def __iter__(self):
        for lvl in self._element.lvl_lst:
            yield NumberingLevelDefinition(lvl, self._element)

    def __len__(self):
        return len(self._element.lvl_lst)


class NumberingLevelDefinition(ElementProxy):
    @property
    def start(self):
        start = self._element.start
        if start is None:
            return None
        return start.val

    @start.setter
    def start(self, value):
        start = self._element.get_or_add_start()
        start.val = value

    @property
    def number_format(self):
        numFmt = self._element.numFmt
        if numFmt is None:
            return None
        return numFmt.val

    @number_format.setter
    def number_format(self, value):
        if value not in ("bullet", "decimal", "none"):
            raise ValueError
        numFmt = self._element.get_or_add_numFmt()
        numFmt.val = value

    @property
    def numbering_level(self):
        return self._element.ilvl

    @property
    def restart_numbering_level(self):
        lvlRestart = self._element.lvlRestart
        if lvlRestart is None:
            return None
        return lvlRestart.val

    @restart_numbering_level.setter
    def restart_numbering_level(self, value):
        lvlRestart = self._element.get_or_add_lvlRestart()
        lvlRestart.val = value

    @property
    def numbering_level_text(self):
        lvlText = self._element.lvlText
        if lvlText is None:
            return None
        return lvlText.val

    @numbering_level_text.setter
    def numbering_level_text(self, value):
        lvlText = self._element.get_or_add_lvlText()
        lvlText.val = value

    @property
    def justification(self):
        lvlJc = self._element.lvlJc
        if lvlJc is None:
            return None
        return lvlJc.val

    @justification.setter
    def justification(self, value):
        lvlJc = self._element.get_or_add_lvlJc()
        lvlJc.val = value

    @property
    def paragraph_properties(self):
        """
        """
        from docx.text.parfmt import ParagraphFormat

        pPr = self._element.pPr
        if pPr is None:
            return None

        return ParagraphFormat.numbering_level_wrapper(pPr)

    @paragraph_properties.setter
    def paragraph_properties(self, value):
        """
        Sets the numbering level associated paragraph properties for this
        numbering level. These properties set the formatting for paragraphs (body text)
        at this numbering level.

        Sets the properties by creating a clone of *value* and setting the clone,
        overwriting any existing values. Pass in |None| to unset this property
        and fall back to specification defaults.
        """
        from docx.text.parfmt import ParagraphFormat
        from docx.oxml.text.parfmt import CT_PPr
        if value is None:
            self._element.set_pPr(None)
        elif isinstance(value, ParagraphFormat):
            val = deepcopy(value._element.pPr)
            self._element.set_pPr(val)
        elif isinstance(value, CT_PPr):
            self._element.set_pPr(deepcopy(value))

    def create_new_paragraph_properties(self):
        from docx.text.parfmt import ParagraphFormat

        self._element._remove_pPr()
        pPr = self._element.get_or_add_pPr()
        return ParagraphFormat.numbering_level_wrapper(pPr)
