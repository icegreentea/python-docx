# encoding: utf-8

"""
|NumberingPart| and closely related objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..opc.part import XmlPart
from ..shared import Emu, lazyproperty
from docx.shared import Inches
from docx.numbering import AbstractNumberingDefinition, NumberingInstance


class NumberingPart(XmlPart):
    """
    Proxy for the numbering.xml part containing numbering definitions for
    a document or glossary.
    """
    @classmethod
    def new(cls):
        """
        Return newly created empty numbering part, containing only the root
        ``<w:numbering>`` element.
        """
        raise NotImplementedError
    '''
    @lazyproperty
    def numbering_definitions(self):
        """
        The |_NumberingDefinitions| instance containing the numbering
        definitions (<w:num> element proxies) for this numbering part.
        """
        return _NumberingDefinitions(self._element)
    '''
    @lazyproperty
    def _document_part(self):
        """|DocumentPart| object for this package."""
        return self.package.main_document_part

    @property
    def abstract_numbering_definitions(self):
        return [AbstractNumberingDefinition(x, self._element)
                for x in self._element.abstractNum_lst]

    @property
    def numbering_instances(self):
        return [NumberingInstance(x, self, self._document_part) for x in
                self._element.num_lst]

    def create_new_abstract_numbering_definition(self,
                                                 name=None,
                                                 hanging_indent=Inches(0.25),
                                                 leading_indent=Inches(0.5),
                                                 tabsize=Inches(0.25)
                                                 ):
        abstractNum_el = self._element.add_abstractNum()
        abstractNum_el.abstractNumId = self._element.next_abstract_num_id
        if name is not None:
            _name = abstractNum_el.get_or_add_name()
            _name.val = name
        for i in range(0, 9):
            lvl = abstractNum_el.add_lvl()
            lvl.ilvl = i
            pPr = lvl.get_or_add_pPr()
            indent = pPr.get_or_add_ind()
            indent.left = Emu(leading_indent).emu + i * Emu(tabsize).emu
            indent.hanging = Emu(hanging_indent).emu
        return AbstractNumberingDefinition(abstractNum_el)

    def create_new_bullet_definition(self, name=None,
                                     hanging_indent=Inches(0.25),
                                     leading_indent=Inches(0.5),
                                     tabsize=Inches(0.25),
                                     bullet_text="\u2022"):
        abstract_num = \
            self.create_new_abstract_numbering_definition(name,
                                                          hanging_indent=hanging_indent,
                                                          leading_indent=leading_indent,
                                                          tabsize=tabsize)
        abstract_num.set_level_number_format("bullet")
        abstract_num.set_level_text(bullet_text)
        return abstract_num

    def create_new_simple_decimal_definition(self, name=None,
                                             hanging_indent=Inches(0.25),
                                             leading_indent=Inches(0.5),
                                             tabsize=Inches(0.25)):
        abstract_num = \
            self.create_new_abstract_numbering_definition(name,
                                                          hanging_indent=hanging_indent,
                                                          leading_indent=leading_indent,
                                                          tabsize=tabsize)
        abstract_num.set_level_number_format("decimal")
        for i, lvl in enumerate(abstract_num):
            lvl.numbering_level_text = "%{}.".format(lvl.numbering_level + 1)
            lvl.start = 1
        return abstract_num

    def create_new_numbering_instance(self, abstract_numbering_definition):
        num_el = self._element.add_num(abstract_numbering_definition.abstract_num_id)
        return NumberingInstance(num_el,
                                 self, self._document_part)

    def clear_abstract_numbering(self):
        self._element.remove_all('w:abstractNum')

    def clear_numbering_instances(self):
        self._element.remove_all('w:num')
