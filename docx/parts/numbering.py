# encoding: utf-8

"""
|NumberingPart| and closely related objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from docx.enum.text import WD_TAB_ALIGNMENT

from ..opc.part import XmlPart
from ..shared import Twips, lazyproperty
from docx.shared import Inches
from docx.numbering import AbstractNumberingDefinition, NumberingInstance


class NumberingPart(XmlPart):
    """
    Proxy for the numbering.xml part containing numbering definitions for
    a document or glossary.
    """
    '''
    @classmethod
    def new(cls):
        """
        Return newly created empty numbering part, containing only the root
        ``<w:numbering>`` element.
        """
        raise NotImplementedError

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
        return [NumberingInstance(x, self, self._document_part) for x in self._element.num_lst]

    def create_new_abstract_numbering_definition(self, name=None):
        abstractNum_el = self._element.add_abstractNum()
        abstractNum_el.abstractNumId = self._element.next_abstract_num_id
        if name is not None:
            _name = abstractNum_el.get_or_add_name()
            _name.val = name
        for i in range(0, 9):
            lvl = abstractNum_el.add_lvl()
            lvl.ilvl = i
        return AbstractNumberingDefinition(abstractNum_el)

    def create_new_bullet_definition(self, name=None, indent_size=Inches(0.25),
                                     tabsize=Inches(0.25), bullet_text="\u2022"):
        abstract_num = self.create_new_abstract_numbering_definition(name)
        abstract_num_el = abstract_num._element
        for i, lvl in enumerate(abstract_num_el.lvl_lst):
            _tabsize_emu = Twips(tabsize * (i+2)).emu
            _indent_emu = Twips(indent_size).emu

            lvl.numFmt.val = "bullet"
            lvl.lvlText.val = bullet_text
            pPr = lvl.get_or_add_pPr()

            tabstops = pPr._add_tabs()
            tabstops.insert_tab_in_order(_indent_emu, WD_TAB_ALIGNMENT.NUM, None)

            indent = pPr.get_or_add_ind()
            indent.left = _tabsize_emu
            indent.hanging = _indent_emu

        return abstract_num

    def create_new_numbering_instance(self, abstract_numbering_definition):
        num_el = self._element.add_num(abstract_numbering_definition.abstract_num_id)
        return NumberingInstance(num_el,
                                 self, self._document_part)


'''
class _NumberingDefinitions(object):
    """
    Collection of |_NumberingDefinition| instances corresponding to the
    ``<w:num>`` elements in a numbering part.
    """
    def __init__(self, numbering_elm):
        super(_NumberingDefinitions, self).__init__()
        self._numbering = numbering_elm

    def __len__(self):
        return len(self._numbering.num_lst)
'''