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
        """
        Sequence of |AbstractNumberingDefinition| contained in document.
        """
        return [AbstractNumberingDefinition(x, self._element)
                for x in self._element.abstractNum_lst]

    @property
    def numbering_instances(self):
        """
        Sequence of |NumberingInstance| contained in the document.
        """
        return [NumberingInstance(x, self, self._document_part) for x in
                self._element.num_lst]

    def create_new_abstract_numbering_definition(self,
                                                 name=None,
                                                 hanging_indent=Inches(0.25),
                                                 leading_indent=Inches(0.5),
                                                 tabsize=Inches(0.25),
                                                 levels=9
                                                 ):
        """
        Create and return |AbstractNumberingDefinition| instance with next
        free ``abstractNumId``.

        *hanging_indent* is the additional indent used on body text after the first
        line. Use of *hanging_indent* allows the start margin of body text to be aligned
        across multiple lines.
        *leading_indent* is the indent from document start margin to start marign of
        body text on the first line. It is NOT the indent to the list marker.
        *tabsize* is the additional indent to be applied for each additional numbering
        level.
        *levels* is the number of child ``<w:lvl>`` elements to create. The maximum is
        9.
        """
        abstractNum_el = self._element.add_abstractNum()
        abstractNum_el.abstractNumId = self._element.next_abstract_num_id
        return AbstractNumberingDefinition.\
            initialize_element(abstractNum_el, name=name, hanging_indent=hanging_indent,
                               leading_indent=leading_indent, tabsize=tabsize, 
                               levels=levels)

    def create_new_numbering_instance(self, abstract_numbering_definition):
        """
        Create and return a new |NumberingInstance| referencing 
        *abstract_numbering_definition*.
        """
        num_el = self._element.add_num(abstract_numbering_definition.abstract_num_id)
        return NumberingInstance(num_el,
                                 self, self._document_part)

    def clear_abstract_numbering(self):
        """
        Remove all abstract numbering definitions.
        """
        self._element.remove_all('w:abstractNum')

    def clear_numbering_instances(self):
        """
        Remove all numbering definitions.
        """
        self._element.remove_all('w:num')
