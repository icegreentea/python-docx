# encoding: utf-8

"""Unit test suite for the docx.numbering module"""

from __future__ import absolute_import, division, print_function, unicode_literals

import pytest

from docx.numbering import Numbering, AbstractNumbering, AbstractNumberingLevel, NumberingInstance

from .unitutil.cxml import element, xml

class DescribeNumbering(object):
    
    def it_provides_list_of_abstract_numberings(self):
        pass

    def it_provides_list_of_concrete_numberings(self):
        pass

    def it_can_clear_abstract_numberings(self):
        pass

    def it_can_clear_concrete_numberings(self):
        pass

    def it_provides_abstract_numbering_by_id(self):
        pass

    def it_provides_abstract_numbering_by_name(self):
        pass

    def it_provides_concrete_numbering_by_numid(self):
        pass

    def it_provides_concrete_numberings_by_abstract_name(self):
        pass

    def it_provides_concrete_numberings_by_abstract_id(self):
        pass

    # fixture --------------------------------------------------------
    @pytest.fixture(params=[
        ('w:numbering/(w:abstractNum{w:abstractNumId="1"})')
    ])
    def getabstractnum_id_fixture(self, request):
        pass

    def getbasesetup_fixture(self):
        pass