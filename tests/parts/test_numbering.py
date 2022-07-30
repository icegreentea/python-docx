# encoding: utf-8

"""
Test suite for the docx.parts.numbering module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from docx.api import Document as OpenDocument
from docx.numbering import AbstractNumberingDefinition, NumberingInstance
from docx.parts.numbering import NumberingPart

from ..unitutil.cxml import element


class DescribeNumberingPart(object):
    def it_can_create_abstract_numbering_definition(self, empty_numbering_fixture):
        doc, num_part = empty_numbering_fixture
        ab_num = num_part.create_new_abstract_numbering_definition(name="test_name")
        assert isinstance(ab_num, AbstractNumberingDefinition)
        assert "test_name" == ab_num.name
        
        assert 1 == len(num_part.abstract_numbering_definitions)
        assert ab_num == num_part.abstract_numbering_definitions[0]

    def it_can_create_numbering_instances(self, empty_numbering_fixture):
        doc, num_part = empty_numbering_fixture
        ab_num = num_part.create_new_abstract_numbering_definition(name="test_name")
        num_inst = num_part.create_new_numbering_instance(ab_num)
        assert isinstance(num_inst, NumberingInstance)

        assert 1 == len(num_part.numbering_instances)

    def it_has(self, empty_numbering_fixture):
        doc, num_part = empty_numbering_fixture
        assert isinstance(num_part, NumberingPart)
        assert 0 == len(num_part.abstract_numbering_definitions)

    @pytest.fixture
    def template_document_fixture(self, request):
        doc = OpenDocument()
        return doc, doc._part.numbering_part

    @pytest.fixture
    def empty_numbering_fixture(self, request):
        doc = OpenDocument()
        doc._part.numbering_part._element = element("w:numbering")
        return doc, doc._part.numbering_part

