# encoding: utf-8

"""
Test suite for the docx.oxml.text.paragraph module.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)
from pyparsing import re

import pytest

from ...unitutil.cxml import element, xml


class DescribeCT_P(object):
    def it_can_return_sectPr(self, get_sectPr_fixture):
        p, expected_return = get_sectPr_fixture
        ret = p.get_sectPr()
        if expected_return is None:
            assert ret is None
        else:
            assert ret.xml == expected_return.xml
    
    def it_can_set_sectPr(self, set_sectPr_fixture):
        p, set_element, expected_xml = set_sectPr_fixture
        p.set_sectPr(set_element)
        assert p.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('w:p', None),
        ('w:p/(w:pPr/(w:sectPr))', 'w:sectPr'),
        ('w:p/(w:pPr/(w:sectPr/w:type{w:val=continuous}))', 'w:sectPr/w:type{w:val=continuous}')
    ])
    def get_sectPr_fixture(self, request):
        initial_cxml, expected_return = request.param
        r = element(initial_cxml)
        if expected_return is None:
            return r, None
        else:
            return r, element(expected_return)

    @pytest.fixture(params=[
        ('w:p', 'w:sectPr', 'w:p/w:pPr/w:sectPr'),
        ('w:p/w:pPr', 'w:sectPr', 'w:p/w:pPr/w:sectPr'),
        ('w:p/w:pPr/w:sectPr/w:type{w:val=continuous}', 'w:sectPr', 'w:p/w:pPr/w:sectPr'),
        ('w:p/w:pPr/w:sectPr/w:type{w:val=continuous}', None, 'w:p/w:pPr'),
    ])
    def set_sectPr_fixture(self, request):
        initial_cxml, set_element_cxml, expected_cxml = request.param
        r = element(initial_cxml)
        expected_xml = element(expected_cxml).xml
        if set_element_cxml is None:
            return r, None, expected_xml
        else:
            return r, element(set_element_cxml), expected_xml