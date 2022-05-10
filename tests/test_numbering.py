# encoding: utf-8

"""Unit test suite for the docx.numbering module"""

from __future__ import absolute_import, division, print_function, unicode_literals
from re import L
from _pytest.assertion import AssertionState
from _pytest.compat import num_mock_patch_args

import pytest

from docx.numbering import Numbering, AbstractNumbering, AbstractNumberingLevel, NumberingInstance
from docx.oxml import numbering
from docx.shared import Twips, Inches
from docx.text.parfmt import ParagraphFormat

from .unitutil.cxml import element, xml

class DescribeNumbering(object):
    
    def it_provides_list_of_abstract_numberings(self, simple_numbering_fixture):
        numbering = Numbering(element(simple_numbering_fixture))
        assert 3 == len(numbering.abstract_numberings)

    def it_provides_list_of_concrete_numberings(self, simple_numbering_fixture):
        numbering = Numbering(element(simple_numbering_fixture))
        assert 4 == len(numbering.numbering_instances)

    def it_can_clear_abstract_numberings(self, simple_numbering_fixture, simple_no_abstract_fixture):
        numbering = Numbering(element(simple_numbering_fixture))
        
        numbering.clear_abstract_numbering()

        assert 0 == len(numbering.abstract_numberings)
        expected_xml = xml(simple_no_abstract_fixture)
        assert expected_xml == numbering._element.xml

    def it_can_clear_numbering_instances(self, simple_numbering_fixture, simple_no_instance_fixture):
        numbering = Numbering(element(simple_numbering_fixture))
        
        numbering.clear_numbering_instances()

        assert 0 == len(numbering.numbering_instances)
        expected_xml = xml(simple_no_instance_fixture)
        assert expected_xml == numbering._element.xml

    def it_provides_abstract_numbering_by_id(self, simple_numbering_fixture):
        numbering = Numbering(element(simple_numbering_fixture))

        for i in range(1,4):
            abnum = numbering.get_abstract_numbering_by_id(i)
            assert abnum is not None
            assert abnum.name == 'list-%s' %i

    def it_provides_abstract_numbering_by_name(self, simple_numbering_fixture):
        numbering = Numbering(element(simple_numbering_fixture))

        for i in range(1,4):
            abnum = numbering.get_abstract_numbering_by_name('list-%s' %i)
            assert abnum is not None
            assert abnum.abstract_num_id == i
            assert abnum.name == 'list-%s' %i

    def it_provides_numbering_instances_by_abstract_name(self, simple_numbering_fixture):
        numbering = Numbering(element(simple_numbering_fixture))
        
        _nums = numbering.get_numbering_instance_by_abstract_numbering("list-1")
        assert 1 == len(_nums)
        assert 1 == _nums[0].numId

        _nums = numbering.get_numbering_instance_by_abstract_numbering("list-2")
        assert 1 == len(_nums)
        assert 2 == _nums[0].numId

        _nums = numbering.get_numbering_instance_by_abstract_numbering("list-3")
        assert 2 == len(_nums)
        assert 3 == _nums[0].numId
        assert 4 == _nums[1].numId

    def it_provides_numbering_instances_by_abstract_id(self, simple_numbering_fixture):
        numbering = Numbering(element(simple_numbering_fixture))

        _nums = numbering.get_numbering_instance_by_abstract_numbering(1)
        assert 1 == len(_nums)
        assert 1 == _nums[0].numId

        _nums = numbering.get_numbering_instance_by_abstract_numbering(2)
        assert 1 == len(_nums)
        assert 2 == _nums[0].numId

        _nums = numbering.get_numbering_instance_by_abstract_numbering(3)
        assert 2 == len(_nums)
        assert 3 == _nums[0].numId
        assert 4 == _nums[1].numId

    def it_provides_numbering_instance_by_id(self, simple_numbering_fixture):
        numbering = Numbering(element(simple_numbering_fixture))
        for i in range(1,5):
            _num = numbering.get_numbering_instance_by_id(i)
            assert _num is not None
            assert i == _num.numId

    def it_can_create_abstract_num_with_name(self, empty_numberings_fixture):
        numbering = Numbering(element(empty_numberings_fixture))
        ab_num = numbering.create_abstract_numbering("name-1")
        assert "name-1" == ab_num.name
        expected_tmpl = '''w:numbering/
            w:abstractNum{w:abstractNumId=%s}/(
                w:name{w:val=%s},
                w:multiLevelType{w:val=multiLevel}
            )
        ''' % (ab_num.abstract_num_id, 'name-1')
        expected_xml = xml(expected_tmpl)
        assert expected_xml == numbering._element.xml

    def it_can_create_numbering_instance(self, empty_numberings_fixture):
        numbering = Numbering(element(empty_numberings_fixture))
        ab_num = numbering.create_abstract_numbering("name-1")

        expected_tmpl = '''
            w:num{w:numId=%s}/w:abstractNumId{w:val=1}
        '''

        num1 = numbering.create_numbering_instance(ab_num)
        assert xml(expected_tmpl % 1) == num1.element.xml

        num2 = numbering.create_numbering_instance(ab_num.name)
        assert xml(expected_tmpl % 2) == num2.element.xml

        num3 = numbering.create_numbering_instance(ab_num.abstract_num_id)
        assert xml(expected_tmpl % 3) == num3.element.xml

    def it_can_create_bullet_definition(self, empty_numberings_fixture):
        numbering = Numbering(element(empty_numberings_fixture))
        abs_num = numbering.create_bullet_abstract_numbering(
                "bullet-list", tab_width=Inches(0.25))

        assert 9 == len(abs_num.levels)
        for i, lvl in enumerate(abs_num.levels):
            assert "bullet" == lvl.numFmt
            assert i == lvl.ilvl
            assert Twips(360 * (i+2)).emu == lvl.left_indent
            #assert Twips(360).emu == lvl.indent_hanging

    def it_can_create_decimal_definition(self, empty_numberings_fixture):
        numbering = Numbering(element(empty_numberings_fixture))
        abs_num = numbering.create_decimal_abstract_numbering(
                "decimal-list", tab_width=Inches(0.25))

        assert 9 == len(abs_num.levels)
        for i, lvl in enumerate(abs_num.levels):
            assert "decimal" == lvl.numFmt
            assert i == lvl.ilvl
            assert Twips(360 * (i+2)).emu == lvl.left_indent
            #assert Twips(360).emu == lvl.indent_hanging
            assert "%%%s." % (i+1) == lvl.lvlText
            assert 1 == lvl.start


    # fixture --------------------------------------------------------
    @pytest.fixture
    def simple_numbering_fixture(self):
        tmpl = (
            'w:numbering/('
                'w:abstractNum{w:abstractNumId=1}/w:name{w:val=list-1},'
                'w:abstractNum{w:abstractNumId=2}/w:name{w:val=list-2},'
                'w:abstractNum{w:abstractNumId=3}/w:name{w:val=list-3},'
                'w:num{w:numId=1}/w:abstractNumId{w:val=1},'
                'w:num{w:numId=2}/w:abstractNumId{w:val=2},'
                'w:num{w:numId=3}/w:abstractNumId{w:val=3},'
                'w:num{w:numId=4}/w:abstractNumId{w:val=3}'
            ')'
        )
        return tmpl

    @pytest.fixture
    def simple_no_abstract_fixture(self):
        tmpl = (
            'w:numbering/('
                'w:num{w:numId=1}/w:abstractNumId{w:val=1},'
                'w:num{w:numId=2}/w:abstractNumId{w:val=2},'
                'w:num{w:numId=3}/w:abstractNumId{w:val=3},'
                'w:num{w:numId=4}/w:abstractNumId{w:val=3}'
            ')'
        )
        return tmpl

    @pytest.fixture
    def simple_no_instance_fixture(self):
        tmpl = (
            'w:numbering/('
                'w:abstractNum{w:abstractNumId=1}/w:name{w:val=list-1},'
                'w:abstractNum{w:abstractNumId=2}/w:name{w:val=list-2},'
                'w:abstractNum{w:abstractNumId=3}/w:name{w:val=list-3}'            ')'
        )
        return tmpl

    @pytest.fixture
    def empty_numberings_fixture(self):
        tmpl = 'w:numbering'
        return tmpl

class DescribeAbstractNumbering(object):
    def it_provides_abstract_num_id_and_name(self, named_abstract_numbering_fixture):
        tmpl, expected_id, expected_name = named_abstract_numbering_fixture
        ab_num = AbstractNumbering(element(tmpl))
        assert expected_id == ab_num.abstract_num_id
        assert expected_name == ab_num.name

    def it_can_get_or_add_level(self, empty_abstract_numbering_fixture):
        ab_num = AbstractNumbering(element(empty_abstract_numbering_fixture))
        lvl = ab_num.get_or_add_level(0)
        assert 1 == len(ab_num.levels)
        expected_xml = xml("""
            w:abstractNum{w:abstractNumId=1}/w:lvl{w:ilvl=0}
        """)
        assert expected_xml == ab_num._element.xml

    def it_can_create_bullet_level(self, empty_abstract_numbering_fixture):
        ab_num = AbstractNumbering(element(empty_abstract_numbering_fixture))
        expected_template = """w:lvl{w:ilvl=0}/(
            w:numFmt{w:val=bullet},
            w:lvlText{w:val=.},
            w:pPr/w:ind{w:left=%s,w:hanging=%s}
        )""" % (2 * Inches(0.25).twips, Inches(0.25).twips)
        expected_xml = xml(expected_template)
        assert 0 == len(ab_num.levels)
        lvl = ab_num.create_bullet_level(0, lvlText=".")
        assert expected_xml == lvl._element.xml
        assert 1 == len(ab_num.levels)
        

    def it_can_create_decimal_level(self, empty_abstract_numbering_fixture):
        ab_num = AbstractNumbering(element(empty_abstract_numbering_fixture))
        expected_template = """w:lvl{w:ilvl=0}/(
            w:start{w:val=1},
            w:numFmt{w:val=decimal},
            w:lvlText{w:val=%%1.},
            w:pPr/w:ind{w:left=%s,w:hanging=%s}
        )""" % (2 * Inches(0.25).twips, Inches(0.25).twips)
        expected_xml = xml(expected_template)
        assert 0 == len(ab_num.levels)
        lvl = ab_num.create_decimal_level(0)
        assert expected_xml == lvl._element.xml
        assert 1 == len(ab_num.levels)

    @pytest.fixture
    def empty_abstract_numbering_fixture(self):
        tmpl = """
            w:abstractNum{w:abstractNumId=1}
        """
        return tmpl

    @pytest.fixture
    def named_abstract_numbering_fixture(self):
        tmpl = """
            w:abstractNum{w:abstractNumId=1}/w:name{w:val=list1}
        """
        return tmpl, 1, "list1"

class DescribeAbstractNumberingLevel(object):

    def it_provides_access_to_empty_attributes(self, empty_lvl):
        tmpl, expected_ilvl = empty_lvl
        lvl = AbstractNumberingLevel(element(tmpl))
        assert expected_ilvl == lvl.ilvl
        assert lvl.start is None
        assert lvl.numFmt is None
        assert lvl.lvlRestart is None
        assert lvl.lvlText is None
        assert isinstance(lvl.pPr, ParagraphFormat)
        assert lvl.left_indent is None
        assert lvl.right_indent is None
        assert lvl.first_line_indent is None

    def it_provides_setters_to_empty_attributes(self, empty_lvl):
        tmpl, expected_ilvl = empty_lvl
        lvl = AbstractNumberingLevel(element(tmpl))
        lvl.start = 1
        lvl.numFmt = "decimal"
        lvl.lvlRestart = 2
        lvl.lvlText = "%1.%2."
        lvl.left_indent = Inches(0.25)
        lvl.right_indent = Inches(0.5)
        lvl.first_line_indent = -Inches(0.25)
        expected_tmpl = """
            w:lvl{w:ilvl=%s}/(
                w:start{w:val=1},
                w:numFmt{w:val=decimal},
                w:lvlRestart{w:val=2},
                w:lvlText{w:val=%%1.%%2.},
                w:pPr/w:ind{w:left=%s,w:right=%s,w:hanging=%s}
            )
        """ % (expected_ilvl, Inches(0.25).twips, Inches(0.50).twips, Inches(0.25).twips)
        expected_xml = xml(expected_tmpl)
        assert expected_xml == lvl._element.xml

    def it_provides_paragraph_access(self, empty_lvl):
        tmpl, expected_ilvl = empty_lvl
        lvl = AbstractNumberingLevel(element(tmpl))
        lvl.pPr.left_indent = Inches(0.25)
        assert lvl.left_indent == Inches(0.25)


    @pytest.fixture
    def empty_lvl(self):
        tmpl = """
            w:lvl{w:ilvl=0}
        """
        return tmpl, 0

class DescribeNumberingInstance(object):
    def it_provides_numId(self, simple_numbering_fixture):
        tmpl, abstract_num_id, abstract_num_name, num_id = simple_numbering_fixture
        numbering = Numbering(element(tmpl))
        num_inst = numbering.get_numbering_instance_by_id(num_id)
        assert num_inst is not None
        assert num_inst.numId == num_id

    def it_provides_abstract_num_id(self, simple_numbering_fixture):
        tmpl, abstract_num_id, abstract_num_name, num_id = simple_numbering_fixture
        numbering = Numbering(element(tmpl))
        num_inst = numbering.get_numbering_instance_by_id(num_id)
        assert num_inst is not None
        assert num_inst.abstract_num_id == abstract_num_id

    def it_provides_abstract_num(self, simple_numbering_fixture):
        tmpl, abstract_num_id, abstract_num_name, num_id = simple_numbering_fixture
        numbering = Numbering(element(tmpl))
        num_inst = numbering.get_numbering_instance_by_id(num_id)
        assert num_inst is not None

        abs_num = num_inst.abstract_num
        assert abstract_num_id == abs_num.abstract_num_id

    def it_can_set_abstract_num_id(self, extra_abstract_fixture):
        numbering = Numbering(element(extra_abstract_fixture))
        num_inst = numbering.get_numbering_instance_by_id(1)
        assert num_inst is not None
        assert 1 == num_inst.abstract_num_id
        num_inst.abstract_num_id = 2
        assert 2 == num_inst.abstract_num_id
        assert 2 == num_inst.abstract_num.abstract_num_id

    def it_can_set_abstract_num(self, extra_abstract_fixture):
        numbering = Numbering(element(extra_abstract_fixture))
        num_inst = numbering.get_numbering_instance_by_id(1)
        abs_num = numbering.get_abstract_numbering_by_id(2)
        assert num_inst is not None
        assert 1 == num_inst.abstract_num_id
        num_inst.abstract_num = abs_num
        assert 2 == num_inst.abstract_num_id
        assert 2 == num_inst.abstract_num.abstract_num_id

    @pytest.fixture
    def simple_numbering_fixture(self):
        abstract_num_id = 1
        abstract_num_name = "list-1"
        num_id = 2
        tmpl = (
            'w:numbering/('
                'w:abstractNum{w:abstractNumId=%s}/w:name{w:val=%s},'
                'w:num{w:numId=%s}/w:abstractNumId{w:val=%s}'
            ')'
        ) % (abstract_num_id, abstract_num_name, num_id, abstract_num_id)
        
        return tmpl, abstract_num_id, abstract_num_name, num_id

    @pytest.fixture
    def extra_abstract_fixture(self):
        tmpl = """
            w:numbering/(
                w:abstractNum{w:abstractNumId=1},
                w:abstractNum{w:abstractNumId=2},
                w:num{w:numId=1}/w:abstractNumId{w:val=1}
            )
        """
        return tmpl