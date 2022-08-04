# encoding: utf-8

"""
Custom element classes related to the numbering part
"""

from . import OxmlElement
from .shared import CT_DecimalNumber
from .simpletypes import ST_DecimalNumber, ST_OnOff, ST_String
from .xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, RequiredAttribute, ZeroOrMore,
    ZeroOrOne
)


class CT_Num(BaseOxmlElement):
    """
    ``<w:num>`` element, which represents a concrete list definition
    instance, having a required child <w:abstractNumId> that references an
    abstract numbering definition that defines most of the formatting details.
    """
    abstractNumId = OneAndOnlyOne('w:abstractNumId')
    lvlOverride = ZeroOrMore('w:lvlOverride')
    numId = RequiredAttribute('w:numId', ST_DecimalNumber)

    def add_lvlOverride(self, ilvl):
        """
        Return a newly added CT_NumLvl (<w:lvlOverride>) element having its
        ``ilvl`` attribute set to *ilvl*.
        """
        return self._add_lvlOverride(ilvl=ilvl)

    @classmethod
    def new(cls, num_id, abstractNum_id):
        """
        Return a new ``<w:num>`` element having numId of *num_id* and having
        a ``<w:abstractNumId>`` child with val attribute set to
        *abstractNum_id*.
        """
        num = OxmlElement('w:num')
        num.numId = num_id
        abstractNumId = CT_DecimalNumber.new(
            'w:abstractNumId', abstractNum_id
        )
        num.append(abstractNumId)
        return num


class CT_NumLvl(BaseOxmlElement):
    """
    ``<w:lvlOverride>`` element, which identifies a level in a list
    definition to override with settings it contains.
    """
    startOverride = ZeroOrOne('w:startOverride', successors=('w:lvl',))
    lvl = ZeroOrOne('w:lvl')

    ilvl = RequiredAttribute('w:ilvl', ST_DecimalNumber)

    def add_startOverride(self, val):
        """
        Return a newly added CT_DecimalNumber element having tagname
        ``w:startOverride`` and ``val`` attribute set to *val*.
        """
        return self._add_startOverride(val=val)


class CT_NumPr(BaseOxmlElement):
    """
    A ``<w:numPr>`` element, a container for numbering properties applied to
    a paragraph.
    """
    ilvl = ZeroOrOne('w:ilvl', successors=(
        'w:numId', 'w:numberingChange', 'w:ins'
    ))
    numId = ZeroOrOne('w:numId', successors=('w:numberingChange', 'w:ins'))

    # @ilvl.setter
    # def _set_ilvl(self, val):
    #     """
    #     Get or add a <w:ilvl> child and set its ``w:val`` attribute to *val*.
    #     """
    #     ilvl = self.get_or_add_ilvl()
    #     ilvl.val = val

    # @numId.setter
    # def numId(self, val):
    #     """
    #     Get or add a <w:numId> child and set its ``w:val`` attribute to
    #     *val*.
    #     """
    #     numId = self.get_or_add_numId()
    #     numId.val = val


class CT_Numbering(BaseOxmlElement):
    """
    ``<w:numbering>`` element, the root element of a numbering part, i.e.
    numbering.xml
    """
    _tag_seq = (
        'w:numPicBullet', 'w:abstractNum', 'w:num', 'w:numIdMacAtCleanup'
    )
    abstractNum = ZeroOrMore('w:abstractNum', successors=_tag_seq[2:])
    num = ZeroOrMore('w:num', successors=('w:numIdMacAtCleanup',))

    del _tag_seq

    def add_num(self, abstractNum_id):
        """
        Return a newly added CT_Num (<w:num>) element referencing the
        abstract numbering definition identified by *abstractNum_id*.
        """
        next_num_id = self._next_numId
        num = CT_Num.new(next_num_id, abstractNum_id)
        return self._insert_num(num)

    def num_having_numId(self, numId):
        """
        Return the ``<w:num>`` child element having ``numId`` attribute
        matching *numId*.
        """
        xpath = './w:num[@w:numId="%d"]' % numId
        try:
            return self.xpath(xpath)[0]
        except IndexError:
            raise KeyError('no <w:num> element with numId %d' % numId)

    @property
    def _next_numId(self):
        """
        The first ``numId`` unused by a ``<w:num>`` element, starting at
        1 and filling any gaps in numbering between existing ``<w:num>``
        elements.
        """
        numId_strs = self.xpath('./w:num/@w:numId')
        num_ids = [int(numId_str) for numId_str in numId_strs]
        for num in range(1, len(num_ids)+2):
            if num not in num_ids:
                break
        return num

    @property
    def next_abstract_num_id(self):
        """
        The first ``abstractNumId`` unused by a ``<w:abstractNum>`` element.
        """
        abstractNumId_strs = self.xpath('./w:abstractNum/@w:abstractNumId')
        abstractNumIds = [int(x) for x in abstractNumId_strs]
        for num in range(1, len(abstractNumIds)+2):
            if num not in abstractNumIds:
                break
        return num


class CT_AbstractNum(BaseOxmlElement):
    """
    ``<w:abstractNum>`` element. Defines a base numbering style.
    """
    _tag_seq = (
        'w:nsid', 'w:multiLevelType', 'w:tmpl', 'w:name', 'w:styleLink',
        'w:numStyleLink', 'w:lvl'
    )

    name = ZeroOrOne('w:name', successors=_tag_seq[4:])
    lvl = ZeroOrMore('w:lvl')

    del _tag_seq

    abstractNumId = RequiredAttribute("w:abstractNumId", ST_DecimalNumber)


class CT_Lvl(BaseOxmlElement):
    """
    ``<w:lvl>`` element. Defines both Numbering Level Definition in an
    Abstract Numbering Definition and a Numbering Level Override Definition
    in an Numbering Instance.
    """
    _tag_seq = (
        'w:start', 'w:numFmt', 'w:lvlRestart', 'w:pStyle', 'w:isLgl', 'w:suff',
        'w:lvlText', 'w:lvlPicBulletId', 'w:legacy', 'w:lvlJc', 'w:pPr', 'w:rPr'
    )
    start = ZeroOrOne("w:start", successors=_tag_seq[1:])
    numFmt = ZeroOrOne("w:numFmt", successors=_tag_seq[2:])
    lvlRestart = ZeroOrOne("w:lvlRestart", successors=_tag_seq[3:])
    lvlText = ZeroOrOne("w:lvlText", successors=_tag_seq[7:])
    lvlJc = ZeroOrOne("w:lvlJc", successors=_tag_seq[10:])
    pPr = ZeroOrOne("w:pPr", successors=_tag_seq[11:])
    rPr = ZeroOrOne("w:rPr")

    del _tag_seq

    ilvl = RequiredAttribute('w:ilvl', ST_DecimalNumber)


class CT_LevelText(BaseOxmlElement):
    val = OptionalAttribute("w:val", ST_String)
    null = OptionalAttribute("w:null", ST_OnOff)
