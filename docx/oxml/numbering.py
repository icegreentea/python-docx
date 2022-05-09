# encoding: utf-8

"""
Custom element classes related to the numbering part
"""

from . import OxmlElement
from .shared import CT_DecimalNumber, CT_String
from ..shared import Twips
from .simpletypes import ST_DecimalNumber, ST_MultiLevelType
from .xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrMore, ZeroOrOne
)
from ..enum.text import (
    WD_TAB_ALIGNMENT,
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

    @classmethod
    def new(cls, ilvl, numId):
        numPr = OxmlElement('w:numPr')
        _ilvl = CT_DecimalNumber.new("w:ilvl", ilvl)
        numPr.append(_ilvl)
        _numId = CT_DecimalNumber.new("w:numId", numId)
        numPr.append(_numId)
        return numPr
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
    abstract_num = ZeroOrMore('w:abstractNum', successors=('w:num',
                                                           'w:numIdMacAtCleanup'))
    num = ZeroOrMore('w:num', successors=('w:numIdMacAtCleanup',))

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

    def add_abstract_num(self, abstractNum_id=None, name=None,
                         multiLevelType="multiLevel"):
        """
        Return a newly created CT_AbstractNum (<w:abstractNum>) element with
        abstractNum_id and name.
        """
        if abstractNum_id is None:
            abstractNum_id = self._next_abstractNumId
        else:
            abstractNumId_strs = self.xpath('./w:abstractNum/@w:abstractNumId')
            abstractNum_ids = [int(id_str) for id_str in abstractNumId_strs]
            if abstractNum_id in abstractNum_ids:
                raise ValueError("abstractNum_id already exists")
        abstract_num = CT_AbstractNum.new(abstractNum_id, name, multiLevelType)
        return self._insert_abstract_num(abstract_num)

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
    def _next_abstractNumId(self):
        """
        The first ``abstractNumId`` usused by a ``<w:abstractNum>`` element,
        starting at 1 and filling any gaps in numbering between existing
        ``<w:abstractNum>`` elements.
        """
        abstractNumId_strs = self.xpath('./w:abstractNum/@w:abstractNumId')
        abstractNum_ids = [int(id_str) for id_str in abstractNumId_strs]
        for num in range(1, len(abstractNum_ids)+2):
            if num not in abstractNum_ids:
                break
        return num

    def get_abstract_num_by_id(self, abstract_num_id):
        """
        Return the ``<w:abstractNum>`` child element having ``abstractNumId`` attribute
        matching *abstractNumId*, or |None| if not found.
        """
        xpath = 'w:abstractNum[@w:abstractNumId="%s"]' % abstract_num_id
        try:
            return self.xpath(xpath)[0]
        except IndexError:
            return None

    def get_num_by_id(self, numbering_id):
        """
        Return the ``<w:num>`` child element having ``numId`` attribute
        matching *numbering_id*, or |None| if not found.
        """
        xpath = 'w:num[@w:numId="%s"]' % numbering_id
        try:
            return self.xpath(xpath)[0]
        except IndexError:
            return None

    def get_nums_by_abstract_num_id(self, abstract_num_id):
        """
        Return the ``<w:abstractNum>`` child element having ``abstractNumId``
        attribute matching *abstract_num_id* or |None| if not found.
        """
        xpath = 'w:num[w:abstractNumId[@w:val="%s"]]' % abstract_num_id
        return self.xpath(xpath)

    def get_abstract_num_by_name(self, name):
        """
        Return the ``<w:abstractNum>`` child element having ``name``
        attribute matching *name* or |None| if not found.
        """
        xpath = 'w:abstractNum[w:name[@w:val="%s"]]' % name
        try:
            return self.xpath(xpath)[0]
        except IndexError:
            return None


class CT_MultiLevelType(BaseOxmlElement):
    """
    ``<w:multiLevelType>`` element. Sets the multiLevelType for an
    ``<w:abstractNum>``.

    TBH I'm not really sure what this does since the spec says the
    that it's value does not constrict options.
    """

    val = RequiredAttribute("w:val", ST_MultiLevelType)

    @classmethod
    def MultiLevel(cls):
        elem = OxmlElement('w:multiLevelType')
        elem.val = ST_MultiLevelType.MULTI_LEVEL
        return elem

    @classmethod
    def SingleLevel(cls):
        elem = OxmlElement('w:multiLevelType')
        elem.val = ST_MultiLevelType.SINGLE_LEVEL
        return elem

    @classmethod
    def HybridMultiLevel(cls):
        elem = OxmlElement('w:multiLevelType')
        elem.val = ST_MultiLevelType.HYBRID_MULTI_LEVEL
        return elem


class CT_AbstractNum(BaseOxmlElement):
    """
    ``<w:abstractNum>`` element. Formatting options for a numbering (list) system
    is actually defined here.
    """
    _tag_seq = (
        'w:nsid', 'w:multiLevelType', 'w:tmpl', 'w:name', 'w:styleLink',
        'w:numStyleLink', 'w:lvl'
    )
    # nsid
    multiLevelType = ZeroOrOne('w:multiLevelType', successors=_tag_seq[2:])
    # tmpl
    name = ZeroOrOne('w:name', successors=_tag_seq[4:])
    styleLink = ZeroOrOne('w:styleLink', successors=_tag_seq[5:])
    numStyleLink = ZeroOrOne('w:numStyleLink', successors=_tag_seq[6:])
    levels = ZeroOrMore('w:lvl')
    del _tag_seq
    abstractNumId = RequiredAttribute("w:abstractNumId", ST_DecimalNumber)

    @classmethod
    def new(cls, abstractNum_id, name=None, multiLevelType="multiLevel"):
        """
        Return a new ``<w:abstractNumId>`` element.
        """
        ab_num = OxmlElement('w:abstractNum')
        ab_num.abstractNumId = abstractNum_id
        if name is not None:
            name_elem = CT_String.new("w:name", name)
            ab_num._insert_name(name_elem)
        if multiLevelType == "multiLevel":
            level_elem = CT_MultiLevelType.MultiLevel()
        elif multiLevelType == "singleLevel":
            level_elem = CT_MultiLevelType.SingleLevel()
        elif multiLevelType == "hybridMultiLevel":
            level_elem = CT_MultiLevelType.HybridMultiLevel()
        else:
            raise ValueError("`multiLevelType` must be one of `multiLevel`,"
                             "`singleLevel` or `hybridMultiLevel`")
        ab_num.append(level_elem)

        return ab_num


class CT_Lvl(BaseOxmlElement):
    """
    ``<w:lvl>`` element. Defines the appearance of a single level
    in an abstract numbering scheme. Multiple levels make up a numbering
    scheme.
    """
    _tag_seq = (
        'w:start', 'w:numFmt', 'w:lvlRestart', 'w:pStyle', 'w:isLgl',
        'w:suff', 'w:lvlText', 'w:lvlPicBulletId', 'w:lvlJc',
        'w:pPr', 'w:rPr'
    )
    start = ZeroOrOne('w:start', successors=_tag_seq[1:])
    # if numFmt is omitted, the level shall be assumed to be of type decimal (17.9.17)
    # there are a whole pile of possible values (see ST_NumberFormat)
    # probably should just focus on `decimal` and `bullet` for now
    numFmt = ZeroOrOne('w:numFmt', successors=_tag_seq[2:])

    # if set defines the indent level that when reached will restart
    # this level's counter
    lvlRestart = ZeroOrOne('w:lvlRestart', successors=_tag_seq[3:])

    # names a paragraph style. that paragraph style will be forced
    # to use this numbering scheme, overriding the paragraph's numPr definition
    pStyle = ZeroOrOne('w:pStyle', successors=_tag_seq[4:])
    isLgl = ZeroOrOne('w:isLgl', successors=_tag_seq[5:])
    suff = ZeroOrOne('w:suff', successors=_tag_seq[6:])
    lvlText = ZeroOrOne('w:lvlText', successors=_tag_seq[7:])
    lvlPicBulletId = ZeroOrOne('w:lvlPicBulletId', successors=_tag_seq[8:])
    lvlJc = ZeroOrOne('w:lvlJc', successors=_tag_seq[9:])
    pPr = ZeroOrOne('w:pPr', successors=_tag_seq[10:])
    rPr = ZeroOrOne('w:rPr')
    del _tag_seq

    ilvl = RequiredAttribute("ilvl", ST_DecimalNumber)
    # tplc
    # tentative

    @classmethod
    def create_bullet(cls, ilvl, tabsize_twips=360, lvlText="\u2022", indent_twips=360):
        """
        Returns a <w:lvl> with sensible defaults for bullet lists. *ilvl* sets the
        depth of the level and is 0-indexed. *lvlText* sets the character of the bullet
        and defaults to a unicode bullet characteer. *tabsize_twips* is the size of the
        indent in twips.
        """
        val_key = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
        lvl = OxmlElement('w:lvl')
        lvl.ilvl = ilvl
        numFmt = lvl._add_numFmt()
        numFmt.attrib[val_key] = "bullet"
        _lvlText = lvl._add_lvlText()
        _lvlText.attrib[val_key] = lvlText
        lvlJc = lvl._add_lvlJc()
        lvlJc.attrib[val_key] = "left"
        pPr = create_indented_pPr(tabsize_twips*(ilvl+2), indent_twips)
        lvl._insert_pPr(pPr)
        return lvl

    @classmethod
    def create_decimal(cls, ilvl, tabsize_twips=360, indent_twips=360):
        """
        Returns a <w:lvl> with sensible defaults for decimal list. *ilvl* sets the
        depth of the level and is 0-indexed. *tabsize_twips* is the size of the indent
        in twips.
        """
        val_key = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
        lvl = OxmlElement('w:lvl')
        lvl.ilvl = ilvl
        numFmt = lvl._add_numFmt()
        numFmt.attrib[val_key] = "decimal"
        _start = lvl._add_start()
        _start.attrib[val_key] = "1"
        lvlJc = lvl._add_lvlJc()
        lvlJc.attrib[val_key] = "left"
        lvlText = lvl._add_lvlText()
        lvlText.attrib[val_key] = "%{}.".format(ilvl + 1)

        pPr = create_indented_pPr(tabsize_twips*(ilvl + 2), indent_twips)
        lvl._insert_pPr(pPr)
        return lvl


def create_indented_pPr(indent_twips, hanging_twips):
    pPr = OxmlElement('w:pPr')
    tabstops = pPr._add_tabs()
    indent_emu = Twips(indent_twips).emu
    hanging_emu = Twips(hanging_twips).emu
    tabstops.insert_tab_in_order(indent_emu, WD_TAB_ALIGNMENT.NUM, None)
    _indent = pPr._add_ind()
    _indent.left = indent_emu
    _indent.hanging = hanging_emu
    return pPr


def create_bullet_series(tabsize_twips=360, lvlText="\u2022", indent_twips=360):
    val_key = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
    lvl_lst = []
    for ilvl in range(0, 9):
        lvl = OxmlElement('w:lvl')
        lvl.ilvl = ilvl
        numFmt = lvl._add_numFmt()
        numFmt.attrib[val_key] = "bullet"
        _lvlText = lvl._add_lvlText()
        _lvlText.attrib[val_key] = lvlText
        lvlJc = lvl._add_lvlJc()
        lvlJc.attrib[val_key] = "left"
        pPr = create_indented_pPr(tabsize_twips*(ilvl + 2), indent_twips)
        lvl._insert_pPr(pPr)
        lvl_lst.append(lvl)
    return lvl_lst


def create_decimal_series(tabsize_twips=360, indent_twips=360):
    val_key = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
    lvl_lst = []
    for ilvl in range(0, 9):
        lvl = OxmlElement('w:lvl')
        lvl.ilvl = ilvl
        numFmt = lvl._add_numFmt()
        numFmt.attrib[val_key] = "decimal"
        _start = lvl._add_start()
        _start.attrib[val_key] = "1"
        lvlJc = lvl._add_lvlJc()
        lvlJc.attrib[val_key] = "left"
        lvlText = lvl._add_lvlText()
        lvlText.attrib[val_key] = "%{}.".format(ilvl + 1)

        pPr = create_indented_pPr(tabsize_twips*(ilvl + 2), indent_twips)
        lvl._insert_pPr(pPr)
        lvl_lst.append(lvl)
    return lvl_lst
