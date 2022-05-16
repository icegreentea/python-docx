# encoding: utf-8

from __future__ import absolute_import, division, print_function, unicode_literals
import re
from typing import Type

from docx.shared import ElementProxy, Inches
from docx.text.parfmt import ParagraphFormat


QUARTER_INCH = Inches(0.25).emu

from docx.oxml.numbering import CT_Lvl

_VAL_KEY = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"

class Numbering(ElementProxy):
    """Wrapper around numbering part/element.

    Defines the numbering definitions used in a document.
    Not intended to be constructed directly. Should be retrieved from
    `docx.Document.numbering`.
    """
    def clear_abstract_numbering(self):
        """ Delete all |AbstractNumbering| defintions."""
        self._element.remove_all("w:abstractNum")

    def clear_numbering_instances(self):
        """ Delete all |NumberingInstance|."""
        self._element.remove_all("w:num")

    def remove_abstract_numbering(self, abstract_num, 
                                  remove_numbering_instances=False,
                                  remove_document_refs=False):
        if isinstance(abstract_num, int):
            ab_num = self.get_abstract_numbering_by_id(abstract_num)
            self.remove_abstract_numbering(ab_num, remove_numbering_instances, 
                                           remove_document_refs)
        elif isinstance(abstract_num, str):
            ab_num = self.get_abstract_numbering_by_name(abstract_num)
            self.remove_abstract_numbering(ab_num, remove_numbering_instances, 
                                           remove_document_refs)
        elif isinstance(abstract_num, AbstractNumbering):
            if remove_document_refs == True:
                raise NotImplementedError("Propgating to all document references is not implemented!")    
            if remove_numbering_instances == True:
                num_insts = self.get_numbering_instance_by_abstract_numbering(abstract_num)
                if num_insts is not None:
                    for num_inst  in num_insts:
                        self.remove_numbering_instance(num_inst)
            elem = abstract_num._element
            elem.getparent().remove(elem)
            abstract_num._element = None
        else:
            raise TypeError

    def remove_numbering_instance(self, num, remove_document_refs=False):
        if remove_document_refs == True:
            raise NotImplementedError("Propgating references is not implemented!")
        if isinstance(num, int):
            #self._element.remove_all("w:num[w:numId=%s" % num)
            num_inst = self.get_numbering_instance_by_id(num)
            self.remove_numbering_instance(num_inst)
        elif isinstance(num, NumberingInstance):
            #self._element.remove_all("w:num[w:numId=%s" % num)
            num_elem = num._element
            num_elem.getparent().remove(num_elem)
            num._element = None
            #num_elem._element = None
        else: 
            raise TypeError

    def get_abstract_numbering_by_id(self, abstract_num_id):
        """ Get |AbstractNumbering| by it's id.
        
        Returns None if no match.
        """
        elem = self._element.get_abstract_num_by_id(abstract_num_id)
        if elem is not None:
            return AbstractNumbering(elem, self)

    def get_abstract_numbering_by_name(self, name):
        """ Get |AbstractNumbering| by it's name.

        Returns None if no match.
        """
        elem = self._element.get_abstract_num_by_name(name)
        if elem is not None:
            return AbstractNumbering(elem, self)

    def get_numbering_instance_by_id(self, num_id):
        """ Get |NumberingInstance| by it's *num_id*.

        Returns None if no match.
        """
        elem = self._element.get_num_by_id(num_id)
        if elem is not None:
            return NumberingInstance(elem, self)

    def get_numbering_instance_by_abstract_numbering(self, abstract_num):
        """ Get |NumberingInstance| (s) that match a given |AbstractNumbering|.

        *abstract_num* can be an int (interpreted as abstract_num_id), string 
        (interpreted as name) or instance of |AbstractNumbering|.
        """
        if isinstance(abstract_num, int):
            elems = self._element.get_nums_by_abstract_num_id(abstract_num)
        elif isinstance(abstract_num, str):
            _abstract_num = self._element.get_abstract_num_by_name(abstract_num)
            return self.get_numbering_instance_by_abstract_numbering(_abstract_num.abstractNumId)
        elif isinstance(abstract_num, AbstractNumbering):
            return self.get_numbering_instance_by_abstract_numbering(
                abstract_num.abstract_num_id)
        else:
            raise TypeError

        return [NumberingInstance(e, self) for e in elems]

    @property
    def abstract_numberings(self):
        """ List of all |AbstractNumbering|. """
        return [AbstractNumbering(e, self) for e in self._element.abstract_num_lst]

    @property
    def numbering_instances(self):
        """ List of all |NumberingInstance|. """
        return [NumberingInstance(e, self) for e in self._element.num_lst]

    def create_bullet_abstract_numbering(self, name, tab_width=QUARTER_INCH, 
                                         num_lvls=9, bulletTxt="\u2022"):
        """ Create and return |AbstractNumbering| defining a bullet list (unordered list).

        :param name: The `name` assigned to created abstract numbering.
        :param tab_width: Sets indent tab width, and hanging indent width. 
            Defaults to quarter inch. 
        :param num_lvls: Number of indentation levels to create. Defaults to 9.
        :param bulletTxt: Symbol to use for bullet. Default to `\u2022` (unicode bullet)

        All indentation levels use the same symbol. The created abstract numbering
        has `

        The created abstract numbering will use a number format (`numFmt`) of "bullet".
        """
        new_abstract_numbering = self._element.add_abstract_num(name=name)
        ab_num = AbstractNumbering(new_abstract_numbering)
        for i in range(0, num_lvls):
            ab_num.create_bullet_level(i, indent_step_size=tab_width,
                first_line_indent=-tab_width, lvlText=bulletTxt)
        return ab_num

    def create_decimal_abstract_numbering(self, name, tab_width=QUARTER_INCH, num_lvls=9):
        """
        Create and return |AbstractNumbering| defining a decimal list (ordered list).

        The created abstract numbering will have name of `name`.
        See func
        """
        new_abstract_numbering = self._element.add_abstract_num(name=name)
        ab_num = AbstractNumbering(new_abstract_numbering)
        for i in range(0, num_lvls):
            ab_num.create_decimal_level(i, indent_step_size=tab_width,
                first_line_indent=-tab_width)
        return ab_num


    def create_abstract_numbering(self, name):
        """ Create and return a new |AbstractNumbering| with ``name``.
        """
        new_abstract_numbering = self._element.add_abstract_num(name=name)
        return AbstractNumbering(new_abstract_numbering)

    def create_numbering_instance(self, abstract_num):
        """ Create and return a new |NumberingInstance| connected to ``abstract_num``.

        :params abstract_num: The abstract numbering to reference. If str, will
            be interpeted as abstract numbering name. If int, will interpret as 
            abstract numbering id. Or can be |AbstractNumbering|.

        :returns |NumberingInstance|: The new numbering created.
        """
        if isinstance(abstract_num, str):
            _elem = self._element.get_abstract_num_by_name(abstract_num)
            return NumberingInstance(self._element.add_num(_elem.abstractNumId), parent=self)
        elif isinstance(abstract_num, int):
            return NumberingInstance(self._element.add_num(abstract_num), parent=self)
        elif isinstance(abstract_num, AbstractNumbering):
            return NumberingInstance(
                self._element.add_num(abstract_num.abstract_num_id), parent=self)
        else:
            raise TypeError


class AbstractNumbering(ElementProxy):
    """ Wrapper around ``<w:abstractNum>``.
    """

    @property
    def abstract_num_id(self):
        """ Unique ID - `w:abstractNum/@w:abstractNumId`. """
        return self._element.abstractNumId

    @property
    def name(self):
        """ (Optional) name - `w:abstractNum/w:Name`. """
        return self._element.name.val

    @property
    def levels(self):
        """ Returns listed of defined |AbstractNumberingLevel|.
        """
        return [AbstractNumberingLevel(lvl, self) for lvl in self._element.levels_lst]

    def get_or_add_level(self, ilvl):
        """ Creates or gets |AbstractNumberingLevel| at ``ilvl``.
        """
        _matches = [x for x in self.levels if x.ilvl == ilvl]
        if len(_matches) == 0:
            new_level = self._element.add_levels()
            new_level.ilvl = ilvl
            return AbstractNumberingLevel(new_level, self)
        else:
            return _matches[0]

    def create_bullet_level(self, ilvl, indent_step_size=QUARTER_INCH,
                            first_line_indent=-QUARTER_INCH, lvlText="\u2022"):
        """ Create a |AbstractNumberingLevel| suitable for unordered list.

        :param ilvl: The ``w:ilvl`` to set the level to.
        :param indent_step_size: The tab/indent tab width to use. Default quarter inch.
        :param first_line_indent: The first line indent to use. Default quarter inch.
        :param lvlText: The bullet to use. Default is \u2022.

        :returns: The created |AbstractNumberingLevel|        
        """
        lvl = self.get_or_add_level(ilvl)
        
        lvl.numFmt = "bullet"
        lvl.lvlText = lvlText
        lvl.left_indent = indent_step_size * (ilvl+2)
        lvl.first_line_indent =  first_line_indent
        return lvl

    def create_decimal_level(self, ilvl, indent_step_size=QUARTER_INCH,
                            first_line_indent=-QUARTER_INCH, lvlText=None):
        """ Create a |AbstractNumberingLevel| suitable for ordered list.

        :param ilvl: The ``w:ilvl`` to set the level to.
        :param indent_step_size: The tab/indent tab width to use. Default quarter inch.
        :param first_line_indent: The first line indent to use. Default quarter inch.
        :param lvlText: The bullet to use. Default is |None| which maps to single
            level numbering.

        :returns: The created |AbstractNumberingLevel|        
        """
        if lvlText is None:
            lvlText = "%" + "%s." % (ilvl+1)
        lvl = self.get_or_add_level(ilvl)
        lvl.numFmt = "decimal"
        lvl.lvlText = lvlText
        lvl.left_indent = indent_step_size * (ilvl+2)
        lvl.first_line_indent =  first_line_indent
        lvl.start = 1
        return lvl

class AbstractNumberingLevel(ElementProxy):
    """ Wrapped around ``<w:lvl>`` / :class: docx.oxml.numbering.CT_Lvl
    """
    @property
    def start(self):
        """ Starting Value of numbering.

        Wrapper around ``<w:start>``. This is the value used in for the list
        counter when this level is used for the first time, or when restarted.

        A value of |None| means that no value has been set, and a OOXML 
        specification default of ``0`` will be used.
        """
        if self._element.start is None:
            return None
        return self._element.start.val

    @start.setter
    def start(self, value):
        _start = self._element.get_or_add_start()
        _start.val = value

    @property
    def numFmt(self):
        """ Numbering Format of numbering.

        Used to design the type of numbering to use. Typical values would be
        `decimal` (for ordered list) or `bullet` (for unordered list).
        Wrapper around ``<w:numFnt>``. 

        A value of |None| means that no value has been set and a OOXML
        specification default of ``decimal`` will be used.
        """
        if self._element.numFmt is None:
            return None
        return self._element.numFmt.attrib[_VAL_KEY]

    @numFmt.setter
    def numFmt(self, value):
        _numFmt = self._element.get_or_add_numFmt()
        _numFmt.attrib[_VAL_KEY] = value

    @property
    def lvlRestart(self): 
        """ Restart Numbering Level Symbol
        
        Wrapper around ``<w:lvlRestart>``. Sets when this level will restart 
        it's counter. 1 indexed. Whenever a level of `lvlRestart` (or higher) is 
        encountered, the counter will restart. Setting to 0 will cause this level 
        to never reset.

        A value of |None| means that no value has been set and a OOXML 
        sepcification default of restart anytime a previous level or higher is 
        used will be used.
        """
        if self._element.lvlRestart is None:
            return None
        return self._element.lvlRestart.val

    @lvlRestart.setter
    def lvlRestart(self, value):
        _lvlRestart = self._element.get_or_add_lvlRestart()
        _lvlRestart.val = value

    @property
    def paragraph_style(self):
        """
        Name of |_ParagraphStyle| set as `pStyle`. 
        |None| indicates no value has been set.
        Named paragraph style will automatically bind to this numbering level (ilvl),
        BUT NOT to the number id - the paragraph still needs to set ``numId`` to
        reference a concrete numbering instances.
        """
        if self._element.pStyle is None:
            return None
        return self._element.pStyle[_VAL_KEY]

    @paragraph_style.setter
    def paragraph_style(self, value):
        _pStyle = self._element.get_or_add_pStyle()
        if isinstance(value, str):
            print("ping")
            _pStyle.attrib[_VAL_KEY] = value
        else:
            from .styles.style import _ParagraphStyle
            if isinstance(str, _ParagraphStyle):
                _pStyle.attrib[_VAL_KEY] = _ParagraphStyle.name
            else:
                raise TypeError

    @property
    def is_legal(self):
        """
        Is "legal" numbering format (show all decimals at all numbering levels)
        enabled? 
        """
        if self._element.isLgl is None:
            return False
        return True

    @is_legal.setter
    def is_legal(self, value):
        if value:
            self._element.get_or_add_isLgl()
        else:
            self._element._remove_isLgl()


    @property
    def suffix(self):
        """
        Returns the content between numbering symbol and paragraph text.
        If |None|, it means no value is set, and OOXML specification default
        of "tab" is assumed.

        Otherwise must be one of "tab", "space" or "nothing"
        """
        if self._element.suff is None:
            return None
        return self._element.suff[_VAL_KEY]

    @suffix.setter
    def suffix(self, value):
        if value is None:
            self.element._remove_suff()
        else:
            value = value.lower()
            if value not in ("tab", "space", "nothing"):
                raise ValueError("Suffix must be one of `tab`, `space` or `nothing`")
            _suff = self._element.get_or_add_suff()
            _suff.attrib[_VAL_KEY] = value

    @property
    def lvlText(self): 
        """ Numerbing Level Text.

        Wrapper around ``<w:lvlText>``. Sets the text content to be displayed
        as that "bullet" or "list" point.

        Can be set as a single character (for example for bullet lists).
        Can use percent symbol (%) followed by a number to insert in index counter.
        
        ``"%1."`` for example would mean use the counter value for the 1st 
        numbering level.

        For example, imagine a list like::

            1. First item
                a. First-First item
                b. First-Second item
            2. Second item
                a. Second-first item

        Using ``lvlText`` of ``%1.`` for ``ilvl==1`` and ``%2.`` for ``ilvl==2``
        would yield::

            1. First item
                1. First-First item
                2. First-Second item
            2. Second item
                1. Second-first item

        Using ``lvlText`` of ``%1.`` for ``ilvl==1`` and ``%1.%2.`` for 
        ``ilvl==2`` would yield::

            1. First item
                1.1. First-First item
                1.2. First-Second item
            2. Second item
                2.1. Second-first item
        """
        if self._element.lvlText is None:
            return None
        return self._element.lvlText.attrib[_VAL_KEY]

    @lvlText.setter
    def lvlText(self, value):
        _lvlTxt = self._element.get_or_add_lvlText()
        _lvlTxt.attrib[_VAL_KEY] = value

    @property
    def justification(self):
        if self._element.lvlJc is None:
            return None
        return self._element.lvlJc.val

    @justification.setter
    def justification(self, value):
        if value is None:
            self._element._remove_lvlJc()
        else:
            self._element.lvlJc.val = value


    @property
    def paragraph_format(self):
        """ |ParagraphFormat|. Numbering Level Associated Paragraph Properties.
        
        Wrapper around ``<w:pPr>``. 
        Defines the paragraph properties associated with this numbering level.
        Paragraph properties set on the actual paragraph will override properties
        set here.

        Tabs (indentation) is defined within here.
        """
        pPr = self._element.get_or_add_pPr()
        return ParagraphFormat(self._element, self)

    @property
    def run_format(self):
        """ 
        Creates or gets Numerbing Symbol Run properties.

        Creates or gets instance of :class:`docx.oxml.text.font.CT_RPr`. This
        sets the styling of the numbering level text when applied. For example,
        you can use this to set the font size of the numbering component to a 
        specific value.
        """
        rPr = self._element.get_or_add_rPr()
        return rPr

    @property
    def ilvl(self): 
        """ Numbering Level Reference.

        Wrapper around ``<w:ilvl>``. Defines the numbering level.
        Used by ``<w:num>`` / |NumberingInstance| to reference a given numbering
        to apply.
        """
        return self._element.ilvl

    @ilvl.setter
    def ilvl(self, value):
        self._element.ilvl = value

    @property
    def left_indent(self):
        """
        |Length| value specifying the space between the left margin and the
        left side of the paragraph. |None| indicates the left indent value is
        inherited from the style hierarchy. Use an |Inches| value object as
        a convenient way to apply indentation in units of inches.

        Wrapper around inner |ParagraphFormat| object's property.
        """
        return self.paragraph_format.left_indent

    @left_indent.setter
    def left_indent(self, value):
        self.paragraph_format.left_indent = value
    
    @property
    def right_indent(self):
        """
        |Length| value specifying the space between the right margin and the
        right side of the paragraph. |None| indicates the right indent value
        is inherited from the style hierarchy. Use a |Cm| value object as
        a convenient way to apply indentation in units of centimeters.

        Wrapper around inner |ParagraphFormat| object's property.
        """
        return self.paragraph_format.right_indent

    @right_indent.setter
    def right_indent(self, value):
        self.paragraph_format.right_indent = value

    @property
    def first_line_indent(self):
        """
        |Length| value specifying the relative difference in indentation for
        the first line of the paragraph. A positive value causes the first
        line to be indented. A negative value produces a hanging indent.
        |None| indicates first line indentation is inherited from the style
        hierarchy.

        Wrapper around inner |ParagraphFormat| object's property.
        """
        return self.paragraph_format.first_line_indent

    @first_line_indent.setter
    def first_line_indent(self, value):
        self.paragraph_format.first_line_indent = value

    

        


class NumberingInstance(ElementProxy):
    """
    Wrapper around ``<w:num>``. Concrete numbering implementation of an 
    abstract numbering defintion (as represented by |AbstractNumbering|).
    Paragraph styles (|ParagraphFormat|) gain numbering style by referencing
    ``numId`` of specific numbering instances.

    Numbering instaces may override the underlying |AbstractNumbering| by
    setting ``ilvl_overrides``. These overrides occur on a per level basis.
    """
    @property
    def ilvl_overrides(self):
        pass

    @property
    def numId(self):
        """ `w:num/@w:numId`. Used as unique identifier for numbering instance"""
        return self._element.numId

    @property
    def abstract_num_id(self):
        """ The ``w:abstractNumId`` of the base |AbstractNumbering|. """
        return self._element.abstractNumId.val

    @abstract_num_id.setter
    def abstract_num_id(self, value):
        self._element.abstractNumId.val = value

    @property
    def abstract_num(self):
        """ Helper reference to base |AbstractNumbering|. """
        return self._parent.get_abstract_numbering_by_id(self.abstract_num_id)

    @abstract_num.setter
    def abstract_num(self, abstract_num):
        self.abstract_num_id = abstract_num.abstract_num_id