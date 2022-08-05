from collections.abc import Sequence
from copy import deepcopy

from docx.shared import ElementProxy, Inches, Emu


class NumberingInstance:
    """
    Wrapper/proxy around ``<w:num>`` element. Represents an instance of a numbering
    or list. In general, you will need a new |NumberingInstance| for each new list that
    you have to restart the counter.

    Requires handles to it's wrapped ``<w:num>`` element, it's parent ``<w:numbering>``
    part, and the overall document - since we use this object to add additional
    list items/paragraphs.

    *start_override* (default 1) will automatically create an ``<w:lvlOverride>``
    override for the first numbering level with a ``<w:startOverride>`` value
    set to *start_override*. The default value of 1 means that each instance
    of |NumberingInstance| referncing the same abstract numbering definition restarts
    it's numbering count. If *start_override* is set to |None|, then none of the 
    overrides will be present.
    """
    def __init__(self, numbering_element, numbering_part, document_part,
                 start_override=1):
        from docx.document import Document
        self._numbering_part = numbering_part
        self._doc_part = document_part
        self._element = numbering_element
        self._doc = Document(self._doc_part.element, self._doc_part)
        if start_override is not None:
            override = self.add_level_override(0)
            override.start_override = start_override

    @property
    def numbering_id(self):
        return self._element.numId

    @property
    def level_overrides(self):
        """
        Sequence of defined |NumberingLevelOverride| objects.
        """
        return [NumberLevelOverride(x, self) for x in self._element.lvlOverride_lst]

    def get_level_override_by_ilvl(self, ilvl):
        """
        Return |NumberingLevelOverride| with matching numbering level of *ilvl*.
        If no override exists with matching ilvl, returns |None|.
        """
        for x in self._element.lvlOverride_lst:
            if x.ilvl == ilvl:
                return NumberLevelOverride(x, self)
        return None

    def add_level_override(self, ilvl):
        """
        Return (create if needed) a |NumberingLevelOverride| with numbering level of
        *ilvl*.
        """
        existing = self.get_level_override_by_ilvl(ilvl)
        if existing:
            return existing
        else:
            lvlOverride_elm = self._element.add_lvlOverride(ilvl)
            numlvloverride = NumberLevelOverride(lvlOverride_elm, self)
            numlvloverride.start_override = 1
            return numlvloverride

    def add_paragraph(self, indent_level, text=''):
        """
        Create and return a new |Paragraph|. *indent_level* is zero-indexed value
        representing the level of indentation.
        """
        para = self._doc.add_paragraph(text)
        p_elm = para._element
        pPr = p_elm.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        ilvl = numPr.get_or_add_ilvl()
        ilvl.val = indent_level

        numId = numPr.get_or_add_numId()
        numId.val = self._element.numId

        return para

    def add_unlabeled_paragraph(self, indent_level, text=''):
        pass

    def _get_indentation(self, indent_level):
        pass


class AbstractNumberingDefinition(ElementProxy, Sequence):
    """
    Wrapper/proxy around ``<w:abstractNum>`` element.
    This element describes an "abstract" or "base" numbering format.
    The element consists of a number of |NumberingLevelDefinition|, each of which
    defines the formatting at a particularing numbering/indentation level.
    """
    @property
    def name(self):
        return self._element.name.val

    @name.setter
    def name(self, value):
        self._element.name.val = value

    @property
    def abstract_num_id(self):
        return self._element.abstractNumId

    def __getitem__(self, key):
        if isinstance(key, slice):
            return [
                NumberingLevelDefinition(lvl, self._element)
                for lvl in self._element.lvl_lst[key]
            ]
        return NumberingLevelDefinition(self._element.lvl_lst[key], self._element)

    def __iter__(self):
        for lvl in self._element.lvl_lst:
            yield NumberingLevelDefinition(lvl, self._element)

    def __len__(self):
        return len(self._element.lvl_lst)

    def set_level_number_format(self, number_format):
        """
        Set the ``number_format`` property (``<w:numFmt>`` element) of child
        |NumberingLevelDefinition|. If *number_format* is a string, then the same
        value will be applied to all child elements. Otherwise, will iterate over
        *number_format* and apply each sub-element to its matching child element
        in sequence.
        """
        if isinstance(number_format, str):
            for i, lvl in enumerate(self):
                lvl.number_format = number_format
        else:
            for i, lvl in enumerate(self):
                lvl.number_format = number_format[i]
        return self

    def set_level_text(self, numbering_level_text):
        """
        Set the ``numbering_level_text`` property (``<w:lvlText)`` element) of
        child |NumberingLevelDefinition|. If *numbering_level_text* is a string,
        then the same value will be applied to all child elements. Otherwise, will
        iterate over *numbering_level_text* and apply each sub-element to its matching
        child element in sequence.
        """
        if isinstance(numbering_level_text, str):
            for i, lvl in enumerate(self):
                lvl.numbering_level_text = numbering_level_text
        else:
            for i, lvl in enumerate(self):
                lvl.numbering_level_text = numbering_level_text[i]
        return self

    def set_level_start(self, start):
        """
        Set the ``start`` property (``<w:start)`` element) of
        child |NumberingLevelDefinition|. If *start* is a int,
        then the same value will be applied to all child elements. Otherwise, will
        iterate over *start* and apply each sub-element to its matching
        child element in sequence.
        """
        if isinstance(start, int):
            for i, lvl in enumerate(self):
                lvl.start = start
        else:
            for i, lvl in enumerate(self):
                lvl.start = start[i]
        return self

    @classmethod
    def initialize_element(cls, abstractNum_elem, name=None,
                           hanging_indent=Inches(0.25),
                           leading_indent=Inches(0.5),
                           tabsize=Inches(0.25), levels=9,
                           abstract_num_id=None):
        """
        Create, initialize and return an |AbstractNumberingDefinition| object.
        *abstractNum_elem* should be a ``<w:abstractNum>`` element with no child
        ``<w:lvl>`` elements defined.

        *hanging_indent* is the additional indent used on body text after the first
        line. Use of *hanging_indent* allows the start margin of body text to be aligned
        across multiple lines.
        *leading_indent* is the indent from document start margin to start marign of
        body text on the first line. It is NOT the indent to the list marker.
        *tabsize* is the additional indent to be applied for each additional numbering
        level.

        If *abstract_num_id* is provided, will override any existing ``w:abstractNumId``
        attribute on *abstractNum_elem*.

        *levels* is the number of child ``<w:lvl>`` elements to create. The maximum is
        9.
        """
        if abstract_num_id is not None:
            abstractNum_elem.abstractNumId = abstract_num_id
        if name is not None:
            _name = abstractNum_elem.get_or_add_name()
            _name.val = name
        for i in range(0, levels):
            lvl = abstractNum_elem.add_lvl()
            lvl.ilvl = i
            pPr = lvl.get_or_add_pPr()
            indent = pPr.get_or_add_ind()
            indent.left = Emu(leading_indent).emu + i * Emu(tabsize).emu
            indent.hanging = Emu(hanging_indent).emu
            start = lvl.get_or_add_start()
            start.val = 1
        return cls(abstractNum_elem)

    @staticmethod
    def alternate_bullet_definition():
        """
        Returns numbering format and numbering level text (as tuple) suitable for bullet
        list that alternates between solid dot, hollow dot, and square bullets.
        """
        return "bullet", [
            '\u2022', '\u25CB', '\u25aa',
            '\u2022', '\u25CB', '\u25aa',
            '\u2022', '\u25CB', '\u25aa',
        ]

    @staticmethod
    def simple_bullet_definition():
        """
        Returns numbering format and numbering level text (as tuple) suitable for 
        bullet list of solid dots.
        """
        return "bullet", '\u2022'

    @staticmethod
    def simple_decimal_definition():
        """
        Returns numbering format and numbering level text (as tuple) suitable for
        a decimal list where each level is labelled with "X.". For example::

            1.
            2.
                1.
                2.
        """
        return "decimal", ["%{}.".format(x+1) for x in range(9)]

    @staticmethod
    def simple_bracket_decimal_definition():
        """
        Returns numbering format and numbering level text (as tuple) suitable for
        a decimal list where each level is labelled with "X)". For example::

            1)
            2)
                1)
                2)
        """
        return "decimal", ["%{})".format(x+1) for x in range(9)]

    @staticmethod
    def fully_defined_decimal_definition():
        """
        Returns numbering format and numbering level text (as tuple) suitable for
        a decimal list where each level is fully defined - including parent level. 
        For example::

            1.
            2.
                2.1.
                2.2
        """
        _store = ["%1."]
        for i in range(1, 9):
            _store.append(_store[-1] + "%{}.".format(i+1))
        return "decimal", _store


class NumberLevelOverride(ElementProxy):
    """
    Wrapper around ``<w:lvlOverride>`` element.
    """
    def __init__(self, element, parent=None):
        super().__init__(element, parent)

    @property
    def numbering_level(self):
        return self._element.ilvl

    @numbering_level.setter
    def numbering_level(self, value):
        self._element.ilvl = value

    @property
    def start_override(self):
        startOverride = self._element.startOverride
        if startOverride is None:
            return None
        return startOverride.val

    @start_override.setter
    def start_override(self, value):
        if value is None:
            self._element._remove_startOverride()
        else:
            startOverride = self._element.get_or_add_startOverride()
            startOverride.val = value

    @property
    def override_level_definition(self):
        lvl = self._element.lvl
        if lvl is None:
            return None
        return NumberingLevelDefinition(lvl, self)

    def create_override_level_definition(self):
        lvl = self._element.get_or_add_lvl()
        return NumberingLevelDefinition(lvl, self)

    def remove_override_level_definition(self):
        self._element._remove_lvl()


class NumberingLevelDefinition(ElementProxy):
    """
    Wrapper around ``<w:lvl>`` element.
    Defines the formatting and behavior of a single numbering/indentation level
    in an abstract numbering definition, or in a numbering instance level override.
    """
    @property
    def start(self):
        """
        The numbering value to start at for a given level. If undefined, will default
        to 0. Wrapper around ``<w:start>`` element.
        """
        start = self._element.start
        if start is None:
            return None
        return start.val

    @start.setter
    def start(self, value):
        start = self._element.get_or_add_start()
        start.val = value

    @property
    def number_format(self):
        """
        The numbering format to use. If undefined, will default to "decimal". Wrapper
        around ``<w:numFmt>`` element.

        Common choices include: "decimal", "bullet", "lowerLetter", "lowerRoman",
        "none", "upperLetter", "upperRoman".

        See defintions of ``ST_NumberFormat`` for more detailed explanations.
        """
        numFmt = self._element.numFmt
        if numFmt is None:
            return None
        return numFmt.val

    @number_format.setter
    def number_format(self, value):
        numFmt = self._element.get_or_add_numFmt()
        numFmt.val = value

    @property
    def numbering_level(self):
        """
        The numbering level. Wrapper around ``w:ilvl`` attribute. Zero indexed.
        """
        return self._element.ilvl

    @property
    def restart_numbering_level(self):
        """
        Wrapper around ``<w:lvlRestart>`` element. Determines when this numbering level
        should restart. One indexed.

        When a numbering level of *restart_numbering_level* occurs in this document,
        this numbering level will reset to it's ``start`` value on next occurence.
        *restart_numbering_level* should be higher than this numbering level.

        For example, assuming that on the second numbering level, we set
        *restart_numbering_level* to 3. Then we would expect this type of behavior::

            1.
            2.
                1.
                2.
                    1.
                1. (reset due to 3rd level appearing)

        When unset (the default), this level will reset whenever a LOWER numbering
        level occurs (normal expected behavior)::

                1.
                2.
                    1.
                3.
                    1. (reset due to 1st level appearing)
        """
        lvlRestart = self._element.lvlRestart
        if lvlRestart is None:
            return None
        return lvlRestart.val

    @restart_numbering_level.setter
    def restart_numbering_level(self, value):
        if value is None:
            self._element._remove_lvlRestart()
        else:
            lvlRestart = self._element.get_or_add_lvlRestart()
            lvlRestart.val = value

    @property
    def numbering_level_text(self):
        """
        Numbering Level Text. Wrapper around ``<w:lvlText>`` element.
        Defines the text content used in numbering. The "1." or specific bullet
        symbol for example.

        Use ``%X`` where X is a one indexed reference to a numbering level. For
        example if ``numbering_level_text`` for the 3rd numbering level was set to
        ``%1.%2-%3)`` then you would see something like::

            1.
            2.
                1.
                    2.1-1)
                    2.1-2)
                2.
                    2.2-1)
                    2.2-2)
        """

        lvlText = self._element.lvlText
        if lvlText is None:
            return None
        return lvlText.val

    @numbering_level_text.setter
    def numbering_level_text(self, value):
        lvlText = self._element.get_or_add_lvlText()
        lvlText.val = value

    @property
    def justification(self):
        """Justification of the numbering element."""
        lvlJc = self._element.lvlJc
        if lvlJc is None:
            return None
        return lvlJc.val

    @justification.setter
    def justification(self, value):
        lvlJc = self._element.get_or_add_lvlJc()
        lvlJc.val = value

    @property
    def paragraph_properties(self):
        """
        |ParagraphProperty| wrapping the ``<w:pPr>`` element of this numbering level.
        Defines the paragraph property of the body text of this numbering level.
        """
        from docx.text.parfmt import ParagraphFormat

        pPr = self._element.pPr
        if pPr is None:
            return None

        return ParagraphFormat.numbering_level_wrapper(pPr)

    @paragraph_properties.setter
    def paragraph_properties(self, value):
        """
        Sets the numbering level associated paragraph properties for this
        numbering level. These properties set the formatting for paragraphs (body text)
        at this numbering level.

        Sets the properties by creating a clone of *value* and setting the clone,
        overwriting any existing values. Pass in |None| to unset this property
        and fall back to specification defaults.
        """
        from docx.text.parfmt import ParagraphFormat
        from docx.oxml.text.parfmt import CT_PPr
        if value is None:
            self._element.set_pPr(None)
        elif isinstance(value, ParagraphFormat):
            val = deepcopy(value._element.pPr)
            self._element.set_pPr(val)
        elif isinstance(value, CT_PPr):
            self._element.set_pPr(deepcopy(value))

    def create_new_paragraph_properties(self):
        """
        Clears any existing paragraph properties, and creates and returns a new
        |ParagraphFormat| for this numbering level.
        """
        from docx.text.parfmt import ParagraphFormat

        self._element._remove_pPr()
        pPr = self._element.get_or_add_pPr()
        return ParagraphFormat.numbering_level_wrapper(pPr)
