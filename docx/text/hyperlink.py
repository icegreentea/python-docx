from docx.shared import RGBColor
from ..runcntnr import RunItemContainer
from .run import Run


class Hyperlink(RunItemContainer):
    """
    Proxy object wrapping ``<w:hyperlink>`` element.
    """
    def __init__(self, element, parent):
        super(Hyperlink, self).__init__(element, parent)

    @property
    def anchor(self):
        return self._element.anchor

    @anchor.setter
    def anchor(self, value):
        self._element.anchor = value

    def clear(self):
        self._element.clear_content()
        return self

    @property
    def is_external(self):
        _id = self._element.id
        if _id is None:
            return False
        else:
            return True

    @property
    def target(self):
        if self.is_external:
            return self.relationship_id
        else:
            return self.anchor

    @property
    def relationship_id(self):
        return self._element.id

    @relationship_id.setter
    def relationship_id(self, value):
        self._element.id = value

    @property
    def runs(self):
        """
        Sequence of |Run| instances corresponding to ``<w:r>`` elements in
        this hyperlink.
        """
        return [Run(r, self) for r in self._element.r_lst]

    @classmethod
    def add_hyperlink_styles(cls, document, link_color="0563C1"):
        """
        Helper method that adds a realized and suitable "Hyperlink" character style
        to the document if it does not already exist.
        """
        from docx.enum.style import WD_STYLE_TYPE
        from docx.enum.dml import MSO_THEME_COLOR
        styles = document.styles
        try:
            hyperlink_style = styles.get_style_id("Hyperlink", WD_STYLE_TYPE.CHARACTER)
        except (ValueError, KeyError):
            _style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
            _style.base_style = styles["Default Paragraph Font"]
            _style.font.color.rgb = RGBColor.from_string(link_color)
            _style.font.color.theme_color = MSO_THEME_COLOR.HYPERLINK
            _style.font.underline = True
        else:
            pass

    def update_external_target(self, new_target, document):
        """
        Update the external target in document relationships. Use to change what
        external hyperlink points to for example.
        """
        if self.relationship_id is None:
            raise ValueError("Cannot update a relationship without having a"
                             " relationship id.")
        document._part.rels._update_external_rel_target(self.relationship_id,
                                                        new_target)
        