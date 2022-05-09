# encoding: utf-8

from __future__ import absolute_import, division, print_function, unicode_literals

from docx.shared import ElementProxy

from docx.oxml.numbering import CT_Lvl


class Numbering(ElementProxy):

    def clear_abstract_numbering(self):
        self._element.remove_all("w:abstractNum")

    def clear_numbering_instances(self):
        self._element.remove_all("w:num")

    def get_abstract_numbering_by_id(self, abstract_num_id):
        elem = self._element.get_abstract_num_by_id(abstract_num_id)
        if elem is not None:
            return AbstractNumbering(elem, self)

    def get_numbering_instance_by_id(self, num_id):
        elem = self._element.get_num_by_id(num_id)
        if elem is not None:
            return NumberingInstance(elem, self)

    def get_numbering_by_abstract_numbering(self, abstract_num):
        if isinstance(abstract_num, str):
            elems = self._element.get_nums_by_abstract_id(abstract_num)
        elif isinstance(abstract_num, int):
            _abstract_num = self._element.get_abstract_num_by_name(abstract_num)
            return self.get_numbering_by_abstract_numbering(_abstract_num.abstractNumId)
        elif isinstance(abstract_num, AbstractNumbering):
            return self.get_numbering_by_abstract_numbering(
                abstract_num.abstract_num_id)
        else:
            raise TypeError

        return [NumberingInstance(e, self) for e in elems]

    @property
    def abstract_numberings(self):
        return [AbstractNumbering(e, self) for e in self._element.abstract_num_lst]

    @property
    def numbering_instances(self):
        return [NumberingInstance(e, self) for e in self._element.num_lst]

    def create_bullet_abstract_numbering(self, name, tab_width_twips=360):
        new_abstract_numbering = self._element.add_abstract_num(name=name)
        levels = [CT_Lvl.create_bullet(i) for i in range(0, 9)]
        for level in levels:
            new_abstract_numbering._insert_levels(level)
        return AbstractNumbering(new_abstract_numbering)

    def create_decimal_abstract_numbering(self, name, tab_width_twips=360):
        new_abstract_numbering = self._element.add_abstract_num(name=name)
        levels = [CT_Lvl.create_decimal(i) for i in range(0, 9)]
        for level in levels:
            new_abstract_numbering._insert_levels(level)
        return AbstractNumbering(new_abstract_numbering)

    def create_abstract_numbering(self, name):
        new_abstract_numbering = self._element.add_abstract_num(name=name)
        return AbstractNumbering(new_abstract_numbering)

    def create_numbering_instance(self, abstract_num):
        if isinstance(abstract_num, str):
            _elem = self._element.get_abstract_num_by_name(abstract_num)
            return NumberingInstance(self._element.add_num(_elem.abstract_num_id))
        elif isinstance(abstract_num, int):
            return NumberingInstance(self._element.add_num(abstract_num))
        elif isinstance(abstract_num, AbstractNumbering):
            return NumberingInstance(
                self._element.add_num(abstract_num.abstract_num_id))
        else:
            raise TypeError


class AbstractNumbering(ElementProxy):
    @property
    def abstract_num_id(self):
        return self._element.abstractNumId

    @property
    def name(self):
        return self._element.name.val


class AbstractNumberingLevel(ElementProxy):
    pass


class NumberingInstance(ElementProxy):
    @property
    def ilvl(self):
        return self._element.ilvl.val

    @property
    def numId(self):
        return self._element.numId
