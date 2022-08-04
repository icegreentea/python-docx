.. _understanding_numbering:

Understanding Numbering
=======================

The word document standard uses three different parts to implement numbering and lists 
in general - unordered lists are just a special case of a numbered list. The three parts
are:

Abstract Numbering Definition
    This is basically a "base" numbering style. The document standard does not allow 
    paragraphs to reference an abstract numbering definition directly - they must 
    reference a numbering instance.

Numbering Instance
    A "concrete" numbering definition instance. Each numbering instance references
    an abstract numbering definition (and has the ability to override aspects of it).
    You need a new numbering instance to restart the numbering - ie for each new list
    where you want to restart from 1), you'll need a new numbering instance.

Paragraphs referencing numbering instances
    Paragraphs reference a "concrete" numbering definition to become a list element.

Any given numbering definition defines a number (up to 9) numbering level definitions.
Each numbering level definition essentially corresponds to an "indentation" level.
A numbering instance can override any number of numbering level definitions of its
parent abstract numbering defintion.

Conceptual Map
--------------

- Paragraph references a numbering instance and numbering level
- Numbering instance references an abstract numbering definition
    - A numbering instance and override any number of numbering level definitions
- Abstract numbering definition contains a number of numbering level definitions
    - Each numbering level definition defines the paragraph and labelling formatting for
      a numbering/indentation level

Key Identifiers
---------------

A paragraph needs to reference a numbering instance (via its numbering instance ID) and
an indentation level to properly acquire the right style.

A numbering instance needs to reference an abstract numbering definition ID to properly
inherit its parent style. 

Common Styling Elements
-----------------------

* You can style the format/style of the body paragraph.
* The labelling text can be styled as a run. For example you can make the "1." in 
  "**1.** First list element" bold.
* You can change the numbering format of the labelling text. For example using bullets,
  or decimals, or roman numerals.
* You can change text of the labelling text. For example if using decimals, you can 
  change between something like "1." and "1)"

Because of all of these styling elements are defined at the numbering level definition, 
you can mix and match these at different levels. 

Types of Numbering Formats
--------------------------

The word document format supports a wide range of different numbering formats. A few 
common formats are listed below.

decimal
    Decimal numbers (1, 2, 3,...)
upperRoman
    Uppercase roman numbers (I, II, III, ...)
lowerRoman
    Lowercase roman numbers (i, ii, iii, ...)
upperLetter
    Uppercase latin alphabet (A, B, C, ...)
lowerLetter
    Lowercase latin alphabet (a, b, c, ...)
bullet
    Use bullets

For a more complete listing, look at the definition of ``w:ST_NumberFormat``

Numbering Text
--------------

Seperate from the "type" of numbering that numbering format gives you, you can control
the exact textual content of the label using "numbering text". For example, if you used
a "decimal" numbering format, this could let you choose between "1." and "1)" as a 
label.

You can also use this to create fully defined labels such as::

    1. First element
    2. Second element
        2.1 Another element
        2.2 Another element
    3. Third element
        3.1 Another element
            3.1.1 Another element

The numbering text format uses ``%X`` as a form of string interpolation. All other 
characters are displayed litterally. In ``%X`` the ``X`` is interpreted as a one-indexed
reference to an indentation/numbering level. The current count of that 
indentation/numbering level is then substituted in.

So for example, for a simple list like::

    1. First element
    2. Second element
        1. Another element

You would use numbering text of ``%1.`` for the first level nad ``%2.`` for the second
numbering level.

For a simple list like::

    1) First element
    2) Second element
        1) Another element

You would use numbering text of ``%1)`` for the first level nad ``%2)`` for the second
numbering level.

For a more complicated list like::

    1. First element
    2. Second element
        2.1. Another element

You would use numbering text of ``%1.`` for the first level nad ``%1.%2.`` for the 
second numbering level.