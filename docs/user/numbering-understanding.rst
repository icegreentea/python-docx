.. _numbering_styles:

Understanding Numbering
=======================

How does Numbering Work?
------------------------

Word documents can define any number of numbering (list) formats. These numbering
formats can then by applied to different paragraphs to get different numbering.
Both "ordered" (numbered) and "unordered" (bullet) lists are implemented using the 
same system. 

There are three main entities that we need to keep track of here: abstract numberings,
numbering instances, and paragraphs.

Abstract numberings are the root definition of what a given numbering/list scheme looks 
like. Numbering instances form a concrete reference, and each numbering instance 
references an abstract numbering. Finally paragraphs reference a numbering instasnce
to apply that numbering to a paragraph.

It's important to note that a fully empty word document does not have any 
numbering defined by default. When using Microsoft Word or LibreOffice, the 
default numbering styles are built into the editor, which are then serialized/saved
into the word file.

The network of relevant XML elements is roughly:

::

    numbering part (numbering.xml)
        w:numbering
            w:abstractNum w:abstractNumId="1" 
                w:name w:val="list" 
                w:lvl w:ilvl="0"
                    w:numFmt
                    w:pPr
                w:lvl w:ilvl="1"
                w:lvl w:ilvl="2"
                ...
                w:lvl w:ilvl="8"
                
            w:num w:numId="1"
                w:abstractNumId="1"
    document part (document.xml)
        w:document
            w:body
                w:p
                    w:pPr
                        w:numPr
                            w:ilvl="0"
                            w:numId="1"

How Does Restarting Numbering Work?
-----------------------------------

Restarting a numbering sequence works by creating a new numbering instance of the same
base abstract numbering. Therefore, each time we start a new numbering sequence, we
have to create a new numbering instance.

For example, imagine a document that looks like:

::

    Body Paragraph 1
    1. List 1-1
        1.1 List 1-2
    2. List 1-3
    Body Paragraph 2
    1. List 2-1
        1.1 List 2-2
    2. List 2-2

This would be implemented with something like:

::

    w:numbering
        w:abstractNum w:abstractNumId="1"
        w:num w:numId="1"
            w:abstractNumId="1"
        w:num w:numId="2"
            w:abstractNumId="1"
    w:document
        w:body
            w:p // "Body Paragraph 1"
            w:p // "List 1-1"
                w:pPr/w:numPr
                    w:ilvl="0"
                    w:numId="1"
            w:p // "List 1-2"
                w:pPr/w:numPr
                    w:ilvl="1"
                    w:numId="1"
            w:p // "List 1-3"
                w:pPr/w:numPr
                    w:ilvl="0"
                    w:numId="1"
            w:p // "Body Paragraph 2"
            w:p // "List 2-1"
                w:pPr/w:numPr
                    w:ilvl="0"
                    w:numId="2"
            w:p // "List 2-2"
                w:pPr/w:numPr
                    w:ilvl="1"
                    w:numId="2"
            w:p // "List 2-3"
                w:pPr/w:numPr
                    w:ilvl="0"
                    w:numId="2"

How Does Formatting Work?
-------------------------

Formatting belongs to the Numbering Level Definition (``<w:lvl>``). There are two types
of formatting that can be applied - paragraph level which is applied to the paragraph using
the numbering, and run level which is applied to the numbering label/text (for example ``3.a.b``).

Paragraph level (formally the Numbering Level Associated Paragraph Properties) override 
existing paragraph properties on any numbered paragraphs that reference the given numbering
instance and level.

How Does Overriding Work?
-------------------------



How Does Numbered Headings Work?
--------------------------------

How To Format Numbering?
------------------------



Glossary
--------

Abstract Numbering Definition
    A ``<w:abstractNum>`` element in the numbering part of a document that 
    defines the attributes of one numbering scheme. Paragraphs in a document 
    reference numbering instances, NOT abstract numberings.
    
    Within an abstract nuimbering is a collection (list) of ``<w:lvl>``/Numbering Level 
    Definition  instances. Each numbering level definition defines the behavior and
    format for a given nesting/indentation level.

Numbering Level Definition
    A ``<w:lvl>`` element. A child of abstract numbering definitions.

Numbering Defintion Instance
    A ``<w:num>`` element in the numbering part of the document. A numbering
    instance has it's own ID (``numId``), and references (and may override)
    an abstract numbering. Paragraphs in a document reference numbering instances, 
    NOT abstract numberings. 

Identifying a Numbering
-----------------------

An abstract numbering has a primary (unique) identifier of ``abstractNumId``.
These are unique within a document only. An abstract numbering can also have an
optional ``name`` field,

A numbering instance also has a unique identifer of ``numId``. Paragraphs 
reference a numbering instance using ``numId``.