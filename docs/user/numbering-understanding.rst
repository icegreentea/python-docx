.. _numbering_styles:

Understanding Numbering
=======================

How does Numbering Work?
------------------------

Word documents can define any number of numbering (list) formats. These numbering
formats can then by applied to different paragraphs to get different numbering.
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


Glossary
--------

abstract numbering
    A ``<w:abstractNum>`` element in the numbering part of a document that 
    defines the attributes of one numbering scheme. Paragraphs in a document 
    reference numbering instances, NOT abstract numberings.

numbering instance
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