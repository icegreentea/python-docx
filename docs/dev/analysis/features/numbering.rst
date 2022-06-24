
Numbering
=========

In a WordprocessingML document, the numbering definitions is separated from the main
body of text, and is stored in a seperate part (``/word/numbering.xml``). 
Numbering is used to define both ordered and unordered (bullet) lists.

The numbering part contains a root ``<w:numbering>`` element (see ``CT_Numbering``).

Numbering definitions occur in two parts: the base *abstract* definition, and the 
actual (concrete) numbering instance. Numbering instances (not the abstract definition)
are what is actually refereneced by paragraphs in the main document. This is similar
to some programming languages use of abstract classes, where a concrete implementations
must be created to be actually used, even if the abstract itself defines all of the
functionality.

More pragmatically, use of new numbering instances are used to restart numberings in
ordered lists. 

Note that the default numbering schemes in a word processor are actually stored in
the editor, and saved into the document as needed. Unlike styles, there is not any
concept of latent numbering.

High Level Structure
--------------------

The high level structure of the numbering part is::

  numbering part
    contains: <w:numbering> (Numbering Definitions)
      contains: 0 or more <w:abstractNum> (Abstract Numbering Defintiions)
        identified by: <w:abstractNumId>
        contains: 0 or more <w:lvl> (Numbering Level Definition)
          identified by <w:ilvl> (sets the numbering level)
          contains: <w:pPr> (Numbering Level Associated Paragraph Properties)
          contains: <w:numFmt> (Numbering Format)
          contains: <w:start> (Starting Value)
          contains: <w:lvlRestart> (Restart Numbering Level Symbol)
          contains: <w:lvlText> (Numbering Level Text)
      contains: 0 or more <w:num> (Numbering Definition Instance)
        identified by: <w:numId>
        contains: 0 or 1 <w:abstractNumId> (reference to base Abstract Numbering)

The interconnections from numbering to application (at the paragraph) is::

  In the main document part:

    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val=0/> <!-- References the ilvl of the targeted numbering-->
          <w:numId w:val=5/> <!-- References a numId of the targeted <w:num> -->
        </w:numPr>
      </w:pPr>
    </w:p>

Numbering (``w:numbering``)
---------------------------

Protocol
~~~~~~~~

Get numbering from document::

  >>> numbering = doc.numbering

Use the ``numbering`` object as entry point for generating abstract and 
concrete numbering::

  >>> ab_num = numbering.create_abstract_numbering("new-numbering-name")
  >>> numbering = numbering.create_numbering_instance(ab_num)

Has helpers for creating bullet or decimal abstract numbering::

  >>> bullet_ab_num = numbering.create_bullet_abstract_numbering(
    "new_bullet_list"
  )
  >>> decimal_ab_num = numbering.create_decimal_abstract_numbering(
    "new_decimal_list"
  )

Abstract Numbering (``w:abstractNum``)
--------------------------------------

This is the heart of a numbering definition. 

Unimplemented Components
~~~~~~~~~~~~~~~~~~~~~~~~

These components have not be implemented into the top level protocol, but are mapped
in :class:`docx.oxml.numbering.CT_AbstractNum()`.

``nsid``
  Used to attempt to link abstract numbering across documents. 

``multiLevelType``
  Apparently this can be used by the editor to change GUI behavior. Should 
  be safe just to default to `multiLevel`

``tmpl``
  Numbering Template Code. Unique hex code. Used to define a location in the 
  GUI where the abstract numbering definition will be displayed.

``styleLink``
  Numbering Style Definition. Specifies that this abstract numbering definition
  is the base numbering style referenced. This basically allows the referenced
  ``styleLink`` to act as another name for the abstract number? To be honest, 
  I am not really sure what this does.

``numStyleLink``
  I am not sure what this does either.


Numbering Level Definition (``w:lvl``)
--------------------------------------

These define the appearance and behavior of a level within an abstract numbering.

XML Semantics
~~~~~~~~~~~~~

Unimplemented Components
~~~~~~~~~~~~~~~~~~~~~~~~

These components have not be implemented into the top level protocol, but are mapped
in :class:`docx.oxml.numbering.CT_Lvl()`.

Attribute ``tentative``

Attribute ``tplc``

``pStyle``
  This "reverse-binds" a paragraph style to a numbering style. The named paragraph 
  style will be forced to apply this level's numbering style/definition when applied.
  This could be a way to apply headering numbering for example.
  This causes ``w:ilvl`` references by the paragraph (in ``w:numPr``) to be ignored.
  Note that the paragraph still needs to define ``w:numId``.

``isLgl`` (Display All Levels Using Arabic Numerals)
  When present, forces all numbering 


``suff``


Numbering Definition Instance (``w:num``)
-----------------------------------------

Applying to Paragraph
---------------------


Schema excerpt
--------------

.. highlight:: xml

::

  <xsd:complexType name="CT_Numbering">
    <xsd:sequence>
      <xsd:element name="numPicBullet"      type="CT_NumPicBullet"  minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="abstractNum"       type="CT_AbstractNum"   minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="num"               type="CT_Num"           minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="numIdMacAtCleanup" type="CT_DecimalNumber" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_AbstractNum">
    <xsd:sequence>
      <xsd:element name="nsid"            type="CT_LongHexNumber"   minOccurs="0"/> 
      <xsd:element name="multiLevelType"  type="CT_MultiLevelType"  minOccurs="0"/>
      <xsd:element name="tmpl"            type="CT_LongHexNumber"   minOccurs="0"/>
      <xsd:element name="name"            type="CT_String"          minOccurs="0"/>
      <xsd:element name="styleLink"       type="CT_String"          minOccurs="0"/>
      <xsd:element name="numStyleLink"    type="CT_String"          minOccurs="0"/>
      <xsd:element name="lvl"             type="CT_Lvl"             minOccurs="0" maxOccurs="9"/> 
    </xsd:sequence>
    <xsd:attribute name="abstractNumId" type="ST_DecimalNumber" use="Required">
  </xsd:complexType>

  <xsd:complexType name="CT_Lvl">
    <xsd:sequence>
      <xsd:element name="start"           type="CT_DecimalNumber"   minOccurs="0"/>
      <xsd:element name="numFmt"          type="CT_NumFmt"          minOccurs="0"/>
      <xsd:element name="lvlRestart"      type="CT_DecimalNumber"   minOccurs="0"/>
      <xsd:element name="pStyle"          type="CT_String"          minOccurs="0"/>
      <xsd:element name="isLgl"           type="CT_OnOff"           minOccurs="0"/>
      <xsd:element name="suff"            type="CT_LevelSuffix"     minOccurs="0"/>
      <xsd:element name="lvlText"         type="CT_LevelText"       minOccurs="0"/>
      <xsd:element name="lvlPicBulletId"  type="CT_DecimalNumber"   minOccurs="0"/>
      <xsd:element name="lvlJc"           type="CT_Jc"              minOccurs="0"/>
      <xsd:element name="pPr"             type="CT_PPrGeneral"      minOccurs="0"/>
      <xsd:element name="rPr"             type="CT_RPr"             minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="ilvl"      type="ST_DecimalNumber" use="required"/>
    <xsd:attribute name="tplc"      type="ST_LongHexNumber" use="optional"/>
    <xsd:attribute name="tentative" type="s:ST_OnOff"       use="optional"/>
  </xsd:complexType> 

  <xsd:complexType name="CT_Num">
    <xsd:sequence>
      <xsd:element name="abstractNumId" type="CT_DecimalNumber"/>
      <xsd:element name="lvlOverride"   type="CT_NumLvl"        minOccurs="0" maxOccurs="9"/>
    </xsd:sequence>
    <xsd:attribute name="numId" type="ST_DecimalNumber" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_NumLvl">
    <xsd:sequence>
      <xsd:element name="startOverride" type="CT_DecimalNumber" minOccurs="0"/>
      <xsd:element name="lvl"           type="CT_Lvl"           minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="ilvl" type="ST_DecimalNumber" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_NumPr">
    <xsd:sequence>
      <xsd:element name="ilvl"            type="CT_DecimalNumber"        minOccurs="0"/>
      <xsd:element name="numId"           type="CT_DecimalNumber"        minOccurs="0"/>
      <xsd:element name="numberingChange" type="CT_TrackChangeNumbering" minOccurs="0"/>
      <xsd:element name="ins"             type="CT_TrackChange"          minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_DecimalNumber">
    <xsd:attribute name="val" type="ST_DecimalNumber" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_DecimalNumber">
    <xsd:restriction base="xsd:integer"/>
  </xsd:simpleType>
