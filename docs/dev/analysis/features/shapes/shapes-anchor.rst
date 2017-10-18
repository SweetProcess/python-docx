Anchor shape
============

Word allows a graphical object to be placed into a document as a floating
object. A floating shape appears as a ``<w:drawing>`` element as a child of
a ``<w:r>`` element and has a ``<w:anchor>`` child.


Candidate protocol -- anchor shape access
-----------------------------------------

The following interactive session illustrates the protocol for accessing an
anchor shape::

    >>> shapes = document.body.anchor_shapes
    >>> shape = shapes[0]
    >>> assert shape.type == MSO_SHAPE_TYPE.PICTURE


Resources
---------

* `Document Members (Word) on MSDN`_
* `Shape Members (Word) on MSDN`_

.. _Document Members (Word) on MSDN:
   http://msdn.microsoft.com/en-us/library/office/ff840898.aspx

.. _Shape Members (Word) on MSDN:
   http://msdn.microsoft.com/en-us/library/office/ff195191.aspx


MS API
------

The Shapes and InlineShapes properties on Document hold references to things
like pictures in the MS API.

* Height and Width
* Borders
* Shadow
* Hyperlink
* PictureFormat (providing brightness, color, crop, transparency, contrast)
* ScaleHeight and ScaleWidth
* HasChart
* HasSmartArt
* Type (Chart, LockedCanvas, Picture, SmartArt, etc.)


Spec references
---------------

* 17.3.3.9 drawing (DrawingML Object)
* 20.4.2.3 anchor (Anchor DrawingML Object)
* 20.4.2.7 extent (Drawing Object Size)


Minimal XML
-----------

.. highlight:: xml

This XML represents my best guess of the minimal inline shape container that
Word will load::

    <w:r>
      <w:drawing>
        <wp:anchor>
          <wp:simplePos x="0" y="0" />
          <wp:positionH relativeFrom="margin">
           <wp:align>right</wp:align>
          </wp:positionH>
          <wp:positionV relativeFrom="margin">
           <wp:align>center</wp:align>
          </wp:positionV>
          <wp:wrapSquare wrapText="bothSides"/>
          <wp:extent cx="914400" cy="914400"/>
          <wp:docPr id="1" name="Picture 1"/>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">

              <!-- might not have to put anything here for a start -->

            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>


Specimen XML
------------

.. highlight:: xml

A ``CT_Drawing`` (``<w:drawing>``) element can appear in a run, as a peer of,
for example, a ``<w:t>`` element. This element contains a DrawingML object.
WordprocessingML drawings are discussed in section 20.4 of the ISO/IEC spec.

This XML represents an inline shape inserted inline on a paragraph by itself.
The particulars of the graphical object itself are redacted::

    <w:p>
      <w:r>
        <w:rPr/>
          <w:noProof/>
        </w:rPr>
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="0" distR="0" behindDoc="0" locked="0" allowOverlap="1" simplePos="0" wp14:anchorId="1BDE1558" wp14:editId="31E593BB">
            <wp:extent cx="859536" cy="343814"/>
            <wp:effectExtent l="0" t="0" r="4445" b="12065"/>
            <wp:docPr id="1" name="Picture 1"/>
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
            </wp:cNvGraphicFramePr>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">

                <!-- graphical object, such as pic:pic, goes here -->

              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </w:r>
    </w:p>


Schema definitions
------------------

.. highlight:: xml

::

  <xsd:complexType name="CT_Drawing">
    <xsd:choice minOccurs="1" maxOccurs="unbounded">
      <xsd:element ref="wp:anchor" minOccurs="0"/>
      <xsd:element ref="wp:inline" minOccurs="0"/>
    </xsd:choice>
  </xsd:complexType>

  <xsd:complexType name="CT_Anchor">
    <xsd:sequence>
      <xsd:element name="simplePos" type="a:CT_Point2D"/>
      <xsd:element name="positionH" type="CT_PosH"/>
      <xsd:element name="positionV" type="CT_PosV"/>
      <xsd:element name="extent" type="a:CT_PositiveSize2D"/>
      <xsd:element name="effectExtent" type="CT_EffectExtent" minOccurs="0"/>
      <xsd:group ref="EG_WrapType"/>
      <xsd:element name="docPr" type="a:CT_NonVisualDrawingProps" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="cNvGraphicFramePr" type="a:CT_NonVisualGraphicFrameProperties"
        minOccurs="0" maxOccurs="1"/>
      <xsd:element ref="a:graphic" minOccurs="1" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="distT" type="ST_WrapDistance" use="optional"/>
    <xsd:attribute name="distB" type="ST_WrapDistance" use="optional"/>
    <xsd:attribute name="distL" type="ST_WrapDistance" use="optional"/>
    <xsd:attribute name="distR" type="ST_WrapDistance" use="optional"/>
    <xsd:attribute name="simplePos" type="xsd:boolean"/>
    <xsd:attribute name="relativeHeight" type="xsd:unsignedInt" use="required"/>
    <xsd:attribute name="behindDoc" type="xsd:boolean" use="required"/>
    <xsd:attribute name="locked" type="xsd:boolean" use="required"/>
    <xsd:attribute name="layoutInCell" type="xsd:boolean" use="required"/>
    <xsd:attribute name="hidden" type="xsd:boolean" use="optional"/>
    <xsd:attribute name="allowOverlap" type="xsd:boolean" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PosV">
    <xsd:sequence>
      <xsd:choice minOccurs="1" maxOccurs="1">
        <xsd:element name="align" type="ST_AlignV" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="posOffset" type="ST_PositionOffset" minOccurs="1" maxOccurs="1"/>
      </xsd:choice>
    </xsd:sequence>
    <xsd:attribute name="relativeFrom" type="ST_RelFromV" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PosH">
    <xsd:sequence>
      <xsd:choice minOccurs="1" maxOccurs="1">
        <xsd:element name="align" type="ST_AlignH" minOccurs="1" maxOccurs="1"/>
        <xsd:element name="posOffset" type="ST_PositionOffset" minOccurs="1" maxOccurs="1"/>
      </xsd:choice>
    </xsd:sequence>
    <xsd:attribute name="relativeFrom" type="ST_RelFromH" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_PositiveSize2D">
    <xsd:attribute name="cx" type="ST_PositiveCoordinate" use="required"/>
    <xsd:attribute name="cy" type="ST_PositiveCoordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_EffectExtent">
    <xsd:attribute name="l" type="a:ST_Coordinate" use="required"/>
    <xsd:attribute name="t" type="a:ST_Coordinate" use="required"/>
    <xsd:attribute name="r" type="a:ST_Coordinate" use="required"/>
    <xsd:attribute name="b" type="a:ST_Coordinate" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualDrawingProps">
    <xsd:sequence>
      <xsd:element name="hlinkClick" type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="hlinkHover" type="CT_Hyperlink"              minOccurs="0"/>
      <xsd:element name="extLst"     type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="id"     type="ST_DrawingElementId" use="required"/>
    <xsd:attribute name="name"   type="xsd:string"          use="required"/>
    <xsd:attribute name="descr"  type="xsd:string"          default=""/>
    <xsd:attribute name="hidden" type="xsd:boolean"         default="false"/>
    <xsd:attribute name="title"  type="xsd:string"          default=""/>
  </xsd:complexType>

  <xsd:complexType name="CT_NonVisualGraphicFrameProperties">
    <xsd:sequence>
      <xsd:element name="graphicFrameLocks" type="CT_GraphicalObjectFrameLocking" minOccurs="0"/>
      <xsd:element name="extLst"            type="CT_OfficeArtExtensionList"      minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GraphicalObject">
    <xsd:sequence>
      <xsd:element name="graphicData" type="CT_GraphicalObjectData"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_GraphicalObjectData">
    <xsd:sequence>
      <xsd:any minOccurs="0" maxOccurs="unbounded" processContents="strict"/>
    </xsd:sequence>
    <xsd:attribute name="uri" type="xsd:token" use="required"/>
  </xsd:complexType>
