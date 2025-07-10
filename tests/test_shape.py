# pyright: reportPrivateUsage=false

"""Test suite for the docx.shape module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.document import Document
from docx.enum.shape import WD_INLINE_SHAPE, WD_ANCHOR_SHAPE
from docx.oxml.document import CT_Body
from docx.oxml.ns import nsmap
from docx.oxml.shape import CT_Inline
from docx.shape import AnchorShape, InlineShape, InlineShapes
from docx.shared import Length, Emu


from .oxml.unitdata.dml import (
    a_blip, a_blipFill, a_graphic, a_graphicData, a_pic, an_inline,
    an_anchor,
)
from .unitutil.cxml import element, xml
from .unitutil.mock import FixtureRequest, Mock, instance_mock


class DescribeInlineShapes:
    """Unit-test suite for `docx.shape.InlineShapes` objects."""

    def it_knows_how_many_inline_shapes_it_contains(self, body: CT_Body, document_: Mock):
        inline_shapes = InlineShapes(body, document_)
        assert len(inline_shapes) == 2

    def it_can_iterate_over_its_InlineShape_instances(self, body: CT_Body, document_: Mock):
        inline_shapes = InlineShapes(body, document_)
        assert all(isinstance(s, InlineShape) for s in inline_shapes)
        assert len(list(inline_shapes)) == 2

    def it_provides_indexed_access_to_inline_shapes(self, body: CT_Body, document_: Mock):
        inline_shapes = InlineShapes(body, document_)
        for idx in range(-2, 2):
            assert isinstance(inline_shapes[idx], InlineShape)

    def it_raises_on_indexed_access_out_of_range(self, body: CT_Body, document_: Mock):
        inline_shapes = InlineShapes(body, document_)

        with pytest.raises(IndexError, match=r"inline shape index \[-3\] out of range"):
            inline_shapes[-3]
        with pytest.raises(IndexError, match=r"inline shape index \[2\] out of range"):
            inline_shapes[2]

    def it_knows_the_part_it_belongs_to(self, body: CT_Body, document_: Mock):
        inline_shapes = InlineShapes(body, document_)
        assert inline_shapes.part is document_.part

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def body(self) -> CT_Body:
        return cast(
            CT_Body, element("w:body/w:p/(w:r/w:drawing/wp:inline, w:r/w:drawing/wp:inline)")
        )

    @pytest.fixture
    def document_(self, request: FixtureRequest):
        return instance_mock(request, Document)


class DescribeAncorShape:

    def it_knows_what_type_of_shape_it_is(self, shape_type_fixture):
        shape, shape_type = shape_type_fixture
        assert shape.type == shape_type

    @pytest.fixture(params=[
        'embed pic', 'link pic', 'link+embed pic', 'chart', 'smart art',
        'not implemented'
    ])
    def shape_type_fixture(self, request):
        if request.param == 'embed pic':
            inline = self._with_picture(embed=True)
            shape_type = WD_ANCHOR_SHAPE.PICTURE

        elif request.param == 'link pic':
            inline = self._with_picture(link=True)
            shape_type = WD_ANCHOR_SHAPE.LINKED_PICTURE

        elif request.param == 'link+embed pic':
            inline = self._with_picture(embed=True, link=True)
            shape_type = WD_ANCHOR_SHAPE.LINKED_PICTURE

        elif request.param == 'chart':
            inline = self._with_uri(nsmap['c'])
            shape_type = WD_ANCHOR_SHAPE.CHART

        elif request.param == 'smart art':
            inline = self._with_uri(nsmap['dgm'])
            shape_type = WD_ANCHOR_SHAPE.SMART_ART

        elif request.param == 'not implemented':
            inline = self._with_uri('foobar')
            shape_type = WD_ANCHOR_SHAPE.NOT_IMPLEMENTED

        return AnchorShape(inline), shape_type

    def _with_picture(self, embed=False, link=False):
        picture_ns = nsmap['pic']

        blip_bldr = a_blip()
        if embed:
            blip_bldr.with_embed('rId1')
        if link:
            blip_bldr.with_link('rId2')

        image = (
            an_anchor().with_nsdecls('wp', 'r').with_child(
                a_graphic().with_nsdecls().with_child(
                    a_graphicData().with_uri(picture_ns).with_child(
                        a_pic().with_nsdecls().with_child(
                            a_blipFill().with_child(
                                blip_bldr)))))
        ).element
        return image

    def _with_uri(self, uri):
        inline = (
            an_anchor().with_nsdecls('wp').with_child(
                a_graphic().with_nsdecls().with_child(
                    a_graphicData().with_uri(uri)))
        ).element
        return inline


class DescribeInlineShape:
    """Unit-test suite for `docx.shape.InlineShape` objects."""

    @pytest.mark.parametrize(
        ("uri", "content_cxml", "expected_value"),
        [
            # -- embedded picture --
            (nsmap["pic"], "/pic:pic/pic:blipFill/a:blip{r:embed=rId1}", WD_INLINE_SHAPE.PICTURE),
            # -- linked picture --
            (
                nsmap["pic"],
                "/pic:pic/pic:blipFill/a:blip{r:link=rId2}",
                WD_INLINE_SHAPE.LINKED_PICTURE,
            ),
            # -- linked and embedded picture (not expected) --
            (
                nsmap["pic"],
                "/pic:pic/pic:blipFill/a:blip{r:embed=rId1,r:link=rId2}",
                WD_INLINE_SHAPE.LINKED_PICTURE,
            ),
            # -- chart --
            (nsmap["c"], "", WD_INLINE_SHAPE.CHART),
            # -- SmartArt --
            (nsmap["dgm"], "", WD_INLINE_SHAPE.SMART_ART),
            # -- something else we don't know about --
            ("foobar", "", WD_INLINE_SHAPE.NOT_IMPLEMENTED),
        ],
    )
    def it_knows_what_type_of_shape_it_is(
        self, uri: str, content_cxml: str, expected_value: WD_INLINE_SHAPE
    ):
        cxml = "wp:inline/a:graphic/a:graphicData{uri=%s}%s" % (uri, content_cxml)
        inline = cast(CT_Inline, element(cxml))
        inline_shape = InlineShape(inline)
        assert inline_shape.type == expected_value

    def it_knows_its_display_dimensions(self):
        inline = cast(CT_Inline, element("wp:inline/wp:extent{cx=333, cy=666}"))
        inline_shape = InlineShape(inline)

        width, height = inline_shape.width, inline_shape.height

        assert isinstance(width, Length)
        assert width == 333
        assert isinstance(height, Length)
        assert height == 666

    def it_can_change_its_display_dimensions(self):
        inline_shape = InlineShape(
            cast(
                CT_Inline,
                element(
                    "wp:inline/(wp:extent{cx=333,cy=666},a:graphic/a:graphicData/pic:pic/"
                    "pic:spPr/a:xfrm/a:ext{cx=333,cy=666})"
                ),
            )
        )

        inline_shape.width = Emu(444)
        inline_shape.height = Emu(888)

        assert inline_shape._inline.xml == xml(
            "wp:inline/(wp:extent{cx=444,cy=888},a:graphic/a:graphicData/pic:pic/pic:spPr/"
            "a:xfrm/a:ext{cx=444,cy=888})"
        )
