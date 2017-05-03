"""Enumerations related to DrawingML shapes in WordprocessingML files."""

import enum


class WD_INLINE_SHAPE_TYPE(enum.Enum):
    """Corresponds to WdInlineShapeType enumeration.

    http://msdn.microsoft.com/en-us/library/office/ff192587.aspx.
    """

    CHART = 12
    LINKED_PICTURE = 4
    PICTURE = 3
    SMART_ART = 15
    NOT_IMPLEMENTED = -6


WD_INLINE_SHAPE = WD_INLINE_SHAPE_TYPE

class WD_ANCHOR_SHAPE_TYPE(enum.Enum):
    """
    Corresponds to WdInlineShapeType enumeration
    http://msdn.microsoft.com/en-us/library/office/ff192587.aspx
    """
    CHART = 12
    LINKED_PICTURE = 4
    PICTURE = 3
    SMART_ART = 15
    NOT_IMPLEMENTED = -6


WD_ANCHOR_SHAPE = WD_ANCHOR_SHAPE_TYPE


class WRAP_SHAPE_TYPE(enum.Enum):
    """Enumeration for shape wrapping types."""

    SQUARE_BOTH_SIDES = 'bothSides' # 'A square wrapped shape with text wrapping on both sides'
    TOP_AND_BOTTOM = '' # 'A square on its own line cleared on left and right'
