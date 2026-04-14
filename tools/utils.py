#!/usr/bin/env python3

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.xmlchemy import BaseOxmlElement
from docx.text.paragraph import Paragraph
from docx.text.run import Run


# Path to the template document, which defines page style and numbering and contains the EPSYL logo
TEMPLATE_DC_PATH = "assets/template.docx"


# Default font property values used when no local values can be found
DEFAULT_FONT = {"name": "Arial", "size": None, "color": None,
                "bold": False, "italic": False, "underline": False}


# Default paragraph format property values used when no local values can be found
DEFAULT_FMT = {"alignment": None}


def get_public_attrs(o) -> list[str]:
    """
    Return the list of `o`'s attributes whose names do not start with a "_".
    """
    return [e for e in dir(o) if not e.startswith("_")]


def get_ilvl(p: Paragraph) -> int | None:
    """
    Return `p`'s list indentation level, or `None` if `p` is not part of a list.
    """
    try:
        ilvl = p._p.pPr.numPr.ilvl.val
    except AttributeError:
        ilvl = None
    return ilvl


def get_font_props(run: Run, parent: Paragraph) -> dict:
    """
    Return a dictionary containing `run`'s basic font properties: size, color, bold, italic and
    underline. Each property is looked for individually and in multiple font objects following style
    hierarchy (the next object is looked up if a property equals `None` in the previous one).
    """
    font = DEFAULT_FONT.copy()
    font["size"] = run.font.size or run.style.font.size or parent.style.font.size or font["size"]
    color = run.font.color.rgb or run.style.font.color.rgb or parent.style.font.color.rgb
    font["color"] = str(color) if color is not None else font["color"]
    for prop in ("bold", "italic", "underline"):
        for src in (run, run.font, run.style.font, parent.style.font):
            if (val := getattr(src, prop)) is not None:
                font[prop] = val
                break
    return font


def get_format_props(p: Paragraph) -> dict:
    """
    Return a dictionary containing paragraph `p`'s format properties like alignment.
    """
    fmt = DEFAULT_FMT.copy()
    fmt["alignment"] = p.alignment or p.paragraph_format.alignment or fmt["alignment"]
    return fmt


def rec_add_xml_children(elm: BaseOxmlElement, children: dict[str, dict | str]) -> None:
    """
    Fill an XML element `elm` with the children tree described by `children`.
    The most nested key-value pairs represent the attributes of their respective parent elements.
    Each element/attribute name is automatically prefixed with "w:".
    """
    for child_name, child in children.items():
        if isinstance(child, str):
            # the dict item represents an attribute
            elm.set(qn(f"w:{child_name}"), child)
        else:  # dict
            # the dict item represents a child element
            child_elm = OxmlElement(f"w:{child_name}")
            rec_add_xml_children(child_elm, child)
            elm.append(child_elm)
