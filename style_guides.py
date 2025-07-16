from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_UNDERLINE
from docx.enum.section import WD_SECTION_START, WD_ORIENTATION
from docx.enum.style import WD_STYLE_TYPE

STYLE_GUIDES = {
    "apa": {
        "meta": {
            "title_page": True,
            "page_numbers": "header-right",
            "requires_abstract": True,
            "default_font": "Times New Roman",
            "citation_format": "({Author}, {Year})",
            "page_size": (Inches(8.5), Inches(11)),  # Letter size
            "orientation": WD_ORIENTATION.PORTRAIT,
            "default_tab_stops": Inches(0.5)
        },
        "margins": {
            "left": Inches(1),
            "right": Inches(1),
            "top": Inches(1),
            "bottom": Inches(1),
            "header": Inches(0.5),
            "footer": Inches(0.5),
            "gutter": Inches(0)
        },
        "styles": {
            "Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(18),
                    "bold": True,
                    "all_caps": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0),
                    "space_after": Pt(24),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "widow_control": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 1": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(24),
                    "space_after": Pt(6),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 1
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 2": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(12),
                    "space_after": Pt(6),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 2
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 3": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(12),
                    "space_after": Pt(6),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 3
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Normal": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0.5),
                    "keep_together": True,
                    "widow_control": True,
                    "orphan_control": True
                },
                "base_style": None,
                "next_style": "Normal"
            },
            "Abstract": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(12),
                    "space_after": Pt(12),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Block Quote": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.5),
                    "right_indent": Inches(0.5),
                    "space_before": Pt(12),
                    "space_after": Pt(12),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "References": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "left_indent": Inches(0.5),
                    "first_line_indent": Inches(-0.5),
                    "keep_together": True,
                    "widow_control": True,
                    "orphan_control": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Figure": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(12),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE
                },
                "based_on": "Normal",
                "next_style": "Figure Caption"
            },
            "Figure Caption": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "italic": True
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0),
                    "space_after": Pt(12),
                    "line_spacing": WD_LINE_SPACING.DOUBLE
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Table": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(12),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE
                },
                "based_on": "Normal",
                "next_style": "Table Title"
            },
            "Table Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "italic": True
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0),
                    "space_after": Pt(6),
                    "line_spacing": WD_LINE_SPACING.DOUBLE
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Table Note": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(10)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(6),
                    "space_after": Pt(12),
                    "line_spacing": WD_LINE_SPACING.SINGLE
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Footnote Text": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(10),
                    "bold": False,
                    "italic": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "space_before": Pt(0),
                    "space_after": Pt(0)
                }
            },
            "Appendix Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "color": RGBColor(0, 0, 0)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(24),
                    "space_after": Pt(6),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Appendices Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(24),
                    "space_after": Pt(6),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 1
                },
                "based_on": "Normal",
                "next_style": "Normal"
            }
        }
    },
    
    "mla": {
        "meta": {
            "title_page": False,
            "page_numbers": "header-right",
            "requires_abstract": False,
            "default_font": "Times New Roman",
            "citation_format": "({Author} {Page})",
            "page_size": (Inches(8.5), Inches(11)),
            "orientation": WD_ORIENTATION.PORTRAIT,
            "default_tab_stops": Inches(0.5)
        },
        "margins": {
            "left": Inches(1),
            "right": Inches(1),
            "top": Inches(1),
            "bottom": Inches(1),
            "header": Inches(0.5),
            "footer": Inches(0.5),
            "gutter": Inches(0)
        },
        "styles": {
            "Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "all_caps": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0)
                }
            },
            "Heading 1": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "outline_level": 1
                }
            },
            "Normal": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0.5)
                }
            },
            "Block Quote": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(1),
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0.5)
                }
            },
            "References": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "left_indent": Inches(0.5),
                    "first_line_indent": Inches(-0.5)
                }
            }
        }
    },
    
    "chicago": {
        "meta": {
            "title_page": True,
            "page_numbers": "header-right",
            "requires_abstract": False,
            "default_font": "Times New Roman",
            "citation_format": "({Author}, {Year}, {Page})",
            "page_size": (Inches(8.5), Inches(11)),
            "orientation": WD_ORIENTATION.PORTRAIT,
            "default_tab_stops": Inches(0.5)
        },
        "margins": {
            "left": Inches(1),
            "right": Inches(1),
            "top": Inches(1),
            "bottom": Inches(1),
            "header": Inches(0.5),
            "footer": Inches(0.5),
            "gutter": Inches(0)
        },
        "styles": {
            "Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(16),
                    "bold": True,
                    "all_caps": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0),
                    "space_after": Pt(12),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0)
                }
            },
            "Heading 1": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(12),
                    "space_after": Pt(6),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "outline_level": 1
                }
            },
            "Normal": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0.5)
                }
            },
            "Block Quote": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(11),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.5),
                    "right_indent": Inches(0.5),
                    "space_before": Pt(6),
                    "space_after": Pt(6),
                    "line_spacing": WD_LINE_SPACING.SINGLE,
                    "first_line_indent": Inches(0)
                }
            },
            "References": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.SINGLE,
                    "left_indent": Inches(0.5),
                    "first_line_indent": Inches(-0.5)
                }
            }
        }
    }
}
