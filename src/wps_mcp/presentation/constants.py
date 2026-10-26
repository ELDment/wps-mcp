"""WPS Presentation and MSO COM Automation Constants.

Central repository of all enumeration constants used across the MCP
server. All constants are defined as module-level integers grouped by
enumeration type.
"""

MSO_TRUE = -1
MSO_FALSE = 0
MSO_TRI_STATE_MIXED = -2
PP_LAYOUT_BLANK = 12
MSO_FREEFORM = 5
MSO_GROUP = 6
MSO_LINKED_PICTURE = 11
MSO_PICTURE = 13
MSO_PLACEHOLDER = 14
MSO_SMART_ART = 24
MSO_SHAPE_RECTANGLE = 1
PP_PLACEHOLDER_TITLE = 1
PP_PLACEHOLDER_BODY = 2
PP_PLACEHOLDER_CENTER_TITLE = 3
PP_PLACEHOLDER_SUBTITLE = 4
PP_ALIGN_LEFT = 1
PP_ALIGN_CENTER = 2
PP_ALIGN_RIGHT = 3
PP_ALIGN_JUSTIFY = 4
PP_ALIGN_DISTRIBUTE = 5
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
MSO_TEXT_ORIENTATION_UPWARD = 2
MSO_TEXT_ORIENTATION_DOWNWARD = 3
MSO_TEXT_ORIENTATION_VERTICAL = 5
PP_AUTO_SIZE_NONE = 0
PP_AUTO_SIZE_SHAPE_TO_FIT_TEXT = 1
PP_AUTO_SIZE_TEXT_TO_FIT_SHAPE = 2
MSO_BRING_TO_FRONT = 0
MSO_SEND_TO_BACK = 1
MSO_BRING_FORWARD = 2
MSO_SEND_BACKWARD = 3
MSO_LINE_SOLID = 1
MSO_LINE_SQUARE_DOT = 2
MSO_LINE_ROUND_DOT = 3
MSO_LINE_DASH = 4
MSO_LINE_DASH_DOT = 5
MSO_LINE_DASH_DOT_DOT = 6
MSO_LINE_LONG_DASH = 7
MSO_LINE_LONG_DASH_DOT = 8
MSO_LINE_LONG_DASH_DOT_DOT = 9
MSO_LINE_SYS_DASH = 10
MSO_LINE_SYS_DOT = 11
MSO_LINE_SYS_DASH_DOT = 12
DASH_STYLE_MAP = {
    "solid": MSO_LINE_SOLID,
    "square_dot": MSO_LINE_SQUARE_DOT,
    "round_dot": MSO_LINE_ROUND_DOT,
    "dash": MSO_LINE_DASH,
    "dash_dot": MSO_LINE_DASH_DOT,
    "dash_dot_dot": MSO_LINE_DASH_DOT_DOT,
    "long_dash": MSO_LINE_LONG_DASH,
    "long_dash_dot": MSO_LINE_LONG_DASH_DOT,
    "long_dash_dot_dot": MSO_LINE_LONG_DASH_DOT_DOT,
    "sys_dash": MSO_LINE_SYS_DASH,
    "sys_dot": MSO_LINE_SYS_DOT,
    "sys_dash_dot": MSO_LINE_SYS_DASH_DOT,
}
DASH_STYLE_NAMES = {v: k for k, v in DASH_STYLE_MAP.items()}
DASH_STYLE_DESCRIPTION = "Dash style: " + ", ".join(DASH_STYLE_MAP)


def dash_style_value(name: str) -> int:
    """The MsoLineDashStyle number for a dash style name.

    Case and surrounding space are ignored. Unknown names raise ValueError.
    """
    key = str(name).strip().lower()
    value = DASH_STYLE_MAP.get(key)
    if value is None:
        raise ValueError(
            f"Unknown dash_style '{name}'. Valid "
            f"values: {list(DASH_STYLE_MAP.keys())}"
        )
    return value


def check_dash_style(name: str | None) -> str | None:
    """Normalize a dash style name or reject an unknown style."""
    if name is None:
        return None
    dash_style_value(name)
    return str(name).strip().lower()


MSO_ARROWHEAD_NONE = 1
MSO_ARROWHEAD_TRIANGLE = 2
MSO_ARROWHEAD_OPEN = 3
MSO_ARROWHEAD_STEALTH = 4
MSO_ARROWHEAD_DIAMOND = 5
MSO_ARROWHEAD_OVAL = 6
PP_BULLET_NONE = 0
PP_BULLET_UNNUMBERED = 1
PP_BULLET_NUMBERED = 2
PP_BULLET_ALPHA_LC_PERIOD = 0
PP_BULLET_ALPHA_UC_PERIOD = 1
PP_BULLET_ARABIC_PAREN_RIGHT = 2
PP_BULLET_ARABIC_PERIOD = 3
PP_BULLET_ROMAN_LC_PAREN_BOTH = 4
PP_BULLET_ROMAN_LC_PAREN_RIGHT = 5
PP_BULLET_ROMAN_LC_PERIOD = 6
PP_BULLET_ROMAN_UC_PERIOD = 7
PP_BULLET_ALPHA_LC_PAREN_BOTH = 8
PP_BULLET_ALPHA_LC_PAREN_RIGHT = 9
PP_BULLET_ALPHA_UC_PAREN_BOTH = 10
PP_BULLET_ALPHA_UC_PAREN_RIGHT = 11
PP_BULLET_ARABIC_PAREN_BOTH = 12
PP_BULLET_ARABIC_PLAIN = 13
PP_BULLET_ROMAN_UC_PAREN_BOTH = 14
PP_BULLET_ROMAN_UC_PAREN_RIGHT = 15
PP_BULLET_KANJI_KOREAN_PLAIN = 26
PP_BULLET_KANJI_KOREAN_PERIOD = 27
PP_BULLET_ARABIC_DB_PLAIN = 28
PP_BULLET_ARABIC_DB_PERIOD = 29
PP_SELECTION_NONE = 0
PP_SELECTION_SLIDES = 1
PP_SELECTION_SHAPES = 2
PP_SELECTION_TEXT = 3
PP_SHOW_TYPE_SPEAKER = 1
PP_SHOW_TYPE_WINDOW = 2
PP_SHOW_TYPE_KIOSK = 3
PP_SHOW_ALL = 1
PP_SHOW_SLIDE_RANGE = 2
PP_FIXED_FORMAT_TYPE_PDF = 2
SHAPE_TYPE_NAMES = {
    1: "AutoShape",
    2: "Callout",
    3: "Chart",
    4: "Comment",
    5: "Freeform",
    6: "Group",
    7: "EmbeddedOLEObject",
    8: "FormControl",
    9: "Line",
    10: "LinkedOLEObject",
    11: "LinkedPicture",
    12: "OLEControlObject",
    13: "Picture",
    14: "Placeholder",
    15: "TextEffect",
    16: "Media",
    17: "TextBox",
    19: "Table",
    24: "SmartArt",
}
PLACEHOLDER_TYPE_NAMES = {
    1: "Title",
    2: "Body",
    3: "CenterTitle",
    4: "Subtitle",
    5: "VerticalTitle",
    6: "VerticalBody",
    7: "Object",
    8: "Chart",
    9: "Bitmap",
    10: "MediaClip",
    11: "OrgChart",
    12: "Table",
    13: "SlideNumber",
    14: "Header",
    15: "Footer",
    16: "Date",
    17: "VerticalObject",
    18: "Picture",
}
WINDOW_STATE_NAMES = {1: "normal", 2: "minimized", 3: "maximized"}
SLIDESHOW_STATE_NAMES = {
    1: "running",
    2: "paused",
    3: "black_screen",
    4: "white_screen",
    5: "done",
}
PP_ACTION_NONE = 0
PP_ACTION_HYPERLINK = 7
PP_MOUSE_CLICK = 1
PP_MOUSE_OVER = 2
ANIMATION_EFFECT_NAMES = {
    1: "appear",
    2: "fly",
    3: "blinds",
    4: "box",
    5: "checkerboard",
    6: "circle",
    8: "diamond",
    9: "dissolve",
    10: "fade",
    16: "split",
    22: "wipe",
    23: "zoom",
    26: "bounce",
    30: "float",
    31: "grow_and_turn",
    54: "change_fill_color",
    55: "change_font",
    56: "change_font_color",
    57: "change_font_size",
    59: "grow_shrink",
    61: "spin",
    62: "transparency",
    63: "bold_flash",
    69: "color_wave",
    73: "darken",
    74: "desaturate",
    75: "flash_bulb",
    78: "lighten",
    80: "teeter",
    82: "wave",
    86: "path_circle",
    88: "path_diamond",
    90: "path_star",
    92: "path_square",
    94: "path_heart",
    109: "path_loop",
    120: "path_left",
    122: "path_arc_down",
    123: "path_zigzag",
    125: "path_sine_wave",
    126: "path_bounce_left",
    127: "path_down",
    129: "path_arc_up",
    131: "path_spiral_right",
    132: "path_wave",
    134: "path_diagonal_down_right",
    136: "path_arc_left",
    137: "path_funnel",
    138: "path_spring",
    139: "path_bounce_right",
    141: "path_diagonal_up_right",
    143: "path_arc_right",
    148: "path_up",
    149: "path_right",
}
ANIMATION_TRIGGER_NAMES = {
    0: "none",
    1: "on_click",
    2: "with_previous",
    3: "after_previous",
    4: "on_shape_click",
}
ANIM_DIRECTION_MAP = {
    "none": 0,
    "up": 1,
    "right": 2,
    "down": 3,
    "left": 4,
    "up_left": 6,
    "up_right": 7,
    "down_right": 8,
    "down_left": 9,
    "top": 10,
    "bottom": 11,
    "top_left": 12,
    "top_right": 13,
    "bottom_right": 14,
    "bottom_left": 15,
    "horizontal": 16,
    "vertical": 17,
    "across": 18,
    "in": 19,
    "out": 20,
    "clockwise": 21,
    "counterclockwise": 22,
    "horizontal_in": 23,
    "horizontal_out": 24,
    "vertical_in": 25,
    "vertical_out": 26,
}
ANIM_DIRECTION_NAMES = {v: k for k, v in ANIM_DIRECTION_MAP.items()}
ALIGN_CMD_MAP = {
    "left": 0,
    "center": 1,
    "right": 2,
    "top": 3,
    "middle": 4,
    "bottom": 5,
}
DISTRIBUTE_CMD_MAP = {"horizontal": 0, "vertical": 1}
FLIP_CMD_MAP = {"horizontal": 0, "vertical": 1}
MERGE_CMD_MAP = {
    "union": 1,
    "combine": 2,
    "intersect": 3,
    "subtract": 4,
    "fragment": 5,
}
SLIDE_SIZE_MAP = {
    "4:3": 1,
    "letter": 2,
    "a4": 3,
    "35mm": 4,
    "overhead": 5,
    "banner": 6,
    "custom": 7,
    "a3": 8,
    "16:9": 9,
    "16:10": 10,
    "widescreen": 9,
}
GRADIENT_STYLE_MAP = {
    "horizontal": 1,
    "vertical": 2,
    "diagonal_up": 3,
    "diagonal_down": 4,
    "from_corner": 5,
    "from_title": 6,
    "from_center": 7,
}
SHAPE_FORMAT_MAP = {"gif": 0, "jpg": 1, "png": 2, "bmp": 3, "wmf": 4, "emf": 5}
VIEW_TYPE_MAP = {
    "normal": 1,
    "slide_master": 2,
    "notes_page": 3,
    "handout_master": 4,
    "notes_master": 5,
    "outline": 6,
    "slide_sorter": 7,
    "title_master": 8,
    "reading": 10,
}
VIEW_TYPE_NAMES = {v: k for k, v in VIEW_TYPE_MAP.items()}
PP_BORDER_TOP = 1
PP_BORDER_LEFT = 2
PP_BORDER_BOTTOM = 3
PP_BORDER_RIGHT = 4
PP_BORDER_DIAGONAL_DOWN = 5
PP_BORDER_DIAGONAL_UP = 6
MSO_ANCHOR_TOP = 1
MSO_ANCHOR_TOP_BASELINE = 2
MSO_ANCHOR_MIDDLE = 3
MSO_ANCHOR_BOTTOM = 4
MSO_ANCHOR_BOTTOM_BASE_LINE = 5
MSO_EDITING_AUTO = 0
MSO_SEGMENT_LINE = 0
EDITING_TYPE_MAP = {"auto": 0, "corner": 1, "smooth": 2, "symmetric": 3}
EDITING_TYPE_NAMES = {0: "auto", 1: "corner", 2: "smooth", 3: "symmetric"}
SEGMENT_TYPE_MAP = {"line": 0, "curve": 1}
SEGMENT_TYPE_NAMES = {0: "line", 1: "curve"}
PICTURE_COLOR_TYPE_MAP = {
    "automatic": 1,
    "grayscale": 2,
    "black_and_white": 3,
    "watermark": 4,
}
PICTURE_COLOR_TYPE_NAMES = {v: k for k, v in PICTURE_COLOR_TYPE_MAP.items()}
AFTER_EFFECT_MAP = {"none": 0, "dim": 1, "hide": 2, "hide_on_next_click": 3}
AFTER_EFFECT_NAMES = {v: k for k, v in AFTER_EFFECT_MAP.items()}
BUILD_LEVEL_MAP = {
    "none": 0,
    "all_levels": 1,
    "first_level": 2,
    "second_level": 3,
    "third_level": 4,
    "fourth_level": 5,
    "fifth_level": 6,
}
BUILD_LEVEL_NAMES = {v: k for k, v in BUILD_LEVEL_MAP.items()}
TEXT_UNIT_EFFECT_MAP = {"by_paragraph": 0, "by_character": 1, "by_word": 2}
TEXT_UNIT_EFFECT_NAMES = {v: k for k, v in TEXT_UNIT_EFFECT_MAP.items()}
ICON_PACKAGE_BASE = (
    "https://cdn.jsdelivr.net/npm/@material-symbols/svg-400@0.31.3"
)
