// Shared configuration for CUP Metadata Processor
const CONFIG = {
    // Valid binding types
    VALID_BINDINGS: new Set([
        'Cased',
        'Limp'
    ]),
    
    // Valid paper types
    VALID_PAPERS: new Set([
        'CUP MunkenPure 80 gsm',
        'Navigator 80 gsm',
        'Clairjet 90 gsm',
        'Magno Matt 90 gsm'
    ]),
    
    // Valid colour options
    VALID_COLOURS: new Set([
        'Mono',
        'Colour'
    ]),
    
    // Valid quality/route options
    VALID_ROUTES: new Set([
        'Standard',
        'Premium'
    ]),
    
    // Valid trim sizes
    VALID_TRIM_SIZES: new Set([
        '140x216',
        '152x229',
        '156x234',
        '170x244',
        '189x246',
        '178x254',
        '203x254',
        '216x280'
    ]),
    
    // Paper-specific compatibility rules
    PAPER_COMPATIBILITY: {
        'CUP MunkenPure 80 gsm': {
            allowedColours: new Set(['Mono']),
            allowedRoutes: new Set(['Standard'])
        },
        'Navigator 80 gsm': {
            allowedColours: new Set(['Mono']),
            allowedRoutes: new Set(['Standard'])
        },
        'Clairjet 90 gsm': {
            allowedColours: new Set(['Colour']),
            allowedRoutes: new Set(['Standard'])
        },
        'Magno Matt 90 gsm': {
            allowedColours: new Set(['Mono', 'Colour']),
            allowedRoutes: new Set(['Premium'])
        }
    },
    
    // Paper specifications for spine calculation
    PAPER_SPECS: {
        "CUP MunkenPure 80 gsm": { grammage: 80, volume: 13 },
        "Navigator 80 gsm": { grammage: 80, volume: 12.5 },
        "Clairjet 90 gsm": { grammage: 90, volume: 10 },
        "Magno Matt 90 gsm": { grammage: 90, volume: 10 }
    },
    
    // Column mapping from Excel headers to XML attribute names
    COLUMN_MAPPING: {
        "Master ISBN 13": "isbn",
        "full_title_for_export": "title",
        "pod_ExtentMod2": "extent",
        "pod_xml_TrimHeight": "trim_height",
        "pod_xml_TrimWidth": "trim_width",
        "pod_xml_TJ_PaperStock": "paper",
        "pod_Colourspace": "colour",
        "pod_xml_TJ_QualityOption": "quality",
        "pod_ColourBook_fullBleed": "bleed",
        "pod_TJ_BindType_c": "binding_style",
        "pod_CoverLaminate": "lamination",
        "pod_xml_Jacket": "jacket",
        "pod_xml_TJ_JacketLamination": "jacket_lamination",
        "pod_xml_TJ_FlapSize": "flap_size",
        "pod_xml_TJ_WibalinColour": "wibalin_colour",
        "pod_xml_SetISBN": "set_isbn"
    },
    
    // Desired column order for display and XML
    COLUMN_ORDER: [
        "isbn",
        "title",
        "trim_height",
        "trim_width",
        "extent",
        "spine",
        "paper",
        "gsm",
        "colour",
        "quality",
        "bleed",
        "binding_style",
        "lamination",
        "jacket",
        "jacket_lamination",
        "flap_size",
        "wibalin_colour",
        "set_isbn"
    ],
    
    // Value remapping for specific fields
    VALUE_MAPPING: {
        "paper": {
            "Cream 80gsm": "CUP MunkenPure 80 gsm",
            "White 80gsm": "Navigator 80 gsm"
        },
        "lamination": {
            "Matte": "Matt"
        }
    },
    
    // Constants for spine calculation
    SPINE_CALCULATION_FACTOR: 20000,
    HARDBACK_SPINE_ADDITION: 4,
    
    // Extent rounding thresholds
    EXTENT_NARROW_WIDTH_THRESHOLD: 156,
    EXTENT_NARROW_DIVISOR: 6,
    EXTENT_WIDE_DIVISOR: 4
};
