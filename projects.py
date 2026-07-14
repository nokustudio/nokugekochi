# Per-project configuration for the Good Earth Kochi walkthrough site.
# Shared assets (Materials, Moodboard, style.css, Noku mark) live at repo root and are
# identical for every project, so they are NOT configured here — they stay inline in
# update_site.py. Everything project-specific lives in PROJECTS below.

# Fallback quote if a room has no copy written yet.
_PLACEHOLDER_QUOTE = "Content for this space is being composed. Renders, products and dimensions to follow."


def _rooms(base, quotes):
    """Build a room config from `base` (list of id/folder/title/label tuples).
    `quotes` maps room id -> quote string; rooms without one get the placeholder."""
    return [
        {"id": rid, "folder": folder, "title": title, "label": label,
         "quote": quotes.get(rid, _PLACEHOLDER_QUOTE)}
        for rid, folder, title, label in base
    ]


# Umang: 11 rooms (room-materials is skipped in the walkthrough; it acts as the
# "01 = Plan Key" numbering offset so Foyer reads as 02).
_UMANG_BASE = [
    ("room-materials", "Materials", "Materials & Moodboard", "Palette"),
    ("room-01", "01. Foyer", "Foyer", "Entry"),
    ("room-02", "02. Living Room 01", "Living Room 01", "Reading"),
    ("room-03", "03. Living Room 02", "Living Room 02", "Seating"),
    ("room-04", "04. Living Room 03", "Living Room 03", "Work"),
    ("room-05", "05. Master Bedroom", "Master Bedroom", "Primary suite"),
    ("room-06", "06. Bedroom 01", "Bedroom 01", "Secondary"),
    ("room-07", "07.  Dining", "Dining", "Eat-in"),
    ("room-08", "08. Bedroom 02", "Bedroom 02", "Sleeping"),
    ("room-09", "09. Bedroom 03", "Bedroom 03", "Sleeping"),
    ("room-10", "10. Bar & Lounge", "Bar & Lounge", "Repose"),
    ("room-11", "11. Outdoor", "Outdoor", "Outdoor"),
]

# Modern Times: 9 rooms matching the folders on disk. The leading room-materials
# entry is a numbering sentinel only (skipped everywhere, like Umang) so the Plan Key
# reads as 01 and the rooms as 02–10, keeping section numbers and the index list in sync.
_MT_BASE = [
    ("room-materials", "Materials", "Materials & Moodboard", "Palette"),
    ("room-01", "01. Living Room", "Living Room", "Living"),
    ("room-02", "02. Dining", "Dining", "Dining"),
    ("room-03", "03. GF Balcony", "Ground Floor Balcony", "Outdoor"),
    ("room-04", "04. Breakfast Counter", "Breakfast Counter", "Counter"),
    ("room-05", "05. Bedroom 01", "Bedroom 01", "Sleeping"),
    ("room-06", "06. Bedroom 02", "Bedroom 02", "Sleeping"),
    ("room-07", "07. Bedroom 03", "Bedroom 03", "Sleeping"),
    ("room-08", "08. Master Bedroom", "Master Bedroom", "Primary suite"),
    ("room-09", "09. FF Balcony", "First Floor Balcony", "Outdoor"),
]

_MT_QUOTES = {
    "room-01": "Where the city softens. Brick, timber and greenery frame a living space that breathes with the building.",
    "room-02": "Gathered in warm light. A table set against the garden's edge, where meals slow to the pace of evening.",
    "room-03": "Ground-level green. The garden meets the threshold, blurring the line between home and landscape.",
    "room-04": "The quiet start. A compact counter for morning light, coffee, and the unhurried first hour.",
    "room-05": "A calm retreat. Grounded tones and soft texture compose a room made for rest.",
    "room-06": "Simple and serene. Clean lines and natural materials shape a private, restful corner.",
    "room-07": "Light and ease. A restful space where greenery and warm wood meet at the window.",
    "room-08": "A sanctuary above the city. Rich timber and layered light craft a suite for deep repose.",
    "room-09": "Between home and sky. The terrace garden brings the ecosystem to the doorstep — green, open, alive.",
}

# Modern Times is a two-floor home. Hotspot rects are in each plan's native
# 2400x1800 space (room id -> SVG rect attrs), split by the floor the room sits on.
_MT_GF_HOTSPOTS = {
    "room-01": 'x="1284" y="333" width="412" height="522"',    # Living Room
    "room-02": 'x="817" y="333" width="467" height="522"',     # Dining
    "room-03": 'x="817" y="856" width="1108" height="598"',   # Ground Floor Balcony
    "room-04": 'x="387" y="333" width="430" height="522"',     # Breakfast Counter
    "room-05": 'x="1925" y="333" width="417" height="915"',   # Bedroom 01
    "room-06": 'x="321" y="856" width="496" height="393"',    # Bedroom 02
}
_MT_FF_HOTSPOTS = {
    "room-07": 'x="1925" y="333" width="417" height="915"',   # Bedroom 03
    "room-08": 'x="321" y="333" width="496" height="915"',    # Master Bedroom
    "room-09": 'x="400" y="998" width="417" height="250"',    # First Floor Balcony
}

_MT_LAYOUT_IMAGES = {
    "room-01": {"file": "Room wise layout/Living Room.jpg", "title": "Living Room Layout"},
    "room-02": {"file": "Room wise layout/Dining.jpg", "title": "Dining Layout"},
    "room-03": {"file": "Room wise layout/Balcony GF.jpg", "title": "Ground Floor Balcony Layout"},
    "room-04": {"file": "Room wise layout/Breakfast Counter.jpg", "title": "Breakfast Counter Layout"},
    "room-05": {"file": "Room wise layout/Bedroom 01.jpg", "title": "Bedroom 01 Layout"},
    "room-06": {"file": "Room wise layout/Bedroom 02.jpg", "title": "Bedroom 02 Layout"},
    "room-07": {"file": "Room wise layout/Bedroom 03.jpg", "title": "Bedroom 03 Layout"},
    "room-08": {"file": "Room wise layout/Master Bedroom.jpg", "title": "Master Bedroom Layout"},
    # room-09 First Floor Balcony has no dedicated room-wise layout image yet.
}


_UMANG_QUOTES = {
    "room-materials": "The tactile foundation of Umang. A curation of warm wood, rich leathers, woven fabrics, and patinated brass.",
    "room-01": "A threshold of arrival. Setting a quiet tone with raw textures, warm wood, and gentle light.",
    "room-02": "A sanctuary of light and form. Where furniture becomes sculptural, crafting spaces of ease.",
    "room-03": "A space for quiet pause. Natural elements invite conversation and bring warmth to the heart of the home.",
    "room-04": "The art of focus and rest. A corner carved for deep thoughts, surrounded by tactile timber.",
    "room-05": "A retreat within a retreat. Soft light plays on rich wood, creating a sanctuary for quiet dreams.",
    "room-06": "Restful horizons. Clean lines and warm tones compose a space of absolute ease and privacy.",
    "room-07": "Gathered around solid timber. A celebration of craft, shared meals, and daily rituals.",
    "room-08": "Quiet clarity. The tactile warmth of natural wood framing moments of rest and reflection.",
    "room-09": "A peaceful haven. Simplicity in design offers a canvas for calm thoughts and restful sleep.",
    "room-10": "An intimate envelope. Rich tones, refined cane, and deep leather set a mood of relaxed sophistication.",
    "room-11": "Between earth and sky. Grounded timber frames the garden breeze, bridging the indoors with nature.",
}

# Umang keyplan hotspot rects on the 1980x1980 plan-key grid (room id -> SVG rect attrs).
_UMANG_HOTSPOTS = {
    "room-01": 'x="442" y="530" width="148" height="160"',
    "room-02": 'x="590" y="360" width="310" height="330"',
    "room-03": 'x="900" y="360" width="160" height="420"',
    "room-04": 'x="1060" y="360" width="120" height="100"',
    "room-05": 'x="1300" y="360" width="330" height="240"',
    "room-06": 'x="1060" y="460" width="210" height="260"',
    "room-07": 'x="590" y="690" width="270" height="220"',
    "room-08": 'x="442" y="910" width="218" height="300"',
    "room-09": 'x="660" y="1010" width="200" height="250"',
    "room-10": 'x="860" y="720" width="270" height="290"',
    "room-11": 'x="1130" y="720" width="300" height="290"',
}

_UMANG_LAYOUT_IMAGES = {
    "room-01": {"file": "Room wise layout/Foyer.png", "title": "Foyer Layout"},
    "room-02": {"file": "Room wise layout/Family.png", "title": "Living Room 01 Layout"},
    "room-03": {"file": "Room wise layout/Living.png", "title": "Living Room 02 Layout"},
    "room-04": {"file": "Room wise layout/Family.png", "title": "Living Room 03 Layout"},
    "room-05": {"file": "Room wise layout/Master Bedroom.png", "title": "Master Bedroom Layout"},
    "room-06": {"file": "Room wise layout/Bedroom 1.png", "title": "Bedroom 01 Layout"},
    "room-07": {"file": "Room wise layout/Dining.png", "title": "Dining Layout"},
    "room-08": {"file": "Room wise layout/Bedroom 2.png", "title": "Bedroom 02 Layout"},
    "room-09": {"file": "Room wise layout/Bedroom 3.png", "title": "Bedroom 03 Layout"},
    "room-10": {"file": "Room wise layout/Bar & Lounge.png", "title": "Bar & Lounge Layout"},
    "room-11": {"file": "Room wise layout/Balcony.png", "title": "Outdoor Layout"},
}

# Umang product-name -> rate-key overrides where simple normalization fails.
_UMANG_RATES_MAPPING = {
    "bed01": "Bed 01 A", "bed01a": "Bed 01 A", "bed47b": "Bed 47", "bed47": "Bed 47",
    "bench03": "Bench 03 A", "bench03a": "Bench 03 A", "bench14b": "Bench 14 A",
    "centertable13": "Centre Table 13 A", "centertable13a": "Centre Table 13 A",
    "studychair07": "Chair 07", "chair07": "Chair 07",
    "studychair12c": "Chair 12 C", "chair12c": "Chair 12 C",
    "chair27": "Chair 27 A", "chair27a": "Chair 27 A",
    "diningchair44c": "Chair 44 C", "chair44c": "Chair 44 C",
    "diningchair05": "Chair 05", "chair05": "Chair 05",
    "diningchair09a": "Chair 09 A", "chair09a": "Chair 09 A",
    "sidetable08": "Side Table 08 B", "sidetable08b": "Side Table 08 B",
    "sidetable31": "Side Table 31 C", "sidetable31b": "Side Table 31 C", "sidetable31c": "Side Table 31 C",
    "sidetable42": "Side Table 42 A", "sidetable42b": "Side Table 42 A", "sidetable42a": "Side Table 42 A",
    "storageunit29": "Storage 29 C", "storageunit29c": "Storage 29 C",
    "storageunit34": "Storage 34 A", "storageunit34a": "Storage 34 A",
    "studytable02": "Study Table 02 A", "studytable02a": "Study Table 02 A",
    "swing04": "Swing 04 A", "swing04a": "Swing 04 A",
    "tvunit09a": "TV Console 09 R2", "tvunit09r2": "TV Console 09 R2",
    "tvunit29": "TV Console 29 A R2", "tvunit29ar2": "TV Console 29 A R2",
}


UMANG = {
    "id": "umang",
    "name": "Umang Residence",
    "short_name": "Umang",
    "output": "umang.html",
    "base": "projects/umang",
    "rooms_config": _rooms(_UMANG_BASE, _UMANG_QUOTES),
    "custom_rates_mapping": _UMANG_RATES_MAPPING,
    # Plan key + furniture layout are per-floor lists (Umang is a single floor).
    "keyplan": {
        "floors": [
            {"label": None, "image": "Layout Simplified 3b with labels.png",
             "viewbox": "0 0 1980 1980", "hotspots": _UMANG_HOTSPOTS},
        ],
    },
    "layout_images": _UMANG_LAYOUT_IMAGES,
    "furniture_layout": {
        "floors": [
            {"label": None, "file": "Layout Original 6.png", "title": "Furniture Layout"},
        ],
    },
    "xlsx": {"database": "../../Product_Database.xlsx", "rates_ref": "../../GE Kochi Product Prices.xlsx"},
    "hero": {
        "article": "The", "name": "Umang", "subtitle": "Residence.",
        "eyebrow": "Good Earth &nbsp;·&nbsp; On the Periyar, Edappally",
        "theme_chip": "Theme One &nbsp;·&nbsp; Collection of Stillness",
        "room_count": "11 Rooms",
        # Researched facts (umang.goodearth.org.in) shown as the cover fact strip
        "facts": ["Edappally, Kochi", "68 Vertical Villas", "19 Floors", "Sky Gardens", "2028"],
        # Full-bleed render slideshow (official Good Earth renders, downloaded to web/)
        "cover_style": "kenburns",
        "cover_images": [
            "projects/umang/web/umang-front-elevation-view.webp",
            "projects/umang/web/umang-sky-deck-party-terrace.webp",
        ],
    },
    "copy": {
        "render_sub": "Umang Project",
        "customisation_intro": "Each piece in this proposal can be tuned for fabric, finish, wood, dimension and joinery to fit Umang specifically. Some examples:",
        "thankyou_descriptor": "We look forward to collaborating with Good Earth to bring <strong>Umang</strong> to life.",
        "thankyou_side_ticker": "Umang Residence &nbsp;·&nbsp; Good Earth Kochi",
        "thankyou_footer_client": "Umang Residence &nbsp;·&nbsp; Pitch",
        "meta_description": "A furnishing proposal for the Umang Residence. Composed by Noku Studio, Bengaluru, for the architects at Good Earth — a walkthrough of every room, and every piece in it.",
    },
}

MODERN_TIMES = {
    "id": "modern-times",
    "name": "Modern Times Residence",
    "short_name": "Modern Times",
    "output": "modern-times.html",
    "base": "projects/modern-times",
    "rooms_config": _rooms(_MT_BASE, _MT_QUOTES),
    # Most Modern Times pieces are the same catalog items as Umang, so reuse
    # Umang's rate-key overrides for the cases simple normalization can't match.
    "custom_rates_mapping": _UMANG_RATES_MAPPING,
    # Two-floor plan key: Ground Floor + First Floor, each with its own hotspots.
    "keyplan": {
        "floors": [
            {"label": "Ground Floor", "image": "Layout Simplified - GF.jpg",
             "viewbox": "0 0 2400 1800", "hotspots": _MT_GF_HOTSPOTS},
            {"label": "First Floor", "image": "Layout Simplified - FF.jpg",
             "viewbox": "0 0 2400 1800", "hotspots": _MT_FF_HOTSPOTS},
        ],
    },
    "layout_images": _MT_LAYOUT_IMAGES,
    "furniture_layout": {
        "floors": [
            {"label": "Ground Floor", "file": "Layout Original - GF 2.jpg", "title": "Ground Floor — Furniture Layout"},
            {"label": "First Floor", "file": "Layout Original - FF 2.jpg", "title": "First Floor — Furniture Layout"},
        ],
    },
    # Products are mostly the same catalog as Umang. Dimensions still point at
    # Umang's sheet (single source of truth); rates were copied into Modern
    # Times' own Furniture_Rates.xlsx so it's a self-contained, editable file.
    # Matches fall through the same fuzzy name/code matching already used
    # per-project; unmatched pieces stay blank either way.
    "xlsx": {"database": "../../Product_Database.xlsx", "rates_ref": "../../GE Kochi Product Prices.xlsx"},
    "hero": {
        "article": "The", "name": "Modern Times", "subtitle": "Residence.",
        "eyebrow": "Good Earth &nbsp;·&nbsp; Kalamassery, Kochi",
        "theme_chip": "The high-rise as an ecosystem",
        "room_count": "9 Rooms",
        # Researched facts (goodearthmoderntimes.com) shown as the cover fact strip
        "facts": ["Kalamassery, Kochi", "4 BHK Simplex & Duplex", "Ready to Move", "CREDAI 2025 &nbsp;·&nbsp; Most Sustainable"],
        # Built-photo triptych (official Good Earth photos, downloaded to web/).
        # No signage/text in any frame — the gate photo's project board clashed with the title.
        "cover_style": "triptych",
        "cover_images": [
            "projects/modern-times/web/mt-exterior-2.jpg",
            "projects/modern-times/web/mt-exterior-1.jpg",
            "projects/modern-times/web/mt-interior.jpg",
        ],
    },
    "copy": {
        "render_sub": "Modern Times Project",
        "customisation_intro": "Each piece can be tuned for fabric, finish, wood, dimension and joinery to fit Modern Times specifically. Some examples:",
        "thankyou_descriptor": "We look forward to collaborating with Good Earth to bring <strong>Modern Times</strong> to life.",
        "thankyou_side_ticker": "Modern Times Residence &nbsp;·&nbsp; Good Earth Kochi",
        "thankyou_footer_client": "Modern Times Residence &nbsp;·&nbsp; Pitch",
        "meta_description": "A furnishing proposal for the Modern Times Residence, composed by Noku Studio for Good Earth, Kochi.",
    },
}

PROJECTS = [UMANG, MODERN_TIMES]
