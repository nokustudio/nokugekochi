import os
import re
import json
import html
import openpyxl

def main():
    print("Starting Noku Pitch Webpage Auto-Updater...")
    
    # 1. Paths relative to this script
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base_dir, "Product Dimensions.xlsx")
    html_path = os.path.join(base_dir, "index.html")
    
    if not os.path.exists(excel_path):
        print(f"Error: Excel file not found at: {excel_path}")
        return
        
    if not os.path.exists(html_path):
        print(f"Error: index.html not found at: {html_path}")
        return

    # 2. Get list of room directories (looking for folders starting with digits)
    room_dirs = []
    for item in sorted(os.listdir(base_dir)):
        item_path = os.path.join(base_dir, item)
        if os.path.isdir(item_path) and re.match(r'^\d+\.', item):
            room_dirs.append(item)
            
    print(f"Found {len(room_dirs)} room directories on disk.")

    # 3. Load excel database
    print("\nReading Excel database...")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    db_products = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        
        # Headers: Space | Code | Product | W | D | H | SH | MH | Dimension ...
        for row in rows[1:]:
            if len(row) < 3 or row[2] is None:
                continue
            db_products.append({
                "sheet": sheet_name,
                "space": str(row[0]).strip() if row[0] is not None else "",
                "code": str(row[1]).strip() if row[1] is not None else "",
                "name": str(row[2]).strip(),
                "w": row[3],
                "d": row[4],
                "h": row[5],
                "sh": row[6],
                "mh": row[7],
                "dim_str": str(row[8]).strip() if len(row) > 8 and row[8] is not None else ""
            })
    print(f"Loaded {len(db_products)} product rows from Excel.")

    # 3b. Load rates database
    rates_path = os.path.join(base_dir, "Furniture_Rates.xlsx")
    rates_db = {}
    if os.path.exists(rates_path):
        print("Reading Furniture Rates database...")
        wb_rates = openpyxl.load_workbook(rates_path, data_only=True)
        ws_rates = wb_rates.active
        for row in ws_rates.iter_rows(min_row=2, values_only=True):
            if row[0] and row[1] is not None:
                rates_db[str(row[0]).strip()] = row[1]
        print(f"Loaded {len(rates_db)} rate entries from Excel.")

    # 5. 11-room sequence config definition matching folders exactly
    rooms_config = [
        {"id": "room-materials", "folder": "Materials", "title": "Materials & Moodboard", "label": "Palette", "quote": "The tactile foundation of Umang. A curation of warm wood, rich leathers, woven fabrics, and patinated brass."},
        {"id": "room-01", "folder": "01. Foyer", "title": "Foyer", "label": "Entry", "quote": "A threshold of arrival. Setting a quiet tone with raw textures, warm wood, and gentle light."},
        {"id": "room-02", "folder": "02. Living Room 01", "title": "Living Room 01", "label": "Seating", "quote": "A space for quiet pause. Natural elements invite conversation and bring warmth to the heart of the home."},
        {"id": "room-03", "folder": "03. Living Room 02", "title": "Living Room 02", "label": "Reading", "quote": "A sanctuary of light and form. Where furniture becomes sculptural, crafting spaces of ease."},
        {"id": "room-04", "folder": "04. Living Room 03", "title": "Living Room 03", "label": "Work", "quote": "The art of focus and rest. A corner carved for deep thoughts, surrounded by tactile timber."},
        {"id": "room-05", "folder": "05. Master Bedroom", "title": "Master Bedroom", "label": "Primary suite", "quote": "A retreat within a retreat. Soft light plays on rich wood, creating a sanctuary for quiet dreams."},
        {"id": "room-06", "folder": "06. Bedroom 01", "title": "Bedroom 01", "label": "Secondary", "quote": "Restful horizons. Clean lines and warm tones compose a space of absolute ease and privacy."},
        {"id": "room-07", "folder": "07.  Dining", "title": "Dining", "label": "Eat-in", "quote": "Gathered around solid timber. A celebration of craft, shared meals, and daily rituals."},
        {"id": "room-08", "folder": "08. Bedroom 02", "title": "Bedroom 02", "label": "Sleeping", "quote": "Quiet clarity. The tactile warmth of natural wood framing moments of rest and reflection."},
        {"id": "room-09", "folder": "09. Bedroom 03", "title": "Bedroom 03", "label": "Sleeping", "quote": "A peaceful haven. Simplicity in design offers a canvas for calm thoughts and restful sleep."},
        {"id": "room-10", "folder": "10. Bar & Lounge", "title": "Bar & Lounge", "label": "Repose", "quote": "An intimate envelope. Rich tones, refined cane, and deep leather set a mood of relaxed sophistication."},
        {"id": "room-11", "folder": "11. Outdoor", "title": "Outdoor", "label": "Outdoor", "quote": "Between earth and sky. Grounded timber frames the garden breeze, bridging the indoors with nature."}
    ]

    # 4. Helper to normalize names for mapping
    def normalize_name(name):
        n = name.lower()
        n = re.sub(r'\.png$|\.jpg$|\.jpeg$', '', n)
        n = re.sub(r'\(.*?\)', '', n) # Strip parentheses and contents
        n = re.sub(r'\bv2\b|\bv1\b|\bfull image\b|\bcopy\b', '', n)
        n = re.sub(r'[^a-z0-9]', '', n)
        return n

    # Custom mapping for rate items where simple normalization might fail
    custom_rates_mapping = {
        "bed01": "Bed 01 A",
        "bed01a": "Bed 01 A",
        "bed47b": "Bed 47",
        "bed47": "Bed 47",
        "bench03": "Bench 03 A",
        "bench03a": "Bench 03 A",
        "bench14b": "Bench 14 A",
        "centertable13": "Centre Table 13 A",
        "centertable13a": "Centre Table 13 A",
        "studychair07": "Chair 07",
        "chair07": "Chair 07",
        "studychair12c": "Chair 12 C",
        "chair12c": "Chair 12 C",
        "chair27": "Chair 27 A",
        "chair27a": "Chair 27 A",
        "diningchair44c": "Chair 44 C",
        "chair44c": "Chair 44 C",
        "diningchair05": "Chair 05",
        "chair05": "Chair 05",
        "diningchair09a": "Chair 09 A",
        "chair09a": "Chair 09 A",
        "sidetable08": "Side Table 08 B",
        "sidetable08b": "Side Table 08 B",
        "sidetable31": "Side Table 31 C",
        "sidetable31b": "Side Table 31 C",
        "sidetable31c": "Side Table 31 C",
        "sidetable42": "Side Table 42 A",
        "sidetable42b": "Side Table 42 A",
        "sidetable42a": "Side Table 42 A",
        "storageunit29": "Storage Unit 29 C",
        "storageunit29c": "Storage Unit 29 C",
        "storageunit34": "Storage Unit 34 A",
        "storageunit34a": "Storage Unit 34 A",
        "studytable02": "Study Table 02 A",
        "studytable02a": "Study Table 02 A",
        "swing04": "Swing 04 A",
        "swing04a": "Swing 04 A",
        "tvunit09a": "TV Unit 09 R2",
        "tvunit09r2": "TV Unit 09 R2",
        "tvunit29": "TV Unit 29 A R2",
        "tvunit29ar2": "TV Unit 29 A R2",
    }

    def make_keyplan_svg(room_id):
        # Precise coordinates mapped for each room on the 1980x1980 plan key grid
        highlights = {
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
            "room-11": 'x="1130" y="720" width="300" height="290"'
        }
        hl = highlights.get(room_id, '')
        if not hl:
            return f"""<svg viewBox="0 0 1980 1980" fill="none" xmlns="http://www.w3.org/2000/svg">
                          <image href="Layout Simplified 3b with labels.png" x="0" y="0" width="1980" height="1980" />
                        </svg>"""
        return f"""<svg viewBox="0 0 1980 1980" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <image href="Layout Simplified 3b with labels.png" x="0" y="0" width="1980" height="1980" />
                      <rect {hl} fill="#a27b5c" fill-opacity="0.4" stroke="#a27b5c" stroke-width="15" />
                    </svg>"""

    # 7. Generate Index Page HTML (Using the Image on the left side)
    def generate_index_page():
        list_rows = ""
        num = 1
        for r_conf in rooms_config:
            rid = r_conf["id"]
            if rid == "room-materials":
                continue  # Exclude Materials & Moodboard from the index list
            num += 1
            idx_str = f"{num:02d}"
            title = r_conf["title"]
            label = r_conf["label"]
            
            list_rows += f"""          <!-- Row: {title} -->
          <div class="index-row" id="index-row-{rid}" onclick="scrollToRoom('{rid}')" onmouseover="highlightRoom('{rid}')" onmouseout="unhighlightRoom('{rid}')">
            <div class="index-row-left">
              <span class="index-row-num">{idx_str}</span>
              <span class="index-row-name">{title}</span>
            </div>
            <span class="index-row-label">{label}</span>
          </div>\n"""
          
        return f"""    <!-- ==========================================================================
         ROOM 01: INDEX PAGE (PLAN KEY)
         ========================================================================== -->
    <main class="room-section index-section" id="room-00" data-room-title="Plan Key">
      
      <!-- Left Column: Title + Blueprint Image -->
      <section class="left-column">
        <div class="space-title-container">
          <span class="space-number">01</span>
          <h1 class="space-title">Plan Key</h1>
        </div>
        <div class="blueprint-large-container" style="padding: 0; overflow: hidden; background: none; border: none; box-shadow: none;">
          <svg class="blueprint-svg-large" viewBox="0 0 1980 1980" xmlns="http://www.w3.org/2000/svg" style="border-radius: var(--radius-card); border: var(--border-width) solid var(--line);">
            <!-- Background blueprint plan -->
            <image href="Layout Simplified 3b with labels.png" x="0" y="0" width="1980" height="1980" />
            <!-- Interactive Room Hotspots -->
            <g class="svg-rooms-group">
              <!-- Foyer -->
              <rect class="svg-room-rect" id="svg-rect-room-01" x="442" y="530" width="148" height="160" onclick="scrollToRoom('room-01')" onmouseover="highlightRoom('room-01')" onmouseout="unhighlightRoom('room-01')" />
              <!-- Living Room 01 -->
              <rect class="svg-room-rect" id="svg-rect-room-02" x="590" y="360" width="310" height="330" onclick="scrollToRoom('room-02')" onmouseover="highlightRoom('room-02')" onmouseout="unhighlightRoom('room-02')" />
              <!-- Living Room 02 -->
              <rect class="svg-room-rect" id="svg-rect-room-03" x="900" y="360" width="160" height="420" onclick="scrollToRoom('room-03')" onmouseover="highlightRoom('room-03')" onmouseout="unhighlightRoom('room-03')" />
              <!-- Living Room 03 -->
              <rect class="svg-room-rect" id="svg-rect-room-04" x="1060" y="360" width="120" height="100" onclick="scrollToRoom('room-04')" onmouseover="highlightRoom('room-04')" onmouseout="unhighlightRoom('room-04')" />
              <!-- Master Bedroom -->
              <rect class="svg-room-rect" id="svg-rect-room-05" x="1300" y="360" width="330" height="240" onclick="scrollToRoom('room-05')" onmouseover="highlightRoom('room-05')" onmouseout="unhighlightRoom('room-05')" />
              <!-- Bedroom 01 -->
              <rect class="svg-room-rect" id="svg-rect-room-06" x="1060" y="460" width="210" height="260" onclick="scrollToRoom('room-06')" onmouseover="highlightRoom('room-06')" onmouseout="unhighlightRoom('room-06')" />
              <!-- Dining -->
              <rect class="svg-room-rect" id="svg-rect-room-07" x="590" y="690" width="270" height="220" onclick="scrollToRoom('room-07')" onmouseover="highlightRoom('room-07')" onmouseout="unhighlightRoom('room-07')" />
              <!-- Bedroom 02 -->
              <rect class="svg-room-rect" id="svg-rect-room-08" x="442" y="910" width="218" height="300" onclick="scrollToRoom('room-08')" onmouseover="highlightRoom('room-08')" onmouseout="unhighlightRoom('room-08')" />
              <!-- Bedroom 03 -->
              <rect class="svg-room-rect" id="svg-rect-room-09" x="660" y="1010" width="200" height="250" onclick="scrollToRoom('room-09')" onmouseover="highlightRoom('room-09')" onmouseout="unhighlightRoom('room-09')" />
              <!-- Bar & Lounge -->
              <rect class="svg-room-rect" id="svg-rect-room-10" x="860" y="720" width="270" height="290" onclick="scrollToRoom('room-10')" onmouseover="highlightRoom('room-10')" onmouseout="unhighlightRoom('room-10')" />
              <!-- Outdoor -->
              <rect class="svg-room-rect" id="svg-rect-room-11" x="1130" y="720" width="300" height="290" onclick="scrollToRoom('room-11')" onmouseover="highlightRoom('room-11')" onmouseout="unhighlightRoom('room-11')" />
            </g>
          </svg>
        </div>
      </section>

      <!-- Right Column: List Key Table -->
      <section class="right-column index-right-column" style="overflow-y: hidden;">
        <div class="index-meta-header">
          <span class="index-label">01 · PLAN KEY</span>
          <h2 class="index-heading">Eleven rooms,<br>one walkthrough.</h2>
        </div>
        
        <div class="index-list-container">
{list_rows}        </div>
      </section>
    </main>\n"""

    # 8. Loop through all 11 rooms and build their HTML content
    print("\nProcessing room folders and compiling cards...")
    html_blocks = []
    js_room_renders = {}
    active_indices = {"room-00": 0}

    # Add Index page block first
    html_blocks.append(generate_index_page())

    # Default SVG Layout details for room views
    layout_svg = """<svg viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <!-- Outer Walls -->
                      <path d="M10 10 H90 V90 H10 Z" stroke="#D8D2C8" stroke-width="1.5" />
                      <path d="M10 10 V90 M90 10 V90" stroke="#6B6258" stroke-width="1" />
                      <!-- Bathroom Divider -->
                      <path d="M68 10 V90" stroke="#6B6258" stroke-width="1" stroke-dasharray="2 2" />
                      <!-- Bedroom door swing -->
                      <path d="M10 75 A15 15 0 0 1 25 90" stroke="#a27b5c" stroke-width="0.75" />
                      <line x1="25" y1="90" x2="25" y2="75" stroke="#a27b5c" stroke-width="0.75" />
                      <!-- Bed outline (Four Poster) -->
                      <rect x="25" y="30" width="34" height="40" fill="none" stroke="#a27b5c" stroke-width="1" />
                      <!-- Pillows -->
                      <rect x="29" y="32" width="11" height="7" rx="1" fill="none" stroke="#6B6258" stroke-width="0.75" />
                      <rect x="44" y="32" width="11" height="7" rx="1" fill="none" stroke="#6B6258" stroke-width="0.75" />
                      <!-- Side tables -->
                      <rect x="17" y="30" width="8" height="8" fill="none" stroke="#6B6258" stroke-width="0.75" />
                      <rect x="59" y="30" width="8" height="8" fill="none" stroke="#6B6258" stroke-width="0.75" />
                      <!-- Study Table and Chair -->
                      <rect x="25" y="80" width="22" height="10" fill="none" stroke="#6B6258" stroke-width="0.75" />
                      <circle cx="36" cy="74" r="3.5" fill="none" stroke="#6B6258" stroke-width="0.75" />
                    </svg>"""

    for r_idx, r_conf in enumerate(rooms_config):
        rid = r_conf["id"]
        folder = r_conf["folder"]
        title = r_conf["title"]
        quote = r_conf["quote"]
        if rid == "room-materials":
            idx_str = "00"
        else:
            r_idx_num = r_idx + 1
            idx_str = f"{r_idx_num:02d}"
        active_indices[rid] = 0
        
        if rid == "room-materials":
            # 1. Custom materials grid HTML (Categories & Circular Swatches) — auto-updated from disk
            materials_html = """          <div class="materials-categories">
            <!-- Category: Wood -->
            <div class="material-category-group">
              <h4 class="material-category-title">Wood</h4>
              <div class="material-swatches">
                <div class="material-swatch" onclick="openLightbox('room-materials', 0)">
                  <div class="swatch-circle">
                    <img src="Materials/Wood/Teak.png" alt="Burma Teak">
                  </div>
                  <span class="swatch-name">Burma Teak</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 1)">
                  <div class="swatch-circle">
                    <img src="Materials/Wood/Reclaimed teak.jpg" alt="Reclaimed Teak">
                  </div>
                  <span class="swatch-name">Reclaimed Teak</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 2)">
                  <div class="swatch-circle">
                    <img src="Materials/Wood/White_Ash_Crown.jpg" alt="White Ash">
                  </div>
                  <span class="swatch-name">White Ash</span>
                </div>
              </div>
            </div>

            <!-- Category: Leather -->
            <div class="material-category-group">
              <h4 class="material-category-title">Leather</h4>
              <div class="material-swatches">
                <div class="material-swatch" onclick="openLightbox('room-materials', 3)">
                  <div class="swatch-circle">
                    <img src="Materials/leather/Vagabond Cognac.jpeg" alt="Vagabond Cognac">
                  </div>
                  <span class="swatch-name">Vagabond Cognac</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 4)">
                  <div class="swatch-circle">
                    <img src="Materials/leather/Montana Chestnut.jpg" alt="Montana Chestnut">
                  </div>
                  <span class="swatch-name">Montana Chestnut</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 5)">
                  <div class="swatch-circle">
                    <img src="Materials/leather/Emperor Brick.jpeg" alt="Emperor Brick">
                  </div>
                  <span class="swatch-name">Emperor Brick</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 6)">
                  <div class="swatch-circle">
                    <img src="Materials/leather/Eternity Olive.jpeg" alt="Eternity Olive">
                  </div>
                  <span class="swatch-name">Eternity Olive</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 7)">
                  <div class="swatch-circle">
                    <img src="Materials/leather/Glory Honey.jpeg" alt="Glory Honey">
                  </div>
                  <span class="swatch-name">Glory Honey</span>
                </div>
              </div>
            </div>

            <!-- Category: Fabric -->
            <div class="material-category-group">
              <h4 class="material-category-title">Fabric</h4>
              <div class="material-swatches">
                <div class="material-swatch" onclick="openLightbox('room-materials', 8)">
                  <div class="swatch-circle">
                    <img src="Materials/Fabric/Vienna Army.jpg" alt="Vienna Army">
                  </div>
                  <span class="swatch-name">Vienna Army</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 9)">
                  <div class="swatch-circle">
                    <img src="Materials/Fabric/Rubik Linen.jpg" alt="Rubik Linen">
                  </div>
                  <span class="swatch-name">Rubik Linen</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 10)">
                  <div class="swatch-circle">
                    <img src="Materials/Fabric/Blush.jpeg" alt="Blush">
                  </div>
                  <span class="swatch-name">Blush</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 11)">
                  <div class="swatch-circle">
                    <img src="Materials/Fabric/Cloud.jpg" alt="Cloud">
                  </div>
                  <span class="swatch-name">Cloud</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 12)">
                  <div class="swatch-circle">
                    <img src="Materials/Fabric/Charcoal.png" alt="Charcoal">
                  </div>
                  <span class="swatch-name">Charcoal</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 13)">
                  <div class="swatch-circle">
                    <img src="Materials/Fabric/Flute.jpeg" alt="Flute">
                  </div>
                  <span class="swatch-name">Flute</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 14)">
                  <div class="swatch-circle">
                    <img src="Materials/Fabric/Opal.png" alt="Opal">
                  </div>
                  <span class="swatch-name">Opal</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 15)">
                  <div class="swatch-circle">
                    <img src="Materials/Fabric/Rosebud.png" alt="Rosebud">
                  </div>
                  <span class="swatch-name">Rosebud</span>
                </div>
                <div class="material-swatch" onclick="openLightbox('room-materials', 16)">
                  <div class="swatch-circle">
                    <img src="Materials/Fabric/Silver.jpeg" alt="Silver">
                  </div>
                  <span class="swatch-name">Silver</span>
                </div>
              </div>
            </div>

            <!-- Category: Metal -->
            <div class="material-category-group">
              <h4 class="material-category-title">Metal</h4>
              <div class="material-swatches">
                <div class="material-swatch" onclick="openLightbox('room-materials', 17)">
                  <div class="swatch-circle">
                    <img src="Materials/Metal/Brass.jpg" alt="Brass">
                  </div>
                  <span class="swatch-name">Brass</span>
                </div>
              </div>
            </div>
          </div>"""

            # 2. Custom moodboard asymmetric grid HTML (3-column layout)
            moodboard_html = """          <div class="moodboard-grid">
            <div class="moodboard-col">
              <div class="moodboard-item tall" onclick="openLightbox('moodboard', 0)">
                <img src="Moodboard - of stillness/57ec27a5377b392bd27f4250d5c39461.jpg" alt="Umang Moodboard 1">
              </div>
              <div class="moodboard-item short" onclick="openLightbox('moodboard', 1)">
                <img src="Moodboard - of stillness/01.jpg" alt="Umang Moodboard 2">
              </div>
            </div>
            <div class="moodboard-col">
              <div class="moodboard-item short" onclick="openLightbox('moodboard', 2)">
                <img src="Moodboard - of stillness/fdf0f8ab8c964cf68f39f626d0922461.jpg" alt="Umang Moodboard 3">
              </div>
              <div class="moodboard-item tall" onclick="openLightbox('moodboard', 3)">
                <img src="Moodboard - of stillness/1d1c1af230ec5bf3fed2f9ac7fd9fdf5.jpg" alt="Umang Moodboard 4">
              </div>
            </div>
            <div class="moodboard-col">
              <div class="moodboard-item tall" onclick="openLightbox('moodboard', 4)">
                <img src="Moodboard - of stillness/wood.jpg" alt="Umang Moodboard 5">
              </div>
              <div class="moodboard-item short" onclick="openLightbox('moodboard', 5)">
                <img src="Moodboard - of stillness/80becd3aab10a0f715442c912fa0600b.jpg" alt="Umang Moodboard 6">
              </div>
            </div>
          </div>"""

            # 3. Combine into the room block
            room_block = f"""    <!-- ==========================================================================
         ROOM {idx_str}: {title.upper()}
         ========================================================================== -->
    <main class="room-section" id="{rid}" data-room-title="{title}">
      
      <!-- Left Column: Title + Materials Grid -->
      <section class="left-column">
        <div class="space-title-container">
          <span class="space-number">{idx_str}</span>
          <h1 class="space-title">{title}</h1>
        </div>
        <div class="showcase-container" style="overflow: hidden; max-height: calc(100vh - 160px);">
{materials_html}
        </div>
      </section>

      <!-- Right Column: Moodboard Header + Asymmetric Grid -->
      <section class="right-column" style="overflow-y: auto; padding-right: 5px;">
        <!-- Top Row: Quote -->
        <div class="room-meta-header" style="grid-template-columns: 1fr;">
          <blockquote class="quote-container">
            "{quote}"
          </blockquote>
        </div>

        <!-- Bottom Row: Asymmetric Moodboard Grid -->
        <div class="products-section-container" style="height: 100%; display: flex; flex-direction: column;">
          <h2 class="section-label">Collection Moodboard</h2>
          <div style="flex-grow: 1; min-height: 0;">
{moodboard_html}
          </div>
        </div>
      </section>
    </main>\n"""
            html_blocks.append(room_block)
            
            # Setup lightbox data for these two keys
            js_room_renders["room-materials"] = [
                # Wood (0-2)
                {"src": "Materials/Wood/Teak.png", "title": "Burma Teak", "sub": "Wood"},
                {"src": "Materials/Wood/Reclaimed teak.jpg", "title": "Reclaimed Teak", "sub": "Wood"},
                {"src": "Materials/Wood/White_Ash_Crown.jpg", "title": "White Ash", "sub": "Wood"},
                # Leather (3-7)
                {"src": "Materials/leather/Vagabond Cognac.jpeg", "title": "Vagabond Cognac", "sub": "Leather"},
                {"src": "Materials/leather/Montana Chestnut.jpg", "title": "Montana Chestnut", "sub": "Leather"},
                {"src": "Materials/leather/Emperor Brick.jpeg", "title": "Emperor Brick", "sub": "Leather"},
                {"src": "Materials/leather/Eternity Olive.jpeg", "title": "Eternity Olive", "sub": "Leather"},
                {"src": "Materials/leather/Glory Honey.jpeg", "title": "Glory Honey", "sub": "Leather"},
                # Fabric (8-16)
                {"src": "Materials/Fabric/Vienna Army.jpg", "title": "Vienna Army", "sub": "Fabric"},
                {"src": "Materials/Fabric/Rubik Linen.jpg", "title": "Rubik Linen", "sub": "Fabric"},
                {"src": "Materials/Fabric/Blush.jpeg", "title": "Blush", "sub": "Fabric"},
                {"src": "Materials/Fabric/Cloud.jpg", "title": "Cloud", "sub": "Fabric"},
                {"src": "Materials/Fabric/Charcoal.png", "title": "Charcoal", "sub": "Fabric"},
                {"src": "Materials/Fabric/Flute.jpeg", "title": "Flute", "sub": "Fabric"},
                {"src": "Materials/Fabric/Opal.png", "title": "Opal", "sub": "Fabric"},
                {"src": "Materials/Fabric/Rosebud.png", "title": "Rosebud", "sub": "Fabric"},
                {"src": "Materials/Fabric/Silver.jpeg", "title": "Silver", "sub": "Fabric"},
                # Metal (17)
                {"src": "Materials/Metal/Brass.jpg", "title": "Brass", "sub": "Metal"}
            ]
            js_room_renders["moodboard"] = [
                {"src": "Moodboard - of stillness/57ec27a5377b392bd27f4250d5c39461.jpg", "title": "Umang Moodboard 1", "sub": "Collection Moodboard"},
                {"src": "Moodboard - of stillness/01.jpg", "title": "Umang Moodboard 2", "sub": "Collection Moodboard"},
                {"src": "Moodboard - of stillness/fdf0f8ab8c964cf68f39f626d0922461.jpg", "title": "Umang Moodboard 3", "sub": "Collection Moodboard"},
                {"src": "Moodboard - of stillness/1d1c1af230ec5bf3fed2f9ac7fd9fdf5.jpg", "title": "Umang Moodboard 4", "sub": "Collection Moodboard"},
                {"src": "Moodboard - of stillness/wood.jpg", "title": "Umang Moodboard 5", "sub": "Collection Moodboard"},
                {"src": "Moodboard - of stillness/80becd3aab10a0f715442c912fa0600b.jpg", "title": "Umang Moodboard 6", "sub": "Collection Moodboard"}
            ]
            continue
            
        rpath = os.path.join(base_dir, folder)
        
        # Determine layout image details
        layout_images = {
            "room-01": {"path": "Room wise layout/Foyer.png", "title": "Foyer Layout"},
            "room-02": {"path": "Room wise layout/Living.png", "title": "Living Room 01 Layout"},
            "room-03": {"path": "Room wise layout/Family.png", "title": "Living Room 02 Layout"},
            "room-04": {"path": "Room wise layout/Family.png", "title": "Living Room 03 Layout"},
            "room-05": {"path": "Room wise layout/Master Bedroom.png", "title": "Master Bedroom Layout"},
            "room-06": {"path": "Room wise layout/Bedroom 1.png", "title": "Bedroom 01 Layout"},
            "room-07": {"path": "Room wise layout/Dining.png", "title": "Dining Layout"},
            "room-08": {"path": "Room wise layout/Bedroom 2.png", "title": "Bedroom 02 Layout"},
            "room-09": {"path": "Room wise layout/Bedroom 3.png", "title": "Bedroom 03 Layout"},
            "room-10": {"path": "Room wise layout/Bar & Lounge.png", "title": "Bar & Lounge Layout"},
            "room-11": {"path": "Room wise layout/Balcony.png", "title": "Outdoor Layout"}
        }
        layout_info = layout_images.get(rid, {"path": "", "title": ""})
        layout_img_path = layout_info["path"]
        layout_img_title = layout_info["title"]
        
        # Load main renders
        renders = []
        if os.path.exists(rpath):
            for item in sorted(os.listdir(rpath)):
                if os.path.isfile(os.path.join(rpath, item)) and item.lower().endswith(('.png', '.jpg', '.jpeg')):
                    renders.append(item)
                        
        # Load products
        products = []
        prod_dir = os.path.join(rpath, "Products")
        if os.path.exists(prod_dir):
            for pitem in sorted(os.listdir(prod_dir)):
                if pitem.lower().endswith(('.png', '.jpg', '.jpeg')):
                    img_norm = normalize_name(pitem)
                    best_match = None
                    best_score = 0
                    
                    for db_p in db_products:
                        db_norm_name = normalize_name(db_p["name"])
                        db_norm_code = normalize_name(db_p["code"])
                        
                        if img_norm == db_norm_name or img_norm == db_norm_code:
                            best_match = db_p
                            break
                        if img_norm in db_norm_name or db_norm_name in img_norm:
                            score = min(len(img_norm), len(db_norm_name)) / max(len(img_norm), len(db_norm_name))
                            if score > best_score:
                                best_score = score
                                best_match = db_p
                                
                    dim_label = ""
                    if best_match:
                        w, d, h = best_match["w"], best_match["d"], best_match["h"]
                        sh, mh = best_match["sh"], best_match["mh"]
                        dim_parts = []
                        if w and d and h:
                            dim_parts.append(f"{w} x {d} x {h} cm (H)")
                        elif best_match["dim_str"]:
                            dim_parts.append(best_match["dim_str"])
                        if sh:
                            dim_parts.append(f"SH-{sh}cm")
                        if mh:
                            dim_parts.append(f"MH-{mh}cm")
                        dim_label = " &nbsp;·&nbsp; ".join(dim_parts)
                        prod_name = best_match["name"]
                    else:
                        prod_name = re.sub(r'\.png$|\.jpg$|\.jpeg$', '', pitem)
                        prod_name = re.sub(r'\bv2\b|\bv1\b|\bfull image\b', '', prod_name).strip()
                        dim_label = "" # Leave empty if missing
                    
                    # Map rates
                    rate_key = None
                    for k, v in custom_rates_mapping.items():
                        if img_norm == normalize_name(k) or (best_match and normalize_name(best_match["name"]) == normalize_name(k)) or (best_match and normalize_name(best_match["code"]) == normalize_name(k)):
                            rate_key = v
                            break
                    
                    if not rate_key:
                        for r_key in rates_db.keys():
                            r_norm = normalize_name(r_key)
                            if img_norm == r_norm or (best_match and normalize_name(best_match["name"]) == r_norm) or (best_match and normalize_name(best_match["code"]) == r_norm):
                                rate_key = r_key
                                break
                    
                    rate_val = rates_db.get(rate_key) if rate_key else None
                    
                    products.append({
                        "img_path": f"{folder}/Products/{pitem}",
                        "name": prod_name,
                        "dimensions": dim_label,
                        "rate": rate_val
                    })

        # JS Render List
        js_renders = []
        for r in renders:
            js_renders.append({
                "src": f"{folder}/{r}",
                "title": f"{title} Render",
                "sub": "Umang Project"
            })
        js_room_renders[rid] = js_renders

        # Slider HTML
        if len(renders) > 1:
            slider_html = f"""          <div class="main-render-wrapper" onclick="openLightbox('{rid}', activeRenderIndex['{rid}'])">
            <!-- Navigation Arrows on Slider -->
            <button class="slider-arrow slider-prev" onclick="navigateRender('{rid}', -1); event.stopPropagation();">&#8249;</button>
            <button class="slider-arrow slider-next" onclick="navigateRender('{rid}', 1); event.stopPropagation();">&#8250;</button>
            
            <img src="{folder}/{renders[0]}" alt="{title} Render 1" class="main-render-image">
          </div>"""
        elif len(renders) == 1:
            slider_html = f"""          <div class="main-render-wrapper" onclick="openLightbox('{rid}', 0)">
            <img src="{folder}/{renders[0]}" alt="{title} Render" class="main-render-image">
          </div>"""
        else:
            slider_html = f"""          <div class="main-render-wrapper">
            <div style="width:100%; height:100%; display:flex; flex-direction:column; justify-content:center; align-items:center; background:var(--cream); color:var(--muted);">
              <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.2" style="margin-bottom:10px;">
                <rect x="3" y="3" width="18" height="18" rx="2" />
                <circle cx="8.5" cy="8.5" r="1.5" />
                <path d="M21 15l-5-5L5 21" />
              </svg>
              <span style="font-family:var(--font-display); font-size: 0.95rem; font-weight:600; letter-spacing:0.05em; text-transform:uppercase;">{title} Renders Placeholder</span>
            </div>
          </div>"""

        # Products cards HTML
        products_html = ""
        if products:
            for p_idx, p in enumerate(products):
                p_name_escaped = html.escape(p["name"])
                
                rate_str = ""
                rate_lbl = ""
                if p.get("rate"):
                    rate_str = f"  —  ₹ {p['rate']:,}"
                    rate_lbl = f"""\n                  <span class="product-card__price" style="font-family: var(--font-display); font-size: 0.82rem; font-weight: 600; color: var(--chamoisee); white-space: nowrap;">₹ {p['rate']:,}</span>"""

                if p["dimensions"]:
                    details_label = f"Teak / White Ash &nbsp;·&nbsp; {p['dimensions']}"
                    lightbox_details = f"Teak / White Ash — {p['dimensions'].replace('&nbsp;·&nbsp;', ' — ')}{rate_str}"
                else:
                    details_label = "Teak / White Ash"
                    lightbox_details = f"Teak / White Ash{rate_str}"
                
                lightbox_details_escaped = html.escape(lightbox_details)
                
                products_html += f"""            <!-- Product Card: {p['name']} -->
            <div class="product-card">
              <div class="product-image-wrapper" onclick="openLightbox('{rid}-prod-{p_idx}', 0, '{p['img_path']}', '{p_name_escaped}', '{lightbox_details_escaped}')">
                <img src="{p['img_path']}" alt="{p_name_escaped}" class="product-image">
              </div>
              <div class="product-card__body">
                <div style="display: flex; justify-content: space-between; align-items: baseline; gap: 8px;">
                  <h3 class="product-card__name" style="flex-grow: 1; margin-bottom: 0;">{p['name']}</h3>{rate_lbl}
                </div>
                <span class="product-card__material">{details_label}</span>
              </div>
            </div>\n"""
        else:
            products_html = f"""            <div style="grid-column: span 2; font-family:var(--font-body); font-size:0.85rem; color:var(--muted); font-style:italic; padding: 40px 0; text-align:center; border: 1px dashed var(--line); border-radius: var(--radius-card); background: var(--light-secondary);">
              Custom bespoke furniture. Details TBD.
            </div>\n"""

        keyplan_svg = make_keyplan_svg(rid)

        room_block = f"""    <!-- ==========================================================================
         ROOM {idx_str}: {title.upper()}
         ========================================================================== -->
    <main class="room-section" id="{rid}" data-room-title="{title}">
      
      <!-- Left Column: Title + Main Renders -->
      <section class="left-column">
        <div class="space-title-container">
          <span class="space-number">{idx_str}</span>
          <h1 class="space-title">{title}</h1>
        </div>

        <div class="showcase-container">
{slider_html}
        </div>
      </section>

      <!-- Right Column: Quote + Maps + Products Grid -->
      <section class="right-column">
        <!-- Top Row: Quote & Skeletal Maps -->
        <div class="room-meta-header">
          <blockquote class="quote-container">
            "{quote}"
          </blockquote>
          
          <div class="maps-container">
            <!-- Detailed Layout -->
            <div class="map-placeholder-card" title="Detailed Room Layout" onclick="openLightbox(null, 0, '{layout_img_path}', '{layout_img_title}', 'Room wise layout')">
              <div class="map-svg-wrapper">
                <img src="{layout_img_path}" alt="{layout_img_title}" class="map-layout-image">
              </div>
              <span class="map-label">Layout</span>
            </div>
            
            <!-- Key Plan (larger, no label) -->
            <div class="map-keyplan-large" title="Key House Plan — click to return to Plan Key" onclick="scrollToRoom('room-00')">
              <div class="map-svg-wrapper">
                {keyplan_svg}
              </div>
            </div>
          </div>
        </div>

        <!-- Bottom Row: Products Grid -->
        <div class="products-section-container">
          <h2 class="section-label">Furniture Used</h2>
          
          <div class="products-grid">
{products_html}          </div>
        </div>
      </section>
    </main>\n"""
        html_blocks.append(room_block)

    # 8b. Generate Customisation Page (last section)
    def generate_customisation_page():
        return """    <!-- ==========================================================================
         CUSTOMISATION: MADE FOR YOU
         ========================================================================== -->
    <main class="room-section custom-section-new" id="room-custom" data-room-title="Customisation">

      <!-- Top Row: Split Header -->
      <div class="custom-header-row">
        <div class="custom-header-left">
          <div class="custom-brand-group">
            <img src="Noku mark.png" alt="N" class="custom-brand-logo">
            <div class="custom-brand-text">
              <span class="custom-brand-title">CUSTOMISATION</span>
              <span class="custom-brand-sub">13 &nbsp;·&nbsp; MADE FOR YOU</span>
            </div>
          </div>
          <h1 class="custom-hero-heading">Every piece, tailored<br>to the residence.</h1>
        </div>
        
        <div class="custom-header-right">
          <div class="custom-chapter-tag">Chapter Four &nbsp;·&nbsp; Tailoring</div>
          <blockquote class="custom-description">
            Each piece in this proposal can be tuned for fabric, finish, wood, dimension and joinery to fit Umang specifically. Some examples:
          </blockquote>
        </div>
      </div>

      <!-- Bottom Row: 4-Column Cards Grid -->
      <div class="custom-cards-row">
        <div class="custom-cards-grid-new">

          <!-- Card 1: Wood -->
          <div class="custom-card-new" title="Wood Customisations">
            <div class="custom-card-new__image-wrapper">
              <img src="Materials/Wood/Teak.png" alt="Wood Options" class="custom-card-new__image">
            </div>
            <div class="custom-card-new__body">
              <span class="custom-card-new__tag">Wood</span>
              <h3 class="custom-card-new__title">Teak, Reclaimed Teak, White Ash.</h3>
              <p class="custom-card-new__desc">Mix-and-match across pieces. Reclaimed available subject to availability.</p>
            </div>
          </div>

          <!-- Card 2: Upholstery -->
          <div class="custom-card-new" title="Upholstery Options">
            <div class="custom-card-new__image-wrapper">
              <img src="Materials/Fabric/Rubik Linen.jpg" alt="Upholstery Options" class="custom-card-new__image">
            </div>
            <div class="custom-card-new__body">
              <span class="custom-card-new__tag">Upholstery</span>
              <h3 class="custom-card-new__title">Fabric or Italian Leather.</h3>
              <p class="custom-card-new__desc">Library of 40+ swatches across our partners &mdash; Vagabond, D&rsquo;Decor, Rubik.</p>
            </div>
          </div>

          <!-- Card 3: Dimensions -->
          <div class="custom-card-new custom-card-new--dim" title="Dimension Customisations">
            <div class="custom-card-new__image-wrapper custom-card-new__image-wrapper--dim">
              <svg viewBox="0 0 120 120" fill="none" xmlns="http://www.w3.org/2000/svg" class="custom-dim-svg">
                <rect x="35" y="25" width="50" height="50" stroke="#a27b5c" stroke-width="1.2"/>
                <line x1="35" y1="25" x2="85" y2="75" stroke="#a27b5c" stroke-width="1.0"/>
                <line x1="85" y1="25" x2="35" y2="75" stroke="#a27b5c" stroke-width="1.0"/>
                <text x="60" y="96" text-anchor="middle" font-family="Josefin Sans, sans-serif" font-size="8.5" fill="#a27b5c" letter-spacing="1">DIM &plusmn;10%</text>
              </svg>
            </div>
            <div class="custom-card-new__body">
              <span class="custom-card-new__tag">Dimensions</span>
              <h3 class="custom-card-new__title">Scale to your space.</h3>
              <p class="custom-card-new__desc">Up to &plusmn;10% on most pieces with no tooling surcharge. Larger changes treated as a new study.</p>
            </div>
          </div>

        </div><!-- /custom-cards-grid-new -->
      </div><!-- /custom-cards-row -->
    </main>\n"""

    html_blocks.append(generate_customisation_page())

    # 9. Load current index.html and merge
    print("\nMerging blocks into index.html...")
    with open(html_path, 'r', encoding='utf-8') as f:
        orig_html = f.read()

    # Find header and lightbox boundaries
    header_end = orig_html.find('</header>')
    if header_end != -1:
        header_end += 9
        lightbox_start = orig_html.find('<!-- ==========================================================================\n       FULLSCREEN CINEMATIC LIGHTBOX MODAL')
        if lightbox_start == -1:
            lightbox_start = orig_html.find('<!-- ==========================================================================\r\n       FULLSCREEN CINEMATIC LIGHTBOX MODAL')
            
        if lightbox_start != -1:
            # Furniture Layout Section (Unnumbered)
            furniture_layout_block = """    <!-- ==========================================================================
         FURNITURE LAYOUT: CENTERED PLAN (UNNUMBERED)
         ========================================================================== -->
    <main class="room-section furniture-layout-section" id="room-layout" data-room-title="Furniture Layout">
      
      <!-- Top Title Bar -->
      <div class="layout-header-row">
        <div class="space-title-container">
          <h1 class="space-title">Furniture Layout</h1>
        </div>
      </div>

      <!-- Centered Image Container -->
      <div class="layout-image-container-centered" onclick="openLightbox(null, 0, 'Layout Original 6.png', 'Furniture Layout')">
        <img src="Layout Original 6.png" alt="Furniture Layout" class="furniture-layout-image">
      </div>

    </main>"""

            # Reorder so that Materials & Moodboard (index 1) is the first slide, followed by Plan Key (index 0), then Furniture Layout
            ordered_blocks = [html_blocks[1], html_blocks[0], furniture_layout_block] + html_blocks[2:]
            gen_rooms_str = "\n".join(ordered_blocks)
            # Scroll-to-top button (inserted just before the lightbox)
            scroll_top_html = """
  <!-- Scroll to Top Button -->
  <button id="scroll-top-btn" onclick="scrollToRoom('room-00')" title="Back to Plan Key">
    <svg viewBox="0 0 24 24"><polyline points="18 15 12 9 6 15"/></svg>
  </button>

  """
            thank_you_html = """  <!-- ======================================================================
       SECTION 4: THANK YOU (dark)
       ====================================================================== -->
  <section class="room-section landing-thankyou-section" id="thank-you" aria-label="Thank You">
    <!-- Background panel with subtle image and heavy dark vignette -->
    <div class="ty-bg-panel" aria-hidden="true">
      <img src="02. Living Room 01/Living Room 01 v1.png" alt="">
    </div>
    <div class="ty-bg-vignette" aria-hidden="true"></div>

    <!-- Side labels -->
    <div class="ty-side-ticker" aria-hidden="true"><span>Good Earth Kochi · Umang Residence · Volume One</span></div>
    <div class="ty-edition-mark" aria-hidden="true"><span>Noku Studio · N°1 · 2026</span></div>

    <!-- Nav -->
    <nav class="ty-nav" aria-label="Section navigation">
      <div class="ty-nav__brand">
        <img src="Noku mark.png" alt="Noku Studio" class="ty-nav__logo">
        <span class="ty-nav__wordmark">Noku Studio</span>
      </div>
      <span class="ty-nav__label">End &nbsp;·&nbsp; Walkthrough</span>
    </nav>

    <!-- Thank You content -->
    <div class="ty-stage">
      <div class="ty-eyebrow">
        <div class="ty-eyebrow-line"></div>
        <span class="ty-eyebrow-text">Volume One &nbsp;·&nbsp; Of Stillness</span>
        <div class="ty-eyebrow-line"></div>
      </div>
      <h1 class="ty-title" id="ty-title">Thank You.</h1>
      <div class="ty-divider"></div>
      <p class="ty-descriptor" id="ty-desc">
        A project by <strong>Noku Studio</strong> in collaboration with <strong>Good Earth</strong>.<br>
        Thank you for exploring our volume of stillness.
      </p>

      <!-- CTAs: Return to Plan Key -->
      <div class="ty-ctas" id="ty-ctas">
        <button class="ty-btn ty-btn--primary" onclick="scrollToRoom('room-00')">
          <span>Plan Key</span>
        </button>
      </div>
    </div>

    <!-- Footer strip -->
    <footer class="ty-footer">
      <div class="ty-footer__left">
        <span class="ty-footer__credit">NOKU STUDIO · BENGALURU</span>
        <span class="ty-footer__client">hello@nokustudio.com &nbsp;·&nbsp; nokustudio.com</span>
      </div>
      <div class="ty-footer__right">
        <span class="ty-pagination">30 &nbsp;/&nbsp; 30</span>
      </div>
    </footer>
  </section>"""
            middle_section = f"""\n\n    <!-- Room Snap Scroll Container -->
    <div class="scroll-container">
{gen_rooms_str}
{thank_you_html}
    </div>
  </div>\n\n  """
            orig_html = orig_html[:header_end] + middle_section + scroll_top_html + orig_html[lightbox_start:]
        else:
            print("Error: Could not find lightbox start.")
            return
    else:
        print("Error: Could not find header end.")
        return

    # Replace JS script variables
    script_start = orig_html.find('<script>')
    if script_start != -1:
        script_start += 8
        lightbox_group_start = orig_html.find('let currentLightboxGroup = null;')
        if lightbox_group_start != -1:
            active_indices["room-layout"] = 0
            js_data = f"""const roomRenders = {json.dumps(js_room_renders, indent=2)};\n\nconst activeRenderIndex = {json.dumps(active_indices, indent=2)};\n"""
            
            js_inject = f"""
    {js_data}
    function scrollToRoom(roomId) {{
      const target = document.getElementById(roomId);
      if (target) {{
        target.scrollIntoView({{ behavior: 'smooth' }});
      }}
    }}

    function highlightRoom(roomId) {{
      const row = document.getElementById('index-row-' + roomId);
      if (row) {{
        row.classList.add('highlighted');
      }}
      const rect = document.getElementById('svg-rect-' + roomId);
      if (rect) {{
        rect.classList.add('highlighted');
      }}
    }}

    function unhighlightRoom(roomId) {{
      const row = document.getElementById('index-row-' + roomId);
      if (row) {{
        row.classList.remove('highlighted');
      }}
      const rect = document.getElementById('svg-rect-' + roomId);
      if (rect) {{
        rect.classList.remove('highlighted');
      }}
    }}

    // Update active room badge in header based on scroll position
    document.addEventListener('DOMContentLoaded', () => {{
      const roomSections = document.querySelectorAll('.room-section');
      const badge = document.getElementById('active-room-badge');
      const scrollContainer = document.querySelector('.scroll-container');
      
      // Dynamic scroll enabling for rooms with more than 4 products
      roomSections.forEach(room => {{
        const productGrid = room.querySelector('.products-grid');
        if (productGrid) {{
          const productCount = productGrid.querySelectorAll('.product-card').length;
          const rightCol = room.querySelector('.right-column');
          if (rightCol && productCount > 4) {{
            rightCol.classList.add('has-scroll');
          }}
        }}
      }});
            const isDesktop = window.innerWidth > 1200;
      const observerOptions = {{
        root: isDesktop ? scrollContainer : null,
        rootMargin: '0px',
        threshold: 0.4
      }};

      const observer = new IntersectionObserver((entries) => {{
        entries.forEach(entry => {{
          if (entry.isIntersecting) {{
            const roomSection = entry.target;
            const idxEl = roomSection.querySelector('.space-number');
            const titleEl = roomSection.querySelector('.space-title');
            const header = document.querySelector('.app-header');
            
            if (roomSection.id === 'thank-you') {{
              if (header) {{
                header.style.opacity = '0';
                header.style.pointerEvents = 'none';
              }}
              document.body.style.backgroundColor = '#1E1915';
            }} else {{
              if (header) {{
                header.style.opacity = '1';
                header.style.pointerEvents = 'auto';
              }}
              document.body.style.backgroundColor = '#f4f5f1';
              if (idxEl && titleEl) {{
                badge.innerText = `${{idxEl.innerText}} ${{titleEl.innerText}}`;
              }} else if (roomSection.id === 'room-layout') {{
                badge.innerText = `Furniture Layout`;
              }} else if (roomSection.id === 'room-custom') {{
                badge.innerText = `13 Customisation`;
              }}
            }}
          }}
        }});
      }}, observerOptions);

      roomSections.forEach(sec => observer.observe(sec));

      // Scroll-to-top button visibility: show when user scrolls past the first section
      const scrollTopBtn = document.getElementById('scroll-top-btn');
      if (scrollTopBtn && scrollContainer) {{
        scrollContainer.addEventListener('scroll', () => {{
          // Show if scrolled past the first snap section (~80% of viewport height)
          const threshold = scrollContainer.clientHeight * 0.8;
          if (scrollContainer.scrollTop > threshold) {{
            scrollTopBtn.classList.add('visible');
          }} else {{
            scrollTopBtn.classList.remove('visible');
          }}
        }});
      }}
    }});

    """
            orig_html = orig_html[:script_start] + js_inject + orig_html[lightbox_group_start:]
        else:
            print("Error: Could not find lightbox group start in javascript.")
            return
    else:
        print("Error: Could not find script start.")
        return

    # Write back to index.html
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(orig_html)
        
    print("\nSuccess! index.html updated and matched with local folder structure.")
    print("Double click index.html to view the updated site.")

if __name__ == "__main__":
    main()
