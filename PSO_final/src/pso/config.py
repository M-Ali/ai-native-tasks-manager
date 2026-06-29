"""Central configuration: product classifications, column names, constants."""

# ── Source file ────────────────────────────────────────────────────────────────
HEADER_ROW = 1          # 0-indexed (row 2 in Excel = index 1 after read_excel skiprows)
DATA_SHEET = 0          # first sheet

# ── Column names (exactly as they appear in source) ───────────────────────────
COL_ORG       = "Sales _Org_ Desc.2"
COL_PRODUCT   = "ProductCategory"
COL_CATEGORY  = "Category"
COL_REGION    = "Sales office Region"
COL_CITY      = "City"
COL_CORP_GRP  = "Corporate Group"
COL_CUST_NUM  = "Customer Number"
COL_CUST_NAME = "Name 1"

COL_GRS_CY    = "SalesGRS_CY"
COL_GRS_LY    = "SalesGRS_LY"
COL_VOL_CY    = "SalesLtr_CY"
COL_VOL_LY    = "SalesLtr_LY"
COL_MT_CY     = "SalesMT_CY"
COL_MT_LY     = "SalesMT_LY"
COL_DISC_CY   = "Disc_CY"
COL_DISC_LY   = "Disc_LY"
COL_DISC_L_CY = "Discount/Ltr_CY"
COL_DISC_L_LY = "Discount/Ltr_LY"
COL_PMGN_CY   = "MarignPrimary_CY"
COL_PMGN_LY   = "MarignPrimary_LY"
COL_MGN_L_CY  = "Margin/Ltr_CY"
COL_MGN_L_LY  = "Margin/Ltr_LY"
COL_NMGN_CY   = "NetMargin_CY"
COL_NMGN_LY   = "NetMargin_LY"
COL_NMGN_L_CY = "NetMargin/Ltr_CY"
COL_NMGN_L_LY = "NetMargin/Ltr_LY"

# %SPLY source columns — pre-calculated in Excel (Same Period Last Year % change)
COL_PCT_SPLY_VOL  = "%SPLYSalesLtr"
COL_PCT_SPLY_GRS  = "%SPLYGRS"
COL_PCT_SPLY_NMGN = "%SPLYNetMargin"
COL_PCT_SPLY_DISC = "%SPLYDisc"
COL_PCT_SPLY_PMGN = "%SPLYPMargin"

# Derived SPLY absolute columns — computed by Ingest from CY ÷ (1 + %SPLY/100)
COL_VOL_SPLY  = "SalesLtr_SPLY"
COL_GRS_SPLY  = "SalesGRS_SPLY"
COL_NMGN_SPLY = "NetMargin_SPLY"
COL_DISC_SPLY = "Disc_SPLY"
COL_PMGN_SPLY = "MarignPrimary_SPLY"

REQUIRED_COLUMNS = [
    COL_ORG, COL_PRODUCT, COL_CATEGORY, COL_REGION, COL_CITY,
    COL_CORP_GRP, COL_CUST_NUM, COL_CUST_NAME,
    COL_GRS_CY, COL_GRS_LY, COL_VOL_CY, COL_VOL_LY,
    COL_DISC_CY, COL_DISC_LY, COL_DISC_L_CY, COL_DISC_L_LY,
    COL_PMGN_CY, COL_PMGN_LY, COL_MGN_L_CY, COL_MGN_L_LY,
    COL_NMGN_CY, COL_NMGN_LY, COL_NMGN_L_CY, COL_NMGN_L_LY,
]

ADDITIVE_METRICS = [
    COL_GRS_CY, COL_GRS_LY,
    COL_VOL_CY, COL_VOL_LY,
    COL_MT_CY,  COL_MT_LY,
    COL_DISC_CY, COL_DISC_LY,
    COL_PMGN_CY, COL_PMGN_LY,
    COL_NMGN_CY, COL_NMGN_LY,
]

# ── Derived column names (added by Ingest Agent) ──────────────────────────────
COL_FUEL_SEG   = "FuelSegment"     # Diesel / Petrol / Lubricants / Other Fuels / LPG / ...
COL_LUBE_CAT   = "LubeCategory"   # DEO / PCMO / MCO / LOW GRADE / Greases / ...
COL_CITY_NORM  = "CityNorm"        # normalized city name
COL_IS_RETAIL  = "IsRetail"        # True if Sales Org == 'Retail Business'
COL_IS_INTL    = "IsInternational" # True if city is outside Pakistan

# ── Sales channel ─────────────────────────────────────────────────────────────
RETAIL_ORG = "Retail Business"

# ── Product → Fuel Segment mapping ────────────────────────────────────────────
PRODUCT_SEGMENTS = {
    "HSD":                        "Diesel",
    "LDO":                        "Diesel",
    "SLUDGE-HSD":                 "Diesel",
    "SLUDGE":                     "Diesel",
    "PMG":                        "Petrol",
    "R95":                        "Petrol",
    "SKO":                        "Other Fuels",
    "FO":                         "Other Fuels",
    "JeT-A1":                     "Aviation Fuel",
    "LPG-PSO":                    "LPG",
    "LNG":                        "LNG",
    "Chemicals":                  "Chemicals",
    "Ethyl Glycol":               "Chemicals",
    "PSO Viscosity Index Improver": "Lubricants",
    "Lubes":                      "Lubricants",
}

DIESEL_PRODUCTS  = {"HSD", "LDO", "SLUDGE-HSD", "SLUDGE"}
PETROL_PRODUCTS  = {"PMG", "R95"}
LUBE_PRODUCTS    = {"Lubes", "PSO Viscosity Index Improver"}
FUEL_PRODUCTS    = DIESEL_PRODUCTS | PETROL_PRODUCTS

RETAIL_FUEL_SEGMENTS = {"Diesel", "Petrol", "Lubricants", "Other Fuels", "LPG"}

# ── Lubricant sub-categories (from Category column) ──────────────────────────
LUBE_CATEGORIES = ["DEO", "PCMO", "MCO", "LOW GRADE", "INDUSTRIAL GRADE", "Greases", "OTHERS"]

# ── Known international cities to flag ────────────────────────────────────────
INTERNATIONAL_CITIES = {
    "DUBAI", "SHARJAH", "ABU DHABI", "UAE", "SHARJAH UAE",
    "PO BOX 8043 SHARJAH UAE", "UNITED ARAB EMIRATES",
    "PARIS", "FRANCE", "PARIS, FRANCE",
    "SINGAPORE", "GUANGZHOU", "BEIJING", "BEIGING", "HANGZHOU",
    "SHENZHEN", "GUANGZHOU CITY", "CHINA", "HONG KONG",
    "RIYADH", "RIYAD", "KUWAIT", "QATAR", "MUSCAT",
    "KUALA LUMPUR", "DUSHANBE", "BAKU",
    "MISSISSAUGA", "MELBOURNE", "ISTANBUL",
    "MIAMI", "UNITED KINGDOM",
}

# ── City normalization table ───────────────────────────────────────────────────
# Maps variant spellings → canonical name
CITY_NORM = {
    # Islamabad / Rawalpindi cluster
    "ISLAMABAD":          "Islamabad",
    "ISALAMABAD":         "Islamabad",
    "ISLMABAD":           "Islamabad",
    "ISLAMABD":           "Islamabad",
    "Islamabad":          "Islamabad",
    "Rawalpindi":         "Rawalpindi",
    "RAWALPINDI":         "Rawalpindi",
    "RAWALPIND":          "Rawalpindi",
    "Taxila":             "Taxila",
    "Wah":                "Wah",
    "WAH CANTONMENT":     "Wah",
    "WAH CANTT":          "Wah",
    "WAH CANT":           "Wah",
    "Sihala":             "Sihala",

    # Lahore cluster
    "Lahore":             "Lahore",
    "LAHORE":             "Lahore",
    "LAHORE CANTT":       "Lahore",
    "DISST, LAHORE PUNJAB": "Lahore",
    "Kotlakhpat":         "Lahore",
    "MURIDKE":            "Muridke",
    "MURIDKEY":           "Muridke",
    "SHEIKHUPURA":        "Sheikhupura",
    "SHEKHUPURA":         "Sheikhupura",
    "SHEKHOPURA":         "Sheikhupura",
    "SHEIKHPURA":         "Sheikhupura",
    "SHEIKUPURA":         "Sheikhupura",
    "NANKANA":            "Nankana Sahib",
    "NANKANA SAHIB":      "Nankana Sahib",
    "NANKANASAHAB":       "Nankana Sahib",
    "NANKANASAHIB":       "Nankana Sahib",
    "KASUR":              "Kasur",
    "KASURE":             "Kasur",
    "KHUDDIAN KASUR":     "Kasur",

    # Faisalabad cluster
    "Faisalabad":         "Faisalabad",
    "FAISALABAD":         "Faisalabad",
    "FAISLABAD":          "Faisalabad",
    "FISALABAD":          "Faisalabad",
    "FAISLABAD":          "Faisalabad",
    "Machike":            "Machike",
    "MACHIKEY":           "Machike",
    "TOBA TEK SING":      "Toba Tek Singh",
    "TOBA TAKE SINGH":    "Toba Tek Singh",
    "Toba tek Singh":     "Toba Tek Singh",
    "TT SINGH":           "Toba Tek Singh",
    "Chiniot":            "Chiniot",
    "GOJRA":              "Gojra",
    "JHANG":              "Jhang",

    # Multan cluster
    "Multan":             "Multan",
    "MULTAN":             "Multan",
    "KHANEWAL":           "Khanewal",
    "MUZAFFARGARH":       "Muzaffargarh",
    "MUZAFARGARH":        "Muzaffargarh",
    "MUZAFAR GARH":       "Muzaffargarh",
    "Muzafar Garh":       "Muzaffargarh",
    "Muzaffar Garh":      "Muzaffargarh",
    "DISTT MUZAFFAR GARH": "Muzaffargarh",
    "DISTT  MUZAFFAR GARH": "Muzaffargarh",
    "MUZAFARGARH":        "Muzaffargarh",
    "VEHARI":             "Vehari",
    "Vehari":             "Vehari",
    "LODHRAN":            "Lodhran",
    "LODRAN":             "Lodhran",
    "SHUJABAD":           "Shujabad",
    "SHUJA ABAD":         "Shujabad",

    # Karachi cluster
    "Karachi":            "Karachi",
    "KARACHI":            "Karachi",
    "KARCHI":             "Karachi",
    "KATACHI":            "Karachi",
    "KARARO":             "Karachi",
    "KORANGI":            "Karachi",
    "BIN QASIM":          "Karachi",

    # Hyderabad cluster
    "Hyderabad":          "Hyderabad",
    "HYDERABAD":          "Hyderabad",
    "New Hyderabad":      "Hyderabad",
    "Kotri":              "Kotri",
    "JAMSHORO":           "Jamshoro",

    # Sukkur cluster
    "Sukkur":             "Sukkur",
    "SUKKUR":             "Sukkur",
    "SUKKAR":             "Sukkur",
    "DISTT  SUKKUR":      "Sukkur",

    # Peshawar cluster
    "Peshawar":           "Peshawar",
    "PESHAWAR":           "Peshawar",
    "PESHWAR":            "Peshawar",
    "?PESHAWAR":          "Peshawar",
    "PESHAWAR HAYATABAD": "Peshawar",
    "HAYATABAD":          "Peshawar",

    # Quetta cluster
    "Quetta":             "Quetta",
    "QUETTA":             "Quetta",

    # Gujranwala cluster
    "Gujranwala":         "Gujranwala",
    "GUJRANWALA":         "Gujranwala",
    "GUJJRANWALA":        "Gujranwala",
    "GUJRAN":             "Gujranwala",
    "WAZIRABAD":          "Wazirabad",
    "Sialkot":            "Sialkot",
    "SIALKOT":            "Sialkot",
    "SILAKOT":            "Sialkot",
    "Gujrat":             "Gujrat",
    "GUJRAT":             "Gujrat",
    "QASBA GUJRAT":       "Gujrat",

    # Sargodha cluster
    "Sargodha":           "Sargodha",
    "SARGODHA":           "Sargodha",
    "MIANWALI":           "Mianwali",
    "BHAKKAR":            "Bhakkar",
    "KHUSHAB":            "Khushab",

    # Bahawalpur cluster
    "BAHAWALPUR":         "Bahawalpur",
    "BHAWALPUR":          "Bahawalpur",
    "BAHAHAWALPUR":       "Bahawalpur",
    "BHAWALNAGAR":        "Bahawalnagar",
    "BAHAWALNAGAR":       "Bahawalnagar",
    "RAHIM YAR KHAN":     "Rahim Yar Khan",
    "RAHIMYARKHAN II":    "Rahim Yar Khan",
    "RAHIMYAR KHAN":      "Rahim Yar Khan",
    "RAHIMYARKHAN":       "Rahim Yar Khan",
    "R.Y.KHAN":           "Rahim Yar Khan",
    "DIST RAHIMYAR KHAN": "Rahim Yar Khan",

    # Sahiwal cluster
    "Sahiwal":            "Sahiwal",
    "SAHIWAL":            "Sahiwal",
    "OKARA":              "Okara",
    "PAKPATTAN":          "Pakpattan",
    "PAK PATTAN":         "Pakpattan",

    # DG Khan cluster
    "D.G.KHAN":           "DG Khan",
    "DERA GHAZI KHAN":    "DG Khan",
    "DERA GHAZI":         "DG Khan",
    "D-G KHAN":           "DG Khan",
    "DG KHAN":            "DG Khan",
    "RAJANPUR":           "Rajanpur",
    "RAJAN PUR":          "Rajanpur",

    # Mardan cluster
    "MARDAN":             "Mardan",
    "Mardan":             "Mardan",
    "SWABI":              "Swabi",
    "CHARSADDA":          "Charsadda",
    "NOWSHERA":           "Nowshera",
    "NOWSHEHRA":          "Nowshera",
    "NOSHERA":            "Nowshera",
    "NOSHERA":            "Nowshera",
    "RISALPUR":           "Nowshera",

    # Abbottabad / Hazara cluster
    "ABBOTABAD":          "Abbottabad",
    "ABBOTTABAD":         "Abbottabad",
    "ABBOTABAD / GB":     "Abbottabad",
    "Mansehra":           "Mansehra",
    "HARIPUR":            "Haripur",
    "HAVELIAN":           "Havelian",
    "Havelyan":           "Havelian",
    "BALAKOT":            "Balakot",

    # Kohat / DI Khan
    "Kohat":              "Kohat",
    "KOHAT":              "Kohat",
    "D.I.KHAN":           "DI Khan",
    "DERA ISMAIL KHAN":   "DI Khan",
    "DI KHAN":            "DI Khan",
    "DIKKHAN":            "DI Khan",
    "D.I. KHAN":          "DI Khan",

    # Jhelum / Chakwal
    "JHELUM":             "Jhelum",
    "JEHLUM":             "Jhelum",
    "JHELUM / AJK":       "Jhelum",
    "CHAKWAL":            "Chakwal",
    "TALAGANG":           "Talagang",

    # Gilgit / Northern Areas
    "Gilgit":             "Gilgit",
    "GILGIT":             "Gilgit",
    "GILGIT NORTHERN-AREA": "Gilgit",
    "Skardu":             "Skardu",
    "BALTISTAN":          "Skardu",
    "CHITRAL":            "Chitral",
    "Chitral":            "Chitral",

    # AJK
    "AZAD KASHMIR":       "Azad Kashmir",
    "AJK":                "Azad Kashmir",
    "Muzaffarabad":       "Muzaffarabad",
    "MIRPUR":             "Mirpur AJK",
    "MIRPUR A.K":         "Mirpur AJK",

    # Larkana cluster
    "LARKANA":            "Larkana",
    "JACOBABAD":          "Jacobabad",
    "SHIKARPUR":          "Shikarpur",
    "Shikarpur":          "Shikarpur",

    # Nawabshah cluster
    "Nawabshah":          "Nawabshah",
    "NAWABSHAH":          "Nawabshah",
    "SANGHAR":            "Sanghar",
    "MIRPURKHAS":         "Mirpur Khas",
    "Mirpur Khas":        "Mirpur Khas",
    "THARPARKAR":         "Tharparkar",
    "THATTA":             "Thatta",
    "DADU":               "Dadu",
    "KHAIRPUR":           "Khairpur",
    "Khairpur":           "Khairpur",
    "NASEERABAD":         "Naseerabad",
    "NASIRABAD":          "Naseerabad",
}

# ── Excel report colours ───────────────────────────────────────────────────────
COLOUR_GREEN_FILL  = "C6EFCE"
COLOUR_RED_FILL    = "FFC7CE"
COLOUR_YELLOW_FILL = "FFEB9C"
COLOUR_HEADER_BG   = "1F3864"   # dark navy
COLOUR_HEADER_FG   = "FFFFFF"
COLOUR_SUBHDR_BG   = "2E75B6"
COLOUR_SUBHDR_FG   = "FFFFFF"
COLOUR_ALT_ROW     = "EBF0F8"

GROWTH_THRESHOLD_PCT = 5.0   # ±5% for yellow band

# ── Opportunity schema (future) ────────────────────────────────────────────────
OPPORTUNITY_SCHEMA = {
    "CityNorm":               str,
    "Population":             float,
    "Households":             float,
    "HouseholdsWithVehicles": float,   # optional
    "RegisteredVehicles":     float,   # optional
}

# ── AI model ──────────────────────────────────────────────────────────────────
AI_MODEL      = "claude-sonnet-4-6"
AI_MAX_TOKENS = 8192     # raised — Gemini 2.5 thinking tokens eat into budget
