"""
Canonical entity registry + file-prefix mapping, built from Section 4 of
the handoff doc plus positive identification of each new WTB/Tax file
pair via the Taxprep CSV's IDENT.Ident311 company-name field (WTB files
carry no entity identity internally, so this is the reliable source).
"""

# EntityId -> (EntityName, Segment)
ENTITIES = {
    "9357-7427 QI": ("9357-7427 Québec Inc.", "HoldCo"),
    "9366-1676 QI": ("9366-1676 Québec Inc.", "HoldCo"),
    "9425-4547 QI": ("9425-4547 Québec Inc.", "HoldCo"),
    "9455-4169 QI": ("9455-4169 Québec Inc.", "HoldCo"),
    "9455-4177 QI": ("9455-4177 Québec Inc.", "HoldCo"),
    "9455-4201 QI": ("Sandhu Holdco", "HoldCo"),
    "9108-6876 QI": ("Auberge St Louis", "Hospitality OpCo"),
    "N/A_BisGur": ("Bistro Guru Inc.", "Restaurant OpCo"),
    "9096-0436 QI": ("Lobby Lounge", "Restaurant OpCo"),
    "9504-7510 QI": ("Restaurant Heritage", "Restaurant OpCo"),
    "9455-8236 QI": ("Restaurant IndiaRosa 2", "Restaurant OpCo"),
    "9525-2854 QI": ("Restaurant IndiaRosa 3", "Restaurant OpCo"),
    "9366-1049 QI": ("Restaurant IndiaRosa 1", "Restaurant OpCo"),
    "9098-0558 QI": ("Restaurant Sandhu", "Restaurant OpCo"),
    "2749-8567 QI": ("Sandhu Leasing", "RealCo"),
    "N/A_SanSan": ("Sandhu & Sandhu Enr.", "RealCo"),
}

# FilePrefix (shared by the WTB and Tax export filenames) -> EntityId.
# Positively identified via each Tax CSV's IDENT.Ident311 field, not
# guessed from the filename abbreviation alone.
FILE_PREFIX_TO_ENTITY = {
    "935_742": "9357-7427 QI",
    "936_167": "9366-1676 QI",
    "942_454": "9425-4547 QI",
    "945_416": "9455-4169 QI",
    "945_417": "9455-4177 QI",
    "945_420": "9455-4201 QI",
    "Aub_Stl": "9108-6876 QI",
    "Bis_Gur": "N/A_BisGur",
    "Lob_Lou": "9096-0436 QI",
    "Res_Her": "9504-7510 QI",
    "Res_In2": "9455-8236 QI",
    "Res_In3": "9525-2854 QI",
    "Res_Ind": "9366-1049 QI",
    "Res_Sa2": "9098-0558 QI",
    "San_Lea": "2749-8567 QI",
    "San_San": "N/A_SanSan",
}

# Known, not in the handoff's 18-entity list at all — flagged, not guessed into place.
UNEXPECTED_ENTITIES = {
    "947_338": "9475-3381 Québec Inc.",
}
