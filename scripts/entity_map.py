"""
Canonical entity registry + file-prefix mapping.

CompanyName / LegalName / Industry come from the user-supplied Dim_Entity
reference table (their "EntityId" column is what this module calls
FilePrefix — the internal EntityId keys below are unchanged from the
original handoff-doc scheme to avoid a wide refactor of every formula
that already keys off them).
"""

# Internal EntityId -> (CompanyName, LegalName, Industry)
# CompanyName/LegalName/Industry per the user's Dim_Entity table (2025).
ENTITIES = {
    "9357-7427 QI": ("9357-7427 Québec Inc. - Sandhu H.", "9357-7427 Québec Inc.", "Construction"),
    "9366-1676 QI": ("9366-1676 Québec Inc. - Sandhu G.", "9366-1676 Québec Inc.", "Real estate"),
    "9425-4547 QI": ("9425-4547 Québec Inc. - Sandhu G", "9425-4547 Québec Inc.", "Real estate"),
    "9455-4169 QI": ("9455-4169 Québec Inc. - Sandhu", "9455-4169 Québec Inc", "Investments"),
    "9455-4177 QI": ("9455-4177 Québec Inc. - Sandhu H.", "9455-4177 Québec Inc.", "Investments"),
    "9455-4201 QI": ("9455-4201 Québec Inc. - Sandhu G", "9455-4201 Québec inc.", "Investments"),
    "9108-6876 QI": ("Auberge du Carre St-Louis", "9108-6876 Québec Inc.", "Hotel"),
    "N/A_BisGur": ("Bistro Guru Inc.", "Bistro Guru Inc.", "Restaurant"),
    "9096-0436 QI": ("Lobby Bar Lounge", "9096-0436 Québec Inc.", "Restaurant"),
    "9504-7510 QI": ("Restaurant Héritaj", "9504-7510 Québec Inc.", "Restaurant"),
    "9455-8236 QI": ("Restaurant Indiarosa II", "9455-8236 Québec Inc.", "Restaurant"),
    "9525-2854 QI": ("Restaurant Indiarosa III", "9525-2854 Québec Inc.", "Restaurant"),
    "9366-1049 QI": ("Restaurant Indiarosa", "9366-1049 Québec Inc.", "Restaurant"),
    "9098-0558 QI": ("Restaurant Pizzeria Sandhu", "9098-0558 Québec Inc.", "Restaurant"),
    "2749-8567 QI": ("Sandhu Leasing", "2749-8567 Québec Inc.", "Rental"),
    # Not present in the user's Dim_Entity table — Industry inferred from
    # its sibling entity (Sandhu Leasing, also a real-estate rental
    # business) so it groups sensibly in the workbook; flagged in the
    # Notes & Validation sheet, not silently presented as given data.
    "N/A_SanSan": ("Sandhu & Sandhu Enr.", "Sandhu & Sandhu Enr.", "Rental"),
}

# FilePrefix (shared by the WTB and Tax export filenames, and matching the
# "EntityId" column of the user's Dim_Entity table) -> internal EntityId.
# Positively identified via each Tax CSV's IDENT.Ident311 field.
# "Rres_Ind" in the user's table is "Res_Ind" here (same LegalName,
# 9366-1049 Québec Inc. — a typo in their EntityId column, not a new entity).
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
