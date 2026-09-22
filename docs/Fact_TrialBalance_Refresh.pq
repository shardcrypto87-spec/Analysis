let
    // ====== EDIT THESE TWO LINES BEFORE EACH YEARLY REFRESH ======
    FolderPath = "C:\Sandhu\WTB_2026\",
    CurrentFY  = "FY2026",
    // ==============================================================

    // Derive the two comparison-year labels from CurrentFY (e.g. FY2026 -> FY2025, FY2024).
    // The WTB's own "Final: <date>" column header is NOT trusted for this —
    // it can be stale by a year — so this is set explicitly instead.
    CurrentYearNum = Number.FromText(Text.End(CurrentFY, 4)),
    PriorFY  = "FY" & Text.From(CurrentYearNum - 1),
    Prior2FY = "FY" & Text.From(CurrentYearNum - 2),

    // Reference tables already embedded in this workbook.
    GifiMap       = Excel.CurrentWorkbook(){[Name="PQ_GIFI_Mapping"]}[Content],
    FallbackMap   = Excel.CurrentWorkbook(){[Name="PQ_Fallback_MapNo"]}[Content],
    ExceptionsTbl = Excel.CurrentWorkbook(){[Name="PQ_Exceptions"]}[Content],
    FilePrefixMap = Excel.CurrentWorkbook(){[Name="PQ_Dim_FilePrefix"]}[Content],
    DimEntity     = Excel.CurrentWorkbook(){[Name="Dim_Entity"]}[Content],

    // ---- 1. Read every WTB file in the folder ----
    SourceFiles = Folder.Files(FolderPath),
    WTBFilesOnly = Table.SelectRows(SourceFiles, each Text.EndsWith([Name], "_WTB.xlsx")),

    // FilePrefix = filename with the trailing "_<year>_WTB" stripped, e.g.
    // "Res_In2_2026_WTB.xlsx" -> "Res_In2". Works regardless of the year value.
    AddPrefix = Table.AddColumn(WTBFilesOnly, "FilePrefix", each
        let
            baseName = Text.BeforeDelimiter([Name], ".xlsx"),
            parts = Text.Split(baseName, "_"),
            prefixParts = List.RemoveLastN(parts, 2)
        in
            Text.Combine(prefixParts, "_")
    ),

    // Read each file's first (only) worksheet, raw — no header promotion,
    // because column POSITION is trusted, not the header text (which
    // includes a date that can be stale). Row 1 is skipped manually below.
    AddSheet = Table.AddColumn(AddPrefix, "SheetData", each
        let
            wb = Excel.Workbook([Content], null, true),
            firstSheet = wb{0}[Data]
        in
            Table.Skip(firstSheet, 1)
    ),

    Combined = Table.SelectColumns(AddSheet, {"FilePrefix", "SheetData"}),
    Expanded = Table.ExpandTableColumn(Combined, "SheetData",
        {"Column1", "Column2", "Column4", "Column5", "Column13", "Column14", "Column15", "Column16"},
        {"AccountNo", "Name", "MapNo", "Type", "Final", "GIFI", "Prior", "Prior2"}
    ),

    // ---- 2. Income statement rows only ----
    IncomeOnly = Table.SelectRows(Expanded, each [Type] = "Income statement"),

    // ---- 3. Unpivot the 3 year columns into (FiscalYear, Amount) rows ----
    UnpivotFinal = Table.AddColumn(
        Table.SelectColumns(IncomeOnly, {"FilePrefix", "AccountNo", "Name", "MapNo", "GIFI", "Final"}),
        "FiscalYear", each CurrentFY
    ),
    RenameFinal = Table.RenameColumns(UnpivotFinal, {{"Final", "Amount"}}),

    UnpivotPrior = Table.AddColumn(
        Table.SelectColumns(IncomeOnly, {"FilePrefix", "AccountNo", "Name", "MapNo", "GIFI", "Prior"}),
        "FiscalYear", each PriorFY
    ),
    RenamePrior = Table.RenameColumns(UnpivotPrior, {{"Prior", "Amount"}}),

    UnpivotPrior2 = Table.AddColumn(
        Table.SelectColumns(IncomeOnly, {"FilePrefix", "AccountNo", "Name", "MapNo", "GIFI", "Prior2"}),
        "FiscalYear", each Prior2FY
    ),
    RenamePrior2 = Table.RenameColumns(UnpivotPrior2, {{"Prior2", "Amount"}}),

    AllYears = Table.Combine({RenameFinal, RenamePrior, RenamePrior2}),
    NonZero = Table.SelectRows(AllYears, each [Amount] <> null and [Amount] <> 0),

    // ---- 4. Categorize: GIFI first, then Map No fallback ----
    WithGifiText = Table.AddColumn(NonZero, "GIFIText", each
        if [GIFI] = null then "" else Text.Trim(Text.From([GIFI]))),

    MergedGifi = Table.NestedJoin(WithGifiText, {"GIFIText"}, GifiMap, {"GIFICode"}, "GifiJoin", JoinKind.LeftOuter),
    ExpandGifi = Table.ExpandTableColumn(MergedGifi, "GifiJoin", {"Category"}, {"CategoryFromGifi"}),

    WithMapNoNorm = Table.AddColumn(ExpandGifi, "MapNoNorm", each
        if [MapNo] = null then "" else Text.Combine(Text.Split(Text.From([MapNo]), " "), "")),

    MergedFallback = Table.NestedJoin(WithMapNoNorm, {"FilePrefix", "MapNoNorm"}, FallbackMap,
        {"FilePrefix", "MapNoNormalized"}, "FallbackJoin", JoinKind.LeftOuter),
    ExpandFallback = Table.ExpandTableColumn(MergedFallback, "FallbackJoin", {"Category"}, {"CategoryFromFallback"}),

    WithCategory = Table.AddColumn(ExpandFallback, "Category", each
        if [CategoryFromGifi] <> null then [CategoryFromGifi]
        else if [CategoryFromFallback] <> null then [CategoryFromFallback]
        else null
    ),

    // ---- 5. Account-level exceptions override whatever was found above ----
    MergedExceptions = Table.NestedJoin(WithCategory, {"FilePrefix", "AccountNo"}, ExceptionsTbl,
        {"FilePrefix", "AccountNo"}, "ExceptionJoin", JoinKind.LeftOuter),
    ExpandExceptions = Table.ExpandTableColumn(MergedExceptions, "ExceptionJoin", {"Category"}, {"CategoryOverride"}),

    FinalCategory = Table.AddColumn(ExpandExceptions, "FinalCategory", each
        if [CategoryOverride] <> null then [CategoryOverride] else [Category]
    ),

    // ---- 6. Map FilePrefix -> EntityId -> Industry ----
    // LeftOuter, NOT Inner: a file whose FilePrefix isn't in PQ_Dim_FilePrefix
    // (a new/renamed file, or an unexpected entity like "947_338" this year)
    // must not be silently dropped — it's excluded below and surfaced
    // instead in Unmapped_FilePrefix_Check, which must be empty before you
    // trust a refresh.
    MergedEntity = Table.NestedJoin(FinalCategory, {"FilePrefix"}, FilePrefixMap, {"FilePrefix"}, "EntityJoin", JoinKind.LeftOuter),
    ExpandEntity = Table.ExpandTableColumn(MergedEntity, "EntityJoin", {"EntityId"}, {"EntityId"}),

    MergedIndustry = Table.NestedJoin(ExpandEntity, {"EntityId"}, DimEntity, {"EntityId"}, "IndustryJoin", JoinKind.LeftOuter),
    ExpandIndustry = Table.ExpandTableColumn(MergedIndustry, "IndustryJoin", {"Industry"}, {"Industry"}),

    // ---- 7. Drop rows with no category or no entity mapping (both checked
    // separately before you trust a refresh — see Uncategorized_Check and
    // Unmapped_FilePrefix_Check) ----
    Categorized = Table.SelectRows(ExpandIndustry, each [FinalCategory] <> null and [EntityId] <> null),

    // ---- 8. Aggregate to the Fact_TrialBalance shape ----
    Grouped = Table.Group(Categorized, {"EntityId", "FinalCategory", "FiscalYear", "Industry"},
        {{"Amount", each List.Sum([Amount]), type number}}),

    Renamed = Table.RenameColumns(Grouped, {{"FinalCategory", "Category"}}),
    Reordered = Table.SelectColumns(Renamed, {"EntityId", "Category", "FiscalYear", "Amount", "Industry"})
in
    Reordered
