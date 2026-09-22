let
    // ====== MUST MATCH the same two lines in Fact_TrialBalance_Refresh.pq ======
    FolderPath = "C:\Sandhu\WTB_2026\",
    CurrentFY  = "FY2026",
    // ============================================================================

    CurrentYearNum = Number.FromText(Text.End(CurrentFY, 4)),
    PriorFY  = "FY" & Text.From(CurrentYearNum - 1),
    Prior2FY = "FY" & Text.From(CurrentYearNum - 2),

    GifiMap       = Excel.CurrentWorkbook(){[Name="PQ_GIFI_Mapping"]}[Content],
    FallbackMap   = Excel.CurrentWorkbook(){[Name="PQ_Fallback_MapNo"]}[Content],
    ExceptionsTbl = Excel.CurrentWorkbook(){[Name="PQ_Exceptions"]}[Content],

    SourceFiles = Folder.Files(FolderPath),
    WTBFilesOnly = Table.SelectRows(SourceFiles, each Text.EndsWith([Name], "_WTB.xlsx")),

    AddPrefix = Table.AddColumn(WTBFilesOnly, "FilePrefix", each
        let
            baseName = Text.BeforeDelimiter([Name], ".xlsx"),
            parts = Text.Split(baseName, "_"),
            prefixParts = List.RemoveLastN(parts, 2)
        in
            Text.Combine(prefixParts, "_")
    ),

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

    IncomeOnly = Table.SelectRows(Expanded, each [Type] = "Income statement"),

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

    MergedExceptions = Table.NestedJoin(WithCategory, {"FilePrefix", "AccountNo"}, ExceptionsTbl,
        {"FilePrefix", "AccountNo"}, "ExceptionJoin", JoinKind.LeftOuter),
    ExpandExceptions = Table.ExpandTableColumn(MergedExceptions, "ExceptionJoin", {"Category"}, {"CategoryOverride"}),

    FinalCategory = Table.AddColumn(ExpandExceptions, "FinalCategory", each
        if [CategoryOverride] <> null then [CategoryOverride] else [Category]
    ),

    // This is the whole point of this query: every row here is real money
    // (non-zero) that didn't match a GIFI code, the Map No fallback, OR an
    // exception. Before trusting a refresh, this table should be EMPTY.
    StillUncategorized = Table.SelectRows(FinalCategory, each [FinalCategory] = null),
    Result = Table.SelectColumns(StillUncategorized,
        {"FilePrefix", "AccountNo", "Name", "MapNo", "GIFI", "FiscalYear", "Amount"})
in
    Result
