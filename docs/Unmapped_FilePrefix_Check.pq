let
    // ====== Only line that matters here ======
    FolderPath = "C:\Sandhu\WTB_2026\",
    // ==========================================

    FilePrefixMap = Excel.CurrentWorkbook(){[Name="PQ_Dim_FilePrefix"]}[Content],

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

    Files = Table.SelectColumns(AddPrefix, {"Name", "FilePrefix"}),

    // Every file in the folder whose FilePrefix has no row in
    // PQ_Dim_FilePrefix — its data was silently excluded from
    // Fact_TrialBalance_Refresh rather than guessed into the wrong
    // entity. This table must be empty before you trust a refresh.
    // If a name here is a genuinely new or renamed entity, add a row to
    // PQ_Dim_FilePrefix (and Dim_Entity, for its CompanyName/Industry)
    // and refresh again.
    Unmapped = Table.SelectRows(Files, each not List.Contains(FilePrefixMap[FilePrefix], [FilePrefix]))
in
    Unmapped
