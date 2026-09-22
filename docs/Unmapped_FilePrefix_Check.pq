let
    // ====== Only line that matters here ======
    FilePath = "C:\Sandhu\Sandhu_G__2027.xlsx",
    // ==========================================

    FilePrefixMap = Excel.CurrentWorkbook(){[Name="PQ_Dim_FilePrefix"]}[Content],
    KnownPrefixesUpper = List.Transform(FilePrefixMap[FilePrefix], each Text.Upper(_)),

    Workbook = Excel.Workbook(File.Contents(FilePath), null, true),
    Sheets = Table.SelectColumns(Workbook, {"Name"}),

    // Every tab in the consolidated file whose name doesn't case-insensitively
    // match a row in PQ_Dim_FilePrefix — its data was silently excluded from
    // Fact_TrialBalance_Refresh rather than guessed into the wrong entity.
    // This table must be empty before you trust a refresh. This is exactly
    // what would have caught the FY2026 example's un-renamed "Sheet16" tab
    // (it turned out to be Sandhu & Sandhu Enr., found only by manually
    // inspecting its account names — never assume, always check this table).
    // If a name here is a genuinely new or renamed entity, add a row to
    // PQ_Dim_FilePrefix (and Dim_Entity, for its CompanyName/Industry) and
    // refresh again. If it's a stray non-entity tab (a cover sheet, notes),
    // it's fine to ignore — but confirm what it is before assuming that.
    Unmapped = Table.SelectRows(Sheets, each not List.Contains(KnownPrefixesUpper, Text.Upper([Name])))
in
    Unmapped
