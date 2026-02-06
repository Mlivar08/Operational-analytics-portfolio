Option Explicit

Sub Export_Replenishment_TXT_ByStore()

    Dim categories As Variant
    Dim storeCols As Variant
    Dim storeFolders As Variant

    Dim i As Long, colIndex As Long
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range

    Dim fieldNum As Long
    Dim outputFile As String
    Dim hasData As Boolean

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' High-level product categories (anonymized / generalized)
    categories = Array("CAT_A", "CAT_B", "CAT_C", "CAT_D", "CAT_E", "CAT_F", "CAT_G", "CAT_H", "CAT_I", "CAT_J")

    ' Columns representing store allocation quantities (example columns)
    storeCols = Array("AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", _
                      "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP")

    ' Store folder identifiers (generalized)
    storeFolders = Array("S01", "S02", "S03", "S04", "S05", "S06", "S07", "S08", "S09", "S10", _
                         "S11", "S12", "S13", "S14", "S15", "S16", "S17", "S18", "S19", "S20", _
                         "S21", "S22", "S23", "S24", "S25", "S26", "S27", "S28", "S29")

    ' Worksheet and table names anonymized
    Set ws = ThisWorkbook.Worksheets("Orders")
    Set tbl = ws.ListObjects("OrdersTable")
    Set rng = tbl.Range

    ' Base output path (anonymized)
    Dim basePath As String
    basePath = "C:\Exports\ReplenishmentFiles\"

    ' Iterate through each store column and each category
    For colIndex = LBound(storeCols) To UBound(storeCols)

        ' Convert column letter to field number within the table
        fieldNum = ws.Range(storeCols(colIndex) & "1").Column - rng.Column + 1

        ' Skip if calculated field is outside the table range
        If fieldNum < 1 Or fieldNum > rng.Columns.Count Then
            GoTo NextStore
        End If

        For i = LBound(categories) To UBound(categories)

            ' Clear filters if active
            On Error Resume Next
            If ws.AutoFilterMode Then
                ws.AutoFilter.ShowAllData
            End If
            On Error GoTo 0

            ' Apply filters:
            ' 1) Store allocation column must be non-empty
            ' 2) Category is matched in Field 1 (adjust if your category column differs)
            rng.AutoFilter Field:=fieldNum, Criteria1:="<>"
            rng.AutoFilter Field:=1, Criteria1:=categories(i)

            ' Check if any visible rows remain (excluding header)
            hasData = False
            On Error Resume Next
            If Application.WorksheetFunction.Subtotal(103, rng.Columns(1)) > 1 Then
                hasData = True
            End If
            On Error GoTo 0

            If hasData Then

                ' Copy the store allocation column only
                ws.Columns(storeCols(colIndex) & ":" & storeCols(colIndex)).Copy

                ' Create a new workbook with one sheet
                Workbooks.Add
                With ActiveSheet

                    ' Paste values only
                    .Range("A1").PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False

                    ' Remove header row
                    .Rows(1).Delete

                    ' Export only if there is any content
                    If Application.WorksheetFunction.CountA(.UsedRange) > 0 Then

                        ' Ensure destination folder exists
                        Dim folderPath As String
                        folderPath = basePath & storeFolders(colIndex) & "\"
                        If Dir(folderPath, vbDirectory) = vbNullString Then
                            MkDir folderPath
                        End If

                        ' Build output file name
                        outputFile = folderPath & storeFolders(colIndex) & "_" & categories(i) & ".txt"

                        ' Save as TXT
                        ActiveWorkbook.SaveAs Filename:=outputFile, FileFormat:=xlTextPrinter, CreateBackup:=False
                    End If

                    ' Close the temporary workbook without prompts
                    ActiveWorkbook.Close SaveChanges:=False
                End With
            End If

        Next i

        ' Clear filters after finishing store
        On Error Resume Next
        ws.ShowAllData
        On Error GoTo 0

NextStore:
    Next colIndex

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Export completed.", vbInformation

End Sub
