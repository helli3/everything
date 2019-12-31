Sub GetSheets()
Dim WriteRow As Long, _
    LastCell As Range, _
    WbDest As Workbook, _
    WbSrc As Workbook, _
    WsDest As Worksheet, _
    WsSrc As Worksheet

Set WbDest = ThisWorkbook
Set WsDest = WbDest.Sheets.Add
WsDest.Cells(1, 1) = "Na naglowek"

Path = "I:\Inwentaryzacja 2019 - skladniki ruchome\Inwentaryzacja\Załaczniki\"
Filename = Dir(Path & "*.xls")

Do While Filename <> ""
    Set WbSrc = Workbooks.Open(Filename:=Path & Filename, ReadOnly:=True)
    Set WsSrc = WbSrc.Sheets(1)
    With WsSrc
        Set LastCell = .Cells.Find(What:="*", _
                      After:=.Range("A1"), _
                      Lookat:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByRows, _
                      SearchDirection:=xlPrevious, _
                      MatchCase:=False)
        .Range(.Range("A1"), LastCell).Copy
    End With
    With WsDest
        WriteRow = .Cells.Find(What:="*", _
                      After:=.Range("A1"), _
                      Lookat:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByRows, _
                      SearchDirection:=xlPrevious, _
                      MatchCase:=False).Row + 1
        .Range("A" & WriteRow).PasteSpecial
    End With
    
    Application.CutCopyMode = False

    WbSrc.Close
    Filename = Dir()
Loop

End Sub



