Attribute VB_Name = "Main"
Option Explicit

Public Tables As Collection

Public Sub Start()
    Application.ScreenUpdating = False
    Globals.InitializeLetters
    ReadTablesInfo
    DoWork
    Application.ScreenUpdating = True
    MsgBox "Finished"
End Sub

Private Sub Initialize()
    
    Dim i As Integer
    
    'Capital letters
    For i = 1 To 26
        Globals.Letters(i) = Chr(i + 64)
    Next i
    
    'Lower case letters
    For i = 27 To 52
        GlobalsLetters(i) = Chr(i + 70)
    Next i

End Sub

Private Sub ReadTablesInfo()

    Set Tables = New Collection
    
    Dim i As Integer
    For i = 2 To Globals.TableCount + 1
        Dim t As MH_Tables.Table
        Set t = New MH_Tables.Table
        With ThisWorkbook.Worksheets("Index")
            t.TableName = .Cells(i, 1)
            t.WorksheetName = .Cells(i, 2)
            t.HeaderRows = .Cells(i, 3)
            t.TitleLevels = .Cells(i, 4)
            t.BannerLevels = .Cells(i, 5)
            t.BaseIn = .Cells(i, 6)
            t.SigTestVs = .Cells(i, 7)
        End With
        Main.Tables.Add t
    Next i

End Sub

Private Sub DoWork()
    
    Dim t As MH_Tables.Table
    For Each t In Main.Tables
        Set t.Worksheet = ActiveWorkbook.Worksheets(t.WorksheetName)
        t.SetDataRanges
        t.InsertNewRows
        t.SetupSubTables
        t.PreFormatting
        t.WriteLegend
        t.HighlightBases
        t.HighlightSignificances
        t.PostFormatting
    Next t

End Sub

