Attribute VB_Name = "Main"
Option Explicit

Public Tables As Collection

Public Sub Start()

    If ThisWorkbook.name = ActiveWorkbook.name Then
        MsgBox ("Please select tables you want to modify")
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Initialize
    CheckTableNames
    ReadTablesInfo
    DoWork
    Application.ScreenUpdating = True
    MsgBox "Finished"
End Sub

Private Sub Initialize()
    
    '1. Populates array containing higher and lower case letters, which is used for significance testing
    Dim i As Integer
    
    For i = 1 To 26
        Globals.Letters(i) = Chr(i + 64)
    Next i
    For i = 27 To 52
        Globals.Letters(i) = Chr(i + 70)
    Next i

    Set Main.Tables = New Collection
    
    ThisWorkbook.Worksheets("Index").Range("B2:B" & Globals.TableCount + 1).ClearContents

End Sub

Private Sub CheckTableNames()

    Dim i As Integer, j As Integer, TableName As String, tableFound As Boolean
    
    For i = 1 To ActiveWorkbook.Worksheets.Count
        TableName = ActiveWorkbook.Worksheets(i).Range(Globals.TableNameCellAddress)
        tableFound = False
        For j = 2 To Globals.TableCount + 1
            If TableName = ThisWorkbook.Worksheets("Index").Cells(j, 1) Then
                tableFound = True
                ThisWorkbook.Worksheets("index").Cells(j, 2) = ActiveWorkbook.Worksheets(i).name
            End If
        Next j
        
        If Not tableFound Then MsgBox "Unknown Table: " + TableName
    Next i
    
End Sub

Private Sub ReadTablesInfo()

    Dim i As Integer, worksheetName As String
    
    For i = 2 To Globals.TableCount + 1
        worksheetName = ThisWorkbook.Worksheets("Index").Cells(i, 2)
        If worksheetName <> "" Then
            Dim t As MH_Tables.Table
            Set t = New MH_Tables.Table
            t.Initialize ThisWorkbook.Worksheets("Index").Rows(i), ActiveWorkbook.Worksheets(worksheetName)
            Main.Tables.Add t
        End If
    Next i

End Sub

Private Sub DoWork()
    
    Dim t As MH_Tables.Table
    For Each t In Main.Tables
        t.InsertNewRows
        t.SetupSubTables
        t.PreFormatting
        t.WriteLegend
        t.HighlightBases
        t.HighlightSignificances
        t.PostFormatting
    Next t

    'comes back to the 1st sheet
    ActiveWorkbook.Worksheets(1).Select
    
End Sub

