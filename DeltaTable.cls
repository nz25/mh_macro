VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeltaTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITable

Public ParentTable As MH_Tables.Table

Public ColumnOffset As Integer

Public Worksheet As Worksheet

Public DataRange As Range
Public ReferenceRange As Range
Public BaseRange As Range
Public TestRange As Range
Public TitleRange As Range
Public BannerRange As Range

Public SubTables As Collection

Public Function ITable_GetSigTestVs() As String
    ITable_GetSigTestVs = "Delta"
End Function

Public Function ITable_GetBannerRange() As Range
    Set ITable_GetBannerRange = Me.BannerRange
End Function

Public Sub Initialize(parent As MH_Tables.Table)

    Set Me.ParentTable = parent
    Me.ColumnOffset = Me.ParentTable.DataRange.Columns.Count + 2

    CopyContentFromParent
    
    'Inserts column for average value
    Me.ParentTable.DataRange.Offset(0, Me.ColumnOffset).Cells(1, 1).EntireColumn.Insert
    
    Set Me.DataRange = Me.ParentTable.DataRange.Offset(0, Me.ColumnOffset).Resize(, Me.ParentTable.DataRange.Columns.Count + 1)
    Set Me.ReferenceRange = Me.DataRange.Columns(1)
    Set Me.TestRange = Me.DataRange.Offset(0, 1).Resize(, Me.DataRange.Columns.Count - 1)
    Set Me.TitleRange = Me.DataRange.Offset(0, -1).Columns(1)
    Set Me.BannerRange = Me.DataRange.Offset(-2, 0).Resize(2)
    Me.BannerRange.Cells(2, 1) = "Average all brands"
    
    'Fills in average column
    Dim i As Integer
    For i = 1 To Me.DataRange.Rows.Count
        If Me.TitleRange.Rows(i) <> "Base" Then
            Me.ReferenceRange.Rows(i) = WorksheetFunction.Average(Me.ParentTable.DataRange.Rows(i))
        End If
    Next i
        
End Sub

Private Sub CopyContentFromParent()
    
    Dim sourceRange As Range
    Set sourceRange = Me.ParentTable.DataRange.Offset(-3, -1).Resize(Me.ParentTable.DataRange.Rows.Count + 3, Me.ParentTable.DataRange.Columns.Count + 1)
    sourceRange.Copy sourceRange.Offset(, Me.ColumnOffset).Cells(1, 1)
    
End Sub

Public Sub SetupSubTables()

    Set Me.SubTables = New Collection

    Dim st As SubTable, pst As SubTable
    
    For Each pst In Me.ParentTable.SubTables
        Set st = New SubTable
        st.Initialize Me, pst.DataRange.Offset(0, Me.ColumnOffset).Resize(, Me.DataRange.Columns.Count)
        Me.SubTables.Add st
    Next pst

End Sub

Public Sub PreFormatting()

    Me.TitleRange.Columns.AutoFit
    Me.BannerRange.Font.Size = 8
    Me.BannerRange.Font.Bold = False
    Me.BannerRange.Rows.AutoFit
    Me.BannerRange.HorizontalAlignment = xlCenter
    Me.BannerRange.WrapText = True
    Me.DataRange.HorizontalAlignment = xlCenter
    
    Dim st As SubTable
    For Each st In Me.SubTables
        st.TestRange.NumberFormat = "+0;-0;0"
    Next st
    
End Sub

Public Sub HighlightBases()

    Dim st As SubTable
    For Each st In Me.SubTables
        Me.ParentTable.HighlightBaseRange st.BaseRange
    Next st
    
End Sub

Public Sub HighlightSignificances()
    
    Dim st As SubTable
    For Each st In Me.SubTables
        st.HighlightSignificances
    Next st

End Sub

Public Sub CalculateDeltas()

    Dim i As Integer, j As Integer
    For i = 1 To Me.TestRange.Rows.Count
        If Me.TitleRange.Rows(i) <> "Base" And Me.TitleRange.Rows(i) <> "" Then
            For j = 1 To Me.TestRange.Columns.Count
                Me.TestRange.Cells(i, j) = Me.TestRange.Cells(i, j) - Me.ReferenceRange.Rows(i)
            Next j
        End If
    Next i
    
End Sub
