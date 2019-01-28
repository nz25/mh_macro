VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TableName As String
Public WorksheetName As String
Public HeaderRows As Integer
Public TitleLevels As Integer
Public BannerLevels As Integer
Public BaseIn As String
Public SigTestVs As String

Public Worksheet As Worksheet

Public DataRange As Range
Public ReferenceRange As Range
Public BaseRange As Range
Public TestRange As Range
Public HeaderRange As Range
Public LegendRange As Range
Public TitleRange As Range
Public BannerRange As Range

Public SubTables As Collection

Public Property Get FirstDataRow() As Integer
    FirstDataRow = Me.HeaderRows + Me.BannerLevels * 2 + 2
End Property

Public Property Get FirstDataColumn() As Integer
    FirstDataColumn = Me.TitleLevels * 2 + 1
End Property

Public Sub SetDataRanges()

    Dim lastRow As Integer, lastColumn As Integer
    
    lastRow = Me.Worksheet.UsedRange.Rows.Count
    lastColumn = Me.Worksheet.UsedRange.Columns.Count
    
    Set Me.DataRange = Me.Worksheet.Range(Me.Worksheet.Cells(Me.FirstDataRow, Me.FirstDataColumn), Me.Worksheet.Cells(lastRow, lastColumn))
    Set Me.TitleRange = Me.DataRange.Offset(0, -1).Columns(1)
    Set Me.BannerRange = Me.DataRange.Offset(-1, 0).Rows(1)
    Set Me.HeaderRange = Me.Worksheet.Range(Me.Worksheet.Cells(1, 1), Me.Worksheet.Cells(Me.HeaderRows, 1))
    Set Me.LegendRange = Me.HeaderRange.Offset(0, 2).Resize(4)
    
    If Me.BaseIn = "Row" Then
        Set Me.BaseRange = Me.DataRange.Rows(1)
        If Me.SigTestVs = "Ref" Then
            Set Me.ReferenceRange = Me.DataRange.Columns(1)
            Set Me.TestRange = Me.DataRange.Offset(1, 1).Resize(Me.DataRange.Rows.Count - 1, Me.DataRange.Columns.Count - 1)
        End If
    ElseIf Me.BaseIn = "Column" Then
        Set Me.BaseRange = Me.DataRange.Columns(1)
        If Me.SigTestVs = "Ref" Then
            Set Me.ReferenceRange = Me.DataRange.Rows(1)
            Set Me.TestRange = Me.DataRange.Offset(1, 1).Resize(Me.DataRange.Rows.Count - 1, Me.DataRange.Columns.Count - 1)
        End If
    ElseIf Me.BaseIn = "SubTable" Then
        'All ranges are defined in subtable class
    End If
End Sub

Public Sub InsertNewRows()
'Insert new rows in the tables for significance tests

    If SigTestVs = "All" Then
    
        Me.BannerRange.Offset(1, 0).EntireRow.Insert Shift:=xlUp
        Set Me.BannerRange = Me.BannerRange.Resize(2)
        Dim i As Integer
        For i = 1 To Me.BannerRange.Columns.Count
            Me.BannerRange.Cells(2, i) = Globals.Letters(i)
        Next i
            
        Dim r As Range
        Set r = Me.TitleRange.Rows(1)
        Do While r <> ""
            If r <> "Base" And Left(r, 10) <> "Conversion" Then
                r.Offset(1, 0).EntireRow.Insert Shift:=xlUp
                Set r = r.Offset(1, 0)
            End If
            Set r = r.Offset(1, 0)
        Loop
        Set Me.DataRange = Me.DataRange.Resize(Me.DataRange.Rows.Count + 1)
        Set Me.TitleRange = Me.TitleRange.Resize(Me.TitleRange.Rows.Count + 1)
        
    End If
End Sub

Public Sub SetupSubTables()

    If Me.BaseIn = "SubTable" Then
        
        Set Me.SubTables = New Collection
        
        Dim i As Integer, firstRow As Integer, lastRow As Integer
        For i = Me.TitleRange.Rows.Count To 1 Step -1
            If lastRow = 0 Then lastRow = i
            If Me.TitleRange.Rows(i) = "Base" Then
                firstRow = i
                Dim st As SubTable
                Set st = New SubTable
                st.Initialize Me, Me.DataRange.Offset(firstRow - 1).Resize(lastRow - firstRow + 1)
                Me.SubTables.Add st
                lastRow = 0
            End If
        Next
        
    End If

End Sub

Public Sub PreFormatting()
    
    ResetFormatting
    
    Me.Worksheet.name = Left(Me.HeaderRange.Rows(2), 31)
    Me.HeaderRange.Rows(Me.HeaderRows).Font.Bold = True
    
    With Me.Worksheet.UsedRange.Resize(Me.Worksheet.UsedRange.Rows.Count + 1)
        .Font.name = Globals.Font
        .numberFormat = "0"
        .VerticalAlignment = xlCenter
    End With
    
    Me.TitleRange.Columns.AutoFit
    Me.BannerRange.Font.Size = 8
    Me.BannerRange.Font.Bold = False
    Me.BannerRange.Rows.AutoFit
    Me.BannerRange.HorizontalAlignment = xlCenter
        
    Me.DataRange.HorizontalAlignment = xlCenter
    Me.DataRange.Replace What:="XXXXX", Replacement:="n.a."
    
End Sub

Private Sub ResetFormatting()
    With Me.Worksheet.UsedRange
        .Interior.Pattern = xlNone
        .Font.Color = XlRgbColor.rgbBlack
        .Font.Italic = False
    End With
End Sub

Public Sub WriteLegend()

    LegendRange.Font.Size = 10
    If Me.BaseIn <> "" Then
        LegendRange.Rows(1).FormulaR1C1 = "Low base n<" & Globals.BaseLow
        LegendRange.Rows(1).Interior.Color = Globals.ColorBaseLow
        LegendRange.Rows(2).FormulaR1C1 = "Low base n<" & Globals.BaseTooLow
        LegendRange.Rows(2).Interior.Color = Globals.ColorBaseTooLow
    End If
    
    If Me.SigTestVs <> "" Then
        LegendRange.Rows(3).FormulaR1C1 = "x% = significantly higher than Total Genpop Sample (95% confidence)"
        LegendRange.Rows(3).Font.Color = Globals.ColorSignificantlyHigher
        LegendRange.Rows(4).FormulaR1C1 = "x% = significantly lower than Total Genpop Sample (95% confidence)"
        LegendRange.Rows(4).Font.Color = Globals.ColorSignificantlyLower
    End If
    
End Sub

Public Sub HighlightBases()

    If Me.BaseIn = "SubTable" Then
        Dim st As SubTable
        For Each st In Me.SubTables
            HighlightBaseRange st.BaseRange
        Next st
    ElseIf Me.BaseIn = "Row" Or Me.BaseIn = "Column" Then
        HighlightBaseRange Me.BaseRange
    End If
    
End Sub

Private Sub HighlightBaseRange(rg As Range)
    Dim c As Range
    For Each c In rg.Cells
        If c < Globals.BaseLow Then c.Interior.Color = Globals.ColorBaseLow
        If c < Globals.BaseTooLow Then c.Interior.Color = Globals.ColorBaseTooLow
    Next c
End Sub

Public Sub HighlightSignificances()
    
    If Me.BaseIn = "Row" Or Me.BaseIn = "Column" Then
        If Me.SigTestVs = "Ref" Then
            HighlightSignficancesVsRef
        End If
    ElseIf Me.BaseIn = "SubTable" Then
        Dim st As SubTable
        For Each st In Me.SubTables
            st.HighlightSignificances
        Next st
    End If
    
End Sub

Private Sub HighlightSignficancesVsRef()

    Dim i As Integer, j As Integer
    Dim x1 As Double, x2 As Double, n1 As Double, n2 As Double

    For i = 1 To Me.TestRange.Rows.Count
        For j = 1 To Me.TestRange.Columns.Count
    
            If Me.BaseIn = "Row" Then
                n1 = Me.ReferenceRange.Cells(1, 1).Value 'reference base
                x1 = Me.ReferenceRange.Cells(i + 1, 1).Value / 100 'reference value
                n2 = Me.BaseRange.Cells(1, j + 1).Value 'test base
                x2 = Me.TestRange.Cells(i, j).Value / 100 'test value
            ElseIf Me.BaseIn = "Column" Then
                n1 = Me.ReferenceRange.Cells(1, 1).Value 'reference base
                x1 = Me.ReferenceRange.Cells(1, j + 1).Value / 100 'reference value
                n2 = Me.BaseRange.Cells(i + 1, 1).Value 'test base
                x2 = Me.TestRange.Cells(i, j).Value / 100 'test value
            End If
            
            If Stats.IsSignificant(Globals.SignificanceLevel, Stats.GetProbability(n1, x1, n2, x2)) Then
                If x2 > x1 Then
                    Me.TestRange.Cells(i, j).Font.Color = Globals.ColorSignificantlyHigher
                Else
                    Me.TestRange.Cells(i, j).Font.Color = Globals.ColorSignificantlyLower
                End If
            End If
        
        Next j
    Next i
    
End Sub


Public Sub PostFormatting()

    If Me.SigTestVs = "All" Then
        Union(Me.BannerRange, Me.DataRange).Columns.AutoFit
    End If
    
    Dim c As Range
    For Each c In Me.TitleRange.Cells
        If Left(c, 10) = "Conversion" Then
            c.EntireRow.Font.Italic = True
            c.EntireRow.Font.Color = rgbGrey
        End If
    Next c
    
End Sub