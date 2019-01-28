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
Implements ITable

Public TableName As String
Public worksheetName As String
Public TitleLevels As Integer
Public BannerLevels As Integer
Public BaseIn As String
Public SigTestVs As String
Public SigText As String
Public HasDelta As Boolean
Public UseFilter As Boolean
Public OpenEnds As Boolean
Public SigLevel As Double
Public TabColor As Long

Public Worksheet As Worksheet

Public DataRange As Range
Public ReferenceRange As Range
Public BaseRange As Range
Public TestRange As Range
Public HeaderRange As Range
Public LegendRange As Range
Public TitleRange As Range
Public BannerRange As Range
Public AutoFilterRange As Range

Public SubTables As Collection
Public DeltaTable As DeltaTable

Public TableIsEmpty As Boolean

Public Function ITable_GetSigLevel() As Double
    ITable_GetSigLevel = Me.SigLevel
End Function

Public Function ITable_GetSigTestVs() As String
    ITable_GetSigTestVs = Me.SigTestVs
End Function

Public Function ITable_GetBannerRange() As Range
    Set ITable_GetBannerRange = Me.BannerRange
End Function

Public Property Get FirstDataRow() As Integer
    FirstDataRow = Me.HeaderRange.Rows.Count + Me.BannerLevels * 2 + 2
End Property

Public Property Get FirstDataColumn() As Integer
    FirstDataColumn = Me.TitleLevels * 2 + 1
End Property

Public Sub Initialize(index As Range, ws As Worksheet)

    Me.TableName = index.Columns(1)
    Me.worksheetName = index.Columns(2)
    Me.TitleLevels = index.Columns(3)
    Me.BannerLevels = index.Columns(4)
    Me.BaseIn = index.Columns(5)
    Me.SigTestVs = index.Columns(6)
    Me.SigText = index.Columns(7)
    Me.HasDelta = index.Columns(8)
    Me.UseFilter = index.Columns(9)
    Me.OpenEnds = index.Columns(10)
    If Me.OpenEnds Then
        Me.SigLevel = 0.999
    Else
        Me.SigLevel = 0.95
    End If
    
    Me.TabColor = index.Columns(1).Interior.Color
    
    Set Me.Worksheet = ws
    
    Dim lastRow As Integer, lastColumn As Integer
    
    lastRow = Me.Worksheet.UsedRange.Rows.Count
    lastColumn = Me.Worksheet.UsedRange.Columns.Count
    
    Set Me.HeaderRange = Me.Worksheet.Cells(1, 1).CurrentRegion
    Set Me.DataRange = Me.Worksheet.Range(Me.Worksheet.Cells(Me.FirstDataRow, Me.FirstDataColumn), Me.Worksheet.Cells(lastRow, lastColumn))
    Set Me.TitleRange = Me.DataRange.Offset(0, -1).Columns(1)
    Set Me.BannerRange = Me.DataRange.Offset(-1, 0).Rows(1)
    Set Me.LegendRange = Me.HeaderRange.Offset(0, 3).Resize(4)
    
    If Me.DataRange.Rows.Count <= 1 Then Me.TableIsEmpty = True
    If Me.TableIsEmpty Then Exit Sub
        
    If Me.SigTestVs = "Trend" Or Me.SigTestVs = "2Columns" Then
        If Me.BaseIn = "Row" Then
            Set Me.BaseRange = Me.DataRange.Rows(1)
            Set Me.TestRange = Me.DataRange.Offset(1, 0).Resize(Me.DataRange.Rows.Count - 1)
        ElseIf Me.BaseIn = "Column" Then
            Set Me.BaseRange = Me.DataRange.Resize(, Globals.NumberOfWaves)
            Set Me.TestRange = Me.DataRange.Offset(0, Globals.NumberOfWaves).Resize(, Me.DataRange.Columns.Count - Globals.NumberOfWaves)
        ElseIf Me.BaseIn = "SubTable" Then
            'All ranges are defined in subtable class
        End If
    Else
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
    End If
    
    If Me.HasDelta Then
        Set Me.DeltaTable = New DeltaTable
        Me.DeltaTable.Initialize Me
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
    
    If Me.HasDelta Then Me.DeltaTable.SetupSubTables

End Sub

Public Sub PreFormatting()
    
    With Me.Worksheet.UsedRange
        .Interior.Pattern = xlNone
        .Font.Color = XlRgbColor.rgbBlack
        .Font.Italic = False
    End With
    
    Me.Worksheet.Name = Left(Me.HeaderRange.Rows(2), 31)
    Me.Worksheet.Tab.Color = Me.TabColor
    Me.HeaderRange.Rows(Me.HeaderRange.Rows.Count).Font.Bold = True
    
    With Me.Worksheet.UsedRange.Resize(Me.Worksheet.UsedRange.Rows.Count + 1)
        .Font.Name = Globals.Font
        .NumberFormat = "0"
        '.VerticalAlignment = xlCenter
    End With
    
    If Me.OpenEnds Then
        'Formatting open-ends?
    End If
    
    Me.TitleRange.Columns.AutoFit
    Me.BannerRange.Font.Size = 8
    Me.BannerRange.Font.Bold = False
    Me.BannerRange.Rows.AutoFit
    Me.BannerRange.WrapText = True
    Me.BannerRange.HorizontalAlignment = xlCenter
        
    Me.DataRange.HorizontalAlignment = xlCenter
    Me.DataRange.Replace What:="XXXXX", Replacement:="n.a."
    
    If Me.HasDelta Then Me.DeltaTable.PreFormatting
    
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
        LegendRange.Rows(3).FormulaR1C1 = Replace(Me.SigText, "(higher/lower)", "higher")
        LegendRange.Rows(3).Font.Color = Globals.ColorSignificantlyHigher
        LegendRange.Rows(4).FormulaR1C1 = Replace(Me.SigText, "(higher/lower)", "lower")
        LegendRange.Rows(4).Font.Color = Globals.ColorSignificantlyLower
    End If
    
End Sub

Public Sub HighlightBases()

    If Me.BaseIn = "SubTable" Then
        Dim st As SubTable
        For Each st In Me.SubTables
        If Me.SigTestVs = "Ref" Or Me.SigTestVs = "Delta" Then
            HighlightBaseRange st.BaseRange.Offset(0, -1).Resize(, st.BaseRange.Columns.Count + 1)
        Else
            HighlightBaseRange st.BaseRange
        End If
                
        Next st
    ElseIf Me.BaseIn = "Row" Or Me.BaseIn = "Column" Then
        HighlightBaseRange Me.BaseRange
    End If
    
    If Me.HasDelta Then Me.DeltaTable.HighlightBases
    
End Sub

Public Sub HighlightBaseRange(rg As Range)
    Dim c As Range
    For Each c In rg.Cells
        If c < Globals.BaseLow Then c.Interior.Color = Globals.ColorBaseLow
        If c < Globals.BaseTooLow Then c.Interior.Color = Globals.ColorBaseTooLow
    Next c
End Sub

Public Sub HighlightSignificances()
    
    Dim testType As SignificanceTestType
    
    If Me.BaseIn = "SubTable" Then
        Dim st As SubTable
        For Each st In Me.SubTables
            st.HighlightSignificances
        Next st
    ElseIf Me.SigTestVs = "Ref" Then
        HighlightSignficancesVsRef SignificanceTestType.Linda
    ElseIf Me.SigTestVs = "Trend" Then
        HighlightSignificancesTrend
    ElseIf Me.SigTestVs = "2Columns" Then
        HighlightSignificances2Columns
    End If
    
    If Me.HasDelta Then Me.DeltaTable.HighlightSignificances
    
End Sub

Private Sub HighlightSignficancesVsRef(testType As SignificanceTestType)

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
                        
            If Stats.IsSignificant(Me.SigLevel, Stats.GetProbability(n1, x1, n2, x2, testType)) Then
                If x2 > x1 Then
                    Me.TestRange.Cells(i, j).Font.Color = Globals.ColorSignificantlyHigher
                Else
                    Me.TestRange.Cells(i, j).Font.Color = Globals.ColorSignificantlyLower
                End If
            End If
        
        Next j
    Next i
    
End Sub

Private Sub HighlightSignificancesTrend()

    Dim i As Integer, j As Integer, currentWave As Integer
    Dim x1 As Double, x2 As Double, n1 As Double, n2 As Double

    For i = 1 To Me.TestRange.Rows.Count
        For j = 1 To Me.TestRange.Columns.Count
        
            currentWave = j Mod Globals.NumberOfWaves
            If currentWave = 0 Then currentWave = Globals.NumberOfWaves
        
            If currentWave > 1 Then 'conduct the test starting from the 2nd wave (skip the 1st wave)
        
                If Me.BaseIn = "Row" Then
                    n1 = Me.BaseRange.Cells(1, j - 1).Value 'reference base
                    x1 = Me.TestRange.Cells(i, j - 1).Value / 100 'reference value
                    n2 = Me.BaseRange.Cells(1, j).Value 'test base
                    x2 = Me.TestRange.Cells(i, j).Value / 100 'test value
                ElseIf Me.BaseIn = "Column" Then
                    n1 = Me.BaseRange.Cells(i, currentWave - 1).Value 'reference base
                    x1 = Me.TestRange.Cells(i, j - 1).Value / 100 'reference value
                    n2 = Me.BaseRange.Cells(i, currentWave).Value 'test base
                    x2 = Me.TestRange.Cells(i, j).Value / 100 'test value
                End If
                
                If Stats.IsSignificant(Me.SigLevel, Stats.GetProbability(n1, x1, n2, x2, SignificanceTestType.Trend)) Then
                    If x2 > x1 Then
                        Me.TestRange.Cells(i, j).Font.Color = Globals.ColorSignificantlyHigher
                    Else
                        Me.TestRange.Cells(i, j).Font.Color = Globals.ColorSignificantlyLower
                    End If
                End If
            End If
        Next j
    Next i
    
End Sub

Private Sub HighlightSignificances2Columns()

    Dim i As Integer, j As Integer
    Dim x1 As Double, x2 As Double, n1 As Double, n2 As Double

    For i = 1 To Me.TestRange.Rows.Count
        For j = 1 To Me.TestRange.Columns.Count
        
            If j Mod 2 = 0 Then 'conduct the test only for 2nd column
        
                n1 = Me.BaseRange.Cells(1, j - 1).Value 'reference base
                x1 = Me.TestRange.Cells(i, j - 1).Value / 100 'reference value
                n2 = Me.BaseRange.Cells(1, j).Value 'test base
                x2 = Me.TestRange.Cells(i, j).Value / 100 'test value
            
                If Stats.IsSignificant(Me.SigLevel, Stats.GetProbability(n1, x1, n2, x2, SignificanceTestType.Trend)) Then
                    If x2 > x1 Then
                        Me.TestRange.Cells(i, j).Font.Color = Globals.ColorSignificantlyHigher
                    Else
                        Me.TestRange.Cells(i, j).Font.Color = Globals.ColorSignificantlyLower
                    End If
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
    
    If Me.HasDelta Then Me.DeltaTable.CalculateDeltas
    
    If Me.UseFilter Then
        Set Me.AutoFilterRange = Me.BannerRange.Cells(Me.BannerRange.Rows.Count, 1).Offset(0, -1)
        Me.AutoFilterRange.Offset(0, -1).Resize(1, 2).AutoFilter
        Me.AutoFilterRange = "Show/hide rows with filter"
        Me.AutoFilterRange.Font.Italic = True
    End If
    
    FixWindow
    
End Sub

Private Sub FixWindow()

    Dim rowToFix As Integer, columnToFix As Integer
    
    If Me.BaseIn = "SubTable" And Me.SigTestVs = "Ref" Then
        'ohne basen mit ref
        rowToFix = Me.DataRange.Rows(1).Row
        columnToFix = Me.DataRange.Columns(1).Column + 1
    ElseIf Me.BaseIn = "SubTable" And Me.SigTestVs <> "Ref" Then
        'ohne basen ohne ref
        rowToFix = Me.DataRange.Rows(1).Row
        columnToFix = Me.DataRange.Columns(1).Column
    ElseIf Me.SigTestVs = "Ref" Then
        'mit basen mit ref
        rowToFix = Me.BaseRange.Row + 1
        columnToFix = Me.ReferenceRange.Column + 1
    ElseIf Me.SigTestVs <> "Ref" Then
        'mit basen ohne ref
        rowToFix = Me.DataRange.Rows(1).Row + 1
        columnToFix = Me.DataRange.Columns(1).Column
    End If

    Me.Worksheet.Select
    Me.Worksheet.Cells(rowToFix, columnToFix).Select
    ActiveWindow.FreezePanes = True
    
End Sub
