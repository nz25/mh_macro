VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DataRange As Range
Public BaseRange As Range
Public ReferenceRange As Range
Public TestRange As Range
Public TitleRange As Range
Public ParentTable As ITable

Public Sub Initialize(t As ITable, rg As Range)
    Set Me.ParentTable = t
    Set Me.DataRange = rg
    Set Me.BaseRange = rg.Resize(1)
    Set Me.TestRange = rg.Offset(1).Resize(rg.Rows.Count - 1)
    Set Me.TitleRange = Me.TestRange.Offset(0, -1).Columns(1)
    
    If Me.ParentTable.GetSigTestVs = "Ref" Or Me.ParentTable.GetSigTestVs = "Delta" Then
        Set Me.ReferenceRange = Me.DataRange.Columns(1)
        Set Me.TestRange = Me.TestRange.Offset(0, 1).Resize(, Me.TestRange.Columns.Count - 1)
        Set Me.BaseRange = Me.BaseRange.Offset(0, 1).Resize(, Me.BaseRange.Columns.Count - 1)
    End If
    
End Sub

Public Sub HighlightSignificances()
    
    If Me.ParentTable.GetSigTestVs = "Ref" Then
        HighlightSignficancesVsRef
    ElseIf Me.ParentTable.GetSigTestVs = "All" Then
        HighlightSignificancesVsAll
    ElseIf Me.ParentTable.GetSigTestVs = "Delta" Then
        HighlightSignficancesVsDelta
    End If
    
End Sub

Private Sub HighlightSignficancesVsRef()

    Dim i As Integer, j As Integer
    Dim x1 As Double, x2 As Double, n1 As Double, n2 As Double

    For i = 1 To Me.TestRange.Rows.Count
        For j = 1 To Me.TestRange.Columns.Count

            If Left(Me.TitleRange.Rows(i), 10) <> "Conversion" Then
                n1 = Me.ReferenceRange.Cells(1, 1).Value 'reference base
                x1 = Me.ReferenceRange.Cells(i + 1, 1).Value / 100 'reference value
                n2 = Me.BaseRange.Cells(1, j).Value 'test base
                x2 = Me.TestRange.Cells(i, j).Value / 100 'test value
            
                If Stats.IsSignificant(Globals.SignificanceLevel, Stats.GetProbability(n1, x1, n2, x2, SignificanceTestType.Linda)) Then
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

Private Sub HighlightSignficancesVsDelta()

    Dim i As Integer, j As Integer
    Dim ref As Double, x1 As Double, n1 As Double

    For i = 1 To Me.TestRange.Rows.Count

        For j = 1 To Me.TestRange.Columns.Count

            If Me.TitleRange.Rows(i) <> "" Then
                ref = Me.ReferenceRange.Cells(i + 1, 1).Value / 100 'reference value
                n1 = Me.BaseRange.Cells(1, j).Value 'test base
                x1 = Me.TestRange.Cells(i, j).Value / 100 'test value
            
                If Stats.IsSignificant(Globals.SignificanceLevel, Stats.GetProbabilityRef(n1, x1, ref)) Then
                    If x1 > ref Then
                        Me.TestRange.Cells(i, j).Font.Color = Globals.ColorSignificantlyHigher
                    Else
                        Me.TestRange.Cells(i, j).Font.Color = Globals.ColorSignificantlyLower
                    End If
                End If
            
            End If
        
        Next j
    Next i
 
End Sub

Private Sub HighlightSignificancesVsAll()

    Dim i As Integer, j As Integer, k As Integer
    Dim n1 As Double, x1 As Double, n2 As Double, x2 As Double
    Dim significanceDictionary As Scripting.Dictionary
    
    'Calculates signficances for "real" variable categories.
    'It's assumed that derived categories (conversion rates) are between all "real" categories.
    'These are excluded by using "Step 4" below
        
    For i = 1 To Me.TestRange.Rows.Count
    
        If Left(Me.TitleRange.Rows(i), 10) <> "Conversion" And Me.TitleRange.Rows(i) <> "" Then
            
            For j = 1 To Me.TestRange.Columns.Count 'test columns
                n1 = Me.BaseRange.Cells(1, j).Value
                x1 = Me.TestRange.Cells(i, j).Value / 100
                
                If n1 < Globals.BaseTooLow Then
                    WriteSignificancesIfBaseTooSmall Me.TestRange.Cells(i + 1, j)
                Else
                    Set significanceDictionary = New Scripting.Dictionary
                    For k = 1 To Me.TestRange.Columns.Count 'reference columns
                        n2 = Me.BaseRange.Cells(1, k).Value
                        x2 = Me.TestRange.Cells(i, k).Value / 100
                        If Stats.IsSignificant(Globals.SignificanceLevel, Stats.GetProbability(n1, x1, n2, x2, SignificanceTestType.Normal)) Then
                            Dim key As String
                            key = GetHeaderKey(k)
                            If x1 > x2 Then
                                significanceDictionary.Add key, Globals.ColorSignificantlyHigher
                            Else
                                significanceDictionary.Add key, Globals.ColorSignificantlyLower
                            End If
                        End If
                    Next k
                    WriteSignificances Me.TestRange.Cells(i + 1, j), significanceDictionary
                End If
            Next j
        End If
    
    Next i
    
End Sub

Private Function GetHeaderKey(columnIndex As Integer) As String
    GetHeaderKey = Me.ParentTable.GetBannerRange.Cells(2, columnIndex)
End Function

Private Function WriteSignificances(cell As Range, dict As Scripting.Dictionary)
    Dim stringToWrite As String
    Dim key As Variant
    For Each key In dict.Keys
        stringToWrite = stringToWrite & key
    Next
    cell.Value = stringToWrite
    Dim i As Integer
    For i = 1 To dict.Count
        cell.Characters(i, 1).Font.Color = dict.Items(i - 1)
    Next i
End Function

Private Function WriteSignificancesIfBaseTooSmall(cell As Range)
    cell.Value = Globals.SignificanceCharToUseIfBaseIsTooLow
End Function

