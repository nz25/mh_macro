Attribute VB_Name = "Globals"
Option Explicit

'GLOBAL SETTINGS

Public Const BaseLow As Integer = 80
Public Const BaseTooLow As Integer = 60
Public Const SignificanceLevel As Double = 0.95
Public Const SignificanceCharToUseIfBaseIsTooLow As String = "*"

Public Const ColorBaseLow As Long = 49407
Public Const ColorBaseTooLow As Long = 26367
Public Const ColorSignificantlyHigher As Long = 5287936
Public Const ColorSignificantlyLower As Long = 255

Public Const IsSignficantErrorCode As Integer = 1001
Public Const GetProbabilityParameterErrorCode As Integer = 1002
Public Const DataRangeErrorCode As Integer = 1003

Public Const Font As String = "Arial Narrow"
Public Const TableNameCellAddress As String = "A2"

Public Letters(1 To 52) As String

