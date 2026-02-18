Attribute VB_Name = "UtilityFunctions"
'---------------------------------------------------------------------------------------
' Module    : UtilityFunctions
' Author    : personalityson
' Date      : 01.05.2025
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public Const MATH_PI As Double = 3.14159265358979    'pi
Public Const MATH_2PI As Double = 6.28318530717959   '2*pi
Public Const MATH_PI2 As Double = 1.5707963267949    'pi/2
Public Const MATH_RPI As Double = 0.318309886183791  '1/pi
Public Const MATH_SQRT2 As Double = 1.4142135623731  'sqrt(2)
Public Const MATH_SQRT3 As Double = 1.73205080756888 'sqrt(3)
Public Const MATH_LN2 As Double = 0.693147180559945  'ln(2)
Public Const MATH_LN3 As Double = 1.09861228866811   'ln(3)
Public Const MATH_E As Double = 2.71828182845905     'e
Public Const MATH_RE As Double = 0.367879441171442   '1/e

Public Const DOUBLE_MIN_ABS As Double = 4.94065645841247E-324
Public Const DOUBLE_MAX_ABS As Double = 1.79769313486231E+308
Public Const DOUBLE_MIN_LOG As Double = -744.440071921381
Public Const DOUBLE_MAX_LOG As Double = 709.782712893384
Public Const DOUBLE_EPSILON As Double = 2.22044604925031E-16

Public Const SIZEOF_INTEGER As Long = 2
Public Const SIZEOF_LONG As Long = 4
Public Const SIZEOF_SINGLE As Long = 4
Public Const SIZEOF_DOUBLE As Long = 8

#If Win64 Then
    Public Const NULL_PTR As LongPtr = 0^
    Public Const SIZEOF_LONGPTR As Long = 8
    Public Const SIZEOF_VARIANT As Long = 24
#Else
    Public Const NULL_PTR As LongPtr = 0&
    Public Const SIZEOF_LONGPTR As Long = 4
    Public Const SIZEOF_VARIANT As Long = 16
#End If

Public Enum RoundingType
    rtNearest
    rtDown
    rtUp
    rtTowardsZero
    rtTowardsInfinity
End Enum

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMillisecond As Integer
End Type

Public Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                                                                ByRef Source As Any, _
                                                                                ByVal Length As LongPtr)

Public Declare PtrSafe Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, _
                                                                                ByVal Length As LongPtr)

Public Declare PtrSafe Function VarPtrArray Lib "VBE7.dll" Alias "VarPtr" (ByRef Var() As Any) As LongPtr

Public Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long

Public Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Sub GetSystemTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME)

Public Function MinLng2(ByVal A As Long, _
                        ByVal B As Long) As Long
    If A < B Then
        MinLng2 = A
    Else
        MinLng2 = B
    End If
End Function

Public Function MaxLng2(ByVal A As Long, _
                        ByVal B As Long) As Long
    If A > B Then
        MaxLng2 = A
    Else
        MaxLng2 = B
    End If
End Function

Public Function MinPtr2(ByVal A As LongPtr, _
                        ByVal B As LongPtr) As LongPtr
    If A < B Then
        MinPtr2 = A
    Else
        MinPtr2 = B
    End If
End Function

Public Function MaxPtr2(ByVal A As LongPtr, _
                        ByVal B As LongPtr) As LongPtr
    If A > B Then
        MaxPtr2 = A
    Else
        MaxPtr2 = B
    End If
End Function

Public Function MinDbl2(ByVal A As Double, _
                        ByVal B As Double) As Double
    If A < B Then
        MinDbl2 = A
    Else
        MinDbl2 = B
    End If
End Function

Public Function MaxDbl2(ByVal A As Double, _
                        ByVal B As Double) As Double
    If A > B Then
        MaxDbl2 = A
    Else
        MaxDbl2 = B
    End If
End Function

Public Function MinLng3(ByVal A As Long, _
                        ByVal B As Long, _
                        ByVal C As Long) As Long
    MinLng3 = A
    If MinLng3 > B Then
        MinLng3 = B
    End If
    If MinLng3 > C Then
        MinLng3 = C
    End If
End Function

Public Function MaxLng3(ByVal A As Long, _
                        ByVal B As Long, _
                        ByVal C As Long) As Long
    MaxLng3 = A
    If MaxLng3 < B Then
        MaxLng3 = B
    End If
    If MaxLng3 < C Then
        MaxLng3 = C
    End If
End Function

Public Function MinPtr3(ByVal A As LongPtr, _
                        ByVal B As LongPtr, _
                        ByVal C As LongPtr) As LongPtr
    MinPtr3 = A
    If MinPtr3 > B Then
        MinPtr3 = B
    End If
    If MinPtr3 > C Then
        MinPtr3 = C
    End If
End Function

Public Function MaxPtr3(ByVal A As LongPtr, _
                        ByVal B As LongPtr, _
                        ByVal C As LongPtr) As LongPtr
    MaxPtr3 = A
    If MaxPtr3 < B Then
        MaxPtr3 = B
    End If
    If MaxPtr3 < C Then
        MaxPtr3 = C
    End If
End Function

Public Function MinDbl3(ByVal A As Double, _
                        ByVal B As Double, _
                        ByVal C As Double) As Double
    MinDbl3 = A
    If MinDbl3 > B Then
        MinDbl3 = B
    End If
    If MinDbl3 > C Then
        MinDbl3 = C
    End If
End Function

Public Function MaxDbl3(ByVal A As Double, _
                        ByVal B As Double, _
                        ByVal C As Double) As Double
    MaxDbl3 = A
    If MaxDbl3 < B Then
        MaxDbl3 = B
    End If
    If MaxDbl3 < C Then
        MaxDbl3 = C
    End If
End Function

Public Function RoundToMultiple(ByVal dblValue As Double, _
                                ByVal dblMultiple As Double, _
                                ByVal eRoundingType As RoundingType) As Double
    If dblMultiple = 0 Then
        RoundToMultiple = dblValue
        Exit Function
    End If
    dblMultiple = Abs(dblMultiple)
    Select Case eRoundingType
        Case rtNearest
            RoundToMultiple = Round(dblValue / dblMultiple) * dblMultiple
        Case rtDown
            RoundToMultiple = Int(dblValue / dblMultiple) * dblMultiple
        Case rtUp
            RoundToMultiple = -Int(-dblValue / dblMultiple) * dblMultiple
        Case rtTowardsZero
            RoundToMultiple = Sgn(dblValue) * Int(Abs(dblValue) / dblMultiple) * dblMultiple
        Case rtTowardsInfinity
            RoundToMultiple = Sgn(dblValue) * -Int(-Abs(dblValue) / dblMultiple) * dblMultiple
    End Select
End Function

Public Function RoundToSignificantDigits(ByVal dblValue As Double, _
                                         ByVal lNumDigits As Long, _
                                         ByVal eRoundingType As RoundingType) As Double
    Dim dblMultiple As Double
    
    If dblValue = 0 Then
        Exit Function
    End If
    If lNumDigits < 1 Then
        Exit Function
    End If
    dblMultiple = 10 ^ (Int(Log(Abs(dblValue)) / Log(10)) + 1 - lNumDigits)
    RoundToSignificantDigits = RoundToMultiple(dblValue, dblMultiple, eRoundingType)
End Function

Public Function GetRank(ByVal vArray As Variant) As Integer
    Const VARIANT_OFFSET_parray As Long = 8
    Dim iVarType As Integer
    Dim pSafeArray As LongPtr

    CopyMemory iVarType, vArray, SIZEOF_INTEGER
    If (iVarType And vbArray) = 0 Then
        GetRank = -1 'A scalar
        Exit Function
    End If
    CopyMemory pSafeArray, ByVal VarPtr(vArray) + VARIANT_OFFSET_parray, SIZEOF_LONGPTR
    If pSafeArray = NULL_PTR Then
        GetRank = 0 'Uninitialized
        Exit Function
    End If
    CopyMemory GetRank, ByVal pSafeArray, SIZEOF_INTEGER
End Function

Public Function EnsureArray(ByVal vValueOrArray As Variant) As Variant
    Const PROCEDURE_NAME As String = "EnsureArray"
    
    Select Case GetRank(vValueOrArray)
        Case -1
            EnsureArray = Array(vValueOrArray)
        Case 0
            EnsureArray = Array()
        Case 1
            EnsureArray = vValueOrArray
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Expecting a acalar, an uninitialized array, or a one-dimensional array."
    End Select
End Function

Public Sub ParseVariantToLongArray(ByVal vValueOrArray As Variant, _
                                   ByRef lNumElements As Long, _
                                   ByRef alArray() As Long)
    Const PROCEDURE_NAME As String = "UtilityFunctions.ParseVariantToLongArray"
    Dim i As Long
    Dim lLBound As Long
    Dim lUBound As Long
    Dim vArray As Variant

    Select Case GetRank(vValueOrArray)
        Case -1
            vArray = Array(vValueOrArray)
        Case 0
            vArray = Array()
        Case 1
            vArray = vValueOrArray
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Expected a single value, an uninitialized array, or a one-dimensional array."
    End Select
    lLBound = LBound(vArray)
    lUBound = UBound(vArray)
    If lLBound > lUBound Then
        lNumElements = 0
        Erase alArray
    Else
        lNumElements = lUBound - lLBound + 1
        ReDim alArray(1 To lNumElements)
        For i = 1 To lNumElements
            alArray(i) = CLng(vArray(lLBound + i - 1))
        Next i
    End If
End Sub

Public Sub ParseVariantToDoubleArray(ByVal vValueOrArray As Variant, _
                                     ByRef lNumElements As Long, _
                                     ByRef adblArray() As Double)
    Const PROCEDURE_NAME As String = "UtilityFunctions.ParseVariantToLongArray"
    Dim i As Long
    Dim lLBound As Long
    Dim lUBound As Long
    Dim vArray As Variant

    Select Case GetRank(vValueOrArray)
        Case -1
            vArray = Array(vValueOrArray)
        Case 0
            vArray = Array()
        Case 1
            vArray = vValueOrArray
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Expected a single value, an uninitialized array, or a one-dimensional array."
    End Select
    lLBound = LBound(vArray)
    lUBound = UBound(vArray)
    If lLBound > lUBound Then
        lNumElements = 0
        Erase adblArray
    Else
        lNumElements = lUBound - lLBound + 1
        ReDim adblArray(1 To lNumElements)
        For i = 1 To lNumElements
            adblArray(i) = CDbl(vArray(lLBound + i - 1))
        Next i
    End If
End Sub

Public Function GetIdentityPermutationArray(ByVal lNumElements As Long) As Long()
    Dim i As Long
    Dim alResult() As Long

    If lNumElements < 1 Then
        Exit Function
    End If
    ReDim alResult(1 To lNumElements)
    For i = 1 To lNumElements
        alResult(i) = i
    Next i
    GetIdentityPermutationArray = alResult
End Function

Public Function GetRandomPermutationArray(ByVal lNumElements As Long) As Long()
    Dim i As Long
    Dim j As Long
    Dim alResult() As Long
    
    If lNumElements < 1 Then
        Exit Function
    End If
    ReDim alResult(1 To lNumElements)
    For i = 1 To lNumElements
        j = Int(i * Rnd + 1)
        alResult(i) = alResult(j)
        alResult(j) = i
    Next i
    GetRandomPermutationArray = alResult
End Function

Public Function Union(ByVal oRangeA As Range, _
                      ByVal oRangeB As Range) As Range
    Const PROCEDURE_NAME As String = "UtilityFunctions.Union"
    
    If oRangeA Is Nothing Then
        Set Union = oRangeB
        Exit Function
    End If
    If oRangeB Is Nothing Then
        Set Union = oRangeA
        Exit Function
    End If
    If Not oRangeA.Worksheet Is oRangeB.Worksheet Then
        Err.Raise 5, PROCEDURE_NAME, "Specified ranges are not on the same worksheet."
    End If
    Set Union = Application.Union(oRangeA, oRangeB)
End Function

Public Function Intersect(ByVal oRangeA As Range, _
                          ByVal oRangeB As Range) As Range
    If oRangeA Is Nothing Then
        Exit Function
    End If
    If oRangeB Is Nothing Then
        Exit Function
    End If
    If Not oRangeA.Worksheet Is oRangeB.Worksheet Then
        Exit Function
    End If
    Set Intersect = Application.Intersect(oRangeA, oRangeB)
End Function

Public Function Complement(ByVal oRangeA As Range, _
                           ByVal oRangeB As Range) As Range
    Dim oAreaA As Range
    Dim oAreaB As Range
    Dim lStartRowA As Long
    Dim lStartColA As Long
    Dim lEndRowA As Long
    Dim lEndColA As Long
    Dim lStartRowB As Long
    Dim lStartColB As Long
    Dim lEndRowB As Long
    Dim lEndColB As Long
    Dim lIntersectStartRow As Long
    Dim lIntersectStartCol As Long
    Dim lIntersectEndRow As Long
    Dim lIntersectEndCol As Long
    Dim oResult As Range
    Dim oResultCopy As Range

    If oRangeA Is Nothing Then
        Exit Function
    End If
    If oRangeB Is Nothing Then
        Set Complement = oRangeA
        Exit Function
    End If
    If Not oRangeA.Worksheet Is oRangeB.Worksheet Then
        Set Complement = oRangeA
        Exit Function
    End If
    Set oResult = oRangeA
    With oRangeA.Worksheet
        For Each oAreaB In oRangeB.Areas
            If oResult Is Nothing Then
                Exit For
            End If
            lStartRowB = oAreaB.Row
            lStartColB = oAreaB.Column
            lEndRowB = lStartRowB + oAreaB.Rows.Count - 1
            lEndColB = lStartColB + oAreaB.Columns.Count - 1
            Set oResultCopy = oResult
            Set oResult = Nothing
            For Each oAreaA In oResultCopy.Areas
                lStartRowA = oAreaA.Row
                lStartColA = oAreaA.Column
                lEndRowA = lStartRowA + oAreaA.Rows.Count - 1
                lEndColA = lStartColA + oAreaA.Columns.Count - 1
                lIntersectStartRow = MaxLng2(lStartRowA, lStartRowB)
                lIntersectStartCol = MaxLng2(lStartColA, lStartColB)
                lIntersectEndRow = MinLng2(lEndRowA, lEndRowB)
                lIntersectEndCol = MinLng2(lEndColA, lEndColB)
                If lIntersectStartRow <= lIntersectEndRow And lIntersectStartCol <= lIntersectEndCol Then
                    If lIntersectStartRow > lStartRowA Then
                        Set oResult = Union(oResult, .Range(.Cells(lStartRowA, lStartColA), .Cells(lIntersectStartRow - 1, lEndColA)))
                    End If
                    If lIntersectStartCol > lStartColA Then
                        Set oResult = Union(oResult, .Range(.Cells(lIntersectStartRow, lStartColA), .Cells(lIntersectEndRow, lIntersectStartCol - 1)))
                    End If
                    If lEndColA > lIntersectEndCol Then
                        Set oResult = Union(oResult, .Range(.Cells(lIntersectStartRow, lIntersectEndCol + 1), .Cells(lIntersectEndRow, lEndColA)))
                    End If
                    If lEndRowA > lIntersectEndRow Then
                        Set oResult = Union(oResult, .Range(.Cells(lIntersectEndRow + 1, lStartColA), .Cells(lEndRowA, lEndColA)))
                    End If
                Else
                    Set oResult = Union(oResult, oAreaA)
                End If
            Next oAreaA
        Next oAreaB
    End With
    Set Complement = oResult
End Function

Public Function GetFirstRow(ByVal oWorksheet As Worksheet, _
                            Optional ByVal lColumn As Long) As Long
    Const PROCEDURE_NAME As String = "UtilityFunctions.GetFirstRow"
    Dim oNonEmptyCell As Range
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Worksheet object is required."
    End If
    If lColumn > 0 Then
        Set oNonEmptyCell = oWorksheet.Columns(lColumn).Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    Else
        Set oNonEmptyCell = oWorksheet.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    End If
    If Not oNonEmptyCell Is Nothing Then
        GetFirstRow = oNonEmptyCell.Row
    End If
End Function

Public Function GetLastRow(ByVal oWorksheet As Worksheet, _
                           Optional ByVal lColumn As Long) As Long
    Const PROCEDURE_NAME As String = "UtilityFunctions.GetLastRow"
    Dim oNonEmptyCell As Range
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Worksheet object is required."
    End If
    If lColumn > 0 Then
        Set oNonEmptyCell = oWorksheet.Columns(lColumn).Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    Else
        Set oNonEmptyCell = oWorksheet.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    End If
    If Not oNonEmptyCell Is Nothing Then
        GetLastRow = oNonEmptyCell.Row
    End If
End Function

Public Function GetFirstColumn(ByVal oWorksheet As Worksheet, _
                               Optional ByVal lRow As Long) As Long
    Const PROCEDURE_NAME As String = "UtilityFunctions.GetFirstColumn"
    Dim oNonEmptyCell As Range
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Worksheet object is required."
    End If
    If lRow > 0 Then
        Set oNonEmptyCell = oWorksheet.Rows(lRow).Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    Else
        Set oNonEmptyCell = oWorksheet.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    End If
    If Not oNonEmptyCell Is Nothing Then
        GetFirstColumn = oNonEmptyCell.Column
    End If
End Function

Public Function GetLastColumn(ByVal oWorksheet As Worksheet, _
                              Optional ByVal lRow As Long) As Long
    Const PROCEDURE_NAME As String = "UtilityFunctions.GetLastColumn"
    Dim oNonEmptyCell As Range
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Worksheet object is required."
    End If
    If lRow > 0 Then
        Set oNonEmptyCell = oWorksheet.Rows(lRow).Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    Else
        Set oNonEmptyCell = oWorksheet.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    End If
    If Not oNonEmptyCell Is Nothing Then
        GetLastColumn = oNonEmptyCell.Column
    End If
End Function

Public Function Fso() As Object
    Static s_oFso As Object
    
    If s_oFso Is Nothing Then
        Set s_oFso = CreateObject("Scripting.FileSystemObject")
    End If
    Set Fso = s_oFso
End Function

Public Function SanitizeFileName(ByVal sName As String) As String
    Const MAX_LENGTH As Long = 255
    Static s_oIllegalCharacters As Object
    Dim sBaseName As String
    Dim sExtension As String
    
    sName = Trim$(sName)
    If s_oIllegalCharacters Is Nothing Then
        Set s_oIllegalCharacters = CreateObject("VBScript.RegExp")
        With s_oIllegalCharacters
            .Global = True
            .IgnoreCase = True
            '"*/:<>?[\]|
            .Pattern = "[\x00-\x1F\x22\x2A\x2F\x3A\x3C\x3E\x3F\x5B-\x5D\x7C\x7F]|[\s.]$|^(CON|PRN|AUX|NUL|COM\d|LPT\d)(\..*)?$"
        End With
    End If
    sName = s_oIllegalCharacters.Replace(sName, "_")
    sBaseName = Fso.GetBaseName(sName)
    sExtension = Fso.GetExtensionName(sName)
    If sExtension <> "" Then
        sExtension = "." & sExtension
    End If
    If Len(sExtension) < MAX_LENGTH Then
        SanitizeFileName = Left$(sBaseName, MAX_LENGTH - Len(sExtension)) & sExtension
    ElseIf Len(sBaseName) > 0 Then
        SanitizeFileName = Left$(sBaseName, 1) & Left$(sExtension, MAX_LENGTH - 1)
    Else
        SanitizeFileName = Left$(sExtension, MAX_LENGTH)
    End If
End Function

Public Function FormatExtension(ByVal lFileFormat As XlFileFormat) As String
    Select Case lFileFormat
        Case xlAddIn: FormatExtension = "xla"
        Case xlAddIn8: FormatExtension = "xla"
        Case xlCSV: FormatExtension = "csv"
        Case xlCSVMac: FormatExtension = "csv"
        Case xlCSVMSDOS: FormatExtension = "csv"
        'Case xlCSVUTF8: FormatExtension = "csv"
        Case xlCSVWindows: FormatExtension = "csv"
        Case xlCurrentPlatformText: FormatExtension = "txt"
        Case xlDBF2: FormatExtension = "dbf"
        Case xlDBF3: FormatExtension = "dbf"
        Case xlDBF4: FormatExtension = "dbf"
        Case xlDIF: FormatExtension = "dif"
        Case xlExcel12: FormatExtension = "xlsb"
        Case xlExcel2: FormatExtension = "xls"
        Case xlExcel2FarEast: FormatExtension = "xls"
        Case xlExcel3: FormatExtension = "xls"
        Case xlExcel4: FormatExtension = "xls"
        Case xlExcel4Workbook: FormatExtension = "xlw"
        Case xlExcel5: FormatExtension = "xls"
        Case xlExcel7: FormatExtension = "xls"
        Case xlExcel8: FormatExtension = "xls"
        Case xlExcel9795: FormatExtension = "xls"
        Case xlHtml: FormatExtension = "html"
        Case xlIntlAddIn: FormatExtension = ""
        Case xlIntlMacro: FormatExtension = ""
        Case xlOpenDocumentSpreadsheet: FormatExtension = "ods"
        Case xlOpenXMLAddIn: FormatExtension = "xlam"
        'Case xlOpenXMLStrictWorkbook: FormatExtension = "xlsx"
        Case xlOpenXMLTemplate: FormatExtension = "xltx"
        Case xlOpenXMLTemplateMacroEnabled: FormatExtension = "xltm"
        Case xlOpenXMLWorkbook: FormatExtension = "xlsx"
        Case xlOpenXMLWorkbookMacroEnabled: FormatExtension = "xlsm"
        Case xlSYLK: FormatExtension = "slk"
        Case xlTemplate: FormatExtension = "xlt"
        Case xlTemplate8: FormatExtension = "xlt"
        Case xlTextMac: FormatExtension = "txt"
        Case xlTextMSDOS: FormatExtension = "txt"
        Case xlTextPrinter: FormatExtension = "prn"
        Case xlTextWindows: FormatExtension = "txt"
        Case xlUnicodeText: FormatExtension = "txt"
        Case xlWebArchive: FormatExtension = "mhtml"
        Case xlWJ2WD1: FormatExtension = "wj2"
        Case xlWJ3: FormatExtension = "wj3"
        Case xlWJ3FJ3: FormatExtension = "wj3"
        Case xlWK1: FormatExtension = "wk1"
        Case xlWK1ALL: FormatExtension = "wk1"
        Case xlWK1FMT: FormatExtension = "wk1"
        Case xlWK3: FormatExtension = "wk3"
        Case xlWK3FM3: FormatExtension = "wk3"
        Case xlWK4: FormatExtension = "wk4"
        Case xlWKS: FormatExtension = "wks"
        Case xlWorkbookDefault: FormatExtension = "xlsx"
        Case xlWorkbookNormal: FormatExtension = "xls"
        Case xlWorks2FarEast: FormatExtension = "wks"
        Case xlWQ1: FormatExtension = "wq1"
        Case xlXMLSpreadsheet: FormatExtension = "xml"
        Case Else: FormatExtension = "xlsx"
    End Select
End Function

Public Function CreateWorkbook(ByVal sDirectory As String, _
                               ByVal sName As String, _
                               Optional ByVal lFileFormat As XlFileFormat = xlWorkbookDefault, _
                               Optional ByVal bOverwrite As Boolean, _
                               Optional ByRef bIsWorkbookNew As Boolean) As Workbook
    Const PROCEDURE_NAME As String = "UtilityFunctions.CreateWorkbook"
    Dim sExtension As String
    Dim sFilename As String
    Dim sPath As String
    Dim oResult As Workbook
    
    If Not Fso.FolderExists(sDirectory) Then
        Err.Raise 9, PROCEDURE_NAME, "Directory does not exist."
    End If
    sExtension = FormatExtension(lFileFormat)
    sFilename = SanitizeFileName(sName & IIf(sExtension = "", "", "." & sExtension))
    sPath = Fso.BuildPath(sDirectory, sFilename)
    If Fso.FileExists(sPath) Then
        If bOverwrite Then
            Kill sPath
        Else
            Set oResult = Workbooks.Open(Filename:=sPath, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, Local:=True)
            bIsWorkbookNew = False
            Set CreateWorkbook = oResult
            Exit Function
        End If
    End If
    Set oResult = Workbooks.Add
    oResult.Title = sName
    oResult.SaveAs Filename:=sPath, FileFormat:=lFileFormat, Local:=True
    bIsWorkbookNew = True
    Set CreateWorkbook = oResult
End Function

Public Function SanitizeWorksheetName(ByVal sName As String) As String
    Const MAX_LENGTH As Long = 31
    Static s_oIllegalCharacters As Object
    
    sName = Trim$(sName)
    If s_oIllegalCharacters Is Nothing Then
        Set s_oIllegalCharacters = CreateObject("VBScript.RegExp")
        With s_oIllegalCharacters
            .Global = True
            ''*/:?[\]
            .Pattern = "[\x00-\x1F\x27\x2A\x2F\x3A\x3F\x5B-\x5D\x7F]"
        End With
    End If
    sName = s_oIllegalCharacters.Replace(sName, "_")
    SanitizeWorksheetName = Left$(sName, MAX_LENGTH)
End Function

Public Function WorksheetExists(ByVal oWorkbook As Workbook, _
                                ByVal sName As String) As Boolean
    Const PROCEDURE_NAME As String = "UtilityFunctions.WorksheetExists"
    
    If oWorkbook Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Workbook object is required."
    End If
    On Error Resume Next
    WorksheetExists = Not oWorkbook.Worksheets(sName) Is Nothing
End Function

Public Function CreateWorksheet(ByVal oWorkbook As Workbook, _
                                ByVal sName As String, _
                                Optional ByVal bOverwrite As Boolean, _
                                Optional ByRef bIsWorksheetNew As Boolean) As Worksheet
    Const PROCEDURE_NAME As String = "UtilityFunctions.CreateWorksheet"
    Dim bDisplayAlertsSave As Boolean
    Dim oResult As Worksheet
    
    If oWorkbook Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Workbook object is required."
    End If
    sName = SanitizeWorksheetName(sName)
    If WorksheetExists(oWorkbook, sName) Then
        If bOverwrite Then
            Set oResult = oWorkbook.Worksheets.Add(After:=oWorkbook.Worksheets(sName))
            bDisplayAlertsSave = Application.DisplayAlerts
            Application.DisplayAlerts = False
            oWorkbook.Worksheets(sName).Delete
            Application.DisplayAlerts = bDisplayAlertsSave
            oResult.Name = sName
            bIsWorksheetNew = True
        Else
            Set oResult = oWorkbook.Worksheets(sName)
            bIsWorksheetNew = False
        End If
    Else
        Set oResult = oWorkbook.Worksheets.Add(After:=oWorkbook.Worksheets(oWorkbook.Worksheets.Count))
        oResult.Name = sName
        bIsWorksheetNew = True
    End If
    oResult.Activate
    ActiveWindow.Zoom = 80
    Set CreateWorksheet = oResult
End Function

Public Function GetUtcTime() As Date
    Dim uNow As SYSTEMTIME
    
    GetSystemTime uNow
    With uNow
        GetUtcTime = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Public Function GetUtcTimestamp() As String
    Dim uNow As SYSTEMTIME
    
    GetSystemTime uNow
    With uNow
        GetUtcTimestamp = Format$(DateDiff("s", DateSerial(1970, 1, 1), DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)), "0000000000") & Format$(.wMillisecond, "000")
    End With
End Function

Public Function ConvertDateToTimestamp(ByVal dtmDate As Date) As Long
    ConvertDateToTimestamp = DateDiff("s", DateSerial(1970, 1, 1), dtmDate)
End Function

Public Function ConvertTimestampToDate(ByVal lTimestamp As Long) As Date
    ConvertTimestampToDate = DateAdd("s", lTimestamp, DateSerial(1970, 1, 1))
End Function

Public Sub LogToWorksheet(ByVal sName As String, _
                          ParamArray avArgs() As Variant)
    Dim i As Long
    Dim lLastRow As Long
    Dim vHeader As Variant
    Dim vHeaderCol As Variant
    Dim bIsWorksheetNew As Boolean
    Dim oLog As Worksheet
    
    Set oLog = CreateWorksheet(ThisWorkbook, sName, False, bIsWorksheetNew)
    lLastRow = MaxLng2(1, GetLastRow(oLog))
    With oLog
        On Error Resume Next
        For i = 0 To UBound(avArgs) - 1 Step 2
            vHeader = avArgs(i)
            vHeaderCol = Application.Match(vHeader, .Rows(1), 0)
            If IsError(vHeaderCol) Then
                vHeaderCol = GetLastColumn(oLog) + 1
                .Cells(1, vHeaderCol) = vHeader
            End If
            .Cells(lLastRow + 1, vHeaderCol) = avArgs(i + 1)
            Application.GoTo .Cells(lLastRow + 1, vHeaderCol)
            DoEvents
        Next i
        On Error GoTo 0
        If bIsWorksheetNew Then
            .Activate
            .Cells(2, 1).Select
            ActiveWindow.FreezePanes = True
        End If
    End With
End Sub
