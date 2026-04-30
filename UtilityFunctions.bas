Attribute VB_Name = "UtilityFunctions"
'---------------------------------------------------------------------------------------
' Module    : UtilityFunctions
' Author    : personalityson
' Date      : 01.05.2026
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
    rndNearest
    rndDown
    rndUp
    rndTowardsZero
    rndTowardsInfinity
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
                                                                                ByRef source As Any, _
                                                                                ByVal Length As LongPtr)

Public Declare PtrSafe Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, _
                                                                                ByVal Length As LongPtr)

Public Declare PtrSafe Function VarPtrArray Lib "VBE7.dll" Alias "VarPtr" (ByRef Var() As Any) As LongPtr

Public Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long

Public Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Sub GetSystemTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME)

Public Function MinLng(ByVal A As Long, _
                       ByVal B As Long) As Long
    If A < B Then
        MinLng = A
    Else
        MinLng = B
    End If
End Function

Public Function MaxLng(ByVal A As Long, _
                       ByVal B As Long) As Long
    If A > B Then
        MaxLng = A
    Else
        MaxLng = B
    End If
End Function

Public Function MinPtr(ByVal A As LongPtr, _
                       ByVal B As LongPtr) As LongPtr
    If A < B Then
        MinPtr = A
    Else
        MinPtr = B
    End If
End Function

Public Function MaxPtr(ByVal A As LongPtr, _
                       ByVal B As LongPtr) As LongPtr
    If A > B Then
        MaxPtr = A
    Else
        MaxPtr = B
    End If
End Function

Public Function MinDbl(ByVal A As Double, _
                       ByVal B As Double) As Double
    If A < B Then
        MinDbl = A
    Else
        MinDbl = B
    End If
End Function

Public Function MaxDbl(ByVal A As Double, _
                       ByVal B As Double) As Double
    If A > B Then
        MaxDbl = A
    Else
        MaxDbl = B
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
        Case rndNearest
            RoundToMultiple = Round(dblValue / dblMultiple) * dblMultiple
        Case rndDown
            RoundToMultiple = Int(dblValue / dblMultiple) * dblMultiple
        Case rndUp
            RoundToMultiple = -Int(-dblValue / dblMultiple) * dblMultiple
        Case rndTowardsZero
            RoundToMultiple = Sgn(dblValue) * Int(Abs(dblValue) / dblMultiple) * dblMultiple
        Case rndTowardsInfinity
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

Private Function GetSafeArrayPtr(ByRef vArray As Variant, _
                                 ByRef pSafeArray As LongPtr) As Boolean
    Const VARIANT_OFFSET_parray As Long = 8
    Const VT_BYREF As Long = &H4000
    Dim iVarType As Integer
    
    pSafeArray = NULL_PTR
    CopyMemory iVarType, vArray, SIZEOF_INTEGER
    If (iVarType And vbArray) = 0 Then
        Exit Function
    End If
    CopyMemory pSafeArray, ByVal VarPtr(vArray) + VARIANT_OFFSET_parray, SIZEOF_LONGPTR
    If (iVarType And VT_BYREF) <> 0 Then
        If pSafeArray <> NULL_PTR Then 'To be safe, should not happen
            CopyMemory pSafeArray, ByVal pSafeArray, SIZEOF_LONGPTR
        End If
    End If
    GetSafeArrayPtr = True
End Function

Public Function GetRank(ByVal vArray As Variant) As Integer
    Dim pSafeArray As LongPtr
    
    Select Case True
        Case Not GetSafeArrayPtr(vArray, pSafeArray)
            GetRank = -1 'Scalar
        Case pSafeArray = NULL_PTR
            GetRank = 0 'Uninitialized
        Case Else
            CopyMemory GetRank, ByVal pSafeArray, SIZEOF_INTEGER
    End Select
End Function

Public Function GetSize(ByVal vArray As Variant, _
                        Optional ByVal iDimension As Integer = 1) As Long
    Const PROCEDURE_NAME As String = "UtilityFunctions.GetSize"
    #If Win64 Then
        Const SA_OFFSET_rgsabound As Long = 24
    #Else
        Const SA_OFFSET_rgsabound As Long = 16
    #End If
    Dim pSafeArray As LongPtr
    Dim iNumDimensions As Integer
    
    If Not GetSafeArrayPtr(vArray, pSafeArray) Then
        Err.Raise 5, PROCEDURE_NAME, "Argument is not an array."
    End If
    If pSafeArray = NULL_PTR Then
        Err.Raise 9, PROCEDURE_NAME, "Array is uninitialized."
    End If
    CopyMemory iNumDimensions, ByVal pSafeArray, SIZEOF_INTEGER
    If iDimension < 1 Or iDimension > iNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    CopyMemory GetSize, _
               ByVal pSafeArray + SA_OFFSET_rgsabound + (iNumDimensions - iDimension) * 2 * SIZEOF_LONG, _
               SIZEOF_LONG
End Function

Public Sub SetLBound(ByRef vArray As Variant, _
                     ByVal iDimension As Integer, _
                     ByVal lLBound As Long)
    Const PROCEDURE_NAME As String = "UtilityFunctions.SetLBound"
    #If Win64 Then
        Const SA_OFFSET_rgsabound As Long = 24
    #Else
        Const SA_OFFSET_rgsabound As Long = 16
    #End If
    Dim pSafeArray As LongPtr
    Dim iNumDimensions As Integer

    If Not GetSafeArrayPtr(vArray, pSafeArray) Then
        Err.Raise 5, PROCEDURE_NAME, "Argument is not an array."
    End If
    If pSafeArray = NULL_PTR Then
        Err.Raise 9, PROCEDURE_NAME, "Array is uninitialized."
    End If
    CopyMemory iNumDimensions, ByVal pSafeArray, SIZEOF_INTEGER
    If iDimension < 1 Or iDimension > iNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    CopyMemory ByVal pSafeArray + SA_OFFSET_rgsabound + (iNumDimensions - iDimension) * 2 * SIZEOF_LONG + SIZEOF_LONG, _
               lLBound, _
               SIZEOF_LONG
End Sub

Public Function EnsureArray(ByVal vValueOrArray As Variant, _
                            Optional ByVal vLBound As Variant) As Variant
    Const PROCEDURE_NAME As String = "UtilityFunctions.EnsureArray"
    Select Case GetRank(vValueOrArray)
        Case -1
            EnsureArray = Array(vValueOrArray)
        Case 0
            EnsureArray = Array()
        Case 1
            EnsureArray = vValueOrArray
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Expecting a scalar, an uninitialized array, or a one-dimensional array."
    End Select
    If Not IsMissing(vLBound) Then
        SetLBound EnsureArray, 1, CLng(vLBound)
    End If
End Function

Public Sub ParseVariantToLongArray(ByVal vArray As Variant, _
                                   ByRef lNumElements As Long, _
                                   ByRef alArray() As Long)
    Dim i As Long
    Dim lLBound As Long
    Dim lUBound As Long

    vArray = EnsureArray(vArray)
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

Public Sub ParseVariantToDoubleArray(ByVal vArray As Variant, _
                                     ByRef lNumElements As Long, _
                                     ByRef adblArray() As Double)
    Dim i As Long
    Dim lLBound As Long
    Dim lUBound As Long

    vArray = EnsureArray(vArray)
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

Public Function GetFirstRow(ByVal oWorksheet As Worksheet, _
                            Optional ByVal lColumn As Long) As Long
    Const PROCEDURE_NAME As String = "UtilityFunctions.GetFirstRow"
    Dim oNonEmptyCell As Range
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Worksheet object is required."
    End If
    If lColumn > 0 Then
        Set oNonEmptyCell = oWorksheet.Columns(lColumn).Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    Else
        Set oNonEmptyCell = oWorksheet.Cells.Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
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
        Set oNonEmptyCell = oWorksheet.Columns(lColumn).Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    Else
        Set oNonEmptyCell = oWorksheet.Cells.Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
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
        Set oNonEmptyCell = oWorksheet.Rows(lRow).Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    Else
        Set oNonEmptyCell = oWorksheet.Cells.Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
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
        Set oNonEmptyCell = oWorksheet.Rows(lRow).Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    Else
        Set oNonEmptyCell = oWorksheet.Cells.Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
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
    Static s_oIllegalChars As Object
    Static s_oReservedNames As Object
    Dim sBaseName As String
    Dim sExtension As String
    Dim lDotPosition As Long
    
    sName = Trim$(sName)
    If s_oIllegalChars Is Nothing Then
        Set s_oIllegalChars = CreateObject("VBScript.RegExp")
        With s_oIllegalChars
            .Global = True
            .IgnoreCase = False
            .Pattern = "[\x00-\x1F\x22\x2A\x2F\x3A\x3C\x3E\x3F\x5B-\x5D\x7C\x7F]|[\s.]+$"
        End With
    End If
    If s_oReservedNames Is Nothing Then
        Set s_oReservedNames = CreateObject("VBScript.RegExp")
        With s_oReservedNames
            .Global = False
            .IgnoreCase = True
            .Pattern = "^(CON|PRN|AUX|NUL|COM\d|LPT\d)$"
        End With
    End If
    sName = s_oIllegalChars.Replace(sName, "_")
    lDotPosition = InStrRev(sName, ".")
    If lDotPosition > 1 Then
        sBaseName = Left$(sName, lDotPosition - 1)
        sExtension = Mid$(sName, lDotPosition)
    Else
        sBaseName = sName
        sExtension = ""
    End If
    If s_oReservedNames.Test(sBaseName) Then
        sBaseName = "_" & sBaseName
    End If
    If Len(sExtension) < MAX_LENGTH Then
        SanitizeFileName = Left$(sBaseName, MAX_LENGTH - Len(sExtension)) & sExtension
    ElseIf Len(sBaseName) > 0 Then
        SanitizeFileName = Left$(sBaseName, 1) & Left$(sExtension, MAX_LENGTH - 1)
    Else
        SanitizeFileName = Left$(sExtension, MAX_LENGTH)
    End If
    If Len(SanitizeFileName) = 0 Or SanitizeFileName = "." Then
        SanitizeFileName = "_"
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
                               Optional ByRef bWorkbookIsNew As Boolean) As Workbook
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
            bWorkbookIsNew = False
            Set CreateWorkbook = oResult
            Exit Function
        End If
    End If
    Set oResult = Workbooks.Add
    oResult.Title = sName
    oResult.SaveAs Filename:=sPath, FileFormat:=lFileFormat, Local:=True
    bWorkbookIsNew = True
    Set CreateWorkbook = oResult
End Function

Public Function SanitizeWorksheetName(ByVal sName As String) As String
    Const MAX_LENGTH As Long = 31
    Static s_oIllegalChars As Object
    
    sName = Trim$(sName)
    If s_oIllegalChars Is Nothing Then
        Set s_oIllegalChars = CreateObject("VBScript.RegExp")
        With s_oIllegalChars
            .Global = True
            .IgnoreCase = False
            .Pattern = "[\x00-\x1F\x2A\x2F\x3A\x3F\x5B-\x5D\x7F]"
        End With
    End If
    sName = s_oIllegalChars.Replace(sName, "_")
    sName = Left$(sName, MAX_LENGTH)
    If Len(sName) > 0 And Left$(sName, 1) = "'" Then
        sName = "_" & Mid$(sName, 2)
    End If
    If Len(sName) > 0 And Right$(sName, 1) = "'" Then
        sName = Left$(sName, Len(sName) - 1) & "_"
    End If
    If StrComp(sName, "History", vbTextCompare) = 0 Then
        sName = "History_"
    End If
    If Len(sName) = 0 Then
        sName = "_"
    End If
    SanitizeWorksheetName = sName
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
                                Optional ByRef bWorksheetIsNew As Boolean) As Worksheet
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
            bWorksheetIsNew = True
        Else
            Set oResult = oWorkbook.Worksheets(sName)
            bWorksheetIsNew = False
        End If
    Else
        Set oResult = oWorkbook.Worksheets.Add(After:=oWorkbook.Worksheets(oWorkbook.Worksheets.Count))
        oResult.Name = sName
        bWorksheetIsNew = True
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

Public Sub WriteLog(ByVal sName As String, _
                    ParamArray avArgs() As Variant)
    Dim i As Long
    Dim lLastRow As Long
    Dim vHeader As Variant
    Dim vHeaderCol As Variant
    Dim bWorksheetIsNew As Boolean
    Dim oLog As Worksheet
    
    Set oLog = CreateWorksheet(ThisWorkbook, sName, False, bWorksheetIsNew)
    lLastRow = MaxLng(1, GetLastRow(oLog))
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
            'Application.GoTo .Cells(lLastRow + 1, vHeaderCol)
            DoEvents
        Next i
        On Error GoTo 0
        If bWorksheetIsNew Then
            .Activate
            .Cells(2, 1).Select
            ActiveWindow.FreezePanes = True
        End If
    End With
End Sub

