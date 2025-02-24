Attribute VB_Name = "UtilityFunctions"
'---------------------------------------------------------------------------------------
' Module    : UtilityFunctions
' Author    : personalityson
' Date      : 01.09.2024
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public Const MATH_PI As Double = 3.14159265358979
Public Const MATH_2PI As Double = 6.28318530717959
Public Const MATH_PI2 As Double = 1.5707963267949
Public Const MATH_E As Double = 2.71828182845905

Public Const DOUBLE_MIN_ABS As Double = 4.94065645841247E-324
Public Const DOUBLE_MAX_ABS As Double = 1.79769313486231E+308
Public Const DOUBLE_MIN_LOG As Double = -744.440071921381
Public Const DOUBLE_MAX_LOG As Double = 709.782712893384

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
                                                                                ByRef Source As Any, _
                                                                                ByVal Length As LongPtr)

Public Declare PtrSafe Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, _
                                                                                ByVal Length As LongPtr)

Public Declare PtrSafe Function VarPtrArray Lib "VBE7.dll" Alias "VarPtr" (ByRef Var() As Any) As LongPtr

Public Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long

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

Public Function GetRank(ByVal vArray As Variant) As Integer
    Const VARIANT_OFFSET_parray As Long = 8
    Dim iVarType As Integer
    Dim pSafeArray As LongPtr

    CopyMemory iVarType, vArray, SIZEOF_INTEGER
    If (iVarType And vbArray) = 0 Then
        GetRank = -1
        Exit Function
    End If
    CopyMemory pSafeArray, ByVal VarPtr(vArray) + VARIANT_OFFSET_parray, SIZEOF_LONGPTR
    If pSafeArray = NULL_PTR Then
        GetRank = 0
        Exit Function
    End If
    CopyMemory GetRank, ByVal pSafeArray, SIZEOF_INTEGER
End Function

Public Sub ParseVariantToLongArray(ByVal vValueOrArray As Variant, _
                                   ByRef lNumElements As Long, _
                                   ByRef alArray() As Long)
    Const PROCEDURE_NAME As String = "Tensor.ParseVariantToLongArray"
    Dim lRank As Long
    Dim lLBound As Long
    Dim lUBound As Long
    Dim i As Long
    
    lRank = GetRank(vValueOrArray)
    Select Case lRank
        Case -1
            lNumElements = 1
            ReDim alArray(1 To lNumElements)
            alArray(1) = CLng(vValueOrArray)
        Case 0
            lNumElements = 0
            Erase alArray
        Case 1
            lLBound = LBound(vValueOrArray)
            lUBound = UBound(vValueOrArray)
            If lLBound > lUBound Then
                lNumElements = 0
                Erase alArray
            Else
                lNumElements = lUBound - lLBound + 1
                ReDim alArray(1 To lNumElements)
                For i = 1 To lNumElements
                    alArray(i) = CLng(vValueOrArray(lLBound + i - 1))
                Next i
            End If
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Expected a single value, an uninitialized array, or a one-dimensional array."
    End Select
End Sub

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

Public Function FileFormatToExtension(ByVal lFileFormat As XlFileFormat) As String
    Select Case lFileFormat
        Case xlAddIn: FileFormatToExtension = "xla"
        Case xlAddIn8: FileFormatToExtension = "xla"
        Case xlCSV: FileFormatToExtension = "csv"
        Case xlCSVMac: FileFormatToExtension = "csv"
        Case xlCSVMSDOS: FileFormatToExtension = "csv"
        Case xlCSVUTF8: FileFormatToExtension = "csv"
        Case xlCSVWindows: FileFormatToExtension = "csv"
        Case xlCurrentPlatformText: FileFormatToExtension = "txt"
        Case xlDBF2: FileFormatToExtension = "dbf"
        Case xlDBF3: FileFormatToExtension = "dbf"
        Case xlDBF4: FileFormatToExtension = "dbf"
        Case xlDIF: FileFormatToExtension = "dif"
        Case xlExcel12: FileFormatToExtension = "xlsb"
        Case xlExcel2: FileFormatToExtension = "xls"
        Case xlExcel2FarEast: FileFormatToExtension = "xls"
        Case xlExcel3: FileFormatToExtension = "xls"
        Case xlExcel4: FileFormatToExtension = "xls"
        Case xlExcel4Workbook: FileFormatToExtension = "xlw"
        Case xlExcel5: FileFormatToExtension = "xls"
        Case xlExcel7: FileFormatToExtension = "xls"
        Case xlExcel8: FileFormatToExtension = "xls"
        Case xlExcel9795: FileFormatToExtension = "xls"
        Case xlHtml: FileFormatToExtension = "html"
        Case xlIntlAddIn: FileFormatToExtension = ""
        Case xlIntlMacro: FileFormatToExtension = ""
        Case xlOpenDocumentSpreadsheet: FileFormatToExtension = "ods"
        Case xlOpenXMLAddIn: FileFormatToExtension = "xlam"
        Case xlOpenXMLStrictWorkbook: FileFormatToExtension = "xlsx"
        Case xlOpenXMLTemplate: FileFormatToExtension = "xltx"
        Case xlOpenXMLTemplateMacroEnabled: FileFormatToExtension = "xltm"
        Case xlOpenXMLWorkbook: FileFormatToExtension = "xlsx"
        Case xlOpenXMLWorkbookMacroEnabled: FileFormatToExtension = "xlsm"
        Case xlSYLK: FileFormatToExtension = "slk"
        Case xlTemplate: FileFormatToExtension = "xlt"
        Case xlTemplate8: FileFormatToExtension = "xlt"
        Case xlTextMac: FileFormatToExtension = "txt"
        Case xlTextMSDOS: FileFormatToExtension = "txt"
        Case xlTextPrinter: FileFormatToExtension = "prn"
        Case xlTextWindows: FileFormatToExtension = "txt"
        Case xlUnicodeText: FileFormatToExtension = "txt"
        Case xlWebArchive: FileFormatToExtension = "mhtml"
        Case xlWJ2WD1: FileFormatToExtension = "wj2"
        Case xlWJ3: FileFormatToExtension = "wj3"
        Case xlWJ3FJ3: FileFormatToExtension = "wj3"
        Case xlWK1: FileFormatToExtension = "wk1"
        Case xlWK1ALL: FileFormatToExtension = "wk1"
        Case xlWK1FMT: FileFormatToExtension = "wk1"
        Case xlWK3: FileFormatToExtension = "wk3"
        Case xlWK3FM3: FileFormatToExtension = "wk3"
        Case xlWK4: FileFormatToExtension = "wk4"
        Case xlWKS: FileFormatToExtension = "wks"
        Case xlWorkbookDefault: FileFormatToExtension = "xlsx"
        Case xlWorkbookNormal: FileFormatToExtension = "xls"
        Case xlWorks2FarEast: FileFormatToExtension = "wks"
        Case xlWQ1: FileFormatToExtension = "wq1"
        Case xlXMLSpreadsheet: FileFormatToExtension = "xml"
        Case Else: FileFormatToExtension = "xlsx"
    End Select
End Function

Public Function CreateWorkbook(ByVal sDirectory As String, _
                               ByVal sName As String, _
                               Optional ByVal lFileFormat As XlFileFormat = xlWorkbookDefault, _
                               Optional ByVal bOverwrite As Boolean, _
                               Optional ByRef bIsWorkbookNew As Boolean) As Workbook
    Dim sExtension As String
    Dim sFileName As String
    Dim sPath As String
    Dim oResult As Workbook
    
    sExtension = FileFormatToExtension(lFileFormat)
    sFileName = SanitizeFileName(sName & IIf(sExtension = "", "", "." & sExtension))
    sPath = Fso.BuildPath(sDirectory, sFileName)
    If Fso.FileExists(sPath) Then
        If bOverwrite Then
            Kill sPath
        Else
            bIsWorkbookNew = False
            Set CreateWorkbook = Workbooks.Open(FileName:=sPath, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, Local:=True)
            Exit Function
        End If
    End If
    Set oResult = Workbooks.Add
    oResult.Title = sName
    oResult.SaveAs FileName:=sPath, FileFormat:=lFileFormat, Local:=True
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

Public Function WorksheetExists(ByVal oParent As Workbook, _
                                ByVal sName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not oParent.Worksheets(sName) Is Nothing
End Function

Public Function CreateWorksheet(ByVal oParent As Workbook, _
                                ByVal sName As String, _
                                Optional ByVal bOverwrite As Boolean, _
                                Optional ByRef bIsWorksheetNew As Boolean) As Worksheet
    Dim bDisplayAlertsSave As Boolean
    Dim oResult As Worksheet
    
    bDisplayAlertsSave = Application.DisplayAlerts
    sName = SanitizeWorksheetName(sName)
    If WorksheetExists(oParent, sName) Then
        If bOverwrite Then
            Set oResult = oParent.Worksheets.Add(after:=Parent.Worksheets(sName))
            Application.DisplayAlerts = False
            oParent.Worksheets(sName).Delete
            Application.DisplayAlerts = bDisplayAlertsSave
        Else
            bIsWorksheetNew = False
            Set CreateWorksheet = oParent.Worksheets(sName)
            Exit Function
        End If
    Else
        Set oResult = oParent.Worksheets.Add(after:=oParent.Worksheets(oParent.Worksheets.Count))
    End If
    oResult.Name = sName
    oResult.Activate
    ActiveWindow.Zoom = 80
    bIsWorksheetNew = True
    Set CreateWorksheet = oResult
End Function

Public Function DumpWorksheet(ByVal oWorksheet As Worksheet, _
                              ByVal sDirectory As String, _
                              ByVal sName As String, _
                              Optional ByVal lFileFormat As XlFileFormat = xlWorkbookDefault, _
                              Optional ByVal bOverwrite As Boolean) As Workbook
    Dim bDisplayAlertsSave As Boolean
    Dim i As Long
    Dim oResult As Workbook
    
    Set oResult = CreateWorkbook(sDirectory, sName, lFileFormat, bOverwrite)
    oWorksheet.Copy Before:=oResult.Sheets(1)
    bDisplayAlertsSave = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For i = oResult.Worksheets.Count To 2 Step -1
        oResult.Worksheets(i).Delete
    Next i
    Application.DisplayAlerts = bDisplayAlertsSave
    oResult.Sheets(1).Name = oWorksheet.Name
    oResult.Save
    'oResult.Close SaveChanges:=True
    Set DumpWorksheet = oResult
End Function

Public Function GetUtcTime() As Date
    Dim uNow As SYSTEMTIME
    
    GetSystemTime uNow
    With uNow
        GetUtcTime = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Public Function GetUtcTimestamp() As Long
    Dim uNow As SYSTEMTIME
    
    GetSystemTime uNow
    With uNow
        GetUtcTimestamp = DateDiff("s", DateSerial(1970, 1, 1), DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond))
    End With
End Function

Public Function ConvertDateToTimestamp(ByVal dtmDate As Date) As Long
    ConvertDateToTimestamp = DateDiff("s", DateSerial(1970, 1, 1), dtmDate)
End Function

Public Function ConvertTimestampToDate(ByVal lTimestamp As Long) As Date
    ConvertTimestampToDate = DateAdd("s", lTimestamp, DateSerial(1970, 1, 1))
End Function

Public Sub LogError(ByVal sSource As String, _
                    ByVal lErrorNumber As Long, _
                    ByVal sErrorDescription As String)
    Dim oErrors As Worksheet
    Dim bIsWorksheetNew As Boolean
    Dim lLastRow As Long
    
    Set oErrors = CreateWorksheet(ThisWorkbook, "Errors", False, bIsWorksheetNew)
    With oErrors
        If bIsWorksheetNew Then
            .Cells(1, 1) = "Time"
            .Cells(1, 2) = "Source"
            .Cells(1, 3) = "Error Number"
            .Cells(1, 4) = "Error Description"
            lLastRow = 1
        Else
            lLastRow = GetLastRow(oErrors)
        End If
        .Cells(lLastRow + 1, 1) = GetUtcTime()
        .Cells(lLastRow + 1, 2) = sSource
        .Cells(lLastRow + 1, 3) = lErrorNumber
        .Cells(lLastRow + 1, 4) = sErrorDescription
        .Cells(lLastRow + 1, 1).Resize(1, 4).WrapText = False
        Application.GoTo .Cells(lLastRow + 1, 1)
    End With
End Sub

Public Sub LogToWorksheet(ByVal sName As String, _
                          ParamArray avArgs() As Variant)
    Dim oLog As Worksheet
    Dim bIsWorksheetNew As Boolean
    Dim lLastRow As Long
    Dim i As Long
    Dim vHeader As Variant
    Dim vHeaderCol As Variant
    Dim vValue As Variant
    
    Set oLog = CreateWorksheet(ThisWorkbook, sName, False, bIsWorksheetNew)
    If bIsWorksheetNew Then
        lLastRow = 1
    Else
        lLastRow = GetLastRow(oLog)
    End If
    With oLog
        On Error Resume Next
        For i = 0 To UBound(avArgs) - 1 Step 2
            vHeader = avArgs(i)
            vHeaderCol = Application.Match(vHeader, .Rows(1), 0)
            vValue = avArgs(i + 1)
            If IsError(vHeaderCol) Then
                vHeaderCol = GetLastColumn(oLog) + 1
                .Cells(1, vHeaderCol) = vHeader
            End If
            .Cells(lLastRow + 1, vHeaderCol) = vValue
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
