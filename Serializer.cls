VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Serializer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SERIALIZABLE_TYPE_MEMO As String = "Memo"
Private Const SERIALIZABLE_TYPE_NOTHING As String = "Nothing"

'Dependencies
Private m_oWorksheet As Worksheet

'State
Private m_oSerializedObjects As Collection
Private m_lCursor As Long
Private m_lIndent As Long

Public Property Get Worksheet() As Worksheet
    Set Worksheet = m_oWorksheet
End Property

Public Sub Init(ByVal sName As String, _
                ByVal bOverwrite As Boolean)
    Clear
    Set m_oWorksheet = CreateWorksheet(ThisWorkbook, sName, bOverwrite)
End Sub

Public Sub WriteObject(ByVal oObject As ISerializable)
    Dim lMemoIndex As Long
    
    If oObject Is Nothing Then
        WriteNothing
    Else
        lMemoIndex = IndexOf(oObject)
        If lMemoIndex > 0 Then
            WriteMemo lMemoIndex
        Else
            WriteSerializable oObject
        End If
    End If
End Sub

Public Sub WriteLong(ByVal lValue As Long)
    GetCell() = lValue
End Sub

Public Sub WriteLongArray(ByRef alArray() As Long)
    WriteArray alArray
End Sub

Public Sub WriteDouble(ByVal dblValue As Double)
    GetCell() = dblValue
End Sub

Public Sub WriteDoubleArray(ByRef adblArray() As Double)
    WriteArray adblArray
End Sub

Public Sub WriteString(ByVal sValue As String)
    With GetCell()
        .NumberFormat = "@"
        .value = sValue
    End With
End Sub

Public Sub WriteStringArray(ByRef asArray() As String)
    WriteArray asArray
End Sub

Public Sub WriteBoolean(ByVal bValue As Boolean)
    GetCell() = bValue
End Sub

Public Sub WriteBooleanArray(ByRef abArray() As Boolean)
    WriteArray abArray
End Sub

Public Function ReadObject() As ISerializable
    Dim sTypeName As String
    
    sTypeName = ReadString()
    Select Case sTypeName
        Case SERIALIZABLE_TYPE_NOTHING
            Set ReadObject = ReadNothing()
        Case SERIALIZABLE_TYPE_MEMO
            Set ReadObject = ReadMemo()
        Case Else
            Set ReadObject = ReadSerializable(sTypeName)
    End Select
End Function

Public Function ReadLong() As Long
    ReadLong = GetCell()
End Function

Public Function ReadLongArray() As Long()
    ReadArray ReadLongArray
End Function

Public Function ReadDouble() As Double
    ReadDouble = GetCell()
End Function

Public Function ReadDoubleArray() As Double()
    ReadArray ReadDoubleArray
End Function

Public Function ReadString() As String
    ReadString = GetCell()
End Function

Public Function ReadStringArray() As String()
    ReadArray ReadStringArray
End Function

Public Function ReadBoolean() As Boolean
    ReadBoolean = GetCell()
End Function

Public Function ReadBooleanArray() As Boolean()
    ReadArray ReadBooleanArray
End Function

Private Sub Memoize(ByVal oObject As ISerializable)
    m_oSerializedObjects.Add oObject
End Sub

Private Function IndexOf(ByVal oObject As ISerializable) As Long
    Dim i As Long
    
    For i = 1 To m_oSerializedObjects.Count
        If m_oSerializedObjects(i) Is oObject Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

Private Function GetCell() As Range
    Set GetCell = m_oWorksheet.Cells(m_lCursor, m_lIndent)
    m_lCursor = m_lCursor + 1
End Function

Private Function GetRange(ByVal lNumRows As Long, _
                          ByVal lNumCols As Long) As Range
    Const PROCEDURE_NAME As String = "Serializer.GetRange"

    If lNumRows < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of rows must be >= 1."
    End If
    If lNumCols < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of columns must be >= 1."
    End If
    Set GetRange = m_oWorksheet.Cells(m_lCursor, m_lIndent).Resize(lNumRows, lNumCols)
    m_lCursor = m_lCursor + lNumRows
End Function

Private Sub WriteNothing()
    WriteString SERIALIZABLE_TYPE_NOTHING
End Sub

Private Sub WriteMemo(ByVal lMemoIndex As Long)
    WriteString SERIALIZABLE_TYPE_MEMO
    m_lIndent = m_lIndent + 1
    WriteLong lMemoIndex
    m_lIndent = m_lIndent - 1
End Sub

Private Sub WriteSerializable(ByVal oObject As ISerializable)
    WriteString TypeName(oObject)
    Memoize oObject
    m_lIndent = m_lIndent + 1
    oObject.Save Me
    m_lIndent = m_lIndent - 1
End Sub

Private Sub WriteArray(ByVal vArray As Variant)
    Const PROCEDURE_NAME As String = "Serializer.WriteArray"
    Dim lNumDimensions As Long
    Dim lNumRows As Long
    Dim lNumCols As Long
    
    lNumDimensions = GetRank(vArray)
    Select Case lNumDimensions
        Case 0
            WriteLong 0 'Number of dimensions
            Exit Sub
        Case 1
            lNumRows = 1
            lNumCols = UBound(vArray, 1) - LBound(vArray, 1) + 1
            WriteLong 1 'Number of dimensions
            WriteLong lNumCols
        Case 2
            lNumRows = UBound(vArray, 1) - LBound(vArray, 1) + 1
            lNumCols = UBound(vArray, 2) - LBound(vArray, 2) + 1
            WriteLong 2 'Number of dimensions
            WriteLong lNumRows
            WriteLong lNumCols
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Expected an uninitialized array, a one-dimensional array, or a two-dimensional array."
    End Select
    With GetRange(lNumRows, lNumCols)
        If (VarType(vArray) And Not vbArray) = vbString Then
            .NumberFormat = "@"
        End If
        .value = vArray
    End With
End Sub

Private Function ReadNothing() As ISerializable
    'Do nothing
End Function

Private Function ReadMemo() As ISerializable
    Dim lMemoIndex As Long
    
    m_lIndent = m_lIndent + 1
    lMemoIndex = ReadLong()
    m_lIndent = m_lIndent - 1
    Set ReadMemo = m_oSerializedObjects(lMemoIndex)
End Function

Private Function ReadSerializable(ByVal sTypeName As String) As ISerializable
    Set ReadSerializable = CreateInstance(sTypeName)
    Memoize ReadSerializable
    m_lIndent = m_lIndent + 1
    ReadSerializable.Load Me
    m_lIndent = m_lIndent - 1
End Function

Private Function CreateInstance(ByVal sTypeName As String) As ISerializable
    Const PROCEDURE_NAME As String = "Serializer.CreateInstance"
    
    Select Case sTypeName
        Case "AdamW"
            Set CreateInstance = New AdamW
        Case "BackpropTrainer"
            Set CreateInstance = New BackpropTrainer
        Case "BCELoss"
            Set CreateInstance = New BCELoss
        Case "CCELoss"
            Set CreateInstance = New CCELoss
        Case "DropoutLayer"
            Set CreateInstance = New DropoutLayer
        Case "FullyConnectedLayer"
            Set CreateInstance = New FullyConnectedLayer
        Case "InputNormalizationLayer"
            Set CreateInstance = New InputNormalizationLayer
        Case "L1Loss"
            Set CreateInstance = New L1Loss
        Case "L2Loss"
            Set CreateInstance = New L2Loss
        Case "LeakyReLULayer"
            Set CreateInstance = New LeakyReLULayer
        Case "Parameter"
            Set CreateInstance = New Parameter
        Case "Sequential"
            Set CreateInstance = New Sequential
        Case "SGDW"
            Set CreateInstance = New SGDW
        Case "SigmoidLayer"
            Set CreateInstance = New SigmoidLayer
        Case "SoftmaxLayer"
            Set CreateInstance = New SoftmaxLayer
        Case "Storage"
            Set CreateInstance = New Storage
        Case "TanhLayer"
            Set CreateInstance = New TanhLayer
        Case "Tensor"
            Set CreateInstance = New Tensor
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Unsupported serializable type."
    End Select
End Function

Private Sub ReadArray(ByRef vArray As Variant)
    Const PROCEDURE_NAME As String = "Serializer.ReadArray"
    Dim i As Long
    Dim j As Long
    Dim lNumDimensions As Long
    Dim lNumRows As Long
    Dim lNumCols As Long
    
    lNumDimensions = ReadLong()
    Select Case lNumDimensions
        Case 0
            'Do nothing
        Case 1
            lNumRows = 1
            lNumCols = ReadLong()
            ReDim vArray(1 To lNumCols)
            With GetRange(lNumRows, lNumCols)
                For i = 1 To lNumCols
                    vArray(i) = .Cells(1, i)
                Next i
            End With
        Case 2
            lNumRows = ReadLong()
            lNumCols = ReadLong()
            ReDim vArray(1 To lNumRows, 1 To lNumCols)
            With GetRange(lNumRows, lNumCols)
                For i = 1 To lNumRows
                    For j = 1 To lNumCols
                        vArray(i, j) = .Cells(i, j)
                    Next j
                Next i
            End With
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Number of dimensions must be 0, 1, or 2."
    End Select
End Sub

Private Sub Clear()
    Set m_oWorksheet = Nothing
    Set m_oSerializedObjects = New Collection
    m_lCursor = 1
    m_lIndent = 1
End Sub
