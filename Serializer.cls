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

Private Const SERIALIZABLE_TYPE_ADAM As String = "Adam"
Private Const SERIALIZABLE_TYPE_BCE_LOSS As String = "BCELoss"
Private Const SERIALIZABLE_TYPE_CCE_LOSS As String = "CCELoss"
Private Const SERIALIZABLE_TYPE_DROPOUT_LAYER As String = "DropoutLayer"
Private Const SERIALIZABLE_TYPE_FULLY_CONNECTED_LAYER As String = "FullyConnectedLayer"
Private Const SERIALIZABLE_TYPE_INPUT_NORMALIZATION_LAYER As String = "InputNormalizationLayer"
Private Const SERIALIZABLE_TYPE_L1_LOSS As String = "L1Loss"
Private Const SERIALIZABLE_TYPE_L2_LOSS As String = "L2Loss"
Private Const SERIALIZABLE_TYPE_LEAKY_RELU_LAYER As String = "LeakyReLULayer"
Private Const SERIALIZABLE_TYPE_MEMO As String = "Memo"
Private Const SERIALIZABLE_TYPE_NOTHING As String = "Nothing"
Private Const SERIALIZABLE_TYPE_PARAMETER As String = "Parameter"
Private Const SERIALIZABLE_TYPE_SEQUENTIAL As String = "Sequential"
Private Const SERIALIZABLE_TYPE_SGDM As String = "SGDM"
Private Const SERIALIZABLE_TYPE_SIGMOID_LAYER As String = "SigmoidLayer"
Private Const SERIALIZABLE_TYPE_SOFTMAX_LAYER As String = "SoftmaxLayer"
Private Const SERIALIZABLE_TYPE_STORAGE As String = "Storage"
Private Const SERIALIZABLE_TYPE_TENSOR As String = "Tensor"

Private m_oWorksheet As Worksheet
Private m_lRow As Long
Private m_lNumSerializedObjects As Long
Private m_aoMemo() As ISerializable

Public Property Get Worksheet() As Worksheet
    Set Worksheet = m_oWorksheet
End Property

Public Sub Init(ByVal sName As String, _
                ByVal bOverwrite As Boolean)
    Set m_oWorksheet = CreateWorksheet(ThisWorkbook, sName, bOverwrite)
    MoveFirst
    m_lNumSerializedObjects = 0
    Erase m_aoMemo
End Sub

Public Sub MoveFirst()
    m_lRow = 1
End Sub

Public Sub MoveLast()
    m_lRow = GetLastRow(m_oWorksheet) + 1
End Sub

Public Sub WriteObject(ByVal oObject As ISerializable)
    Dim lIndex As Long

    If oObject Is Nothing Then
        WriteNothing
    Else
        lIndex = IndexOf(oObject)
        If lIndex > 0 Then
            WriteMemo lIndex
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
        .Value = sValue
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
    m_lNumSerializedObjects = m_lNumSerializedObjects + 1
    ReDim Preserve m_aoMemo(1 To m_lNumSerializedObjects)
    Set m_aoMemo(m_lNumSerializedObjects) = oObject
End Sub

Private Function IndexOf(ByVal oObject As ISerializable) As Long
    Dim i As Long
    
    For i = 1 To m_lNumSerializedObjects
        If m_aoMemo(i) Is oObject Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

Private Function GetCell() As Range
    Set GetCell = m_oWorksheet.Cells(m_lRow, 1)
    m_lRow = m_lRow + 1
End Function

Private Function GetRange(ByVal lNumRows As Long, _
                          ByVal lNumCols As Long) As Range
    Const PROCEDURE_NAME As String = "Serializer.GetRange"
    
    If lNumRows < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of rows must be greater than 0."
    End If
    If lNumCols < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of columns must be greater than 0."
    End If
    Set GetRange = m_oWorksheet.Cells(m_lRow, 1).Resize(lNumRows, lNumCols)
    m_lRow = m_lRow + lNumRows
End Function

Private Sub WriteNothing()
    WriteString SERIALIZABLE_TYPE_NOTHING
End Sub

Private Sub WriteMemo(ByVal lIndex As Long)
    WriteString SERIALIZABLE_TYPE_MEMO
    WriteLong lIndex
End Sub

Private Sub WriteSerializable(ByVal oObject As ISerializable)
    WriteString TypeName(oObject)
    Memoize oObject
    oObject.Save Me
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
        .Value = vArray
    End With
End Sub

Private Function ReadNothing() As ISerializable
    'Do nothing
End Function

Private Function ReadMemo() As ISerializable
    Dim lIndex As Long
    
    lIndex = ReadLong()
    Set ReadMemo = m_aoMemo(lIndex)
End Function

Private Function ReadSerializable(ByVal sTypeName As String) As ISerializable
    Set ReadSerializable = CreateInstance(sTypeName)
    Memoize ReadSerializable
    ReadSerializable.Load Me
End Function

Private Function CreateInstance(ByVal sTypeName As String) As ISerializable
    Const PROCEDURE_NAME As String = "Serializer.CreateInstance"
    
    Select Case sTypeName
        Case SERIALIZABLE_TYPE_ADAM
            Set CreateInstance = New Adam
        Case SERIALIZABLE_TYPE_BCE_LOSS
            Set CreateInstance = New BCELoss
        Case SERIALIZABLE_TYPE_CCE_LOSS
            Set CreateInstance = New CCELoss
        Case SERIALIZABLE_TYPE_DROPOUT_LAYER
            Set CreateInstance = New DropoutLayer
        Case SERIALIZABLE_TYPE_FULLY_CONNECTED_LAYER
            Set CreateInstance = New FullyConnectedLayer
        Case SERIALIZABLE_TYPE_INPUT_NORMALIZATION_LAYER
            Set CreateInstance = New InputNormalizationLayer
        Case SERIALIZABLE_TYPE_L1_LOSS
            Set CreateInstance = New L1Loss
        Case SERIALIZABLE_TYPE_L2_LOSS
            Set CreateInstance = New L2Loss
        Case SERIALIZABLE_TYPE_LEAKY_RELU_LAYER
            Set CreateInstance = New LeakyReLULayer
        Case SERIALIZABLE_TYPE_PARAMETER
            Set CreateInstance = New Parameter
        Case SERIALIZABLE_TYPE_SEQUENTIAL
            Set CreateInstance = New Sequential
        Case SERIALIZABLE_TYPE_SGDM
            Set CreateInstance = New SGDM
        Case SERIALIZABLE_TYPE_SIGMOID_LAYER
            Set CreateInstance = New SigmoidLayer
        Case SERIALIZABLE_TYPE_SOFTMAX_LAYER
            Set CreateInstance = New SoftmaxLayer
        Case SERIALIZABLE_TYPE_STORAGE
            Set CreateInstance = New Storage
        Case SERIALIZABLE_TYPE_TENSOR
            Set CreateInstance = New Tensor
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Unsupported serializable type."
    End Select
End Function

Private Sub ReadArray(ByRef vArray As Variant)
    Const PROCEDURE_NAME As String = "Serializer.ReadArray"
    Dim lNumDimensions As Long
    Dim lNumRows As Long
    Dim lNumCols As Long
    Dim i As Long
    Dim j As Long
    
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
