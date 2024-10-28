VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Storage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

Private Const DATA_TYPE As Long = vbDouble
Private Const ELEMENT_SIZE As Long = SIZEOF_DOUBLE

#If Win64 Then
    Private Const SA_OFFSET_pvData As Long = 16
#Else
    Private Const SA_OFFSET_pvData As Long = 12
#End If

Private m_lNumRows As Long
Private m_lNumCols As Long
Private m_lNumElements As Long
Private m_adblBuffer() As Double
Private m_pAddress As LongPtr

Public Property Get DataType() As VbVarType
    DataType = DATA_TYPE
End Property

Public Property Get ElementSize() As Long
    ElementSize = ELEMENT_SIZE
End Property

Public Property Get NumRows() As Long
    NumRows = m_lNumRows
End Property

Public Property Get NumCols() As Long
    NumCols = m_lNumCols
End Property

Public Property Get NumElements() As Long
    NumElements = m_lNumElements
End Property

Public Property Get Address() As LongPtr
    Address = m_pAddress
End Property

Public Sub Resize(ByVal lNumRows As Long, _
                  ByVal lNumCols As Long)
    Const PROCEDURE_NAME As String = "Storage.Resize"
    
    If lNumRows < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of rows must be greater than or equal to 0."
    End If
    If lNumCols < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of columns must be greater than or equal to 0."
    End If
    Reset
    m_lNumRows = lNumRows
    m_lNumCols = lNumCols
    m_lNumElements = m_lNumRows * m_lNumCols
    If m_lNumElements < 1 Then
        Exit Sub
    End If
    ReDim m_adblBuffer(1 To m_lNumRows, 1 To m_lNumCols)
    m_pAddress = VarPtr(m_adblBuffer(1, 1))
End Sub

Public Function Clone() As Storage
    Set Clone = New Storage
    Clone.Resize m_lNumRows, m_lNumCols
    CopyMemory ByVal Clone.Address, ByVal m_pAddress, m_lNumElements * ELEMENT_SIZE
End Function

Public Sub FromArray(ByRef adblArray() As Double)
    Dim lNumDimensions As Long
    Dim lNumRows As Long
    Dim lNumCols As Long
    Dim i As Long
    Dim pDescriptor As LongPtr
    Dim pData As LongPtr
    
    lNumDimensions = GetRank(adblArray)
    If lNumDimensions < 1 Then
        Reset
        Exit Sub
    End If
    lNumRows = 1
    lNumCols = 1
    For i = 1 To lNumDimensions
        lNumRows = lNumRows * lNumCols
        lNumCols = UBound(adblArray, i) - LBound(adblArray, i) + 1
    Next i
    Resize lNumRows, lNumCols
    CopyMemory pDescriptor, ByVal VarPtrArray(adblArray), SIZEOF_LONGPTR
    CopyMemory pData, ByVal pDescriptor + SA_OFFSET_pvData, SIZEOF_LONGPTR
    CopyMemory ByVal m_pAddress, ByVal pData, m_lNumElements * ELEMENT_SIZE
End Sub

Public Function ToArray() As Double()
    ToArray = m_adblBuffer
End Function

Public Sub FromRange(ByVal oRange As Range, _
                     Optional bTrans As Boolean)
    Dim lNumRows As Long
    Dim lNumCols As Long
    Dim i As Long
    Dim j As Long
    
    If oRange Is Nothing Then
        Reset
        Exit Sub
    End If
    With oRange.Areas(1)
        lNumRows = .Rows.Count
        lNumCols = .Columns.Count
        If bTrans Then
            Resize lNumCols, lNumRows
            For i = 1 To lNumRows
                For j = 1 To lNumCols
                    m_adblBuffer(j, i) = .Cells(i, j)
                Next j
            Next i
        Else
            Resize lNumRows, lNumCols
            For i = 1 To lNumRows
                For j = 1 To lNumCols
                    m_adblBuffer(i, j) = .Cells(i, j)
                Next j
            Next i
        End If
    End With
End Sub

Private Sub Class_Terminate()
    Reset
End Sub

Private Sub Reset()
    m_lNumRows = 0
    m_lNumCols = 0
    m_lNumElements = 0
    Erase m_adblBuffer
    m_pAddress = NULL_PTR
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Storage.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteDoubleArray m_adblBuffer
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Storage.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        FromArray .ReadDoubleArray()
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Sub ISerializable_Save(ByVal oWorksheet As Object)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Object)
        Load oWorksheet
    End Sub
#Else
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If
