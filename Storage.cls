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

Private Const VARIANT_TYPE As Long = vbDouble
Private Const ELEMENT_SIZE As Long = SIZEOF_DOUBLE

#If Win64 Then
    Private Const SA_OFFSET_pvData As Long = 16
#Else
    Private Const SA_OFFSET_pvData As Long = 12
#End If

'Dimensions
Private m_lNumRows As Long
Private m_lNumCols As Long
Private m_lNumElements As Long

'Buffer
Private m_adblBuffer() As Double
Private m_pAddress As LongPtr

Public Property Get VarType() As VbVarType
    VarType = VARIANT_TYPE
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
        Err.Raise 5, PROCEDURE_NAME, "Number of rows must be >= 0."
    End If
    If lNumCols < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of columns must be >= 0."
    End If
    Clear
    m_lNumRows = lNumRows
    m_lNumCols = lNumCols
    m_lNumElements = m_lNumRows * m_lNumCols
    If m_lNumElements > 0 Then
        ReDim m_adblBuffer(1 To m_lNumRows, 1 To m_lNumCols)
        m_pAddress = VarPtr(m_adblBuffer(1, 1))
    End If
End Sub

Public Function Clone() As Storage
    Set Clone = New Storage
    Clone.Resize m_lNumRows, m_lNumCols
    If m_lNumElements > 0 Then
        CopyMemory ByVal Clone.Address, ByVal m_pAddress, m_lNumElements * ELEMENT_SIZE
    End If
End Function

Public Sub FromArray(ByRef adblArray() As Double)
    Dim i As Long
    Dim lNumDimensions As Long
    Dim lNumRows As Long
    Dim lNumCols As Long
    Dim pDescriptor As LongPtr
    Dim pData As LongPtr
    
    Clear
    lNumDimensions = GetRank(adblArray)
    If lNumDimensions < 1 Then
        Exit Sub
    End If
    lNumRows = 1
    lNumCols = 1
    For i = 1 To lNumDimensions
        lNumRows = lNumRows * lNumCols
        lNumCols = UBound(adblArray, i) - LBound(adblArray, i) + 1
    Next i
    Resize lNumRows, lNumCols
    If m_lNumElements > 0 Then
        CopyMemory pDescriptor, ByVal VarPtrArray(adblArray), SIZEOF_LONGPTR
        CopyMemory pData, ByVal pDescriptor + SA_OFFSET_pvData, SIZEOF_LONGPTR
        CopyMemory ByVal m_pAddress, ByVal pData, m_lNumElements * ELEMENT_SIZE
    End If
End Sub

Public Function ToArray() As Double()
    ToArray = m_adblBuffer
End Function

Public Sub FromRange(ByVal oRange As Range, _
                     Optional ByVal bTranspose As Boolean)
    Const CHUNK_ROWS As Long = 1000
    Dim i As Long
    Dim j As Long
    Dim lMinRow As Long
    Dim lMaxRow As Long
    Dim lMinCol As Long
    Dim lMaxCol As Long
    Dim lNumRows As Long
    Dim lNumCols As Long
    Dim lRowOffset As Long
    Dim lColOffset As Long
    Dim lStartRow As Long
    Dim lChunkRows As Long
    Dim vValues As Variant
    Dim oArea As Range

    Clear
    If oRange Is Nothing Then
        Exit Sub
    End If
    With oRange.Areas(1)
        lMinRow = .Row
        lMaxRow = .Row + .Rows.Count - 1
        lMinCol = .Column
        lMaxCol = .Column + .Columns.Count - 1
    End With
    For Each oArea In oRange.Areas
        lMinRow = MinLng(lMinRow, oArea.Row)
        lMinCol = MinLng(lMinCol, oArea.Column)
        lMaxRow = MaxLng(lMaxRow, oArea.Row + oArea.Rows.Count - 1)
        lMaxCol = MaxLng(lMaxCol, oArea.Column + oArea.Columns.Count - 1)
    Next oArea
    lNumRows = lMaxRow - lMinRow + 1
    lNumCols = lMaxCol - lMinCol + 1
    If bTranspose Then
        Resize lNumCols, lNumRows
    Else
        Resize lNumRows, lNumCols
    End If
    For Each oArea In oRange.Areas
        lRowOffset = oArea.Row - lMinRow
        lColOffset = oArea.Column - lMinCol
        For lStartRow = 1 To oArea.Rows.Count Step CHUNK_ROWS
            lChunkRows = MinLng(CHUNK_ROWS, oArea.Rows.Count - lStartRow + 1)
            vValues = oArea.Rows(lStartRow).Resize(lChunkRows).Value2
            If lChunkRows = 1 And oArea.Columns.Count = 1 Then '.Value2 returns a scalar
                If bTranspose Then
                    m_adblBuffer(lColOffset + 1, lRowOffset + lStartRow) = vValues
                Else
                    m_adblBuffer(lRowOffset + lStartRow, lColOffset + 1) = vValues
                End If
            Else
                If bTranspose Then
                    For i = 1 To lChunkRows
                        For j = 1 To oArea.Columns.Count
                            m_adblBuffer(lColOffset + j, lRowOffset + lStartRow + i - 1) = vValues(i, j)
                        Next j
                    Next i
                Else
                    For i = 1 To lChunkRows
                        For j = 1 To oArea.Columns.Count
                            m_adblBuffer(lRowOffset + lStartRow + i - 1, lColOffset + j) = vValues(i, j)
                        Next j
                    Next i
                End If
            End If
            Application.StatusBar = "FromRange progress: " & Format((lRowOffset + lStartRow + CHUNK_ROWS - 1) / lNumRows, "0.0%")
            DoEvents
        Next lStartRow
    Next oArea
End Sub

Private Sub Class_Terminate()
    Clear
End Sub

Private Sub Clear()
    m_lNumRows = 0
    m_lNumCols = 0
    m_lNumElements = 0
    Erase m_adblBuffer
    m_pAddress = NULL_PTR
End Sub

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "Storage.Save"
    
    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteDoubleArray m_adblBuffer
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "Storage.Load"
    
    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        FromArray .ReadDoubleArray()
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Sub ISerializable_Save(ByVal oSerializer As Object)
        Save oSerializer
    End Sub
    
    Private Sub ISerializable_Load(ByVal oSerializer As Object)
        Load oSerializer
    End Sub
#Else
    Private Sub ISerializable_Save(ByVal oSerializer As Serializer)
        Save oSerializer
    End Sub
    
    Private Sub ISerializable_Load(ByVal oSerializer As Serializer)
        Load oSerializer
    End Sub
#End If
