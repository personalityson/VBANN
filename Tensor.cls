VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tensor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements ISerializable

Private Const S_OK As Long = 0
Private Const SA_OFFSET_fFeatures As Long = 2
Private Const SA_OFFSET_cbElements As Long = 4
Private Const SA_OFFSET_cLocks As Long = 8

#If Win64 Then
    Private Const SA_OFFSET_pvData As Long = 16
    Private Const SA_OFFSET_rgsabound As Long = 24
#Else
    Private Const SA_OFFSET_pvData As Long = 12
    Private Const SA_OFFSET_rgsabound As Long = 16
#End If

Private Declare PtrSafe Function SafeArrayAllocDescriptorEx Lib "oleaut32" (ByVal vt As Integer, _
                                                                            ByVal cDims As Long, _
                                                                            ByRef ppsaOut As LongPtr) As Long

Private Declare PtrSafe Function SafeArrayDestroyDescriptor Lib "oleaut32" (ByVal psa As LongPtr) As Long

Private m_lNumDimensions As Long
Private m_alShape() As Long
Private m_alStrides() As Long
Private m_lNumElements As Long
Private m_oData As Storage
Private m_pDescriptor As LongPtr

Private Sub Class_Initialize()
    Resize Array()
End Sub

Private Sub Class_Terminate()
    Reset
End Sub

Public Property Get NumDimensions() As Long
    NumDimensions = m_lNumDimensions
End Property

Public Property Get Shape() As Long()
    Shape = m_alShape
End Property

Public Property Get Size(ByVal lDimension As Long) As Long
    Size = m_alShape(lDimension)
End Property

Public Property Get Stride(ByVal lDimension As Long) As Long
    Stride = m_alStrides(lDimension)
End Property

Public Property Get NumElements() As Long
    NumElements = m_lNumElements
End Property

Public Property Get Address() As LongPtr
    Address = m_oData.Address
End Property

Public Sub Resize(ByVal vShape As Variant, _
                  Optional ByVal oData As Storage)
    Reset
    SetShape vShape
    If oData Is Nothing Then
        AllocateData
    Else
        AttachData oData
    End If
    AllocateDescriptor
End Sub

Public Sub Reshape(ByVal vShape As Variant)
    Resize vShape, m_oData
End Sub

Public Function View(ByVal vShape As Variant) As Tensor
    Set View = New Tensor
    View.Resize vShape, m_oData
End Function

Public Function Clone() As Tensor
    Set Clone = New Tensor
    Clone.Resize m_alShape, m_oData.Clone
End Function

Public Sub Fill(ByVal dblValue As Double)
    m_oData.Fill dblValue
End Sub

Public Function Slice(ByVal lDimension As Long, _
                      ByVal lFrom As Long, _
                      ByVal lTo As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Slice"
    Dim alShape() As Long
    Dim lStrideSize As Long
    Dim lNumSteps As Long
    Dim i As Long
    Dim oResult As Tensor
    
    If lFrom < 1 Then
        Err.Raise 9, PROCEDURE_NAME, "Start index must be greater than or equal to 1."
    End If
    If lFrom > m_alShape(lDimension) Then
        Err.Raise 9, PROCEDURE_NAME, "Start index exceeds the dimension size."
    End If
    If lTo < lFrom Then
        Err.Raise 9, PROCEDURE_NAME, "End index must be greater than or equal to the start index."
    End If
    If lTo > m_alShape(lDimension) Then
        Err.Raise 9, PROCEDURE_NAME, "End index exceeds the dimension size."
    End If
    alShape = m_alShape
    alShape(lDimension) = lTo - lFrom + 1
    Set oResult = New Tensor
    oResult.Resize alShape
    lStrideSize = m_alStrides(lDimension) * m_oData.ElementSize
    lNumSteps = m_lNumElements / (m_alShape(lDimension) * m_alStrides(lDimension))
    For i = 1 To lNumSteps
        CopyMemory ByVal oResult.Address + ((i - 1) * alShape(lDimension) + 0) * lStrideSize, _
                   ByVal m_oData.Address + ((i - 1) * m_alShape(lDimension) + (lFrom - 1)) * lStrideSize, _
                   alShape(lDimension) * lStrideSize
    Next i
    Set Slice = oResult
End Function

Public Function Gather(ByVal lDimension As Long, _
                       ByVal vIndices As Variant) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Gather"
    Dim lNumIndices As Long
    Dim alIndices() As Long
    Dim alShape() As Long
    Dim lStrideSize As Long
    Dim lNumSteps As Long
    Dim lIndex As Long
    Dim i As Long
    Dim j As Long
    Dim oResult As Tensor
    
    VariantToLongs vIndices, lNumIndices, alIndices
    If lNumIndices < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "At least one index is required."
    End If
    For i = 1 To lNumIndices
        lIndex = alIndices(i)
        If lIndex < 1 Then
            Err.Raise 9, PROCEDURE_NAME, "Index must be greater than or equal to 1."
        End If
        If lIndex > m_alShape(lDimension) Then
            Err.Raise 9, PROCEDURE_NAME, "Index exceeds the dimension size."
        End If
    Next i
    alShape = m_alShape
    alShape(lDimension) = lNumIndices
    Set oResult = New Tensor
    oResult.Resize alShape
    lStrideSize = m_alStrides(lDimension) * m_oData.ElementSize
    lNumSteps = m_lNumElements / (m_alShape(lDimension) * m_alStrides(lDimension))
    For i = 1 To lNumSteps
        For j = 1 To lNumIndices
            lIndex = alIndices(j)
            CopyMemory ByVal oResult.Address + ((i - 1) * alShape(lDimension) + (j - 1)) * lStrideSize, _
                       ByVal m_oData.Address + ((i - 1) * m_alShape(lDimension) + (lIndex - 1)) * lStrideSize, _
                       lStrideSize
        Next j
    Next i
    Set Gather = oResult
End Function

'Public Function Squeeze() As Tensor
'    Dim lNumDimensions As Long
'    Dim alShape() As Long
'    Dim i As Long
'
'    alShape = m_alShape
'    For i = 1 To m_lNumDimensions
'        If alShape(i) > 1 Then
'            lNumDimensions = lNumDimensions + 1
'            alShape(lNumDimensions) = alShape(i)
'        End If
'    Next i
'
'End Function

Public Function Flatten(ByVal lFrom As Long, _
                        ByVal lTo As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Flatten"
    Dim lNumDimensions As Long
    Dim alShape() As Long
    Dim i As Long
    
    If lFrom < 1 Then
        Err.Raise 9, PROCEDURE_NAME, "Start dimension must be greater than or equal to 1."
    End If
    If lFrom > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Start dimension exceeds the number of dimensions."
    End If
    If lTo < lFrom Then
        Err.Raise 9, PROCEDURE_NAME, "End dimension must be greater than or equal to the start dimension."
    End If
    If lTo > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "End dimension exceeds the number of dimensions."
    End If
    alShape = m_alShape
    For i = lFrom + 1 To lTo
        alShape(lFrom) = alShape(lFrom) * alShape(i)
    Next i
    For i = lTo + 1 To m_lNumDimensions
        alShape(lFrom + i - lTo) = alShape(i)
    Next i
    lNumDimensions = m_lNumDimensions - (lTo - lFrom)
    ReDim Preserve alShape(1 To lNumDimensions)
    Set Flatten = View(alShape)
End Function

Public Function Ravel() As Tensor
    Static s_oFlatView As Tensor
    Dim bIsDirty As Boolean
    
    If s_oFlatView Is Nothing Then
        bIsDirty = True
    ElseIf s_oFlatView.Address <> m_oData.Address Then
        bIsDirty = True
    End If
    If bIsDirty Then
        If m_lNumDimensions > 1 Then
            Set s_oFlatView = View(m_lNumElements)
        Else
            Set s_oFlatView = Me
        End If
    End If
    Set Ravel = s_oFlatView
End Function

Public Sub FromRange(ByVal rngRange As Range, _
                     Optional ByVal bTranspose As Boolean)
    Const PROCEDURE_NAME As String = "Tensor.FromRange"
    Dim lNumRows As Long
    Dim lNumCols As Long
    Dim oData As Storage
    
    If rngRange Is Nothing Then
        Resize Array()
    Else
        With rngRange.Areas(1)
            lNumRows = .Rows.Count
            lNumCols = .Columns.Count
            Set oData = New Storage
            oData.FromRange .Cells, bTranspose
            If bTranspose Then
                Resize Array(lNumCols, lNumRows), oData
            Else
                Resize Array(lNumRows, lNumCols), oData
            End If
        End With
    End If
End Sub

Public Function ShapeEquals(ByVal vShape As Variant) As Boolean
    Dim lNumDimensions As Long
    Dim alShape() As Long
    Dim i As Long
    Dim lSize As Long
    
    VariantToLongs vShape, lNumDimensions, alShape
    If lNumDimensions <> m_lNumDimensions Then
        Exit Function
    End If
    For i = 1 To lNumDimensions
        lSize = alShape(i)
        If lSize <> -1 And lSize <> m_alShape(i) Then
            Exit Function
        End If
    Next i
    ShapeEquals = True
End Function

Public Sub CreateAlias(ByRef A() As Double)
    CopyMemory ByVal VarPtrArray(A), m_pDescriptor, SIZEOF_VBA_LONGPTR
End Sub

Public Sub RemoveAlias(ByRef A() As Double)
    ZeroMemory ByVal VarPtrArray(A), SIZEOF_VBA_LONGPTR
End Sub

Public Sub Save(ByVal oModel As Serializer)
    With oModel
        .WriteLongArray m_alShape
        .WriteObject m_oData
    End With
End Sub

Public Sub Load(ByVal oModel As Serializer)
    Dim alShape() As Long
    Dim oData As Storage

    With oModel
        alShape = .ReadLongArray()
        Set oData = .ReadObject()
    End With
    Resize alShape, oData
End Sub

Private Sub Reset()
    DestroyDescriptor
    m_lNumDimensions = 0
    Erase m_alShape
    Erase m_alStrides
    m_lNumElements = 0
    Set m_oData = Nothing
    m_pDescriptor = NULL_PTR
End Sub

Private Sub VariantToLongs(ByVal vValueOrArray As Variant, _
                           ByRef lNumElements As Long, _
                           ByRef alArray() As Long)
    Const PROCEDURE_NAME As String = "Tensor.VariantToLongs"
    Dim lRank As Long
    Dim lLBound As Long
    Dim lUBound As Long
    Dim i As Long
    
    lRank = GetRank(vValueOrArray)
    Select Case lRank
        Case -1 'Value
            lNumElements = 1
            ReDim alArray(1 To lNumElements)
            alArray(1) = CLng(vValueOrArray)
        Case 0 'Uninitialized array
            lNumElements = 0
            Erase alArray
        Case 1 'One-dimensional array
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
            Err.Raise 5, PROCEDURE_NAME, "Arrays with more than one dimension are not supported."
    End Select
End Sub

Private Sub SetShape(ByVal vShape As Variant)
    Const PROCEDURE_NAME As String = "Tensor.ParseShape"
    Dim i As Long
    Dim lSize As Long
    
    VariantToLongs vShape, m_lNumDimensions, m_alShape
    If m_lNumDimensions = 0 Then
        Erase m_alStrides
        m_lNumElements = 0
    Else
        ReDim m_alStrides(1 To m_lNumDimensions)
        m_lNumElements = 1
        For i = 1 To m_lNumDimensions
            lSize = m_alShape(i)
            If lSize < 1 Then
                Err.Raise 5, PROCEDURE_NAME, "Each dimension size must be greater than 0."
            End If
            m_alStrides(i) = m_lNumElements
            m_lNumElements = m_lNumElements * lSize
        Next i
    End If
End Sub

Private Sub AttachData(ByVal oData As Storage)
    Const PROCEDURE_NAME As String = "Tensor.AttachData"
    
    If oData Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "A valid Storage object is required."
    End If
    If oData.NumElements <> m_lNumElements Then
        Err.Raise 5, PROCEDURE_NAME, "The number of elements in the data does not match the tensor's number of elements."
    End If
    Set m_oData = oData
End Sub

Private Sub AllocateData()
    Dim lNumRows As Long
    Dim lNumCols As Long
    
    Set m_oData = New Storage
    If m_lNumDimensions > 0 Then
        lNumRows = m_alStrides(m_lNumDimensions)
        lNumCols = m_alShape(m_lNumDimensions)
        m_oData.Resize lNumRows, lNumCols
    End If
End Sub

Private Sub AllocateDescriptor()
    Const PROCEDURE_NAME As String = "Tensor.AllocateDescriptor"
    Const FADF_STATIC As Integer = &H2
    Const FADF_HAVEVARTYPE As Integer = &H80
    Dim i As Long
    Dim lSize As Long
    Dim lResult As Long
    
    If m_lNumDimensions = 0 Then
        Exit Sub
    End If
    lResult = SafeArrayAllocDescriptorEx(m_oData.DataType, m_lNumDimensions, m_pDescriptor)
    If lResult <> S_OK Then
        Err.Raise lResult, PROCEDURE_NAME, GetSystemMessage(lResult)
    End If
    CopyMemory ByVal m_pDescriptor + SA_OFFSET_fFeatures, FADF_STATIC Or FADF_HAVEVARTYPE, SIZEOF_VBA_INTEGER
    CopyMemory ByVal m_pDescriptor + SA_OFFSET_cbElements, m_oData.ElementSize, SIZEOF_VBA_LONG
    CopyMemory ByVal m_pDescriptor + SA_OFFSET_pvData, m_oData.Address, SIZEOF_VBA_LONGPTR
    For i = 1 To m_lNumDimensions
        lSize = m_alShape(i)
        CopyMemory ByVal m_pDescriptor + SA_OFFSET_rgsabound + (m_lNumDimensions - i) * 2 * SIZEOF_VBA_LONG, lSize, SIZEOF_VBA_LONG
        CopyMemory ByVal m_pDescriptor + SA_OFFSET_rgsabound + (m_lNumDimensions - i) * 2 * SIZEOF_VBA_LONG + SIZEOF_VBA_LONG, 1&, SIZEOF_VBA_LONG
    Next i
End Sub

Private Sub DestroyDescriptor()
    Const PROCEDURE_NAME As String = "Tensor.DestroyDescriptor"
    Dim lResult As Long
    
    lResult = SafeArrayDestroyDescriptor(m_pDescriptor)
    If lResult <> S_OK Then
        Err.Raise lResult, PROCEDURE_NAME, GetSystemMessage(lResult)
    End If
End Sub

Private Sub ISerializable_Save(ByVal oModel As Serializer)
    Save oModel
End Sub

Private Sub ISerializable_Load(ByVal oModel As Serializer)
    Load oModel
End Sub
