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

#Const USE_LATE_BINDING_IN_INTERFACES = True

Implements ISerializable

Private Const S_OK As Long = 0
Private Const SA_OFFSET_fFeatures As Long = 2
Private Const SA_OFFSET_cbElements As Long = 4

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

Public Property Get NumDimensions() As Long
    NumDimensions = m_lNumDimensions
End Property

Public Property Get Shape() As Long()
    Shape = m_alShape
End Property

Public Property Get Size(ByVal lDimension As Long) As Long
    Const PROCEDURE_NAME As String = "Tensor.Size"
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    Size = m_alShape(lDimension)
End Property

Public Property Get Stride(ByVal lDimension As Long) As Long
    Const PROCEDURE_NAME As String = "Tensor.Stride"
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    Stride = m_alStrides(lDimension)
End Property

Public Property Get NumElements() As Long
    NumElements = m_lNumElements
End Property

Public Property Get Address() As LongPtr
    Address = m_oData.Address
End Property

Private Sub Class_Initialize()
    Resize Array()
End Sub

Private Sub Class_Terminate()
    Reset
End Sub

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

Public Function ShapeEquals(ByVal vShape As Variant) As Boolean
    Dim lNumDimensions As Long
    Dim alOtherShape() As Long
    Dim i As Long
    Dim lDimensionSize As Long
    
    VariantToLongs vShape, lNumDimensions, alOtherShape
    If lNumDimensions <> m_lNumDimensions Then
        Exit Function
    End If
    For i = 1 To lNumDimensions
        lDimensionSize = alOtherShape(i)
        If lDimensionSize <> -1 And lDimensionSize <> m_alShape(i) Then
            Exit Function
        End If
    Next i
    ShapeEquals = True
End Function

Public Function Slice(ByVal lDimension As Long, _
                      ByVal lStart As Long, _
                      ByVal lEnd As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Slice"
    Dim lDimensionSize As Long
    Dim alNewShape() As Long
    Dim lStrideBytes As Long
    Dim lNumSteps As Long
    Dim i As Long
    Dim oResult As Tensor
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    If lStart < 1 Or lStart > m_alShape(lDimension) Then
        Err.Raise 9, PROCEDURE_NAME, "Start index must be between 1 and the size of the dimension."
    End If
    If lEnd < 1 Or lEnd > m_alShape(lDimension) Then
        Err.Raise 9, PROCEDURE_NAME, "End index must be between 1 and the size of the dimension."
    End If
    lDimensionSize = MaxLng2(0, lEnd - lStart + 1)
    alNewShape = m_alShape
    alNewShape(lDimension) = lDimensionSize
    Set oResult = New Tensor
    oResult.Resize alNewShape
    lStrideBytes = m_alStrides(lDimension) * m_oData.ElementSize
    lNumSteps = SizeBetween(lDimension + 1, m_lNumDimensions)
    For i = 1 To lNumSteps
        CopyMemory ByVal oResult.Address + ((i - 1) * lDimensionSize + 0) * lStrideBytes, _
                   ByVal m_oData.Address + ((i - 1) * m_alShape(lDimension) + (lStart - 1)) * lStrideBytes, _
                   lDimensionSize * lStrideBytes
    Next i
    Set Slice = oResult
End Function

Public Function Gather(ByVal lDimension As Long, _
                       ByVal vIndices As Variant) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Gather"
    Dim lNumIndices As Long
    Dim alIndices() As Long
    Dim alNewShape() As Long
    Dim lStrideBytes As Long
    Dim lNumSteps As Long
    Dim lIndex As Long
    Dim i As Long
    Dim j As Long
    Dim oResult As Tensor
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    VariantToLongs vIndices, lNumIndices, alIndices
    For i = 1 To lNumIndices
        lIndex = alIndices(i)
        If lIndex < 1 Or lIndex > m_alShape(lDimension) Then
            Err.Raise 9, PROCEDURE_NAME, "Index must be between 1 and the size of the dimension."
        End If
    Next i
    alNewShape = m_alShape
    alNewShape(lDimension) = lNumIndices
    Set oResult = New Tensor
    oResult.Resize alNewShape
    lStrideBytes = m_alStrides(lDimension) * m_oData.ElementSize
    lNumSteps = SizeBetween(lDimension + 1, m_lNumDimensions)
    For i = 1 To lNumSteps
        For j = 1 To lNumIndices
            lIndex = alIndices(j)
            CopyMemory ByVal oResult.Address + ((i - 1) * lNumIndices + (j - 1)) * lStrideBytes, _
                       ByVal m_oData.Address + ((i - 1) * m_alShape(lDimension) + (lIndex - 1)) * lStrideBytes, _
                       lStrideBytes
        Next j
    Next i
    Set Gather = oResult
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

Public Sub FromArray(ByRef adblArray() As Double)
    Dim lNumDimensions As Long
    Dim alShape() As Long
    Dim i As Long
    Dim lDimensionSize As Long
    Dim oData As Storage
    
    lNumDimensions = GetRank(adblArray)
    If lNumDimensions < 1 Then
        Resize Array()
        Exit Sub
    End If
    ReDim alShape(1 To lNumDimensions)
    For i = 1 To lNumDimensions
        lDimensionSize = UBound(adblArray, i) - LBound(adblArray, i) + 1
        alShape(i) = lDimensionSize
    Next i
    Set oData = New Storage
    oData.FromArray adblArray
    Resize alShape, oData
End Sub

Public Function ToArray() As Double()
    Dim A_() As Double
    
    CreateAlias A_
    ToArray = A_
    RemoveAlias A_
End Function

Public Sub FromRange(ByVal rngRange As Range, _
                     Optional ByVal bTranspose As Boolean)
    Dim lNumRows As Long
    Dim lNumCols As Long
    Dim oData As Storage
    
    If rngRange Is Nothing Then
        Resize Array()
        Exit Sub
    End If
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
End Sub

Public Sub CreateAlias(ByRef adblAlias() As Double)
    CopyMemory ByVal VarPtrArray(adblAlias), m_pDescriptor, SIZEOF_LONGPTR
End Sub

Public Sub RemoveAlias(ByRef adblAlias() As Double)
    ZeroMemory ByVal VarPtrArray(adblAlias), SIZEOF_LONGPTR
End Sub

Public Function Sum() As Double
    Dim i As Long
    Dim A_() As Double
    Dim dblSum As Double

    Ravel.CreateAlias A_
    For i = 1 To m_lNumElements
        dblSum = dblSum + A_(i)
    Next i
    Ravel.RemoveAlias A_
    Sum = dblSum
End Function

Public Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Tensor.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteLongArray m_alShape
        .WriteObject m_oData
    End With
End Sub

Public Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Tensor.Load"
    Dim alShape() As Long
    Dim oData As Storage

    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
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

Private Sub SetShape(ByVal vShape As Variant)
    Const PROCEDURE_NAME As String = "Tensor.SetShape"
    Dim i As Long
    Dim lDimensionSize As Long
    
    VariantToLongs vShape, m_lNumDimensions, m_alShape
    If m_lNumDimensions < 1 Then
        Erase m_alStrides
        m_lNumElements = 0
        Exit Sub
    End If
    ReDim m_alStrides(1 To m_lNumDimensions)
    m_lNumElements = 1
    For i = 1 To m_lNumDimensions
        lDimensionSize = m_alShape(i)
        If lDimensionSize < 0 Then
            Err.Raise 5, PROCEDURE_NAME, "Dimension size must be greater than or equal to 0."
        End If
        m_alStrides(i) = m_lNumElements
        m_lNumElements = m_lNumElements * lDimensionSize
    Next i
End Sub

Private Function SizeBetween(ByVal lStart As Long, _
                             ByVal lEnd As Long) As Long
    Dim i As Long
    Dim lDimensionSize As Long
    Dim lResult As Long

    If m_lNumDimensions < 1 Then
        Exit Function
    End If
    lStart = MaxLng2(1, lStart)
    lEnd = MinLng2(lEnd, m_lNumDimensions)
    lResult = 1
    For i = lStart To lEnd
        lDimensionSize = m_alShape(i)
        lResult = lResult * lDimensionSize
    Next
    SizeBetween = lResult
End Function

Private Sub AttachData(ByVal oData As Storage)
    Const PROCEDURE_NAME As String = "Tensor.AttachData"
    
    If oData Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Storage object is required."
    End If
    If oData.NumElements <> m_lNumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Number of elements in the data does not match the tensor's number of elements."
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
    Dim lErrorCode As Long
    Dim i As Long
    Dim lDimensionSize As Long
    
    If m_lNumDimensions = 0 Then
        Exit Sub
    End If
    lErrorCode = SafeArrayAllocDescriptorEx(m_oData.DataType, m_lNumDimensions, m_pDescriptor)
    If lErrorCode <> S_OK Then
        Err.Raise lErrorCode, PROCEDURE_NAME, GetSystemMessage(lErrorCode)
    End If
    CopyMemory ByVal m_pDescriptor + SA_OFFSET_fFeatures, FADF_STATIC Or FADF_HAVEVARTYPE, SIZEOF_INTEGER
    CopyMemory ByVal m_pDescriptor + SA_OFFSET_cbElements, m_oData.ElementSize, SIZEOF_LONG
    CopyMemory ByVal m_pDescriptor + SA_OFFSET_pvData, m_oData.Address, SIZEOF_LONGPTR
    For i = 1 To m_lNumDimensions
        lDimensionSize = m_alShape(i)
        CopyMemory ByVal m_pDescriptor + SA_OFFSET_rgsabound + (m_lNumDimensions - i) * 2 * SIZEOF_LONG, lDimensionSize, SIZEOF_LONG
        CopyMemory ByVal m_pDescriptor + SA_OFFSET_rgsabound + (m_lNumDimensions - i) * 2 * SIZEOF_LONG + SIZEOF_LONG, 1&, SIZEOF_LONG
    Next i
End Sub

Private Sub DestroyDescriptor()
    Const PROCEDURE_NAME As String = "Tensor.DestroyDescriptor"
    Dim lErrorCode As Long
    
    lErrorCode = SafeArrayDestroyDescriptor(m_pDescriptor)
    If lErrorCode <> S_OK Then
        Err.Raise lErrorCode, PROCEDURE_NAME, GetSystemMessage(lErrorCode)
    End If
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
