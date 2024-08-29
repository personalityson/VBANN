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

Private Declare PtrSafe Function SafeArrayAllocDescriptorEx Lib "oleaut32.dll" (ByVal vt As Integer, _
                                                                                ByVal cDims As Long, _
                                                                                ByRef ppsaOut As LongPtr) As Long

Private Declare PtrSafe Function SafeArrayDestroyDescriptor Lib "oleaut32.dll" (ByVal psa As LongPtr) As Long

Public Enum ReductionType
    rdcSum
    rdcMean
    rdcMin
    rdcMax
    rdcProduct
End Enum

Private m_lNumDimensions As Long
Private m_alShape() As Long
Private m_alInnerStride() As Long
Private m_alOuterStride() As Long
Private m_lNumElements As Long
Private m_oData As Storage
Private m_pDescriptor As LongPtr
Private m_bIsDescriptorAllocated As Boolean

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

Public Property Get NumElements() As Long
    NumElements = m_lNumElements
End Property

Public Property Get Address() As LongPtr
    Address = m_oData.Address
End Property

Private Sub Class_Initialize()
    Set m_oData = New Storage
End Sub

Private Sub Class_Terminate()
    Reset
End Sub

Public Sub Resize(ByVal vShape As Variant, _
                  Optional ByVal oData As Storage)
    Reset
    SetShape vShape
    CalculateStrides
    If Not oData Is Nothing Then
        AttachData oData
    Else
        AllocateData
    End If
End Sub

Public Sub Reshape(ByVal vShape As Variant)
    Me.Resize vShape, m_oData
End Sub

Public Function View(ByVal vShape As Variant) As Tensor
    Set View = New Tensor
    View.Resize vShape, m_oData
End Function

Public Function Clone() As Tensor
    Set Clone = New Tensor
    Clone.Resize m_alShape, m_oData.Clone
End Function

Public Function Flatten() As Tensor
    Static s_oFlatView As Tensor
    Dim bIsDirty As Boolean
    
    If s_oFlatView Is Nothing Then
        bIsDirty = True
    ElseIf s_oFlatView.Address <> m_oData.Address Then
        bIsDirty = True
    End If
    If bIsDirty Then
        If m_lNumDimensions > 1 Then
            Set s_oFlatView = Me.View(m_lNumElements)
        Else
            Set s_oFlatView = Me
        End If
    End If
    Set Flatten = s_oFlatView
End Function

Public Function Slice(ByVal lDimension As Long, _
                      ByVal lStart As Long, _
                      ByVal lEnd As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Slice"

    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    If lStart < 1 Or lStart > m_alShape(lDimension) Then
        Err.Raise 9, PROCEDURE_NAME, "Start index must be between 1 and the size of the dimension."
    End If
    If lEnd < 1 Or lEnd > m_alShape(lDimension) Then
        Err.Raise 9, PROCEDURE_NAME, "End index must be between 1 and the size of the dimension."
    End If
    Set Slice = SliceInternal(lDimension, lStart, lEnd)
End Function

Public Function Gather(ByVal lDimension As Long, _
                       ByVal vIndices As Variant) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Gather"
    Dim lNumIndices As Long
    Dim alIndices() As Long
    Dim i As Long
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    ParseVariantToLongArray vIndices, lNumIndices, alIndices
    For i = 1 To lNumIndices
        If alIndices(i) < 1 Or alIndices(i) > m_alShape(lDimension) Then
            Err.Raise 9, PROCEDURE_NAME, "Index must be between 1 and the size of the dimension."
        End If
    Next i
    Set Gather = GatherInternal(lDimension, lNumIndices, alIndices)
End Function

Public Function Tile(ByVal lDimension As Long, _
                     ByVal lNumRepeats As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Tile"
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    If lNumRepeats < 0 Then
        Err.Raise 9, PROCEDURE_NAME, "Number of repeats must be greater than or equal to 0."
    End If
    Set Tile = TileInternal(lDimension, lNumRepeats)
End Function

Public Function Reduce(ByVal lDimension As Long, _
                       ByVal eReduction As ReductionType) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Reduce"
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    Set Reduce = ReduceInternal(lDimension, eReduction)
End Function

Public Function Sum() As Double
    Sum = AggregateInternal(rdcSum)
End Function

Public Function Mean() As Double
    Mean = AggregateInternal(rdcMean)
End Function

Public Function Min() As Double
    Min = AggregateInternal(rdcMin)
End Function

Public Function Max() As Double
    Max = AggregateInternal(rdcMax)
End Function

Public Function Product() As Double
    Product = AggregateInternal(rdcProduct)
End Function

Public Sub FromRange(ByVal rngRange As Range, _
                     Optional ByVal bTrans As Boolean)
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
        oData.FromRange .Cells, bTrans
        If bTrans Then
            Resize Array(lNumCols, lNumRows), oData
        Else
            Resize Array(lNumRows, lNumCols), oData
        End If
    End With
End Sub

Public Sub FromArray(ByRef adblArray() As Double)
    Dim lNumDimensions As Long
    Dim alShape() As Long
    Dim i As Long
    Dim oData As Storage
    
    lNumDimensions = GetRank(adblArray)
    If lNumDimensions < 1 Then
        Resize Array()
        Exit Sub
    End If
    ReDim alShape(1 To lNumDimensions)
    For i = 1 To lNumDimensions
        alShape(i) = UBound(adblArray, i) - LBound(adblArray, i) + 1
    Next i
    Set oData = New Storage
    oData.FromArray adblArray
    Resize alShape, oData
End Sub

Public Function ToArray() As Double()
    Dim A_() As Double
    
    Me.CreateAlias A_
    ToArray = A_
    Me.RemoveAlias A_
End Function

Public Sub FillConstant(ByVal dblValue As Double)
    If m_lNumElements < 1 Then
        Exit Sub
    End If
    TileMemory m_oData.Address, m_oData.NumElements * m_oData.ElementSize, VarPtr(dblValue), SIZEOF_DOUBLE
End Sub

Public Sub FillUniform(Optional ByVal dblLow As Double = 0, _
                       Optional ByVal dblHigh As Double = 1)
    Dim i As Long
    Dim A_() As Double
    
    If m_lNumElements < 1 Then
        Exit Sub
    End If
    Me.Flatten.CreateAlias A_
    For i = 1 To m_lNumElements
        A_(i) = dblLow + (dblHigh - dblLow) * Rnd()
    Next i
    Me.Flatten.RemoveAlias A_
End Sub

Public Sub FillNormal(Optional ByVal dblMu As Double = 0, _
                      Optional ByVal dblSigma As Double = 1)
    Dim i As Long
    Dim A_() As Double
    
    If m_lNumElements < 1 Then
        Exit Sub
    End If
    Me.Flatten.CreateAlias A_
    For i = 1 To m_lNumElements
        A_(i) = dblMu + dblSigma * NormRand()
    Next i
    Me.Flatten.RemoveAlias A_
End Sub

'-1 = wildcard
Public Function ShapeEquals(ByVal vShape As Variant) As Boolean
    Dim lOtherNumDimensions As Long
    Dim alOtherShape() As Long
    Dim i As Long
    
    ParseVariantToLongArray vShape, lOtherNumDimensions, alOtherShape
    If lOtherNumDimensions <> m_lNumDimensions Then
        Exit Function
    End If
    For i = 1 To lOtherNumDimensions
        If alOtherShape(i) <> -1 And alOtherShape(i) <> m_alShape(i) Then
            Exit Function
        End If
    Next i
    ShapeEquals = True
End Function

Public Sub CreateAlias(ByRef adblAlias() As Double)
    AllocateDescriptorLazily
    CopyMemory ByVal VarPtrArray(adblAlias), m_pDescriptor, SIZEOF_LONGPTR
End Sub

Public Sub RemoveAlias(ByRef adblAlias() As Double)
    ZeroMemory ByVal VarPtrArray(adblAlias), SIZEOF_LONGPTR
End Sub

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
    Erase m_alInnerStride
    Erase m_alOuterStride
    m_lNumElements = 0
    Set m_oData = New Storage
    m_pDescriptor = NULL_PTR
    m_bIsDescriptorAllocated = False
End Sub

Private Sub SetShape(ByVal vShape As Variant)
    Const PROCEDURE_NAME As String = "Tensor.SetShape"
    Dim i As Long

    ParseVariantToLongArray vShape, m_lNumDimensions, m_alShape
    For i = 1 To m_lNumDimensions
        If m_alShape(i) < 0 Then
            Err.Raise 5, PROCEDURE_NAME, "Dimension size must be greater than or equal to 0."
        End If
    Next i
End Sub

Private Sub CalculateStrides()
    Dim i As Long
    Dim lInnerStride As Long
    Dim lOuterStride As Long

    If m_lNumDimensions < 1 Then
        Erase m_alInnerStride
        Erase m_alOuterStride
        m_lNumElements = 0
        Exit Sub
    End If
    ReDim m_alInnerStride(1 To m_lNumDimensions)
    ReDim m_alOuterStride(1 To m_lNumDimensions)
    lInnerStride = 1
    lOuterStride = 1
    For i = 1 To m_lNumDimensions
        m_alInnerStride(i) = lInnerStride
        m_alOuterStride(m_lNumDimensions - i + 1) = lOuterStride
        lInnerStride = lInnerStride * m_alShape(i)
        lOuterStride = lOuterStride * m_alShape(m_lNumDimensions - i + 1)
    Next i
    m_lNumElements = lInnerStride
End Sub

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
        lNumRows = m_alInnerStride(m_lNumDimensions)
        lNumCols = m_alShape(m_lNumDimensions)
'        lNumRows = m_alShape(1)
'        lNumCols = m_alOuterStride(1)
        m_oData.Resize lNumRows, lNumCols
    End If
End Sub

Private Sub AllocateDescriptor()
    Const PROCEDURE_NAME As String = "Tensor.AllocateDescriptor"
    Const FADF_STATIC As Integer = &H2
    Const FADF_HAVEVARTYPE As Integer = &H80
    Dim lErrorCode As Long
    Dim i As Long
    
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
        CopyMemory ByVal m_pDescriptor + SA_OFFSET_rgsabound + (m_lNumDimensions - i) * 2 * SIZEOF_LONG, m_alShape(i), SIZEOF_LONG
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

Private Sub AllocateDescriptorLazily()
    If Not m_bIsDescriptorAllocated Then
        AllocateDescriptor
        m_bIsDescriptorAllocated = True
    End If
End Sub

Private Function SliceInternal(ByVal lDimension As Long, _
                               ByVal lStart As Long, _
                               ByVal lEnd As Long) As Tensor
    Dim lOldSize As Long
    Dim lNewSize As Long
    Dim alNewShape() As Long
    Dim lInnerStrideBytes As Long
    Dim lOuterStride As Long
    Dim i As Long
    Dim oResult As Tensor
    
    lOldSize = m_alShape(lDimension)
    lNewSize = MaxLng2(0, lEnd - lStart + 1)
    alNewShape = m_alShape
    alNewShape(lDimension) = lNewSize
    Set oResult = New Tensor
    oResult.Resize alNewShape
    lInnerStrideBytes = m_alInnerStride(lDimension) * m_oData.ElementSize
    lOuterStride = m_alOuterStride(lDimension)
    For i = 1 To lOuterStride
        CopyMemory ByVal oResult.Address + ((i - 1) * lNewSize + (1 - 1)) * lInnerStrideBytes, _
                   ByVal m_oData.Address + ((i - 1) * lOldSize + (lStart - 1)) * lInnerStrideBytes, _
                   lNewSize * lInnerStrideBytes
    Next i
    Set SliceInternal = oResult
End Function

Public Function GatherInternal(ByVal lDimension As Long, _
                               ByVal lNumIndices As Long, _
                               ByRef alIndices() As Long) As Tensor
    Dim lOldSize As Long
    Dim lNewSize As Long
    Dim alNewShape() As Long
    Dim lInnerStrideBytes As Long
    Dim lOuterStride As Long
    Dim i As Long
    Dim j As Long
    Dim lIndex As Long
    Dim oResult As Tensor
    
    lOldSize = m_alShape(lDimension)
    lNewSize = MinLng2(lOldSize, lNumIndices)
    alNewShape = m_alShape
    alNewShape(lDimension) = lNewSize
    Set oResult = New Tensor
    oResult.Resize alNewShape
    lInnerStrideBytes = m_alInnerStride(lDimension) * m_oData.ElementSize
    lOuterStride = m_alOuterStride(lDimension)
    For i = 1 To lOuterStride
        For j = 1 To lNewSize
            lIndex = alIndices(j)
            CopyMemory ByVal oResult.Address + ((i - 1) * lNewSize + (j - 1)) * lInnerStrideBytes, _
                       ByVal m_oData.Address + ((i - 1) * lOldSize + (lIndex - 1)) * lInnerStrideBytes, _
                       lInnerStrideBytes
        Next j
    Next i
    Set GatherInternal = oResult
End Function

Private Function TileInternal(ByVal lDimension As Long, _
                              ByVal lNumRepeats As Long) As Tensor
    Dim lOldSize As Long
    Dim lNewSize As Long
    Dim alNewShape() As Long
    Dim lInnerStrideBytes As Long
    Dim lOuterStride As Long
    Dim i As Long
    Dim oResult As Tensor
    
    lOldSize = m_alShape(lDimension)
    lNewSize = lNumRepeats * lOldSize
    alNewShape = m_alShape
    alNewShape(lDimension) = lNewSize
    Set oResult = New Tensor
    oResult.Resize alNewShape
    lInnerStrideBytes = m_alInnerStride(lDimension) * m_oData.ElementSize
    lOuterStride = m_alOuterStride(lDimension)
    For i = 1 To lOuterStride
        TileMemory oResult.Address + ((i - 1) * lNewSize) * lInnerStrideBytes, _
                   lNewSize * lInnerStrideBytes, _
                   m_oData.Address + ((i - 1) * lOldSize) * lInnerStrideBytes, _
                   lOldSize * lInnerStrideBytes
    Next i
    Set TileInternal = oResult
End Function

Private Function ReduceInternal(ByVal lDimension As Long, _
                                ByVal eReduction As ReductionType) As Tensor
    Dim lOldSize As Long
    Dim lNewSize As Long
    Dim alNewShape() As Long
    Dim lInnerStride As Long
    Dim lOuterStride As Long
    Dim i As Long
    Dim k As Long
    Dim A_() As Double
    Dim R_() As Double
    Dim A As Tensor
    Dim R As Tensor
    Dim oResult As Tensor
    
    lOldSize = m_alShape(lDimension)
    lNewSize = 1
    alNewShape = m_alShape
    alNewShape(lDimension) = lNewSize
    Set oResult = New Tensor
    oResult.Resize alNewShape
    lInnerStride = m_alInnerStride(lDimension)
    lOuterStride = m_alOuterStride(lDimension)
    Set A = Me.View(Array(lInnerStride, lOldSize, lOuterStride))
    Set R = oResult.View(Array(lInnerStride, lNewSize, lOuterStride))
    A.CreateAlias A_
    R.CreateAlias R_
    For i = 1 To lInnerStride
        For k = 1 To lOuterStride
            Select Case eReduction
                Case rdcSum
                    R_(i, 1, k) = SumInSlice(A_, i, lOldSize, k)
                Case rdcMean
                    R_(i, 1, k) = MeanInSlice(A_, i, lOldSize, k)
                Case rdcMin
                    R_(i, 1, k) = MinInSlice(A_, i, lOldSize, k)
                Case rdcMax
                    R_(i, 1, k) = MaxInSlice(A_, i, lOldSize, k)
                Case rdcProduct
                    R_(i, 1, k) = ProductInSlice(A_, i, lOldSize, k)
            End Select
        Next k
    Next i
    A.RemoveAlias A_
    R.RemoveAlias R_
    Set ReduceInternal = oResult
End Function

Private Function AggregateInternal(ByVal eReduction As ReductionType) As Double
    Dim A_() As Double
    Dim dblResult As Double
    
    With GetFullSlice
        .CreateAlias A_
        Select Case eReduction
            Case rdcSum
                dblResult = SumInSlice(A_, 1, m_lNumElements, 1)
            Case rdcMean
                dblResult = MeanInSlice(A_, 1, m_lNumElements, 1)
            Case rdcMin
                dblResult = MinInSlice(A_, 1, m_lNumElements, 1)
            Case rdcMax
                dblResult = MaxInSlice(A_, 1, m_lNumElements, 1)
            Case rdcProduct
                dblResult = ProductInSlice(A_, 1, m_lNumElements, 1)
        End Select
        .RemoveAlias A_
    End With
    AggregateInternal = dblResult
End Function

Private Function GetFullSlice() As Tensor
    Static s_oFullSlice As Tensor
    Dim bIsDirty As Boolean
    
    If s_oFullSlice Is Nothing Then
        bIsDirty = True
    ElseIf s_oFullSlice.Address <> m_oData.Address Then
        bIsDirty = True
    End If
    If bIsDirty Then
        Set s_oFullSlice = Me.View(Array(1, m_lNumElements, 1))
    End If
    Set GetFullSlice = s_oFullSlice
End Function

Private Function SumInSlice(ByRef A_() As Double, _
                            ByVal i As Long, _
                            ByVal lNumElements As Long, _
                            ByVal k As Long) As Double
    Dim j As Long
    Dim dblResult As Double
    
    If lNumElements < 1 Then
        Exit Function
    End If
    dblResult = 0
    For j = 1 To lNumElements
         dblResult = dblResult + A_(i, j, k)
    Next j
    SumInSlice = dblResult
End Function

Private Function MeanInSlice(ByRef A_() As Double, _
                             ByVal i As Long, _
                             ByVal lNumElements As Long, _
                             ByVal k As Long) As Double
    If lNumElements < 1 Then
        Exit Function
    End If
    MeanInSlice = SumInSlice(A_, i, lNumElements, k) / lNumElements
End Function

Private Function MinInSlice(ByRef A_() As Double, _
                            ByVal i As Long, _
                            ByVal lNumElements As Long, _
                            ByVal k As Long) As Double
    Const PROCEDURE_NAME As String = "Tensor.MinInSlice"
    Dim j As Long
    Dim dblResult As Double
    
    If lNumElements < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Reducing over an empty dimension is not allowed."
    End If
    dblResult = A_(i, 1, k)
    For j = 2 To lNumElements
         dblResult = MinLng2(dblResult, A_(i, j, k))
    Next j
    MinInSlice = dblResult
End Function

Private Function MaxInSlice(ByRef A_() As Double, _
                            ByVal i As Long, _
                            ByVal lNumElements As Long, _
                            ByVal k As Long) As Double
    Const PROCEDURE_NAME As String = "Tensor.MaxInSlice"
    Dim j As Long
    Dim dblResult As Double
    
    If lNumElements < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Reducing over an empty dimension is not allowed."
    End If
    dblResult = A_(i, 1, k)
    For j = 2 To lNumElements
         dblResult = MaxLng2(dblResult, A_(i, j, k))
    Next j
    MaxInSlice = dblResult
End Function

Private Function ProductInSlice(ByRef A_() As Double, _
                                ByVal i As Long, _
                                ByVal lNumElements As Long, _
                                ByVal k As Long) As Double
    Dim j As Long
    Dim dblResult As Double
    
    If lNumElements < 1 Then
        Exit Function
    End If
    dblResult = 1
    For j = 1 To lNumElements
        dblResult = dblResult * A_(i, j, k)
    Next j
    ProductInSlice = dblResult
End Function

Private Sub TileMemory(ByVal pTarget As LongPtr, _
                       ByVal pTargetLength As LongPtr, _
                       ByVal pSource As LongPtr, _
                       ByVal pSourceLength As LongPtr)
    Const PROCEDURE_NAME As String = "Tensor.TileMemory"
    Dim i As LongPtr
    Dim pLength As LongPtr
    
    If pTargetLength < 1 Then
        Exit Sub
    End If
    If pTarget = NULL_PTR Then
        Err.Raise 5, PROCEDURE_NAME, "Target memory address cannot be NULL."
    End If
    If pSourceLength < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Source memory length must be greater than 0."
    End If
    If pSource = NULL_PTR Then
        Err.Raise 5, PROCEDURE_NAME, "Source memory address cannot be NULL."
    End If
    pLength = MinPtr2(pTargetLength, pSourceLength)
    CopyMemory ByVal pTarget, ByVal pSource, pLength
    i = pLength
    Do While i < pTargetLength
        pLength = MinPtr2(i, pTargetLength - i)
        CopyMemory ByVal pTarget + i, ByVal pTarget, pLength
        i = i + pLength
    Loop
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
