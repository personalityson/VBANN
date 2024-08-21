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

Private Enum ReductionType
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
    AllocateDescriptor
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
    Dim lOldDimensionSize As Long
    Dim lNewDimensionSize As Long
    Dim alNewShape() As Long
    Dim lInnerStride As Long
    Dim lOuterStride As Long
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
    lOldDimensionSize = m_alShape(lDimension)
    lNewDimensionSize = MaxLng2(0, lEnd - lStart + 1)
    alNewShape = m_alShape
    alNewShape(lDimension) = lNewDimensionSize
    Set oResult = New Tensor
    oResult.Resize alNewShape
    lInnerStride = m_alInnerStride(lDimension)
    lOuterStride = m_alOuterStride(lDimension)
    For i = 1 To lOuterStride
        CopyMemory ByVal oResult.Address + ((i - 1) * lNewDimensionSize + (1 - 1)) * lInnerStride * m_oData.ElementSize, _
                   ByVal m_oData.Address + ((i - 1) * lOldDimensionSize + (lStart - 1)) * lInnerStride * m_oData.ElementSize, _
                   lNewDimensionSize * lInnerStride * m_oData.ElementSize
    Next i
    Set Slice = oResult
End Function

Public Function Gather(ByVal lDimension As Long, _
                       ByVal vIndices As Variant) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Gather"
    Dim lNumIndices As Long
    Dim alIndices() As Long
    Dim lOldDimensionSize As Long
    Dim lNewDimensionSize As Long
    Dim alNewShape() As Long
    Dim lInnerStride As Long
    Dim lOuterStride As Long
    Dim i As Long
    Dim j As Long
    Dim lIndex As Long
    Dim oResult As Tensor
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    ParseVariantToLongArray vIndices, lNumIndices, alIndices
    For i = 1 To lNumIndices
        lIndex = alIndices(i)
        If lIndex < 1 Or lIndex > m_alShape(lDimension) Then
            Err.Raise 9, PROCEDURE_NAME, "Index must be between 1 and the size of the dimension."
        End If
    Next i
    lOldDimensionSize = m_alShape(lDimension)
    lNewDimensionSize = MinLng2(lOldDimensionSize, lNumIndices)
    alNewShape = m_alShape
    alNewShape(lDimension) = lNewDimensionSize
    Set oResult = New Tensor
    oResult.Resize alNewShape
    lInnerStride = m_alInnerStride(lDimension)
    lOuterStride = m_alOuterStride(lDimension)
    For i = 1 To lOuterStride
        For j = 1 To lNewDimensionSize
            lIndex = alIndices(j)
            CopyMemory ByVal oResult.Address + ((i - 1) * lNewDimensionSize + (j - 1)) * lInnerStride * m_oData.ElementSize, _
                       ByVal m_oData.Address + ((i - 1) * lOldDimensionSize + (lIndex - 1)) * lInnerStride * m_oData.ElementSize, _
                       lInnerStride * m_oData.ElementSize
        Next j
    Next i
    Set Gather = oResult
End Function

Public Function ReduceSum(ByVal lDimension As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.ReduceSum"
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    Set ReduceSum = ReduceWithDimension(lDimension, rdcSum)
End Function

Public Function ReduceMean(ByVal lDimension As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.ReduceMean"
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    Set ReduceMean = ReduceWithDimension(lDimension, rdcMean)
End Function

Public Function ReduceMin(ByVal lDimension As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.ReduceMin"
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    Set ReduceMin = ReduceWithDimension(lDimension, rdcMin)
End Function

Public Function ReduceMax(ByVal lDimension As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.ReduceMax"
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    Set ReduceMax = ReduceWithDimension(lDimension, rdcMax)
End Function

Public Function ReduceProduct(ByVal lDimension As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.ReduceProduct"
    
    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension must be between 1 and the number of dimensions."
    End If
    Set ReduceProduct = ReduceWithDimension(lDimension, rdcProduct)
End Function

Public Function Sum() As Double
    Sum = SumInStride(1, 1, m_lNumElements)
End Function

Public Function Mean() As Double
    Mean = MeanInStride(1, 1, m_lNumElements)
End Function

Public Function Min() As Double
    Min = MinInStride(1, 1, m_lNumElements)
End Function

Public Function Max() As Double
    Max = MaxInStride(1, 1, m_lNumElements)
End Function

Public Function Product() As Double
    Product = ProductInStride(1, 1, m_lNumElements)
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
    
    CreateAlias A_
    ToArray = A_
    RemoveAlias A_
End Function

Public Sub FillConstant(ByVal dblValue As Double)
    Dim i As Long
    Dim lLength As Long
    
    If m_lNumElements < 1 Then
        Exit Sub
    End If
    CopyMemory ByVal m_oData.Address, dblValue, m_oData.ElementSize
    i = 1
    Do While i < m_lNumElements
        lLength = MinLng2(i, m_lNumElements - i)
        CopyMemory ByVal m_oData.Address + i * m_oData.ElementSize, ByVal m_oData.Address, lLength * m_oData.ElementSize
        i = i + lLength
    Loop
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

Private Function ReduceWithDimension(ByVal lDimension As Long, _
                                     ByVal eReduction As ReductionType) As Tensor
    Dim lOldDimensionSize As Long
    Dim lNewDimensionSize As Long
    Dim alNewShape() As Long
    Dim lInnerStride As Long
    Dim lOuterStride As Long
    Dim lIndexA As Long
    Dim lIndexR As Long
    Dim i As Long
    Dim k As Long
    Dim R_() As Double
    Dim oResult As Tensor
    
    lOldDimensionSize = m_alShape(lDimension)
    lNewDimensionSize = 1
    alNewShape = m_alShape
    alNewShape(lDimension) = lNewDimensionSize
    Set oResult = New Tensor
    oResult.Resize alNewShape
    oResult.Flatten.CreateAlias R_
    lInnerStride = m_alInnerStride(lDimension)
    lOuterStride = m_alOuterStride(lDimension)
    For i = 1 To lInnerStride
        For k = 1 To lOuterStride
            lIndexA = (k - 1) * lOldDimensionSize * lInnerStride + i
            lIndexR = (k - 1) * lNewDimensionSize * lInnerStride + i
            Select Case eReduction
                Case rdcSum
                    R_(lIndexR) = SumInStride(lIndexA, lInnerStride, lOldDimensionSize)
                Case rdcMean
                    R_(lIndexR) = MeanInStride(lIndexA, lInnerStride, lOldDimensionSize)
                Case rdcMin
                    R_(lIndexR) = MinInStride(lIndexA, lInnerStride, lOldDimensionSize)
                Case rdcMax
                    R_(lIndexR) = MaxInStride(lIndexA, lInnerStride, lOldDimensionSize)
                Case rdcProduct
                    R_(lIndexR) = ProductInStride(lIndexA, lInnerStride, lOldDimensionSize)
            End Select
        Next k
    Next i
    oResult.Flatten.RemoveAlias R_
    Set ReduceWithDimension = oResult
End Function

Private Function SumInStride(ByVal lStart As Long, _
                             ByVal lStride As Long, _
                             ByVal lNumElements As Long) As Double
    Dim j As Long
    Dim A_() As Double
    Dim dblResult As Double

    If lNumElements < 1 Then
        Exit Function
    End If
    Me.Flatten.CreateAlias A_
    dblResult = 0
    For j = 1 To lNumElements
         dblResult = dblResult + A_(lStart + (j - 1) * lStride)
    Next j
    Me.Flatten.RemoveAlias A_
    SumInStride = dblResult
End Function

Private Function MeanInStride(ByVal lStart As Long, _
                              ByVal lStride As Long, _
                              ByVal lNumElements As Long) As Double
    If lNumElements < 1 Then
        Exit Function
    End If
    MeanInStride = SumInStride(lStart, lStride, lNumElements) / lNumElements
End Function

Private Function MinInStride(ByVal lStart As Long, _
                             ByVal lStride As Long, _
                             ByVal lNumElements As Long) As Double
    Const PROCEDURE_NAME As String = "Tensor.MinInStride"
    Dim i As Long
    Dim A_() As Double
    Dim dblResult As Double
    
    If lNumElements < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of elements must be greater than 0."
    End If
    Me.Flatten.CreateAlias A_
    dblResult = A_(lStart)
    For i = 2 To lNumElements
         dblResult = MinLng2(dblResult, A_(lStart + (i - 1) * lStride))
    Next i
    Me.Flatten.RemoveAlias A_
    MinInStride = dblResult
End Function

Private Function MaxInStride(ByVal lStart As Long, _
                             ByVal lStride As Long, _
                             ByVal lNumElements As Long) As Double
    Const PROCEDURE_NAME As String = "Tensor.MaxInStride"
    Dim i As Long
    Dim A_() As Double
    Dim dblResult As Double
    
    If lNumElements < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of elements must be greater than 0."
    End If
    Me.Flatten.CreateAlias A_
    dblResult = A_(lStart)
    For i = 2 To lNumElements
         dblResult = MaxLng2(dblResult, A_(lStart + (i - 1) * lStride))
    Next i
    Me.Flatten.RemoveAlias A_
    MaxInStride = dblResult
End Function

Private Function ProductInStride(ByVal lStart As Long, _
                                 ByVal lStride As Long, _
                                 ByVal lNumElements As Long) As Double
    Dim i As Long
    Dim A_() As Double
    Dim dblResult As Double
    
    If lNumElements < 1 Then
        Exit Function
    End If
    Me.Flatten.CreateAlias A_
    dblResult = 1
    For i = 1 To lNumElements
        dblResult = dblResult * A_(lStart + (i - 1) * lStride)
    Next i
    Me.Flatten.RemoveAlias A_
    ProductInStride = dblResult
End Function

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
