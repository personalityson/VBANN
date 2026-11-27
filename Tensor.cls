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

#Const USE_LATE_BINDING_IN_INTERFACES = True

Private Const MAX_DIMENSIONS As Long = 8
Private Const FADF_STATIC As Integer = &H2
Private Const FADF_FIXEDSIZE As Integer = &H10
Private Const FADF_HAVEVARTYPE As Integer = &H80

Private Type SAFEARRAYBOUND
    cElements As Long
    lLBound As Long
End Type

Private Type SAFEARRAY
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound(1 To MAX_DIMENSIONS) As SAFEARRAYBOUND
End Type

Private Type SafeArrayDescriptorHeader
    Reserved(1 To 12) As Byte
    VarType As Long
End Type

Private Type SafeArrayDescriptorEx
    Header As SafeArrayDescriptorHeader
    Descriptor As SAFEARRAY
End Type

Public Enum ShapePattern
    spWildcard = -1
    spEllipsis = -2
End Enum

Public Enum ReductionType
    rtSum
    rtProduct
    rtMin
    rtMax
    rtMean
    rtMedian
    rtArgMin
    rtArgMax
End Enum

'Dimensions
Private m_lNumDimensions As Long
Private m_alShape() As Long

'Storage
Private m_oData As Storage
Private m_uSafeArray As SafeArrayDescriptorEx
Private m_uFlatSafeArray As SafeArrayDescriptorEx

Public Property Get NumDimensions() As Long
    NumDimensions = m_lNumDimensions
End Property

Public Property Get Shape() As Long()
    Shape = m_alShape
End Property

Public Property Get Size(ByVal lDimension As Long) As Long
    Const PROCEDURE_NAME As String = "Tensor.Size"

    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    Size = m_alShape(lDimension)
End Property

Public Property Get SizeBetween(ByVal lStart As Long, _
                                ByVal lEnd As Long) As Long
    Const PROCEDURE_NAME As String = "Tensor.SizeBetween"
    Dim i As Long
    Dim lResult As Long

    If lStart < 1 Then
        Err.Raise 9, PROCEDURE_NAME, "Start dimension index is out of range."
    End If
    lResult = 1
    For i = lStart To MinLng(lEnd, m_lNumDimensions)
        lResult = lResult * m_alShape(i)
    Next i
    SizeBetween = lResult
End Property

Public Property Get NumElements() As Long
    NumElements = m_oData.NumElements
End Property

Public Property Get Address() As LongPtr
    Address = m_oData.Address
End Property

Public Sub Init(ByVal vShape As Variant)
    Dim lNumDimensions As Long
    Dim alShape() As Long

    ParseShape vShape, lNumDimensions, alShape
    Clear
    m_lNumDimensions = lNumDimensions
    m_alShape = alShape
    AllocateData
    AllocateDescriptor
    AllocateFlatDescriptor
End Sub

Friend Sub Wrap(ByVal vShape As Variant, _
                ByVal oData As Storage)
    Const PROCEDURE_NAME As String = "Tensor.Wrap"
    Dim lNumDimensions As Long
    Dim lNumElements As Long
    Dim alShape() As Long

    If oData Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Storage object is required."
    End If
    ParseShape vShape, lNumDimensions, alShape, lNumElements
    If oData.NumElements <> lNumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Number of elements in the data does not match the tensor's number of elements."
    End If
    Clear
    m_lNumDimensions = lNumDimensions
    m_alShape = alShape
    Set m_oData = oData
    AllocateDescriptor
    AllocateFlatDescriptor
End Sub

Public Sub FromArray(ByRef adblArray() As Double)
    Dim i As Long
    Dim lNumDimensions As Long
    Dim alShape() As Long
    Dim oData As Storage

    lNumDimensions = GetRank(adblArray)
    If lNumDimensions < 1 Then
        Me.Init Array()
        Exit Sub
    End If
    ReDim alShape(1 To lNumDimensions)
    For i = 1 To lNumDimensions
        alShape(i) = UBound(adblArray, i) - LBound(adblArray, i) + 1
    Next i
    Set oData = New Storage
    oData.FromArray adblArray
    Me.Wrap alShape, oData
End Sub

Public Sub FromRange(ByVal oRange As Range, _
                     Optional ByVal bTranspose As Boolean)
    Dim oData As Storage

    If oRange Is Nothing Then
        Me.Init Array()
        Exit Sub
    End If
    Set oData = New Storage
    oData.FromRange oRange, bTranspose
    Me.Wrap Array(oData.NumRows, oData.NumCols), oData
End Sub

Public Sub Fill(ByVal dblValue As Double)
    TileMemory m_oData.Address, CLngPtr(m_oData.NumElements) * m_oData.ElementSize, VarPtr(dblValue), SIZEOF_DOUBLE
End Sub

Public Sub Copy(ByVal oSource As Tensor)
    Const PROCEDURE_NAME As String = "Tensor.Copy"

    If oSource Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If oSource.NumElements <> m_oData.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Source tensor must have the same number of elements as the target tensor."
    End If
    CopyMemory ByVal m_oData.Address, ByVal oSource.Address, CLngPtr(m_oData.NumElements) * m_oData.ElementSize
End Sub

Public Function Clone() As Tensor
    Set Clone = New Tensor
    Clone.Wrap m_alShape, m_oData.Clone
End Function

Public Sub AsArray(ByRef adblAlias() As Double)
    CopyMemory ByVal VarPtrArray(adblAlias), VarPtr(m_uSafeArray.Descriptor), SIZEOF_LONGPTR
End Sub

Public Sub AsFlatArray(ByRef adblAlias() As Double)
    CopyMemory ByVal VarPtrArray(adblAlias), VarPtr(m_uFlatSafeArray.Descriptor), SIZEOF_LONGPTR
End Sub

Public Function ToArray() As Double()
    Dim A_() As Double

    Me.AsArray A_
    ToArray = CVar(A_) 'Cvar returns a fresh descriptor without inherited flags or locks
End Function

Public Function View(ByVal vShape As Variant) As Tensor
    Set View = New Tensor
    View.Wrap vShape, m_oData
End Function

Public Function Flatten() As Tensor
    Set Flatten = Me.View(m_oData.NumElements)
End Function

Public Function Expand(ByVal lNumDimensions As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Expand"
    Dim i As Long
    Dim alNewShape() As Long

    If m_lNumDimensions < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor must have at least one dimension."
    End If
    If lNumDimensions < 1 Or lNumDimensions > MAX_DIMENSIONS Then
        Err.Raise 5, PROCEDURE_NAME, "Number of dimensions must be between 1 and " & MAX_DIMENSIONS & "."
    End If
    alNewShape = m_alShape
    If lNumDimensions > m_lNumDimensions Then
        ReDim Preserve alNewShape(1 To lNumDimensions)
        For i = m_lNumDimensions + 1 To lNumDimensions
            alNewShape(i) = 1
        Next i
    End If
    Set Expand = Me.View(alNewShape)
End Function

Public Function Gather(ByVal lDimension As Long, _
                       ByVal vIndices As Variant) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Gather"
    Dim i As Long
    Dim lNumIndices As Long
    Dim alIndices() As Long

    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    ParseVariantToLongArray vIndices, lNumIndices, alIndices
    For i = 1 To lNumIndices
        If alIndices(i) < 1 Or alIndices(i) > m_alShape(lDimension) Then
            Err.Raise 9, PROCEDURE_NAME, "Index is out of range."
        End If
    Next i
    Set Gather = GatherInternal(lDimension, lNumIndices, alIndices)
End Function

Public Function Slice(ByVal lDimension As Long, _
                      ByVal lStart As Long, _
                      ByVal lEnd As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Slice"

    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    If lStart < 1 Or lStart > m_alShape(lDimension) Then
        Err.Raise 9, PROCEDURE_NAME, "Start index is out of range."
    End If
    If lEnd < 1 Or lEnd > m_alShape(lDimension) Then
        Err.Raise 9, PROCEDURE_NAME, "End index is out of range."
    End If
    Set Slice = SliceInternal(lDimension, lStart, lEnd)
End Function

Public Function Tile(ByVal lDimension As Long, _
                     ByVal lNumRepeats As Long) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Tile"

    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    If lNumRepeats < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of repeats must be >= 0."
    End If
    Set Tile = TileInternal(lDimension, lNumRepeats)
End Function

Public Function Transpose(Optional ByVal lDimension1 As Long = 1, _
                          Optional ByVal lDimension2 As Long = 2) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Transpose"

    If lDimension1 < 1 Or lDimension1 > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    If lDimension2 < 1 Or lDimension2 > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    Select Case True
        Case lDimension1 = lDimension2
            Set Transpose = Me.Clone
        Case lDimension1 > lDimension2
            Set Transpose = TransposeInternal(lDimension2, lDimension1)
        Case Else
            Set Transpose = TransposeInternal(lDimension1, lDimension2)
    End Select
End Function

Public Function Reduce(ByVal lDimension As Long, _
                       ByVal eReduction As ReductionType) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Reduce"

    If lDimension < 1 Or lDimension > m_lNumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    If m_alShape(lDimension) = 0 Then
        Select Case eReduction
            Case rtMin, rtMax, rtMean, rtMedian, rtArgMin, rtArgMax
                Err.Raise 5, PROCEDURE_NAME, "Reducing over an empty dimension is not allowed."
        End Select
    End If
    Set Reduce = ReduceInternal(lDimension, eReduction)
End Function

Public Function Sum() As Double
    Sum = ReduceAll(rtSum)
End Function

Public Function Product() As Double
    Product = ReduceAll(rtProduct)
End Function

Public Function Min() As Double
    Min = ReduceAll(rtMin)
End Function

Public Function Max() As Double
    Max = ReduceAll(rtMax)
End Function

Public Function Mean() As Double
    Mean = ReduceAll(rtMean)
End Function

Public Function Median() As Double
    Median = ReduceAll(rtMedian)
End Function

Public Function ArgMin() As Long
    ArgMin = ReduceAll(rtArgMin)
End Function

Public Function ArgMax() As Long
    ArgMax = ReduceAll(rtArgMax)
End Function

Public Function ShapeEquals(ByVal vShape As Variant) As Boolean
    Dim i As Long
    Dim lOtherNumDimensions As Long
    Dim alOtherShape() As Long
    Dim lEllipsisIndex As Long
    Dim lPrefix As Long
    Dim lSuffix As Long

    ParseVariantToLongArray vShape, lOtherNumDimensions, alOtherShape
    For i = 1 To lOtherNumDimensions
        If alOtherShape(i) = spEllipsis Then
            If lEllipsisIndex > 0 Then
                Err.Raise 5, "Tensor.ShapeEquals", "Shape can contain at most one ellipsis."
            End If
            lEllipsisIndex = i
        End If
    Next i
    If lEllipsisIndex = 0 Then
        If m_lNumDimensions <> lOtherNumDimensions Then
            Exit Function
        End If
        For i = 1 To lOtherNumDimensions
            If alOtherShape(i) <> spWildcard And alOtherShape(i) <> m_alShape(i) Then
                Exit Function
            End If
        Next i
    Else
        lPrefix = lEllipsisIndex - 1
        lSuffix = lOtherNumDimensions - lEllipsisIndex
        If lPrefix + lSuffix > m_lNumDimensions Then
            Exit Function
        End If
        For i = 1 To lPrefix
            If alOtherShape(i) <> spWildcard And alOtherShape(i) <> m_alShape(i) Then
                Exit Function
            End If
        Next i
        For i = 1 To lSuffix
            If alOtherShape(lEllipsisIndex + i) <> spWildcard And alOtherShape(lEllipsisIndex + i) <> m_alShape(m_lNumDimensions - lSuffix + i) Then
                Exit Function
            End If
        Next i
    End If
    ShapeEquals = True
End Function

Public Function IsConformable(ByVal A As Tensor) As Boolean
    Const PROCEDURE_NAME As String = "Tensor.IsConformable"
    Dim i As Long

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> m_oData.NumElements Then
        Exit Function
    End If
    For i = 1 To MinLng(m_lNumDimensions, A.NumDimensions)
        If A.Size(i) <> m_alShape(i) Then
            Exit Function
        End If
    Next i
    For i = A.NumDimensions + 1 To m_lNumDimensions
        If m_alShape(i) <> 1 Then
            Exit Function
        End If
    Next i
    For i = m_lNumDimensions + 1 To A.NumDimensions
        If A.Size(i) <> 1 Then
            Exit Function
        End If
    Next i
    IsConformable = True
End Function

Private Sub Class_Terminate()
    Clear
End Sub

Private Sub Clear()
    m_lNumDimensions = 0
    Erase m_alShape
    Set m_oData = Nothing
End Sub

Private Sub ParseShape(ByVal vShape As Variant, _
                       ByRef lNumDimensions As Long, _
                       ByRef alShape() As Long, _
                       Optional ByRef lNumElements As Long)
    Const PROCEDURE_NAME As String = "Tensor.ParseShape"
    Dim i As Long
    
    ParseVariantToLongArray vShape, lNumDimensions, alShape
    If lNumDimensions > MAX_DIMENSIONS Then
        Err.Raise 5, PROCEDURE_NAME, "Number of dimensions exceeds the maximum allowed (" & MAX_DIMENSIONS & ")."
    End If
    lNumElements = Sgn(lNumDimensions) 'Empty tensor has 0 elements, not a scalar
    For i = 1 To lNumDimensions
        If alShape(i) < 0 Then
            Err.Raise 5, PROCEDURE_NAME, "Dimension size must be >= 0."
        End If
        lNumElements = lNumElements * alShape(i)
    Next i
End Sub

Private Sub AllocateData()
    Dim lNumRows As Long
    Dim lNumCols As Long

    Set m_oData = New Storage
    If m_lNumDimensions > 0 Then
        lNumRows = Me.SizeBetween(1, m_lNumDimensions - 1)
        lNumCols = m_alShape(m_lNumDimensions)
        m_oData.Init lNumRows, lNumCols
    End If
End Sub

Private Sub AllocateDescriptor()
    Dim i As Long

    With m_uSafeArray
        .Header.VarType = m_oData.VarType
        .Descriptor.cDims = m_lNumDimensions
        .Descriptor.fFeatures = FADF_STATIC Or FADF_FIXEDSIZE Or FADF_HAVEVARTYPE
        .Descriptor.cbElements = m_oData.ElementSize
        .Descriptor.cLocks = 1
        .Descriptor.pvData = m_oData.Address
        For i = 1 To m_lNumDimensions
            .Descriptor.rgsabound(i).cElements = m_alShape(m_lNumDimensions - i + 1)
            .Descriptor.rgsabound(i).lLBound = 1
        Next i
    End With
End Sub

Private Sub AllocateFlatDescriptor()
    With m_uFlatSafeArray
        .Header.VarType = m_oData.VarType
        .Descriptor.cDims = 1
        .Descriptor.fFeatures = FADF_STATIC Or FADF_FIXEDSIZE Or FADF_HAVEVARTYPE
        .Descriptor.cbElements = m_oData.ElementSize
        .Descriptor.cLocks = 1
        .Descriptor.pvData = m_oData.Address
        .Descriptor.rgsabound(1).cElements = m_oData.NumElements
        .Descriptor.rgsabound(1).lLBound = 1
    End With
End Sub

Private Function GatherInternal(ByVal lDimension As Long, _
                                ByVal lNumIndices As Long, _
                                ByRef alIndices() As Long) As Tensor
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim lOldSize As Long
    Dim lNewSize As Long
    Dim lBefore As Long
    Dim lAfter As Long
    Dim pStrideBytes As LongPtr
    Dim alNewShape() As Long
    Dim oResult As Tensor

    lOldSize = m_alShape(lDimension)
    lNewSize = lNumIndices
    alNewShape = m_alShape
    alNewShape(lDimension) = lNewSize
    Set oResult = New Tensor
    oResult.Init alNewShape
    lBefore = Me.SizeBetween(1, lDimension - 1)
    lAfter = Me.SizeBetween(lDimension + 1, m_lNumDimensions)
    pStrideBytes = CLngPtr(lBefore) * m_oData.ElementSize
    For i = 1 To lAfter
        j = 1
        Do While j <= lNewSize
            For k = j + 1 To lNewSize
                If alIndices(k) <> alIndices(k - 1) + 1 Then Exit For
            Next k
            CopyMemory ByVal oResult.Address + ((i - 1) * lNewSize + (j - 1)) * pStrideBytes, _
                       ByVal m_oData.Address + ((i - 1) * lOldSize + (alIndices(j) - 1)) * pStrideBytes, _
                       (k - j) * pStrideBytes
            j = k
        Loop
    Next i
    Set GatherInternal = oResult
End Function

Private Function SliceInternal(ByVal lDimension As Long, _
                               ByVal lStart As Long, _
                               ByVal lEnd As Long) As Tensor
    Dim i As Long
    Dim lOldSize As Long
    Dim lNewSize As Long
    Dim lBefore As Long
    Dim lAfter As Long
    Dim pStrideBytes As LongPtr
    Dim alNewShape() As Long
    Dim oResult As Tensor

    lOldSize = m_alShape(lDimension)
    lNewSize = MaxLng(0, lEnd - lStart + 1)
    alNewShape = m_alShape
    alNewShape(lDimension) = lNewSize
    Set oResult = New Tensor
    oResult.Init alNewShape
    lBefore = Me.SizeBetween(1, lDimension - 1)
    lAfter = Me.SizeBetween(lDimension + 1, m_lNumDimensions)
    pStrideBytes = CLngPtr(lBefore) * m_oData.ElementSize
    For i = 1 To lAfter
        CopyMemory ByVal oResult.Address + ((i - 1) * lNewSize) * pStrideBytes, _
                   ByVal m_oData.Address + ((i - 1) * lOldSize + (lStart - 1)) * pStrideBytes, _
                   lNewSize * pStrideBytes
    Next i
    Set SliceInternal = oResult
End Function

Private Function TileInternal(ByVal lDimension As Long, _
                              ByVal lNumRepeats As Long) As Tensor
    Dim i As Long
    Dim lOldSize As Long
    Dim lNewSize As Long
    Dim lBefore As Long
    Dim lAfter As Long
    Dim pStrideBytes As LongPtr
    Dim alNewShape() As Long
    Dim oResult As Tensor

    lOldSize = m_alShape(lDimension)
    lNewSize = MaxLng(0, lNumRepeats * lOldSize)
    alNewShape = m_alShape
    alNewShape(lDimension) = lNewSize
    Set oResult = New Tensor
    oResult.Init alNewShape
    lBefore = Me.SizeBetween(1, lDimension - 1)
    lAfter = Me.SizeBetween(lDimension + 1, m_lNumDimensions)
    pStrideBytes = CLngPtr(lBefore) * m_oData.ElementSize
    For i = 1 To lAfter
        TileMemory oResult.Address + ((i - 1) * lNewSize) * pStrideBytes, _
                   lNewSize * pStrideBytes, _
                   m_oData.Address + ((i - 1) * lOldSize) * pStrideBytes, _
                   lOldSize * pStrideBytes
    Next i
    Set TileInternal = oResult
End Function

Private Function TransposeInternal(ByVal lDimension1 As Long, _
                                   ByVal lDimension2 As Long) As Tensor
    Const COPY_THRESHOLD As Long = 40
    Dim i As Long
    Dim j As Long
    Dim P As Long
    Dim q As Long
    Dim s As Long
    Dim M As Long
    Dim n As Long
    Dim lBefore As Long
    Dim lBetween As Long
    Dim lAfter As Long
    Dim pStrideBytes As LongPtr
    Dim alNewShape() As Long
    Dim A_() As Double
    Dim R_() As Double
    Dim A As Tensor
    Dim R As Tensor
    Dim oResult As Tensor
    
    M = m_alShape(lDimension1)
    n = m_alShape(lDimension2)
    alNewShape = m_alShape
    alNewShape(lDimension1) = n
    alNewShape(lDimension2) = M
    Set oResult = New Tensor
    oResult.Init alNewShape
    lBefore = Me.SizeBetween(1, lDimension1 - 1)
    lBetween = Me.SizeBetween(lDimension1 + 1, lDimension2 - 1)
    lAfter = Me.SizeBetween(lDimension2 + 1, m_lNumDimensions)
    pStrideBytes = CLngPtr(lBefore) * m_oData.ElementSize
    Set A = Me.View(Array(lBefore, M, lBetween, n, lAfter))
    Set R = oResult.View(Array(lBefore, n, lBetween, M, lAfter))
    A.AsArray A_
    R.AsArray R_
    If lBefore >= COPY_THRESHOLD Then
        For s = 1 To lAfter
            For i = 1 To M
                For q = 1 To lBetween
                    For j = 1 To n
                        CopyMemory R_(1, j, q, i, s), A_(1, i, q, j, s), pStrideBytes
                    Next j
                Next q
            Next i
        Next s
    Else
        For s = 1 To lAfter
            For i = 1 To M
                For q = 1 To lBetween
                    For j = 1 To n
                        For P = 1 To lBefore
                            R_(P, j, q, i, s) = A_(P, i, q, j, s)
                        Next P
                    Next j
                Next q
            Next i
        Next s
    End If
    Set TransposeInternal = oResult
End Function

Private Function ReduceInternal(ByVal lDimension As Long, _
                                ByVal eReduction As ReductionType) As Tensor
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim lOldSize As Long
    Dim lNewSize As Long
    Dim lBefore As Long
    Dim lAfter As Long
    Dim lIndex As Long
    Dim dblAcc As Double
    Dim alNewShape() As Long
    Dim adblValues() As Double
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
    oResult.Init alNewShape
    lBefore = Me.SizeBetween(1, lDimension - 1)
    lAfter = Me.SizeBetween(lDimension + 1, m_lNumDimensions)
    Set A = Me.View(Array(lBefore, lOldSize, lAfter))
    Set R = oResult.View(Array(lBefore, lNewSize, lAfter))
    A.AsArray A_
    R.AsArray R_
    Select Case eReduction
        Case rtSum
            For k = 1 To lAfter
                For i = 1 To lBefore
                    dblAcc = 0
                    For j = 1 To lOldSize
                        dblAcc = dblAcc + A_(i, j, k)
                    Next j
                    R_(i, 1, k) = dblAcc
                Next i
            Next k
        Case rtProduct
            For k = 1 To lAfter
                For i = 1 To lBefore
                    dblAcc = 1
                    For j = 1 To lOldSize
                        dblAcc = dblAcc * A_(i, j, k)
                    Next j
                    R_(i, 1, k) = dblAcc
                Next i
            Next k
        Case rtMin
            For k = 1 To lAfter
                For i = 1 To lBefore
                    dblAcc = A_(i, 1, k)
                    For j = 2 To lOldSize
                        dblAcc = MinDbl(dblAcc, A_(i, j, k))
                    Next j
                    R_(i, 1, k) = dblAcc
                Next i
            Next k
        Case rtMax
            For k = 1 To lAfter
                For i = 1 To lBefore
                    dblAcc = A_(i, 1, k)
                    For j = 2 To lOldSize
                        dblAcc = MaxDbl(dblAcc, A_(i, j, k))
                    Next j
                    R_(i, 1, k) = dblAcc
                Next i
            Next k
        Case rtMean
            For k = 1 To lAfter
                For i = 1 To lBefore
                    dblAcc = 0
                    For j = 1 To lOldSize
                        dblAcc = dblAcc + A_(i, j, k)
                    Next j
                    R_(i, 1, k) = dblAcc / lOldSize
                Next i
            Next k
        Case rtMedian
            For k = 1 To lAfter
                For i = 1 To lBefore
                    ReDim adblValues(1 To lOldSize)
                    For j = 1 To lOldSize
                        adblValues(j) = A_(i, j, k)
                    Next j
                    R_(i, 1, k) = MedianInPlace(adblValues, lOldSize)
                Next i
            Next k
        Case rtArgMin
            For k = 1 To lAfter
                For i = 1 To lBefore
                    lIndex = 1
                    dblAcc = A_(i, 1, k)
                    For j = 2 To lOldSize
                        If A_(i, j, k) < dblAcc Then
                            dblAcc = A_(i, j, k)
                            lIndex = j
                        End If
                    Next j
                    R_(i, 1, k) = lIndex
                Next i
            Next k
        Case rtArgMax
            For k = 1 To lAfter
                For i = 1 To lBefore
                    lIndex = 1
                    dblAcc = A_(i, 1, k)
                    For j = 2 To lOldSize
                        If A_(i, j, k) > dblAcc Then
                            dblAcc = A_(i, j, k)
                            lIndex = j
                        End If
                    Next j
                    R_(i, 1, k) = lIndex
                Next i
            Next k
    End Select
    Set ReduceInternal = oResult
End Function

Private Function ReduceAll(ByVal eReduction As ReductionType) As Double
    ReduceAll = Me.Flatten.Reduce(1, eReduction).ToArray()(1)
End Function

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "Tensor.Save"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteLongArray m_alShape
        .WriteObject m_oData
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "Tensor.Load"
    Dim alShape() As Long
    Dim oData As Storage

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        alShape = .ReadLongArray()
        Set oData = .ReadObject()
    End With
    If oData Is Nothing Then
        Me.Init alShape
    Else
        Me.Wrap alShape, oData
    End If
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

