Attribute VB_Name = "TensorOps"
Option Explicit

Private Const OPENBLAS_PATH As String = "C:\Users\hello\OneDrive\Documents\VBANN\libopenblas.dll"

Private m_vIsBlasAvailable As Variant

Private Declare PtrSafe Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" (ByVal lpPathName As String) As Long

Private Declare PtrSafe Sub dscal Lib "libopenblas.dll" (ByRef n As Long, _
                                                         ByRef alpha As Double, _
                                                         ByVal X As LongPtr, _
                                                         ByRef incX As Long)

Private Declare PtrSafe Sub daxpby Lib "libopenblas.dll" (ByRef n As Long, _
                                                          ByRef alpha As Double, _
                                                          ByVal X As LongPtr, _
                                                          ByRef incX As Long, _
                                                          ByRef beta As Double, _
                                                          ByVal Y As LongPtr, _
                                                          ByRef incY As Long)

Private Declare PtrSafe Sub dgemm Lib "libopenblas.dll" (ByVal transA As String, _
                                                         ByVal transB As String, _
                                                         ByRef m As Long, _
                                                         ByRef n As Long, _
                                                         ByRef k As Long, _
                                                         ByRef alpha As Double, _
                                                         ByVal a As LongPtr, _
                                                         ByRef ldA As Long, _
                                                         ByVal b As LongPtr, _
                                                         ByRef ldB As Long, _
                                                         ByRef beta As Double, _
                                                         ByVal C As LongPtr, _
                                                         ByRef ldC As Long)

Public Function IsBlasAvailable() As Boolean
    If IsEmpty(m_vIsBlasAvailable) Then
        If Fso.FileExists(OPENBLAS_PATH) Then
            m_vIsBlasAvailable = SetDllDirectory(Fso.GetParentFolderName(OPENBLAS_PATH)) <> 0
        Else
            m_vIsBlasAvailable = False
        End If
    End If
    IsBlasAvailable = m_vIsBlasAvailable
End Function

'VecAdd                 Y = A + B
'VecAdd_I               A = A + B (In-place)
'VecAddC                Y = A + scalar
'VecSub                 Y = A - B
'VecSubC                Y = A - scalar
'VecSubCRev             Y = scalar - A
'VecMul                 Y = A * B
'VecMulC                Y = A * scalar
'VecDiv                 Y = A / B
'VecDivC                Y = A / scalar
'VecDivCRev             Y = scalar / A
'VecDivSqrtAddC         Y = A / Sqrt(B + scalar)
'VecAbs                 Y = Abs(A)
'VecSign                Y = Sign(A)
'VecPow2                Y = A^2
'VecSqrt                Y = Sqrt(A)
'VecExp                 Y = Exp(A)
'VecLog                 Y = Log(A)
'VecLeakyReLU           Y = If A > 0 Then A Else dblNegativeSlope * A
'VecLeakyReLUDerivative Y = If A > 0 Then 1 Else dblNegativeSlope
'VecSigmoid             Y = 1 / (1 + Exp(-A))
'VecSigmoidDerivative   Y = A * (1 - A)
'VecTanh                Y = Tanh(A)
'VecTanhDerivative      Y = 1 - A^2
'VecLinComb             Y = alpha * A + beta * B
'VecLinComb_I           A = alpha * A + beta * B (In-place)
'MatMul                 Y = A * B
'MatMul_I               C = C + A * B (In-place)

'Y = A + B
Public Function VecAdd(ByVal a As Tensor, _
                       ByVal b As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecAdd"

    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If b Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If a.NumElements <> b.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If IsBlasAvailable() Then
        Set VecAdd = VecLinCombBlas(1, a, 1, b)
    Else
        Set VecAdd = VecAddNaive(a, b)
    End If
End Function

'A = A + B
Public Sub VecAdd_I(ByVal a As Tensor, _
                    ByVal b As Tensor)
    Const PROCEDURE_NAME As String = "TensorOps.VecAdd_I"

    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If b Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If a.NumElements <> b.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If IsBlasAvailable() Then
        VecLinCombBlas_I 1, a, 1, b
    Else
        VecAddNaive_I a, b
    End If
End Sub

'Y = A + scalar
Public Function VecAddC(ByVal a As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecAddC"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecAddC = VecLinCombBlas(1, a, 1, Full(a.Shape, dblScalar))
    Else
        Set VecAddC = VecAddCNaive(a, dblScalar)
    End If
End Function

'Y = A - B
Public Function VecSub(ByVal a As Tensor, _
                       ByVal b As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSub"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If b Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If a.NumElements <> b.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If IsBlasAvailable() Then
        Set VecSub = VecLinCombBlas(1, a, -1, b)
    Else
        Set VecSub = VecSubNaive(a, b)
    End If
End Function

'Y = A - scalar
Public Function VecSubC(ByVal a As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSubC"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecSubC = VecLinCombBlas(1, a, -1, Full(a.Shape, dblScalar))
    Else
        Set VecSubC = VecSubCNaive(a, dblScalar)
    End If
End Function

'Y = scalar - A
Public Function VecSubCRev(ByVal a As Tensor, _
                           ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSubCRev"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecSubCRev = VecLinCombBlas(1, Full(a.Shape, dblScalar), -1, a)
    Else
        Set VecSubCRev = VecSubCRevNaive(a, dblScalar)
    End If
End Function

'Y = A * B
Public Function VecMul(ByVal a As Tensor, _
                       ByVal b As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecMul"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If b Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If a.NumElements <> b.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    Set VecMul = VecMulNaive(a, b)
End Function

'Y = A * scalar
Public Function VecMulC(ByVal a As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecMulC"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecMulC = VecMulCBlas(a, dblScalar)
    Else
        Set VecMulC = VecMulCNaive(a, dblScalar)
    End If
End Function

'Y = A / B
Public Function VecDiv(ByVal a As Tensor, _
                       ByVal b As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecDiv"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If b Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If a.NumElements <> b.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    Set VecDiv = VecDivNaive(a, b)
End Function

'Y = A / scalar
Public Function VecDivC(ByVal a As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Set VecDivC = VecMulC(a, 1 / dblScalar)
End Function

'Y = scalar / A
Public Function VecDivCRev(ByVal a As Tensor, _
                           ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecDivCRev"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecDivCRev = VecDivCRevNaive(a, dblScalar)
End Function

'Y = A / (Sqrt(B) + scalar)
Public Function VecDivSqrtAddC(ByVal a As Tensor, _
                               ByVal b As Tensor, _
                               ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecDivSqrtAddC"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If b Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If a.NumElements <> b.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    Set VecDivSqrtAddC = VecDivSqrtAddCNaive(a, b, dblScalar)
End Function

'Y = Abs(A)
Public Function VecAbs(ByVal a As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecAbs"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecAbs = VecAbsNaive(a)
End Function

'Y = Sign(A)
Public Function VecSign(ByVal a As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSign"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSign = VecSignNaive(a)
End Function

'Y = A^2
Public Function VecPow2(ByVal a As Tensor) As Tensor
    Set VecPow2 = VecMul(a, a)
End Function

'Y = Sqrt(A)
Public Function VecSqrt(ByVal a As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSqrt"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSqrt = VecSqrtNaive(a)
End Function

'Y = Exp(A)
Public Function VecExp(ByVal a As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecExp"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecExp = VecExpNaive(a)
End Function

'Y = Log(A)
Public Function VecLog(ByVal a As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLog"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecLog = VecLogNaive(a)
End Function

'Y = If A > 0 Then A Else dblNegativeSlope * A
Public Function VecLeakyReLU(ByVal a As Tensor, _
                             ByVal dblNegativeSlope As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLeakyReLU"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecLeakyReLU = VecLeakyReLUNaive(a, dblNegativeSlope)
End Function

'Y = If A > 0 Then 1 Else dblNegativeSlope
Public Function VecLeakyReLUDerivative(ByVal a As Tensor, _
                                       ByVal dblNegativeSlope As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLeakyReLUDerivative"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecLeakyReLUDerivative = VecLeakyReLUDerivativeNaive(a, dblNegativeSlope)
End Function

'Y = 1 / (1 + Exp(-A))
Public Function VecSigmoid(ByVal a As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSigmoid"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSigmoid = VecSigmoidNaive(a)
End Function

'Y = A * (1 - A)
Public Function VecSigmoidDerivative(ByVal a As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSigmoidDerivative"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSigmoidDerivative = VecSigmoidDerivativeNaive(a)
End Function

'Y = Tanh(A)
Public Function VecTanh(ByVal a As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecTanh"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecTanh = VecTanhNaive(a)
End Function

'Y = 1 - A^2
Public Function VecTanhDerivative(ByVal a As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecTanhDerivative"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecTanhDerivative = VecTanhDerivativeNaive(a)
End Function

'Y = alpha * A + beta * B
Public Function VecLinComb(ByVal dblAlpha As Double, _
                           ByVal a As Tensor, _
                           ByVal dblBeta As Double, _
                           ByVal b As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLinComb"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If b Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If a.NumElements <> b.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If IsBlasAvailable() Then
        Set VecLinComb = VecLinCombBlas(dblAlpha, a, dblBeta, b)
    Else
        Set VecLinComb = VecLinCombNaive(dblAlpha, a, dblBeta, b)
    End If
End Function

'A = alpha * A + beta * B
Public Sub VecLinComb_I(ByVal dblAlpha As Double, _
                        ByVal a As Tensor, _
                        ByVal dblBeta As Double, _
                        ByVal b As Tensor)
    Const PROCEDURE_NAME As String = "TensorOps.VecLinComb_I"
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If b Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If a.NumElements <> b.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If IsBlasAvailable() Then
        VecLinCombBlas_I dblAlpha, a, dblBeta, b
    Else
        VecLinCombNaive_I dblAlpha, a, dblBeta, b
    End If
End Sub

'Y = A * B
Public Function MatMul(ByVal a As Tensor, _
                       ByVal b As Tensor, _
                       Optional ByVal bTransposeA As Boolean, _
                       Optional ByVal bTransposeB As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.MatMul"
    Dim lNumColsA As Long
    Dim lNumRowsB As Long
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If b Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If a.NumDimensions < 1 Or a.NumDimensions > 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor A must have 1 or 2 dimensions."
    End If
    If b.NumDimensions < 1 Or b.NumDimensions > 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor B must have 1 or 2 dimensions."
    End If
    If a.NumDimensions = 1 Then
        Set a = a.View(Array(1, a.Size(1)))
    End If
    If b.NumDimensions = 1 Then
        Set b = b.View(Array(b.Size(1), 1))
    End If
    lNumColsA = IIf(bTransposeA, a.Size(1), a.Size(2))
    lNumRowsB = IIf(bTransposeB, b.Size(2), b.Size(1))
    If lNumColsA <> lNumRowsB Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B are incompatible for matrix multiplication."
    End If
    If IsBlasAvailable() Then
        Set MatMul = MatMulBlas(a, b, bTransposeA, bTransposeB)
    Else
        Set MatMul = MatMulNaive(a, b, bTransposeA, bTransposeB)
    End If
End Function

'C = C + A * B
Public Sub MatMul_I(ByVal C As Tensor, _
                    ByVal a As Tensor, _
                    ByVal b As Tensor, _
                    Optional ByVal bTransposeA As Boolean, _
                    Optional ByVal bTransposeB As Boolean)
    Const PROCEDURE_NAME As String = "TensorOps.MatMul_I"
    Dim lNumRowsA As Long
    Dim lNumColsA As Long
    Dim lNumRowsB As Long
    Dim lNumColsB As Long
    
    If a Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If b Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If C Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If a.NumDimensions < 1 Or a.NumDimensions > 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor A must have 1 or 2 dimensions."
    End If
    If b.NumDimensions < 1 Or b.NumDimensions > 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor B must have 1 or 2 dimensions."
    End If
    If C.NumDimensions < 1 Or C.NumDimensions > 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor C must have 1 or 2 dimensions."
    End If
    If a.NumDimensions = 1 Then
        Set a = a.View(Array(1, a.Size(1)))
    End If
    If b.NumDimensions = 1 Then
        Set b = b.View(Array(b.Size(1), 1))
    End If
    lNumRowsA = IIf(bTransposeA, a.Size(2), a.Size(1))
    lNumColsA = IIf(bTransposeA, a.Size(1), a.Size(2))
    lNumRowsB = IIf(bTransposeB, b.Size(2), b.Size(1))
    lNumColsB = IIf(bTransposeB, b.Size(1), b.Size(2))
    If lNumColsA <> lNumRowsB Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B are incompatible for matrix multiplication."
    End If
    If C.NumDimensions = 1 Then
        Select Case 1
            Case lNumRowsA
                Set C = C.View(Array(1, C.Size(1)))
            Case lNumColsB
                Set C = C.View(Array(C.Size(1), 1))
        End Select
    End If
    If Not C.ShapeEquals(Array(lNumRowsA, lNumColsB)) Then
        Err.Raise 5, PROCEDURE_NAME, "Output tensor does not match the expected shape for matrix multiplication."
    End If
    If IsBlasAvailable() Then
        MatMulBlas_I C, a, b, bTransposeA, bTransposeB
    Else
        MatMulNaive_I C, a, b, bTransposeA, bTransposeB
    End If
End Sub

Private Function SafeSigmoid(ByVal dblValue As Double) As Double
    If dblValue < -DOUBLE_MAX_LOG Then
        SafeSigmoid = 0
    Else
        SafeSigmoid = 1 / (1 + Exp(-dblValue))
    End If
End Function

Private Function SafeTanh(ByVal dblValue As Double) As Double
    Dim dblExp As Double
    
    dblValue = 2 * dblValue
    If dblValue > DOUBLE_MAX_LOG Then
        SafeTanh = 1
    Else
        dblExp = Exp(dblValue)
        SafeTanh = (dblExp - 1) / (dblExp + 1)
    End If
End Function

Private Function VecAddNaive(ByVal a As Tensor, _
                             ByVal b As Tensor) As Tensor
    Set a = a.Clone
    VecAddNaive_I a, b
    Set VecAddNaive = a
End Function

Private Sub VecAddNaive_I(ByVal a As Tensor, _
                          ByVal b As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double

    a.Flatten.CreateAlias A_
    b.Flatten.CreateAlias B_
    For i = 1 To a.NumElements
        A_(i) = A_(i) + B_(i)
    Next i
    a.Flatten.RemoveAlias A_
    b.Flatten.RemoveAlias B_
End Sub

Private Function VecAddCNaive(ByVal a As Tensor, _
                              ByVal dblScalar As Double) As Tensor
    Set a = a.Clone
    VecAddCNaive_I a, dblScalar
    Set VecAddCNaive = a
End Function

Private Sub VecAddCNaive_I(ByVal a As Tensor, _
                           ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = A_(i) + dblScalar
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecSubNaive(ByVal a As Tensor, _
                             ByVal b As Tensor) As Tensor
    Set a = a.Clone
    VecSubNaive_I a, b
    Set VecSubNaive = a
End Function

Private Sub VecSubNaive_I(ByVal a As Tensor, _
                          ByVal b As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    a.Flatten.CreateAlias A_
    b.Flatten.CreateAlias B_
    For i = 1 To a.NumElements
        A_(i) = A_(i) - B_(i)
    Next i
    a.Flatten.RemoveAlias A_
    b.Flatten.RemoveAlias B_
End Sub

Private Function VecSubCNaive(ByVal a As Tensor, _
                              ByVal dblScalar As Double) As Tensor
    Set a = a.Clone
    VecSubCNaive_I a, dblScalar
    Set VecSubCNaive = a
End Function

Private Sub VecSubCNaive_I(ByVal a As Tensor, _
                           ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = A_(i) - dblScalar
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecSubCRevNaive(ByVal a As Tensor, _
                                 ByVal dblScalar As Double) As Tensor
    Set a = a.Clone
    VecSubCRevNaive_I a, dblScalar
    Set VecSubCRevNaive = a
End Function

Private Sub VecSubCRevNaive_I(ByVal a As Tensor, _
                              ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = dblScalar - A_(i)
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecMulNaive(ByVal a As Tensor, _
                             ByVal b As Tensor) As Tensor
    Set a = a.Clone
    VecMulNaive_I a, b
    Set VecMulNaive = a
End Function

Private Sub VecMulNaive_I(ByVal a As Tensor, _
                          ByVal b As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    a.Flatten.CreateAlias A_
    b.Flatten.CreateAlias B_
    For i = 1 To a.NumElements
        A_(i) = A_(i) * B_(i)
    Next i
    a.Flatten.RemoveAlias A_
    b.Flatten.RemoveAlias B_
End Sub

Private Function VecMulCNaive(ByVal a As Tensor, _
                              ByVal dblScalar As Double) As Tensor
    Set a = a.Clone
    VecMulCNaive_I a, dblScalar
    Set VecMulCNaive = a
End Function

Private Sub VecMulCNaive_I(ByVal a As Tensor, _
                           ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = dblScalar * A_(i)
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecMulCBlas(ByVal a As Tensor, _
                             ByVal dblScalar As Double) As Tensor
    Set a = a.Clone
    VecMulCBlas_I a, dblScalar
    Set VecMulCBlas = a
End Function

Private Sub VecMulCBlas_I(ByVal a As Tensor, _
                          ByVal dblScalar As Double)
    dscal a.NumElements, dblScalar, a.Address, 1&
End Sub

Private Function VecDivNaive(ByVal a As Tensor, _
                             ByVal b As Tensor) As Tensor
    Set a = a.Clone
    VecDivNaive_I a, b
    Set VecDivNaive = a
End Function

Private Sub VecDivNaive_I(ByVal a As Tensor, _
                          ByVal b As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    a.Flatten.CreateAlias A_
    b.Flatten.CreateAlias B_
    For i = 1 To a.NumElements
        A_(i) = A_(i) / B_(i)
    Next i
    a.Flatten.RemoveAlias A_
    b.Flatten.RemoveAlias B_
End Sub

Private Function VecDivCRevNaive(ByVal a As Tensor, _
                                 ByVal dblScalar As Double) As Tensor
    Set a = a.Clone
    VecDivCRevNaive_I a, dblScalar
    Set VecDivCRevNaive = a
End Function

Private Sub VecDivCRevNaive_I(ByVal a As Tensor, _
                              ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = dblScalar / A_(i)
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecDivSqrtAddCNaive(ByVal a As Tensor, _
                                     ByVal b As Tensor, _
                                     ByVal dblScalar As Double) As Tensor
    Set a = a.Clone
    VecDivSqrtAddCNaive_I a, b, dblScalar
    Set VecDivSqrtAddCNaive = a
End Function

Private Sub VecDivSqrtAddCNaive_I(ByVal a As Tensor, _
                                  ByVal b As Tensor, _
                                  ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    a.Flatten.CreateAlias A_
    b.Flatten.CreateAlias B_
    For i = 1 To a.NumElements
        A_(i) = A_(i) / (Sqr(B_(i)) + dblScalar)
    Next i
    a.Flatten.RemoveAlias A_
    b.Flatten.RemoveAlias B_
End Sub

Private Function VecAbsNaive(ByVal a As Tensor) As Tensor
    Set a = a.Clone
    VecAbsNaive_I a
    Set VecAbsNaive = a
End Function

Private Sub VecAbsNaive_I(ByVal a As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = Abs(A_(i))
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecSignNaive(ByVal a As Tensor) As Tensor
    Set a = a.Clone
    VecSignNaive_I a
    Set VecSignNaive = a
End Function

Private Sub VecSignNaive_I(ByVal a As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = Sgn(A_(i))
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecSqrtNaive(ByVal a As Tensor) As Tensor
    Set a = a.Clone
    VecSqrtNaive_I a
    Set VecSqrtNaive = a
End Function

Private Sub VecSqrtNaive_I(ByVal a As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = Sqr(A_(i))
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecExpNaive(ByVal a As Tensor) As Tensor
    Set a = a.Clone
    VecExpNaive_I a
    Set VecExpNaive = a
End Function

Private Sub VecExpNaive_I(ByVal a As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = Exp(A_(i))
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecLogNaive(ByVal a As Tensor) As Tensor
    Set a = a.Clone
    VecLogNaive_I a
    Set VecLogNaive = a
End Function

Private Sub VecLogNaive_I(ByVal a As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = Log(A_(i))
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecLeakyReLUNaive(ByVal a As Tensor, _
                                   ByVal dblNegativeSlope As Double) As Tensor
    Set a = a.Clone
    VecLeakyReLUNaive_I a, dblNegativeSlope
    Set VecLeakyReLUNaive = a
End Function

Private Sub VecLeakyReLUNaive_I(ByVal a As Tensor, _
                                ByVal dblNegativeSlope As Double)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        If A_(i) < 0 Then
            A_(i) = dblNegativeSlope * A_(i)
        End If
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecLeakyReLUDerivativeNaive(ByVal a As Tensor, _
                                             ByVal dblNegativeSlope As Double) As Tensor
    Set a = a.Clone
    VecLeakyReLUDerivativeNaive_I a, dblNegativeSlope
    Set VecLeakyReLUDerivativeNaive = a
End Function

Private Sub VecLeakyReLUDerivativeNaive_I(ByVal a As Tensor, _
                                          ByVal dblNegativeSlope As Double)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        If A_(i) < 0 Then
            A_(i) = dblNegativeSlope
        Else
            A_(i) = 1
        End If
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecSigmoidNaive(ByVal a As Tensor) As Tensor
    Set a = a.Clone
    VecSigmoidNaive_I a
    Set VecSigmoidNaive = a
End Function

Private Sub VecSigmoidNaive_I(ByVal a As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = SafeSigmoid(A_(i))
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecSigmoidDerivativeNaive(ByVal a As Tensor) As Tensor
    Set a = a.Clone
    VecSigmoidDerivativeNaive_I a
    Set VecSigmoidDerivativeNaive = a
End Function

Private Sub VecSigmoidDerivativeNaive_I(ByVal a As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = A_(i) * (1 - A_(i))
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecTanhNaive(ByVal a As Tensor) As Tensor
    Set a = a.Clone
    VecTanhNaive_I a
    Set VecTanhNaive = a
End Function

Private Sub VecTanhNaive_I(ByVal a As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = SafeTanh(A_(i))
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecTanhDerivativeNaive(ByVal a As Tensor) As Tensor
    Set a = a.Clone
    VecTanhDerivativeNaive_I a
    Set VecTanhDerivativeNaive = a
End Function

Private Sub VecTanhDerivativeNaive_I(ByVal a As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    a.Flatten.CreateAlias A_
    For i = 1 To a.NumElements
        A_(i) = 1 - A_(i) * A_(i)
    Next i
    a.Flatten.RemoveAlias A_
End Sub

Private Function VecLinCombNaive(ByVal dblAlpha As Double, _
                                 ByVal a As Tensor, _
                                 ByVal dblBeta As Double, _
                                 ByVal b As Tensor) As Tensor
    Set a = a.Clone
    VecLinCombNaive_I dblAlpha, a, dblBeta, b
    Set VecLinCombNaive = a
End Function

Private Sub VecLinCombNaive_I(ByVal dblAlpha As Double, _
                              ByVal a As Tensor, _
                              ByVal dblBeta As Double, _
                              ByVal b As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    a.Flatten.CreateAlias A_
    b.Flatten.CreateAlias B_
    For i = 1 To a.NumElements
        A_(i) = dblAlpha * A_(i) + dblBeta * B_(i)
    Next i
    a.Flatten.RemoveAlias A_
    b.Flatten.RemoveAlias B_
End Sub

Private Function VecLinCombBlas(ByVal dblAlpha As Double, _
                                ByVal a As Tensor, _
                                ByVal dblBeta As Double, _
                                ByVal b As Tensor) As Tensor
    Set a = a.Clone
    VecLinCombBlas_I dblAlpha, a, dblBeta, b
    Set VecLinCombBlas = a
End Function

Private Sub VecLinCombBlas_I(ByVal dblAlpha As Double, _
                             ByVal a As Tensor, _
                             ByVal dblBeta As Double, _
                             ByVal b As Tensor)
    daxpby a.NumElements, dblBeta, b.Address, 1&, dblAlpha, a.Address, 1&
End Sub

Private Function MatMulNaive(ByVal a As Tensor, _
                             ByVal b As Tensor, _
                             ByVal bTransposeA As Boolean, _
                             ByVal bTransposeB As Boolean) As Tensor
    Dim m As Long
    Dim n As Long
    Dim C As Tensor
    
    m = IIf(bTransposeA, a.Size(2), a.Size(1))
    n = IIf(bTransposeB, b.Size(1), b.Size(2))
    Set C = Zeros(Array(m, n))
    MatMulNaive_I C, a, b, bTransposeA, bTransposeB
    Set MatMulNaive = C
End Function

Private Sub MatMulNaive_I(ByVal C As Tensor, _
                          ByVal a As Tensor, _
                          ByVal b As Tensor, _
                          ByVal bTransposeA As Boolean, _
                          ByVal bTransposeB As Boolean)
    Dim i As Long
    Dim j As Long
    Dim p As Long
    Dim m As Long
    Dim n As Long
    Dim k As Long
    Dim dblSum As Double
    Dim A_() As Double
    Dim B_() As Double
    Dim C_() As Double
    
    m = C.Size(1)
    n = C.Size(2)
    k = IIf(bTransposeA, a.Size(1), a.Size(2))
    a.CreateAlias A_
    b.CreateAlias B_
    C.CreateAlias C_
    For i = 1 To m
        For j = 1 To n
            dblSum = 0
            Select Case True
                Case Not bTransposeA And Not bTransposeB
                    For p = 1 To k
                        dblSum = dblSum + A_(i, p) * B_(p, j)
                    Next p
                Case bTransposeA And Not bTransposeB
                    For p = 1 To k
                        dblSum = dblSum + A_(p, i) * B_(p, j)
                    Next p
                Case Not bTransposeA And bTransposeB
                    For p = 1 To k
                        dblSum = dblSum + A_(i, p) * B_(j, p)
                    Next p
                Case bTransposeA And bTransposeB
                    For p = 1 To k
                        dblSum = dblSum + A_(p, i) * B_(j, p)
                    Next p
            End Select
            C_(i, j) = C_(i, j) + dblSum
        Next j
    Next i
    a.RemoveAlias A_
    b.RemoveAlias B_
    C.RemoveAlias C_
End Sub

Private Function MatMulBlas(ByVal a As Tensor, _
                            ByVal b As Tensor, _
                            ByVal bTransposeA As Boolean, _
                            ByVal bTransposeB As Boolean) As Tensor
    Dim m As Long
    Dim n As Long
    Dim C As Tensor
    
    m = IIf(bTransposeA, a.Size(2), a.Size(1))
    n = IIf(bTransposeB, b.Size(1), b.Size(2))
    Set C = Zeros(Array(m, n))
    MatMulBlas_I C, a, b, bTransposeA, bTransposeB
    Set MatMulBlas = C
End Function

Private Sub MatMulBlas_I(ByVal C As Tensor, _
                         ByVal a As Tensor, _
                         ByVal b As Tensor, _
                         ByVal bTransposeA As Boolean, _
                         ByVal bTransposeB As Boolean)
    Dim sTransposeA As String
    Dim sTransposeB As String
    Dim m As Long
    Dim n As Long
    Dim k As Long
    
    sTransposeA = IIf(bTransposeA, "T", "N")
    sTransposeB = IIf(bTransposeB, "T", "N")
    m = C.Size(1)
    n = C.Size(2)
    k = IIf(bTransposeA, a.Size(1), a.Size(2))
    dgemm sTransposeA, sTransposeB, m, n, k, 1#, a.Address, a.Size(1), b.Address, b.Size(1), 1#, C.Address, m
End Sub
