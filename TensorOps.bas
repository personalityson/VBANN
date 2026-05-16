Attribute VB_Name = "TensorOps"
Option Explicit

Private Const OPENBLAS_PATH As String = "C:\Users\hello\OneDrive\Documents\VBANN\libopenblas.dll"

Private m_bBlasInitialized As Boolean
Private m_bIsBlasAvailable As Boolean

Private Declare PtrSafe Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" (ByVal lpPathName As String) As Long

Private Declare PtrSafe Function ddot Lib "libopenblas.dll" (ByRef n As Long, _
                                                             ByVal X As LongPtr, _
                                                             ByRef incX As Long, _
                                                             ByVal Y As LongPtr, _
                                                             ByRef incY As Long) As Double

Private Declare PtrSafe Sub dscal Lib "libopenblas.dll" (ByRef n As Long, _
                                                         ByRef alpha As Double, _
                                                         ByVal X As LongPtr, _
                                                         ByRef incX As Long)

Private Declare PtrSafe Function dnrm2 Lib "libopenblas.dll" (ByRef n As Long, _
                                                              ByVal X As LongPtr, _
                                                              ByRef incX As Long) As Double

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
                                                         ByVal A As LongPtr, _
                                                         ByRef ldA As Long, _
                                                         ByVal B As LongPtr, _
                                                         ByRef ldB As Long, _
                                                         ByRef beta As Double, _
                                                         ByVal C As LongPtr, _
                                                         ByRef ldC As Long)

Private Declare PtrSafe Sub domatcopy Lib "libopenblas.dll" (ByVal order As String, _
                                                             ByVal trans As String, _
                                                             ByRef m As Long, _
                                                             ByRef n As Long, _
                                                             ByRef alpha As Double, _
                                                             ByVal A As LongPtr, _
                                                             ByRef ldA As Long, _
                                                             ByVal B As LongPtr, _
                                                             ByRef ldB As Long)

Public Function IsBlasAvailable() As Boolean
    If Not m_bBlasInitialized Then
        If Fso.FileExists(OPENBLAS_PATH) Then
            m_bIsBlasAvailable = SetDllDirectory(Fso.GetParentFolderName(OPENBLAS_PATH)) <> 0
        End If
        m_bBlasInitialized = True
    End If
    IsBlasAvailable = m_bIsBlasAvailable
End Function

'VecDot                 Y = Sum(A * B)
'VecNorm2               Y = ||A||_2
'VecAdd                 Y = A + B
'VecAdd_I               A = A + B (In-place)
'VecAddC                Y = A + scalar
'VecSub                 Y = A - B
'VecSubC                Y = A - scalar
'VecSubCRev             Y = scalar - A
'VecMul                 Y = A * B
'VecMulC                Y = A * scalar
'VecMulC_I              A = A * scalar (In-place)
'VecDiv                 Y = A / B
'VecDivC                Y = A / scalar
'VecDivCRev             Y = scalar / A
'VecDivRms              Y = A / (Sqrt(B + inner_epsilon) + outer_epsilon)
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
'MatTranspose           Y = A^T

'Y = Sum(A * B)
Public Function VecDot(ByVal A As Tensor, _
                       ByVal B As Tensor) As Double
    Const PROCEDURE_NAME As String = "TensorOps.VecDot"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If IsBlasAvailable() Then
        VecDot = VecDotBlas(A, B)
    Else
        VecDot = VecDotNaive(A, B)
    End If
End Function

'Y = ||A||_2
Public Function VecNorm2(ByVal A As Tensor) As Double
    Const PROCEDURE_NAME As String = "TensorOps.VecNorm2"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        VecNorm2 = VecNorm2Blas(A)
    Else
        VecNorm2 = VecNorm2Naive(A)
    End If
End Function

'Y = A + B
Public Function VecAdd(ByVal A As Tensor, _
                       ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecAdd"

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If IsBlasAvailable() Then
        Set VecAdd = VecLinCombBlas(1, A, 1, B)
    Else
        Set VecAdd = VecAddNaive(A, B)
    End If
End Function

'A = A + B
Public Sub VecAdd_I(ByVal A As Tensor, _
                    ByVal B As Tensor)
    Const PROCEDURE_NAME As String = "TensorOps.VecAdd_I"

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If IsBlasAvailable() Then
        VecLinCombBlas_I 1, A, 1, B
    Else
        VecAddNaive_I A, B
    End If
End Sub

'Y = A + scalar
Public Function VecAddC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecAddC"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecAddC = VecAddCNaive(A, dblScalar)
End Function

'Y = A - B
Public Function VecSub(ByVal A As Tensor, _
                       ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSub"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If IsBlasAvailable() Then
        Set VecSub = VecLinCombBlas(1, A, -1, B)
    Else
        Set VecSub = VecSubNaive(A, B)
    End If
End Function

'Y = A - scalar
Public Function VecSubC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSubC"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSubC = VecSubCNaive(A, dblScalar)
End Function

'Y = scalar - A
Public Function VecSubCRev(ByVal A As Tensor, _
                           ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSubCRev"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSubCRev = VecSubCRevNaive(A, dblScalar)
End Function

'Y = A * B
Public Function VecMul(ByVal A As Tensor, _
                       ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecMul"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    Set VecMul = VecMulNaive(A, B)
End Function

'Y = A * scalar
Public Function VecMulC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecMulC"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecMulC = VecMulCBlas(A, dblScalar)
    Else
        Set VecMulC = VecMulCNaive(A, dblScalar)
    End If
End Function

'A = A * scalar
Public Sub VecMulC_I(ByVal A As Tensor, _
                     ByVal dblScalar As Double)
    Const PROCEDURE_NAME As String = "TensorOps.VecMulC_I"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        VecMulCBlas_I A, dblScalar
    Else
        VecMulCNaive_I A, dblScalar
    End If
End Sub

'Y = A / B
Public Function VecDiv(ByVal A As Tensor, _
                       ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecDiv"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    Set VecDiv = VecDivNaive(A, B)
End Function

'Y = A / scalar
Public Function VecDivC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Set VecDivC = VecMulC(A, 1 / dblScalar)
End Function

'Y = scalar / A
Public Function VecDivCRev(ByVal A As Tensor, _
                           ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecDivCRev"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecDivCRev = VecDivCRevNaive(A, dblScalar)
End Function

'Y = A / (Sqrt(B + inner_epsilon) + outer_epsilon)
Public Function VecDivRms(ByVal A As Tensor, _
                          ByVal B As Tensor, _
                          ByVal dblInnerEpsilon As Double, _
                          ByVal dblOuterEpsilon As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecDivRms"

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If dblInnerEpsilon < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Inner epsilon must be >= 0."
    End If
    If dblOuterEpsilon < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Outer epsilon must be >= 0."
    End If
    Set VecDivRms = VecDivRmsNaive(A, B, dblInnerEpsilon, dblOuterEpsilon)
End Function

'Y = Abs(A)
Public Function VecAbs(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecAbs"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecAbs = VecAbsNaive(A)
End Function

'Y = Sign(A)
Public Function VecSign(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSign"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSign = VecSignNaive(A)
End Function

'Y = A^2
Public Function VecPow2(ByVal A As Tensor) As Tensor
    Set VecPow2 = VecMul(A, A)
End Function

'Y = Sqrt(A)
Public Function VecSqrt(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSqrt"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSqrt = VecSqrtNaive(A)
End Function

'Y = Exp(A)
Public Function VecExp(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecExp"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecExp = VecExpNaive(A)
End Function

'Y = Log(A)
Public Function VecLog(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLog"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecLog = VecLogNaive(A)
End Function

'Y = If A > 0 Then A Else dblNegativeSlope * A
Public Function VecLeakyReLU(ByVal A As Tensor, _
                             ByVal dblNegativeSlope As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLeakyReLU"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecLeakyReLU = VecLeakyReLUNaive(A, dblNegativeSlope)
End Function

'Y = If A > 0 Then 1 Else dblNegativeSlope
Public Function VecLeakyReLUDerivative(ByVal A As Tensor, _
                                       ByVal dblNegativeSlope As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLeakyReLUDerivative"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecLeakyReLUDerivative = VecLeakyReLUDerivativeNaive(A, dblNegativeSlope)
End Function

'Y = 1 / (1 + Exp(-A))
Public Function VecSigmoid(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSigmoid"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSigmoid = VecSigmoidNaive(A)
End Function

'Y = A * (1 - A)
Public Function VecSigmoidDerivative(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSigmoidDerivative"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSigmoidDerivative = VecSigmoidDerivativeNaive(A)
End Function

'Y = Tanh(A)
Public Function VecTanh(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecTanh"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecTanh = VecTanhNaive(A)
End Function

'Y = 1 - A^2
Public Function VecTanhDerivative(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecTanhDerivative"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecTanhDerivative = VecTanhDerivativeNaive(A)
End Function

'Y = alpha * A + beta * B
Public Function VecLinComb(ByVal dblAlpha As Double, _
                           ByVal A As Tensor, _
                           ByVal dblBeta As Double, _
                           ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLinComb"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If IsBlasAvailable() Then
        Set VecLinComb = VecLinCombBlas(dblAlpha, A, dblBeta, B)
    Else
        Set VecLinComb = VecLinCombNaive(dblAlpha, A, dblBeta, B)
    End If
End Function

'A = alpha * A + beta * B
Public Sub VecLinComb_I(ByVal dblAlpha As Double, _
                        ByVal A As Tensor, _
                        ByVal dblBeta As Double, _
                        ByVal B As Tensor)
    Const PROCEDURE_NAME As String = "TensorOps.VecLinComb_I"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    If IsBlasAvailable() Then
        VecLinCombBlas_I dblAlpha, A, dblBeta, B
    Else
        VecLinCombNaive_I dblAlpha, A, dblBeta, B
    End If
End Sub

'Y = A * B
Public Function MatMul(ByVal A As Tensor, _
                       ByVal B As Tensor, _
                       Optional ByVal bTransposeA As Boolean, _
                       Optional ByVal bTransposeB As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.MatMul"
    Dim lNumColsA As Long
    Dim lNumRowsB As Long
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumDimensions < 1 Or A.NumDimensions > 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor A must have 1 or 2 dimensions."
    End If
    If B.NumDimensions < 1 Or B.NumDimensions > 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor B must have 1 or 2 dimensions."
    End If
    If A.NumDimensions = 1 Then
        Set A = A.View(Array(1, A.Size(1)))
    End If
    If B.NumDimensions = 1 Then
        Set B = B.View(Array(B.Size(1), 1))
    End If
    lNumColsA = IIf(bTransposeA, A.Size(1), A.Size(2))
    lNumRowsB = IIf(bTransposeB, B.Size(2), B.Size(1))
    If lNumColsA <> lNumRowsB Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B are incompatible for matrix multiplication."
    End If
    If IsBlasAvailable() Then
        Set MatMul = MatMulBlas(A, B, bTransposeA, bTransposeB)
    Else
        Set MatMul = MatMulNaive(A, B, bTransposeA, bTransposeB)
    End If
End Function

'C = C + A * B
Public Sub MatMul_I(ByVal C As Tensor, _
                    ByVal A As Tensor, _
                    ByVal B As Tensor, _
                    Optional ByVal bTransposeA As Boolean, _
                    Optional ByVal bTransposeB As Boolean)
    Const PROCEDURE_NAME As String = "TensorOps.MatMul_I"
    Dim lNumRowsA As Long
    Dim lNumColsA As Long
    Dim lNumRowsB As Long
    Dim lNumColsB As Long
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If C Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumDimensions < 1 Or A.NumDimensions > 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor A must have 1 or 2 dimensions."
    End If
    If B.NumDimensions < 1 Or B.NumDimensions > 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor B must have 1 or 2 dimensions."
    End If
    If C.NumDimensions < 1 Or C.NumDimensions > 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor C must have 1 or 2 dimensions."
    End If
    If A.NumDimensions = 1 Then
        Set A = A.View(Array(1, A.Size(1)))
    End If
    If B.NumDimensions = 1 Then
        Set B = B.View(Array(B.Size(1), 1))
    End If
    lNumRowsA = IIf(bTransposeA, A.Size(2), A.Size(1))
    lNumColsA = IIf(bTransposeA, A.Size(1), A.Size(2))
    lNumRowsB = IIf(bTransposeB, B.Size(2), B.Size(1))
    lNumColsB = IIf(bTransposeB, B.Size(1), B.Size(2))
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
        MatMulBlas_I C, A, B, bTransposeA, bTransposeB
    Else
        MatMulNaive_I C, A, B, bTransposeA, bTransposeB
    End If
End Sub

Public Function MatTranspose(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.MatTranspose"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumDimensions <> 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor must be 2-dimensional."
    End If
    If IsBlasAvailable() Then
        Set MatTranspose = MatTransposeBlas(A)
    Else
        Set MatTranspose = MatTransposeNaive(A)
    End If
End Function

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

Private Function VecDotNaive(ByVal A As Tensor, _
                             ByVal B As Tensor) As Double
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    Dim dblSum As Double
    
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        dblSum = dblSum + A_(i) * B_(i)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
    VecDotNaive = dblSum
End Function

Private Function VecDotBlas(ByVal A As Tensor, _
                            ByVal B As Tensor) As Double
    VecDotBlas = ddot(A.NumElements, A.Address, 1&, B.Address, 1&)
End Function

Private Function VecNorm2Naive(ByVal A As Tensor) As Double
    Dim i As Long
    Dim dblSum As Double
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        dblSum = dblSum + A_(i) * A_(i)
    Next i
    A.Flatten.RemoveAlias A_
    VecNorm2Naive = Sqr(dblSum)
End Function

Private Function VecNorm2Blas(ByVal A As Tensor) As Double
    VecNorm2Blas = dnrm2(A.NumElements, A.Address, 1&)
End Function

Private Function VecAddNaive(ByVal A As Tensor, _
                             ByVal B As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecAddNaive_I Y, B
    Set VecAddNaive = Y
End Function

Private Sub VecAddNaive_I(ByVal A As Tensor, _
                          ByVal B As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double

    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) + B_(i)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
End Sub

Private Function VecAddCNaive(ByVal A As Tensor, _
                              ByVal dblScalar As Double) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecAddCNaive_I Y, dblScalar
    Set VecAddCNaive = Y
End Function

Private Sub VecAddCNaive_I(ByVal A As Tensor, _
                           ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = A_(i) + dblScalar
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecSubNaive(ByVal A As Tensor, _
                             ByVal B As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecSubNaive_I Y, B
    Set VecSubNaive = Y
End Function

Private Sub VecSubNaive_I(ByVal A As Tensor, _
                          ByVal B As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) - B_(i)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
End Sub

Private Function VecSubCNaive(ByVal A As Tensor, _
                              ByVal dblScalar As Double) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecSubCNaive_I Y, dblScalar
    Set VecSubCNaive = Y
End Function

Private Sub VecSubCNaive_I(ByVal A As Tensor, _
                           ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = A_(i) - dblScalar
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecSubCRevNaive(ByVal A As Tensor, _
                                 ByVal dblScalar As Double) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecSubCRevNaive_I Y, dblScalar
    Set VecSubCRevNaive = Y
End Function

Private Sub VecSubCRevNaive_I(ByVal A As Tensor, _
                              ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = dblScalar - A_(i)
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecMulNaive(ByVal A As Tensor, _
                             ByVal B As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecMulNaive_I Y, B
    Set VecMulNaive = Y
End Function

Private Sub VecMulNaive_I(ByVal A As Tensor, _
                          ByVal B As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) * B_(i)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
End Sub

Private Function VecMulCNaive(ByVal A As Tensor, _
                              ByVal dblScalar As Double) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecMulCNaive_I Y, dblScalar
    Set VecMulCNaive = Y
End Function

Private Sub VecMulCNaive_I(ByVal A As Tensor, _
                           ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = dblScalar * A_(i)
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecMulCBlas(ByVal A As Tensor, _
                             ByVal dblScalar As Double) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecMulCBlas_I Y, dblScalar
    Set VecMulCBlas = Y
End Function

Private Sub VecMulCBlas_I(ByVal A As Tensor, _
                          ByVal dblScalar As Double)
    dscal A.NumElements, dblScalar, A.Address, 1&
End Sub

Private Function VecDivNaive(ByVal A As Tensor, _
                             ByVal B As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecDivNaive_I Y, B
    Set VecDivNaive = Y
End Function

Private Sub VecDivNaive_I(ByVal A As Tensor, _
                          ByVal B As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) / B_(i)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
End Sub

Private Function VecDivCRevNaive(ByVal A As Tensor, _
                                 ByVal dblScalar As Double) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecDivCRevNaive_I Y, dblScalar
    Set VecDivCRevNaive = Y
End Function

Private Sub VecDivCRevNaive_I(ByVal A As Tensor, _
                              ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = dblScalar / A_(i)
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecDivRmsNaive(ByVal A As Tensor, _
                                ByVal B As Tensor, _
                                ByVal dblInnerEpsilon As Double, _
                                ByVal dblOuterEpsilon As Double) As Tensor
    Dim Y As Tensor

    Set Y = A.Clone
    VecDivRmsNaive_I Y, B, dblInnerEpsilon, dblOuterEpsilon
    Set VecDivRmsNaive = Y
End Function

Private Sub VecDivRmsNaive_I(ByVal A As Tensor, _
                             ByVal B As Tensor, _
                             ByVal dblInnerEpsilon As Double, _
                             ByVal dblOuterEpsilon As Double)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double

    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) / (Sqr(B_(i) + dblInnerEpsilon) + dblOuterEpsilon)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
End Sub

Private Function VecAbsNaive(ByVal A As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecAbsNaive_I Y
    Set VecAbsNaive = Y
End Function

Private Sub VecAbsNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = Abs(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecSignNaive(ByVal A As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecSignNaive_I Y
    Set VecSignNaive = Y
End Function

Private Sub VecSignNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = Sgn(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecSqrtNaive(ByVal A As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecSqrtNaive_I Y
    Set VecSqrtNaive = Y
End Function

Private Sub VecSqrtNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = Sqr(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecExpNaive(ByVal A As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecExpNaive_I Y
    Set VecExpNaive = Y
End Function

Private Sub VecExpNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = Exp(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecLogNaive(ByVal A As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecLogNaive_I Y
    Set VecLogNaive = Y
End Function

Private Sub VecLogNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = Log(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecLeakyReLUNaive(ByVal A As Tensor, _
                                   ByVal dblNegativeSlope As Double) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecLeakyReLUNaive_I Y, dblNegativeSlope
    Set VecLeakyReLUNaive = Y
End Function

Private Sub VecLeakyReLUNaive_I(ByVal A As Tensor, _
                                ByVal dblNegativeSlope As Double)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        If A_(i) < 0 Then
            A_(i) = dblNegativeSlope * A_(i)
        End If
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecLeakyReLUDerivativeNaive(ByVal A As Tensor, _
                                             ByVal dblNegativeSlope As Double) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecLeakyReLUDerivativeNaive_I Y, dblNegativeSlope
    Set VecLeakyReLUDerivativeNaive = Y
End Function

Private Sub VecLeakyReLUDerivativeNaive_I(ByVal A As Tensor, _
                                          ByVal dblNegativeSlope As Double)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        If A_(i) < 0 Then
            A_(i) = dblNegativeSlope
        Else
            A_(i) = 1
        End If
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecSigmoidNaive(ByVal A As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecSigmoidNaive_I Y
    Set VecSigmoidNaive = Y
End Function

Private Sub VecSigmoidNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = SafeSigmoid(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecSigmoidDerivativeNaive(ByVal A As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecSigmoidDerivativeNaive_I Y
    Set VecSigmoidDerivativeNaive = Y
End Function

Private Sub VecSigmoidDerivativeNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = A_(i) * (1 - A_(i))
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecTanhNaive(ByVal A As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecTanhNaive_I Y
    Set VecTanhNaive = Y
End Function

Private Sub VecTanhNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = SafeTanh(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecTanhDerivativeNaive(ByVal A As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecTanhDerivativeNaive_I Y
    Set VecTanhDerivativeNaive = Y
End Function

Private Sub VecTanhDerivativeNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = 1 - A_(i) * A_(i)
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecLinCombNaive(ByVal dblAlpha As Double, _
                                 ByVal A As Tensor, _
                                 ByVal dblBeta As Double, _
                                 ByVal B As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecLinCombNaive_I dblAlpha, Y, dblBeta, B
    Set VecLinCombNaive = Y
End Function

Private Sub VecLinCombNaive_I(ByVal dblAlpha As Double, _
                              ByVal A As Tensor, _
                              ByVal dblBeta As Double, _
                              ByVal B As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = dblAlpha * A_(i) + dblBeta * B_(i)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
End Sub

Private Function VecLinCombBlas(ByVal dblAlpha As Double, _
                                ByVal A As Tensor, _
                                ByVal dblBeta As Double, _
                                ByVal B As Tensor) As Tensor
    Dim Y As Tensor
    
    Set Y = A.Clone
    VecLinCombBlas_I dblAlpha, Y, dblBeta, B
    Set VecLinCombBlas = Y
End Function

Private Sub VecLinCombBlas_I(ByVal dblAlpha As Double, _
                             ByVal A As Tensor, _
                             ByVal dblBeta As Double, _
                             ByVal B As Tensor)
    daxpby A.NumElements, dblBeta, B.Address, 1&, dblAlpha, A.Address, 1&
End Sub

Private Function MatMulNaive(ByVal A As Tensor, _
                             ByVal B As Tensor, _
                             ByVal bTransposeA As Boolean, _
                             ByVal bTransposeB As Boolean) As Tensor
    Dim m As Long
    Dim n As Long
    Dim C As Tensor
    
    m = IIf(bTransposeA, A.Size(2), A.Size(1))
    n = IIf(bTransposeB, B.Size(1), B.Size(2))
    Set C = Zeros(Array(m, n))
    MatMulNaive_I C, A, B, bTransposeA, bTransposeB
    Set MatMulNaive = C
End Function

Private Sub MatMulNaive_I(ByVal C As Tensor, _
                          ByVal A As Tensor, _
                          ByVal B As Tensor, _
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
    k = IIf(bTransposeA, A.Size(1), A.Size(2))
    A.CreateAlias A_
    B.CreateAlias B_
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
    A.RemoveAlias A_
    B.RemoveAlias B_
    C.RemoveAlias C_
End Sub

Private Function MatMulBlas(ByVal A As Tensor, _
                            ByVal B As Tensor, _
                            ByVal bTransposeA As Boolean, _
                            ByVal bTransposeB As Boolean) As Tensor
    Dim m As Long
    Dim n As Long
    Dim C As Tensor
    
    m = IIf(bTransposeA, A.Size(2), A.Size(1))
    n = IIf(bTransposeB, B.Size(1), B.Size(2))
    Set C = Zeros(Array(m, n))
    MatMulBlas_I C, A, B, bTransposeA, bTransposeB
    Set MatMulBlas = C
End Function

Private Sub MatMulBlas_I(ByVal C As Tensor, _
                         ByVal A As Tensor, _
                         ByVal B As Tensor, _
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
    k = IIf(bTransposeA, A.Size(1), A.Size(2))
    dgemm sTransposeA, sTransposeB, m, n, k, 1#, A.Address, A.Size(1), B.Address, B.Size(1), 1#, C.Address, m
End Sub

Private Function MatTransposeNaive(ByVal A As Tensor) As Tensor
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim n As Long
    Dim A_() As Double
    Dim Y_() As Double
    Dim Y As Tensor
    
    m = A.Size(1)
    n = A.Size(2)
    Set Y = Zeros(Array(n, m))
    A.CreateAlias A_
    Y.CreateAlias Y_
    For j = 1 To m
        For i = 1 To n
            Y_(i, j) = A_(j, i)
        Next i
    Next j
    A.RemoveAlias A_
    Y.RemoveAlias Y_
    Set MatTransposeNaive = Y
End Function

Private Function MatTransposeBlas(ByVal A As Tensor) As Tensor
    Dim m As Long
    Dim n As Long
    Dim Y As Tensor
    
    m = A.Size(1)
    n = A.Size(2)
    Set Y = Zeros(Array(n, m))
    domatcopy "C", "T", m, n, 1#, A.Address, m, Y.Address, n
    Set MatTransposeBlas = Y
End Function
