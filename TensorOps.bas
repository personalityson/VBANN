Attribute VB_Name = "TensorOps"
Option Explicit

'https://github.com/OpenMathLib/OpenBLAS/releases
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
'VecMul_I               A = A * B (In-place)
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
'VecClamp               Y = Min(Max(A, dblMin), dblMax)
'VecWhere               Y = IIf(C, A, B)
'VecLeakyReLU           Y = If A > 0 Then A Else dblNegativeSlope * A
'VecLeakyReLUDerivative Y = If A > 0 Then 1 Else dblNegativeSlope
'VecSigmoid             Y = 1 / (1 + Exp(-A))
'VecSigmoidDerivative   Y = A * (1 - A)
'VecSoftplus            Y = Log(1 + Exp(A))
'VecTanh                Y = Tanh(A)
'VecTanhDerivative      Y = 1 - A^2
'VecLinComb             Y = alpha * A + beta * B
'VecLinComb_I           A = alpha * A + beta * B (In-place)
'MatMul                 Y = A * B
'MatMul_I               C = C + A * B (In-place)
'Softmax                Y = Exp(A) / Sum(Exp(A))
'LogSoftmax             Y = A - Log(Sum(Exp(A)))
'Transpose              Y = A^T

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
    If Not TryBroadcast(A, B) Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B cannot be broadcast."
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
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not TryBroadcast(A, B) Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B cannot be broadcast."
    End If
    Set Y = A.Clone
    If IsBlasAvailable() Then
        VecLinCombBlas_I 1, Y, 1, B
    Else
        VecAddNaive_I Y, B
    End If
    Set VecAdd = Y
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
    If Not TryBroadcast(A, B, True) Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor B cannot be broadcast to the shape of tensor A."
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
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    If IsBlasAvailable() Then
        VecAddCBlas_I Y, dblScalar
    Else
        VecAddCNaive_I Y, dblScalar
    End If
    Set VecAddC = Y
End Function

'Y = A - B
Public Function VecSub(ByVal A As Tensor, _
                       ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSub"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not TryBroadcast(A, B) Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B cannot be broadcast."
    End If
    Set Y = A.Clone
    If IsBlasAvailable() Then
        VecLinCombBlas_I 1, Y, -1, B
    Else
        VecSubNaive_I Y, B
    End If
    Set VecSub = Y
End Function

'Y = A - scalar
Public Function VecSubC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSubC"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    If IsBlasAvailable() Then
        VecSubCBlas_I Y, dblScalar
    Else
        VecSubCNaive_I Y, dblScalar
    End If
    Set VecSubC = Y
End Function

'Y = scalar - A
Public Function VecSubCRev(ByVal A As Tensor, _
                           ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSubCRev"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    If IsBlasAvailable() Then
        VecSubCRevBlas_I Y, dblScalar
    Else
        VecSubCRevNaive_I Y, dblScalar
    End If
    Set VecSubCRev = Y
End Function

'Y = A * B
Public Function VecMul(ByVal A As Tensor, _
                       ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecMul"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not TryBroadcast(A, B) Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B cannot be broadcast."
    End If
    Set Y = A.Clone
    VecMulNaive_I Y, B
    Set VecMul = Y
End Function

'A = A * B
Public Sub VecMul_I(ByVal A As Tensor, _
                    ByVal B As Tensor)
    Const PROCEDURE_NAME As String = "TensorOps.VecMul_I"

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not TryBroadcast(A, B, True) Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor B cannot be broadcast to the shape of tensor A."
    End If
    VecMulNaive_I A, B
End Sub

'Y = A * scalar
Public Function VecMulC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecMulC"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    If IsBlasAvailable() Then
        VecMulCBlas_I Y, dblScalar
    Else
        VecMulCNaive_I Y, dblScalar
    End If
    Set VecMulC = Y
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
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not TryBroadcast(A, B) Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B cannot be broadcast."
    End If
    Set Y = A.Clone
    VecDivNaive_I Y, B
    Set VecDiv = Y
End Function

'Y = A / scalar
Public Function VecDivC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecDivC"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecDivC = VecMulC(A, 1 / dblScalar)
End Function

'Y = scalar / A
Public Function VecDivCRev(ByVal A As Tensor, _
                           ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecDivCRev"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecDivCRevNaive_I Y, dblScalar
    Set VecDivCRev = Y
End Function

'Y = A / (Sqrt(B + inner_epsilon) + outer_epsilon)
Public Function VecDivRms(ByVal A As Tensor, _
                          ByVal B As Tensor, _
                          ByVal dblInnerEpsilon As Double, _
                          ByVal dblOuterEpsilon As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecDivRms"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If dblInnerEpsilon < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Inner epsilon must be >= 0."
    End If
    If dblOuterEpsilon < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Outer epsilon must be >= 0."
    End If
    If Not TryBroadcast(A, B) Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B cannot be broadcast."
    End If
    Set Y = A.Clone
    VecDivRmsNaive_I Y, B, dblInnerEpsilon, dblOuterEpsilon
    Set VecDivRms = Y
End Function

'Y = Abs(A)
Public Function VecAbs(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecAbs"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecAbsNaive_I Y
    Set VecAbs = Y
End Function

'Y = Sign(A)
Public Function VecSign(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSign"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecSignNaive_I Y
    Set VecSign = Y
End Function

'Y = A^2
Public Function VecPow2(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecPow2"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecMulNaive_I Y, A
    Set VecPow2 = Y
End Function

'Y = Sqrt(A)
Public Function VecSqrt(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSqrt"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecSqrtNaive_I Y
    Set VecSqrt = Y
End Function

'Y = Exp(A)
Public Function VecExp(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecExp"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecExpNaive_I Y
    Set VecExp = Y
End Function

'Y = Log(A)
Public Function VecLog(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLog"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecLogNaive_I Y
    Set VecLog = Y
End Function

'Y = Min(Max(A, dblMin), dblMax)
Public Function VecClamp(ByVal A As Tensor, _
                         ByVal dblMin As Double, _
                         ByVal dblMax As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecClamp"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If dblMin > dblMax Then
        Err.Raise 5, PROCEDURE_NAME, "Lower bound must be <= upper bound."
    End If
    Set Y = A.Clone
    VecClampNaive_I Y, dblMin, dblMax
    Set VecClamp = Y
End Function

'Y(i) = IIf(C, A, B)
Public Function VecWhere(ByVal C As Tensor, _
                         ByVal A As Tensor, _
                         ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecWhere"
    Dim Y As Tensor

    If C Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not TryBroadcast(A, B) Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B cannot be broadcast."
    End If
    If Not TryBroadcast(C, A) Then
        Err.Raise 5, PROCEDURE_NAME, "Shape of the mask cannot be broadcast with the input tensors."
    End If
    'If the mask widened A, B catches up. This cannot fail, as B already matched A.
    TryBroadcast A, B, True
    Set Y = A.Clone
    VecWhereNaive_I C, Y, B
    Set VecWhere = Y
End Function

'Y = If A > 0 Then A Else dblNegativeSlope * A
Public Function VecLeakyReLU(ByVal A As Tensor, _
                             ByVal dblNegativeSlope As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLeakyReLU"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecLeakyReLUNaive_I Y, dblNegativeSlope
    Set VecLeakyReLU = Y
End Function

'Y = If A > 0 Then 1 Else dblNegativeSlope
Public Function VecLeakyReLUDerivative(ByVal A As Tensor, _
                                       ByVal dblNegativeSlope As Double) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLeakyReLUDerivative"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecLeakyReLUDerivativeNaive_I Y, dblNegativeSlope
    Set VecLeakyReLUDerivative = Y
End Function

'Y = 1 / (1 + Exp(-A))
Public Function VecSigmoid(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSigmoid"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecSigmoidNaive_I Y
    Set VecSigmoid = Y
End Function

'Y = A * (1 - A)
Public Function VecSigmoidDerivative(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSigmoidDerivative"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecSigmoidDerivativeNaive_I Y
    Set VecSigmoidDerivative = Y
End Function

'Y = Log(1 + Exp(A))
Public Function VecSoftplus(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecSoftplus"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecSoftplusNaive_I Y
    Set VecSoftplus = Y
End Function

'Y = Tanh(A)
Public Function VecTanh(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecTanh"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecTanhNaive_I Y
    Set VecTanh = Y
End Function

'Y = 1 - A^2
Public Function VecTanhDerivative(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecTanhDerivative"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Y = A.Clone
    VecTanhDerivativeNaive_I Y
    Set VecTanhDerivative = Y
End Function

'Y = alpha * A + beta * B
Public Function VecLinComb(ByVal dblAlpha As Double, _
                           ByVal A As Tensor, _
                           ByVal dblBeta As Double, _
                           ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.VecLinComb"
    Dim Y As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not TryBroadcast(A, B) Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B cannot be broadcast."
    End If
    Set Y = A.Clone
    If IsBlasAvailable() Then
        VecLinCombBlas_I dblAlpha, Y, dblBeta, B
    Else
        VecLinCombNaive_I dblAlpha, Y, dblBeta, B
    End If
    Set VecLinComb = Y
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
    If Not TryBroadcast(A, B, True) Then
        Err.Raise 5, PROCEDURE_NAME, "Tensor B cannot be broadcast to the shape of tensor A."
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
    Dim lNumRowsA As Long
    Dim lNumColsA As Long
    Dim lNumRowsB As Long
    Dim lNumColsB As Long
    Dim Y As Tensor

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
    lNumRowsA = IIf(bTransposeA, A.Size(2), A.Size(1))
    lNumColsA = IIf(bTransposeA, A.Size(1), A.Size(2))
    lNumRowsB = IIf(bTransposeB, B.Size(2), B.Size(1))
    lNumColsB = IIf(bTransposeB, B.Size(1), B.Size(2))
    If lNumColsA <> lNumRowsB Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B are incompatible for matrix multiplication."
    End If
    Set Y = Zeros(Array(lNumRowsA, lNumColsB))
    If IsBlasAvailable() Then
        MatMulBlas_I Y, A, B, bTransposeA, bTransposeB
    Else
        MatMulNaive_I Y, A, B, bTransposeA, bTransposeB
    End If
    Set MatMul = Y
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

Public Function Softmax(ByVal A As Tensor, _
                        Optional ByVal lDimension As Long = 1) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.Softmax"
    Dim oExp As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If lDimension < 1 Or lDimension > A.NumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    Set oExp = VecExp(VecSub(A, A.Reduce(lDimension, rtMax)))
    Set Softmax = VecDiv(oExp, oExp.Reduce(lDimension, rtSum))
End Function

Public Function LogSoftmax(ByVal A As Tensor, _
                           Optional ByVal lDimension As Long = 1) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.LogSoftmax"
    Dim oShifted As Tensor

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If lDimension < 1 Or lDimension > A.NumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    Set oShifted = VecSub(A, A.Reduce(lDimension, rtMax))
    Set LogSoftmax = VecSub(oShifted, VecLog(VecExp(oShifted).Reduce(lDimension, rtSum)))
End Function

Public Function Transpose(ByVal A As Tensor, _
                          Optional ByVal lDimension1 As Long = 1, _
                          Optional ByVal lDimension2 As Long = 2) As Tensor
    Const PROCEDURE_NAME As String = "TensorOps.Transpose"

    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If lDimension1 < 1 Or lDimension1 > A.NumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    If lDimension2 < 1 Or lDimension2 > A.NumDimensions Then
        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
    End If
    Select Case True
        Case lDimension1 = lDimension2
            Set Transpose = A.Clone
        Case IsBlasAvailable() And A.NumDimensions = 2
            Set Transpose = TransposeBlas(A)
        Case Else
            Set Transpose = TransposeNaive(A, MinLng(lDimension1, lDimension2), MaxLng(lDimension1, lDimension2))
    End Select
End Function

'Broadcasts A and B to a common shape, shapes align from dimension 1 from left to right
Private Function TryBroadcast(ByRef A As Tensor, _
                              ByRef B As Tensor, _
                              Optional ByVal bInPlace As Boolean) As Boolean
    Dim i As Long
    Dim lMaxDimensions As Long

    If Not A.IsConformable(B) Then
        lMaxDimensions = MaxLng(A.NumDimensions, B.NumDimensions)
        Set A = A.Expand(lMaxDimensions)
        Set B = B.Expand(lMaxDimensions)
        For i = 1 To lMaxDimensions
            Select Case True
                Case A.Size(i) = B.Size(i)
                    'Do nothing
                Case B.Size(i) = 1
                    Set B = B.Tile(i, A.Size(i))
                Case A.Size(i) = 1 And Not bInPlace
                    Set A = A.Tile(i, B.Size(i))
                Case Else
                    Exit Function
            End Select
        Next i
    End If
    TryBroadcast = True
End Function

Private Function SafeSigmoid(ByVal dblValue As Double) As Double
    If dblValue < -DOUBLE_MAX_LOG Then
        SafeSigmoid = 0
    Else
        SafeSigmoid = 1 / (1 + Exp(-dblValue))
    End If
End Function

Private Function SafeSoftplus(ByVal dblValue As Double) As Double
    SafeSoftplus = MaxDbl(dblValue, 0) + Log(1 + Exp(-Abs(dblValue)))
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

    A.CreateFlatAlias A_
    B.CreateFlatAlias B_
    For i = 1 To A.NumElements
        dblSum = dblSum + A_(i) * B_(i)
    Next i
    A.RemoveAlias A_
    B.RemoveAlias B_
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

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        dblSum = dblSum + A_(i) * A_(i)
    Next i
    A.RemoveAlias A_
    VecNorm2Naive = Sqr(dblSum)
End Function

Private Function VecNorm2Blas(ByVal A As Tensor) As Double
    VecNorm2Blas = dnrm2(A.NumElements, A.Address, 1&)
End Function

Private Sub VecAddNaive_I(ByVal A As Tensor, _
                          ByVal B As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double

    A.CreateFlatAlias A_
    B.CreateFlatAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) + B_(i)
    Next i
    A.RemoveAlias A_
    B.RemoveAlias B_
End Sub

Private Sub VecAddCNaive_I(ByVal A As Tensor, _
                           ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = A_(i) + dblScalar
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecAddCBlas_I(ByVal A As Tensor, _
                          ByVal dblScalar As Double)
    daxpby A.NumElements, 1#, VarPtr(dblScalar), 0&, 1#, A.Address, 1&
End Sub

Private Sub VecSubNaive_I(ByVal A As Tensor, _
                          ByVal B As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double

    A.CreateFlatAlias A_
    B.CreateFlatAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) - B_(i)
    Next i
    A.RemoveAlias A_
    B.RemoveAlias B_
End Sub

Private Sub VecSubCNaive_I(ByVal A As Tensor, _
                           ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = A_(i) - dblScalar
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecSubCBlas_I(ByVal A As Tensor, _
                          ByVal dblScalar As Double)
    daxpby A.NumElements, -1#, VarPtr(dblScalar), 0&, 1#, A.Address, 1&
End Sub

Private Sub VecSubCRevNaive_I(ByVal A As Tensor, _
                              ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = dblScalar - A_(i)
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecSubCRevBlas_I(ByVal A As Tensor, _
                             ByVal dblScalar As Double)
    daxpby A.NumElements, 1#, VarPtr(dblScalar), 0&, -1#, A.Address, 1&
End Sub

Private Sub VecMulNaive_I(ByVal A As Tensor, _
                          ByVal B As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double

    A.CreateFlatAlias A_
    B.CreateFlatAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) * B_(i)
    Next i
    A.RemoveAlias A_
    B.RemoveAlias B_
End Sub

Private Sub VecMulCNaive_I(ByVal A As Tensor, _
                           ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = dblScalar * A_(i)
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecMulCBlas_I(ByVal A As Tensor, _
                          ByVal dblScalar As Double)
    dscal A.NumElements, dblScalar, A.Address, 1&
End Sub

Private Sub VecDivNaive_I(ByVal A As Tensor, _
                          ByVal B As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double

    A.CreateFlatAlias A_
    B.CreateFlatAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) / B_(i)
    Next i
    A.RemoveAlias A_
    B.RemoveAlias B_
End Sub

Private Sub VecDivCRevNaive_I(ByVal A As Tensor, _
                              ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = dblScalar / A_(i)
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecDivRmsNaive_I(ByVal A As Tensor, _
                             ByVal B As Tensor, _
                             ByVal dblInnerEpsilon As Double, _
                             ByVal dblOuterEpsilon As Double)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double

    A.CreateFlatAlias A_
    B.CreateFlatAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) / (Sqr(B_(i) + dblInnerEpsilon) + dblOuterEpsilon)
    Next i
    A.RemoveAlias A_
    B.RemoveAlias B_
End Sub

Private Sub VecAbsNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = Abs(A_(i))
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecSignNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = Sgn(A_(i))
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecSqrtNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = Sqr(A_(i))
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecExpNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = Exp(A_(i))
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecLogNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = Log(A_(i))
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecClampNaive_I(ByVal A As Tensor, _
                            ByVal dblMin As Double, _
                            ByVal dblMax As Double)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        If A_(i) < dblMin Then
            A_(i) = dblMin
        ElseIf A_(i) > dblMax Then
            A_(i) = dblMax
        End If
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecWhereNaive_I(ByVal C As Tensor, _
                            ByVal A As Tensor, _
                            ByVal B As Tensor)
    Dim i As Long
    Dim C_() As Double
    Dim A_() As Double
    Dim B_() As Double

    C.CreateFlatAlias C_
    A.CreateFlatAlias A_
    B.CreateFlatAlias B_
    For i = 1 To A.NumElements
        If C_(i) = 0 Then
            A_(i) = B_(i)
        End If
    Next i
    C.RemoveAlias C_
    A.RemoveAlias A_
    B.RemoveAlias B_
End Sub

Private Sub VecLeakyReLUNaive_I(ByVal A As Tensor, _
                                ByVal dblNegativeSlope As Double)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        If A_(i) < 0 Then
            A_(i) = dblNegativeSlope * A_(i)
        End If
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecLeakyReLUDerivativeNaive_I(ByVal A As Tensor, _
                                          ByVal dblNegativeSlope As Double)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        If A_(i) > 0 Then
            A_(i) = 1
        Else
            A_(i) = dblNegativeSlope
        End If
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecSigmoidNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = SafeSigmoid(A_(i))
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecSigmoidDerivativeNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = A_(i) * (1 - A_(i))
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecSoftplusNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = SafeSoftplus(A_(i))
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecTanhNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = SafeTanh(A_(i))
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecTanhDerivativeNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double

    A.CreateFlatAlias A_
    For i = 1 To A.NumElements
        A_(i) = 1 - A_(i) * A_(i)
    Next i
    A.RemoveAlias A_
End Sub

Private Sub VecLinCombNaive_I(ByVal dblAlpha As Double, _
                              ByVal A As Tensor, _
                              ByVal dblBeta As Double, _
                              ByVal B As Tensor)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double

    A.CreateFlatAlias A_
    B.CreateFlatAlias B_
    For i = 1 To A.NumElements
        A_(i) = dblAlpha * A_(i) + dblBeta * B_(i)
    Next i
    A.RemoveAlias A_
    B.RemoveAlias B_
End Sub

Private Sub VecLinCombBlas_I(ByVal dblAlpha As Double, _
                             ByVal A As Tensor, _
                             ByVal dblBeta As Double, _
                             ByVal B As Tensor)
    daxpby A.NumElements, dblBeta, B.Address, 1&, dblAlpha, A.Address, 1&
End Sub

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
    Select Case True
        Case Not bTransposeA And Not bTransposeB
            For i = 1 To m
                For j = 1 To n
                    dblSum = 0
                    For p = 1 To k
                        dblSum = dblSum + A_(i, p) * B_(p, j)
                    Next p
                    C_(i, j) = C_(i, j) + dblSum
                Next j
            Next i
        Case bTransposeA And Not bTransposeB
            For i = 1 To m
                For j = 1 To n
                    dblSum = 0
                    For p = 1 To k
                        dblSum = dblSum + A_(p, i) * B_(p, j)
                    Next p
                    C_(i, j) = C_(i, j) + dblSum
                Next j
            Next i
        Case Not bTransposeA And bTransposeB
            For i = 1 To m
                For j = 1 To n
                    dblSum = 0
                    For p = 1 To k
                        dblSum = dblSum + A_(i, p) * B_(j, p)
                    Next p
                    C_(i, j) = C_(i, j) + dblSum
                Next j
            Next i
        Case bTransposeA And bTransposeB
            For i = 1 To m
                For j = 1 To n
                    dblSum = 0
                    For p = 1 To k
                        dblSum = dblSum + A_(p, i) * B_(j, p)
                    Next p
                    C_(i, j) = C_(i, j) + dblSum
                Next j
            Next i
    End Select
    A.RemoveAlias A_
    B.RemoveAlias B_
    C.RemoveAlias C_
End Sub

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

Private Function TransposeNaive(ByVal A As Tensor, _
                                ByVal lDimension1 As Long, _
                                ByVal lDimension2 As Long) As Tensor
    Dim i As Long
    Dim j As Long
    Dim p As Long
    Dim q As Long
    Dim s As Long
    Dim m As Long
    Dim n As Long
    Dim lBefore As Long
    Dim lBetween As Long
    Dim lAfter As Long
    Dim alNewShape() As Long
    Dim X_() As Double
    Dim Y_() As Double
    Dim X As Tensor
    Dim Y As Tensor
    Dim R As Tensor

    lBefore = 1
    For i = 1 To lDimension1 - 1
        lBefore = lBefore * A.Size(i)
    Next i
    lBetween = 1
    For i = lDimension1 + 1 To lDimension2 - 1
        lBetween = lBetween * A.Size(i)
    Next i
    lAfter = 1
    For i = lDimension2 + 1 To A.NumDimensions
        lAfter = lAfter * A.Size(i)
    Next i
    m = A.Size(lDimension1)
    n = A.Size(lDimension2)
    alNewShape = A.Shape
    alNewShape(lDimension1) = n
    alNewShape(lDimension2) = m
    Set R = Zeros(alNewShape)
    Set X = A.View(Array(lBefore, m, lBetween, n, lAfter))
    Set Y = R.View(Array(lBefore, n, lBetween, m, lAfter))
    X.CreateAlias X_
    Y.CreateAlias Y_
    For s = 1 To lAfter
        For i = 1 To m
            For q = 1 To lBetween
                For j = 1 To n
                    For p = 1 To lBefore
                        Y_(p, j, q, i, s) = X_(p, i, q, j, s)
                    Next p
                Next j
            Next q
        Next i
    Next s
    X.RemoveAlias X_
    Y.RemoveAlias Y_
    Set TransposeNaive = R
End Function

Private Function TransposeBlas(ByVal A As Tensor) As Tensor
    Dim m As Long
    Dim n As Long
    Dim Y As Tensor

    m = A.Size(1)
    n = A.Size(2)
    Set Y = Zeros(Array(n, m))
    domatcopy "C", "T", m, n, 1#, A.Address, m, Y.Address, n
    Set TransposeBlas = Y
End Function
