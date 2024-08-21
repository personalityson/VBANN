Attribute VB_Name = "MathFunctions"
Option Explicit

Private Const OPENBLAS_PATH As String = "C:\Users\hello\OneDrive\Documents\VBANN\libopenblas.dll"

Private m_vIsBlasAvailable As Variant

Public Declare PtrSafe Sub dscal Lib "libopenblas.dll" (ByRef n As Long, _
                                                        ByRef alpha As Double, _
                                                        ByVal X As LongPtr, _
                                                        ByRef incX As Long)

Public Declare PtrSafe Sub daxpby Lib "libopenblas.dll" (ByRef n As Long, _
                                                         ByRef alpha As Double, _
                                                         ByVal X As LongPtr, _
                                                         ByRef incX As Long, _
                                                         ByRef beta As Double, _
                                                         ByVal Y As LongPtr, _
                                                         ByRef incY As Long)

Public Declare PtrSafe Sub dgemm Lib "libopenblas.dll" (ByVal transA As String, _
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

Public Function IsBlasAvailable() As Boolean
    If IsEmpty(m_vIsBlasAvailable) Then
        m_vIsBlasAvailable = Fso.FileExists(OPENBLAS_PATH)
        If m_vIsBlasAvailable Then
            ChDir Fso.GetParentFolderName(OPENBLAS_PATH)
        End If
    End If
    IsBlasAvailable = m_vIsBlasAvailable
End Function

Public Function NormRand() As Double
    NormRand = Sqr(-2 * Log(Rnd() + DOUBLE_MIN_ABS)) * Cos(MATH_2PI * Rnd())
End Function

Public Function Sigmoid(ByVal dblValue As Double) As Double
    If dblValue >= -DOUBLE_MAX_LOG Then
        Sigmoid = 1 / (1 + Exp(-dblValue))
    End If
End Function

Public Function VecAdd(ByVal A As Tensor, _
                       ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecAdd"
    
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
        Set VecAdd = ComputeLinCombWithBlas(1, A, 1, B)
    Else
        Set VecAdd = ComputeAddNaively(A, B)
    End If
End Function

Public Function VecAddC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecAddC"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecAddC = ComputeLinCombWithBlas(1, A, 1, Full(A.Shape, dblScalar))
    Else
        Set VecAddC = ComputeAddCNaively(A, dblScalar)
    End If
End Function

Public Function VecSub(ByVal A As Tensor, _
                       ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecSub"
    
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
        Set VecSub = ComputeLinCombWithBlas(1, A, -1, B)
    Else
        Set VecSub = ComputeSubNaively(A, B)
    End If
End Function

Public Function VecSubC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecSubC"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecSubC = ComputeLinCombWithBlas(1, A, -1, Full(A.Shape, dblScalar))
    Else
        Set VecSubC = ComputeSubCNaively(A, dblScalar)
    End If
End Function

Public Function VecSubCRev(ByVal dblScalar As Double, _
                           ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecSubCRev"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecSubCRev = ComputeLinCombWithBlas(-1, A, 1, Full(A.Shape, dblScalar))
    Else
        Set VecSubCRev = ComputeSubCRevNaively(dblScalar, A)
    End If
End Function

Public Function VecMul(ByVal A As Tensor, _
                       ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecMul"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    Set VecMul = ComputeMulNaively(A, B)
End Function

Public Function VecMulC(ByVal dblScalar As Double, _
                        ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecMulC"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecMulC = ComputeMulCWithBlas(dblScalar, A)
    Else
        Set VecMulC = ComputeMulCNaively(dblScalar, A)
    End If
End Function

Public Function VecDiv(ByVal A As Tensor, _
                       ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecDiv"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    Set VecDiv = ComputeDivNaively(A, B)
End Function

Public Function VecDivC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Set VecDivC = VecMulC(1 / dblScalar, A)
End Function

Public Function VecDivCRev(ByVal dblScalar As Double, _
                           ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecDivCRev"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecDivCRev = ComputeDivCRevNaively(dblScalar, A)
End Function

Public Function VecSqrt(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecSqrt"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSqrt = ComputeSqrtNaively(A)
End Function

Public Function VecDivSqrtAddC(ByVal A As Tensor, _
                               ByVal B As Tensor, _
                               ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecDivSqrtAddC"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If B Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If A.NumElements <> B.NumElements Then
        Err.Raise 5, PROCEDURE_NAME, "Tensors A and B must have the same number of elements."
    End If
    Set VecDivSqrtAddC = ComputeDivSqrtAddCNaively(A, B, dblScalar)
End Function

Public Function VecPow2(ByVal A As Tensor) As Tensor
    Set VecPow2 = VecMul(A, A)
End Function

Public Function VecAbs(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecAbs"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecAbs = ComputeAbsNaively(A)
End Function

Public Function VecSign(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecSign"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSign = ComputeSignNaively(A)
End Function

Public Function VecSigmoid(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecSigmoid"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSigmoid = ComputeSigmoidNaively(A)
End Function

Public Function VecLinComb(ByVal dblAlpha As Double, _
                           ByVal A As Tensor, _
                           ByVal dblBeta As Double, _
                           ByVal B As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecLinComb"
    
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
        Set VecLinComb = ComputeLinCombWithBlas(dblAlpha, A, dblBeta, B)
    Else
        Set VecLinComb = ComputeLinCombNaively(dblAlpha, A, dblBeta, B)
    End If
End Function

Public Function MatMul(ByVal A As Tensor, _
                       ByVal B As Tensor, _
                       Optional ByVal bTransA As Boolean, _
                       Optional ByVal bTransB As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.MatMul"
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
    lNumColsA = IIf(bTransA, A.Size(1), A.Size(2))
    lNumRowsB = IIf(bTransB, B.Size(2), B.Size(1))
    If lNumColsA <> lNumRowsB Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of tensors A and B are incompatible for matrix multiplication."
    End If
    If IsBlasAvailable() Then
        Set MatMul = ComputeMatMulWithBlas(A, B, bTransA, bTransB)
    Else
        Set MatMul = ComputeMatMulNaively(A, B, bTransA, bTransB)
    End If
End Function

Private Function ComputeAddNaively(ByVal A As Tensor, _
                                   ByVal B As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) + B_(i)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
    Set ComputeAddNaively = A
End Function

Private Function ComputeAddCNaively(ByVal A As Tensor, _
                                    ByVal dblScalar As Double) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    A = A.Clone
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = A_(i) + dblScalar
    Next i
    A.Flatten.RemoveAlias A_
    Set ComputeAddCNaively = A
End Function

Private Function ComputeSubNaively(ByVal A As Tensor, _
                                   ByVal B As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) - B_(i)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
    Set ComputeSubNaively = A
End Function

Private Function ComputeSubCNaively(ByVal A As Tensor, _
                                    ByVal dblScalar As Double) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = A_(i) - dblScalar
    Next i
    A.Flatten.RemoveAlias A_
    Set ComputeSubCNaively = A
End Function

Private Function ComputeSubCRevNaively(ByVal dblScalar As Double, _
                                       ByVal A As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = dblScalar - A_(i)
    Next i
    A.Flatten.RemoveAlias A_
    Set ComputeSubCRevNaively = A
End Function

Private Function ComputeMulNaively(ByVal A As Tensor, _
                                   ByVal B As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) * B_(i)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
    Set ComputeMulNaively = A
End Function

Private Function ComputeMulCNaively(ByVal dblScalar As Double, _
                                    ByVal A As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = dblScalar * A_(i)
    Next i
    A.Flatten.RemoveAlias A_
    Set ComputeMulCNaively = A
End Function

Private Function ComputeMulCWithBlas(ByVal dblScalar As Double, _
                                     ByVal A As Tensor) As Tensor
    Set A = A.Clone
    dscal A.NumElements, dblScalar, A.Address, 1&
    Set ComputeMulCWithBlas = A
End Function

Private Function ComputeDivNaively(ByVal A As Tensor, _
                                   ByVal B As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) / B_(i)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
    Set ComputeDivNaively = A
End Function

Private Function ComputeDivCRevNaively(ByVal dblScalar As Double, _
                                       ByVal A As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = dblScalar / A_(i)
    Next i
    A.Flatten.RemoveAlias A_
    Set ComputeDivCRevNaively = A
End Function

Private Function ComputeSqrtNaively(ByVal A As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = Sqr(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
    Set ComputeSqrtNaively = A
End Function

Private Function ComputeDivSqrtAddCNaively(ByVal A As Tensor, _
                                           ByVal B As Tensor, _
                                           ByVal dblScalar As Double) As Tensor
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) / (Sqr(B_(i)) + dblScalar)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
    Set ComputeDivSqrtAddCNaively = A
End Function

Private Function ComputeAbsNaively(ByVal A As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = Abs(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
    Set ComputeAbsNaively = A
End Function

Private Function ComputeSignNaively(ByVal A As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = Sgn(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
    Set ComputeSignNaively = A
End Function

Private Function ComputeSigmoidNaively(ByVal A As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = Sigmoid(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
    Set ComputeSigmoidNaively = A
End Function

Private Function ComputeLinCombNaively(ByVal dblAlpha As Double, _
                                       ByVal A As Tensor, _
                                       ByVal dblBeta As Double, _
                                       ByVal B As Tensor) As Tensor
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    Set A = A.Clone
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = dblAlpha * A_(i) + dblBeta * B_(i)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
    Set ComputeLinCombNaively = A
End Function

Private Function ComputeLinCombWithBlas(ByVal dblAlpha As Double, _
                                        ByVal A As Tensor, _
                                        ByVal dblBeta As Double, _
                                        ByVal B As Tensor) As Tensor
    Set A = A.Clone
    daxpby A.NumElements, dblBeta, B.Address, 1&, dblAlpha, A.Address, 1&
    Set ComputeLinCombWithBlas = A
End Function

Private Function ComputeMatMulNaively(ByVal A As Tensor, _
                                      ByVal B As Tensor, _
                                      ByVal bTransA As Boolean, _
                                      ByVal bTransB As Boolean) As Tensor
    Dim m As Long
    Dim n As Long
    Dim k As Long
    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim dblSum As Double
    Dim A_() As Double
    Dim B_() As Double
    Dim C_() As Double
    Dim C As Tensor
    
    m = IIf(bTransA, A.Size(2), A.Size(1))
    k = IIf(bTransA, A.Size(1), A.Size(2))
    n = IIf(bTransB, B.Size(1), B.Size(2))
    Set C = Zeros(Array(m, n))
    A.CreateAlias A_
    B.CreateAlias B_
    C.CreateAlias C_
    For i = 1 To m
        For j = 1 To n
            dblSum = 0
            Select Case True
                Case Not bTransA And Not bTransB
                    For l = 1 To k
                        dblSum = dblSum + A_(i, l) * B_(l, j)
                    Next l
                Case bTransA And Not bTransB
                    For l = 1 To k
                        dblSum = dblSum + A_(l, i) * B_(l, j)
                    Next l
                Case Not bTransA And bTransB
                    For l = 1 To k
                        dblSum = dblSum + A_(i, l) * B_(j, l)
                    Next l
                Case bTransA And bTransB
                    For l = 1 To k
                        dblSum = dblSum + A_(l, i) * B_(j, l)
                    Next l
            End Select
            C_(i, j) = dblSum
        Next j
    Next i
    A.RemoveAlias A_
    B.RemoveAlias B_
    C.RemoveAlias C_
    Set ComputeMatMulNaively = C
End Function

Private Function ComputeMatMulWithBlas(ByVal A As Tensor, _
                                       ByVal B As Tensor, _
                                       ByVal bTransA As Boolean, _
                                       ByVal bTransB As Boolean) As Tensor
    Dim sTransA As String
    Dim sTransB As String
    Dim m As Long
    Dim n As Long
    Dim k As Long
    Dim C As Tensor
    
    sTransA = IIf(bTransA, "T", "N")
    sTransB = IIf(bTransB, "T", "N")
    m = IIf(bTransA, A.Size(2), A.Size(1))
    k = IIf(bTransA, A.Size(1), A.Size(2))
    n = IIf(bTransB, B.Size(1), B.Size(2))
    Set C = Zeros(Array(m, n))
    dgemm sTransA, sTransB, m, n, k, 1#, A.Address, A.Size(1), B.Address, B.Size(1), 1#, C.Address, C.Size(1)
    Set ComputeMatMulWithBlas = C
End Function
