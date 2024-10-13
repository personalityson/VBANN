Attribute VB_Name = "MathFunctions"
Option Explicit

Private Const OPENBLAS_PATH As String = "C:\Users\hello\OneDrive\Documents\VBANN\libopenblas.dll"

Private m_vIsBlasAvailable As Variant

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

Public Function Sigmoid(ByVal dblValue As Double) As Double
    If dblValue >= -DOUBLE_MAX_LOG Then
        Sigmoid = 1 / (1 + Exp(-dblValue))
    End If
End Function

'Y = A + B
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
        Set VecAdd = VecLinCombBlas(1, A, 1, B)
    Else
        Set VecAdd = VecAddNaive(A, B)
    End If
End Function

'Y = A + scalar
Public Function VecAddC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecAddC"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecAddC = VecLinCombBlas(1, A, 1, Full(A.Shape, dblScalar))
    Else
        Set VecAddC = VecAddCNaive(A, dblScalar)
    End If
End Function

'Y = A - B
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
        Set VecSub = VecLinCombBlas(1, A, -1, B)
    Else
        Set VecSub = VecSubNaive(A, B)
    End If
End Function

'Y = A - scalar
Public Function VecSubC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecSubC"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecSubC = VecLinCombBlas(1, A, -1, Full(A.Shape, dblScalar))
    Else
        Set VecSubC = VecSubCNaive(A, dblScalar)
    End If
End Function

'Y = scalar - A
Public Function VecSubCRev(ByVal A As Tensor, _
                           ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecSubCRev"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecSubCRev = VecLinCombBlas(-1, A, 1, Full(A.Shape, dblScalar))
    Else
        Set VecSubCRev = VecSubCRevNaive(A, dblScalar)
    End If
End Function

'Y = A .* B
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
    Set VecMul = VecMulNaive(A, B)
End Function

'Y = A .* scalar
Public Function VecMulC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecMulC"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If IsBlasAvailable() Then
        Set VecMulC = VecMulCBlas(A, dblScalar)
    Else
        Set VecMulC = VecMulCNaive(A, dblScalar)
    End If
End Function

'Y = A ./ B
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
    Set VecDiv = VecDivNaive(A, B)
End Function

'Y = A ./ scalar
Public Function VecDivC(ByVal A As Tensor, _
                        ByVal dblScalar As Double) As Tensor
    Set VecDivC = VecMulC(A, 1 / dblScalar)
End Function

'Y = scalar ./ A
Public Function VecDivCRev(ByVal A As Tensor, _
                           ByVal dblScalar As Double) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecDivCRev"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecDivCRev = VecDivCRevNaive(A, dblScalar)
End Function

'Y = A ./ (Sqrt(B) + scalar)
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
    Set VecDivSqrtAddC = VecDivSqrtAddCNaive(A, B, dblScalar)
End Function

'Y = Abs(A)
Public Function VecAbs(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecAbs"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecAbs = VecAbsNaive(A)
End Function

'Y = Sign(A)
Public Function VecSign(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecSign"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSign = VecSignNaive(A)
End Function

'Y = A .^ 2
Public Function VecPow2(ByVal A As Tensor) As Tensor
    Set VecPow2 = VecMul(A, A)
End Function

'Y = Sqrt(A)
Public Function VecSqrt(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecSqrt"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSqrt = VecSqrtNaive(A)
End Function

'Y = Exp(A)
Public Function VecExp(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecExp"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecExp = VecExpNaive(A)
End Function

'Y = Log(A)
Public Function VecLog(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecLog"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecLog = VecLogNaive(A)
End Function

'Y = 1 ./ (1 + Exp(-A))
Public Function VecSigmoid(ByVal A As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "MathFunctions.VecSigmoid"
    
    If A Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set VecSigmoid = VecSigmoidNaive(A)
End Function

'A = alpha .* A + beta .* B
Public Sub VecLinComb_I(ByVal dblAlpha As Double, _
                        ByVal A As Tensor, _
                        ByVal dblBeta As Double, _
                        ByVal B As Tensor)
    Const PROCEDURE_NAME As String = "MathFunctions.VecLinComb_I"
    
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

'Y = alpha .* A + beta .* B
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
        Set VecLinComb = VecLinCombBlas(dblAlpha, A, dblBeta, B)
    Else
        Set VecLinComb = VecLinCombNaive(dblAlpha, A, dblBeta, B)
    End If
End Function

'C = C + A * B
Public Sub MatMul_I(ByVal C As Tensor, _
                    ByVal A As Tensor, _
                    ByVal B As Tensor, _
                    Optional ByVal bTransA As Boolean, _
                    Optional ByVal bTransB As Boolean)
    Const PROCEDURE_NAME As String = "MathFunctions.MatMul_I"
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
    lNumRowsA = IIf(bTransA, A.Size(2), A.Size(1))
    lNumColsA = IIf(bTransA, A.Size(1), A.Size(2))
    lNumRowsB = IIf(bTransB, B.Size(2), B.Size(1))
    lNumColsB = IIf(bTransB, B.Size(1), B.Size(2))
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
        Err.Raise 5, PROCEDURE_NAME, "Output tensor shape does not match the expected shape for matrix multiplication."
    End If
    If IsBlasAvailable() Then
        MatMulBlas_I C, A, B, bTransA, bTransB
    Else
        MatMulNaive_I C, A, B, bTransA, bTransB
    End If
End Sub

'Y = A * B
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
        Set MatMul = MatMulBlas(A, B, bTransA, bTransB)
    Else
        Set MatMul = MatMulNaive(A, B, bTransA, bTransB)
    End If
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

Private Function VecAddNaive(ByVal A As Tensor, _
                             ByVal B As Tensor) As Tensor
    Set A = A.Clone
    VecAddNaive_I A, B
    Set VecAddNaive = A
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

Private Function VecAddCNaive(ByVal A As Tensor, _
                              ByVal dblScalar As Double) As Tensor
    Set A = A.Clone
    VecAddCNaive_I A, dblScalar
    Set VecAddCNaive = A
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

Private Function VecSubNaive(ByVal A As Tensor, _
                             ByVal B As Tensor) As Tensor
    Set A = A.Clone
    VecSubNaive_I A, B
    Set VecSubNaive = A
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

Private Function VecSubCNaive(ByVal A As Tensor, _
                              ByVal dblScalar As Double) As Tensor
    Set A = A.Clone
    VecSubCNaive_I A, dblScalar
    Set VecSubCNaive = A
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

Private Function VecSubCRevNaive(ByVal A As Tensor, _
                                 ByVal dblScalar As Double) As Tensor
    Set A = A.Clone
    VecSubCRevNaive_I A, dblScalar
    Set VecSubCRevNaive = A
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

Private Function VecMulNaive(ByVal A As Tensor, _
                             ByVal B As Tensor) As Tensor
    Set A = A.Clone
    VecMulNaive_I A, B
    Set VecMulNaive = A
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

Private Function VecMulCNaive(ByVal A As Tensor, _
                              ByVal dblScalar As Double) As Tensor
    Set A = A.Clone
    VecMulCNaive_I A, dblScalar
    Set VecMulCNaive = A
End Function

Private Sub VecMulCBlas_I(ByVal A As Tensor, _
                          ByVal dblScalar As Double)
    dscal A.NumElements, dblScalar, A.Address, 1&
End Sub

Private Function VecMulCBlas(ByVal A As Tensor, _
                             ByVal dblScalar As Double) As Tensor
    Set A = A.Clone
    VecMulCBlas_I A, dblScalar
    Set VecMulCBlas = A
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

Private Function VecDivNaive(ByVal A As Tensor, _
                             ByVal B As Tensor) As Tensor
    Set A = A.Clone
    VecDivNaive_I A, B
    Set VecDivNaive = A
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

Private Function VecDivCRevNaive(ByVal A As Tensor, _
                                 ByVal dblScalar As Double) As Tensor
    Set A = A.Clone
    VecDivCRevNaive_I A, dblScalar
    Set VecDivCRevNaive = A
End Function

Private Sub VecDivSqrtAddCNaive_I(ByVal A As Tensor, _
                                  ByVal B As Tensor, _
                                  ByVal dblScalar As Double)
    Dim i As Long
    Dim A_() As Double
    Dim B_() As Double
    
    A.Flatten.CreateAlias A_
    B.Flatten.CreateAlias B_
    For i = 1 To A.NumElements
        A_(i) = A_(i) / (Sqr(B_(i)) + dblScalar)
    Next i
    A.Flatten.RemoveAlias A_
    B.Flatten.RemoveAlias B_
End Sub

Private Function VecDivSqrtAddCNaive(ByVal A As Tensor, _
                                     ByVal B As Tensor, _
                                     ByVal dblScalar As Double) As Tensor
    Set A = A.Clone
    VecDivSqrtAddCNaive_I A, B, dblScalar
    Set VecDivSqrtAddCNaive = A
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

Private Function VecAbsNaive(ByVal A As Tensor) As Tensor
    Set A = A.Clone
    VecAbsNaive_I A
    Set VecAbsNaive = A
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

Private Function VecSignNaive(ByVal A As Tensor) As Tensor
    Set A = A.Clone
    VecSignNaive_I A
    Set VecSignNaive = A
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

Private Function VecSqrtNaive(ByVal A As Tensor) As Tensor
    Set A = A.Clone
    VecSqrtNaive_I A
    Set VecSqrtNaive = A
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

Private Function VecExpNaive(ByVal A As Tensor) As Tensor
    Set A = A.Clone
    VecExpNaive_I A
    Set VecExpNaive = A
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

Private Function VecLogNaive(ByVal A As Tensor) As Tensor
    Set A = A.Clone
    VecLogNaive_I A
    Set VecLogNaive = A
End Function

Private Sub VecSigmoidNaive_I(ByVal A As Tensor)
    Dim i As Long
    Dim A_() As Double
    
    A.Flatten.CreateAlias A_
    For i = 1 To A.NumElements
        A_(i) = Sigmoid(A_(i))
    Next i
    A.Flatten.RemoveAlias A_
End Sub

Private Function VecSigmoidNaive(ByVal A As Tensor) As Tensor
    Set A = A.Clone
    VecSigmoidNaive_I A
    Set VecSigmoidNaive = A
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

Private Function VecLinCombNaive(ByVal dblAlpha As Double, _
                                 ByVal A As Tensor, _
                                 ByVal dblBeta As Double, _
                                 ByVal B As Tensor) As Tensor
    Set A = A.Clone
    VecLinCombNaive_I dblAlpha, A, dblBeta, B
    Set VecLinCombNaive = A
End Function

Private Sub VecLinCombBlas_I(ByVal dblAlpha As Double, _
                             ByVal A As Tensor, _
                             ByVal dblBeta As Double, _
                             ByVal B As Tensor)
    daxpby A.NumElements, dblBeta, B.Address, 1&, dblAlpha, A.Address, 1&
End Sub

Private Function VecLinCombBlas(ByVal dblAlpha As Double, _
                                ByVal A As Tensor, _
                                ByVal dblBeta As Double, _
                                ByVal B As Tensor) As Tensor
    Set A = A.Clone
    VecLinCombBlas_I dblAlpha, A, dblBeta, B
    Set VecLinCombBlas = A
End Function

Private Sub MatMulNaive_I(ByVal C As Tensor, _
                          ByVal A As Tensor, _
                          ByVal B As Tensor, _
                          ByVal bTransA As Boolean, _
                          ByVal bTransB As Boolean)
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
    
    m = C.Size(1)
    n = C.Size(2)
    k = IIf(bTransA, A.Size(1), A.Size(2))
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
            C_(i, j) = C_(i, j) + dblSum
        Next j
    Next i
    A.RemoveAlias A_
    B.RemoveAlias B_
    C.RemoveAlias C_
End Sub

Private Function MatMulNaive(ByVal A As Tensor, _
                             ByVal B As Tensor, _
                             ByVal bTransA As Boolean, _
                             ByVal bTransB As Boolean) As Tensor
    Dim m As Long
    Dim n As Long
    Dim C As Tensor
    
    m = IIf(bTransA, A.Size(2), A.Size(1))
    n = IIf(bTransB, B.Size(1), B.Size(2))
    Set C = Zeros(Array(m, n))
    MatMulNaive_I C, A, B, bTransA, bTransB
    Set MatMulNaive = C
End Function

Private Sub MatMulBlas_I(ByVal C As Tensor, _
                         ByVal A As Tensor, _
                         ByVal B As Tensor, _
                         ByVal bTransA As Boolean, _
                         ByVal bTransB As Boolean)
    Dim sTransA As String
    Dim sTransB As String
    Dim m As Long
    Dim n As Long
    Dim k As Long
    
    sTransA = IIf(bTransA, "T", "N")
    sTransB = IIf(bTransB, "T", "N")
    m = C.Size(1)
    n = C.Size(2)
    k = IIf(bTransA, A.Size(1), A.Size(2))
    dgemm sTransA, sTransB, m, n, k, 1#, A.Address, A.Size(1), B.Address, B.Size(1), 1#, C.Address, m
End Sub

Private Function MatMulBlas(ByVal A As Tensor, _
                            ByVal B As Tensor, _
                            ByVal bTransA As Boolean, _
                            ByVal bTransB As Boolean) As Tensor
    Dim m As Long
    Dim n As Long
    Dim C As Tensor
    
    m = IIf(bTransA, A.Size(2), A.Size(1))
    n = IIf(bTransB, B.Size(1), B.Size(2))
    Set C = Zeros(Array(m, n))
    MatMulBlas_I C, A, B, bTransA, bTransB
    Set MatMulBlas = C
End Function
