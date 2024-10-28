Attribute VB_Name = "TensorFunctions"
Option Explicit

Public Function Zeros(ByVal vShape As Variant) As Tensor
    Set Zeros = New Tensor
    Zeros.Resize vShape
End Function

Public Function Ones(ByVal vShape As Variant) As Tensor
    Set Ones = New Tensor
    With Ones
        .Resize vShape
        .Fill 1
    End With
End Function

Public Function Full(ByVal vShape As Variant, _
                     ByVal dblValue As Double) As Tensor
    Set Full = New Tensor
    With Full
        .Resize vShape
        .Fill dblValue
    End With
End Function

Public Function Uniform(ByVal vShape As Variant, _
                        Optional ByVal dblLow As Double = 0, _
                        Optional ByVal dblHigh As Double = 1) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set Uniform = New Tensor
    With Uniform
        .Resize vShape
        .Flatten.CreateAlias A_
        For i = 1 To .NumElements
            A_(i) = dblLow + (dblHigh - dblLow) * Rnd()
        Next i
        .Flatten.RemoveAlias A_
    End With
End Function

Public Function Normal(ByVal vShape As Variant, _
                       Optional ByVal dblMu As Double = 0, _
                       Optional ByVal dblSigma As Double = 1) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set Normal = New Tensor
    With Normal
        .Resize vShape
        .Flatten.CreateAlias A_
        For i = 1 To .NumElements
            A_(i) = dblMu + dblSigma * NormRand()
        Next i
        .Flatten.RemoveAlias A_
    End With
End Function

Public Function Bernoulli(ByVal vShape As Variant, _
                          Optional ByVal dblProbability As Double = 0.5) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set Bernoulli = New Tensor
    With Bernoulli
        .Resize vShape
        .Flatten.CreateAlias A_
        For i = 1 To .NumElements
            A_(i) = -(Rnd() < dblProbability)
        Next i
        .Flatten.RemoveAlias A_
    End With
End Function

Public Function GlorotUniform(ByVal vShape As Variant, _
                              ByVal lInputSize As Long, _
                              ByVal lOutputSize As Long, _
                              Optional ByVal dblGain As Double = 1) As Tensor
    Dim dblLimit As Double
    
    dblLimit = dblGain * Sqr(6 / (lInputSize + lOutputSize))
    Set GlorotUniform = Uniform(vShape, -dblLimit, dblLimit)
End Function

Public Function GlorotNormal(ByVal vShape As Variant, _
                             ByVal lInputSize As Long, _
                             ByVal lOutputSize As Long, _
                             Optional ByVal dblGain As Double = 1) As Tensor
    Dim dblSigma As Double

    dblSigma = dblGain * Sqr(2 / (lInputSize + lOutputSize))
    Set GlorotNormal = Normal(vShape, 0, dblSigma)
End Function

Public Function HeUniform(ByVal vShape As Variant, _
                          ByVal lInputSize As Long, _
                          Optional ByVal dblGain As Double = 1) As Tensor
    Dim dblLimit As Double

    dblLimit = dblGain * Sqr(6 / (lInputSize))
    Set HeUniform = Uniform(vShape, -dblLimit, dblLimit)
End Function

Public Function HeNormal(ByVal vShape As Variant, _
                         ByVal lInputSize As Long, _
                         Optional ByVal dblGain As Double = 1) As Tensor
    Dim dblSigma As Double

    dblSigma = dblGain * Sqr(2 / lInputSize)
    Set HeNormal = Normal(vShape, 0, dblSigma)
End Function

Public Function TensorFromRange(ByVal oRange As Range, _
                                ByVal bTrans As Boolean) As Tensor
    Set TensorFromRange = New Tensor
    TensorFromRange.FromRange oRange, bTrans
End Function

Public Function TensorFromArray(ByRef adblArray() As Double) As Tensor
    Set TensorFromArray = New Tensor
    TensorFromArray.FromArray adblArray
End Function

Public Function Concacenate(ByVal lDimension As Long, _
                            ByVal vTensors As Variant) As Tensor
    Const PROCEDURE_NAME As String = "Tensor.Gather"
    Dim lNumIndices As Long
    Dim alIndices() As Long
    Dim i As Long
    
'    If lDimension < 1 Or lDimension > m_lNumDimensions Then
'        Err.Raise 9, PROCEDURE_NAME, "Dimension index is out of range."
'    End If
'    ParseVariantToLongArray vIndices, lNumIndices, alIndices
'    If UBound(alShapeA) <> UBound(alShapeB) Then
'        Err.Raise 5, PROCEDURE_NAME, "Tensors must have the same number of dimensions."
'    End If
'
'    For i = 1 To UBound(alShapeA)
'        If i <> lDimension And alShapeA(i) <> alShapeB(i) Then
'            Err.Raise 5, PROCEDURE_NAME, "Shapes must match except on the concatenation dimension."
'        End If
'    Next i
End Function


Public Sub ParseVariantToTensorArray(ByVal vValueOrArray As Variant, _
                                     ByRef lNumTensors As Long, _
                                     ByRef aoTensors() As Long)
    Const PROCEDURE_NAME As String = "Tensor.ParseVariantToTensorArray"
    Dim lRank As Long
    Dim lLBound As Long
    Dim lUBound As Long
    Dim i As Long
    
    lRank = GetRank(vValueOrArray)
    Select Case lRank
        Case -1
            lNumElements = 1
            ReDim aoTensors(1 To lNumElements)
            Set aoTensors(1) = CLng(vValueOrArray)
        Case 0
            lNumElements = 0
            Erase aoTensors
        Case 1
            lLBound = LBound(vValueOrArray)
            lUBound = UBound(vValueOrArray)
            If lLBound > lUBound Then
                lNumElements = 0
                Erase aoTensors
            Else
                lNumElements = lUBound - lLBound + 1
                ReDim aoTensors(1 To lNumElements)
                For i = 1 To lNumElements
                    Set aoTensors(i) = vValueOrArray(lLBound + i - 1)
                Next i
            End If
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Expected a single value, an uninitialized array, or a one-dimensional array."
    End Select
End Sub
