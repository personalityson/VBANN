Attribute VB_Name = "TensorInit"
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
                                Optional ByVal bTranspose As Boolean) As Tensor
    Set TensorFromRange = New Tensor
    TensorFromRange.FromRange oRange, bTranspose
End Function

Public Function TensorFromArray(ByRef adblArray() As Double) As Tensor
    Set TensorFromArray = New Tensor
    TensorFromArray.FromArray adblArray
End Function

Private Function NormRand() As Double
    NormRand = Sqr(-2 * Log(Rnd() + DOUBLE_EPSILON)) * Cos(MATH_2PI * Rnd())
End Function
