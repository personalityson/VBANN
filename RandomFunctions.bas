Attribute VB_Name = "RandomFunctions"
Option Explicit

Public Function NormRand() As Double
    NormRand = Sqr(-2 * Log(Rnd() + DOUBLE_MIN_ABS)) * Cos(MATH_2PI * Rnd())
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

