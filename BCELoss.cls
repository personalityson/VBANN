VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BCELoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILossFunction
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

Private Const BCE_EPSILON As Double = 0.0000001

'Configuration
Private m_bFromLogits As Boolean

Public Property Get FromLogits() As Boolean
    FromLogits = m_bFromLogits
End Property

Public Sub Init(Optional ByVal bFromLogits As Boolean)
    m_bFromLogits = bFromLogits
End Sub

'Logit:       L = Softplus(Y) - T * Y
'Probability: L = -T * Log(Y) - (1 - T) * Log(1 - Y)
Public Function Loss(ByVal Y As Tensor, _
                     ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "BCELoss.Loss"

    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Y.NumDimensions < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Output tensor must have at least one dimension."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    If m_bFromLogits Then
        Set Loss = VecSub(VecSoftplus(Y), VecMul(Y, T))
    Else
        Set Loss = CalculateLossFromProbabilities(Y, T)
    End If
End Function

'Logit:       dL/dY = Sigmoid(Y) - T
'Probability: dL/dY = (Y - T) / (Y * (1 - Y))
Public Function Gradient(ByVal Y As Tensor, _
                         ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "BCELoss.Gradient"

    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Y.NumDimensions < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Output tensor must have at least one dimension."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    If m_bFromLogits Then
        Set Gradient = VecSub(VecSigmoid(Y), T)
    Else
        Set Gradient = CalculateGradientFromProbabilities(Y, T)
    End If
End Function

'Logit:       d2L/dY2 = Sigmoid(Y) * (1 - Sigmoid(Y))
'Probability: d2L/dY2 = T / Y^2 + (1 - T) / (1 - Y)^2
Public Function Hessian(ByVal Y As Tensor, _
                        ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "BCELoss.Hessian"

    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Y.NumDimensions < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Output tensor must have at least one dimension."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    If m_bFromLogits Then
        Set Hessian = VecSigmoidDerivative(VecSigmoid(Y))
    Else
        Set Hessian = CalculateHessianFromProbabilities(Y, T)
    End If
End Function

'Log(M / (1 - M)), M is the mean over samples
Public Function Intercept(ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "BCELoss.Intercept"
    Dim oClamped As Tensor

    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T.NumDimensions < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Target tensor must have at least one dimension."
    End If
    If T.Size(T.NumDimensions) < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Target tensor must contain at least one sample."
    End If
    Set oClamped = VecClamp(T.Reduce(T.NumDimensions, rtMean), BCE_EPSILON, 1 - BCE_EPSILON)
    Set Intercept = VecLog(VecDiv(oClamped, VecSubCRev(oClamped, 1)))
End Function

Public Function Transform(ByVal Y As Tensor) As Tensor
    Set Transform = VecSigmoid(Y)
End Function

Private Function CalculateLossFromProbabilities(ByVal Y As Tensor, _
                                                ByVal T As Tensor) As Tensor
    Dim i As Long
    Dim dblY As Double
    Dim Y_() As Double
    Dim T_() As Double
    Dim L_() As Double
    Dim L As Tensor
    
    Set L = Zeros(Y.Shape)
    Y.CreateFlatAlias Y_
    T.CreateFlatAlias T_
    L.CreateFlatAlias L_
    For i = 1 To Y.NumElements
        dblY = ClampDbl(Y_(i), BCE_EPSILON, 1 - BCE_EPSILON)
        L_(i) = -T_(i) * Log(dblY) - (1 - T_(i)) * Log(1 - dblY)
    Next i
    Y.RemoveAlias Y_
    T.RemoveAlias T_
    L.RemoveAlias L_
    Set CalculateLossFromProbabilities = L
End Function

Private Function CalculateGradientFromProbabilities(ByVal Y As Tensor, _
                                                    ByVal T As Tensor) As Tensor
    Dim i As Long
    Dim dblY As Double
    Dim Y_() As Double
    Dim T_() As Double
    Dim dLdY_() As Double
    Dim dLdY As Tensor
    
    Set dLdY = Zeros(Y.Shape)
    Y.CreateFlatAlias Y_
    T.CreateFlatAlias T_
    dLdY.CreateFlatAlias dLdY_
    For i = 1 To Y.NumElements
        dblY = ClampDbl(Y_(i), BCE_EPSILON, 1 - BCE_EPSILON)
        dLdY_(i) = (dblY - T_(i)) / (dblY * (1 - dblY))
    Next i
    Y.RemoveAlias Y_
    T.RemoveAlias T_
    dLdY.RemoveAlias dLdY_
    Set CalculateGradientFromProbabilities = dLdY
End Function

Private Function CalculateHessianFromProbabilities(ByVal Y As Tensor, _
                                                   ByVal T As Tensor) As Tensor
    Dim i As Long
    Dim dblY As Double
    Dim Y_() As Double
    Dim T_() As Double
    Dim d2LdY2_() As Double
    Dim d2LdY2 As Tensor
    
    Set d2LdY2 = Zeros(Y.Shape)
    Y.CreateFlatAlias Y_
    T.CreateFlatAlias T_
    d2LdY2.CreateFlatAlias d2LdY2_
    For i = 1 To Y.NumElements
        dblY = ClampDbl(Y_(i), BCE_EPSILON, 1 - BCE_EPSILON)
        d2LdY2_(i) = T_(i) / (dblY * dblY) + (1 - T_(i)) / ((1 - dblY) * (1 - dblY))
    Next i
    Y.RemoveAlias Y_
    T.RemoveAlias T_
    d2LdY2.RemoveAlias d2LdY2_
    Set CalculateHessianFromProbabilities = d2LdY2
End Function

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "BCELoss.Save"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteBoolean m_bFromLogits
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "BCELoss.Load"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        m_bFromLogits = .ReadBoolean()
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get ILossFunction_FromLogits() As Boolean
        ILossFunction_FromLogits = Me.FromLogits
    End Property

    Private Function ILossFunction_Loss(ByVal Y As Object, _
                                        ByVal T As Object) As Object
        Set ILossFunction_Loss = Me.Loss(Y, T)
    End Function

    Private Function ILossFunction_Gradient(ByVal Y As Object, _
                                            ByVal T As Object) As Object
        Set ILossFunction_Gradient = Me.Gradient(Y, T)
    End Function

    Private Function ILossFunction_Hessian(ByVal Y As Object, _
                                           ByVal T As Object) As Object
        Set ILossFunction_Hessian = Me.Hessian(Y, T)
    End Function

    Private Function ILossFunction_Intercept(ByVal T As Object) As Object
        Set ILossFunction_Intercept = Me.Intercept(T)
    End Function

    Private Function ILossFunction_Transform(ByVal Y As Object) As Object
        Set ILossFunction_Transform = Me.Transform(Y)
    End Function

    Private Sub ISerializable_Save(ByVal oSerializer As Object)
        Save oSerializer
    End Sub

    Private Sub ISerializable_Load(ByVal oSerializer As Object)
        Load oSerializer
    End Sub
#Else
    Private Property Get ILossFunction_FromLogits() As Boolean
        ILossFunction_FromLogits = Me.FromLogits
    End Property

    Private Function ILossFunction_Loss(ByVal Y As Tensor, _
                                        ByVal T As Tensor) As Tensor
        Set ILossFunction_Loss = Me.Loss(Y, T)
    End Function

    Private Function ILossFunction_Gradient(ByVal Y As Tensor, _
                                            ByVal T As Tensor) As Tensor
        Set ILossFunction_Gradient = Me.Gradient(Y, T)
    End Function

    Private Function ILossFunction_Hessian(ByVal Y As Tensor, _
                                           ByVal T As Tensor) As Tensor
        Set ILossFunction_Hessian = Me.Hessian(Y, T)
    End Function

    Private Function ILossFunction_Intercept(ByVal T As Tensor) As Tensor
        Set ILossFunction_Intercept = Me.Intercept(T)
    End Function

    Private Function ILossFunction_Transform(ByVal Y As Tensor) As Tensor
        Set ILossFunction_Transform = Me.Transform(Y)
    End Function

    Private Sub ISerializable_Save(ByVal oSerializer As Serializer)
        Save oSerializer
    End Sub

    Private Sub ISerializable_Load(ByVal oSerializer As Serializer)
        Load oSerializer
    End Sub
#End If

