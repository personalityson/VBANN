VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HuberLoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILossFunction
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Hyperparameters
Private m_dblDelta As Double

Public Property Get FromLogits() As Boolean
    FromLogits = True
End Property

Public Property Get Delta() As Double
    Delta = m_dblDelta
End Property

Public Sub Init(Optional ByVal dblDelta As Double = 1)
    Const PROCEDURE_NAME As String = "HuberLoss.Init"

    If dblDelta <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Delta must be > 0."
    End If
    m_dblDelta = dblDelta
End Sub

'L = 0.5 * R^2 if |R| <= delta, else delta * (|R| - 0.5 * delta), R = Y - T
'With C = Clamp(R, -Delta, Delta) both cases are L = C * (R - 0.5 * C)
Public Function Loss(ByVal Y As Tensor, _
                     ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "HuberLoss.Loss"
    Dim R As Tensor
    Dim C As Tensor

    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    Set R = VecSub(Y, T)
    Set C = VecClamp(R, -m_dblDelta, m_dblDelta)
    Set Loss = VecMul(C, VecLinComb(1, R, -0.5, C))
End Function

'dL/dY = Clamp(Y - T, -Delta, Delta)
Public Function Gradient(ByVal Y As Tensor, _
                         ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "HuberLoss.Gradient"

    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    Set Gradient = VecClamp(VecSub(Y, T), -m_dblDelta, m_dblDelta)
End Function

'd2L/dY2 = 1 for XGBoost, the true second derivative is 1 inside delta and 0 outside
Public Function Hessian(ByVal Y As Tensor, _
                        ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "HuberLoss.Hessian"

    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    Set Hessian = Ones(Y.Shape)
End Function

'Median over samples, robust to the outliers Huber is meant for
Public Function Intercept(ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "HuberLoss.Intercept"

    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T.NumDimensions < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Target tensor must have at least one dimension."
    End If
    If T.Size(T.NumDimensions) < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Target tensor must contain at least one sample."
    End If
    Set Intercept = T.Reduce(T.NumDimensions, rtMedian)
End Function

Public Function Transform(ByVal Y As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "HuberLoss.Transform"

    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Transform = Y
End Function

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "HuberLoss.Save"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteDouble m_dblDelta
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "HuberLoss.Load"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        m_dblDelta = .ReadDouble()
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

