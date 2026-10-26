VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "L2Loss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILossFunction
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

Public Property Get FromLogits() As Boolean
    FromLogits = True
End Property

Public Sub Init()
    'Do nothing
End Sub

'L = 0.5 * (Y - T) ^ 2
Public Function Loss(ByVal Y As Tensor, _
                     ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "L2Loss.Loss"
    
    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    Set Loss = VecMulC(VecPow2(VecSub(Y, T)), 0.5)
End Function

'dL/dY = Y - T
Public Function Gradient(ByVal Y As Tensor, _
                         ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "L2Loss.Gradient"
    
    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    Set Gradient = VecSub(Y, T)
End Function

'd2L/dY2 = 1
Public Function Hessian(ByVal Y As Tensor, _
                        ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "L2Loss.Hessian"
    
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

'Mean over samples
Public Function Intercept(ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "L2Loss.Intercept"

    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T.NumDimensions < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Target tensor must have at least one dimension."
    End If
    If T.Size(T.NumDimensions) < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Target tensor must contain at least one sample."
    End If
    Set Intercept = T.Reduce(T.NumDimensions, rtMean)
End Function

Public Function Transform(ByVal Y As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "L2Loss.Transform"

    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Transform = Y
End Function

Private Sub Save(ByVal oSerializer As Serializer)
    'Do nothing
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    'Do nothing
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

