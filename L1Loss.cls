VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "L1Loss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICriterion
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

Public Sub Init()
    'Do nothing
End Sub

Public Function Loss(ByVal Y As Tensor, _
                     ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "L1Loss.Loss"
    
    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    Set Loss = CalculateLoss(Y, T)
End Function

Public Function Gradient(ByVal Y As Tensor, _
                         ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "L1Loss.Gradient"
    
    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    Set Gradient = CalculateGradient(Y, T)
End Function

Public Function Hessian(ByVal Y As Tensor, _
                        ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "L1Loss.Hessian"
    
    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set Hessian = CalculateHessian(Y)
End Function

'L = Abs(Y - T)
Private Function CalculateLoss(ByVal Y As Tensor, _
                               ByVal T As Tensor) As Tensor
    Set CalculateLoss = VecAbs(VecSub(Y, T))
End Function

'dL/dY = Sign(Y - T)
Private Function CalculateGradient(ByVal Y As Tensor, _
                                   ByVal T As Tensor) As Tensor
    Set CalculateGradient = VecSign(VecSub(Y, T))
End Function

'd^2L/dY^2 = 0
Private Function CalculateHessian(ByVal Y As Tensor) As Tensor
    Set CalculateHessian = Zeros(Y.Shape)
End Function

Private Sub Save(ByVal oWorksheet As Serializer)
    'Do nothing
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    'Do nothing
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Function ICriterion_Loss(ByVal Y As Object, _
                                        ByVal T As Object) As Object
        Set ICriterion_Loss = Loss(Y, T)
    End Function
    
    Private Function ICriterion_Gradient(ByVal Y As Object, _
                                         ByVal T As Object) As Object
        Set ICriterion_Gradient = Gradient(Y, T)
    End Function
    
    Private Function ICriterion_Hessian(ByVal Y As Object, _
                                        ByVal T As Object) As Object
        Set ICriterion_Hessian = Hessian(Y, T)
    End Function
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Object)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Object)
        Load oWorksheet
    End Sub
#Else
    Private Function ICriterion_Loss(ByVal Y As Tensor, _
                                        ByVal T As Tensor) As Tensor
        Set ICriterion_Loss = Loss(Y, T)
    End Function
    
    Private Function ICriterion_Gradient(ByVal Y As Tensor, _
                                         ByVal T As Tensor) As Tensor
        Set ICriterion_Gradient = Gradient(Y, T)
    End Function
    
    Private Function ICriterion_Hessian(ByVal Y As Tensor, _
                                        ByVal T As Tensor) As Tensor
        Set ICriterion_Hessian = Hessian(Y, T)
    End Function
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If

