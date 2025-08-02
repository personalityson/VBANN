VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCELoss"
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
    Const PROCEDURE_NAME As String = "CCELoss.Loss"
    
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
    Const PROCEDURE_NAME As String = "CCELoss.Gradient"
    
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
    Const PROCEDURE_NAME As String = "CCELoss.Hessian"
    
    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    Set Hessian = CalculateHessian(Y, T)
End Function

'L = -T * Log(Y)
Private Function CalculateLoss(ByVal Y As Tensor, _
                               ByVal T As Tensor) As Tensor
    Dim i As Long
    Dim Y_() As Double
    Dim T_() As Double
    Dim L_() As Double
    Dim L As Tensor
    
    Set L = Zeros(Y.Shape)
    Y.Flatten.CreateAlias Y_
    T.Flatten.CreateAlias T_
    L.Flatten.CreateAlias L_
    For i = 1 To Y.NumElements
        L_(i) = -T_(i) * Log(Y_(i) + DOUBLE_EPSILON)
    Next i
    Y.Flatten.RemoveAlias Y_
    T.Flatten.RemoveAlias T_
    L.Flatten.RemoveAlias L_
    Set CalculateLoss = L
End Function

'dLdY = -T / Y
Private Function CalculateGradient(ByVal Y As Tensor, _
                                   ByVal T As Tensor) As Tensor
    Dim i As Long
    Dim Y_() As Double
    Dim T_() As Double
    Dim dLdY_() As Double
    Dim dLdY As Tensor
    
    Set dLdY = Zeros(Y.Shape)
    Y.Flatten.CreateAlias Y_
    T.Flatten.CreateAlias T_
    dLdY.Flatten.CreateAlias dLdY_
    For i = 1 To Y.NumElements
        dLdY_(i) = -T_(i) / (Y_(i) + DOUBLE_EPSILON)
    Next i
    Y.Flatten.RemoveAlias Y_
    T.Flatten.RemoveAlias T_
    dLdY.Flatten.RemoveAlias dLdY_
    Set CalculateGradient = dLdY
End Function

'd^2L/dY^2 = T / Y^2
Private Function CalculateHessian(ByVal Y As Tensor, _
                                  ByVal T As Tensor) As Tensor
    Dim i As Long
    Dim Y_() As Double
    Dim T_() As Double
    Dim d2LdY2_() As Double
    Dim d2LdY2 As Tensor
    
    Set d2LdY2 = Zeros(Y.Shape)
    Y.Flatten.CreateAlias Y_
    T.Flatten.CreateAlias T_
    d2LdY2.Flatten.CreateAlias d2LdY2_
    For i = 1 To Y.NumElements
        d2LdY2_(i) = T_(i) / ((Y_(i) + DOUBLE_EPSILON) ^ 2)
    Next i
    Y.Flatten.RemoveAlias Y_
    T.Flatten.RemoveAlias T_
    d2LdY2.Flatten.RemoveAlias d2LdY2_
    Set CalculateHessian = d2LdY2
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

