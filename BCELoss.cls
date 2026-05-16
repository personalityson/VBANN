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

Implements ICriterion
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

Private Const EPSILON As Double = 0.0000001

Public Sub Init()
    'Do nothing
End Sub

Public Function Forward(ByVal Y As Tensor, _
                        ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "BCELoss.Forward"
    
    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    Set Forward = CalculateLoss(Y, T)
End Function

Public Function Backward(ByVal Y As Tensor, _
                         ByVal T As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "BCELoss.Backward"
    
    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not Y.ShapeEquals(T.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Output and target tensors must have the same shape."
    End If
    Set Backward = CalculateGradient(Y, T)
End Function

'L = -T * Log(Y) - (1 - T) * Log(1 - Y)
Private Function CalculateLoss(ByVal Y As Tensor, _
                               ByVal T As Tensor) As Tensor
    Dim i As Long
    Dim dblY As Double
    Dim Y_() As Double
    Dim T_() As Double
    Dim L_() As Double
    Dim L As Tensor
    
    Set L = Zeros(Y.Shape)
    Y.Flatten.CreateAlias Y_
    T.Flatten.CreateAlias T_
    L.Flatten.CreateAlias L_
    For i = 1 To Y.NumElements
        dblY = MinDbl(1 - EPSILON, MaxDbl(EPSILON, Y_(i)))
        L_(i) = -T_(i) * Log(dblY) - (1 - T_(i)) * Log(1 - dblY)
    Next i
    Y.Flatten.RemoveAlias Y_
    T.Flatten.RemoveAlias T_
    L.Flatten.RemoveAlias L_
    Set CalculateLoss = L
End Function

'dL/dY = (Y - T) / (Y * (1 - Y))
Private Function CalculateGradient(ByVal Y As Tensor, _
                                   ByVal T As Tensor) As Tensor
    Dim i As Long
    Dim dblY As Double
    Dim Y_() As Double
    Dim T_() As Double
    Dim dLdY_() As Double
    Dim dLdY As Tensor
    
    Set dLdY = Zeros(Y.Shape)
    Y.Flatten.CreateAlias Y_
    T.Flatten.CreateAlias T_
    dLdY.Flatten.CreateAlias dLdY_
    For i = 1 To Y.NumElements
        dblY = MinDbl(1 - EPSILON, MaxDbl(EPSILON, Y_(i)))
        dLdY_(i) = (dblY - T_(i)) / (dblY * (1 - dblY))
    Next i
    Y.Flatten.RemoveAlias Y_
    T.Flatten.RemoveAlias T_
    dLdY.Flatten.RemoveAlias dLdY_
    Set CalculateGradient = dLdY
End Function

Private Sub Save(ByVal oSerializer As Serializer)
    'Do nothing
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    'Do nothing
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Function ICriterion_Forward(ByVal Y As Object, _
                                        ByVal T As Object) As Object
        Set ICriterion_Forward = Forward(Y, T)
    End Function
    
    Private Function ICriterion_Backward(ByVal Y As Object, _
                                         ByVal T As Object) As Object
        Set ICriterion_Backward = Backward(Y, T)
    End Function
    
    Private Sub ISerializable_Save(ByVal oSerializer As Object)
        Save oSerializer
    End Sub
    
    Private Sub ISerializable_Load(ByVal oSerializer As Object)
        Load oSerializer
    End Sub
#Else
    Private Function ICriterion_Forward(ByVal Y As Tensor, _
                                        ByVal T As Tensor) As Tensor
        Set ICriterion_Forward = Forward(Y, T)
    End Function
    
    Private Function ICriterion_Backward(ByVal Y As Tensor, _
                                         ByVal T As Tensor) As Tensor
        Set ICriterion_Backward = Backward(Y, T)
    End Function
    
    Private Sub ISerializable_Save(ByVal oSerializer As Serializer)
        Save oSerializer
    End Sub
    
    Private Sub ISerializable_Load(ByVal oSerializer As Serializer)
        Load oSerializer
    End Sub
#End If

