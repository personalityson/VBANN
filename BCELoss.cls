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

#Const USE_LATE_BINDING_IN_INTERFACES = True

Implements ICriterion
Implements ISerializable

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
        Err.Raise 5, PROCEDURE_NAME, "Shapes of Y and T do not match."
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
        Err.Raise 5, PROCEDURE_NAME, "Shapes of Y and T do not match."
    End If
    Set Backward = CalculateOutputError(Y, T)
End Function

Private Function CalculateLoss(ByVal Y As Tensor, _
                               ByVal T As Tensor) As Tensor
    Dim i As Long
    Dim Y_() As Double
    Dim T_() As Double
    Dim L_() As Double
    Dim l As Tensor
    
    'L = -T * Log(Y) - (1 - T) * Log(1 - Y)
    Set l = Zeros(Y.Shape)
    Y.Flatten.CreateAlias Y_
    T.Flatten.CreateAlias T_
    l.Flatten.CreateAlias L_
    For i = 1 To Y.NumElements
        L_(i) = -T_(i) * Log(Y_(i) + DOUBLE_MIN_ABS) - (1 - T_(i)) * Log(1 - Y_(i) + DOUBLE_MIN_ABS)
    Next i
    Y.Flatten.RemoveAlias Y_
    T.Flatten.RemoveAlias T_
    l.Flatten.RemoveAlias L_
    Set CalculateLoss = l
End Function

Private Function CalculateOutputError(ByVal Y As Tensor, _
                                      ByVal T As Tensor) As Tensor
    'dLdY = Y - T
    Set CalculateOutputError = VecSub(Y, T)
End Function

Private Sub Save(ByVal oWorksheet As Serializer)
    'Do nothing
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
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
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Object)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Object)
        Load oWorksheet
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
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If

