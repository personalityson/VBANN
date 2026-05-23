VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SigmoidLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILayer
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Forward/backward cache
Private m_X As Tensor
Private m_Y As Tensor

Public Property Get Learnables() As Variant
    Learnables = Array()
End Property

Public Sub Init()
    Clear
End Sub

Public Function Forward(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "SigmoidLayer.Forward"
    
    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set m_X = X
    Set m_Y = CalculateOutput(X)
    Set Forward = m_Y
End Function

Public Function Backward(ByVal dLdY As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "SigmoidLayer.Backward"
    
    If m_Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Forward pass must be executed before the backward pass."
    End If
    If dLdY Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not dLdY.ShapeEquals(m_Y.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor does not match the expected shape for backward pass."
    End If
    Set Backward = CalculateInputGradient(m_Y, dLdY)
End Function

Private Function CalculateOutput(ByVal X As Tensor) As Tensor
    Set CalculateOutput = VecSigmoid(X)
End Function

Private Function CalculateInputGradient(ByVal Y As Tensor, _
                                     ByVal dLdY As Tensor) As Tensor
    Set CalculateInputGradient = VecMul(dLdY, VecSigmoidDerivative(Y))
End Function

Private Sub Clear()
    Set m_X = Nothing
    Set m_Y = Nothing
End Sub

Private Sub Save(ByVal oSerializer As Serializer)
    'Do nothing
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Clear
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get ILayer_Learnables() As Variant
        ILayer_Learnables = Me.Learnables
    End Property
    
    Private Function ILayer_Forward(ByVal X As Object, _
                                    Optional ByVal bIsTraining As Boolean) As Object
        Set ILayer_Forward = Me.Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal dLdY As Object) As Object
        Set ILayer_Backward = Me.Backward(dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oSerializer As Object)
        Save oSerializer
    End Sub
    
    Private Sub ISerializable_Load(ByVal oSerializer As Object)
        Load oSerializer
    End Sub
#Else
    Private Property Get ILayer_Learnables() As Variant
        ILayer_Learnables = Me.Learnables
    End Property
    
    Private Function ILayer_Forward(ByVal X As Tensor, _
                                    Optional ByVal bIsTraining As Boolean) As Tensor
        Set ILayer_Forward = Me.Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal dLdY As Tensor) As Tensor
        Set ILayer_Backward = Me.Backward(dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oSerializer As Serializer)
        Save oSerializer
    End Sub
    
    Private Sub ISerializable_Load(ByVal oSerializer As Serializer)
        Load oSerializer
    End Sub
#End If

