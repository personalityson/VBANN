VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DropoutLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILayer
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Hyperparameters
Private m_dblDropoutRate As Double

'Forward/backward cache
Private m_oMask As Tensor
Private m_X As Tensor
Private m_Y As Tensor

Public Property Get LearnableParameters() As Variant
    LearnableParameters = Array()
End Property

Public Sub Init(Optional ByVal dblDropoutRate As Double = 0.5)
    Const PROCEDURE_NAME As String = "DropoutLayer.Init"
    
    If dblDropoutRate < 0 Or dblDropoutRate >= 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Dropout rate must be >= 0 and < 1."
    End If
    Clear
    m_dblDropoutRate = dblDropoutRate
End Sub

Public Function Forward(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean = True) As Tensor
    Const PROCEDURE_NAME As String = "DropoutLayer.Forward"
    
    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set m_X = X
    If bIsTraining Then
        Set m_oMask = Bernoulli(X.Shape, 1 - m_dblDropoutRate)
        Set m_Y = ApplyDropoutMask(X, m_oMask)
    Else
        Set m_Y = X
    End If
    Set Forward = m_Y
End Function

Public Function Backward(ByVal dLdY As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "DropoutLayer.Backward"
    
    If m_oMask Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Forward pass must be executed in training before the backward pass."
    End If
    If dLdY Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not dLdY.ShapeEquals(m_oMask.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor shape does not match the shape of the input from the forward pass."
    End If
    Set Backward = ApplyDropoutMask(dLdY, m_oMask)
End Function

Private Function ApplyDropoutMask(ByVal X As Tensor, _
                                  ByVal oMask As Tensor) As Tensor
    Set ApplyDropoutMask = VecMulC(VecMul(X, oMask), 1 / (1 - m_dblDropoutRate))
End Function

Private Sub Clear()
    m_dblDropoutRate = 0
    Set m_oMask = Nothing
    Set m_X = Nothing
    Set m_Y = Nothing
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "DropoutLayer.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteDouble m_dblDropoutRate
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "DropoutLayer.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oWorksheet
        m_dblDropoutRate = .ReadDouble()
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get ILayer_LearnableParameters() As Variant
        ILayer_LearnableParameters = Me.LearnableParameters
    End Property
    
    Private Function ILayer_Forward(ByVal X As Object, _
                                    Optional ByVal bIsTraining As Boolean) As Object
        Set ILayer_Forward = Me.Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal dLdY As Object) As Object
        Set ILayer_Backward = Me.Backward(dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Object)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Object)
        Load oWorksheet
    End Sub
#Else
    Private Property Get ILayer_LearnableParameters() As Variant
        ILayer_LearnableParameters = Me.LearnableParameters
    End Property
    
    Private Function ILayer_Forward(ByVal X As Tensor, _
                                    Optional ByVal bIsTraining As Boolean) As Tensor
        Set ILayer_Forward = Me.Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal dLdY As Tensor) As Tensor
        Set ILayer_Backward = Me.Backward(dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If

