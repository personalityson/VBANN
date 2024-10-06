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

#Const USE_LATE_BINDING_IN_INTERFACES = True

Implements ILayer
Implements ISerializable

'Cache
Private m_X As Tensor
Private m_Y As Tensor

Public Property Get LearnableParameters() As Variant
    LearnableParameters = Array()
End Property

Public Sub Init()
    'Do nothing
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
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor shape does not match the expected shape for backward pass."
    End If
    Set Backward = CalculateInputError(m_Y, dLdY)
End Function

Private Function CalculateOutput(ByVal X As Tensor) As Tensor
    'Y = Sigmoid(X)
    Set CalculateOutput = VecSigmoid(X)
End Function

Private Function CalculateInputError(ByVal Y As Tensor, _
                                     ByVal dLdY As Tensor) As Tensor
    'dLdX = dLdY * Y * (1 - Y)
    Set CalculateInputError = VecMul(dLdY, VecMul(Y, VecSubCRev(Y, 1)))
End Function

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "SigmoidLayer.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteObject m_X
        .WriteObject m_Y
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "SigmoidLayer.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        Set m_X = .ReadObject()
        Set m_Y = .ReadObject()
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get ILayer_LearnableParameters() As Variant
        ILayer_LearnableParameters = LearnableParameters
    End Property
    
    Private Function ILayer_Forward(ByVal X As Object, _
                                    Optional ByVal bIsTraining As Boolean) As Object
        Set ILayer_Forward = Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal dLdY As Object) As Object
        Set ILayer_Backward = Backward(dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Object)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Object)
        Load oWorksheet
    End Sub
#Else
    Private Property Get ILayer_LearnableParameters() As Variant
        ILayer_LearnableParameters = LearnableParameters
    End Property
    
    Private Function ILayer_Forward(ByVal X As Tensor, _
                                    Optional ByVal bIsTraining As Boolean) As Tensor
        Set ILayer_Forward = Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal dLdY As Tensor) As Tensor
        Set ILayer_Backward = Backward(dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If
