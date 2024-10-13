VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LeakyReLULayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_LATE_BINDING_IN_INTERFACES = True

Implements ILayer
Implements ISerializable

Private m_dblNegativeSlope As Double

'Cache
Private m_X As Tensor
Private m_Y As Tensor

Public Property Get LearnableParameters() As Variant
    LearnableParameters = Array()
End Property

Public Sub Init(Optional ByVal dblNegativeSlope As Double = 0.01)
    m_dblNegativeSlope = dblNegativeSlope
End Sub

Public Function Forward(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "LeakyReLULayer.Forward"
    
    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Set m_X = X
    Set m_Y = CalculateOutput(X)
    Set Forward = m_Y
End Function

Public Function Backward(ByVal dLdY As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "LeakyReLULayer.Backward"
    
    If m_X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Forward pass must be executed before the backward pass."
    End If
    If dLdY Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not dLdY.ShapeEquals(m_X.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor shape does not match the expected shape for backward pass."
    End If
    Set Backward = CalculateInputError(m_X, dLdY)
End Function

Private Function CalculateOutput(ByVal X As Tensor) As Tensor
    Dim i As Long
    Dim Y_() As Double
    Dim Y As Tensor
    
    Set Y = X.Clone
    Y.Flatten.CreateAlias Y_
    For i = 1 To X.NumElements
        If Y_(i) < 0 Then
            Y_(i) = Y_(i) * m_dblNegativeSlope
        End If
    Next i
    Y.Flatten.RemoveAlias Y_
    Set CalculateOutput = Y
End Function

Private Function CalculateInputError(ByVal X As Tensor, _
                                     ByVal dLdY As Tensor) As Tensor
    Dim i As Long
    Dim X_() As Double
    Dim dLdX_() As Double
    Dim dLdX As Tensor

    Set dLdX = dLdY.Clone
    X.Flatten.CreateAlias X_
    dLdX.Flatten.CreateAlias dLdX_
    For i = 1 To X.NumElements
        If X_(i) < 0 Then
            dLdX_(i) = dLdX_(i) * m_dblNegativeSlope
        End If
    Next i
    X.Flatten.RemoveAlias X_
    dLdX.Flatten.RemoveAlias dLdX_
    Set CalculateInputError = dLdX
End Function

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "LeakyReLULayer.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteDouble m_dblNegativeSlope
        .WriteObject m_X
        .WriteObject m_Y
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "LeakyReLULayer.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        m_dblNegativeSlope = .ReadDouble()
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
