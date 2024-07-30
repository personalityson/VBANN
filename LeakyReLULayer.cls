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

Public Property Get NegativeSlope() As Double
    NegativeSlope = m_dblNegativeSlope
End Property

Public Property Get Parameters() As Variant
    Parameters = Array()
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
    Set Forward = CalculateOutput(X)
End Function

Public Function Backward(ByVal X As Tensor, _
                         ByVal Y As Tensor, _
                         ByVal dLdY As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "LeakyReLULayer.Backward"
    
    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If dLdY Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not Y.ShapeEquals(dLdY.Shape) Then
        Err.Raise 5, PROCEDURE_NAME, "Shapes of Y and dLdY do not match."
    End If
    Set Backward = CalculateInputError(Y, dLdY)
End Function

Public Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "LeakyReLULayer.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteDouble m_dblNegativeSlope
    End With
End Sub

Public Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "LeakyReLULayer.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        m_dblNegativeSlope = .ReadDouble()
    End With
End Sub

Private Function CalculateOutput(ByVal X As Tensor) As Tensor
    Dim i As Long
    Dim X_() As Double
    Dim Y_() As Double
    Dim Y As Tensor
    
    Set Y = Zeros(X.Shape)
    X.Ravel.CreateAlias X_
    Y.Ravel.CreateAlias Y_
    For i = 1 To X.NumElements
        If X_(i) < 0 Then
            Y_(i) = X_(i) * m_dblNegativeSlope
        Else
            Y_(i) = X_(i)
        End If
    Next i
    X.Ravel.RemoveAlias X_
    Y.Ravel.RemoveAlias Y_
    Set CalculateOutput = Y
End Function

Private Function CalculateInputError(ByVal Y As Tensor, _
                                     ByVal dLdY As Tensor) As Tensor
    Dim i As Long
    Dim Y_() As Double
    Dim dLdX_() As Double
    Dim dLdY_() As Double
    Dim dLdX As Tensor
    
    Set dLdX = Zeros(Y.Shape)
    Y.Ravel.CreateAlias Y_
    dLdX.Ravel.CreateAlias dLdX_
    dLdY.Ravel.CreateAlias dLdY_
    For i = 1 To Y.NumElements
        If Y_(i) < 0 Then
            dLdX_(i) = dLdY_(i) * m_dblNegativeSlope
        Else
            dLdX_(i) = dLdY_(i)
        End If
    Next i
    Y.Ravel.RemoveAlias Y_
    dLdX.Ravel.RemoveAlias dLdX_
    dLdY.Ravel.RemoveAlias dLdY_
    Set CalculateInputError = dLdX
End Function

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get ILayer_Parameters() As Variant
        ILayer_Parameters = Parameters
    End Property
    
    Private Function ILayer_Forward(ByVal X As Object, _
                                    Optional ByVal bIsTraining As Boolean) As Object
        Set ILayer_Forward = Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal X As Object, _
                                     ByVal Y As Object, _
                                     ByVal dLdY As Object) As Object
        Set ILayer_Backward = Backward(X, Y, dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Object)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Object)
        Load oWorksheet
    End Sub
#Else
    Private Property Get ILayer_Parameters() As Variant
        ILayer_Parameters = Parameters
    End Property
    
    Private Function ILayer_Forward(ByVal X As Tensor, _
                                    Optional ByVal bIsTraining As Boolean) As Tensor
        Set ILayer_Forward = Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal X As Tensor, _
                                     ByVal Y As Tensor, _
                                     ByVal dLdY As Tensor) As Tensor
        Set ILayer_Backward = Backward(X, Y, dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If
