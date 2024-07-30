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

Public Property Get Parameters() As Variant
    Parameters = Array()
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
    Set Forward = CalculateOutput(X)
End Function

Public Function Backward(ByVal X As Tensor, _
                         ByVal Y As Tensor, _
                         ByVal dLdY As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "SigmoidLayer.Backward"
    
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
    'Do nothing
End Sub

Public Sub Load(ByVal oWorksheet As Serializer)
    'Do nothing
End Sub

Private Function SafeSigmoid(ByVal dblValue As Double) As Double
    If dblValue >= -DOUBLE_MAX_LOG Then
        SafeSigmoid = 1 / (1 + Exp(-dblValue))
    End If
End Function

Private Function CalculateOutput(ByVal X As Tensor) As Tensor
    Dim i As Long
    Dim X_() As Double
    Dim Y_() As Double
    Dim Y As Tensor
    
    Set Y = Zeros(X.Shape)
    X.Ravel.CreateAlias X_
    Y.Ravel.CreateAlias Y_
    For i = 1 To X.NumElements
        Y_(i) = SafeSigmoid(X_(i))
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
        dLdX_(i) = dLdY_(i) * Y_(i) * (1 - Y_(i))
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

