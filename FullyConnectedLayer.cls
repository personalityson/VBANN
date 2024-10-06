VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FullyConnectedLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_LATE_BINDING_IN_INTERFACES = True

Implements ILayer
Implements ISerializable

Private m_lInputSize As Long
Private m_lOutputSize As Long
Private m_W As Parameter
Private m_b As Parameter

'Cache
Private m_X As Tensor
Private m_Y As Tensor

Public Property Get Weights() As Parameter
    Set Weights = m_W
End Property

Public Property Get Bias() As Parameter
    Set Bias = m_b
End Property

Public Property Get LearnableParameters() As Variant
    LearnableParameters = Array(m_W, m_b)
End Property

Public Sub Init(ByVal lInputSize As Long, _
                ByVal lOutputSize As Long)
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Init"
    
    If lInputSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Input size must be greater than 0."
    End If
    If lOutputSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Output size must be greater than 0."
    End If
    m_lInputSize = lInputSize
    m_lOutputSize = lOutputSize
    Set m_W = Parameter(GlorotUniform(Array(m_lOutputSize, m_lInputSize), m_lInputSize, m_lOutputSize, 1), 1, 1)
    Set m_b = Parameter(Zeros(Array(m_lOutputSize)), 1, 0)
End Sub

Public Function Forward(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Forward"
    
    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not X.ShapeEquals(Array(m_lInputSize, -1)) Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor shape does not match the expected shape for forward pass."
    End If
    Set m_X = X
    Set m_Y = CalculateOutput(X)
    Set Forward = m_Y
End Function

Public Function Backward(ByVal dLdY As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Backward"
    
    If m_X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Forward pass must be executed before the backward pass."
    End If
    If dLdY Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not dLdY.ShapeEquals(Array(m_lOutputSize, m_X.Size(2))) Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor shape does not match the expected shape for backward pass."
    End If
    UpdateGradients m_X, dLdY
    Set Backward = CalculateInputError(dLdY)
End Function

Private Function CalculateOutput(ByVal X As Tensor) As Tensor
    Dim lBatchSize As Long
    Dim Y As Tensor

    lBatchSize = X.Size(2)
    Set Y = Zeros(Array(m_lOutputSize, lBatchSize))
    'Update in-place, Y = Y + W * X
    MatMul_I Y, m_W.Variable, X
    'Update in-place, Y = Y + b
    MatMul_I Y, m_b.Variable, Ones(lBatchSize), True, True
    Set CalculateOutput = Y
End Function

Private Function CalculateInputError(ByVal dLdY As Tensor) As Tensor
    'dLdX = W' * dLdY
    Set CalculateInputError = MatMul(m_W.Variable, dLdY, True)
End Function

Private Sub UpdateGradients(ByVal X As Tensor, _
                            ByVal dLdY As Tensor)
    Dim lBatchSize As Long
    
    lBatchSize = X.Size(2)
    'Update in-place, dLdW = dLdW + dLdY * X'
    MatMul_I m_W.Gradient, dLdY, X, , True
    'Update in-place, dLdb = dLdb + dLdY
    MatMul_I m_b.Gradient, dLdY, Ones(lBatchSize)
    m_W.NumUpdates = m_W.NumUpdates + lBatchSize
    m_b.NumUpdates = m_b.NumUpdates + lBatchSize
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteLong m_lInputSize
        .WriteLong m_lOutputSize
        .WriteObject m_W
        .WriteObject m_b
        .WriteObject m_X
        .WriteObject m_Y
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        m_lInputSize = .ReadLong()
        m_lOutputSize = .ReadLong()
        Set m_W = .ReadObject()
        Set m_b = .ReadObject()
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
