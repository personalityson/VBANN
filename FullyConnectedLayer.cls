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

Implements ILayer
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Dimensions
Private m_lInputSize As Long
Private m_lOutputSize As Long

'Learnable parameters
Private m_oWeights As Parameter
Private m_oBias As Parameter

'Forward/backward cache
Private m_X As Tensor
Private m_Y As Tensor

Public Property Get Weights() As Parameter
    Set Weights = m_oWeights
End Property

Public Property Get Bias() As Parameter
    Set Bias = m_oBias
End Property

Public Property Get LearnableParameters() As Variant
    LearnableParameters = Array(m_oWeights, m_oBias)
End Property

Public Sub Init(ByVal lInputSize As Long, _
                ByVal lOutputSize As Long)
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Init"
    
    If lInputSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Input size must be >= 1."
    End If
    If lOutputSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Output size must be >= 1."
    End If
    Clear
    m_lInputSize = lInputSize
    m_lOutputSize = lOutputSize
    Set m_oWeights = Parameter(GlorotUniform(Array(m_lOutputSize, m_lInputSize), m_lInputSize, m_lOutputSize, 1), 1, 1)
    Set m_oBias = Parameter(Zeros(Array(m_lOutputSize)), 1, 0)
End Sub

Public Function Forward(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Forward"

    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If X.NumDimensions <> 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor must have two dimensions."
    End If
    If X.Size(1) <> m_lInputSize Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor's first dimension does not match the layer's input size."
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
    If dLdY.NumDimensions <> 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor must have two dimensions."
    End If
    If dLdY.Size(1) <> m_lOutputSize Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor's first dimension must match the layer's output size."
    End If
    If dLdY.Size(2) <> m_X.Size(2) Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor's second dimension must match the batch size from the forward pass."
    End If
    UpdateGradients m_X, dLdY
    Set Backward = CalculateInputError(dLdY)
End Function

'Y = W * X + b
Private Function CalculateOutput(ByVal X As Tensor) As Tensor
    Dim lBatchSize As Long
    Dim Y As Tensor

    lBatchSize = X.Size(2)
    Set Y = Zeros(Array(m_lOutputSize, lBatchSize))
    MatMul_I Y, m_oWeights.Variable, X
    MatMul_I Y, m_oBias.Variable, Ones(lBatchSize), True, True
    Set CalculateOutput = Y
End Function

'dLdX = W' * dLdY
Private Function CalculateInputError(ByVal dLdY As Tensor) As Tensor
    Set CalculateInputError = MatMul(m_oWeights.Variable, dLdY, True)
End Function

'dLdW = dLdW + dLdY * X'
'dLdb = dLdb + dLdY
Private Sub UpdateGradients(ByVal X As Tensor, _
                            ByVal dLdY As Tensor)
    Dim lBatchSize As Long

    lBatchSize = X.Size(2)
    MatMul_I m_oWeights.Gradient, dLdY, X, , True
    MatMul_I m_oBias.Gradient, dLdY, Ones(lBatchSize)
    m_oWeights.NumUpdates = m_oWeights.NumUpdates + lBatchSize
    m_oBias.NumUpdates = m_oBias.NumUpdates + lBatchSize
End Sub

Private Sub Clear()
    m_lInputSize = 0
    m_lOutputSize = 0
    Set m_oWeights = Nothing
    Set m_oBias = Nothing
    Set m_X = Nothing
    Set m_Y = Nothing
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteLong m_lInputSize
        .WriteLong m_lOutputSize
        .WriteObject m_oWeights
        .WriteObject m_oBias
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oWorksheet
        m_lInputSize = .ReadLong()
        m_lOutputSize = .ReadLong()
        Set m_oWeights = .ReadObject()
        Set m_oBias = .ReadObject()
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

