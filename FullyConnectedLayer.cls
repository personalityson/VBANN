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

Public Property Get InputSize() As Long
    InputSize = m_lInputSize
End Property

Public Property Get OutputSize() As Long
    OutputSize = m_lOutputSize
End Property

Public Property Get Weights() As Parameter
    Set Weights = m_W
End Property

Public Property Get Bias() As Parameter
    Set Bias = m_b
End Property

Public Property Get Parameters() As Variant
    Parameters = Array(m_W, m_b)
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
        Err.Raise 5, PROCEDURE_NAME, "Input tensor shape does not match the expected shape."
    End If
    If IsBlasAvailable() Then
        Set Forward = CalculateOutputWithBlas(X)
    Else
        Set Forward = CalculateOutputNaively(X)
    End If
End Function

Public Function Backward(ByVal X As Tensor, _
                         ByVal Y As Tensor, _
                         ByVal dLdY As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Backward"
    
    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Y Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If dLdY Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not X.ShapeEquals(Array(m_lInputSize, -1)) Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor shape does not match the expected shape."
    End If
    If Not Y.ShapeEquals(Array(m_lOutputSize, X.Size(2))) Then
        Err.Raise 5, PROCEDURE_NAME, "Output tensor shape does not match the expected shape."
    End If
    If Not dLdY.ShapeEquals(Array(m_lOutputSize, X.Size(2))) Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor shape does not match the expected shape."
    End If
    If IsBlasAvailable() Then
        UpdateGradientsWithBlas X, dLdY
        Set Backward = CalculateInputErrorWithBlas(dLdY)
    Else
        UpdateGradientsNaively X, dLdY
        Set Backward = CalculateInputErrorNaively(dLdY)
    End If
End Function

Public Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteLong m_lInputSize
        .WriteLong m_lOutputSize
        .WriteObject m_W
        .WriteObject m_b
    End With
End Sub

Public Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "FullyConnectedLayer.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        m_lInputSize = .ReadLong()
        m_lOutputSize = .ReadLong()
        Set m_W = .ReadObject()
        Set m_b = .ReadObject()
    End With
End Sub

Private Function CalculateOutputNaively(ByVal X As Tensor) As Tensor
    Dim lBatchSize As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dblSum As Double
    Dim W_() As Double
    Dim b_() As Double
    Dim X_() As Double
    Dim Y_() As Double
    Dim Y As Tensor

    lBatchSize = X.Size(2)
    Set Y = Zeros(Array(m_lOutputSize, lBatchSize))
    X.CreateAlias X_
    Y.CreateAlias Y_
    m_W.Learnable.CreateAlias W_
    m_b.Learnable.CreateAlias b_
    For k = 1 To lBatchSize
        For i = 1 To m_lOutputSize
            dblSum = b_(i)
            For j = 1 To m_lInputSize
                dblSum = dblSum + W_(i, j) * X_(j, k)
            Next j
            Y_(i, k) = dblSum
        Next i
    Next k
    X.RemoveAlias X_
    Y.RemoveAlias Y_
    m_W.Learnable.RemoveAlias W_
    m_b.Learnable.RemoveAlias b_
    Set CalculateOutputNaively = Y
End Function

Private Function CalculateInputErrorNaively(ByVal dLdY As Tensor) As Tensor
    Dim lBatchSize As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dblSum As Double
    Dim dLdX_() As Double
    Dim dLdY_() As Double
    Dim W_() As Double
    Dim dLdX As Tensor
    
    lBatchSize = dLdY.Size(2)
    Set dLdX = Zeros(Array(m_lInputSize, lBatchSize))
    dLdX.CreateAlias dLdX_
    dLdY.CreateAlias dLdY_
    m_W.Learnable.CreateAlias W_
    For k = 1 To lBatchSize
        For j = 1 To m_lInputSize
            dblSum = 0
            For i = 1 To m_lOutputSize
                dblSum = dblSum + W_(i, j) * dLdY_(i, k)
            Next i
            dLdX_(j, k) = dblSum
        Next j
    Next k
    dLdX.RemoveAlias dLdX_
    dLdY.RemoveAlias dLdY_
    m_W.Learnable.RemoveAlias W_
    Set CalculateInputErrorNaively = dLdX
End Function

Private Sub UpdateGradientsNaively(ByVal X As Tensor, _
                                   ByVal dLdY As Tensor)
    Dim lBatchSize As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dblSum As Double
    Dim X_() As Double
    Dim dLdY_() As Double
    Dim dLdW_() As Double
    Dim dLdb_() As Double

    lBatchSize = X.Size(2)
    X.CreateAlias X_
    dLdY.CreateAlias dLdY_
    m_W.Gradient.CreateAlias dLdW_
    m_b.Gradient.CreateAlias dLdb_
    For i = 1 To m_lOutputSize
        For j = 1 To m_lInputSize
            dblSum = 0
            For k = 1 To lBatchSize
                dblSum = dblSum + X_(j, k) * dLdY_(i, k)
            Next k
            dLdW_(i, j) = dblSum
        Next j
        dblSum = 0
        For k = 1 To lBatchSize
            dblSum = dblSum + dLdY_(i, k)
        Next k
        dLdb_(i) = dblSum
    Next i
    X.RemoveAlias X_
    dLdY.RemoveAlias dLdY_
    m_W.Gradient.RemoveAlias dLdW_
    m_b.Gradient.RemoveAlias dLdb_
    m_W.NumUpdates = m_W.NumUpdates + lBatchSize
    m_b.NumUpdates = m_b.NumUpdates + lBatchSize
End Sub

Private Function CalculateOutputWithBlas(ByVal X As Tensor) As Tensor
    Dim lBatchSize As Long
    Dim Y As Tensor
    
    lBatchSize = X.Size(2)
    Set Y = Zeros(Array(m_lOutputSize, lBatchSize))
    dgemm "N", "N", m_lOutputSize, lBatchSize, m_lInputSize, 1#, m_W.Learnable.Address, m_lOutputSize, X.Address, m_lInputSize, 1#, Y.Address, m_lOutputSize
    dgemm "N", "N", m_lOutputSize, lBatchSize, 1&, 1#, m_b.Learnable.Address, m_lOutputSize, Ones(lBatchSize).Address, 1&, 1#, Y.Address, m_lOutputSize
    Set CalculateOutputWithBlas = Y
End Function

Private Function CalculateInputErrorWithBlas(ByVal dLdY As Tensor) As Tensor
    Dim lBatchSize As Long
    Dim dLdX As Tensor
    
    lBatchSize = dLdY.Size(2)
    Set dLdX = Zeros(Array(m_lInputSize, lBatchSize))
    dgemm "T", "N", m_lInputSize, lBatchSize, m_lOutputSize, 1#, m_W.Learnable.Address, m_lOutputSize, dLdY.Address, m_lOutputSize, 1#, dLdX.Address, m_lInputSize
    Set CalculateInputErrorWithBlas = dLdX
End Function

Private Sub UpdateGradientsWithBlas(ByVal X As Tensor, _
                                    ByVal dLdY As Tensor)
    Dim lBatchSize As Long
    
    lBatchSize = X.Size(2)
    dgemm "N", "T", m_lOutputSize, m_lInputSize, lBatchSize, 1#, dLdY.Address, m_lOutputSize, X.Address, m_lInputSize, 1#, m_W.Gradient.Address, m_lOutputSize
    dgemv "N", m_lOutputSize, lBatchSize, 1#, dLdY.Address, m_lOutputSize, Ones(lBatchSize).Address, 1&, 1#, m_b.Gradient.Address, 1&
    m_W.NumUpdates = m_W.NumUpdates + lBatchSize
    m_b.NumUpdates = m_b.NumUpdates + lBatchSize
End Sub

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
