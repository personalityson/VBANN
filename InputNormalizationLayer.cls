VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputNormalizationLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILayer
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

Private Const EPSILON As Double = 0.00000001

'Dimensions
Private m_alInputShape() As Long

'Normalization stats
Private m_oMean As Tensor
Private m_oVariance As Tensor

'Forward/backward cache
Private m_X As Tensor
Private m_Y As Tensor

Public Property Get LearnableParameters() As Variant
    LearnableParameters = Array()
End Property

Public Sub Init(ByVal oTrainingLoader As DataLoader)
    Const PROCEDURE_NAME As String = "InputNormalizationLayer.Init"

    If oTrainingLoader Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid DataLoader object is required."
    End If
    Clear
    Adapt oTrainingLoader
End Sub

Public Function Forward(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "InputNormalizationLayer.Forward"

    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not X.ShapeEquals(m_alInputShape) Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor does not match the expected shape for forward pass."
    End If
    Set m_X = X
    Set m_Y = CalculateOutput(X)
    Set Forward = m_Y
End Function

Public Function Backward(ByVal dLdY As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "InputNormalizationLayer.Backward"
    
    If dLdY Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not dLdY.ShapeEquals(m_alInputShape) Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor does not match the expected shape for backward pass."
    End If
    Set Backward = CalculateInputError(dLdY)
End Function

'Y = (X - m) / Sqrt(v + Epsilon)
Private Function CalculateOutput(ByVal X As Tensor) As Tensor
    Dim lBatchDimension As Long
    Dim lBatchSize As Long
    Dim oTiledMean As Tensor
    Dim oTiledVariance As Tensor

    lBatchDimension = X.NumDimensions
    lBatchSize = X.Size(lBatchDimension)
    Set oTiledMean = m_oMean.Tile(lBatchDimension, lBatchSize)
    Set oTiledVariance = m_oVariance.Tile(lBatchDimension, lBatchSize)
    Set CalculateOutput = VecDivSqrtAddC(VecSub(X, oTiledMean), oTiledVariance, EPSILON)
End Function

'dLdX = dLdY / Sqrt(v + Epsilon)
Private Function CalculateInputError(ByVal dLdY As Tensor) As Tensor
    Dim lBatchDimension As Long
    Dim lBatchSize As Long
    Dim oTiledVariance As Tensor

    lBatchDimension = dLdY.NumDimensions
    lBatchSize = dLdY.Size(lBatchDimension)
    Set oTiledVariance = m_oVariance.Tile(lBatchDimension, lBatchSize)
    Set CalculateInputError = VecDivSqrtAddC(dLdY, oTiledVariance, EPSILON)
End Function

Private Sub Adapt(ByVal oTrainingLoader As DataLoader)
    Dim lBatchDimension As Long
    Dim lBatchSize As Long
    Dim lNumSamples As Long
    Dim X As Tensor
    Dim oMean As Tensor
    Dim oM2 As Tensor
    Dim oDelta As Tensor
    Dim oBatchMean As Tensor
    Dim oBatchM2 As Tensor
    Dim oBatch As TensorDataset

    With oTrainingLoader
        .Reset
        Do While .Fetch(oBatch)
            Set X = oBatch.Item(1)
            lBatchDimension = X.NumDimensions
            lBatchSize = oBatch.NumSamples
            Set oBatchMean = X.Reduce(lBatchDimension, rdcMean)
            Set oBatchM2 = VecPow2(VecSub(X, oBatchMean.Tile(lBatchDimension, lBatchSize))).Reduce(lBatchDimension, rdcSum)
            If lNumSamples = 0 Then
                m_alInputShape = X.Shape
                m_alInputShape(X.NumDimensions) = -1
                Set oMean = oBatchMean
                Set oM2 = oBatchM2
                lNumSamples = lBatchSize
            Else
                Set oDelta = VecSub(oBatchMean, oMean)
                VecLinComb_I 1, oMean, lBatchSize / (lNumSamples + lBatchSize), oDelta
                VecLinComb_I 1, oM2, 1, oBatchM2
                VecLinComb_I 1, oM2, lNumSamples * lBatchSize / (lNumSamples + lBatchSize), VecPow2(oDelta)
                lNumSamples = lNumSamples + lBatchSize
            End If
        Loop
    End With
    If lNumSamples > 0 Then
        Set m_oMean = oMean
        Set m_oVariance = VecDivC(oM2, lNumSamples)
    End If
End Sub

Private Sub Clear()
    Erase m_alInputShape
    Set m_oMean = Nothing
    Set m_oVariance = Nothing
    Set m_X = Nothing
    Set m_Y = Nothing
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "InputNormalizationLayer.Save"

    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteLongArray m_alInputShape
        .WriteObject m_oMean
        .WriteObject m_oVariance
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "InputNormalizationLayer.Load"

    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oWorksheet
        m_alInputShape = .ReadLongArray()
        Set m_oMean = .ReadObject()
        Set m_oVariance = .ReadObject()
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

