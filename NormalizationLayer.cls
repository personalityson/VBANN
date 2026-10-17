VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NormalizationLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILayer
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

Private Const ADAPT_BATCH_SIZE As Long = 1024
Private Const NORMALIZATION_EPSILON As Double = 0.00001

'Configuration
Private m_bInvert As Boolean

'Dimensions
Private m_alInputShape() As Long

'Normalization stats with last batch dimension = 1
Private m_oShift As Tensor
Private m_oScale As Tensor

'Forward/backward cache
Private m_X As Tensor
Private m_Y As Tensor

Public Property Get Learnables() As Variant
    Learnables = Array()
End Property

Public Sub Init(ByVal oTrainingSet As IDataset, _
                ByVal lIndex As Long, _
                Optional ByVal bInvert As Boolean)
    Const PROCEDURE_NAME As String = "NormalizationLayer.Init"
    Dim oTrainingLoader As DataLoader
    
    If oTrainingSet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IDataset object is required."
    End If
    If oTrainingSet.NumTensors < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Dataset must have at least one tensor."
    End If
    Clear
    m_bInvert = bInvert
    Set oTrainingLoader = New DataLoader
    oTrainingLoader.Init oTrainingSet, ADAPT_BATCH_SIZE, bDropRemainder:=False
    Adapt oTrainingLoader, lIndex
End Sub

Public Function Forward(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "NormalizationLayer.Forward"

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
    Const PROCEDURE_NAME As String = "NormalizationLayer.Backward"

    If dLdY Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not dLdY.ShapeEquals(m_alInputShape) Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor does not match the expected shape for backward pass."
    End If
    Set Backward = CalculateInputGradient(dLdY)
End Function

Private Sub Clear()
    m_bInvert = False
    Erase m_alInputShape
    Set m_oShift = Nothing
    Set m_oScale = Nothing
    Set m_X = Nothing
    Set m_Y = Nothing
End Sub

Private Function CalculateOutput(ByVal X As Tensor) As Tensor
    If m_bInvert Then
        Set CalculateOutput = VecAdd(VecMul(X, m_oScale), m_oShift)
    Else
        Set CalculateOutput = VecDiv(VecSub(X, m_oShift), m_oScale)
    End If
End Function

Private Function CalculateInputGradient(ByVal dLdY As Tensor) As Tensor
    If m_bInvert Then
        Set CalculateInputGradient = VecMul(dLdY, m_oScale)
    Else
        Set CalculateInputGradient = VecDiv(dLdY, m_oScale)
    End If
End Function

Private Sub Adapt(ByVal oTrainingLoader As DataLoader, _
                  ByVal lIndex As Long)
    Dim lBatchDimension As Long
    Dim lBatchSize As Long
    Dim lNumSamples As Long
    Dim alInputShape() As Long
    Dim X As Tensor
    Dim oMean As Tensor
    Dim oM2 As Tensor
    Dim oDelta As Tensor
    Dim oBatchMean As Tensor
    Dim oBatchM2 As Tensor
    Dim oBatch As TensorDataset

    With oTrainingLoader
        .Reset
        Do While .NextBatch(oBatch)
            Set X = oBatch.Tensors(lIndex)
            lBatchDimension = X.NumDimensions
            lBatchSize = oBatch.NumSamples
            Set oBatchMean = X.Reduce(lBatchDimension, rtMean)
            Set oBatchM2 = VecPow2(VecSub(X, oBatchMean)).Reduce(lBatchDimension, rtSum)
            If lNumSamples = 0 Then
                alInputShape = X.Shape
                alInputShape(lBatchDimension) = spWildcard
                Set oMean = oBatchMean
                Set oM2 = oBatchM2
            Else
                Set oDelta = VecSub(oBatchMean, oMean)
                VecLinComb_I 1, oMean, lBatchSize / (lNumSamples + lBatchSize), oDelta
                VecLinComb_I 1, oM2, 1, oBatchM2
                VecLinComb_I 1, oM2, CDbl(lNumSamples) * lBatchSize / (lNumSamples + lBatchSize), VecPow2(oDelta)
            End If
            lNumSamples = lNumSamples + lBatchSize
            Application.StatusBar = "Adapt progress: " & Format(lNumSamples / .NumSamples, "0.0%")
            DoEvents
        Loop
    End With
    If lNumSamples > 0 Then
        m_alInputShape = alInputShape
        Set m_oShift = oMean
        Set m_oScale = VecSqrt(VecAddC(VecDivC(oM2, lNumSamples), NORMALIZATION_EPSILON))
    End If
End Sub

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "NormalizationLayer.Save"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteBoolean m_bInvert
        .WriteLongArray m_alInputShape
        .WriteObject m_oShift
        .WriteObject m_oScale
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "NormalizationLayer.Load"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oSerializer
        m_bInvert = .ReadBoolean()
        m_alInputShape = .ReadLongArray()
        Set m_oShift = .ReadObject()
        Set m_oScale = .ReadObject()
    End With
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

