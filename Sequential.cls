VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sequential"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_LATE_BINDING_IN_INTERFACES = True

Implements ISerializable

Private m_oCriterion As ICriterion
Private m_oOptimizer As IOptimizer
Private m_lNumLayers As Long
Private m_aoLayers() As ILayer

Public Property Get Criterion() As ICriterion
    Set Criterion = m_oCriterion
End Property

Public Property Get Optimizer() As IOptimizer
    Set Optimizer = m_oOptimizer
End Property

Public Function Init(ByVal oCriterion As ICriterion, _
                     ByVal oOptimizer As IOptimizer)
    Const PROCEDURE_NAME As String = "Sequential.Init"

    If oCriterion Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ICriterion object is required."
    End If
    If oOptimizer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IOptimizer object is required."
    End If
    Set m_oCriterion = oCriterion
    Set m_oOptimizer = oOptimizer
    m_lNumLayers = 0
    Erase m_aoLayers
End Function

Public Sub Add(ByVal oLayer As ILayer)
    Const PROCEDURE_NAME As String = "Sequential.Add"

    If oLayer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ILayer object is required."
    End If
    m_lNumLayers = m_lNumLayers + 1
    ReDim Preserve m_aoLayers(1 To m_lNumLayers)
    Set m_aoLayers(m_lNumLayers) = oLayer
End Sub

Public Function Predict(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean) As Tensor
    Dim i As Long

    For i = 1 To m_lNumLayers
        Set X = m_aoLayers(i).Forward(X, bIsTraining)
    Next i
    Set Predict = X
End Function

Public Sub Backpropagate(ByVal X As Tensor, _
                         ByVal T As Tensor)
    Dim i As Long
    Dim Y As Tensor
    Dim dLdY As Tensor

    Set Y = Predict(X, True)
    Set dLdY = m_oCriterion.Backward(Y, T)
    For i = m_lNumLayers To 1 Step -1
        Set dLdY = m_aoLayers(i).Backward(dLdY)
        UpdateParameters m_aoLayers(i).LearnableParameters
    Next i
End Sub

Public Sub Fit(ByVal oTrainingSet As DataLoader, _
               ByVal oTestSet As DataLoader, _
               ByVal lNumEpochs As Long)
    Const PROCEDURE_NAME As String = "Sequential.Fit"
    Dim lEpoch As Long
    Dim dblLoss As Double

    If oTrainingSet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid DataLoader object is required."
    End If
    If oTestSet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid DataLoader object is required."
    End If
    For lEpoch = 1 To lNumEpochs
        FitOneEpoch oTrainingSet
        dblLoss = Loss(oTestSet)
        LogToWorksheet "Log", "Time", GetUtcTime(), "Epoch", lEpoch, "Total number of epochs", lNumEpochs, "Learning rate", m_oOptimizer.LearningRate, "Loss", dblLoss
    Next lEpoch
End Sub

Private Sub FitOneEpoch(ByVal oTrainingSet As DataLoader)
    Dim lBatch As Long
    Dim X As Tensor
    Dim T As Tensor

    With oTrainingSet
        .Shuffle
        For lBatch = 1 To .NumBatches
            With .GetBatch(lBatch)
                Set X = .Item("Input")
                Set T = .Item("Label")
                Backpropagate X, T
            End With
            Application.StatusBar = "FitOneEpoch progress: " & Round(100 * (lBatch / .NumBatches), 1) & "%"
            DoEvents
        Next lBatch
    End With
End Sub

Private Function Loss(ByVal oTestSet As DataLoader) As Double
    Dim lBatch As Long
    Dim lNumSamples As Long
    Dim dblSum As Double
    Dim X As Tensor
    Dim T As Tensor
    Dim Y As Tensor
    
    With oTestSet
        For lBatch = 1 To .NumBatches
            With .GetBatch(lBatch)
                Set X = .Item("Input")
                Set T = .Item("Label")
                Set Y = Predict(X)
                dblSum = dblSum + m_oCriterion.Forward(Y, T).Sum
                lNumSamples = lNumSamples + .NumSamples
            End With
            Application.StatusBar = "Loss progress: " & Round(100 * (lBatch / .NumBatches), 1) & "%"
            DoEvents
        Next lBatch
    End With
    If lNumSamples > 0 Then
        Loss = dblSum / lNumSamples
    End If
End Function

Private Sub UpdateParameters(ByVal vLearnableParameters As Variant)
    Dim W As Variant

    For Each W In vLearnableParameters
        m_oOptimizer.Update W
    Next W
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Sequential.Save"
    Dim i As Long

    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteObject m_oCriterion
        .WriteObject m_oOptimizer
        .WriteLong m_lNumLayers
        For i = 1 To m_lNumLayers
            .WriteObject m_aoLayers(i)
        Next i
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Sequential.Load"
    Dim i As Long

    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        Set m_oCriterion = .ReadObject()
        Set m_oOptimizer = .ReadObject()
        m_lNumLayers = .ReadLong()
        ReDim m_aoLayers(1 To m_lNumLayers)
        For i = 1 To m_lNumLayers
            Set m_aoLayers(i) = .ReadObject()
        Next i
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Sub ISerializable_Save(ByVal oWorksheet As Object)
        Save oWorksheet
    End Sub

    Private Sub ISerializable_Load(ByVal oWorksheet As Object)
        Load oWorksheet
    End Sub
#Else
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub

    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If


