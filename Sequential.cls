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

Private m_lNumLayers As Long
Private m_aoLayers() As ILayer
Private m_aoStates() As Tensor
Private m_oCriterion As ICriterion
Private m_oOptimizer As IOptimizer

Public Property Get NumLayers() As Long
    NumLayers = m_lNumLayers
End Property

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
    m_lNumLayers = 0
    Erase m_aoLayers
    Erase m_aoStates
    Set m_oCriterion = oCriterion
    Set m_oOptimizer = oOptimizer
End Function

Public Function GetLayer(ByVal lIndex As Long) As ILayer
    Const PROCEDURE_NAME As String = "Sequential.GetLayer"
    
    If lIndex < 1 Or lIndex > m_lNumLayers Then
        Err.Raise 9, PROCEDURE_NAME, "Layer index must be between 1 and the number of layers."
    End If
    Set GetLayer = m_aoLayers(lIndex)
End Function

Public Function GetState(ByVal lIndex As Long) As Tensor
    Const PROCEDURE_NAME As String = "Sequential.GetState"
    
    If lIndex < 1 Or lIndex > m_lNumLayers Then
        Err.Raise 9, PROCEDURE_NAME, "Layer index must be between 1 and the number of layers."
    End If
    Set GetState = m_aoStates(lIndex)
End Function

Public Sub Add(ByVal oLayer As ILayer)
    Const PROCEDURE_NAME As String = "Sequential.Add"
    
    If oLayer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ILayer object is required."
    End If
    m_lNumLayers = m_lNumLayers + 1
    ReDim Preserve m_aoLayers(1 To m_lNumLayers)
    ReDim Preserve m_aoStates(0 To m_lNumLayers)
    Set m_aoLayers(m_lNumLayers) = oLayer
End Sub

Public Function Predict(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean) As Tensor
    Dim i As Long
    
    ReDim m_aoStates(0 To m_lNumLayers)
    Set m_aoStates(0) = X
    For i = 1 To m_lNumLayers
        Set X = m_aoStates(i - 1)
        Set m_aoStates(i) = m_aoLayers(i).Forward(X, bIsTraining)
    Next i
    Set Predict = m_aoStates(m_lNumLayers)
End Function

Public Sub Backpropagate(ByVal X As Tensor, _
                         ByVal T As Tensor)
    Dim i As Long
    Dim Y As Tensor
    Dim dLdY As Tensor
    
    Set Y = Predict(X, True)
    Set dLdY = m_oCriterion.Backward(Y, T)
    For i = m_lNumLayers To 1 Step -1
        Set X = m_aoStates(i - 1)
        Set Y = m_aoStates(i)
        Set dLdY = m_aoLayers(i).Backward(X, Y, dLdY)
        UpdateParameters m_aoLayers(i).Parameters
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

Public Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Sequential.Save"
    Dim i As Long
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteLong m_lNumLayers
        For i = 1 To m_lNumLayers
            .WriteObject m_aoLayers(i)
        Next i
        .WriteObject m_oCriterion
        .WriteObject m_oOptimizer
    End With
End Sub

Public Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Sequential.Load"
    Dim i As Long
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        m_lNumLayers = .ReadLong()
        ReDim m_aoLayers(1 To m_lNumLayers)
        ReDim m_aoStates(0 To m_lNumLayers)
        For i = 1 To m_lNumLayers
            Set m_aoLayers(i) = .ReadObject()
        Next i
        Set m_oCriterion = .ReadObject()
        Set m_oOptimizer = .ReadObject()
    End With
End Sub

Private Sub UpdateParameters(ByVal vParameters As Variant)
    Dim W As Variant
    
    For Each W In vParameters
        m_oOptimizer.Update W
    Next W
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
