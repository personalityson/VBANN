VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Evolution Strategies (Salimans et al. 2017) with mirrored sampling, centered-rank fitness shaping and mini-batch fitness.

'Dependencies
Private m_oModel As ILayer
Private m_oLossFunction As ILossFunction
Private m_oOptimizer As IOptimizer

'Hyperparameters
Private m_lPopulationSize As Long
Private m_dblMutationSigma As Double

'Training state
Private m_oParameters As Collection

Public Property Get Model() As ILayer
    Set Model = m_oModel
End Property

Public Property Get LossFunction() As ILossFunction
    Set LossFunction = m_oLossFunction
End Property

Public Property Get Optimizer() As IOptimizer
    Set Optimizer = m_oOptimizer
End Property

Public Property Get PopulationSize() As Long
    PopulationSize = m_lPopulationSize
End Property

Public Property Get MutationSigma() As Double
    MutationSigma = m_dblMutationSigma
End Property

Public Property Let MutationSigma(ByVal dblMutationSigma As Double)
    Const PROCEDURE_NAME As String = "OpenES.MutationSigma"

    If dblMutationSigma <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Mutation sigma must be > 0."
    End If
    m_dblMutationSigma = dblMutationSigma
End Property

Public Sub Init(ByVal oModel As ILayer, _
                ByVal oLossFunction As ILossFunction, _
                ByVal oOptimizer As IOptimizer, _
                Optional ByVal lPopulationSize As Long = 50, _
                Optional ByVal dblMutationSigma As Double = 0.02)
    Const PROCEDURE_NAME As String = "OpenES.Init"

    If oModel Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ILayer object is required."
    End If
    If oLossFunction Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ILossFunction object is required."
    End If
    If oOptimizer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IOptimizer object is required."
    End If
    If lPopulationSize < 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Population size must be >= 2."
    End If
    If lPopulationSize Mod 2 <> 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Population size must be even (mirrored sampling)."
    End If
    If dblMutationSigma <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Mutation sigma must be > 0."
    End If
    Clear
    Set m_oModel = oModel
    Set m_oLossFunction = oLossFunction
    Set m_oOptimizer = oOptimizer
    m_lPopulationSize = lPopulationSize
    m_dblMutationSigma = dblMutationSigma
End Sub

Public Sub Fit(ByVal oTrainingLoader As DataLoader, _
               ByVal lNumEpochs As Long, _
               Optional ByVal bVerbose As Boolean, _
               Optional ByVal oValidationLoader As DataLoader)
    Const PROCEDURE_NAME As String = "OpenES.Fit"
    Dim i As Long
    Dim dblTrainingLoss As Double
    Dim dblValidationLoss As Double

    If oTrainingLoader Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid DataLoader object is required."
    End If
    If oTrainingLoader.NumTensors < 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Data loader must contain an input tensor and a target tensor."
    End If
    If lNumEpochs < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of epochs must be >= 1."
    End If
    CollectParameters
    For i = 1 To lNumEpochs
        FitOne oTrainingLoader
        If bVerbose Then
            dblTrainingLoss = Me.Evaluate(oTrainingLoader)
            WriteLog "Log", True, "Name", TypeName(Me), _
                     "Time", GetUtcTime(), _
                     "Iteration", i, _
                     "Total number of iterations", lNumEpochs, _
                     "Learning rate", m_oOptimizer.LearningRate, _
                     "Mutation sigma", m_dblMutationSigma, _
                     "Training loss", dblTrainingLoss
            If Not oValidationLoader Is Nothing Then
                dblValidationLoss = Me.Evaluate(oValidationLoader)
                WriteLog "Log", False, "Validation loss", dblValidationLoss
            End If
        End If
    Next i
End Sub

Public Function Evaluate(ByVal oDataLoader As DataLoader) As Double
    Const PROCEDURE_NAME As String = "OpenES.Evaluate"
    Dim lNumSamples As Long
    Dim dblLoss As Double
    Dim oBatch As TensorDataset

    If oDataLoader Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid DataLoader object is required."
    End If
    If oDataLoader.NumTensors < 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Data loader must contain an input tensor and a target tensor."
    End If
    With oDataLoader
        .Reset
        Do While .NextBatch(oBatch)
            dblLoss = dblLoss + CalculateLoss(oBatch.Tensors(1), oBatch.Tensors(2))
            lNumSamples = lNumSamples + oBatch.NumSamples
            Application.StatusBar = "Evaluate progress: " & Format(lNumSamples / .NumSamples, "0.0%")
            DoEvents
        Loop
    End With
    If lNumSamples > 0 Then
        Evaluate = dblLoss / lNumSamples
    End If
End Function

Private Sub Clear()
    Set m_oModel = Nothing
    Set m_oLossFunction = Nothing
    Set m_oOptimizer = Nothing
    m_lPopulationSize = 0
    m_dblMutationSigma = 0
    Set m_oParameters = New Collection
End Sub

Private Sub CollectParameters()
    Const PROCEDURE_NAME As String = "OpenES.CollectParameters"
    Dim W As Variant

    Set m_oParameters = New Collection
    For Each W In EnsureIterable(m_oModel.Learnables)
        If Not IsObject(W) Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
        End If
        If W Is Nothing Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
        End If
        If Not TypeOf W Is Parameter Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
        End If
        m_oParameters.Add W
    Next W
End Sub

Private Sub FitOne(ByVal oTrainingLoader As DataLoader)
    Dim lNumSamples As Long
    Dim X As Tensor
    Dim T As Tensor
    Dim oBatch As TensorDataset

    With oTrainingLoader
        .Shuffle
        Do While .NextBatch(oBatch)
            Set X = oBatch.Tensors(1)
            Set T = oBatch.Tensors(2)
            Evolve X, T
            m_oOptimizer.Update m_oModel.Learnables
            lNumSamples = lNumSamples + oBatch.NumSamples
            Application.StatusBar = "FitOne progress: " & Format(lNumSamples / .NumSamples, "0.0%")
            DoEvents
        Loop
    End With
End Sub

Private Sub Evolve(ByVal X As Tensor, _
                   ByVal T As Tensor)
    Dim k As Long
    Dim lNumPairs As Long
    Dim adblLosses() As Double
    Dim aoNoise() As Collection

    lNumPairs = m_lPopulationSize \ 2
    ReDim adblLosses(1 To m_lPopulationSize)
    ReDim aoNoise(1 To lNumPairs)
    For k = 1 To lNumPairs
        Set aoNoise(k) = SampleNoise()
        ShiftWeights aoNoise(k), m_dblMutationSigma
        adblLosses(2 * k - 1) = CalculateLoss(X, T)
        ShiftWeights aoNoise(k), -2 * m_dblMutationSigma
        adblLosses(2 * k) = CalculateLoss(X, T)
        ShiftWeights aoNoise(k), m_dblMutationSigma
    Next k
    AccumulateGradient aoNoise, CalculateCenteredRanks(adblLosses)
End Sub

Private Function SampleNoise() As Collection
    Dim W As Parameter

    Set SampleNoise = New Collection
    For Each W In m_oParameters
        SampleNoise.Add Normal(W.Tensor.Shape)
    Next W
End Function

Private Sub ShiftWeights(ByVal oNoise As Collection, _
                         ByVal dblStep As Double)
    Dim i As Long

    For i = 1 To m_oParameters.Count
        VecLinComb_I 1, m_oParameters(i).Tensor, dblStep, oNoise(i)
    Next i
End Sub

Private Sub AccumulateGradient(ByRef aoNoise() As Collection, _
                               ByRef adblRanks() As Double)
    Dim i As Long
    Dim k As Long
    Dim dblScale As Double

    dblScale = 1 / (m_lPopulationSize * m_dblMutationSigma)
    For i = 1 To m_oParameters.Count
        With m_oParameters(i)
            For k = 1 To m_lPopulationSize \ 2
                VecLinComb_I 1, .Gradient, dblScale * (adblRanks(2 * k - 1) - adblRanks(2 * k)), aoNoise(k).Item(i)
            Next k
            .NumAccumulatedSamples = .NumAccumulatedSamples + 1
        End With
    Next i
End Sub

Private Function CalculateLoss(ByVal X As Tensor, _
                               ByVal T As Tensor) As Double
    CalculateLoss = m_oLossFunction.Loss(m_oModel.Forward(X), T).Sum
End Function

Private Function CalculateCenteredRanks(ByRef adblLosses() As Double) As Double()
    Dim i As Long
    Dim j As Long
    Dim lRank As Long
    Dim adblResult() As Double

    ReDim adblResult(1 To m_lPopulationSize)
    For i = 1 To m_lPopulationSize
        lRank = 0
        For j = 1 To m_lPopulationSize
            If adblLosses(j) < adblLosses(i) Or (adblLosses(j) = adblLosses(i) And j < i) Then
                lRank = lRank + 1
            End If
        Next j
        adblResult(i) = lRank / (m_lPopulationSize - 1) - 0.5
    Next i
    CalculateCenteredRanks = adblResult
End Function

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "OpenES.Save"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteObject m_oModel
        .WriteObject m_oLossFunction
        .WriteObject m_oOptimizer
        .WriteLong m_lPopulationSize
        .WriteDouble m_dblMutationSigma
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "OpenES.Load"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oSerializer
        Set m_oModel = .ReadObject()
        Set m_oLossFunction = .ReadObject()
        Set m_oOptimizer = .ReadObject()
        m_lPopulationSize = .ReadLong()
        m_dblMutationSigma = .ReadDouble()
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Sub ISerializable_Save(ByVal oSerializer As Object)
        Save oSerializer
    End Sub

    Private Sub ISerializable_Load(ByVal oSerializer As Object)
        Load oSerializer
    End Sub
#Else
    Private Sub ISerializable_Save(ByVal oSerializer As Serializer)
        Save oSerializer
    End Sub

    Private Sub ISerializable_Load(ByVal oSerializer As Serializer)
        Load oSerializer
    End Sub
#End If

