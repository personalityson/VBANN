VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeneticTrainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Dependencies
Private m_oModel As ILayer
Private m_oLossFunction As ICriterion

'Hyperparameters
Private m_lPopulationSize As Long
Private m_lNumElites As Long
Private m_lTournamentSize As Long
Private m_dblCrossoverRate As Double
Private m_dblMutationRate As Double
Private m_dblMutationSigma As Double

'State
Private m_aoPopulation() As Object
Private m_adblFitness() As Double
Private m_alRanking() As Long

Public Property Get Model() As ILayer
    Set Model = m_oModel
End Property

Public Property Get Criterion() As ICriterion
    Set Criterion = m_oLossFunction
End Property

Public Property Get PopulationSize() As Long
    PopulationSize = m_lPopulationSize
End Property

Public Property Get NumElites() As Long
    NumElites = m_lNumElites
End Property

Public Property Get TournamentSize() As Long
    TournamentSize = m_lTournamentSize
End Property

Public Property Get CrossoverRate() As Double
    CrossoverRate = m_dblCrossoverRate
End Property

Public Property Get MutationRate() As Double
    MutationRate = m_dblMutationRate
End Property

Public Property Let MutationRate(ByVal dblMutationRate As Double)
    Const PROCEDURE_NAME As String = "GeneticTrainer.MutationRate"

    If dblMutationRate < 0 Or dblMutationRate > 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Mutation rate must be >= 0 and <= 1."
    End If
    m_dblMutationRate = dblMutationRate
End Property

Public Property Get MutationSigma() As Double
    MutationSigma = m_dblMutationSigma
End Property

Public Property Let MutationSigma(ByVal dblMutationSigma As Double)
    Const PROCEDURE_NAME As String = "GeneticTrainer.MutationSigma"

    If dblMutationSigma <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Mutation sigma must be > 0."
    End If
    m_dblMutationSigma = dblMutationSigma
End Property

Public Sub Init(ByVal oModel As ILayer, _
                ByVal oLossFunction As ICriterion, _
                Optional ByVal lPopulationSize As Long = 50, _
                Optional ByVal lNumElites As Long = 2, _
                Optional ByVal lTournamentSize As Long = 3, _
                Optional ByVal dblCrossoverRate As Double = 0.9, _
                Optional ByVal dblMutationRate As Double = 0.1, _
                Optional ByVal dblMutationSigma As Double = 0.1)
    Const PROCEDURE_NAME As String = "GeneticTrainer.Init"

    If oModel Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ILayer object is required."
    End If
    If oLossFunction Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ICriterion object is required."
    End If
    If lPopulationSize < 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Population size must be >= 2."
    End If
    If lNumElites < 1 Or lNumElites >= lPopulationSize Then
        Err.Raise 5, PROCEDURE_NAME, "Number of elites must be >= 1 and < population size."
    End If
    If lTournamentSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Tournament size must be >= 1."
    End If
    If dblCrossoverRate < 0 Or dblCrossoverRate > 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Crossover rate must be >= 0 and <= 1."
    End If
    If dblMutationRate < 0 Or dblMutationRate > 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Mutation rate must be >= 0 and <= 1."
    End If
    If dblMutationSigma <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Mutation sigma must be > 0."
    End If
    Clear
    Set m_oModel = oModel
    Set m_oLossFunction = oLossFunction
    m_lPopulationSize = lPopulationSize
    m_lNumElites = lNumElites
    m_lTournamentSize = lTournamentSize
    m_dblCrossoverRate = dblCrossoverRate
    m_dblMutationRate = dblMutationRate
    m_dblMutationSigma = dblMutationSigma
    InitializePopulation
End Sub

Public Sub Fit(ByVal oTrainingSet As IDataset, _
               ByVal oTestSet As IDataset, _
               ByVal lNumGenerations As Long)
    Const PROCEDURE_NAME As String = "GeneticTrainer.Fit"
    Dim i As Long
    Dim dblTestLoss As Double

    If oTrainingSet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IDataset object is required."
    End If
    If oTestSet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IDataset object is required."
    End If
    If oTrainingSet.NumSamples < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Training dataset must contain at least one sample."
    End If
    If oTestSet.NumSamples < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Test dataset must contain at least one sample."
    End If
    If lNumGenerations < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of generations must be >= 1."
    End If
    For i = 1 To lNumGenerations
        EvaluatePopulation oTrainingSet.Cache
        RankPopulation
        LoadGenome m_aoPopulation(m_alRanking(1))
        dblTestLoss = Evaluate(oTestSet)
        WriteLog "Log", "Name", TypeName(Me), _
                 "Time", GetUtcTime(), _
                 "Epoch", i, _
                 "Total number of epochs", lNumGenerations, _
                 "Mutation sigma", m_dblMutationSigma, _
                 "Training loss", m_adblFitness(m_alRanking(1)), _
                 "Test loss", dblTestLoss
        Evolve
    Next i
End Sub

Public Function Evaluate(ByVal oDataset As IDataset) As Double
    Const PROCEDURE_NAME As String = "GeneticTrainer.Evaluate"

    If oDataset Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IDataset object is required."
    End If
    If oDataset.NumSamples < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Dataset must contain at least one sample."
    End If
    Evaluate = CalculateLoss(oDataset.Cache)
End Function

Private Sub Clear()
    Set m_oModel = Nothing
    Set m_oLossFunction = Nothing
    m_lPopulationSize = 0
    m_lNumElites = 0
    m_lTournamentSize = 0
    m_dblCrossoverRate = 0
    m_dblMutationRate = 0
    m_dblMutationSigma = 0
    Erase m_aoPopulation
    Erase m_adblFitness
    Erase m_alRanking
End Sub

Private Sub InitializePopulation()
    Dim i As Long

    ReDim m_aoPopulation(1 To m_lPopulationSize)
    Set m_aoPopulation(1) = CaptureGenome()
    For i = 2 To m_lPopulationSize
        Set m_aoPopulation(i) = CloneGenome(m_aoPopulation(1))
        Mutate m_aoPopulation(i), 1, m_dblMutationSigma
    Next i
End Sub

Private Function CaptureGenome() As Object
    Const PROCEDURE_NAME As String = "GeneticTrainer.CaptureGenome"
    Dim W As Variant
    Dim oGenome As Object

    Set oGenome = CreateObject("Scripting.Dictionary")
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
        Set oGenome.Item(W) = W.Tensor.Clone
    Next W
    Set CaptureGenome = oGenome
End Function

Private Function CloneGenome(ByVal oGenome As Object) As Object
    Dim W As Variant
    Dim A As Tensor
    Dim oClone As Object

    Set oClone = CreateObject("Scripting.Dictionary")
    For Each W In oGenome.Keys
        Set A = oGenome.Item(W)
        Set oClone.Item(W) = A.Clone
    Next W
    Set CloneGenome = oClone
End Function

'The model and the best genome share weights. This requires NumElites >= 1.
Private Sub LoadGenome(ByVal oGenome As Object)
    Dim W As Variant
    Dim A As Tensor

    For Each W In oGenome.Keys
        Set A = oGenome.Item(W)
        Set W.Tensor = A
    Next W
End Sub

Private Sub EvaluatePopulation(ByVal oTrainingData As TensorDataset)
    Dim i As Long

    ReDim m_adblFitness(1 To m_lPopulationSize)
    For i = 1 To m_lPopulationSize
        LoadGenome m_aoPopulation(i)
        m_adblFitness(i) = CalculateLoss(oTrainingData)
        Application.StatusBar = "Fitness progress: " & Format(i / m_lPopulationSize, "0.0%")
        DoEvents
    Next i
End Sub

Private Function CalculateLoss(ByVal oData As TensorDataset) As Double
    Dim X As Tensor
    Dim T As Tensor
    Dim Y As Tensor

    Set X = oData.Tensor(1)
    Set T = oData.Tensor(2)
    Set Y = m_oModel.Forward(X, False)
    CalculateLoss = m_oLossFunction.Forward(Y, T).Sum / oData.NumSamples
End Function

Private Sub RankPopulation()
    Dim i As Long
    Dim j As Long
    Dim lIndex As Long

    m_alRanking = GetIdentityPermutationArray(m_lPopulationSize)
    For i = 2 To m_lPopulationSize
        lIndex = m_alRanking(i)
        j = i - 1
        Do While j >= 1
            If m_adblFitness(m_alRanking(j)) <= m_adblFitness(lIndex) Then
                Exit Do
            End If
            m_alRanking(j + 1) = m_alRanking(j)
            j = j - 1
        Loop
        m_alRanking(j + 1) = lIndex
    Next i
End Sub

Private Sub Evolve()
    Dim i As Long
    Dim aoNewPopulation() As Object

    ReDim aoNewPopulation(1 To m_lPopulationSize)
    For i = 1 To m_lNumElites
        Set aoNewPopulation(i) = m_aoPopulation(m_alRanking(i))
    Next i
    For i = m_lNumElites + 1 To m_lPopulationSize
        Set aoNewPopulation(i) = Breed()
    Next i
    m_aoPopulation = aoNewPopulation
End Sub

Private Function Breed() As Object
    Dim oChild As Object

    If Rnd() < m_dblCrossoverRate Then
        Set oChild = Crossover(m_aoPopulation(SelectParent()), m_aoPopulation(SelectParent()))
    Else
        Set oChild = CloneGenome(m_aoPopulation(SelectParent()))
    End If
    Mutate oChild, m_dblMutationRate, m_dblMutationSigma
    Set Breed = oChild
End Function

Private Function SelectParent() As Long
    Dim i As Long
    Dim lBestRank As Long

    lBestRank = m_lPopulationSize
    For i = 1 To m_lTournamentSize
        lBestRank = MinLng(lBestRank, Int(m_lPopulationSize * Rnd()) + 1)
    Next i
    SelectParent = m_alRanking(lBestRank)
End Function

Private Function Crossover(ByVal oParentA As Object, _
                           ByVal oParentB As Object) As Object
    Dim W As Variant
    Dim A As Tensor
    Dim B As Tensor
    Dim oChild As Object

    Set oChild = CreateObject("Scripting.Dictionary")
    For Each W In oParentA.Keys
        Set A = oParentA.Item(W)
        Set B = oParentB.Item(W)
        Set oChild.Item(W) = VecWhere(Bernoulli(A.Shape, 0.5), A, B)
    Next W
    Set Crossover = oChild
End Function

Private Sub Mutate(ByVal oGenome As Object, _
                   ByVal dblRate As Double, _
                   ByVal dblSigma As Double)
    Dim A As Variant

    For Each A In oGenome.Items
        VecAdd_I A, VecMul(Bernoulli(A.Shape, dblRate), Normal(A.Shape, 0, dblSigma))
    Next A
End Sub

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "GeneticTrainer.Save"
    Dim i As Long
    Dim W As Variant
    Dim A As Tensor
    Dim oGenome As Object

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteObject m_oModel
        .WriteObject m_oLossFunction
        .WriteLong m_lPopulationSize
        .WriteLong m_lNumElites
        .WriteLong m_lTournamentSize
        .WriteDouble m_dblCrossoverRate
        .WriteDouble m_dblMutationRate
        .WriteDouble m_dblMutationSigma
        For i = 1 To m_lPopulationSize
            Set oGenome = m_aoPopulation(i)
            .WriteLong oGenome.Count
            For Each W In oGenome.Keys
                Set A = oGenome.Item(W)
                .WriteObject W
                .WriteObject A
            Next W
        Next i
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "GeneticTrainer.Load"
    Dim i As Long
    Dim j As Long
    Dim lNumParameters As Long
    Dim W As Parameter
    Dim A As Tensor
    Dim oGenome As Object

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oSerializer
        Set m_oModel = .ReadObject()
        Set m_oLossFunction = .ReadObject()
        m_lPopulationSize = .ReadLong()
        m_lNumElites = .ReadLong()
        m_lTournamentSize = .ReadLong()
        m_dblCrossoverRate = .ReadDouble()
        m_dblMutationRate = .ReadDouble()
        m_dblMutationSigma = .ReadDouble()
        If m_lPopulationSize > 0 Then
            ReDim m_aoPopulation(1 To m_lPopulationSize)
            For i = 1 To m_lPopulationSize
                Set oGenome = CreateObject("Scripting.Dictionary")
                lNumParameters = .ReadLong()
                For j = 1 To lNumParameters
                    Set W = .ReadObject()
                    Set A = .ReadObject()
                    Set oGenome.Item(W) = A
                Next j
                Set m_aoPopulation(i) = oGenome
            Next i
        End If
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

