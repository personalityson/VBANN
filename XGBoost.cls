VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XGBoost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Gradient boosted trees, grown depth first.

Private Const MAX_DEPTH As Long = 12
Private Const MAX_QUANTILE_SAMPLES As Long = 200000

Private Type Bucket
    SumG As Double
    SumH As Double
    Count As Long
End Type

Private Type Split
    Feature As Long
    Bin As Long
    Threshold As Double
    Gain As Double
End Type

'Dependencies
Private m_oLossFunction As ILossFunction

'Hyperparameters
Private m_dblLearningRate As Double
Private m_lMaxDepth As Long
Private m_dblLambda As Double
Private m_dblGamma As Double
Private m_dblMinChildWeight As Double
Private m_dblSubsample As Double
Private m_dblColsample As Double
Private m_lNumBins As Long

'Model, m_aoTrees(output, round)
Private m_lNumFeatures As Long
Private m_lNumOutputs As Long
Private m_lNumRounds As Long
Private m_oBaseScore As Tensor
Private m_aoTrees() As XGBoostTree

'Training state, discarded once Fit returns
Private m_lMaxBins As Long
Private m_alNumCuts() As Long
Private m_adblCuts() As Double
Private m_alBins() As Long

Public Property Get NumFeatures() As Long
    NumFeatures = m_lNumFeatures
End Property

Public Property Get NumOutputs() As Long
    NumOutputs = m_lNumOutputs
End Property

Public Property Get NumRounds() As Long
    NumRounds = m_lNumRounds
End Property

Public Sub Init(ByVal oLossFunction As ILossFunction, _
                Optional ByVal dblLearningRate As Double = 0.1, _
                Optional ByVal lMaxDepth As Long = 6, _
                Optional ByVal dblLambda As Double = 1, _
                Optional ByVal dblGamma As Double = 0, _
                Optional ByVal dblMinChildWeight As Double = 1, _
                Optional ByVal dblSubsample As Double = 1, _
                Optional ByVal dblColsample As Double = 1, _
                Optional ByVal lNumBins As Long = 256)
    Const PROCEDURE_NAME As String = "XGBoost.Init"

    If oLossFunction Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ILossFunction object is required."
    End If
    If Not oLossFunction.FromLogits Then
        Err.Raise 5, PROCEDURE_NAME, "XGBoost requires a loss function that works on logits."
    End If
    If dblLearningRate <= 0 Or dblLearningRate > 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Learning rate must be > 0 and <= 1."
    End If
    If lMaxDepth < 1 Or lMaxDepth > MAX_DEPTH Then
        Err.Raise 5, PROCEDURE_NAME, "Max depth must be >= 1 and <= " & MAX_DEPTH & "."
    End If
    If dblLambda < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Lambda (L2 regularization) must be >= 0."
    End If
    If dblGamma < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Gamma (minimum split gain) must be >= 0."
    End If
    If dblMinChildWeight < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Minimum child weight must be >= 0."
    End If
    If dblSubsample <= 0 Or dblSubsample > 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Subsample (row fraction per tree) must be > 0 and <= 1."
    End If
    If dblColsample <= 0 Or dblColsample > 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Colsample (feature fraction per tree) must be > 0 and <= 1."
    End If
    If lNumBins < 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of bins must be >= 2."
    End If
    Clear
    Set m_oLossFunction = oLossFunction
    m_dblLearningRate = dblLearningRate
    m_lMaxDepth = lMaxDepth
    m_dblLambda = dblLambda
    m_dblGamma = dblGamma
    m_dblMinChildWeight = dblMinChildWeight
    m_dblSubsample = dblSubsample
    m_dblColsample = dblColsample
    m_lNumBins = lNumBins
End Sub

Public Sub Fit(ByVal oTrainingSet As IDataset, _
               ByVal lNumRounds As Long, _
               Optional ByVal bVerbose As Boolean, _
               Optional ByVal oValidationSet As IDataset)
    Const PROCEDURE_NAME As String = "XGBoost.Fit"
    Dim i As Long
    Dim lNumSamples As Long
    Dim dblTrainingLoss As Double
    Dim dblValidationLoss As Double
    Dim X As Tensor
    Dim T As Tensor
    Dim Y As Tensor

    If oTrainingSet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IDataset object is required."
    End If
    If oTrainingSet.NumTensors < 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Dataset must contain an input tensor and a target tensor."
    End If
    Set X = oTrainingSet.Tensors(1)
    Set T = oTrainingSet.Tensors(2)
    If X.NumDimensions <> 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor must have 2 dimensions."
    End If
    If T.NumDimensions <> 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Target tensor must have 2 dimensions."
    End If
    If lNumRounds < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of rounds must be >= 1."
    End If
    m_lNumFeatures = X.Size(1)
    m_lNumOutputs = T.Size(1)
    m_lNumRounds = 0
    Set m_oBaseScore = m_oLossFunction.Intercept(T)
    ReDim m_aoTrees(1 To m_lNumOutputs, 1 To lNumRounds)
    BuildBins X
    lNumSamples = X.Size(2)
    Set Y = m_oBaseScore.Tile(2, lNumSamples)
    For i = 1 To lNumRounds
        FitOne i, X, Y, T
        m_lNumRounds = i
        If bVerbose Then
            dblTrainingLoss = Me.Evaluate(oTrainingSet)
            WriteLog "Log", True, "Name", TypeName(Me), _
                     "Time", GetUtcTime(), _
                     "Iteration", i, _
                     "Total number of iterations", lNumRounds, _
                     "Learning rate", m_dblLearningRate, _
                     "Training loss", dblTrainingLoss
            If Not oValidationSet Is Nothing Then
                dblValidationLoss = Me.Evaluate(oValidationSet)
                WriteLog "Log", False, "Validation loss", dblValidationLoss
            End If
        End If
        Application.StatusBar = "Fit progress: " & Format(i / lNumRounds, "0.0%")
        DoEvents
    Next i
    ReleaseTrainingState
End Sub

Public Function Evaluate(ByVal oDataset As IDataset) As Double
    Const PROCEDURE_NAME As String = "XGBoost.Evaluate"
    Dim X As Tensor
    Dim T As Tensor

    If m_lNumRounds < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Model has not been fitted."
    End If
    If oDataset Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IDataset object is required."
    End If
    If oDataset.NumTensors < 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Dataset must contain an input tensor and a target tensor."
    End If
    Set X = oDataset.Tensors(1)
    Set T = oDataset.Tensors(2)
    If X.NumDimensions <> 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor must have 2 dimensions."
    End If
    If T.NumDimensions <> 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Target tensor must have 2 dimensions."
    End If
    If X.Size(1) <> m_lNumFeatures Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor is not compatible with the trained model."
    End If
    If T.Size(1) <> m_lNumOutputs Then
        Err.Raise 5, PROCEDURE_NAME, "Target tensor is not compatible with the trained model."
    End If
    If X.Size(2) > 0 Then
        Evaluate = m_oLossFunction.Loss(PredictRaw(X), T).Sum / X.Size(2)
    End If
End Function

Public Function Predict(ByVal X As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "XGBoost.Predict"

    If m_lNumRounds < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Model has not been fitted."
    End If
    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If X.NumDimensions <> 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor must have 2 dimensions."
    End If
    If X.Size(1) <> m_lNumFeatures Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor is not compatible with the trained model."
    End If
    Set Predict = m_oLossFunction.Transform(PredictRaw(X))
End Function

Private Function PredictRaw(ByVal X As Tensor) As Tensor
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim lNumSamples As Long
    Dim X_() As Double
    Dim Y_() As Double
    Dim Y As Tensor
    Dim oTree As XGBoostTree

    lNumSamples = X.Size(2)
    Set Y = m_oBaseScore.Tile(2, lNumSamples)
    X.CreateAlias X_
    Y.CreateAlias Y_
    For i = 1 To m_lNumRounds
        For k = 1 To m_lNumOutputs
            Set oTree = m_aoTrees(k, i)
            For j = 1 To lNumSamples
                Y_(k, j) = Y_(k, j) + oTree.PredictRow(X_, j)
            Next j
        Next k
    Next i
    X.RemoveAlias X_
    Y.RemoveAlias Y_
    Set PredictRaw = Y
End Function

Private Sub FitOne(ByVal lRound As Long, _
                   ByVal X As Tensor, _
                   ByVal Y As Tensor, _
                   ByVal T As Tensor)
    Dim k As Long
    Dim lNumSamples As Long
    Dim X_() As Double
    Dim Y_() As Double
    Dim G_() As Double
    Dim H_() As Double
    Dim G As Tensor
    Dim H As Tensor

    lNumSamples = X.Size(2)
    Set G = m_oLossFunction.Gradient(Y, T)
    Set H = m_oLossFunction.Hessian(Y, T)
    X.CreateAlias X_
    Y.CreateAlias Y_
    G.CreateAlias G_
    H.CreateAlias H_
    For k = 1 To m_lNumOutputs
        Set m_aoTrees(k, lRound) = BuildTree(k, lNumSamples, X_, Y_, G_, H_)
    Next k
    X.RemoveAlias X_
    Y.RemoveAlias Y_
    G.RemoveAlias G_
    H.RemoveAlias H_
End Sub

Private Function BuildTree(ByVal lOutput As Long, _
                           ByVal lNumSamples As Long, _
                           ByRef X_() As Double, _
                           ByRef Y_() As Double, _
                           ByRef G_() As Double, _
                           ByRef H_() As Double) As XGBoostTree
    Dim lNumDrawnSamples As Long
    Dim lNumDrawnFeatures As Long
    Dim alIndex() As Long
    Dim alFeatures() As Long
    Dim auHist() As Bucket
    Dim oTree As XGBoostTree

    lNumDrawnSamples = MaxLng(1, Int(m_dblSubsample * lNumSamples))
    lNumDrawnFeatures = MaxLng(1, Int(m_dblColsample * m_lNumFeatures))
    alIndex = GetRandomPermutation(lNumSamples)
    alFeatures = GetRandomPermutation(m_lNumFeatures)
    Set oTree = New XGBoostTree
    oTree.Init CLng(2 ^ (m_lMaxDepth + 1)) - 1
    BuildHistogram auHist, alIndex, 1, lNumDrawnSamples, alFeatures, lNumDrawnFeatures, G_, H_, lOutput
    BuildNode oTree, auHist, alIndex, 1, lNumDrawnSamples, alFeatures, lNumDrawnFeatures, Y_, G_, H_, lOutput, 0
    oTree.Complete
    ScoreHeldOutRows oTree, X_, Y_, alIndex, lNumSamples, lNumDrawnSamples, lOutput
    Set BuildTree = oTree
End Function

Private Sub ScoreHeldOutRows(ByVal oTree As XGBoostTree, _
                             ByRef X_() As Double, _
                             ByRef Y_() As Double, _
                             ByRef alIndex() As Long, _
                             ByVal lNumSamples As Long, _
                             ByVal lNumDrawnSamples As Long, _
                             ByVal lOutput As Long)
    Dim j As Long
    Dim s As Long

    If lNumDrawnSamples >= lNumSamples Then
        Exit Sub
    End If
    For j = lNumDrawnSamples + 1 To lNumSamples
        s = alIndex(j)
        Y_(lOutput, s) = Y_(lOutput, s) + oTree.PredictRow(X_, s)
    Next j
End Sub

Private Function BuildNode(ByVal oTree As XGBoostTree, _
                           ByRef auHist() As Bucket, _
                           ByRef alIndex() As Long, _
                           ByVal lStart As Long, _
                           ByVal lEnd As Long, _
                           ByRef alFeatures() As Long, _
                           ByVal lNumDrawnFeatures As Long, _
                           ByRef Y_() As Double, _
                           ByRef G_() As Double, _
                           ByRef H_() As Double, _
                           ByVal lOutput As Long, _
                           ByVal lDepth As Long) As Long
    Dim j As Long
    Dim s As Long
    Dim lNode As Long
    Dim lCount As Long
    Dim lMid As Long
    Dim lLeftNode As Long
    Dim lRightNode As Long
    Dim dblSumG As Double
    Dim dblSumH As Double
    Dim dblWeight As Double
    Dim uSplit As Split
    Dim auLeftHist() As Bucket
    Dim auRightHist() As Bucket

    lNode = oTree.AddNode()
    BuildNode = lNode
    HistogramTotals auHist, dblSumG, dblSumH, lCount
    If lDepth < m_lMaxDepth Then
        uSplit = FindBestSplit(auHist, alFeatures, lNumDrawnFeatures, dblSumG, dblSumH, lCount)
    End If
    'No split, this node is a leaf
    If uSplit.Feature = 0 Then
        dblWeight = m_dblLearningRate * LeafWeight(dblSumG, dblSumH)
        oTree.SetLeaf lNode, dblWeight
        For j = lStart To lEnd
            s = alIndex(j)
            Y_(lOutput, s) = Y_(lOutput, s) + dblWeight
        Next j
        Exit Function
    End If
    'Left alIndex(lStart To lMid - 1)
    'Right alIndex(lMid To lEnd)
    lMid = PartitionRows(uSplit, alIndex, lStart, lEnd)
    If lMid - lStart <= lEnd - lMid + 1 Then
        BuildHistogram auLeftHist, alIndex, lStart, lMid - 1, alFeatures, lNumDrawnFeatures, G_, H_, lOutput
        SubtractHistogram auHist, auLeftHist, auRightHist, lNumDrawnFeatures
    Else
        BuildHistogram auRightHist, alIndex, lMid, lEnd, alFeatures, lNumDrawnFeatures, G_, H_, lOutput
        SubtractHistogram auHist, auRightHist, auLeftHist, lNumDrawnFeatures
    End If
    Erase auHist
    lLeftNode = BuildNode(oTree, auLeftHist, alIndex, lStart, lMid - 1, alFeatures, lNumDrawnFeatures, Y_, G_, H_, lOutput, lDepth + 1)
    lRightNode = BuildNode(oTree, auRightHist, alIndex, lMid, lEnd, alFeatures, lNumDrawnFeatures, Y_, G_, H_, lOutput, lDepth + 1)
    oTree.SetSplit lNode, uSplit.Feature, uSplit.Threshold, lLeftNode, lRightNode
End Function

Private Function FindBestSplit(ByRef auHist() As Bucket, _
                               ByRef alFeatures() As Long, _
                               ByVal lNumDrawnFeatures As Long, _
                               ByVal dblSumG As Double, _
                               ByVal dblSumH As Double, _
                               ByVal lCount As Long) As Split
    Dim i As Long
    Dim k As Long
    Dim lBin As Long
    Dim lNumCuts As Long
    Dim lCountL As Long
    Dim dblSumGL As Double
    Dim dblSumHL As Double
    Dim dblSumHR As Double
    Dim dblGain As Double
    Dim dblNodeObjective As Double
    Dim uBest As Split

    dblNodeObjective = NodeObjective(dblSumG, dblSumH)
    For k = 1 To lNumDrawnFeatures
        i = alFeatures(k)
        lNumCuts = m_alNumCuts(i)
        dblSumGL = 0
        dblSumHL = 0
        lCountL = 0
        For lBin = 1 To lNumCuts
            With auHist(k, lBin)
                dblSumGL = dblSumGL + .SumG
                dblSumHL = dblSumHL + .SumH
                lCountL = lCountL + .Count
            End With
            dblSumHR = dblSumH - dblSumHL
            If lCountL >= 1 And lCount - lCountL >= 1 Then
                If dblSumHL >= m_dblMinChildWeight And dblSumHR >= m_dblMinChildWeight Then
                    dblGain = dblNodeObjective _
                            - NodeObjective(dblSumGL, dblSumHL) _
                            - NodeObjective(dblSumG - dblSumGL, dblSumHR) _
                            - m_dblGamma
                    If dblGain > uBest.Gain Then
                        uBest.Feature = i
                        uBest.Bin = lBin
                        uBest.Threshold = m_adblCuts(i, lBin)
                        uBest.Gain = dblGain
                    End If
                End If
            End If
        Next lBin
    Next k
    FindBestSplit = uBest
End Function

Private Function LeafWeight(ByVal dblSumG As Double, _
                            ByVal dblSumH As Double) As Double
    If dblSumH + m_dblLambda > 0 Then
        LeafWeight = -dblSumG / (dblSumH + m_dblLambda)
    End If
End Function

Private Function NodeObjective(ByVal dblSumG As Double, _
                               ByVal dblSumH As Double) As Double
    Dim dblWeight As Double

    dblWeight = LeafWeight(dblSumG, dblSumH)
    NodeObjective = dblSumG * dblWeight + 0.5 * (dblSumH + m_dblLambda) * dblWeight * dblWeight
End Function

Private Sub BuildHistogram(ByRef auHist() As Bucket, _
                           ByRef alIndex() As Long, _
                           ByVal lStart As Long, _
                           ByVal lEnd As Long, _
                           ByRef alFeatures() As Long, _
                           ByVal lNumDrawnFeatures As Long, _
                           ByRef G_() As Double, _
                           ByRef H_() As Double, _
                           ByVal lOutput As Long)
    Dim j As Long
    Dim k As Long
    Dim s As Long
    Dim dblG As Double
    Dim dblH As Double

    ReDim auHist(1 To lNumDrawnFeatures, 1 To m_lMaxBins)
    For j = lStart To lEnd
        s = alIndex(j)
        dblG = G_(lOutput, s)
        dblH = H_(lOutput, s)
        For k = 1 To lNumDrawnFeatures
            With auHist(k, m_alBins(alFeatures(k), s))
                .SumG = .SumG + dblG
                .SumH = .SumH + dblH
                .Count = .Count + 1
            End With
        Next k
    Next j
End Sub

'auSibling = auParent - auChild, bin by bin.
Private Sub SubtractHistogram(ByRef auParent() As Bucket, _
                              ByRef auChild() As Bucket, _
                              ByRef auSibling() As Bucket, _
                              ByVal lNumDrawnFeatures As Long)
    Dim k As Long
    Dim lBin As Long

    ReDim auSibling(1 To lNumDrawnFeatures, 1 To m_lMaxBins)
    For k = 1 To lNumDrawnFeatures
        For lBin = 1 To m_lMaxBins
            With auSibling(k, lBin)
                .SumG = auParent(k, lBin).SumG - auChild(k, lBin).SumG
                .SumH = auParent(k, lBin).SumH - auChild(k, lBin).SumH
                .Count = auParent(k, lBin).Count - auChild(k, lBin).Count
            End With
        Next lBin
    Next k
End Sub

Private Sub HistogramTotals(ByRef auHist() As Bucket, _
                            ByRef dblSumG As Double, _
                            ByRef dblSumH As Double, _
                            ByRef lCount As Long)
    Dim lBin As Long

    dblSumG = 0
    dblSumH = 0
    lCount = 0
    For lBin = 1 To m_lMaxBins
        With auHist(1, lBin)
            dblSumG = dblSumG + .SumG
            dblSumH = dblSumH + .SumH
            lCount = lCount + .Count
        End With
    Next lBin
End Sub

Private Function PartitionRows(ByRef uSplit As Split, _
                               ByRef alIndex() As Long, _
                               ByVal lStart As Long, _
                               ByVal lEnd As Long) As Long
    Dim lLeft As Long
    Dim lRight As Long
    Dim lTemp As Long

    lLeft = lStart
    lRight = lEnd
    Do While lLeft <= lRight
        If m_alBins(uSplit.Feature, alIndex(lLeft)) <= uSplit.Bin Then
            lLeft = lLeft + 1
        Else
            lTemp = alIndex(lLeft)
            alIndex(lLeft) = alIndex(lRight)
            alIndex(lRight) = lTemp
            lRight = lRight - 1
        End If
    Loop
    PartitionRows = lLeft
End Function

Private Sub BuildBins(ByVal X As Tensor)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim lNumSamples As Long
    Dim lStride As Long
    Dim lNumSorted As Long
    Dim X_() As Double
    Dim adblSorted() As Double

    lNumSamples = X.Size(2)
    lStride = MaxLng(1, lNumSamples \ MAX_QUANTILE_SAMPLES)
    lNumSorted = (lNumSamples + lStride - 1) \ lStride
    ReDim adblSorted(1 To lNumSorted)
    ReDim m_alNumCuts(1 To m_lNumFeatures)
    ReDim m_adblCuts(1 To m_lNumFeatures, 1 To m_lNumBins - 1)
    ReDim m_alBins(1 To m_lNumFeatures, 1 To lNumSamples)
    m_lMaxBins = 1
    X.CreateAlias X_
    For i = 1 To m_lNumFeatures
        k = 0
        For j = 1 To lNumSamples Step lStride
            k = k + 1
            adblSorted(k) = X_(i, j)
        Next j
        QuickSortDoubles adblSorted, 1, lNumSorted
        m_alNumCuts(i) = BuildCuts(i, adblSorted, lNumSorted)
        m_lMaxBins = MaxLng(m_lMaxBins, m_alNumCuts(i) + 1)
        For j = 1 To lNumSamples
            m_alBins(i, j) = BinOf(i, X_(i, j))
        Next j
        Application.StatusBar = "Binning progress: " & Format(i / m_lNumFeatures, "0.0%")
        DoEvents
    Next i
    X.RemoveAlias X_
End Sub

Private Function BuildCuts(ByVal lFeature As Long, _
                           ByRef adblSorted() As Double, _
                           ByVal lNumSorted As Long) As Long
    Dim k As Long
    Dim p As Long
    Dim lLastCut As Long
    Dim lNumBoundaries As Long
    Dim lNumCuts As Long

    For p = 1 To lNumSorted - 1
        If adblSorted(p) < adblSorted(p + 1) Then
            lNumBoundaries = lNumBoundaries + 1
        End If
    Next p
    If lNumBoundaries <= m_lNumBins - 1 Then
        For p = 1 To lNumSorted - 1
            If adblSorted(p) < adblSorted(p + 1) Then
                lNumCuts = lNumCuts + 1
                m_adblCuts(lFeature, lNumCuts) = (adblSorted(p) + adblSorted(p + 1)) / 2
            End If
        Next p
    Else
        For k = 1 To m_lNumBins - 1
            p = Int(CDbl(k) * lNumSorted / m_lNumBins)
            If p > lLastCut Then
                Do While p < lNumSorted
                    If adblSorted(p) < adblSorted(p + 1) Then Exit Do
                    p = p + 1
                Loop
                If p >= lNumSorted Then Exit For
                lNumCuts = lNumCuts + 1
                m_adblCuts(lFeature, lNumCuts) = (adblSorted(p) + adblSorted(p + 1)) / 2
                lLastCut = p
            End If
        Next k
    End If
    BuildCuts = lNumCuts
End Function

Private Function BinOf(ByVal lFeature As Long, _
                       ByVal dblValue As Double) As Long
    Dim lLo As Long
    Dim lMid As Long
    Dim lHi As Long

    lLo = 1
    lHi = m_alNumCuts(lFeature)
    Do While lLo <= lHi
        lMid = (lLo + lHi) \ 2
        If m_adblCuts(lFeature, lMid) <= dblValue Then
            lLo = lMid + 1
        Else
            lHi = lMid - 1
        End If
    Loop
    BinOf = lLo
End Function

'The bin matrix is NumFeatures * NumSamples * 4 bytes
Private Sub ReleaseTrainingState()
    m_lMaxBins = 0
    Erase m_alNumCuts
    Erase m_adblCuts
    Erase m_alBins
End Sub

Private Sub Clear()
    Set m_oLossFunction = Nothing
    m_dblLearningRate = 0
    m_lMaxDepth = 0
    m_dblLambda = 0
    m_dblGamma = 0
    m_dblMinChildWeight = 0
    m_dblSubsample = 0
    m_dblColsample = 0
    m_lNumBins = 0
    m_lNumFeatures = 0
    m_lNumOutputs = 0
    m_lNumRounds = 0
    Set m_oBaseScore = Nothing
    Erase m_aoTrees
    ReleaseTrainingState
End Sub

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "XGBoost.Save"
    Dim i As Long
    Dim k As Long

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteObject m_oLossFunction
        .WriteDouble m_dblLearningRate
        .WriteLong m_lMaxDepth
        .WriteDouble m_dblLambda
        .WriteDouble m_dblGamma
        .WriteDouble m_dblMinChildWeight
        .WriteDouble m_dblSubsample
        .WriteDouble m_dblColsample
        .WriteLong m_lNumBins
        .WriteLong m_lNumFeatures
        .WriteLong m_lNumOutputs
        .WriteLong m_lNumRounds
        .WriteObject m_oBaseScore
        For i = 1 To m_lNumRounds
            For k = 1 To m_lNumOutputs
                .WriteObject m_aoTrees(k, i)
            Next k
        Next i
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "XGBoost.Load"
    Dim i As Long
    Dim k As Long

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oSerializer
        Set m_oLossFunction = .ReadObject()
        m_dblLearningRate = .ReadDouble()
        m_lMaxDepth = .ReadLong()
        m_dblLambda = .ReadDouble()
        m_dblGamma = .ReadDouble()
        m_dblMinChildWeight = .ReadDouble()
        m_dblSubsample = .ReadDouble()
        m_dblColsample = .ReadDouble()
        m_lNumBins = .ReadLong()
        m_lNumFeatures = .ReadLong()
        m_lNumOutputs = .ReadLong()
        m_lNumRounds = .ReadLong()
        Set m_oBaseScore = .ReadObject()
        If m_lNumOutputs > 0 And m_lNumRounds > 0 Then
            ReDim m_aoTrees(1 To m_lNumOutputs, 1 To m_lNumRounds)
            For i = 1 To m_lNumRounds
                For k = 1 To m_lNumOutputs
                    Set m_aoTrees(k, i) = .ReadObject()
                Next k
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


