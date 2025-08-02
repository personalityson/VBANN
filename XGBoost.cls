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

'Dependencies
Private m_oCriterion As ICriterion

'Hyperparameters
Private m_dblLearningRate As Double
Private m_lMaxDepth As Long
Private m_dblLambda As Double
Private m_dblGamma As Double
Private m_dblMinChildWeight As Double

'State
Private m_lNumFeatures As Long
Private m_dblBaseValue As Double
Private m_lNumTrees As Long
Private m_aoTrees() As XGBoostNode

Public Sub Init(ByVal oCriterion As ICriterion, _
                Optional ByVal dblLearningRate As Double = 0.1, _
                Optional ByVal lMaxDepth As Long = 6, _
                Optional ByVal dblLambda As Double = 1, _
                Optional ByVal dblGamma As Double = 0, _
                Optional ByVal dblMinChildWeight As Double = 1)
    Const PROCEDURE_NAME As String = "XGBoost.Init"
    
    If oCriterion Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ICriterion object is required."
    End If
    If dblLearningRate <= 0 Or dblLearningRate > 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Learning rate must be > 0 and <= 1."
    End If
    If lMaxDepth < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Max depth must be >= 1."
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
    Clear
    Set m_oCriterion = oCriterion
    m_dblLearningRate = dblLearningRate
    m_lMaxDepth = lMaxDepth
    m_dblLambda = dblLambda
    m_dblGamma = dblGamma
    m_dblMinChildWeight = dblMinChildWeight
End Sub

Public Sub Fit(ByVal X As Tensor, _
               ByVal T As Tensor, _
               ByVal lNumTrees As Long)
    Const PROCEDURE_NAME As String = "XGBoost.Fit"
    Dim i As Long
    Dim Y As Tensor

    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If X.NumDimensions <> 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor must have 2 dimensions."
    End If
    If T.NumDimensions <> 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Target tensor must have 1 dimension."
    End If
    If X.Size(2) <> T.Size(1) Then
        Err.Raise 5, PROCEDURE_NAME, "Number of samples in input and target tensors does not match."
    End If
    If lNumTrees < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of trees must be >= 1."
    End If
    m_lNumFeatures = X.Size(1)
    m_dblBaseValue = T.Mean
    m_lNumTrees = lNumTrees
    ReDim m_aoTrees(1 To m_lNumTrees)
    Set Y = Full(T.Shape, m_dblBaseValue)
    For i = 1 To m_lNumTrees
        Set m_aoTrees(i) = FitOne(X, Y, T)
        Application.StatusBar = "Fit progress: " & Format(i / lNumTrees, "0.0%")
        DoEvents
    Next i
End Sub

Public Function Loss(ByVal X As Tensor, _
                     ByVal T As Tensor) As Double
    Const PROCEDURE_NAME As String = "XGBoost.Loss"
    Dim Y As Tensor
    
    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If T Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If X.NumDimensions <> 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor must have 2 dimensions."
    End If
    If T.NumDimensions <> 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Target tensor must have 1 dimension."
    End If
    If X.Size(1) <> m_lNumFeatures Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor is not compatible with the trained model."
    End If
    If X.Size(2) <> T.Size(1) Then
        Err.Raise 5, PROCEDURE_NAME, "Number of samples in input and target tensors does not match."
    End If
    Set Y = Me.Predict(X)
    Loss = m_oCriterion.Loss(Y, T).Mean
End Function

Public Function Predict(ByVal X As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "XGBoost.Predict"
    Dim lNumSamples  As Long
    Dim i As Long
    Dim Y As Tensor
    Dim Y_hat As Tensor
    
    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If X.NumDimensions <> 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor must have 2 dimensions."
    End If
    If X.Size(1) <> m_lNumFeatures Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor is not compatible with the trained model."
    End If
    lNumSamples = X.Size(2)
    Set Y = Full(Array(lNumSamples), m_dblBaseValue)
    For i = 1 To m_lNumTrees
        Set Y_hat = PredictOne(m_aoTrees(i), X)
        VecLinComb_I 1, Y, m_dblLearningRate, Y_hat
        Application.StatusBar = "Predict progress: " & Format(i / m_lNumTrees, "0.0%")
        DoEvents
    Next i
    Set Predict = Y
End Function

Private Sub Clear()
    Set m_oCriterion = Nothing
    m_dblLearningRate = 0
    m_lMaxDepth = 0
    m_dblLambda = 0
    m_dblGamma = 0
    m_dblMinChildWeight = 0
    m_lNumFeatures = 0
    m_dblBaseValue = 0
    m_lNumTrees = 0
    Erase m_aoTrees
End Sub

Private Function BuildNode(ByRef X_() As Double, _
                           ByRef G_() As Double, _
                           ByRef H_() As Double, _
                           ByRef alSampleIndices() As Long, _
                           ByVal lDepth As Long) As XGBoostNode
    Dim i As Long
    Dim j As Long
    Dim s_l As Long
    Dim s_r As Long
    Dim lNumSamples As Long
    Dim dblSumGT As Double
    Dim dblSumHT As Double
    Dim dblSumGL As Double
    Dim dblSumGR As Double
    Dim dblSumHL As Double
    Dim dblSumHR As Double
    Dim dblWeight As Double
    Dim dblGain As Double
    Dim dblBestGain As Double
    Dim dblBestThreshold As Double
    Dim lBestFeature As Long
    Dim alSampleIndicesSorted() As Long
    Dim alBestSampleIndicesL() As Long
    Dim alBestSampleIndicesR() As Long
    Dim oChildL As XGBoostNode
    Dim oChildR As XGBoostNode
    Dim oResult As XGBoostNode
    
    lNumSamples = UBound(alSampleIndices)
    For j = 1 To lNumSamples
        dblSumGT = dblSumGT + G_(alSampleIndices(j))
        dblSumHT = dblSumHT + H_(alSampleIndices(j))
    Next j
    If Abs(dblSumHT + m_dblLambda) < DOUBLE_EPSILON Then
        dblWeight = 0
    Else
        dblWeight = -dblSumGT / (dblSumHT + m_dblLambda)
    End If
    If lDepth >= m_lMaxDepth Or dblSumHT < m_dblMinChildWeight Or lNumSamples < 2 Then
        Set oResult = New XGBoostNode
        oResult.MakeLeaf dblWeight
        Set BuildNode = oResult
        Exit Function
    End If
    dblBestGain = -DOUBLE_MAX_ABS
    For i = 1 To m_lNumFeatures
        alSampleIndicesSorted = alSampleIndices
        QuickSortByFeature X_, alSampleIndicesSorted, i, 1, lNumSamples
        dblSumGL = 0
        dblSumHL = 0
        dblSumGR = dblSumGT
        dblSumHR = dblSumHT
        For j = 1 To lNumSamples - 1
            s_l = alSampleIndicesSorted(j)
            s_r = alSampleIndicesSorted(j + 1)
            dblSumGL = dblSumGL + G_(s_l)
            dblSumHL = dblSumHL + H_(s_l)
            dblSumGR = dblSumGR - G_(s_l)
            dblSumHR = dblSumHR - H_(s_l)
            If dblSumHL >= m_dblMinChildWeight And dblSumHR >= m_dblMinChildWeight Then
                If X_(i, s_l) <> X_(i, s_r) Then
                    dblGain = 0.5 * ((dblSumGL * dblSumGL) / (dblSumHL + m_dblLambda) + _
                                     (dblSumGR * dblSumGR) / (dblSumHR + m_dblLambda) - _
                                     (dblSumGT * dblSumGT) / (dblSumHT + m_dblLambda)) - m_dblGamma
                    If dblGain > dblBestGain Then
                        dblBestGain = dblGain
                        dblBestThreshold = (X_(i, s_l) + X_(i, s_r)) / 2
                        lBestFeature = i
                        ReDim alBestSampleIndicesL(1 To j)
                        ReDim alBestSampleIndicesR(1 To lNumSamples - j)
                        CopyMemory alBestSampleIndicesL(1), alSampleIndicesSorted(1), j * SIZEOF_LONG
                        CopyMemory alBestSampleIndicesR(1), alSampleIndicesSorted(j + 1), (lNumSamples - j) * SIZEOF_LONG
                    End If
                End If
            End If
        Next j
    Next i
    If dblBestGain > 0 Then
        Set oChildL = BuildNode(X_, G_, H_, alBestSampleIndicesL, lDepth + 1)
        Set oChildR = BuildNode(X_, G_, H_, alBestSampleIndicesR, lDepth + 1)
        Set oResult = New XGBoostNode
        oResult.Split lBestFeature, dblBestThreshold, oChildL, oChildR
    Else
        Set oResult = New XGBoostNode
        oResult.MakeLeaf dblWeight
    End If
    Set BuildNode = oResult
End Function

Private Function FitOne(ByVal X As Tensor, _
                        ByVal Y As Tensor, _
                        ByVal T As Tensor) As XGBoostNode
    Dim lNumSamples As Long
    Dim alSampleIndices() As Long
    Dim X_() As Double
    Dim G_() As Double
    Dim H_() As Double
    Dim G As Tensor
    Dim H As Tensor
    Dim Y_hat As Tensor
    Dim oResult As XGBoostNode
    
    lNumSamples = X.Size(2)
    alSampleIndices = GetIdentityPermutationArray(lNumSamples)
    Set G = m_oCriterion.Gradient(Y, T)
    Set H = m_oCriterion.Hessian(Y, T)
    X.CreateAlias X_
    G.CreateAlias G_
    H.CreateAlias H_
    Set oResult = BuildNode(X_, G_, H_, alSampleIndices, 1)
    X.RemoveAlias X_
    G.RemoveAlias G_
    H.RemoveAlias H_
    Set Y_hat = PredictOne(oResult, X)
    VecLinComb_I 1, Y, m_dblLearningRate, Y_hat
    Set FitOne = oResult
End Function

Private Function PredictOne(ByVal oTree As XGBoostNode, _
                            ByVal X As Tensor) As Tensor
    Dim j As Long
    Dim lNumSamples As Long
    Dim X_() As Double
    Dim Y_() As Double
    Dim Y As Tensor
    
    lNumSamples = X.Size(2)
    Set Y = Zeros(Array(lNumSamples))
    X.CreateAlias X_
    Y.CreateAlias Y_
    For j = 1 To lNumSamples
        Y_(j) = oTree.PredictRow(X_, j)
    Next j
    X.RemoveAlias X_
    Y.RemoveAlias Y_
    Set PredictOne = Y
End Function

Private Sub QuickSortByFeature(ByRef X_() As Double, _
                               ByRef alSampleIndices() As Long, _
                               ByVal lFeatureIndex As Long, _
                               ByVal lStart As Long, _
                               ByVal lEnd As Long)
    Dim i As Long
    Dim j As Long
    Dim lTemp As Long
    Dim lPivotIndex As Long
    Dim dblPivotValue As Double
    
    If lStart >= lEnd Then
        Exit Sub
    End If
    lPivotIndex = alSampleIndices(lStart + (lEnd - lStart) \ 2)
    dblPivotValue = X_(lFeatureIndex, lPivotIndex)
    i = lStart
    j = lEnd
    Do
        Do While X_(lFeatureIndex, alSampleIndices(i)) < dblPivotValue
            i = i + 1
        Loop
        Do While X_(lFeatureIndex, alSampleIndices(j)) > dblPivotValue
            j = j - 1
        Loop
        If i <= j Then
            lTemp = alSampleIndices(i)
            alSampleIndices(i) = alSampleIndices(j)
            alSampleIndices(j) = lTemp
            i = i + 1
            j = j - 1
        End If
    Loop While i <= j
    If lStart < j Then
        QuickSortByFeature X_, alSampleIndices, lFeatureIndex, lStart, j
    End If
    If i < lEnd Then
        QuickSortByFeature X_, alSampleIndices, lFeatureIndex, i, lEnd
    End If
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "XGBoost.Save"
    Dim i As Long

    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteObject m_oCriterion
        .WriteDouble m_dblLearningRate
        .WriteLong m_lMaxDepth
        .WriteDouble m_dblLambda
        .WriteDouble m_dblGamma
        .WriteDouble m_dblMinChildWeight
        .WriteLong m_lNumFeatures
        .WriteDouble m_dblBaseValue
        .WriteLong m_lNumTrees
        For i = 1 To m_lNumTrees
            .WriteObject m_aoTrees(i)
        Next i
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "XGBoost.Load"
    Dim i As Long

    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oWorksheet
        Set m_oCriterion = .ReadObject()
        m_dblLearningRate = .ReadDouble()
        m_lMaxDepth = .ReadLong()
        m_dblLambda = .ReadDouble()
        m_dblGamma = .ReadDouble()
        m_dblMinChildWeight = .ReadDouble()
        m_lNumFeatures = .ReadLong()
        m_dblBaseValue = .ReadDouble()
        m_lNumTrees = .ReadLong()
        If m_lNumTrees > 0 Then
            ReDim m_aoTrees(1 To m_lNumTrees)
            For i = 1 To m_lNumTrees
                Set m_aoTrees(i) = .ReadObject()
            Next i
        End If
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
