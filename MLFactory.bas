Attribute VB_Name = "MLFactory"
Option Explicit

Public Function Adam(Optional ByVal dblLearningRate As Double = 0.001, _
                     Optional ByVal dblBeta1 As Double = 0.9, _
                     Optional ByVal dblBeta2 As Double = 0.999, _
                     Optional ByVal dblEpsilon As Double = 0.00000001, _
                     Optional ByVal dblWeightDecay As Double = 0.01) As Adam
    Set Adam = New Adam
    Adam.Init dblLearningRate, dblBeta1, dblBeta2, dblEpsilon, dblWeightDecay
End Function

Public Function BCELoss() As BCELoss
    Set BCELoss = New BCELoss
End Function

Public Function CCELoss() As CCELoss
    Set CCELoss = New CCELoss
End Function

Public Function DataLoader(ByVal oDataset As IDataset, _
                           ByVal lBatchSize As Long, _
                           Optional ByVal bDropRemainder As Boolean) As DataLoader
    Set DataLoader = New DataLoader
    DataLoader.Init oDataset, lBatchSize, bDropRemainder
End Function

Public Function DropoutLayer(Optional ByVal dblDropoutRate As Double = 0.5) As DropoutLayer
    Set DropoutLayer = New DropoutLayer
    DropoutLayer.Init dblDropoutRate
End Function

Public Function FullyConnectedLayer(ByVal lInputSize As Long, _
                                    ByVal lOutputSize As Long) As FullyConnectedLayer
    Set FullyConnectedLayer = New FullyConnectedLayer
    FullyConnectedLayer.Init lInputSize, lOutputSize
End Function

Public Function InputNormalizationLayer(ByVal oTrainingLoader As DataLoader) As InputNormalizationLayer
    Set InputNormalizationLayer = New InputNormalizationLayer
    InputNormalizationLayer.Init oTrainingLoader
End Function

Public Function L1Loss() As L1Loss
    Set L1Loss = New L1Loss
End Function

Public Function L2Loss() As L2Loss
    Set L2Loss = New L2Loss
End Function

Public Function LeakyReLULayer(Optional ByVal dblNegativeSlope As Double = 0.01) As LeakyReLULayer
    Set LeakyReLULayer = New LeakyReLULayer
    LeakyReLULayer.Init dblNegativeSlope
End Function

Public Function Parameter(ByVal oVariable As Tensor, _
                          Optional ByVal dblLearningRateScale As Double = 1, _
                          Optional ByVal dblWeightDecayScale As Double = 1) As Parameter
    Set Parameter = New Parameter
    Parameter.Init oVariable, dblLearningRateScale, dblWeightDecayScale
End Function

Public Function Sequential(ByVal oCriterion As ICriterion, _
                           ByVal oOptimizer As IOptimizer) As Sequential
    Set Sequential = New Sequential
    Sequential.Init oCriterion, oOptimizer
End Function

Public Function SGDM(Optional ByVal dblLearningRate As Double = 0.001, _
                     Optional ByVal dblMomentum As Double = 0.9, _
                     Optional ByVal dblWeightDecay As Double = 0.01) As SGDM
    Set SGDM = New SGDM
    SGDM.Init dblLearningRate, dblMomentum, dblWeightDecay
End Function

Public Function SigmoidLayer() As SigmoidLayer
    Set SigmoidLayer = New SigmoidLayer
End Function

Public Function SoftmaxLayer() As SoftmaxLayer
    Set SoftmaxLayer = New SoftmaxLayer
End Function

Public Function SubsetDataset(ByVal oDataset As IDataset, _
                              ByVal vIndices As Variant) As SubsetDataset
    Set SubsetDataset = New SubsetDataset
    SubsetDataset.Init oDataset, vIndices
End Function

Public Function TanhLayer() As TanhLayer
    Set TanhLayer = New TanhLayer
End Function

Public Function TensorDataset(ByVal vTensors As Variant) As TensorDataset
    Set TensorDataset = New TensorDataset
    TensorDataset.Init vTensors
End Function

Public Sub Serialize(ByVal sName As String, _
                     ByVal oObject As ISerializable)
    With New Serializer
        .Init sName, True
        .WriteObject oObject
    End With
End Sub

Public Function Unserialize(ByVal sName As String) As ISerializable
    With New Serializer
        .Init sName, False
        Set Unserialize = .ReadObject()
    End With
End Function

Public Function ImportDatasetFromWorksheet(ByVal oWorkbook As Workbook, _
                                           ByVal sName As String, _
                                           ByVal vSegmentSizes As Variant, _
                                           Optional ByVal bHasHeaders As Boolean, _
                                           Optional ByVal bSqueeze As Boolean) As TensorDataset
    Const PROCEDURE_NAME As String = "MLFactory.ImportDatasetFromWorksheet"
    Dim i As Long
    Dim lNumSegments As Long
    Dim alSegmentSizes() As Long
    Dim lFirstRow As Long
    Dim lFirstCol As Long
    Dim lNumSamples As Long
    Dim X As Tensor
    Dim alTensors() As Tensor
    Dim oSource As Worksheet
    Dim oResult As TensorDataset
    
    If oWorkbook Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Workbook object is required."
    End If
    If Not WorksheetExists(oWorkbook, sName) Then
        Err.Raise 9, PROCEDURE_NAME, "Specified worksheet does not exist."
    End If
    ParseVariantToLongArray vSegmentSizes, lNumSegments, alSegmentSizes
    For i = 1 To lNumSegments
        If alSegmentSizes(i) < 1 Then
            Err.Raise 5, PROCEDURE_NAME, "Segment size must be >= 1."
        End If
    Next i
    Set oSource = ThisWorkbook.Sheets(sName)
    lFirstRow = GetFirstRow(oSource) + IIf(bHasHeaders, 1, 0)
    lFirstCol = GetFirstColumn(oSource)
    lNumSamples = GetLastRow(oSource) - lFirstRow + 1
    ReDim alTensors(1 To lNumSegments)
    Set oResult = New TensorDataset
    For i = 1 To lNumSegments
        If lNumSamples > 0 Then
            Set X = TensorFromRange(oSource.Cells(lFirstRow, lFirstCol).Resize(lNumSamples, alSegmentSizes(i)), True)
        Else
            Set X = Zeros(Array(alSegmentSizes(i), 0))
        End If
        If bSqueeze Then
            Set X = X.Squeeze
        End If
        Set alTensors(i) = X
        lFirstCol = lFirstCol + alSegmentSizes(i)
    Next i
    oResult.Init alTensors
    Set ImportDatasetFromWorksheet = oResult
End Function

Public Sub RandomSplit(ByVal oDataset As IDataset, _
                       ByVal dblAt As Double, _
                       ByRef A As SubsetDataset, _
                       ByRef B As SubsetDataset)
    Const PROCEDURE_NAME As String = "MLFactory.RandomSplit"
    Dim lSizeA As Long
    Dim lSizeB As Long
    Dim alFullIndices() As Long
    Dim alIndicesA() As Long
    Dim alIndicesB() As Long

    If oDataset Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IDataset object is required."
    End If
    If dblAt < 0 Or dblAt > 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Fraction must be >= 0 and <= 1."
    End If
    alFullIndices = GetRandomPermutationArray(oDataset.NumSamples)
    lSizeA = CLng(dblAt * oDataset.NumSamples + 0.5)
    lSizeB = oDataset.NumSamples - lSizeA
    If lSizeA > 0 Then
        ReDim alIndicesA(1 To lSizeA)
        CopyMemory alIndicesA(1), alFullIndices(1), lSizeA * SIZEOF_LONG
    End If
    If lSizeB > 0 Then
        ReDim alIndicesB(1 To lSizeB)
        CopyMemory alIndicesB(1), alFullIndices(lSizeA + 1), lSizeB * SIZEOF_LONG
    End If
    Set A = oDataset.Subset(alIndicesA)
    Set B = oDataset.Subset(alIndicesB)
End Sub
