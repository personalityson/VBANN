Attribute VB_Name = "FactoryFunctions"
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
                           ByVal lBatchSize As Long) As DataLoader
    Set DataLoader = New DataLoader
    DataLoader.Init oDataset, lBatchSize
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

Public Function InputNormalizationLayer(ByVal oTrainingSet As DataLoader, _
                                        Optional ByVal dblEpsilon As Double = 0.00001) As InputNormalizationLayer
    Set InputNormalizationLayer = New InputNormalizationLayer
    InputNormalizationLayer.Init oTrainingSet, dblEpsilon
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

Public Function TanhLayer() As TanhLayer
    Set TanhLayer = New TanhLayer
End Function

Public Function TensorDataset() As TensorDataset
    Set TensorDataset = New TensorDataset
    TensorDataset.Init
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

Public Function ImportDatasetFromWorksheet(ByVal sName As String, _
                                           ByVal lInputSize As Long, _
                                           ByVal lLabelSize As Long, _
                                           Optional ByVal bHasHeaders As Boolean) As TensorDataset
    Const PROCEDURE_NAME As String = "FactoryFunctions.ImportDatasetFromWorksheet"
    Dim lFirstRow As Long
    Dim lFirstCol As Long
    Dim lNumSamples As Long
    Dim oInputs As Range
    Dim oLabels As Range
    Dim oSource As Worksheet
    Dim oResult As TensorDataset
    
    If Not WorksheetExists(ThisWorkbook, sName) Then
        Err.Raise 9, PROCEDURE_NAME, "Specified worksheet does not exist."
    End If
    If lInputSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Input size must be greater than 0."
    End If
    If lLabelSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Label size must be greater than 0."
    End If
    Set oSource = ThisWorkbook.Sheets(sName)
    lFirstRow = GetFirstRow(oSource) + IIf(bHasHeaders, 1, 0)
    lFirstCol = GetFirstColumn(oSource)
    lNumSamples = GetLastRow(oSource) - lFirstRow + 1
    Set oResult = New TensorDataset
    oResult.Init
    If lNumSamples > 0 Then
        Set oInputs = oSource.Cells(lFirstRow, lFirstCol).Resize(lNumSamples, lInputSize)
        Set oLabels = oSource.Cells(lFirstRow, lFirstCol + lInputSize).Resize(lNumSamples, lLabelSize)
        With oResult
            .Add "Input", TensorFromRange(oInputs, True)
            .Add "Label", TensorFromRange(oLabels, True)
        End With
    Else
        With oResult
            .Add "Input", Zeros(Array(lInputSize, 0))
            .Add "Label", Zeros(Array(lLabelSize, 0))
        End With
    End If
    Set ImportDatasetFromWorksheet = oResult
End Function

Public Function ImportDatasetFromCsv(ByVal strPath As String, _
                                     ByVal lInputSize As Long, _
                                     ByVal lLabelSize As Long, _
                                     Optional ByVal bHasHeaders As Boolean) As TensorDataset
    Const PROCEDURE_NAME As String = "FactoryFunctions.ImportDatasetFromCsv"
    Const CHUNK_SIZE As Long = 10000
    Const ForReading As Long = 1
    Dim lNumRows As Long
    Dim lNumAllocatedRows As Long
    Dim lNumFields As Long
    Dim i As Long
    Dim vFields As Variant
    Dim dblValue As Double
    Dim adblInputs() As Double
    Dim adblLabels() As Double
    Dim oResult As TensorDataset
    
    If Not Fso.FileExists(strPath) Then
        Err.Raise 9, PROCEDURE_NAME, "Specified file does not exist."
    End If
    If lInputSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Input size must be greater than 0."
    End If
    If lLabelSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Label size must be greater than 0."
    End If
    With Fso.OpenTextFile(strPath, ForReading)
        If Not .AtEndOfStream And bHasHeaders Then
            .SkipLine
        End If
        Do While Not .AtEndOfStream
            lNumRows = lNumRows + 1
            If lNumRows > lNumAllocatedRows Then
                lNumAllocatedRows = lNumAllocatedRows + CHUNK_SIZE
                ReDim Preserve adblInputs(1 To lInputSize, 1 To lNumAllocatedRows)
                ReDim Preserve adblLabels(1 To lLabelSize, 1 To lNumAllocatedRows)
            End If
            'vFields = Split(.ReadLine, ",")
            vFields = Split(.ReadLine, ";")
            lNumFields = UBound(vFields) + 1
            If lNumFields < lInputSize + lLabelSize Then
                Err.Raise 5, PROCEDURE_NAME, "Number of fields must be greater than or equal to the sum of input size and label size."
            End If
            For i = 1 To lNumFields
                'dblValue = Val(vFields(i - 1))
                dblValue = CDbl(vFields(i - 1))
                If i <= lInputSize Then
                    adblInputs(i, lNumRows) = dblValue
                ElseIf i <= lInputSize + lLabelSize Then
                    adblLabels(i - lInputSize, lNumRows) = dblValue
                End If
            Next i
            If lNumRows Mod 100 = 0 Then
                Application.StatusBar = "ImportDatasetFromCsv progress: " & lNumRows
                DoEvents
            End If
        Loop
        .Close
    End With
    Set oResult = New TensorDataset
    oResult.Init
    If lNumRows > 0 Then
        ReDim Preserve adblInputs(1 To lInputSize, 1 To lNumRows)
        ReDim Preserve adblLabels(1 To lLabelSize, 1 To lNumRows)
        With oResult
            .Add "Input", TensorFromArray(adblInputs)
            .Add "Label", TensorFromArray(adblLabels)
        End With
    Else
        With oResult
            .Add "Input", Zeros(Array(lInputSize, 0))
            .Add "Label", Zeros(Array(lLabelSize, 0))
        End With
    End If
    Set ImportDatasetFromCsv = oResult
End Function
