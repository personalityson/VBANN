Attribute VB_Name = "FactoryFunctions"
Option Explicit

Public Function Zeros(ByVal vShape As Variant) As Tensor
    Set Zeros = New Tensor
    Zeros.Resize vShape
End Function

Public Function Ones(ByVal vShape As Variant) As Tensor
    Set Ones = Full(vShape, 1)
End Function

Public Function Full(ByVal vShape As Variant, _
                     ByVal dblValue As Double) As Tensor
    Set Full = New Tensor
    With Full
        .Resize vShape
        .Fill dblValue
    End With
End Function

Public Function Uniform(ByVal vShape As Variant, _
                        ByVal dblLow As Double, _
                        ByVal dblHigh As Double) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set Uniform = New Tensor
    With Uniform
        .Resize vShape
        .Ravel.CreateAlias A_
        For i = 1 To .NumElements
            A_(i) = dblLow + (dblHigh - dblLow) * Rnd()
        Next i
        .Ravel.RemoveAlias A_
    End With
End Function

Public Function Normal(ByVal vShape As Variant, _
                       ByVal dblMu As Double, _
                       ByVal dblSigma As Double) As Tensor
    Dim i As Long
    Dim A_() As Double
    
    Set Normal = New Tensor
    With Normal
        .Resize vShape
        .Ravel.CreateAlias A_
        For i = 1 To .NumElements
            A_(i) = dblMu + dblSigma * NormRand()
        Next i
        .Ravel.RemoveAlias A_
    End With
End Function

Public Function GlorotUniform(ByVal vShape As Variant, _
                              ByVal lInputSize As Long, _
                              ByVal lOutputSize As Long, _
                              Optional ByVal dblGain As Double = 1) As Tensor
    Dim dblLimit As Double
    
    dblLimit = dblGain * Sqr(6 / (lInputSize + lOutputSize))
    Set GlorotUniform = Uniform(vShape, -dblLimit, dblLimit)
End Function

Public Function GlorotNormal(ByVal vShape As Variant, _
                             ByVal lInputSize As Long, _
                             ByVal lOutputSize As Long, _
                             Optional ByVal dblGain As Double = 1) As Tensor
    Dim dblSigma As Double

    dblSigma = dblGain * Sqr(2 / (lInputSize + lOutputSize))
    Set GlorotNormal = Normal(vShape, 0, dblSigma)
End Function

Public Function HeUniform(ByVal vShape As Variant, _
                          ByVal lInputSize As Long, _
                          Optional ByVal dblGain As Double = 1) As Tensor
    Dim dblLimit As Double

    dblLimit = dblGain * Sqr(6 / (lInputSize))
    Set HeUniform = Uniform(vShape, -dblLimit, dblLimit)
End Function

Public Function HeNormal(ByVal vShape As Variant, _
                         ByVal lInputSize As Long, _
                         Optional ByVal dblGain As Double = 1) As Tensor
    Dim dblSigma As Double

    dblSigma = dblGain * Sqr(2 / lInputSize)
    Set HeNormal = Normal(vShape, 0, dblSigma)
End Function

Public Function Parameter(ByVal oLearnable As Tensor, _
                          Optional ByVal dblLearningRateFactor As Double = 1, _
                          Optional ByVal dblWeightDecayFactor As Double = 1) As Parameter
    Set Parameter = New Parameter
    Parameter.Init oLearnable, dblLearningRateFactor, dblWeightDecayFactor
End Function

Public Function FullyConnectedLayer(ByVal lInputSize As Long, _
                                    ByVal lOutputSize As Long) As FullyConnectedLayer
    Set FullyConnectedLayer = New FullyConnectedLayer
    FullyConnectedLayer.Init lInputSize, lOutputSize
End Function

Public Function LeakyReLULayer(Optional ByVal dblNegativeSlope As Double = 0.01) As LeakyReLULayer
    Set LeakyReLULayer = New LeakyReLULayer
    LeakyReLULayer.Init dblNegativeSlope
End Function

Public Function SigmoidLayer() As SigmoidLayer
    Set SigmoidLayer = New SigmoidLayer
End Function

Public Function BCELoss() As BCELoss
    Set BCELoss = New BCELoss
End Function

Public Function L1Loss() As L1Loss
    Set L1Loss = New L1Loss
End Function

Public Function L2Loss() As L2Loss
    Set L2Loss = New L2Loss
End Function

Public Function Adam(Optional ByVal dblLearningRate As Double = 0.001, _
                     Optional ByVal dblBeta1 As Double = 0.9, _
                     Optional ByVal dblBeta2 As Double = 0.999, _
                     Optional ByVal dblEpsilon As Double = 0.00000001, _
                     Optional ByVal dblWeightDecay As Double = 0.01) As Adam
    Set Adam = New Adam
    Adam.Init dblLearningRate, dblBeta1, dblBeta2, dblEpsilon, dblWeightDecay
End Function

Public Function SGDM(Optional ByVal dblLearningRate As Double = 0.001, _
                     Optional ByVal dblMomentum As Double = 0.9, _
                     Optional ByVal dblWeightDecay As Double = 0.01) As SGDM
    Set SGDM = New SGDM
    SGDM.Init dblLearningRate, dblMomentum, dblWeightDecay
End Function

Public Function DataLoader(ByVal oDataset As Dataset, _
                           ByVal lBatchSize As Long) As DataLoader
    Set DataLoader = New DataLoader
    DataLoader.Init oDataset, lBatchSize
End Function

Public Function Sequential(ByVal oCriterion As ICriterion, _
                           ByVal oOptimizer As IOptimizer) As Sequential
    Set Sequential = New Sequential
    Sequential.Init oCriterion, oOptimizer
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

Public Function TensorFromRange(ByVal rngRange As Range, _
                                ByVal bTranspose As Boolean) As Tensor
    Set TensorFromRange = New Tensor
    TensorFromRange.FromRange rngRange, bTranspose
End Function

Public Function TensorFromArray(ByRef adblArray() As Double) As Tensor
    Set TensorFromArray = New Tensor
    TensorFromArray.FromArray adblArray
End Function

Public Function ImportDatasetFromWorksheet(ByVal sName As String, _
                                           ByVal lInputSize As Long, _
                                           ByVal lLabelSize As Long, _
                                           Optional ByVal bHasHeaders As Boolean) As Dataset
    Const PROCEDURE_NAME As String = "FactoryFunctions.ImportDatasetFromWorksheet"
    Dim lFirstRow As Long
    Dim lFirstCol As Long
    Dim lNumSamples As Long
    Dim rngInputs As Range
    Dim rngLabels As Range
    Dim wksSource As Worksheet
    Dim oResult As Dataset
    
    If Not WorksheetExists(ThisWorkbook, sName) Then
        Err.Raise 9, PROCEDURE_NAME, "Specified worksheet does not exist."
    End If
    If lInputSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Input size must be greater than 0."
    End If
    If lLabelSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Label size must be greater than 0."
    End If
    Set wksSource = ThisWorkbook.Sheets(sName)
    lFirstRow = GetFirstRow(wksSource) + IIf(bHasHeaders, 1, 0)
    lFirstCol = GetFirstColumn(wksSource)
    lNumSamples = GetLastRow(wksSource) - lFirstRow + 1
    Set oResult = New Dataset
    If lNumSamples > 0 Then
        Set rngInputs = wksSource.Cells(lFirstRow, lFirstCol).Resize(lNumSamples, lInputSize)
        Set rngLabels = wksSource.Cells(lFirstRow, lFirstCol + lInputSize).Resize(lNumSamples, lLabelSize)
        With oResult
            .Add TensorFromRange(rngInputs, True)
            .Add TensorFromRange(rngLabels, True)
        End With
    Else
        With oResult
            .Add Zeros(Array(lInputSize, 0))
            .Add Zeros(Array(lLabelSize, 0))
        End With
    End If
    Set ImportDatasetFromWorksheet = oResult
End Function

Public Function ImportDatasetFromCsv(ByVal strPath As String, _
                                     ByVal lInputSize As Long, _
                                     ByVal lLabelSize As Long, _
                                     Optional ByVal bHasHeaders As Boolean) As Dataset
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
    Dim oResult As Dataset
    
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
        Loop
        .Close
    End With
    Set oResult = New Dataset
    If lNumRows > 0 Then
        ReDim Preserve adblInputs(1 To lInputSize, 1 To lNumRows)
        ReDim Preserve adblLabels(1 To lLabelSize, 1 To lNumRows)
        With oResult
            .Add TensorFromArray(adblInputs)
            .Add TensorFromArray(adblLabels)
        End With
    Else
        With oResult
            .Add Zeros(Array(lInputSize, 0))
            .Add Zeros(Array(lLabelSize, 0))
        End With
    End If
    Set ImportDatasetFromCsv = oResult
End Function

Public Sub LogToWorksheet(ByVal sName As String, _
                          ParamArray avArgs() As Variant)
    Dim wksLog As Worksheet
    Dim bIsWorksheetNew As Boolean
    Dim lLastRow As Long
    Dim i As Long
    Dim vHeader As Variant
    Dim vValue As Variant
    Dim lHeaderCol As Long
    
    Set wksLog = CreateWorksheet(ThisWorkbook, sName, False, bIsWorksheetNew)
    If bIsWorksheetNew Then
        lLastRow = 1
    Else
        lLastRow = GetLastRow(wksLog)
    End If
    With wksLog
        On Error Resume Next
        For i = 0 To UBound(avArgs) - 1 Step 2
            vHeader = avArgs(i)
            vValue = avArgs(i + 1)
            lHeaderCol = 0
            lHeaderCol = WorksheetFunction.Match(vHeader, .Rows(1), 0)
            If lHeaderCol = 0 Then
                lHeaderCol = GetLastColumn(wksLog) + 1
                .Cells(1, lHeaderCol) = vHeader
            End If
            .Cells(lLastRow + 1, lHeaderCol) = vValue
            Application.GoTo .Cells(lLastRow + 1, lHeaderCol)
            DoEvents
        Next i
        On Error GoTo 0
        If bIsWorksheetNew Then
            .Activate
            .Cells(2, 1).Select
            ActiveWindow.FreezePanes = True
        End If
    End With

End Sub