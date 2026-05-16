Attribute VB_Name = "MLFactory"
Option Explicit

Public Function AdamW(Optional ByVal dblLearningRate As Double = 0.001, _
                      Optional ByVal dblBeta1 As Double = 0.9, _
                      Optional ByVal dblBeta2 As Double = 0.999, _
                      Optional ByVal dblEpsilon As Double = 0.00000001, _
                      Optional ByVal dblWeightDecay As Double = 0.01, _
                      Optional ByVal dblGradientThreshold As Double = DOUBLE_MAX_ABS) As AdamW
    Set AdamW = New AdamW
    AdamW.Init dblLearningRate, dblBeta1, dblBeta2, dblEpsilon, dblWeightDecay, dblGradientThreshold
End Function

Public Function BCELoss() As BCELoss
    Set BCELoss = New BCELoss
    BCELoss.Init
End Function

Public Function CCELoss() As CCELoss
    Set CCELoss = New CCELoss
    CCELoss.Init
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
    L1Loss.Init
End Function

Public Function L2Loss() As L2Loss
    Set L2Loss = New L2Loss
    L2Loss.Init
End Function

Public Function LeakyReLULayer(Optional ByVal dblNegativeSlope As Double = 0.01) As LeakyReLULayer
    Set LeakyReLULayer = New LeakyReLULayer
    LeakyReLULayer.Init dblNegativeSlope
End Function

Public Function Parameter(ByVal oTensor As Tensor, _
                          Optional ByVal dblLearningRateScale As Double = 1, _
                          Optional ByVal dblWeightDecayScale As Double = 1, _
                          Optional ByVal bUseGradientClipping As Boolean = True) As Parameter
    Set Parameter = New Parameter
    Parameter.Init oTensor, dblLearningRateScale, dblWeightDecayScale, bUseGradientClipping
End Function

Public Function Sequential(ByVal oCriterion As ICriterion, _
                           ByVal oOptimizer As IOptimizer) As Sequential
    Set Sequential = New Sequential
    Sequential.Init oCriterion, oOptimizer
End Function

Public Function SGDW(Optional ByVal dblLearningRate As Double = 0.01, _
                     Optional ByVal dblMomentum As Double = 0.9, _
                     Optional ByVal dblWeightDecay As Double = 0.0001, _
                     Optional ByVal dblGradientThreshold As Double = DOUBLE_MAX_ABS) As SGDW
    Set SGDW = New SGDW
    SGDW.Init dblLearningRate, dblMomentum, dblWeightDecay, dblGradientThreshold
End Function

Public Function SigmoidLayer() As SigmoidLayer
    Set SigmoidLayer = New SigmoidLayer
    SigmoidLayer.Init
End Function

Public Function SoftmaxLayer() As SoftmaxLayer
    Set SoftmaxLayer = New SoftmaxLayer
    SoftmaxLayer.Init
End Function

Public Function SubsetDataset(ByVal oDataset As IDataset, _
                              ByVal vIndices As Variant) As SubsetDataset
    Set SubsetDataset = New SubsetDataset
    SubsetDataset.Init oDataset, vIndices
End Function

Public Function TanhLayer() As TanhLayer
    Set TanhLayer = New TanhLayer
    TanhLayer.Init
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

Public Function Deserialize(ByVal sName As String) As ISerializable
    With New Serializer
        .Init sName, False
        Set Deserialize = .ReadObject()
    End With
End Function

Public Function ImportDatasetFromWorksheet(ByVal oWorkbook As Workbook, _
                                           ByVal sName As String, _
                                           ByVal vSegmentSizes As Variant, _
                                           Optional ByVal bHasHeaders As Boolean) As TensorDataset
    Const PROCEDURE_NAME As String = "MLFactory.ImportDatasetFromWorksheet"
    Dim i As Long
    Dim lNumSegments As Long
    Dim alSegmentSizes() As Long
    Dim lFirstRow As Long
    Dim lFirstCol As Long
    Dim lNumSamples As Long
    Dim aoTensors() As Tensor
    Dim oSource As Worksheet
    Dim oResult As TensorDataset

    If oWorkbook Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Workbook object is required."
    End If
    If Not WorksheetExists(oWorkbook, sName) Then
        Err.Raise 9, PROCEDURE_NAME, "Specified worksheet does not exist."
    End If
    ParseVariantToLongArray vSegmentSizes, lNumSegments, alSegmentSizes
    If lNumSegments < 1 Then
        Err.Raise 9, PROCEDURE_NAME, "Dataset must have at least one segment."
    End If
    For i = 1 To lNumSegments
        If alSegmentSizes(i) < 1 Then
            Err.Raise 5, PROCEDURE_NAME, "Segment size must be >= 1."
        End If
    Next i
    Set oSource = oWorkbook.Sheets(sName)
    lFirstRow = GetFirstRow(oSource) + IIf(bHasHeaders, 1, 0)
    lFirstCol = GetFirstColumn(oSource)
    
    'lNumSamples = GetLastRow(oSource) - lFirstRow + 1
    lNumSamples = GetLastRow(oSource, 1) - lFirstRow + 1
    
    ReDim aoTensors(1 To lNumSegments)
    Set oResult = New TensorDataset
    For i = 1 To lNumSegments
        If lNumSamples > 0 Then
            Set aoTensors(i) = TensorFromRange(oSource.Cells(lFirstRow, lFirstCol).Resize(lNumSamples, alSegmentSizes(i)), True)
        Else
            Set aoTensors(i) = Zeros(Array(alSegmentSizes(i), 0))
        End If
        lFirstCol = lFirstCol + alSegmentSizes(i)
    Next i
    oResult.Init aoTensors
    Set ImportDatasetFromWorksheet = oResult
End Function

Public Function ImportDatasetFromCsv(ByVal sPath As String, _
                                     ByVal vSegmentSizes As Variant, _
                                     Optional ByVal bHasHeaders As Boolean, _
                                     Optional ByVal sDelimiter As String = ",", _
                                     Optional ByVal sDecimalSeparator As String = ".") As TensorDataset
    Const PROCEDURE_NAME As String = "MLFactory.ImportDatasetFromCsv"
    Const ForReading As Long = 1
    Dim bDecimalComma As Boolean
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim lNumSegments As Long
    Dim alSegmentSizes() As Long
    Dim lNumColumns As Long
    Dim lNumSamples As Long
    Dim lNumFields As Long
    Dim lOffset As Long
    Dim sLine As String
    Dim asFields() As String
    Dim adblRow() As Double
    Dim aoTensors() As Tensor
    Dim oResult As TensorDataset

    If Not Fso.FileExists(sPath) Then
        Err.Raise 53, PROCEDURE_NAME, "File not found."
    End If
    ParseVariantToLongArray vSegmentSizes, lNumSegments, alSegmentSizes
    If lNumSegments < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Dataset must have at least one segment."
    End If
    For i = 1 To lNumSegments
        If alSegmentSizes(i) < 1 Then
            Err.Raise 5, PROCEDURE_NAME, "Segment size must be >= 1."
        End If
        lNumColumns = lNumColumns + alSegmentSizes(i)
    Next i
    If sDelimiter = "" Then
        Err.Raise 5, PROCEDURE_NAME, "Delimiter cannot be empty."
    End If
    Select Case sDecimalSeparator
        Case "."
            bDecimalComma = False
        Case ","
            bDecimalComma = True
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Decimal separator must be either '.' or ','."
    End Select
    If sDelimiter = sDecimalSeparator Then
        Err.Raise 5, PROCEDURE_NAME, "Delimiter and decimal separator must differ."
    End If
    With Fso.OpenTextFile(sPath, ForReading)
        If bHasHeaders And Not .AtEndOfStream Then
            .SkipLine
        End If
        Do While Not .AtEndOfStream
            .SkipLine
            lNumSamples = lNumSamples + 1
        Loop
        .Close
    End With
    ReDim aoTensors(1 To lNumSegments)
    For i = 1 To lNumSegments
        Set aoTensors(i) = Zeros(Array(alSegmentSizes(i), lNumSamples))
    Next i
    ReDim adblRow(1 To lNumColumns)
    With Fso.OpenTextFile(sPath, ForReading)
        If bHasHeaders And Not .AtEndOfStream Then
            .SkipLine
        End If
        For j = 1 To lNumSamples
            sLine = .ReadLine
            If bDecimalComma Then
                sLine = Replace$(sLine, ",", ".")
            End If
            asFields = Split(sLine, sDelimiter)
            lNumFields = UBound(asFields) + 1
            For k = 1 To lNumColumns
                If k > lNumFields Then
                    adblRow(k) = 0
                Else
                    adblRow(k) = Val(asFields(k - 1))
                End If
            Next k
            lOffset = 1
            For i = 1 To lNumSegments
                CopyMemory ByVal aoTensors(i).Address + CLngPtr(j - 1) * alSegmentSizes(i) * SIZEOF_DOUBLE, _
                           adblRow(lOffset), _
                           alSegmentSizes(i) * SIZEOF_DOUBLE
                lOffset = lOffset + alSegmentSizes(i)
            Next i
        Next j
        .Close
    End With
    Set oResult = New TensorDataset
    oResult.Init aoTensors
    Set ImportDatasetFromCsv = oResult
End Function

Public Sub SplitDataset(ByVal oDataset As IDataset, _
                        ByVal dblAt As Double, _
                        ByRef A As SubsetDataset, _
                        ByRef B As SubsetDataset, _
                        Optional ByVal bRandomize As Boolean)
    Const PROCEDURE_NAME As String = "MLFactory.SplitDataset"
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
    If bRandomize Then
        alFullIndices = GetRandomPermutationArray(oDataset.NumSamples)
    Else
        alFullIndices = GetIdentityPermutationArray(oDataset.NumSamples)
    End If
    lSizeA = Int(dblAt * oDataset.NumSamples + 0.5)
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


