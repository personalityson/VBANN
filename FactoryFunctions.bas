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

Public Function XGBoost(ByVal oCriterion As ICriterion, _
                        Optional ByVal dblLearningRate As Double = 0.1, _
                        Optional ByVal lMaxDepth As Long = 6, _
                        Optional ByVal dblLambda As Double = 1, _
                        Optional ByVal dblGamma As Double = 0, _
                        Optional ByVal dblMinChildWeight As Double = 1) As XGBoost
    Set XGBoost = New XGBoost
    XGBoost.Init oCriterion, dblLearningRate, lMaxDepth, dblLambda, dblGamma, dblMinChildWeight
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
    Const PROCEDURE_NAME As String = "FactoryFunctions.ImportDatasetFromWorksheet"
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

'Public Function ImportDatasetFromCsv(ByVal strPath As String, _
'                                     ByVal vSegmentSizes As Variant, _
'                                     Optional ByVal bHasHeaders As Boolean, _
'                                     Optional ByVal bSqueeze As Boolean) As TensorDataset
'    Const PROCEDURE_NAME As String = "FactoryFunctions.ImportDatasetFromCsv"
'    Const CHUNK_SIZE As Long = 4096
'    Const ForReading As Long = 1
'    Dim i As Long
'    Dim j As Long
'    Dim j_chunk As Long
'    Dim k As Long
'    Dim lOffset As Long
'    Dim lTemp As Long
'    Dim lNumSegments As Long
'    Dim alSegmentSizes() As Long
'    Dim lNumRows As Long
'    Dim lNumFields As Long
'    Dim vFields As Variant
'    Dim adblChunk() As Double
'    Dim X_() As Double
'    Dim aoTensors() As Tensor
'
'    If Not Fso.FileExists(strPath) Then
'        Err.Raise 9, PROCEDURE_NAME, "Specified file does not exist."
'    End If
'    ParseVariantToLongArray vSegmentSizes, lNumSegments, alSegmentSizes
'    If lNumSegments < 1 Then
'        Set ImportDatasetFromCsv = TensorDataset(Array())
'        Exit Function
'    End If
'    For k = 1 To lNumSegments
'        If alSegmentSizes(k) < 1 Then
'            Err.Raise 5, PROCEDURE_NAME, "Segment size must be >= 1."
'        End If
'        lTemp = lTemp + alSegmentSizes(k)
'    Next k
'    With Fso.OpenTextFile(strPath, ForReading)
'        If Not .AtEndOfStream And bHasHeaders Then
'            .SkipLine
'        End If
'        Do While Not .AtEndOfStream
'            lNumRows = lNumRows + 1
'        Loop
'        .Close
'    End With
'    ReDim aoTensors(1 To lNumSegments)
'    For k = 1 To lNumSegments
'        Set aoTensors(k) = Zeros(Array(alSegmentSizes(k), lNumRows))
'    Next k
'    If lNumRows < 1 Then
'        Set ImportDatasetFromCsv = TensorDataset(aoTensors)
'        Exit Function
'    End If
'    ReDim adblChunk(1 To lTemp, 1 To CHUNK_SIZE)
'    With Fso.OpenTextFile(strPath, ForReading)
'        If Not .AtEndOfStream And bHasHeaders Then
'            .SkipLine
'        End If
'        Do While Not .AtEndOfStream
'            j = j + 1
'            j_chunk = ((j - 1) Mod CHUNK_SIZE) + 1
'            'vFields = Split(.ReadLine, ",")
'            vFields = Split(.ReadLine, ";")
'            lNumFields = UBound(vFields) + 1
'            If lNumFields <> lTemp Then
'                Err.Raise 5, PROCEDURE_NAME, "Number of fields must be => some number." 'lTemp
'            End If
'            For i = 1 To lNumFields
'                'adblChunk(i, j_chunk) = Val(vFields(i - 1))
'                adblChunk(i, j_chunk) = CDbl(vFields(i - 1))
'            Next i
'            If j_chunk = CHUNK_SIZE Or j = lNumRows Then
'                lOffset = 0
'                For k = 1 To lNumSegments
'                    aoTensors(k).CreateAlias X_
'                    For i = 1 To j_chunk
'                        CopyMemory X_(1, j), adblChunk(1 + lOffset, i), SIZEOF_DOUBLE * alSegmentSizes(k)
'                    Next i
'                    lOffset = lOffset + alSegmentSizes(k)
'                    aoTensors(k).RemoveAlias X_
'                Next k
'            End If
'            If j Mod 100 = 0 Then
'                Application.StatusBar = "ImportDatasetFromCsv progress: " & j
'                DoEvents
'            End If
'        Loop
'        .Close
'    End With
'    If bSqueeze Then
'        For k = 1 To lNumSegments
'            Set aoTensors(k) = aoTensors(k).Squeeze
'        Next k
'    End If
'    Set ImportDatasetFromCsv = TensorDataset(aoTensors)
'End Function

'Public Sub Init(ByVal sPath As String, _
'                ByVal vSegmentSizes As Variant, _
'                Optional ByVal bSqueeze As Boolean, _
'                Optional ByVal bHasHeaders As Boolean)
'    Const PROCEDURE_NAME As String = "FactoryFunctions.ImportDatasetFromWorksheet"
'    Const CHUNK_SIZE As Long = 10000
'    Dim i As Long
'    Dim lNumAllocatedRows As Long
'    Dim sLine As String
'
'    If Not Fso.FileExists(sPath) Then
'        Err.Raise 9, PROCEDURE_NAME, "Specified file does not exist."
'    End If
'    ParseVariantToLongArray vSegmentSizes, m_lNumSegments, m_alSegmentSizes
'    For i = 1 To m_lNumSegments
'        If m_alSegmentSizes(i) < 1 Then
'            Err.Raise 5, PROCEDURE_NAME, "Segment size must be >= 1."
'        End If
'    Next i
'    Clear
'    m_iFileHandle = FreeFile
'    Open sPath For Binary As #m_iFileHandle
'    If bHasHeaders Then
'        Line Input #m_iFileHandle, sLine
'    End If
'    m_lNumSamples = 0
'    Do While Not EOF(m_iFileHandle)
'        m_lNumSamples = m_lNumSamples + 1
'        If m_lNumSamples > lNumAllocatedRows Then
'            lNumAllocatedRows = lNumAllocatedRows + CHUNK_SIZE
'            ReDim Preserve m_apOffsets(1 To lNumAllocatedRows)
'        End If
'        m_apOffsets(m_lNumSamples) = Loc(m_iFileHandle)
'        Line Input #m_iFileHandle, sLine
'    Loop
'    If m_lNumSamples > 0 Then
'        ReDim Preserve m_apOffsets(1 To m_lNumSamples)
'    Else
'        Erase m_apOffsets
'    End If
'End Sub
'

Public Sub RandomSplit(ByVal oDataset As IDataset, _
                       ByVal dblAt As Double, _
                       ByRef A As SubsetDataset, _
                       ByRef B As SubsetDataset)
    Const PROCEDURE_NAME As String = "FactoryFunctions.RandomSplit"
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
