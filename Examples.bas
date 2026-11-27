Attribute VB_Name = "Examples"
Option Explicit

Public Sub SetupAndTrainBackprop()
    Const MODEL_NAME As String = "MySequentialModel"
    Const TRAINER_NAME As String = "MyBackpropTrainer"
    Dim lInputSize As Long
    Dim lLabelSize As Long
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim adblPrediction() As Double
    Dim oFullSet As TensorDataset
    Dim oTrainingSet As SubsetDataset
    Dim oTrainingLoader As DataLoader
    Dim oValidationSet As SubsetDataset
    Dim oValidationLoader As DataLoader
    Dim oBatch As TensorDataset
    Dim oModel As Sequential
    Dim oTrainer As BackpropTrainer

    lInputSize = 8
    lLabelSize = 1
    lBatchSize = 16
    lNumEpochs = 40

    'Prepare training data
    Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Concrete", Array(lInputSize, lLabelSize), bHasHeaders:=True)
    SplitDataset oFullSet, 0.8, oTrainingSet, oValidationSet, bRandomize:=True
    Set oTrainingLoader = DataLoader(oTrainingSet, lBatchSize)
    Set oValidationLoader = DataLoader(oValidationSet, oValidationSet.NumSamples)

    'VBA uses column-major array layout -- the first index moves fastest (same as in Fortran, MATLAB, Julia etc)
    'Because of that, samples go in the last dimension:
    'Data is (features, samples)

    'Inputs are normalized on the way in, last layer maps outputs back to the label scale
    Set oModel = Sequential()
    oModel.Add NormalizationLayer(oTrainingSet, lIndex:=1)
    oModel.Add FullyConnectedLayer(lInputSize, 32)
    oModel.Add LeakyReLULayer()
    oModel.Add DropoutLayer(0.2)
    oModel.Add FullyConnectedLayer(32, 16)
    oModel.Add LeakyReLULayer()
    oModel.Add DropoutLayer(0.1)
    oModel.Add FullyConnectedLayer(16, lLabelSize)
    oModel.Add NormalizationLayer(oTrainingSet, lIndex:=2, bInvert:=True)

    'SGDW optimizer with gradient clipping at a norm of 20
    Set oTrainer = BackpropTrainer(oModel, L2Loss(), SGDW(0.01, 0.9, 0.001, 20))

    'With bVerbose = True the losses go to the "Log" worksheet
    oTrainer.Fit oTrainingLoader, lNumEpochs, True, oValidationLoader

    'Compute validation loss
    MsgBox oTrainer.Evaluate(oValidationLoader)

    'Save everything to a worksheet, incl. loss function and optimizer state
    Serialize TRAINER_NAME, oTrainer

    'Load the trainer
    Set oTrainer = Deserialize(TRAINER_NAME)

    'Check if validation loss is the same with the deserialized model
    MsgBox oTrainer.Evaluate(oValidationLoader)

    'Continue training with a lower learning rate
    oTrainer.Optimizer.LearningRate = 0.003
    oTrainer.Fit oTrainingLoader, lNumEpochs, True, oValidationLoader

    'Save the model only
    Serialize MODEL_NAME, oTrainer.Model

    'Load the model
    Set oModel = Deserialize(MODEL_NAME)

    'Predict the first 5 validation samples
    Set oBatch = oValidationSet.Gather(Array(1, 2, 3, 4, 5))
    adblPrediction = oModel.Forward(oBatch.Tensors(1)).ToArray

    Beep
End Sub

Public Sub SetupAndTrainBackpropClassification()
    Const MODEL_NAME As String = "MyClassifierModel"
    Dim lInputSize As Long
    Dim lLabelSize As Long
    Dim lNumClasses As Long
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim dblAccuracy As Double
    Dim Y As Tensor
    Dim P As Tensor
    Dim L As Tensor
    Dim oFullSet As TensorDataset
    Dim oTrainingSet As SubsetDataset
    Dim oTrainingLoader As DataLoader
    Dim oValidationSet As SubsetDataset
    Dim oValidationLoader As DataLoader
    Dim oModel As Sequential
    Dim oInferenceModel As Sequential
    Dim oTrainer As BackpropTrainer

    lInputSize = 4
    lLabelSize = 1
    lNumClasses = 3
    lBatchSize = 8
    lNumEpochs = 40

    'The first column in training data is the ID, which we will skip
    'The last label column holds class indices 1..lNumClasses
    'Id  SepalLengthCm  SepalWidthCm    PetalLengthCm   PetalWidthCm    Species
    '1   5,1            3,5             1,4             0,2             1
    '2   4,9            3               1,4             0,2             1
    Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Iris", Array(1, lInputSize, 1), bHasHeaders:=True)
    'OneHot turns 1, 2, 3 class indices into (1, 0, 0), (0, 1, 0), (0, 0, 1) targets
    'oFullSet.Tensors(1) is not used
    Set oFullSet = TensorDataset(Array(oFullSet.Tensors(2), OneHot(oFullSet.Tensors(3), lNumClasses)))
    SplitDataset oFullSet, 0.8, oTrainingSet, oValidationSet, bRandomize:=True
    Set oTrainingLoader = DataLoader(oTrainingSet, lBatchSize)
    Set oValidationLoader = DataLoader(oValidationSet, oValidationSet.NumSamples)

    'Inputs are normalized on the way in, the last layer outputs one logit per class
    Set oModel = Sequential()
    oModel.Add NormalizationLayer(oTrainingSet, lIndex:=1)
    oModel.Add FullyConnectedLayer(lInputSize, 8)
    oModel.Add LeakyReLULayer()
    oModel.Add DropoutLayer(0.2)
    oModel.Add FullyConnectedLayer(8, 4)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(4, lNumClasses)

    'CCELoss on logits applies softmax internally, which is more stable
    Set oTrainer = BackpropTrainer(oModel, CCELoss(bFromLogits:=True), SGDW())

    'With bVerbose = True the losses go to the "Log" worksheet
    oTrainer.Fit oTrainingLoader, lNumEpochs, True, oValidationLoader

    'Validation loss
    MsgBox oTrainer.Evaluate(oValidationLoader)

    'Accuracy
    Set Y = oModel.Forward(oValidationSet.Tensors(1))
    Set P = Y.Reduce(1, rtArgMax)
    Set L = oValidationSet.Tensors(2).Reduce(1, rtArgMax)
    dblAccuracy = 1 - VecSign(VecAbs(VecSub(P, L))).Mean
    MsgBox Format(dblAccuracy, "0.0%")

    'Softmax is added at the end for inference, so the saved model outputs probabilities
    Set oInferenceModel = Sequential()
    oInferenceModel.Add oModel
    oInferenceModel.Add SoftmaxLayer()
    
    'Save the model
    Serialize MODEL_NAME, oInferenceModel

    Beep
End Sub

Public Sub SetupAndTrainXGBoost()
    Const MODEL_NAME As String = "MyXGBoostModel"
    Dim lInputSize As Long
    Dim lLabelSize As Long
    Dim lNumRounds As Long
    Dim adblPrediction() As Double
    Dim oFullSet As TensorDataset
    Dim oTrainingSet As SubsetDataset
    Dim oValidationSet As SubsetDataset
    Dim oBatch As TensorDataset
    Dim oModel As XGBoost

    lInputSize = 8
    lLabelSize = 1
    lNumRounds = 40

    'Prepare training data
    Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Concrete", Array(lInputSize, lLabelSize), bHasHeaders:=True)
    SplitDataset oFullSet, 0.8, oTrainingSet, oValidationSet, bRandomize:=True

    'VBA uses column-major array layout -- the first index moves fastest (same as in Fortran, MATLAB, Julia etc)
    'Because of that, samples go in the last dimension:
    'Data is (features, samples)

    'XGBoost needs a loss that works on raw scores (FromLogits = True), L2Loss always does
    Set oModel = XGBoost(L2Loss(), dblLearningRate:=0.1, lMaxDepth:=6, dblSubsample:=0.75, dblColsample:=0.75)

    'With bVerbose = True the losses go to the "Log" worksheet
    oModel.Fit oTrainingSet, lNumRounds, True, oValidationSet

    'Compute validation loss
    MsgBox oModel.Evaluate(oValidationSet)

    'Save
    Serialize MODEL_NAME, oModel

    'Load
    Set oModel = Deserialize(MODEL_NAME)

    'Check if validation loss is the same with the deserialized model
    MsgBox oModel.Evaluate(oValidationSet)

    'Predict the first 5 validation samples
    Set oBatch = oValidationSet.Gather(Array(1, 2, 3, 4, 5))
    adblPrediction = oModel.Predict(oBatch.Tensors(1)).ToArray

    Beep
End Sub

Public Sub SetupAndTrainEvolutionaryStrategies()
    Const MODEL_NAME As String = "MySequentialModel"
    Const TRAINER_NAME As String = "MyOpenES"
    Dim lInputSize As Long
    Dim lLabelSize As Long
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim adblPrediction() As Double
    Dim oFullSet As TensorDataset
    Dim oTrainingSet As SubsetDataset
    Dim oTrainingLoader As DataLoader
    Dim oValidationSet As SubsetDataset
    Dim oValidationLoader As DataLoader
    Dim oBatch As TensorDataset
    Dim oModel As Sequential
    Dim oTrainer As OpenES

    lInputSize = 8
    lLabelSize = 1
    lBatchSize = 16
    lNumEpochs = 40

    'Prepare training data
    Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Concrete", Array(lInputSize, lLabelSize), bHasHeaders:=True)
    SplitDataset oFullSet, 0.8, oTrainingSet, oValidationSet, bRandomize:=True
    Set oTrainingLoader = DataLoader(oTrainingSet, lBatchSize)
    Set oValidationLoader = DataLoader(oValidationSet, oValidationSet.NumSamples)

    'VBA uses column-major array layout -- the first index moves fastest (same as in Fortran, MATLAB, Julia etc)
    'Because of that, samples go in the last dimension:
    'Data is (features, samples)

    'OpenES always runs the model in inference mode, so no dropout here
    Set oModel = Sequential()
    oModel.Add NormalizationLayer(oTrainingSet, lIndex:=1)
    oModel.Add FullyConnectedLayer(lInputSize, 32)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(32, 16)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(16, lLabelSize)
    oModel.Add NormalizationLayer(oTrainingSet, lIndex:=2, bInvert:=True)

    'The population size must be even
    Set oTrainer = OpenES(oModel, L2Loss(), SGDW(dblLearningRate:=0.002), lPopulationSize:=50, dblMutationSigma:=0.02)

    'With bVerbose = True the losses go to the "Log" worksheet
    oTrainer.Fit oTrainingLoader, lNumEpochs, True, oValidationLoader

    'Compute validation loss
    MsgBox oTrainer.Evaluate(oValidationLoader)

    'Save everything to a worksheet, incl. loss function and optimizer state
    Serialize TRAINER_NAME, oTrainer

    'Load the trainer
    Set oTrainer = Deserialize(TRAINER_NAME)

    'Check if validation loss is the same with the deserialized model
    MsgBox oTrainer.Evaluate(oValidationLoader)

    'Continue training with smaller learning rate
    oTrainer.Optimizer.LearningRate = 0.001
    oTrainer.Fit oTrainingLoader, lNumEpochs, True, oValidationLoader

    'Save the model only
    Serialize MODEL_NAME, oTrainer.Model

    'Load the model
    Set oModel = Deserialize(MODEL_NAME)

    'Predict the first 5 validation samples
    Set oBatch = oValidationSet.Gather(Array(1, 2, 3, 4, 5))
    adblPrediction = oModel.Forward(oBatch.Tensors(1)).ToArray

    Beep
End Sub

Public Sub SetupAndTrainBackpropPretraining()
    Dim lInputSize As Long
    Dim lLabelSize As Long
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim X_norm As Tensor
    Dim oFullSet As TensorDataset
    Dim oTrainingSet As SubsetDataset
    Dim oTrainingLoader As DataLoader
    Dim oValidationSet As SubsetDataset
    Dim oValidationLoader As DataLoader
    Dim oNormalSet As TensorDataset
    Dim oNormalLoader As DataLoader
    Dim oInputNormalization As NormalizationLayer
    Dim oEncoder As Sequential
    Dim oAutoencoder As Sequential
    Dim oModel As Sequential
    Dim oPretrainer As BackpropTrainer
    Dim oFinetuner As BackpropTrainer

    lInputSize = 8
    lLabelSize = 1
    lBatchSize = 16
    lNumEpochs = 40

    'Prepare training data
    Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Concrete", Array(lInputSize, lLabelSize), bHasHeaders:=True)
    SplitDataset oFullSet, 0.8, oTrainingSet, oValidationSet, bRandomize:=True
    Set oTrainingLoader = DataLoader(oTrainingSet, lBatchSize)
    Set oValidationLoader = DataLoader(oValidationSet, oValidationSet.NumSamples)

    'The normalized inputs are both the input and the target
    Set oInputNormalization = NormalizationLayer(oTrainingSet, lIndex:=1)
    Set X_norm = oInputNormalization.Forward(oTrainingSet.Tensors(1))
    Set oNormalSet = TensorDataset(Array(X_norm, X_norm))
    Set oNormalLoader = DataLoader(oNormalSet, lBatchSize)

    Set oEncoder = Sequential()
    oEncoder.Add FullyConnectedLayer(lInputSize, 32)
    oEncoder.Add LeakyReLULayer()
    oEncoder.Add FullyConnectedLayer(32, 16)
    oEncoder.Add LeakyReLULayer()

    'Dropout makes this a denoising autoencoder
    Set oAutoencoder = Sequential()
    oAutoencoder.Add DropoutLayer(0.2)
    oAutoencoder.Add oEncoder
    oAutoencoder.Add FullyConnectedLayer(16, lInputSize)

    Set oPretrainer = BackpropTrainer(oAutoencoder, L2Loss(), SGDW())
    oPretrainer.Fit oNormalLoader, lNumEpochs, True

    'The fine-tuned model starts from the pretrained weights, last layer maps outputs back to the label scale
    Set oModel = Sequential()
    oModel.Add oInputNormalization
    oModel.Add oEncoder
    oModel.Add FullyConnectedLayer(16, lLabelSize)
    oModel.Add NormalizationLayer(oTrainingSet, lIndex:=2, bInvert:=True)

    'Fine-tune it with a regression head
    Set oFinetuner = BackpropTrainer(oModel, L2Loss(), SGDW())
    oFinetuner.Fit oTrainingLoader, lNumEpochs, True, oValidationLoader

    'Compute validation loss
    MsgBox oFinetuner.Evaluate(oValidationLoader)

    Beep
End Sub

'Worksheet function, e.g. =PredictInWorksheet(A2:H10)
Public Function PredictInWorksheet(ByVal oInput As Range) As Variant
    Const MODEL_NAME As String = "MySequentialModel"
    Static s_oModel As Sequential
    Dim X As Tensor
    Dim Y As Tensor

    If s_oModel Is Nothing Then
        Set s_oModel = Deserialize(MODEL_NAME)
    End If
    Set X = TensorFromRange(oInput, True)
    Set Y = s_oModel.Forward(X)
    PredictInWorksheet = Y.Transpose.ToArray
End Function

Public Sub WorkingWithTensors()
    Dim i As Long
    Dim M_() As Double
    Dim F_() As Double
    Dim adblArray() As Double
    Dim M As Tensor
    Dim A As Tensor
    Dim B As Tensor
    Dim X As Tensor

    'A tensor of zeros with shape (2, 3, 4)
    Set A = Zeros(Array(2, 3, 4))
    Debug.Print A.NumDimensions         '3
    Debug.Print A.Size(1)               '2
    Debug.Print A.SizeBetween(1, 2)     '6 = 2 * 3
    Debug.Print A.NumElements           '24
    Debug.Print A.Address               'Pointer to the first element

    'Constant fills and the identity matrix
    Set A = Ones(Array(2, 3, 4))
    Set A = Full(Array(2, 3, 4), 777)
    Set A = Eye(3)

    'Random initializers
    Set A = Uniform(Array(2, 3, 4), 0, 1)
    Set A = Normal(Array(2, 3, 4), 0, 1)
    Set A = Bernoulli(Array(2, 3, 4), 0.5)

    'Glorot and He initializers
    Set A = GlorotUniform(Array(2, 3, 4), 8, 4)
    Set A = HeNormal(Array(2, 3, 4), 8)

    'No dimensions means no elements
    Set A = Zeros(Array())
    Debug.Print A.NumElements           '0
    Set A = Zeros(Array(1))
    Debug.Print A.NumElements           '1

    'Size-0 dimensions are allowed
    Set A = Zeros(Array(0, 5))
    Debug.Print A.NumElements           '0

    'AsArray and AsFlatArray point a native array at the tensor's memory -- no copy
    Set M = Zeros(Array(2, 3))
    M.AsArray M_
    M.AsFlatArray F_
    For i = 1 To M.NumElements
        F_(i) = i
    Next i

    'Memory order is (1,1), (2,1), (1,2), (2,2), (1,3), (2,3)
    Debug.Print M_(2, 1)                '2
    Debug.Print M_(1, 2)                '3

    'View reads the same memory with a new shape
    Set B = M.View(Array(3, 2))
    Debug.Print B.ToArray()(3, 1)

    'Writing to B changes the original M, because they share storage
    Set B = M.Flatten
    B.Fill 0
    Debug.Print M_(2, 3)                '0

    'A batch of 5 samples with 3 features each has shape (3, 5)
    Set X = Normal(Array(3, 5))

    'Pick samples 2 and 4 by gathering along the last dimension
    Set B = X.Gather(2, Array(2, 4))    '(3, 2)

    'One sample is a contiguous column of features
    Set B = X.Slice(2, 1, 1)            '(3, 1)

    'In a worksheet each row is usually a sample, bTranspose:=True turns rows into samples
    Set B = TensorFromRange(Worksheets("Concrete").Range("A2:B4"))         '(3, 2)
    Set B = TensorFromRange(Worksheets("Concrete").Range("A2:B4"), True)   '(2, 3)

    'Indexing is 1-based and Slice includes its end
    Set A = Uniform(Array(4, 6))
    Set B = A.Slice(1, 2, 3)            'Rows 2 and 3
    Set B = A.Gather(1, Array(4, 1))    'Rows 4 and 1

    'Reductions keep the reduced dimension with size 1
    Set B = A.Reduce(2, rtMean)         '(4, 1)
    Set B = A.Reduce(1, rtArgMax)       '(1, 6)

    'Views share storage, everything else copies
    Set B = A.Flatten                   'View, shape (24)
    Set B = A.View(Array(1, 24, 1))     'View, any shape with the same number of elements
    Set B = A.Expand(3)                 'View, shape (4, 6, 1)
    Set B = A.Transpose                 'Copy, shape (6, 4)
    Set B = A.Tile(2, 3)                'Copy, shape (4, 18)

    'Every element set to 777
    A.Fill 777
    
    'Copies values, must have the same number of elements
    Set B = Zeros(A.Shape)
    B.Copy A
    
    'Clone with its own copy of the data
    Set B = A.Clone

    'Checking shapes
    Debug.Print A.ShapeEquals(Array(4, 6))              'True
    Debug.Print A.ShapeEquals(Array(4, spWildcard))     'True, spWildcard matches any size
    Debug.Print A.ShapeEquals(Array(spEllipsis, 6))     'True, spEllipsis matches any number of dimensions

    'A copy with the same shape, any number of dimensions
    adblArray = A.ToArray
    
    'Tensor from array
    Set B = TensorFromArray(adblArray)

    Beep
End Sub

Public Sub WorkingWithTensorOps()
    Dim A As Tensor
    Dim B As Tensor
    Dim C As Tensor
    Dim Y As Tensor
    Dim dblValue As Double

    'Two random vectors
    Set A = Uniform(Array(5), 0, 1)
    Set B = Uniform(Array(5), 0, 1)

    'VecDot computes Sum(A * B)
    dblValue = VecDot(A, B)

    'VecNorm2 computes Sqrt(Sum(A^2))
    dblValue = VecNorm2(A)

    'Element-wise binary
    Set Y = VecAdd(A, B)        'Y = A + B
    Set Y = VecSub(A, B)        'Y = A - B
    Set Y = VecMul(A, B)        'Y = A * B  (not matrix multiply)
    Set Y = VecDiv(A, B)        'Y = A / B

    'Element-wise scalar
    Set Y = VecAddC(A, 10)      'Y = A + 10
    Set Y = VecSubC(A, 10)      'Y = A - 10
    Set Y = VecSubCRev(A, 10)   'Y = 10 - A
    Set Y = VecMulC(A, 2)       'Y = A * 2
    Set Y = VecDivC(A, 2)       'Y = A / 2
    Set Y = VecDivCRev(A, 1)    'Y = 1 / A

    'Element-wise unary
    Set Y = VecAbs(A)
    Set Y = VecSign(A)
    Set Y = VecPow2(A)
    Set Y = VecSqrt(A)
    Set Y = VecExp(A)
    Set Y = VecLog(A)
    Set Y = VecClamp(A, 0.2, 0.8)

    'Activations and their derivatives
    Set Y = VecSigmoid(A)
    Set Y = VecSigmoidDerivative(VecSigmoid(A))
    Set Y = VecTanh(A)
    Set Y = VecTanhDerivative(VecTanh(A))
    Set Y = VecLeakyReLU(A, 0.01)
    Set Y = VecLeakyReLUDerivative(A, 0.01)
    Set Y = VecSoftplus(A)

    'Y = alpha * A + beta * B
    Set Y = VecLinComb(0.99, A, 0.01, B)

    'Y = A / (Sqrt(B + inner_epsilon) + outer_epsilon)
    Set C = VecPow2(B)
    Set Y = VecDivRms(A, C, 0, 0.00000001)
    Set Y = VecDivRms(A, C, 0.00001, 0)

    'VecWhere selects element-wise Y = IIf(C <> 0, A, B)
    Set C = Bernoulli(Array(5), 0.5)
    Set Y = VecWhere(C, A, B)

    'In-place A = 2 * A
    VecMulC_I A, 2

    'In-place A = A + B
    VecAdd_I A, B

    'In-place A = A * B
    VecMul_I A, B

    'In-place A = alpha * A + beta * B
    VecLinComb_I 0.99, A, 0.01, B

    'Broadcasting: a dimension of size 1, or a missing trailing dimension, is repeated to match
    Set A = Uniform(Array(3, 4))
    Set B = Uniform(Array(3))
    Set Y = VecAdd(A, B)             'Y is 3x4

    'B with shape (1, 4) scales each column of A by its own factor
    Set B = Uniform(Array(1, 4))
    Set Y = VecMul(A, B)             'Y is 3x4

    'Standard matrix product
    Set A = Uniform(Array(3, 4))     '3x4
    Set B = Uniform(Array(4, 5))     '4x5
    Set Y = MatMul(A, B)             'Y is 3x5

    'MatMul also transposes
    Set A = Uniform(Array(4, 3))     '4x3
    Set Y = MatMul(A, B, True)       'Y = A' * B, so Y is 3x5

    'MatMul_I accumulates C = C + A * B
    Set A = Uniform(Array(3, 4))
    Set B = Uniform(Array(4, 5))
    Set C = Zeros(Array(3, 5))
    MatMul_I C, A, B                 'C = A * B
    MatMul_I C, A, B                 'C = C + A * B  (so now 2 * A * B)

    'Softmax and LogSoftmax along dimension 1 (classes in rows, samples in columns)
    Set A = Normal(Array(3, 4))
    Set Y = Softmax(A, 1)
    Set Y = LogSoftmax(A, 1)

    Beep
End Sub
