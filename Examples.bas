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
    lNumEpochs = 20

    'Prepare training data
    Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Concrete", Array(lInputSize, lLabelSize), bHasHeaders:=True)
    SplitDataset oFullSet, 0.8, oTrainingSet, oValidationSet, bRandomize:=True
    Set oTrainingLoader = DataLoader(oTrainingSet, lBatchSize)
    Set oValidationLoader = DataLoader(oValidationSet, oValidationSet.NumSamples)

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
    Set oTrainer = BackpropTrainer(oModel, L2Loss(), SGDW(0.001, 0.9, 0.0001, 20))

    'With bVerbose = True the losses go to the "Log" worksheet
    oTrainer.Fit oTrainingLoader, lNumEpochs, True, oValidationLoader

    'Compute validation loss
    MsgBox oTrainer.Evaluate(oValidationLoader)

    'Save everything to a worksheet, incl. loss function and optimizer state
    Serialize TRAINER_NAME, oTrainer

    'Load the trainer
    Set oTrainer = Deserialize(TRAINER_NAME)
    MsgBox oTrainer.Evaluate(oValidationLoader)

    'Continue training with a lower learning rate
    oTrainer.Optimizer.LearningRate = 0.0001
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

'Gradient boosted trees on the Binance sheet.
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
    lNumEpochs = 20

    'Prepare training data
    Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Concrete", Array(lInputSize, lLabelSize), bHasHeaders:=True)
    SplitDataset oFullSet, 0.8, oTrainingSet, oValidationSet, bRandomize:=True
    Set oTrainingLoader = DataLoader(oTrainingSet, lBatchSize)
    Set oValidationLoader = DataLoader(oValidationSet, oValidationSet.NumSamples)

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
    MsgBox oTrainer.Evaluate(oValidationLoader)

    'Continue training with smaller mutations
    'oTrainer.MutationSigma = 0.01
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

    'Prepare training data.
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
    PredictInWorksheet = Transpose(Y).ToArray
End Function

Public Sub WorkingWithTensors()
    Dim A As Tensor
    Dim B As Tensor
    Dim A_() As Double
    Dim B_() As Double
    Dim adblArray() As Double

    'Create a tensor A of zeros with shape (2, 3, 4).
    Set A = Zeros(Array(2, 3, 4))

    'Basic properties of A.
    MsgBox A.NumDimensions
    MsgBox A.Size(1)
    MsgBox A.Size(2)
    MsgBox A.Size(3)
    MsgBox A.NumElements
    MsgBox A.Address 'Pointer to the first element

    'Constant fills.
    Set A = Ones(Array(2, 3, 4))
    Set A = Full(Array(2, 3, 4), 777)

    'Identity matrix.
    Set A = Eye(3)

    'Random initializers.
    Set A = Uniform(Array(2, 3, 4), 0, 1)
    Set A = Normal(Array(2, 3, 4), 0, 1)
    Set A = Bernoulli(Array(2, 3, 4), 0.5)

    'Glorot and He initializers. FullyConnectedLayer uses GlorotUniform for its weights.
    Set A = GlorotUniform(Array(2, 3, 4), 8, 4)
    Set A = HeNormal(Array(2, 3, 4), 8)

    'Fill tensor A with a constant value.
    A.Fill 777

    'Copy the values of A into an existing tensor B. B must have the same number of elements.
    Set B = Zeros(A.Shape)
    B.Copy A

    'Clone tensor A into a new tensor B with its own copy of the data.
    Set B = A.Clone

    'Use ShapeEquals to check if A's shape matches (2, 3, 4).
    MsgBox A.ShapeEquals(Array(2, 3, 4))

    'spWildcard matches any size, spEllipsis matches any number of dimensions. Both return True.
    MsgBox A.ShapeEquals(Array(2, spWildcard, 4))
    MsgBox A.ShapeEquals(Array(spEllipsis, 4))

    'Create a different view of A with a new shape (6, 4).
    'This view shares the same underlying data, but has a different layout.
    Set B = A.View(Array(6, 4))

    'Create alias arrays for direct memory access.
    A.CreateAlias A_
    B.CreateAlias B_

    'Modify an element via the alias from A's perspective.
    A_(1, 1, 1) = 123

    'Both return 123.
    MsgBox A_(1, 1, 1)
    MsgBox B_(1, 1)

    'Erase the B_ alias to simulate clearing the fixed-size array.
    Erase B_

    'Both return 0.
    MsgBox A_(1, 1, 1)
    MsgBox B_(1, 1)

    'Remove the aliases to avoid memory deallocation.
    A.RemoveAlias A_
    B.RemoveAlias B_

    'Start from fresh values for the shape operations below.
    Set A = Uniform(Array(2, 3, 4))

    'Create a flattened view of A with shared underlying data. The new shape is (24).
    Set A = A.Flatten

    'Add singleton dimensions on both sides. The new shape is (1, 24, 1).
    Set A = A.View(Array(1, 24, 1))

    'Regroup the same 24 elements. The new shape is (4, 6).
    Set A = A.View(Array(4, 6))

    'Whole-tensor reductions return a Double.
    MsgBox A.Mean

    'Pick rows 4 and 1 along dimension 1, in that order. The new shape is (2, 6).
    Set B = A.Gather(1, Array(4, 1))

    'Reduce A along dimension 2 using mean reduction. The new shape is (4, 1).
    Set A = A.Reduce(2, rtMean)

    'Slice A along dimension 1 from index 3 to 4. The new shape is (2, 1).
    Set A = A.Slice(1, 3, 4)

    'Tile A along dimension 2, repeating it 3 times. The new shape is (2, 3).
    Set A = A.Tile(2, 3)

    'Copy tensor A to a native VBA array with the same shape.
    adblArray = A.ToArray

    'Create tensor A from a native VBA array.
    A.FromArray adblArray

    'Create tensor A from an Excel range. The cells must be numbers or empty. The new shape is (3, 2).
    A.FromRange ActiveSheet.Range("A1:B3")

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
    Set Y = VecClamp(A, 0.2, 0.8)   'Y = Min(Max(A, 0.2), 0.8)

    'Activations and their derivatives
    Set Y = VecSigmoid(A)
    Set Y = VecSigmoidDerivative(VecSigmoid(A))
    Set Y = VecTanh(A)
    Set Y = VecTanhDerivative(VecTanh(A))
    Set Y = VecLeakyReLU(A, 0.01)
    Set Y = VecLeakyReLUDerivative(A, 0.01)
    Set Y = VecSoftplus(A)

    'Linear combination: Y = alpha * A + beta * B.
    Set Y = VecLinComb(0.99, A, 0.01, B)

    'Y = A / (Sqrt(B + inner_epsilon) + outer_epsilon)
    Set C = VecPow2(B)
    Set Y = VecDivRms(A, C, 0, 0.00000001)  'Used by AdamW
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

    'In-place ops can broadcast B but never A
    'VecAdd_I A, B works
    'VecAdd_I B, A would raise an error
    VecAdd_I A, B

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

    'Transpose swaps dimensions 1 and 2 by default, or any two dimensions you name.
    Set A = Uniform(Array(3, 4))
    Set Y = Transpose(A)             'Y is 4x3
    Set A = Uniform(Array(2, 3, 4))
    Set Y = Transpose(A, 1, 3)       'Y is 4x3x2

    Beep
End Sub
