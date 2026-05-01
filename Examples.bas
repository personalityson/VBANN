Attribute VB_Name = "Examples"
Option Explicit

Public Sub SetupAndTrain()
    Const MODEL_NAME As String = "MySequentialModel"
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim lInputSize As Long
    Dim lLabelSize As Long
    Dim oFullSet As TensorDataset
    Dim oTrainingSet As SubsetDataset
    Dim oTrainingLoader As DataLoader
    Dim oTestSet As SubsetDataset
    Dim oTestLoader As DataLoader
    Dim oModel As Sequential

    Randomize Timer

    lInputSize = 8
    lLabelSize = 1
    lBatchSize = 32
    lNumEpochs = 40

    'Prepare training data
    Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Concrete", Array(lInputSize, lLabelSize), True)
    SplitDataset oFullSet, 0.8, oTrainingSet, oTestSet, True
    Set oTrainingLoader = DataLoader(oTrainingSet, lBatchSize)
    Set oTestLoader = DataLoader(oTestSet, lBatchSize)

    'Setup and train
    Set oModel = Sequential(L2Loss(), SGDW())
    oModel.Add InputNormalizationLayer(oTrainingLoader)
    oModel.Add FullyConnectedLayer(lInputSize, 32)
    oModel.Add LeakyReLULayer()
    oModel.Add DropoutLayer(0.3)
    oModel.Add FullyConnectedLayer(32, 16)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(16, lLabelSize)
    oModel.Fit oTrainingLoader, oTestLoader, lNumEpochs

    'Compute test loss
    MsgBox oModel.Loss(oTestLoader)

    'Save to worksheet
    Serialize MODEL_NAME, oModel

    'Load from worksheet
    Set oModel = Unserialize(MODEL_NAME)

    'Compute test loss again with unserialized model
    MsgBox oModel.Loss(oTestLoader)

    Beep
End Sub

Public Sub ContinueTraining()
    Const MODEL_NAME As String = "MySequentialModel"
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim lInputSize As Long
    Dim lLabelSize As Long
    Dim oFullSet As TensorDataset
    Dim oTrainingSet As SubsetDataset
    Dim oTestSet As SubsetDataset
    Dim oTrainingLoader As DataLoader
    Dim oTestLoader As DataLoader
    Dim oModel As Sequential

    lInputSize = 8
    lLabelSize = 1
    lBatchSize = 32
    lNumEpochs = 20

    Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Concrete", Array(lInputSize, lLabelSize), True)
    SplitDataset oFullSet, 0.8, oTrainingSet, oTestSet, True
    Set oTrainingLoader = DataLoader(oTrainingSet, lBatchSize)
    Set oTestLoader = DataLoader(oTestSet, lBatchSize)

    Set oModel = Unserialize(MODEL_NAME)

    MsgBox "Test loss before continued training: " & oModel.Loss(oTestLoader)

    oModel.Fit oTrainingLoader, oTestLoader, lNumEpochs

    MsgBox "Test loss after continued training: " & oModel.Loss(oTestLoader)

    Serialize MODEL_NAME, oModel

    Beep
End Sub

Public Function PredictInWorksheet(ByVal oInput As Range) As Variant
    Const MODEL_NAME As String = "MySequentialModel"
    Static s_oModel As Sequential
    Dim X As Tensor
    Dim Y As Tensor

    If s_oModel Is Nothing Then
        Set s_oModel = Unserialize(MODEL_NAME)
    End If
    Set X = TensorFromRange(oInput, True)
    Set Y = s_oModel.Predict(X)
    PredictInWorksheet = MatTranspose(Y).ToArray
End Function

Public Sub WorkingWithTensors()
    Dim A As Tensor
    Dim B As Tensor
    Dim A_() As Double
    Dim B_() As Double
    Dim adblArray() As Double

    'Create an empty tensor A filled with zeros, with shape (2, 3, 4).
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

    'Glorot and He initializers are what FullyConnectedLayer uses internally.
    Set A = GlorotUniform(Array(4, 8), 8, 4)
    Set A = HeNormal(Array(4, 8), 8)

    'Fill tensor A with a constant value.
    A.Fill 777

    'Copy tensor A into a new tensor B. (B must be resized to match A's shape.)
    Set B = New Tensor
    B.Resize A.Shape
    B.Copy A

    'Clone tensor A into a new tensor B.
    Set B = A.Clone

    'Use ShapeEquals to check if A's shape matches (2, 3, 4).
    MsgBox A.ShapeEquals(Array(2, 3, 4))

    'Create a different view of A with a new shape (6, 4).
    'This view shares the same underlying data, but has a different layout.
    Set B = A.View(Array(6, 4))

    'Create alias arrays for direct memory access.
    A.CreateAlias A_
    B.CreateAlias B_

    ' Modify an element via the alias from A's perspective.
    A_(1, 1, 1) = 777

    'Both return 777.
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

    'Create a flattened view of A with shared underlying data. The new shape is (24).
    Set A = A.Flatten

    'Add singleton dimensions on both sides. The new shape is (1, 24, 1).
    Set A = A.View(Array(1, 24, 1))

    'Reshape A to a 2D tensor (4, 6). Number of elements must remain the same.
    A.Reshape Array(4, 6)

    'Reduce A along dimension 2 using mean reduction. The new shape is (4, 1).
    Set A = A.Reduce(2, rdcMean)

    'Slice A along dimension 1 from index 3 to 4. The new shape is (2, 1).
    Set A = A.Slice(1, 3, 4)

    'Tile A along dimension 2, repeating it 3 times. The new shape is (2, 3).
    Set A = A.Tile(2, 3)

    'Create tensor A from a native VBA array.
    A.FromArray adblArray

    'Copy tensor A to a native VBA array.
    adblArray = A.ToArray

    'Create tensor A from an Excel range.
    A.FromRange ActiveSheet.Range("A1:B3")

    Beep
End Sub


Public Sub WorkingWithTensorOps()
    Dim A As Tensor
    Dim B As Tensor
    Dim C As Tensor
    Dim Y As Tensor
    Dim dblValue As Double

    'Two random vectors to work with.
    Set A = Uniform(Array(5), 0, 1)
    Set B = Uniform(Array(5), 0, 1)

    'VecDot computes Sum(A * B).
    dblValue = VecDot(A, B)

    'VecNorm2 computes Sqrt(Sum(A^2)).
    dblValue = VecNorm2(A)

    'Element-wise binary (A and B must have the same NumElements).
    Set Y = VecAdd(A, B)        'Y = A + B
    Set Y = VecSub(A, B)        'Y = A - B
    Set Y = VecMul(A, B)        'Y = A * B  (not matrix multiply)
    Set Y = VecDiv(A, B)        'Y = A / B

    'Element-wise scalar.
    Set Y = VecAddC(A, 10)      'Y = A + 10
    Set Y = VecSubC(A, 10)      'Y = A - 10
    Set Y = VecSubCRev(A, 10)   'Y = 10 - A
    Set Y = VecMulC(A, 2)       'Y = A * 2
    Set Y = VecDivC(A, 2)       'Y = A / 2
    Set Y = VecDivCRev(A, 1)    'Y = 1 / A

    'Element-wise unary.
    Set Y = VecAbs(A)
    Set Y = VecSign(A)
    Set Y = VecPow2(A)
    Set Y = VecSqrt(A)
    Set Y = VecExp(A)
    Set Y = VecLog(A)

    'Activations and their derivatives.
    Set Y = VecSigmoid(A)
    Set Y = VecSigmoidDerivative(VecSigmoid(A))
    Set Y = VecTanh(A)
    Set Y = VecTanhDerivative(VecTanh(A))
    Set Y = VecLeakyReLU(A, 0.01)
    Set Y = VecLeakyReLUDerivative(A, 0.01)

    'In-place A := 2 * A
    VecMulC_I A, 2

    'In-place A := A + B
    VecAdd_I A, B

    'In-place A := alpha * A + beta * B
    VecLinComb_I 0.9, A, 0.1, B

    'Standard matrix product.
    Set A = Uniform(Array(3, 4))     '3x4
    Set B = Uniform(Array(4, 5))     '4x5
    Set Y = MatMul(A, B)             'Y is 3x5

    'MatMul_I accumulates: C := C + A * B.
    Set A = Uniform(Array(3, 4))
    Set B = Uniform(Array(4, 5))
    Set C = Zeros(Array(3, 5))
    MatMul_I C, A, B                 'C := A * B
    MatMul_I C, A, B                 'C := C + A * B  (so now 2 * A * B)

    'Transpose.
    Set A = Uniform(Array(3, 4))
    Set Y = MatTranspose(A)          'Y is 4x3

    Beep
End Sub
