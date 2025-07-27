Attribute VB_Name = "Examples"
Option Explicit

Public Sub SetupAndTrainSequential()
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

    lInputSize = 8
    lLabelSize = 1
    lBatchSize = 16
    lNumEpochs = 40

    'Prepare training data
    Set oFullSet = ImportDatasetFromWorksheet("Concrete", Array(lInputSize, lLabelSize), True, False)
    RandomSplit oFullSet, 0.8, oTrainingSet, oTestSet
    Set oTrainingLoader = DataLoader(oTrainingSet, lBatchSize)
    Set oTestLoader = DataLoader(oTestSet, lBatchSize)
    
    'Setup and train
    Set oModel = Sequential(L2Loss(), SGDM())
    oModel.Add InputNormalizationLayer(oTrainingLoader)
    oModel.Add FullyConnectedLayer(lInputSize, 200)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(200, 100)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(100, 50)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(50, lLabelSize)
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

Public Sub SetupAndTrainXGBoost()
    Const MODEL_NAME As String = "MyXGBoostModel"
    Dim lInputSize As Long
    Dim lLabelSize As Long
    Dim lNumRounds As Long
    Dim lMaxDepth As Long
    Dim dblLearningRate As Double
    Dim X As Tensor
    Dim T As Tensor
    Dim oFullSet As TensorDataset
    Dim oTrainingSet As SubsetDataset
    Dim oTestSet As SubsetDataset
    Dim oModel As XGBoost

    lInputSize = 8
    lLabelSize = 1
    dblLearningRate = 0.1
    lMaxDepth = 6
    lNumRounds = 100

    'Prepare training data
    Set oFullSet = ImportDatasetFromWorksheet("Concrete", Array(lInputSize, lLabelSize), True, True)
    RandomSplit oFullSet, 0.8, oTrainingSet, oTestSet
    Set X = oTrainingSet.Cache.Tensor(1)
    Set T = oTrainingSet.Cache.Tensor(2)

    'Setup and train
    Set oModel = XGBoost(L2Loss(), dblLearningRate, lMaxDepth)
    oModel.Fit X, T, lNumRounds

    'Compute test loss
    Set X = oTestSet.Cache.Tensor(1)
    Set T = oTestSet.Cache.Tensor(2)
    MsgBox oModel.Loss(X, T)

    'Save to worksheet
    Serialize MODEL_NAME, oModel

    'Load from worksheet
    Set oModel = Unserialize(MODEL_NAME)

    'Compute test loss again with unserialized model
    MsgBox oModel.Loss(X, T)

    Beep
End Sub

Public Function PredictInWorksheet(ByVal oInput As Range) As Double()
    Const MODEL_NAME As String = "MySequentialModel"
    Static s_oModel As Sequential
    Dim X As Tensor
    Dim Y As Tensor
    
    If s_oModel Is Nothing Then
        Set s_oModel = Unserialize(MODEL_NAME)
    End If
    Set X = TensorFromRange(oInput, True)
    Set Y = s_oModel.Predict(X)
    PredictInWorksheet = Y.ToArray
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

    'Create a tensor A filled with constant values.
    Set A = Ones(Array(2, 3, 4))
    Set A = Full(Array(2, 3, 4), 777)

    'Create a tensor A filled with random values.
    Set A = Uniform(Array(2, 3, 4), 0, 1)
    Set A = Normal(Array(2, 3, 4), 0, 1)
    Set A = Bernoulli(Array(2, 3, 4), 0.5)

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
End Sub
