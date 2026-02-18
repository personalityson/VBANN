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

    lInputSize = 8
    lLabelSize = 1
    lBatchSize = 16
    lNumEpochs = 40

    'Prepare training data
    Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Concrete", Array(lInputSize, lLabelSize), True, False)
    RandomSplit oFullSet, 0.8, oTrainingSet, oTestSet
    Set oTrainingLoader = DataLoader(oTrainingSet, lBatchSize)
    Set oTestLoader = DataLoader(oTestSet, lBatchSize)
    
    'Setup and train
    Set oModel = Sequential(L2Loss(), SGDM())
    oModel.Add InputNormalizationLayer(oTrainingLoader)
    oModel.Add FullyConnectedLayer(lInputSize, 64)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(64, 16)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(16, 4)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(4, lLabelSize)
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
    PredictInWorksheet = WorksheetFunction.Transpose(Y.ToArray)
End Function

Public Sub WorkingWithTensors()
    Dim a As Tensor
    Dim b As Tensor
    Dim A_() As Double
    Dim B_() As Double
    Dim adblArray() As Double

    'Create an empty tensor A filled with zeros, with shape (2, 3, 4).
    Set a = Zeros(Array(2, 3, 4))

    'Basic properties of A.
    MsgBox a.NumDimensions
    MsgBox a.Size(1)
    MsgBox a.Size(2)
    MsgBox a.Size(3)
    MsgBox a.NumElements
    MsgBox a.Address 'Pointer to the first element

    'Create a tensor A filled with constant values.
    Set a = Ones(Array(2, 3, 4))
    Set a = Full(Array(2, 3, 4), 777)

    'Create a tensor A filled with random values.
    Set a = Uniform(Array(2, 3, 4), 0, 1)
    Set a = Normal(Array(2, 3, 4), 0, 1)
    Set a = Bernoulli(Array(2, 3, 4), 0.5)

    'Fill tensor A with a constant value.
    a.Fill 777

    'Copy tensor A into a new tensor B. (B must be resized to match A's shape.)
    Set b = New Tensor
    b.Resize a.Shape
    b.Copy a

    'Clone tensor A into a new tensor B.
    Set b = a.Clone

    'Use ShapeEquals to check if A's shape matches (2, 3, 4).
    MsgBox a.ShapeEquals(Array(2, 3, 4))

    'Create a different view of A with a new shape (6, 4).
    'This view shares the same underlying data, but has a different layout.
    Set b = a.View(Array(6, 4))

    'Create alias arrays for direct memory access.
    a.CreateAlias A_
    b.CreateAlias B_

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
    a.RemoveAlias A_
    b.RemoveAlias B_

    'Create a flattened view of A with shared underlying data. The new shape is (24).
    Set a = a.Flatten

    'Add singleton dimensions on both sides. The new shape is (1, 24, 1).
    Set a = a.View(Array(1, 24, 1))

    'Reshape A to a 2D tensor (4, 6). Number of elements must remain the same.
    a.Reshape Array(4, 6)

    'Reduce A along dimension 2 using mean reduction. The new shape is (4, 1).
    Set a = a.Reduce(2, rdcMean)

    'Slice A along dimension 1 from index 3 to 4. The new shape is (2, 1).
    Set a = a.Slice(1, 3, 4)

    'Tile A along dimension 2, repeating it 3 times. The new shape is (2, 3).
    Set a = a.Tile(2, 3)

    'Create tensor A from a native VBA array.
    a.FromArray adblArray

    'Copy tensor A to a native VBA array.
    adblArray = a.ToArray

    'Create tensor A from an Excel range.
    a.FromRange ActiveSheet.Range("A1:B3")
End Sub

Public Sub Test()
    Dim a As Tensor
    Static A_() As Double
    
    Set a = New Tensor
    a.Resize Array()
    a.CreateAlias A_
    
    a.RemoveAlias A_
End Sub

