## VBANN

### What is VBANN?
VBANN is a small machine learning framework implemented in VBA, which can be used to set up and train simple neural networks. I'm writing this project from scratch to deepen my understanding of neural networks.<br/>
VBANN is designed to store everything in the same file. Your training data, your model and the framework itself are all contained in the same workbook.<br/>
VBANN is modular and extensible. You can add your own layer classes.<br/>
You can speed it up 6-10x by downloading and linking to [a prebuilt OpenBLAS dll](https://github.com/OpenMathLib/OpenBLAS/releases) inside the [MathFunctions](MathFunctions.bas) module.

### Contributing
Feel free to open issues or submit custom layers to enhance the functionality of VBANN.

### License
This project is licensed under the [Creative Commons Zero v1.0 Universal](LICENSE.txt).

### Examples
```vba
Const MODEL_NAME As String = "MyModel"

Public Sub SetupAndTrain()
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim oTrainingSet As DataLoader
    Dim oTestSet As DataLoader

    lBatchSize = 10
    lNumEpochs = 40

    Set oTrainingSet = DataLoader(ImportDatasetFromWorksheet("ConcreteTrain", 8, 1, True), lBatchSize)
    Set oTestSet = DataLoader(ImportDatasetFromWorksheet("ConcreteTest", 8, 1, True), lBatchSize)

    Set m_oModel = Sequential(L2Loss(), SGDM())
    m_oModel.Add InputNormalizationLayer(oTrainingSet)
    m_oModel.Add FullyConnectedLayer(8, 200)
    m_oModel.Add LeakyReLULayer()
    m_oModel.Add FullyConnectedLayer(200, 100)
    m_oModel.Add LeakyReLULayer()
    m_oModel.Add FullyConnectedLayer(100, 50)
    m_oModel.Add LeakyReLULayer()
    m_oModel.Add FullyConnectedLayer(50, 1)
    m_oModel.Fit oTrainingSet, oTestSet, lNumEpochs

    Serialize MODEL_NAME, m_oModel

    Beep
End Sub

Public Sub ContinueTraining()
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim oTrainingSet As DataLoader
    Dim oTestSet As DataLoader

    lBatchSize = 10
    lNumEpochs = 5

    Set oTrainingSet = DataLoader(ImportDatasetFromWorksheet("ConcreteTrain", 8, 1, True), lBatchSize)
    Set oTestSet = DataLoader(ImportDatasetFromWorksheet("ConcreteTest", 8, 1, True), lBatchSize)

    Set m_oModel = Unserialize(MODEL_NAME)
    m_oModel.Fit oTrainingSet, oTestSet, lNumEpochs
    Serialize MODEL_NAME, m_oModel

    Beep
End Sub

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
    MsgBox A.Address
    
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
    A.FromRange ThisWorkbook.Worksheets("Sheet1").Range("A1:B3")
End Sub
```
