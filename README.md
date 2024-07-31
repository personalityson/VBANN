## VBANN

```vba
Const MODEL_NAME As String = "MyModel"

Private m_oModel As Sequential

Public Sub SetupAndTrain()
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim oTrainingSet As DataLoader
    Dim oTestSet As DataLoader
    
    VerifyOpenBlasLibrary
    
    lBatchSize = 10
    lNumEpochs = 50
    
    Set oTrainingSet = DataLoader(ImportDatasetFromWorksheet("ConcreteTrain", 8, 1, True), lBatchSize)
    Set oTestSet = DataLoader(ImportDatasetFromWorksheet("ConcreteTest", 8, 1, True), lBatchSize)
    
    Set m_oModel = Sequential(L2Loss(), SGDM())
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
```

### What is VBANN?
VBANN is a small machine learning framework implemented in VBA, which can be used to set up and train simple neural networks.<br/>
VBANN is designed to store everything in the same file. Your training data, your model and the framework itself are all contained in the same workbook.<br/>
VBANN is modular and extensible. You can add your own layer classes.<br/>
You can speed it up 6-10x by downloading and linking to [a prebuilt OpenBLAS dll](https://github.com/OpenMathLib/OpenBLAS/releases) inside the [BlasFunctions](BlasFunctions.bas) module.

### Why do people use VBA?
VBA runs on locked-down mandatory corporate Windows laptops with no software installation privileges, firewalled networking and USB ports disabled.<br/>
Very often it is the only solution that can be implemented without approval from the IT department.

### Contributing
Feel free to open issues or submit pull requests to enhance the functionality of VBANN.

### License
This project is licensed under the [MIT License](LICENSE.txt).
