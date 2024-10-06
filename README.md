## VBANN

```vba
Public Sub SetupAndTrain()
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim oTrainingSet As DataLoader
    Dim oTestSet As DataLoader
    Dim oModel As Sequential
    
    lBatchSize = 10
    lNumEpochs = 50
    
    Set oTrainingSet = DataLoader(ImportDatasetFromWorksheet("Train", 8, 1, True), lBatchSize)
    Set oTestSet = DataLoader(ImportDatasetFromWorksheet("Test", 8, 1, True), lBatchSize)
    
    Set oModel = Sequential(L2Loss(), SGDM())
    oModel.Add InputNormalizationLayer(oTrainingSet)
    oModel.Add FullyConnectedLayer(8, 200)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(200, 100)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(100, 50)
    oModel.Add LeakyReLULayer()
    oModel.Add FullyConnectedLayer(50, 1)
    oModel.Fit oTrainingSet, oTestSet, lNumEpochs
    
    Serialize MODEL_NAME, oModel
    
    Beep
End Sub
```

### What is VBANN?
VBANN is a small machine learning framework implemented in VBA, which can be used to set up and train simple neural networks.<br/>
VBANN is designed to store everything in the same file. Your training data, your model and the framework itself are all contained in the same workbook.<br/>
VBANN is modular and extensible. You can add your own layer classes.<br/>
You can speed it up 6-10x by downloading and linking to [a prebuilt OpenBLAS dll](https://github.com/OpenMathLib/OpenBLAS/releases) inside the [MathFunctions](MathFunctions.bas) module.

### Contributing
Feel free to open issues or submit custom layers to enhance the functionality of VBANN.

### License
This project is licensed under the [Creative Commons Zero v1.0 Universal](LICENSE.txt).
