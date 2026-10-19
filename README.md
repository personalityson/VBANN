## VBANN

### What is VBANN?
VBANN is a small machine learning framework implemented in VBA, which can be used to set up and train simple neural networks. I'm writing this project from scratch to deepen my understanding of neural networks.<br/>
VBANN is designed to store everything in the same file. Your training data, your model and the framework itself are all contained in the same workbook.<br/>
VBANN is modular and extensible. You can add your own layer classes.<br/>
VBANN can be sped up 6-10x by downloading and linking to [a prebuilt OpenBLAS dll](https://github.com/OpenMathLib/OpenBLAS/releases) inside the [TensorOps](TensorOps.bas) module.

### License
This project is licensed under the [Creative Commons Zero v1.0 Universal](LICENSE.txt).

### Quick example

Gradient-boosted trees:
```vba
Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Concrete", Array(8, 1), bHasHeaders:=True)
SplitDataset oFullSet, 0.8, oTrainingSet, oValidationSet, bRandomize:=True

Set oModel = XGBoost(L2Loss(), dblLearningRate:=0.1, lMaxDepth:=6)
oModel.Fit oTrainingSet, 40, True, oValidationSet

MsgBox oModel.Evaluate(oValidationSet)
Serialize "MyXGBoostModel", oModel
```

A neural network:
```vba
Set oModel = Sequential()
oModel.Add NormalizationLayer(oTrainingSet, lIndex:=1)
oModel.Add FullyConnectedLayer(8, 32)
oModel.Add LeakyReLULayer()
oModel.Add FullyConnectedLayer(32, 1)
oModel.Add NormalizationLayer(oTrainingSet, lIndex:=2, bInvert:=True)

Set oTrainer = BackpropTrainer(oModel, L2Loss(), AdamW())
oTrainer.Fit DataLoader(oTrainingSet, 16), 20, True
```

