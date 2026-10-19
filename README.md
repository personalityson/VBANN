## VBANN

### What is VBANN?
- VBANN is a small PyTorch-style machine learning framework implemented in pure VBA, which can be used to set up neural networks and gradient-boosted trees inside Excel. I'm writing this project to deepen my understanding of neural networks.<br/>
- VBANN is designed to store everything in the same file. No installs, no add-ins, no admin rights needed -- the whole stack, including data, model, trainer and optimizer state, lives in one .xlsm file. Models, trainers and optimizer state are saved to a worksheet, so you can reload them and continue training.<br/>
- VBANN is modular and extensible. You can add your own layer classes.<br/>
- VBANN can be sped up 6-10x by downloading and linking to [a prebuilt OpenBLAS dll](https://github.com/OpenMathLib/OpenBLAS/releases) inside the [TensorOps](TensorOps.bas) module.

### License
This project is licensed under the [Creative Commons Zero v1.0 Universal](LICENSE.txt).

### Quick example
Gradient-boosted trees:
```vba
Set oFullSet = ImportDatasetFromWorksheet(ThisWorkbook, "Concrete", Array(8, 1), bHasHeaders:=True)
SplitDataset oFullSet, 0.8, oTrainingSet, oValidationSet, bRandomize:=True

Set oModel = XGBoost(L2Loss(), dblLearningRate:=0.1, lMaxDepth:=6)
oModel.Fit oTrainingSet, lNumRounds:=40, bVerbose:=True, oValidationSet:=oValidationSet

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
oTrainer.Fit DataLoader(oTrainingSet, lBatchSize:=16), lNumEpochs:=20, bVerbose:=True
```
Check out full [Examples.bas](Examples.bas).
