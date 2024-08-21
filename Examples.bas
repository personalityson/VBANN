Attribute VB_Name = "Examples"
Option Explicit

Const MODEL_NAME As String = "MyModel"

Private m_oModel As Sequential

Public Sub SetupAndTrain()
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim oTrainingSet As DataLoader
    Dim oTestSet As DataLoader
    Dim lStart As Long
    Dim lEnd As Long

    Randomize 777
    
    lBatchSize = 10
    lNumEpochs = 5
    
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
    
    lStart = GetTickCount()
    m_oModel.Fit oTrainingSet, oTestSet, lNumEpochs
    lEnd = GetTickCount()

    MsgBox (lEnd - lStart) / 1000
    
    'Serialize MODEL_NAME, m_oModel
    
    Beep
End Sub

Public Sub ContinueTraining()
    Dim lBatchSize As Long
    Dim lNumEpochs As Long
    Dim oTrainingSet As DataLoader
    Dim oTestSet As DataLoader
    
    lBatchSize = 10
    lNumEpochs = 50
    
    Set oTrainingSet = DataLoader(ImportDatasetFromWorksheet("ConcreteTrain", 8, 1, True), lBatchSize)
    Set oTestSet = DataLoader(ImportDatasetFromWorksheet("ConcreteTest", 8, 1, True), lBatchSize)
    
    Set m_oModel = Unserialize(MODEL_NAME)
    m_oModel.Fit oTrainingSet, oTestSet, lNumEpochs
    Serialize MODEL_NAME, m_oModel
    
    Beep
End Sub

Public Function PredictInWorksheet(ByVal rngInput As Range) As Double()
    Dim X As Tensor
    Dim Y As Tensor
    
    If m_oModel Is Nothing Then
        Set m_oModel = Unserialize(MODEL_NAME)
    End If
    Set X = TensorFromRange(rngInput, True)
    Set Y = m_oModel.Predict(X)
    PredictInWorksheet = Y.ToArray
End Function

Public Sub TestEx()
    Dim A As Tensor
    Dim B As Tensor
    Dim C As Tensor
    Dim lStart As Long
    Dim lEnd As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim X_() As Double
    Dim Y_() As Double
    Dim lBatchSize As Long
    Dim dblMean As Double
    Dim dblVariance As Double
    Dim dblTemp As Double
    
    lBatchSize = 10
    
    Randomize 777

    Set A = Uniform(Array(10, 1000, 30))
'    Set B = Uniform(Array(5, 10))
'    Set C = Zeros(Array(10000))
'    'MatMul False, False, 1, A, B, 1, C
'    MsgBox C.NumDimensions

    lStart = GetTickCount()
    For k = 1 To 100
        Set C = A.ReduceSum(2)
    Next k
    lEnd = GetTickCount()
    
    MsgBox (lEnd - lStart) / 1000
    
    'MsgBox C.ToArray()(1, 1)

    'MsgBox X.ToArray()(1, 1)
 '   Set Y = X.Sum(2)
'    X.CreateAlias X_
'    Y.CreateAlias Y_
'    MsgBox X_(1, 1, 1)
'    MsgBox Y_(1, 1, 1)
'    X.RemoveAlias X_
'    Y.RemoveAlias Y_

    'MsgBox dblTemp

    'MsgBox (lEnd - lStart) / 1000
    
    'MsgBox C.ToArray()(1, 1, 1)
    

    
End Sub
