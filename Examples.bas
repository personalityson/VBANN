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
    m_oModel.Add InputNormalizationLayer(oTrainingSet)
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

    Serialize MODEL_NAME, m_oModel

    Beep
End Sub

'Public Sub SetupAndTrain()
'    Dim lBatchSize As Long
'    Dim lNumEpochs As Long
'    Dim oTrainingSet As DataLoader
'    Dim oTestSet As DataLoader
'    Dim lStart As Long
'    Dim lEnd As Long
'
'    Randomize 777
'
'    lBatchSize = 32
'    lNumEpochs = 10
'
'    Set oTrainingSet = DataLoader(ImportDatasetFromCsv("C:\Users\hello\OneDrive\Documents\VBANN\mnist_train.csv", 784, 10, True), lBatchSize)
'    Set oTestSet = DataLoader(ImportDatasetFromCsv("C:\Users\hello\OneDrive\Documents\VBANN\mnist_test.csv", 784, 10, True), lBatchSize)
'
'    Set m_oModel = Sequential(CCELoss(), SGDM(0.001))
'    'm_oModel.Add InputNormalizationLayer(oTrainingSet)
'    m_oModel.Add FullyConnectedLayer(784, 128)
'    m_oModel.Add LeakyReLULayer()
'    m_oModel.Add FullyConnectedLayer(128, 64)
'    m_oModel.Add SigmoidLayer()
'    m_oModel.Add FullyConnectedLayer(64, 32)
'    m_oModel.Add SigmoidLayer()
'    m_oModel.Add FullyConnectedLayer(32, 10)
'    m_oModel.Add SoftmaxLayer()
'
'    lStart = GetTickCount()
'    m_oModel.Fit oTrainingSet, oTestSet, lNumEpochs
'    lEnd = GetTickCount()
'
'    MsgBox (lEnd - lStart) / 1000
'
'    Serialize MODEL_NAME, m_oModel
'
'    Beep
'End Sub

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

'Public Sub ContinueTraining()
'    Dim lBatchSize As Long
'    Dim lNumEpochs As Long
'    Dim oTrainingSet As DataLoader
'    Dim oTestSet As DataLoader
'
'    lBatchSize = 32
'    lNumEpochs = 10
'
'    Set oTrainingSet = DataLoader(ImportDatasetFromCsv("C:\Users\hello\OneDrive\Documents\VBANN\mnist_train.csv", 784, 10, True), lBatchSize)
'    Set oTestSet = DataLoader(ImportDatasetFromCsv("C:\Users\hello\OneDrive\Documents\VBANN\mnist_test.csv", 784, 10, True), lBatchSize)
'
'    Set m_oModel = Unserialize(MODEL_NAME)
'
'    m_oModel.Fit oTrainingSet, oTestSet, lNumEpochs
'
'    Serialize MODEL_NAME, m_oModel
'
'    Beep
'End Sub


Public Function PredictInWorksheet(ByVal oInput As Range) As Double()
    Dim X As Tensor
    Dim Y As Tensor
    
    If m_oModel Is Nothing Then
        Set m_oModel = Unserialize(MODEL_NAME)
    End If
    Set X = TensorFromRange(oInput, True)
    Set Y = m_oModel.Predict(X)
    PredictInWorksheet = Y.ToArray
End Function

Public Sub TestEx()
    Dim A As Tensor
    Dim B As Tensor
    Dim C As Tensor
    Dim D As Tensor
    Dim lStart As Long
    Dim lEnd As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim A_() As Double
    Dim B_() As Double
    Dim C_() As Double
    Dim lBatchSize As Long
    Dim dblMean As Double
    Dim dblVariance As Double
    Dim dblTemp As Double
    Dim v As Variant
    Dim oLayer As ILayer
    
    
    lBatchSize = 10
    
    Randomize 777
    
    
    Set oLayer = SoftmaxLayer()
    
    ReDim A_(1 To 2, 1 To 2)
    A_(1, 1) = 2
    A_(2, 1) = 1
    A_(1, 2) = 0.5
    A_(2, 2) = 0.5
    Set A = New Tensor
    A.FromArray A_
    
    ReDim B_(1 To 2, 1 To 2)
    B_(1, 1) = 0.1
    B_(2, 1) = -0.2
    B_(1, 2) = -0.3
    B_(2, 2) = 0.4
    Set B = New Tensor
    B.FromArray B_
    
    Set C = oLayer.Forward(A)
    
    MsgBox "ok"
    
    Set C = oLayer.Backward(B)
    
    MsgBox "ok"
'    Set A = Bernoulli(Array(3, 4, 5), 0.1)
'    MsgBox "ok"
'    Set B = Zeros(Array(3, 4, 5))
'    Set C = A.Flatten
'    Set D = B.Flatten
    
'    Set C = A.Tile(3, 2)

'    Set oLayer = New Parameter
'    Serialize "jj", oLayer
'
'    Set oLayer = Unserialize("jj")
    
'    A.Flatten.CreateAlias A_
'    For i = 1 To A.NumElements
'        A_(i) = i
'    Next i
'    A.Flatten.RemoveAlias A_
'
'    Serialize "test", A
'    Set B = Uniform(Array(5, 10))
'    Set C = Zeros(Array(10000))
'    'MatMul False, False, 1, A, B, 1, C
'    MsgBox C.NumDimensions

    'MsgBox a.CallMethodDynamically("AddNumbers", 4, 3)

'    lStart = GetTickCount()
'    For k = 1 To 100
'        Set A = Bernoulli(Array(1000, 1000))
'    Next k
'    lEnd = GetTickCount()
'
'    MsgBox (lEnd - lStart) / 1000
    
'    ReDim A_(1 To 10, 1 To 10)
'    ReDim B_(1 To 10, 1 To 10)
'    v = WorksheetFunction.MMult(A_, B_)
'    MsgBox v(1, 1)
'    MsgBox C.ToArray()(1, 1)
'
'    MsgBox X.ToArray()(1, 1)
'    Set Y = X.Sum(2)
'    A.CreateAlias A_
'    C.CreateAlias C_
'    MsgBox A_(1, 1, 1)
'    MsgBox C_(1, 1, 1)
'    A.RemoveAlias A_
'    C.RemoveAlias C_
'
'    MsgBox dblTemp

    'MsgBox (lEnd - lStart) / 1000
    
    'MsgBox C.ToArray()(1, 1, 1)
    

    
End Sub

