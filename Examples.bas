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
