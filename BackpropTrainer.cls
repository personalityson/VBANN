VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BackpropTrainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Dependencies
Private m_oModel As ILayer
Private m_oLossFunction As ICriterion
Private m_oOptimizer As IOptimizer

Public Property Get Model() As ILayer
    Set Model = m_oModel
End Property

Public Property Get Criterion() As ICriterion
    Set Criterion = m_oLossFunction
End Property

Public Property Get Optimizer() As IOptimizer
    Set Optimizer = m_oOptimizer
End Property

Public Sub Init(ByVal oModel As ILayer, _
                ByVal oLossFunction As ICriterion, _
                ByVal oOptimizer As IOptimizer)
    Const PROCEDURE_NAME As String = "BackpropTrainer.Init"

    If oModel Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ILayer object is required."
    End If
    If oLossFunction Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ICriterion object is required."
    End If
    If oOptimizer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IOptimizer object is required."
    End If
    Set m_oModel = oModel
    Set m_oLossFunction = oLossFunction
    Set m_oOptimizer = oOptimizer
End Sub

Public Sub Fit(ByVal oTrainingLoader As DataLoader, _
               ByVal oTestLoader As DataLoader, _
               ByVal lNumEpochs As Long)
    Const PROCEDURE_NAME As String = "BackpropTrainer.Train"
    Dim i As Long
    Dim dblTrainingLoss As Double
    Dim dblTestLoss As Double

    If oTrainingLoader Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid DataLoader object is required."
    End If
    If oTestLoader Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid DataLoader object is required."
    End If
    If lNumEpochs < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of epochs must be >= 1."
    End If
    For i = 1 To lNumEpochs
        FitOne oTrainingLoader
        dblTrainingLoss = Evaluate(oTrainingLoader)
        dblTestLoss = Evaluate(oTestLoader)
        WriteLog "Log", "Name", TypeName(Me), _
                 "Time", GetUtcTime(), _
                 "Epoch", i, _
                 "Total number of epochs", lNumEpochs, _
                 "Learning rate", m_oOptimizer.LearningRate, _
                 "Training loss", dblTrainingLoss, _
                 "Test loss", dblTestLoss
    Next i
End Sub

Public Function Evaluate(ByVal oDataLoader As DataLoader) As Double
    Const PROCEDURE_NAME As String = "BackpropTrainer.Evaluate"
    Dim lNumSamples As Long
    Dim dblLoss As Double
    Dim X As Tensor
    Dim T As Tensor
    Dim Y As Tensor
    Dim oBatch As TensorDataset

    If oDataLoader Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid DataLoader object is required."
    End If
    With oDataLoader
        .Reset
        Do While .NextBatch(oBatch)
            Set X = oBatch.Tensor(1)
            Set T = oBatch.Tensor(2)
            Set Y = m_oModel.Forward(X)
            dblLoss = dblLoss + m_oLossFunction.Forward(Y, T).Sum
            lNumSamples = lNumSamples + oBatch.NumSamples
            If lNumSamples \ .BatchSize Mod 100 = 0 Then
                Application.StatusBar = "Evaluate progress: " & Format(lNumSamples / .NumSamples, "0.0%")
                DoEvents
            End If
        Loop
    End With
    If lNumSamples > 0 Then
        Evaluate = dblLoss / lNumSamples
    End If
End Function

Private Sub FitOne(ByVal oTrainingLoader As DataLoader)
    Dim lNumSamples As Long
    Dim X As Tensor
    Dim T As Tensor
    Dim oBatch As TensorDataset

    With oTrainingLoader
        .Shuffle
        Do While .NextBatch(oBatch)
            Set X = oBatch.Tensor(1)
            Set T = oBatch.Tensor(2)
            Backpropagate X, T
            lNumSamples = lNumSamples + oBatch.NumSamples
            If lNumSamples \ .BatchSize Mod 100 = 0 Then
                Application.StatusBar = "FitOne progress: " & Format(lNumSamples / .NumSamples, "0.0%")
                DoEvents
            End If
        Loop
    End With
End Sub

Private Sub Backpropagate(ByVal X As Tensor, _
                          ByVal T As Tensor)
    Dim Y As Tensor
    Dim dLdY As Tensor

    Set Y = m_oModel.Forward(X, True)
    Set dLdY = m_oLossFunction.Backward(Y, T)
    m_oModel.Backward dLdY
    m_oOptimizer.Update m_oModel.Learnables
End Sub

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "BackpropTrainer.Save"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteObject m_oModel
        .WriteObject m_oLossFunction
        .WriteObject m_oOptimizer
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "BackpropTrainer.Load"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        Set m_oModel = .ReadObject()
        Set m_oLossFunction = .ReadObject()
        Set m_oOptimizer = .ReadObject()
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Sub ISerializable_Save(ByVal oSerializer As Object)
        Save oSerializer
    End Sub

    Private Sub ISerializable_Load(ByVal oSerializer As Object)
        Load oSerializer
    End Sub
#Else
    Private Sub ISerializable_Save(ByVal oSerializer As Serializer)
        Save oSerializer
    End Sub

    Private Sub ISerializable_Load(ByVal oSerializer As Serializer)
        Load oSerializer
    End Sub
#End If

