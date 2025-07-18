VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sequential"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Dependencies
Private m_oCriterion As ICriterion
Private m_oOptimizer As IOptimizer

'State
Private m_lNumLayers As Long
Private m_aoLayers() As ILayer

Public Property Get Criterion() As ICriterion
    Set Criterion = m_oCriterion
End Property

Public Property Get Optimizer() As IOptimizer
    Set Optimizer = m_oOptimizer
End Property

Public Sub Init(ByVal oCriterion As ICriterion, _
                ByVal oOptimizer As IOptimizer)
    Const PROCEDURE_NAME As String = "Sequential.Init"

    If oCriterion Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ICriterion object is required."
    End If
    If oOptimizer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IOptimizer object is required."
    End If
    Clear
    Set m_oCriterion = oCriterion
    Set m_oOptimizer = oOptimizer
End Sub

Public Sub Add(ByVal oLayer As ILayer)
    Const PROCEDURE_NAME As String = "Sequential.Add"

    If oLayer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ILayer object is required."
    End If
    m_lNumLayers = m_lNumLayers + 1
    ReDim Preserve m_aoLayers(1 To m_lNumLayers)
    Set m_aoLayers(m_lNumLayers) = oLayer
End Sub

Public Sub Fit(ByVal oTrainingLoader As DataLoader, _
               ByVal oTestLoader As DataLoader, _
               ByVal lNumEpochs As Long)
    Const PROCEDURE_NAME As String = "Sequential.Fit"
    Dim i As Long
    Dim dblLoss As Double

    If oTrainingLoader Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid DataLoader object is required."
    End If
    If oTestLoader Is Nothing Then
        Set oTestLoader = oTrainingLoader
    End If
    If lNumEpochs < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of epochs must be >= 1."
    End If
    For i = 1 To lNumEpochs
        FitOne oTrainingLoader
        dblLoss = Loss(oTestLoader)
        LogToWorksheet "Log", "Time", GetUtcTime(), "Epoch", i, "Total number of epochs", lNumEpochs, "Learning rate", m_oOptimizer.LearningRate, "Loss", dblLoss
    Next i
End Sub

Public Function Predict(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "Sequential.Predict"
    Dim i As Long
    
    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If X.NumDimensions < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor must have at least one dimension."
    End If
    For i = 1 To m_lNumLayers
        Set X = m_aoLayers(i).Forward(X, bIsTraining)
    Next i
    Set Predict = X
End Function

Private Sub Backpropagate(ByVal X As Tensor, _
                          ByVal T As Tensor)
    Dim i As Long
    Dim Y As Tensor
    Dim dLdY As Tensor
    
    Set Y = Predict(X, True)
    Set dLdY = m_oCriterion.Gradient(Y, T)
    For i = m_lNumLayers To 1 Step -1
        Set dLdY = m_aoLayers(i).Backward(dLdY)
        UpdateParameters m_aoLayers(i).LearnableParameters
    Next i
End Sub

Public Function Loss(ByVal oTestLoader As DataLoader) As Double
    Const PROCEDURE_NAME As String = "Sequential.Loss"
    Dim lNumSamples As Long
    Dim dblLoss As Double
    Dim X As Tensor
    Dim T As Tensor
    Dim Y As Tensor
    Dim oBatch As TensorDataset

    If oTestLoader Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid DataLoader object is required."
    End If
    With oTestLoader
        .Reset
        Do While .Fetch(oBatch)
            Set X = oBatch.Item(1)
            Set T = oBatch.Item(2)
            Set Y = Predict(X)
            dblLoss = dblLoss + m_oCriterion.Loss(Y, T).Sum
            lNumSamples = lNumSamples + oBatch.NumSamples
        Loop
    End With
    If lNumSamples > 0 Then
        Loss = dblLoss / lNumSamples
    End If
End Function

Private Sub FitOne(ByVal oTrainingLoader As DataLoader)
    Dim X As Tensor
    Dim T As Tensor
    Dim oBatch As TensorDataset

    With oTrainingLoader
        .Shuffle
        Do While .Fetch(oBatch)
            Set X = oBatch.Item(1)
            Set T = oBatch.Item(2)
            Backpropagate X, T
        Loop
    End With
End Sub

Private Sub UpdateParameters(ByVal vLearnableParameters As Variant)
    Const PROCEDURE_NAME As String = "Sequential.UpdateParameters"
    Dim W As Variant

    For Each W In vLearnableParameters
        If Not IsObject(W) Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
        End If
        If W Is Nothing Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
        End If
        If Not TypeOf W Is Parameter Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
        End If
        m_oOptimizer.Update W
    Next W
End Sub

Private Sub Clear()
    Set m_oCriterion = Nothing
    Set m_oOptimizer = Nothing
    m_lNumLayers = 0
    Erase m_aoLayers
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Sequential.Save"
    Dim i As Long

    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteObject m_oCriterion
        .WriteObject m_oOptimizer
        .WriteLong m_lNumLayers
        For i = 1 To m_lNumLayers
            .WriteObject m_aoLayers(i)
        Next i
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Sequential.Load"
    Dim i As Long
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oWorksheet
        Set m_oCriterion = .ReadObject()
        Set m_oOptimizer = .ReadObject()
        m_lNumLayers = .ReadLong()
        If m_lNumLayers > 0 Then
            ReDim m_aoLayers(1 To m_lNumLayers)
            For i = 1 To m_lNumLayers
                Set m_aoLayers(i) = .ReadObject()
            Next i
        End If
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Sub ISerializable_Save(ByVal oWorksheet As Object)
        Save oWorksheet
    End Sub

    Private Sub ISerializable_Load(ByVal oWorksheet As Object)
        Load oWorksheet
    End Sub
#Else
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub

    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If
