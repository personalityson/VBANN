VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SGDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IOptimizer
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

Private Const KEY_VELOCITY As String = "Velocity"

'Hyperparameters
Private m_dblLearningRate As Double
Private m_dblMomentum As Double
Private m_dblWeightDecay As Double

'State
Private m_oStates As Object

Public Property Get LearningRate() As Double
    LearningRate = m_dblLearningRate
End Property

Public Property Let LearningRate(ByVal dblLearningRate As Double)
    Const PROCEDURE_NAME As String = "SGDM.LearningRate"
    
    If dblLearningRate <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Learning rate must be > 0."
    End If
    m_dblLearningRate = dblLearningRate
End Property

Public Sub Init(Optional ByVal dblLearningRate As Double = 0.001, _
                Optional ByVal dblMomentum As Double = 0.9, _
                Optional ByVal dblWeightDecay As Double = 0.01)
    Const PROCEDURE_NAME As String = "SGDM.Init"
    
    If dblLearningRate <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Learning rate must be > 0."
    End If
    If dblMomentum < 0 Or dblMomentum >= 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Momentum must be >= 0 and < 1."
    End If
    If dblWeightDecay < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Weight decay must be >= 0."
    End If
    Clear
    m_dblLearningRate = dblLearningRate
    m_dblMomentum = dblMomentum
    m_dblWeightDecay = dblWeightDecay
End Sub

Public Sub Update(ByVal vParameters As Variant)
    Const PROCEDURE_NAME As String = "Adam.Update"
    Dim W As Variant
    
    Select Case GetRank(vParameters)
        Case -1
            vParameters = Array(vParameters)
        Case 0
            Exit Sub
        Case 1
            'Continue
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Expected a single parameter, an uninitialized array, or a one-dimensional array of parameters."
    End Select
    For Each W In vParameters
        If Not IsObject(W) Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
        End If
        If W Is Nothing Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
        End If
        If Not TypeOf W Is Parameter Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
        End If
        ApplyGradients W
    Next W
End Sub

'v = momentum * v + (1 - momentum) * (dW / NumUpdates)
'W = W - lr * (v + wd * W)
Private Sub ApplyGradients(ByVal W As Parameter)
    Dim dblLearningRate As Double
    Dim dblWeightDecay As Double
    Dim v As Tensor
    
    If W.NumUpdates < 1 Then
        Exit Sub
    End If
    GetStates W, v
    dblLearningRate = W.LearningRateScale * m_dblLearningRate
    dblWeightDecay = W.WeightDecayScale * m_dblWeightDecay
    VecLinComb_I m_dblMomentum, v, (1 - m_dblMomentum) / W.NumUpdates, W.Gradient
    VecLinComb_I 1 - dblLearningRate * dblWeightDecay, W.Variable, -dblLearningRate, v
    W.ResetGradient
    'PutStates W, v
End Sub

Private Sub GetStates(ByVal W As Parameter, _
                      ByRef v As Tensor)
    With m_oStates
        If Not .Exists(W) Then
            PutStates W, Zeros(W.Variable.Shape)
        End If
        With .Item(W)
            Set v = .Item(KEY_VELOCITY)
        End With
    End With
End Sub

Private Sub PutStates(ByVal W As Parameter, _
                      ByVal v As Tensor)
    With m_oStates
        If Not .Exists(W) Then
            Set .Item(W) = CreateObject("Scripting.Dictionary")
        End If
        With .Item(W)
            Set .Item(KEY_VELOCITY) = v
        End With
    End With
End Sub

Private Sub Clear()
    m_dblLearningRate = 0
    m_dblMomentum = 0
    m_dblWeightDecay = 0
    Set m_oStates = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "SGDM.Save"
    Dim W As Variant
    Dim v As Tensor
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteDouble m_dblLearningRate
        .WriteDouble m_dblMomentum
        .WriteDouble m_dblWeightDecay
        .WriteLong m_oStates.Count
        For Each W In m_oStates.Keys
            GetStates W, v
            .WriteObject W
            .WriteObject v
        Next W
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "SGDM.Load"
    Dim lNumStates As Long
    Dim i As Long
    Dim W As Variant
    Dim v As Tensor
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oWorksheet
        m_dblLearningRate = .ReadDouble()
        m_dblMomentum = .ReadDouble()
        m_dblWeightDecay = .ReadDouble()
        lNumStates = .ReadLong()
        For i = 1 To lNumStates
            Set W = .ReadObject()
            Set v = .ReadObject()
            PutStates W, v
        Next i
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get IOptimizer_LearningRate() As Double
        IOptimizer_LearningRate = Me.LearningRate
    End Property
    
    Private Property Let IOptimizer_LearningRate(ByVal dblLearningRate As Double)
        Me.LearningRate = dblLearningRate
    End Property
    
    Private Sub IOptimizer_Update(ByVal vParameters As Variant)
        Me.Update vParameters
    End Sub
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Object)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Object)
        Load oWorksheet
    End Sub
#Else
    Private Property Get IOptimizer_LearningRate() As Double
        IOptimizer_LearningRate = Me.LearningRate
    End Property
    
    Private Property Let IOptimizer_LearningRate(ByVal dblLearningRate As Double)
        Me.LearningRate = dblLearningRate
    End Property
    
    Private Sub IOptimizer_Update(ByVal W As Parameter)
        Me.Update W
    End Sub
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If

