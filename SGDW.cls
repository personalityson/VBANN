VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SGDW"
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
Private m_dblGradientThreshold As Double

'State
Private m_oStates As Object

Public Property Get LearningRate() As Double
    LearningRate = m_dblLearningRate
End Property

Public Property Let LearningRate(ByVal dblLearningRate As Double)
    Const PROCEDURE_NAME As String = "SGDW.LearningRate"
    
    If dblLearningRate <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Learning rate must be > 0."
    End If
    m_dblLearningRate = dblLearningRate
End Property

Public Sub Init(Optional ByVal dblLearningRate As Double = 0.001, _
                Optional ByVal dblMomentum As Double = 0.9, _
                Optional ByVal dblWeightDecay As Double = 0.0001, _
                Optional ByVal dblGradientThreshold As Double = DOUBLE_MAX_ABS)
    Const PROCEDURE_NAME As String = "SGDW.Init"
    
    If dblLearningRate <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "LearningRate must be > 0."
    End If
    If dblMomentum < 0 Or dblMomentum >= 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Momentum must be >= 0 and < 1."
    End If
    If dblWeightDecay < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "WeightDecay must be >= 0."
    End If
    If dblGradientThreshold <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "GradientThreshold must be > 0."
    End If
    Clear
    m_dblLearningRate = dblLearningRate
    m_dblMomentum = dblMomentum
    m_dblWeightDecay = dblWeightDecay
    m_dblGradientThreshold = dblGradientThreshold
End Sub

Public Sub Update(ByVal vParameters As Variant)
    Const PROCEDURE_NAME As String = "SGDW.Update"
    Dim W As Variant
    
    vParameters = EnsureArray(vParameters)
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

'SGD with momentum and decoupled weight decay
'v = momentum * v + dW / NumAccumulatedSamples
'W = (1 - lr * wd) * W - lr * v
Private Sub ApplyGradients(ByVal W As Parameter)
    Dim dblLearningRate As Double
    Dim dblWeightDecay As Double
    Dim dblGradientScale As Double
    Dim v As Tensor
    
    If W.NumAccumulatedSamples < 1 Then
        Exit Sub
    End If
    GetStates W, v
    dblLearningRate = W.LearningRateScale * m_dblLearningRate
    dblWeightDecay = W.DecayScale * m_dblWeightDecay
    If W.UseGradientClipping And m_dblGradientThreshold < DOUBLE_MAX_ABS Then
        dblGradientScale = m_dblGradientThreshold / MaxDbl(m_dblGradientThreshold, VecNorm2(W.Gradient) / W.NumAccumulatedSamples)
    Else
        dblGradientScale = 1
    End If
    VecLinComb_I m_dblMomentum, v, dblGradientScale / W.NumAccumulatedSamples, W.Gradient
    VecLinComb_I 1 - dblLearningRate * dblWeightDecay, W.Tensor, -dblLearningRate, v
    W.ResetGradient
    'PutStates W, v
End Sub

Private Sub GetStates(ByVal W As Parameter, _
                      ByRef v As Tensor)
    With m_oStates
        If Not .Exists(W) Then
            PutStates W, Zeros(W.Tensor.Shape)
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
    m_dblGradientThreshold = 0
    Set m_oStates = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "SGDW.Save"
    Dim W As Variant
    Dim v As Tensor
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteDouble m_dblLearningRate
        .WriteDouble m_dblMomentum
        .WriteDouble m_dblWeightDecay
        .WriteDouble m_dblGradientThreshold
        .WriteLong m_oStates.Count
        For Each W In m_oStates.Keys
            GetStates W, v
            .WriteObject W
            .WriteObject v
        Next W
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "SGDW.Load"
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
        m_dblGradientThreshold = .ReadDouble()
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
    
    Private Sub IOptimizer_Update(ByVal vParameters As Variant)
        Me.Update vParameters
    End Sub
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If

