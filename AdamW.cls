VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdamW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IOptimizer
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

Private Const KEY_STEP As String = "Step"
Private Const KEY_1ST_MOMENT As String = "Mean"
Private Const KEY_2ND_MOMENT As String = "Variance"

'Hyperparameters
Private m_dblLearningRate As Double
Private m_dblBeta1 As Double
Private m_dblBeta2 As Double
Private m_dblEpsilon As Double
Private m_dblWeightDecay As Double
Private m_dblGradientThreshold As Double

'State
Private m_oStates As Object

Public Property Get LearningRate() As Double
    LearningRate = m_dblLearningRate
End Property

Public Property Let LearningRate(ByVal dblLearningRate As Double)
    Const PROCEDURE_NAME As String = "AdamW.LearningRate"
    
    If dblLearningRate <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Learning rate must be > 0."
    End If
    m_dblLearningRate = dblLearningRate
End Property

Public Sub Init(Optional ByVal dblLearningRate As Double = 0.001, _
                Optional ByVal dblBeta1 As Double = 0.9, _
                Optional ByVal dblBeta2 As Double = 0.999, _
                Optional ByVal dblEpsilon As Double = 0.00000001, _
                Optional ByVal dblWeightDecay As Double = 0.01, _
                Optional ByVal dblGradientThreshold As Double = DOUBLE_MAX_ABS)
    Const PROCEDURE_NAME As String = "AdamW.Init"
    
    If dblLearningRate <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Learning rate must be > 0."
    End If
    If dblBeta1 < 0 Or dblBeta1 >= 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Beta1 must be >= 0 and < 1."
    End If
    If dblBeta2 < 0 Or dblBeta2 >= 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Beta2 must be >= 0 and < 1."
    End If
    If dblEpsilon <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Epsilon must be > 0."
    End If
    If dblWeightDecay < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Weight decay must be >= 0."
    End If
    If dblGradientThreshold <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient threshold must be > 0."
    End If
    Clear
    m_dblLearningRate = dblLearningRate
    m_dblBeta1 = dblBeta1
    m_dblBeta2 = dblBeta2
    m_dblEpsilon = dblEpsilon
    m_dblWeightDecay = dblWeightDecay
    m_dblGradientThreshold = dblGradientThreshold
End Sub

Public Sub Update(ByVal vParameters As Variant)
    Const PROCEDURE_NAME As String = "AdamW.Update"
    Dim W As Variant
    
    For Each W In EnsureIterable(vParameters)
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

'Adam with decoupled weight decay
'm = beta1 * m + (1 - beta1) * (dW / NumAccumulatedSamples)
'v = beta2 * v + (1 - beta2) * (dW / NumAccumulatedSamples)^2
'alpha = lr * Sqr(1 - beta2^t) / (1 - beta1^t)
'W = (1 - lr * wd) * W - alpha * m / (Sqr(v) + epsilon)
Private Sub ApplyGradients(ByVal W As Parameter)
    Dim lStep As Long
    Dim dblLearningRate As Double
    Dim dblAlpha As Double
    Dim dblWeightDecay As Double
    Dim dblGradientScale As Double
    Dim m As Tensor
    Dim v As Tensor

    If W.NumAccumulatedSamples < 1 Then
        Exit Sub
    End If
    GetStates W, lStep, m, v
    lStep = lStep + 1
    dblLearningRate = W.LearningRateScale * m_dblLearningRate
    dblAlpha = dblLearningRate * Sqr(1 - m_dblBeta2 ^ lStep) / (1 - m_dblBeta1 ^ lStep)
    dblWeightDecay = W.WeightDecayScale * m_dblWeightDecay
    If W.UseGradientClipping And m_dblGradientThreshold < DOUBLE_MAX_ABS Then
        dblGradientScale = m_dblGradientThreshold / MaxDbl(m_dblGradientThreshold, VecNorm2(W.Gradient) / W.NumAccumulatedSamples)
    Else
        dblGradientScale = 1
    End If
    VecLinComb_I m_dblBeta1, m, (1 - m_dblBeta1) * dblGradientScale / W.NumAccumulatedSamples, W.Gradient
    VecLinComb_I m_dblBeta2, v, (1 - m_dblBeta2) * dblGradientScale * dblGradientScale / W.NumAccumulatedSamples / W.NumAccumulatedSamples, VecPow2(W.Gradient)
    VecLinComb_I 1 - dblLearningRate * dblWeightDecay, W.Tensor, -dblAlpha, VecDivRms(m, v, 0, m_dblEpsilon)
    W.ResetGradient
    PutStates W, lStep, m, v
End Sub

Private Sub GetStates(ByVal W As Parameter, _
                      ByRef lStep As Long, _
                      ByRef m As Tensor, _
                      ByRef v As Tensor)
    With m_oStates
        If Not .Exists(W) Then
            PutStates W, 0&, Zeros(W.Tensor.Shape), Zeros(W.Tensor.Shape)
        End If
        With .Item(W)
            lStep = .Item(KEY_STEP)
            Set m = .Item(KEY_1ST_MOMENT)
            Set v = .Item(KEY_2ND_MOMENT)
        End With
    End With
End Sub

Private Sub PutStates(ByVal W As Parameter, _
                      ByVal lStep As Long, _
                      ByVal m As Tensor, _
                      ByVal v As Tensor)
    With m_oStates
        If Not .Exists(W) Then
            Set .Item(W) = CreateObject("Scripting.Dictionary")
        End If
        With .Item(W)
            .Item(KEY_STEP) = lStep
            Set .Item(KEY_1ST_MOMENT) = m
            Set .Item(KEY_2ND_MOMENT) = v
        End With
    End With
End Sub

Private Sub Clear()
    m_dblLearningRate = 0
    m_dblBeta1 = 0
    m_dblBeta2 = 0
    m_dblEpsilon = 0
    m_dblWeightDecay = 0
    m_dblGradientThreshold = 0
    Set m_oStates = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "AdamW.Save"
    Dim W As Variant
    Dim lStep As Long
    Dim m As Tensor
    Dim v As Tensor
    
    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteDouble m_dblLearningRate
        .WriteDouble m_dblBeta1
        .WriteDouble m_dblBeta2
        .WriteDouble m_dblEpsilon
        .WriteDouble m_dblWeightDecay
        .WriteDouble m_dblGradientThreshold
        .WriteLong m_oStates.Count
        For Each W In m_oStates.Keys
            GetStates W, lStep, m, v
            .WriteObject W
            .WriteLong lStep
            .WriteObject m
            .WriteObject v
        Next W
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "AdamW.Load"
    Dim i As Long
    Dim lNumStates As Long
    Dim W As Parameter
    Dim lStep As Long
    Dim m As Tensor
    Dim v As Tensor
    
    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oSerializer
        m_dblLearningRate = .ReadDouble()
        m_dblBeta1 = .ReadDouble()
        m_dblBeta2 = .ReadDouble()
        m_dblEpsilon = .ReadDouble()
        m_dblWeightDecay = .ReadDouble()
        m_dblGradientThreshold = .ReadDouble()
        lNumStates = .ReadLong()
        For i = 1 To lNumStates
            Set W = .ReadObject()
            lStep = .ReadLong()
            Set m = .ReadObject()
            Set v = .ReadObject()
            PutStates W, lStep, m, v
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
    
    Private Sub ISerializable_Save(ByVal oSerializer As Object)
        Save oSerializer
    End Sub
    
    Private Sub ISerializable_Load(ByVal oSerializer As Object)
        Load oSerializer
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
    
    Private Sub ISerializable_Save(ByVal oSerializer As Serializer)
        Save oSerializer
    End Sub
    
    Private Sub ISerializable_Load(ByVal oSerializer As Serializer)
        Load oSerializer
    End Sub
#End If

