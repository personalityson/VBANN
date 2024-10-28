VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Adam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IOptimizer
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

Private Const KEY_STEP As String = "Step"
Private Const KEY_MEAN As String = "Mean"
Private Const KEY_VARIANCE As String = "Variance"

Private m_dblLearningRate As Double
Private m_dblBeta1 As Double
Private m_dblBeta2 As Double
Private m_dblEpsilon As Double
Private m_dblWeightDecay As Double
Private m_oStates As Object

Public Property Get LearningRate() As Double
    LearningRate = m_dblLearningRate
End Property

Public Property Let LearningRate(ByVal dblLearningRate As Double)
    Const PROCEDURE_NAME As String = "Adam.LearningRate"
    
    If dblLearningRate <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Learning rate must be greater than 0."
    End If
    m_dblLearningRate = dblLearningRate
End Property

Public Sub Init(Optional ByVal dblLearningRate As Double = 0.001, _
                Optional ByVal dblBeta1 As Double = 0.9, _
                Optional ByVal dblBeta2 As Double = 0.999, _
                Optional ByVal dblEpsilon As Double = 0.00000001, _
                Optional ByVal dblWeightDecay As Double = 0.01)
    Const PROCEDURE_NAME As String = "Adam.Init"
    
    If dblLearningRate <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Learning rate must be greater than 0."
    End If
    If dblBeta1 <= 0 Or dblBeta1 >= 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Beta1 must be greater than 0 and less than 1."
    End If
    If dblBeta2 <= 0 Or dblBeta2 >= 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Beta2 must be greater than 0 and less than 1."
    End If
    If dblEpsilon <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Epsilon must be greater than 0."
    End If
    If dblWeightDecay < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Weight decay must be greater than or equal to 0."
    End If
    m_dblLearningRate = dblLearningRate
    m_dblBeta1 = dblBeta1
    m_dblBeta2 = dblBeta2
    m_dblEpsilon = dblEpsilon
    m_dblWeightDecay = dblWeightDecay
    Set m_oStates = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Update(ByVal W As Parameter)
    Const PROCEDURE_NAME As String = "Adam.Update"
    
    If W Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
    End If
    ApplyUpdates W
End Sub

Private Sub ApplyUpdates(ByVal W As Parameter)
    Dim lStep As Long
    Dim dblAlpha As Double
    Dim dblWeightDecay As Double
    Dim m As Tensor
    Dim v As Tensor
    
    If W.NumUpdates < 1 Then
        Exit Sub
    End If
    GetStates W, lStep, m, v
    lStep = lStep + 1
    dblAlpha = W.LearningRateScale * m_dblLearningRate * Sqr(1 - m_dblBeta2 ^ lStep) / (1 - m_dblBeta1 ^ lStep)
    dblWeightDecay = W.WeightDecayScale * m_dblWeightDecay
    'm = beta1 * m + (1 - beta1) * (dW / NumUpdates)
    VecLinComb_I m_dblBeta1, m, (1 - m_dblBeta1) / (W.NumUpdates), W.Gradient
    'v = beta2 * v + (1 - beta2) * (dW / NumUpdates)^2
    VecLinComb_I m_dblBeta2, v, (1 - m_dblBeta2) / (W.NumUpdates * W.NumUpdates), VecPow2(W.Gradient)
    'W = W - alpha * (m / (Sqr(v) + epsilon) + wd * W)
    VecLinComb_I 1 - dblAlpha * dblWeightDecay, W.Variable, -dblAlpha, VecDivSqrtAddC(m, v, m_dblEpsilon)
    W.ResetGradient
    PutStates W, lStep, m, v
End Sub

Private Sub GetStates(ByVal W As Parameter, _
                      ByRef lStep As Long, _
                      ByRef m As Tensor, _
                      ByRef v As Tensor)
    With m_oStates
        If Not .Exists(W) Then
            PutStates W, 0&, Zeros(W.Variable.Shape), Zeros(W.Variable.Shape)
        End If
        With .Item(W)
            lStep = .Item(KEY_STEP)
            Set m = .Item(KEY_MEAN)
            Set v = .Item(KEY_VARIANCE)
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
            Set .Item(KEY_MEAN) = m
            Set .Item(KEY_VARIANCE) = v
        End With
    End With
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Adam.Save"
    Dim W As Variant
    Dim lStep As Long
    Dim m As Tensor
    Dim v As Tensor
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteDouble m_dblLearningRate
        .WriteDouble m_dblBeta1
        .WriteDouble m_dblBeta2
        .WriteDouble m_dblEpsilon
        .WriteDouble m_dblWeightDecay
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

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Adam.Load"
    Dim lNumStates As Long
    Dim i As Long
    Dim W As Parameter
    Dim lStep As Long
    Dim m As Tensor
    Dim v As Tensor
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Set m_oStates = CreateObject("Scripting.Dictionary")
    With oWorksheet
        m_dblLearningRate = .ReadDouble()
        m_dblBeta1 = .ReadDouble()
        m_dblBeta2 = .ReadDouble()
        m_dblEpsilon = .ReadDouble()
        m_dblWeightDecay = .ReadDouble()
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
        IOptimizer_LearningRate = LearningRate
    End Property
    
    Private Property Let IOptimizer_LearningRate(ByVal dblLearningRate As Double)
        LearningRate = dblLearningRate
    End Property
    
    Private Sub IOptimizer_Update(ByVal W As Object)
        Update W
    End Sub
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Object)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Object)
        Load oWorksheet
    End Sub
#Else
    Private Property Get IOptimizer_LearningRate() As Double
        IOptimizer_LearningRate = LearningRate
    End Property
    
    Private Property Let IOptimizer_LearningRate(ByVal dblLearningRate As Double)
        LearningRate = dblLearningRate
    End Property
    
    Private Sub IOptimizer_Update(ByVal W As Parameter)
        Update W
    End Sub
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If
