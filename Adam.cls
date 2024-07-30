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

#Const USE_LATE_BINDING_IN_INTERFACES = True

Implements IOptimizer
Implements ISerializable

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

Public Property Get Beta1() As Double
    Beta1 = m_dblBeta1
End Property

Public Property Get Beta2() As Double
    Beta2 = m_dblBeta2
End Property

Public Property Get Epsilon() As Double
    Epsilon = m_dblEpsilon
End Property

Public Property Get WeightDecay() As Double
    WeightDecay = m_dblWeightDecay
End Property

Private Sub Class_Initialize()
    Set m_oStates = CreateObject("Scripting.Dictionary")
End Sub

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
    m_oStates.RemoveAll
End Sub

Public Sub Update(ByVal W As Parameter)
    Const PROCEDURE_NAME As String = "Adam.Update"
    
    If W Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
    End If
    ApplyUpdates W
End Sub

Public Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Adam.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteDouble m_dblLearningRate
        .WriteDouble m_dblBeta1
        .WriteDouble m_dblBeta2
        .WriteDouble m_dblEpsilon
        .WriteDouble m_dblWeightDecay
    End With
End Sub

Public Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Adam.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        m_dblLearningRate = .ReadDouble()
        m_dblBeta1 = .ReadDouble()
        m_dblBeta2 = .ReadDouble()
        m_dblEpsilon = .ReadDouble()
        m_dblWeightDecay = .ReadDouble()
    End With
End Sub

Private Sub GetStates(ByVal W As Parameter, _
                      ByRef lStep As Long, _
                      ByRef m As Tensor, _
                      ByRef v As Tensor)
    With m_oStates
        If Not .Exists(W) Then
            PutStates W, 0&, Zeros(W.Learnable.Shape), Zeros(W.Learnable.Shape)
        End If
        With .Item(W)
            lStep = .Item("Step")
            Set m = .Item("Mean")
            Set v = .Item("Variance")
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
            .Item("Step") = lStep
            Set .Item("Mean") = m
            Set .Item("Variance") = v
        End With
    End With
End Sub

Private Sub ApplyUpdates(ByVal W As Parameter)
    Dim lStep As Long
    Dim lNumUpdates As Long
    Dim i As Long
    Dim dblLearningRate As Double
    Dim dblWeightDecay As Double
    Dim dblAlpha As Double
    Dim dblMeanGradient As Double
    Dim W_() As Double
    Dim dLdW_() As Double
    Dim m_() As Double
    Dim v_() As Double
    Dim m As Tensor
    Dim v As Tensor
    
    If W.NumUpdates < 1 Then
        Exit Sub
    End If
    GetStates W, lStep, m, v
    lStep = lStep + 1
    lNumUpdates = W.NumUpdates
    dblLearningRate = W.LearningRateFactor * m_dblLearningRate
    dblWeightDecay = W.WeightDecayFactor * m_dblWeightDecay
    dblAlpha = dblLearningRate * Sqr(1 - m_dblBeta2 ^ lStep) / (1 - m_dblBeta1 ^ lStep)
    W.Learnable.Ravel.CreateAlias W_
    W.Gradient.Ravel.CreateAlias dLdW_
    m.Ravel.CreateAlias m_
    v.Ravel.CreateAlias v_
    For i = 1 To W.Learnable.NumElements
        dblMeanGradient = dLdW_(i) / lNumUpdates
        m_(i) = m_dblBeta1 * m_(i) + (1 - m_dblBeta1) * dblMeanGradient
        v_(i) = m_dblBeta2 * v_(i) + (1 - m_dblBeta2) * dblMeanGradient * dblMeanGradient
        W_(i) = W_(i) - dblAlpha * (m_(i) / (Sqr(v_(i)) + m_dblEpsilon) + dblWeightDecay * W_(i))
    Next i
    W.Learnable.Ravel.RemoveAlias W_
    W.Gradient.Ravel.RemoveAlias dLdW_
    m.Ravel.RemoveAlias m_
    v.Ravel.RemoveAlias v_
    PutStates W, lStep, m, v
    W.ResetGradient
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