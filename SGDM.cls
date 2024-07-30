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

#Const USE_LATE_BINDING_IN_INTERFACES = True

Implements IOptimizer
Implements ISerializable

Private m_dblLearningRate As Double
Private m_dblMomentum As Double
Private m_dblWeightDecay As Double
Private m_oStates As Object

Public Property Get LearningRate() As Double
    LearningRate = m_dblLearningRate
End Property

Public Property Let LearningRate(ByVal dblLearningRate As Double)
    Const PROCEDURE_NAME As String = "SGDM.LearningRate"
    
    If dblLearningRate <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Learning rate must be greater than 0."
    End If
    m_dblLearningRate = dblLearningRate
End Property

Public Property Get Momentum() As Double
    Momentum = m_dblMomentum
End Property

Public Property Get WeightDecay() As Double
    WeightDecay = m_dblWeightDecay
End Property

Private Sub Class_Initialize()
    Set m_oStates = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Init(Optional ByVal dblLearningRate As Double = 0.001, _
                Optional ByVal dblMomentum As Double = 0.9, _
                Optional ByVal dblWeightDecay As Double = 0.01)
    Const PROCEDURE_NAME As String = "SGDM.Init"
    
    If dblLearningRate <= 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Learning rate must be greater than 0."
    End If
    If dblMomentum < 0 Or dblMomentum >= 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Momentum must be greater than or equal to 0 and less than 1."
    End If
    If dblWeightDecay < 0 Then
        Err.Raise 5, PROCEDURE_NAME, "Weight decay must be greater than or equal to 0."
    End If
    m_dblLearningRate = dblLearningRate
    m_dblMomentum = dblMomentum
    m_dblWeightDecay = dblWeightDecay
    m_oStates.RemoveAll
End Sub

Public Sub Update(ByVal W As Parameter)
    Const PROCEDURE_NAME As String = "SGDM.Update"
    
    If W Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Parameter object is required."
    End If
    If IsBlasAvailable() Then
        ApplyUpdatesWithBlas W
    Else
        ApplyUpdatesNaively W
    End If
End Sub

Public Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "SGDM.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteDouble m_dblLearningRate
        .WriteDouble m_dblMomentum
        .WriteDouble m_dblWeightDecay
    End With
End Sub

Public Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "SGDM.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        m_dblLearningRate = .ReadDouble()
        m_dblMomentum = .ReadDouble()
        m_dblWeightDecay = .ReadDouble()
    End With
End Sub

Private Sub GetStates(ByVal W As Parameter, _
                      ByRef v As Tensor)
    With m_oStates
        If Not .Exists(W) Then
            PutStates W, Zeros(W.Learnable.Shape)
        End If
        With .Item(W)
            Set v = .Item("Velocity")
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
            Set .Item("Velocity") = v
        End With
    End With
End Sub

Private Sub ApplyUpdatesNaively(ByVal W As Parameter)
    Dim lNumUpdates As Long
    Dim i As Long
    Dim dblLearningRate As Double
    Dim dblWeightDecay As Double
    Dim W_() As Double
    Dim dLdW_() As Double
    Dim v_() As Double
    Dim v As Tensor
    
    If W.NumUpdates < 1 Then
        Exit Sub
    End If
    GetStates W, v
    lNumUpdates = W.NumUpdates
    dblLearningRate = W.LearningRateFactor * m_dblLearningRate
    dblWeightDecay = W.WeightDecayFactor * m_dblWeightDecay
    W.Learnable.Ravel.CreateAlias W_
    W.Gradient.Ravel.CreateAlias dLdW_
    v.Ravel.CreateAlias v_
    For i = 1 To W.Learnable.NumElements
        v_(i) = m_dblMomentum * v_(i) + (1 - m_dblMomentum) * dLdW_(i) / lNumUpdates
        W_(i) = W_(i) - dblLearningRate * (v_(i) + dblWeightDecay * W_(i))
    Next i
    W.Learnable.Ravel.RemoveAlias W_
    W.Gradient.Ravel.RemoveAlias dLdW_
    v.Ravel.RemoveAlias v_
    PutStates W, v
    W.ResetGradient
End Sub

Private Sub ApplyUpdatesWithBlas(ByVal W As Parameter)
    Dim lNumUpdates As Long
    Dim dblLearningRate As Double
    Dim dblWeightDecay As Double
    Dim v As Tensor
    
    If W.NumUpdates < 1 Then
        Exit Sub
    End If
    GetStates W, v
    lNumUpdates = W.NumUpdates
    dblLearningRate = W.LearningRateFactor * m_dblLearningRate
    dblWeightDecay = W.WeightDecayFactor * m_dblWeightDecay
    daxpby W.Learnable.NumElements, (1 - m_dblMomentum) / lNumUpdates, W.Gradient.Address, 1&, m_dblMomentum, v.Address, 1&
    daxpby W.Learnable.NumElements, -dblLearningRate, v.Address, 1&, 1 - dblLearningRate * dblWeightDecay, W.Learnable.Address, 1&
    PutStates W, v
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

