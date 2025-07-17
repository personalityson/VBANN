VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Hyperparameters
Private m_dblLearningRateScale As Double
Private m_dblWeightDecayScale As Double

'State
Private m_oVariable As Tensor
Private m_oGradient As Tensor
Private m_lNumUpdates As Long

Public Property Get Variable() As Tensor
    Set Variable = m_oVariable
End Property

Public Property Get Gradient() As Tensor
    Set Gradient = m_oGradient
End Property

Public Property Get NumUpdates() As Long
    NumUpdates = m_lNumUpdates
End Property

Public Property Let NumUpdates(ByVal lNumUpdates As Long)
    m_lNumUpdates = lNumUpdates
End Property

Public Property Get LearningRateScale() As Double
    LearningRateScale = m_dblLearningRateScale
End Property

Public Property Get WeightDecayScale() As Double
    WeightDecayScale = m_dblWeightDecayScale
End Property

Public Sub Init(ByVal oVariable As Tensor, _
                Optional ByVal dblLearningRateScale As Double = 1, _
                Optional ByVal dblWeightDecayScale As Double = 1)
    Const PROCEDURE_NAME As String = "Parameter.Init"
    
    If oVariable Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Clear
    Set m_oVariable = oVariable
    Set m_oGradient = Zeros(m_oVariable.Shape)
    m_lNumUpdates = 0
    m_dblLearningRateScale = dblLearningRateScale
    m_dblWeightDecayScale = dblWeightDecayScale
End Sub

Public Sub ResetGradient()
    m_oGradient.Fill 0
    m_lNumUpdates = 0
End Sub

Private Sub Clear()
    Set m_oVariable = Nothing
    Set m_oGradient = Nothing
    m_lNumUpdates = 0
    m_dblLearningRateScale = 0
    m_dblWeightDecayScale = 0
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Parameter.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteObject m_oVariable
        .WriteObject m_oGradient
        .WriteLong m_lNumUpdates
        .WriteDouble m_dblLearningRateScale
        .WriteDouble m_dblWeightDecayScale
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Parameter.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oWorksheet
        Set m_oVariable = .ReadObject()
        Set m_oGradient = .ReadObject()
        m_lNumUpdates = .ReadLong()
        m_dblLearningRateScale = .ReadDouble()
        m_dblWeightDecayScale = .ReadDouble()
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
