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
Private m_dblDecayScale As Double
Private m_bUseGradientClipping As Boolean

'State
Private m_oTensor As Tensor
Private m_oGradient As Tensor
Private m_lNumAccumulatedSamples As Long

Public Property Get Tensor() As Tensor
    Set Tensor = m_oTensor
End Property

Public Property Get Gradient() As Tensor
    Set Gradient = m_oGradient
End Property

Public Property Get NumAccumulatedSamples() As Long
    NumAccumulatedSamples = m_lNumAccumulatedSamples
End Property

Public Property Let NumAccumulatedSamples(ByVal lNumAccumulatedSamples As Long)
    m_lNumAccumulatedSamples = lNumAccumulatedSamples
End Property

Public Property Get LearningRateScale() As Double
    LearningRateScale = m_dblLearningRateScale
End Property

Public Property Get DecayScale() As Double
    DecayScale = m_dblDecayScale
End Property

Public Property Get UseGradientClipping() As Boolean
    UseGradientClipping = m_bUseGradientClipping
End Property

Public Sub Init(ByVal oTensor As Tensor, _
                Optional ByVal dblLearningRateScale As Double = 1, _
                Optional ByVal dblDecayScale As Double = 1, _
                Optional ByVal bUseGradientClipping As Boolean = True)
    Const PROCEDURE_NAME As String = "Parameter.Init"
    
    If oTensor Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    Clear
    Set m_oTensor = oTensor
    Set m_oGradient = Zeros(m_oTensor.Shape)
    m_lNumAccumulatedSamples = 0
    m_dblLearningRateScale = dblLearningRateScale
    m_dblDecayScale = dblDecayScale
    m_bUseGradientClipping = bUseGradientClipping
End Sub

Public Sub ResetGradient()
    m_oGradient.Fill 0
    m_lNumAccumulatedSamples = 0
End Sub

Private Sub Clear()
    Set m_oTensor = Nothing
    Set m_oGradient = Nothing
    m_lNumAccumulatedSamples = 0
    m_dblLearningRateScale = 0
    m_dblDecayScale = 0
    m_bUseGradientClipping = False
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Parameter.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteObject m_oTensor
        .WriteObject m_oGradient
        .WriteLong m_lNumAccumulatedSamples
        .WriteDouble m_dblLearningRateScale
        .WriteDouble m_dblDecayScale
        .WriteBoolean m_bUseGradientClipping
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "Parameter.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oWorksheet
        Set m_oTensor = .ReadObject()
        Set m_oGradient = .ReadObject()
        m_lNumAccumulatedSamples = .ReadLong()
        m_dblLearningRateScale = .ReadDouble()
        m_dblDecayScale = .ReadDouble()
        m_bUseGradientClipping = .ReadBoolean()
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
