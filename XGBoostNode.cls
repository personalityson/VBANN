VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XGBoostNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Node type
Private m_bIsLeaf As Boolean

'Leaf parameters
Private m_dblWeight As Double

'Split parameters
Private m_lFeatureIndex As Long
Private m_dblThreshold As Double
Private m_oChildL As XGBoostNode
Private m_oChildR As XGBoostNode

Public Sub MakeLeaf(ByVal dblWeight As Double)
    Clear
    m_bIsLeaf = True
    m_dblWeight = dblWeight
End Sub

Public Sub Split(ByVal lFeatureIndex As Long, _
                 ByVal dblThreshold As Double, _
                 ByVal oChildL As XGBoostNode, _
                 ByVal oChildR As XGBoostNode)
    Clear
    m_bIsLeaf = False
    m_lFeatureIndex = lFeatureIndex
    m_dblThreshold = dblThreshold
    Set m_oChildL = oChildL
    Set m_oChildR = oChildR
End Sub

Public Function PredictRow(ByRef X_() As Double, _
                           ByVal lSampleIndex As Long) As Double
    If m_bIsLeaf Then
        PredictRow = m_dblWeight
    Else
        If X_(m_lFeatureIndex, lSampleIndex) < m_dblThreshold Then
            PredictRow = m_oChildL.PredictRow(X_, lSampleIndex)
        Else
            PredictRow = m_oChildR.PredictRow(X_, lSampleIndex)
        End If
    End If
End Function

Private Sub Clear()
    m_bIsLeaf = False
    m_dblWeight = 0
    m_lFeatureIndex = 0
    m_dblThreshold = 0
    Set m_oChildL = Nothing
    Set m_oChildR = Nothing
End Sub

Private Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "XGBoost.Save"

    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteBoolean m_bIsLeaf
        If m_bIsLeaf Then
            .WriteDouble m_dblWeight
        Else
            .WriteLong m_lFeatureIndex
            .WriteDouble m_dblThreshold
            .WriteObject m_oChildL
            .WriteObject m_oChildR
        End If
    End With
End Sub

Private Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "XGBoost.Load"

    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oWorksheet
        m_bIsLeaf = .ReadBoolean
        If m_bIsLeaf Then
            m_dblWeight = .ReadDouble()
        Else
            m_lFeatureIndex = .ReadLong()
            m_dblThreshold = .ReadDouble()
            Set m_oChildL = .ReadObject()
            Set m_oChildR = .ReadObject()
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
