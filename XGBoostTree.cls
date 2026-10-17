VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XGBoostTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Structure
Private m_lNumNodes As Long
Private m_alFeature() As Long
Private m_adblThreshold() As Double
Private m_alLeft() As Long
Private m_alRight() As Long
Private m_adblWeight() As Double

Public Property Get NumNodes() As Long
    NumNodes = m_lNumNodes
End Property

Public Sub Init(ByVal lCapacity As Long)
    Const PROCEDURE_NAME As String = "XGBoostTree.Init"

    If lCapacity < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Capacity must be >= 1."
    End If
    Clear
    ReDim m_alFeature(1 To lCapacity)
    ReDim m_adblThreshold(1 To lCapacity)
    ReDim m_alLeft(1 To lCapacity)
    ReDim m_alRight(1 To lCapacity)
    ReDim m_adblWeight(1 To lCapacity)
End Sub

Public Function AddNode() As Long
    Const PROCEDURE_NAME As String = "XGBoostTree.AddNode"

    If m_lNumNodes >= UBound(m_alFeature) Then
        Err.Raise 5, PROCEDURE_NAME, "Tree has outgrown its reserved capacity."
    End If
    m_lNumNodes = m_lNumNodes + 1
    AddNode = m_lNumNodes
End Function

Public Sub SetLeaf(ByVal lNode As Long, _
                   ByVal dblWeight As Double)
    Const PROCEDURE_NAME As String = "XGBoostTree.SetLeaf"

    If lNode < 1 Or lNode > m_lNumNodes Then
        Err.Raise 9, PROCEDURE_NAME, "Node index is out of range."
    End If
    m_alFeature(lNode) = 0
    m_adblWeight(lNode) = dblWeight
End Sub

Public Sub SetSplit(ByVal lNode As Long, _
                    ByVal lFeature As Long, _
                    ByVal dblThreshold As Double, _
                    ByVal lLeft As Long, _
                    ByVal lRight As Long)
    Const PROCEDURE_NAME As String = "XGBoostTree.SetSplit"

    If lNode < 1 Or lNode > m_lNumNodes Then
        Err.Raise 9, PROCEDURE_NAME, "Node index is out of range."
    End If
    If lFeature < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Feature index must be >= 1."
    End If
    If lLeft <= lNode Or lLeft > m_lNumNodes Then
        Err.Raise 9, PROCEDURE_NAME, "Left child must be a node added after its parent."
    End If
    If lRight <= lNode Or lRight > m_lNumNodes Then
        Err.Raise 9, PROCEDURE_NAME, "Right child must be a node added after its parent."
    End If
    m_alFeature(lNode) = lFeature
    m_adblThreshold(lNode) = dblThreshold
    m_alLeft(lNode) = lLeft
    m_alRight(lNode) = lRight
End Sub

Public Sub Complete()
    Const PROCEDURE_NAME As String = "XGBoostTree.Complete"

    If m_lNumNodes < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Tree has no nodes."
    End If
    ReDim Preserve m_alFeature(1 To m_lNumNodes)
    ReDim Preserve m_adblThreshold(1 To m_lNumNodes)
    ReDim Preserve m_alLeft(1 To m_lNumNodes)
    ReDim Preserve m_alRight(1 To m_lNumNodes)
    ReDim Preserve m_adblWeight(1 To m_lNumNodes)
End Sub

Public Function PredictRow(ByRef X_() As Double, _
                           ByVal j As Long) As Double
    Dim i As Long

    i = 1
    Do While m_alFeature(i) > 0
        If X_(m_alFeature(i), j) < m_adblThreshold(i) Then
            i = m_alLeft(i)
        Else
            i = m_alRight(i)
        End If
    Loop
    PredictRow = m_adblWeight(i)
End Function

Private Sub Clear()
    m_lNumNodes = 0
    Erase m_alFeature
    Erase m_adblThreshold
    Erase m_alLeft
    Erase m_alRight
    Erase m_adblWeight
End Sub

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "XGBoostTree.Save"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteLong m_lNumNodes
        .WriteLongArray m_alFeature
        .WriteDoubleArray m_adblThreshold
        .WriteLongArray m_alLeft
        .WriteLongArray m_alRight
        .WriteDoubleArray m_adblWeight
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "XGBoostTree.Load"

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oSerializer
        m_lNumNodes = .ReadLong()
        m_alFeature = .ReadLongArray()
        m_adblThreshold = .ReadDoubleArray()
        m_alLeft = .ReadLongArray()
        m_alRight = .ReadLongArray()
        m_adblWeight = .ReadDoubleArray()
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Sub ISerializable_Save(ByVal oSerializer As Object)
        Save oSerializer
    End Sub

    Private Sub ISerializable_Load(ByVal oSerializer As Object)
        Load oSerializer
    End Sub
#Else
    Private Sub ISerializable_Save(ByVal oSerializer As Serializer)
        Save oSerializer
    End Sub

    Private Sub ISerializable_Load(ByVal oSerializer As Serializer)
        Load oSerializer
    End Sub
#End If

