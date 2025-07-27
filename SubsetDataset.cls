VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubsetDataset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDataset

#Const USE_LATE_BINDING_IN_INTERFACES = True

'Dependencies
Private m_oDataset As IDataset

'State
Private m_lNumSamples As Long
Private m_alIndices() As Long

'Gather cache
Private m_oGatherCache As TensorDataset

Public Property Get NumSamples() As Long
    NumSamples = m_lNumSamples
End Property

Public Sub Init(ByVal oDataset As IDataset, _
                ByVal vIndices As Variant)
    Const PROCEDURE_NAME As String = "SubsetDataset.Init"
    Dim i As Long
    Dim lNumIndices As Long
    Dim alIndices() As Long
    
    If oDataset Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IDataset object is required."
    End If
    ParseVariantToLongArray vIndices, lNumIndices, alIndices
    For i = 1 To lNumIndices
        If alIndices(i) < 1 Or alIndices(i) > oDataset.NumSamples Then
            Err.Raise 9, PROCEDURE_NAME, "Sample index is out of range."
        End If
    Next i
    Clear
    Set m_oDataset = oDataset
    m_lNumSamples = lNumIndices
    m_alIndices = alIndices
End Sub

Public Function Cache() As TensorDataset
    If m_oGatherCache Is Nothing Then
        Set m_oGatherCache = m_oDataset.Gather(m_alIndices)
    End If
    Set Cache = m_oGatherCache
End Function

Public Function Gather(ByVal vIndices As Variant) As TensorDataset
    Dim alIndices() As Long
    
    alIndices = MapLogicalToPhysical(vIndices)
    Set Gather = m_oDataset.Gather(alIndices)
End Function

Public Function Subset(ByVal vIndices As Variant) As SubsetDataset
    Dim alIndices() As Long
    
    alIndices = MapLogicalToPhysical(vIndices)
    Set Subset = m_oDataset.Subset(alIndices)
End Function

Private Function MapLogicalToPhysical(ByVal vIndices As Variant) As Long()
    Const PROCEDURE_NAME As String = "SubsetDataset.MapLogicalToPhysical"
    Dim i As Long
    Dim lNumIndices As Long
    Dim alIndices() As Long
    
    ParseVariantToLongArray vIndices, lNumIndices, alIndices
    For i = 1 To lNumIndices
        If alIndices(i) < 1 Or alIndices(i) > m_lNumSamples Then
            Err.Raise 9, PROCEDURE_NAME, "Sample index is out of range."
        End If
        alIndices(i) = m_alIndices(alIndices(i))
    Next i
    MapLogicalToPhysical = alIndices
End Function

Private Sub Clear()
    Set m_oDataset = Nothing
    m_lNumSamples = 0
    Erase m_alIndices
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get IDataset_NumSamples() As Long
        IDataset_NumSamples = Me.NumSamples
    End Property
    
    Private Function IDataset_Cache() As Object
        Set IDataset_Cache = Me.Cache()
    End Function
    
    Private Function IDataset_Gather(ByVal vIndices As Variant) As Object
        Set IDataset_Gather = Me.Gather(vIndices)
    End Function
    
    Private Function IDataset_Subset(ByVal vIndices As Variant) As Object
        Set IDataset_Subset = Me.Subset(vIndices)
    End Function
#Else
    Private Property Get IDataset_NumSamples() As Long
        IDataset_NumSamples = Me.NumSamples
    End Property
    
    Private Function IDataset_Cache() As TensorDataset
        Set IDataset_Cache = Me.Cache()
    End Function
    
    Private Function IDataset_Gather(ByVal vIndices As Variant) As TensorDataset
        Set IDataset_Gather = Me.Gather(vIndices)
    End Function
    
    Private Function IDataset_Subset(ByVal vIndices As Variant) As SubsetDataset
        Set IDataset_Subset = Me.Subset(vIndices)
    End Function
#End If
