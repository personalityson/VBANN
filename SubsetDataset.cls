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
Private m_oSource As IDataset

'State
Private m_lNumSamples As Long
Private m_alIndices() As Long

'Cache
Private m_oCache As TensorDataset

Public Property Get NumSamples() As Long
    NumSamples = m_lNumSamples
End Property

Public Property Get NumTensors() As Long
    NumTensors = m_oSource.NumTensors
End Property

Public Property Get Tensors(ByVal lIndex As Long) As Tensor
    If m_oCache Is Nothing Then
        Set m_oCache = m_oSource.Gather(m_alIndices)
    End If
    Set Tensors = m_oCache.Tensors(lIndex)
End Property

Public Sub Init(ByVal oSource As IDataset, _
                ByVal vIndices As Variant)
    Const PROCEDURE_NAME As String = "SubsetDataset.Init"
    Dim i As Long
    Dim lNumIndices As Long
    Dim alIndices() As Long
    
    If oSource Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IDataset object is required."
    End If
    ParseVariantToLongArray vIndices, lNumIndices, alIndices
    For i = 1 To lNumIndices
        If alIndices(i) < 1 Or alIndices(i) > oSource.NumSamples Then
            Err.Raise 9, PROCEDURE_NAME, "Sample index is out of range."
        End If
    Next i
    Clear
    Set m_oSource = oSource
    m_lNumSamples = lNumIndices
    m_alIndices = alIndices
End Sub

Public Function Gather(ByVal vIndices As Variant) As TensorDataset
    Dim alIndices() As Long
    
    alIndices = MapLogicalToPhysical(vIndices)
    Set Gather = m_oSource.Gather(alIndices)
End Function

Public Function Subset(ByVal vIndices As Variant) As SubsetDataset
    Dim alIndices() As Long
    
    alIndices = MapLogicalToPhysical(vIndices)
    Set Subset = m_oSource.Subset(alIndices)
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
    Set m_oSource = Nothing
    m_lNumSamples = 0
    Erase m_alIndices
    Set m_oCache = Nothing
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get IDataset_NumSamples() As Long
        IDataset_NumSamples = Me.NumSamples
    End Property
    
    Private Property Get IDataset_NumTensors() As Long
        IDataset_NumTensors = Me.NumTensors
    End Property
    
    Private Property Get IDataset_Tensors(ByVal lIndex As Long) As Object
        Set IDataset_Tensors = Me.Tensors(lIndex)
    End Property
    
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
    
    Private Property Get IDataset_NumTensors() As Long
        IDataset_NumTensors = Me.NumTensors
    End Property
    
    Private Property Get IDataset_Tensors(ByVal lIndex As Long) As Tensor
        Set IDataset_Tensors = Me.Tensors(lIndex)
    End Property
    
    Private Function IDataset_Gather(ByVal vIndices As Variant) As TensorDataset
        Set IDataset_Gather = Me.Gather(vIndices)
    End Function
    
    Private Function IDataset_Subset(ByVal vIndices As Variant) As SubsetDataset
        Set IDataset_Subset = Me.Subset(vIndices)
    End Function
#End If
