VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TensorDataset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDataset

#Const USE_LATE_BINDING_IN_INTERFACES = True

'State
Private m_lNumSamples As Long
Private m_oTensors As Collection

Public Property Get NumSamples() As Long
    NumSamples = m_lNumSamples
End Property

Public Property Get Cache() As TensorDataset
    Set Cache = Me
End Property

Public Property Get NumTensors() As Long
    NumTensors = m_oTensors.Count
End Property

Public Property Get Tensor(ByVal lTensorIndex As Long) As Tensor
    Const PROCEDURE_NAME As String = "TensorDataset.Item"

    If lTensorIndex < 1 Or lTensorIndex > m_oTensors.Count Then
        Err.Raise 9, PROCEDURE_NAME, "Tensor index is out of range."
    End If
    Set Tensor = m_oTensors(lTensorIndex)
End Property

Public Sub Init(ByVal vTensors As Variant)
    Const PROCEDURE_NAME As String = "TensorDataset.Init"
    Dim i As Long
    Dim lNumTensors As Long
    Dim lNumSamples As Long
    Dim X As Variant
    Dim oTensors As Collection
    
    Set oTensors = New Collection
    For Each X In EnsureIterable(vTensors)
        If Not IsObject(X) Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
        End If
        If X Is Nothing Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
        End If
        If Not TypeOf X Is Tensor Then
            Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
        End If
        If X.NumDimensions < 1 Then
            Err.Raise 5, PROCEDURE_NAME, "Tensor must have at least one dimension."
        End If
        If oTensors.Count = 0 Then
            lNumSamples = X.Size(X.NumDimensions)
        End If
        If X.Size(X.NumDimensions) <> lNumSamples Then
            Err.Raise 5, PROCEDURE_NAME, "All tensors in the dataset must have the same number of samples."
        End If
        oTensors.Add X
    Next X
    Clear
    m_lNumSamples = lNumSamples
    Set m_oTensors = oTensors
End Sub

Public Function Gather(ByVal vSampleIndices As Variant) As TensorDataset
    Dim i As Long
    Dim aoGatheredTensors() As Tensor
    Dim oResult As TensorDataset
    
    If m_oTensors.Count > 0 Then
        ReDim aoGatheredTensors(1 To m_oTensors.Count)
        For i = 1 To m_oTensors.Count
            Set aoGatheredTensors(i) = m_oTensors(i).Gather(m_oTensors(i).NumDimensions, vSampleIndices)
        Next i
    End If
    Set oResult = New TensorDataset
    oResult.Init aoGatheredTensors
    Set Gather = oResult
End Function

Public Function Subset(ByVal vSampleIndices As Variant) As SubsetDataset
    Dim oResult As SubsetDataset

    Set oResult = New SubsetDataset
    oResult.Init Me, vSampleIndices
    Set Subset = oResult
End Function

Private Sub Clear()
    m_lNumSamples = 0
    Set m_oTensors = New Collection
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get IDataset_NumSamples() As Long
        IDataset_NumSamples = Me.NumSamples
    End Property
    
    Private Property Get IDataset_Cache() As Object
        Set IDataset_Cache = Me.Cache()
    End Property
    
    Private Function IDataset_Gather(ByVal vSampleIndices As Variant) As Object
        Set IDataset_Gather = Me.Gather(vSampleIndices)
    End Function
    
    Private Function IDataset_Subset(ByVal vSampleIndices As Variant) As Object
        Set IDataset_Subset = Me.Subset(vSampleIndices)
    End Function
#Else
    Private Property Get IDataset_NumSamples() As Long
        IDataset_NumSamples = Me.NumSamples
    End Property
    
    Private Property Get IDataset_Cache() As TensorDataset
        Set IDataset_Cache = Me.Cache()
    End Property
    
    Private Function IDataset_Gather(ByVal vSampleIndices As Variant) As TensorDataset
        Set IDataset_Gather = Me.Gather(vSampleIndices)
    End Function
    
    Private Function IDataset_Subset(ByVal vSampleIndices As Variant) As SubsetDataset
        Set IDataset_Subset = Me.Subset(vSampleIndices)
    End Function
#End If
