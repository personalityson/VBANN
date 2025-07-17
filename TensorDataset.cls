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

#Const USE_LATE_BINDING_IN_INTERFACES = True

Implements IDataset

'Dependencies
Private m_lNumTensors As Long
Private m_aoTensors() As Tensor

'State
Private m_lNumSamples As Long
Private m_alIndices() As Long

Public Property Get NumSamples() As Long
    NumSamples = m_lNumSamples
End Property

Public Sub Init(ByVal vTensors As Variant, _
                Optional ByVal vSampleIndices As Variant)
    Const PROCEDURE_NAME As String = "TensorDataset.Init"
    
    Select Case GetRank(vTensors)
        Case -1
            Me.Init Array(vTensors), vSampleIndices
        Case 0
            Clear
        Case 1
            Clear
            SetTensors vTensors
            If IsMissing(vSampleIndices) Then
                m_alIndices = GetIdentityPermutationArray(m_lNumSamples)
            Else
                SetIndices vSampleIndices
            End If
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Expected a single tensor, an uninitialized array, or a one-dimensional array of tensors."
    End Select
End Sub

Public Function Item(ByVal lTensorIndex As Long) As Tensor
    Const PROCEDURE_NAME As String = "TensorDataset.Item"

    If lTensorIndex < 1 Or lTensorIndex > m_lNumTensors Then
        Err.Raise 9, PROCEDURE_NAME, "Tensor index is out of range."
    End If
    With m_aoTensors(lTensorIndex)
        Set Item = .Gather(.NumDimensions, m_alIndices)
    End With
End Function

Public Function Subset(ByVal vSampleIndices As Variant) As TensorDataset
    Const PROCEDURE_NAME As String = "TensorDataset.Subset"
    Dim i As Long
    Dim lNumSampleIndices As Long
    Dim alSampleIndices() As Long
    Dim oResult As TensorDataset
    
    ParseVariantToLongArray vSampleIndices, lNumSampleIndices, alSampleIndices
    For i = 1 To lNumSampleIndices
        If alSampleIndices(i) < 1 Or alSampleIndices(i) > m_lNumSamples Then
            Err.Raise 9, PROCEDURE_NAME, "Sample index is out of range."
        End If
        alSampleIndices(i) = m_alIndices(alSampleIndices(i))
    Next i
    Set oResult = New TensorDataset
    oResult.Init m_aoTensors, alSampleIndices
    Set Subset = oResult
End Function

Private Sub Clear()
    m_lNumTensors = 0
    Erase m_aoTensors
    m_lNumSamples = 0
    Erase m_alIndices
End Sub

Private Sub SetTensors(ByVal vTensors As Variant)
    Const PROCEDURE_NAME As String = "TensorDataset.SetTensors"
    Dim i As Long
    Dim lThisNumSamples As Long
    Dim X As Variant
    
    m_lNumTensors = UBound(vTensors) - LBound(vTensors) + 1
    ReDim m_aoTensors(1 To m_lNumTensors)
    For Each X In vTensors
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
        lThisNumSamples = X.Size(X.NumDimensions)
        If i = 0 Then
            m_lNumSamples = lThisNumSamples
        ElseIf lThisNumSamples <> m_lNumSamples Then
            Err.Raise 5, PROCEDURE_NAME, "All tensors in the dataset must have the same number of samples."
        End If
        i = i + 1
        Set m_aoTensors(i) = X
    Next X
End Sub

Private Sub SetIndices(ByVal vSampleIndices As Variant)
    Const PROCEDURE_NAME As String = "TensorDataset.SetIndices"
    Dim i As Long
    Dim lNumSampleIndices As Long
    Dim alSampleIndices() As Long

    ParseVariantToLongArray vSampleIndices, lNumSampleIndices, alSampleIndices
    For i = 1 To lNumSampleIndices
        If alSampleIndices(i) < 1 Or alSampleIndices(i) > m_lNumSamples Then
            Err.Raise 9, PROCEDURE_NAME, "Sample index is out of range."
        End If
    Next i
    m_lNumSamples = lNumSampleIndices
    m_alIndices = alSampleIndices
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get IDataset_NumSamples() As Long
        IDataset_NumSamples = Me.NumSamples
    End Property
    
    Private Function IDataset_Item(ByVal lTensorIndex As Long) As Object
        Set IDataset_Item = Me.Item(lTensorIndex)
    End Function
    
    Private Function IDataset_Subset(ByVal vSampleIndices As Variant) As Object
        Set IDataset_Subset = Me.Subset(vSampleIndices)
    End Function
#Else
    Private Property Get IDataset_NumSamples() As Long
        IDataset_NumSamples = Me.NumSamples
    End Property
    
    Private Function IDataset_Item(ByVal lTensorIndex As Long) As Tensor
        IDataset_Item = Me.Item(lTensorIndex)
    End Function
    
    Private Function IDataset_Subset(ByVal vSampleIndices As Variant) As IDataset
        Set IDataset_Subset = Me.Subset(vSampleIndices)
    End Function
#End If
