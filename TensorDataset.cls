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
Private m_lNumTensors As Long
Private m_aoTensors() As Tensor
Private m_lNumSamples As Long

Public Property Get NumTensors() As Long
    NumTensors = m_lNumTensors
End Property

Public Property Get Tensor(ByVal lTensorIndex As Long) As Tensor
    Const PROCEDURE_NAME As String = "TensorDataset.Item"

    If lTensorIndex < 1 Or lTensorIndex > m_lNumTensors Then
        Err.Raise 9, PROCEDURE_NAME, "Tensor index is out of range."
    End If
    Set Tensor = m_aoTensors(lTensorIndex)
End Property

Public Property Get NumSamples() As Long
    NumSamples = m_lNumSamples
End Property

Public Sub Init(ByVal vTensors As Variant)
    Const PROCEDURE_NAME As String = "TensorDataset.Init"
    Dim i As Long
    Dim lNumTensors As Long
    Dim lNumSamples As Long
    Dim X As Variant
    Dim aoTensors() As Tensor
    
    Select Case GetRank(vTensors)
        Case -1
            vTensors = Array(vTensors)
        Case 0
            Clear
            Exit Sub
        Case 1
            'Continue
        Case Else
            Err.Raise 5, PROCEDURE_NAME, "Expected a single tensor, an uninitialized array, or a one-dimensional array of tensors."
    End Select
    lNumTensors = UBound(vTensors) - LBound(vTensors) + 1
    If lNumTensors > 0 Then
        ReDim aoTensors(1 To lNumTensors)
    End If
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
        If i = 0 Then
            lNumSamples = X.Size(X.NumDimensions)
        End If
        If X.Size(X.NumDimensions) <> lNumSamples Then
            Err.Raise 5, PROCEDURE_NAME, "All tensors in the dataset must have the same number of samples."
        End If
        i = i + 1
        Set aoTensors(i) = X
    Next X
    Clear
    m_lNumTensors = lNumTensors
    m_aoTensors = aoTensors
    m_lNumSamples = lNumSamples
End Sub

Public Function Cache() As TensorDataset
    Set Cache = Me
End Function

Public Function Gather(ByVal vIndices As Variant) As TensorDataset
    Dim i As Long
    Dim aoGatheredTensors() As Tensor
    Dim oResult As TensorDataset
    
    If m_lNumTensors > 0 Then
        ReDim aoGatheredTensors(1 To m_lNumTensors)
        For i = 1 To m_lNumTensors
            Set aoGatheredTensors(i) = m_aoTensors(i).Gather(m_aoTensors(i).NumDimensions, vIndices)
        Next i
    End If
    Set oResult = New TensorDataset
    oResult.Init aoGatheredTensors
    Set Gather = oResult
End Function

Public Function Subset(ByVal vIndices As Variant) As SubsetDataset
    Dim oResult As SubsetDataset

    Set oResult = New SubsetDataset
    oResult.Init Me, vIndices
    Set Subset = oResult
End Function

Private Sub Clear()
    m_lNumTensors = 0
    Erase m_aoTensors
    m_lNumSamples = 0
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
    
    Private Function IDataset_Cache() As Object
        Set IDataset_Cache = Me.Cache()
    End Function
    
    Private Function IDataset_Gather(ByVal vIndices As Variant) As TensorDataset
        Set IDataset_Gather = Me.Gather(vIndices)
    End Function
    
    Private Function IDataset_Subset(ByVal vIndices As Variant) As SubsetDataset
        Set IDataset_Subset = Me.Subset(vIndices)
    End Function
#End If
