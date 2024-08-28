VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleDataset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_LATE_BINDING_IN_INTERFACES = True

Implements IDataset

Private m_lNumSamples As Long
Private m_lNumTensors As Long
Private m_aoTensors() As Tensor

Public Property Get NumSamples() As Long
    NumSamples = m_lNumSamples
End Property

Public Property Get NumTensors() As Long
    NumTensors = m_lNumTensors
End Property

Public Sub Init()
    m_lNumSamples = 0
    m_lNumTensors = 0
    Erase m_aoTensors
End Sub

Public Function GetTensor(ByVal lIndex As Long) As Tensor
    Const PROCEDURE_NAME As String = "Dataset.GetTensor"
    
    If lIndex < 1 Or lIndex > m_lNumTensors Then
        Err.Raise 9, PROCEDURE_NAME, "Tensor index must be between 1 and the number of tensors."
    End If
    Set GetTensor = m_aoTensors(lIndex)
End Function

Public Sub Add(ByVal oTensor As Tensor)
    Const PROCEDURE_NAME As String = "Dataset.Add"
    Dim lNumSamples As Long
    
    If oTensor Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If oTensor.NumDimensions < 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of dimensions must be greater than or equal to 2."
    End If
    lNumSamples = oTensor.Size(oTensor.NumDimensions)
    If m_lNumTensors = 0 Then
        m_lNumSamples = lNumSamples
    ElseIf lNumSamples <> m_lNumSamples Then
        Err.Raise 5, PROCEDURE_NAME, "Number of samples in the Tensor must match the existing number of samples."
    End If
    m_lNumTensors = m_lNumTensors + 1
    ReDim Preserve m_aoTensors(1 To m_lNumTensors)
    Set m_aoTensors(m_lNumTensors) = oTensor
End Sub

Public Function Gather(ByVal vIndices As Variant) As SimpleDataset
    Dim i As Long
    Dim oResult As SimpleDataset
    
    Set oResult = New SimpleDataset
    For i = 1 To m_lNumTensors
        With m_aoTensors(i)
            oResult.Add .Gather(.NumDimensions, vIndices)
        End With
    Next i
    Set Gather = oResult
End Function

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get IDataset_NumSamples() As Long
        IDataset_NumSamples = NumSamples
    End Property
    
    Private Function IDataset_Gather(ByVal vIndices As Variant) As Object
        Set IDataset_Gather = Gather(vIndices)
    End Function
#Else
    Private Property Get IDataset_NumSamples() As Long
        IDataset_NumSamples = NumSamples
    End Property
    
    Private Function IDataset_Gather(ByVal vIndices As Variant) As SimpleDataset
        Set IDataset_Gather = Gather(vIndices)
    End Function
#End If


