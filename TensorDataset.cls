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

Private m_lNumSamples As Long
Private m_oTensors As Object

Public Property Get NumSamples() As Long
    NumSamples = m_lNumSamples
End Property

Public Property Get Item(ByVal vKey As Variant) As Tensor
    Const PROCEDURE_NAME As String = "Dataset.Item"
    
    If Not m_oTensors.Exists(vKey) Then
        Err.Raise 9, PROCEDURE_NAME, "Specified key does not exist in the dataset."
    End If
    Set Item = m_oTensors.Item(vKey)
End Property

Public Sub Init()
    m_lNumSamples = 0
    Set m_oTensors = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Add(ByVal vKey As Variant, _
               ByVal oTensor As Tensor)
    Const PROCEDURE_NAME As String = "Dataset.Add"
    Dim lNumSamples As Long
    
    If oTensor Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If oTensor.NumDimensions < 2 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of dimensions must be greater than or equal to 2."
    End If
    lNumSamples = oTensor.Size(oTensor.NumDimensions)
    If m_oTensors.Count = 0 Then
        m_lNumSamples = lNumSamples
    ElseIf lNumSamples <> m_lNumSamples Then
        Err.Raise 5, PROCEDURE_NAME, "Number of samples in the Tensor must match the existing number of samples."
    End If
    Set m_oTensors.Item(vKey) = oTensor
End Sub

Public Function Gather(ByVal vIndices As Variant) As TensorDataset
    Dim vKey As Variant
    Dim oTensor As Tensor
    Dim oResult As TensorDataset
    
    Set oResult = New TensorDataset
    oResult.Init
    For Each vKey In m_oTensors.Keys
        Set oTensor = m_oTensors.Item(vKey)
        oResult.Add vKey, oTensor.Gather(oTensor.NumDimensions, vIndices)
    Next vKey
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
    
    Private Function IDataset_Gather(ByVal vIndices As Variant) As TensorDataset
        Set IDataset_Gather = Gather(vIndices)
    End Function
#End If
