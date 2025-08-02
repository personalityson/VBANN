VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dependencies
Private m_oDataset As IDataset

'Configuration
Private m_lBatchSize As Long
Private m_bDropRemainder As Boolean

'State
Private m_alIndices() As Long
Private m_lCursor As Long

Public Property Get NumSamples() As Long
    NumSamples = m_oDataset.NumSamples
End Property

Public Property Get BatchSize() As Long
    BatchSize = m_lBatchSize
End Property

Public Sub Init(ByVal oDataset As IDataset, _
                ByVal lBatchSize As Long, _
                Optional ByVal bDropRemainder As Boolean)
    Const PROCEDURE_NAME As String = "DataLoader.Init"
    
    If oDataset Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid IDataset object is required."
    End If
    If lBatchSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Batch size must be >= 1."
    End If
    Clear
    Set m_oDataset = oDataset
    m_lBatchSize = lBatchSize
    m_bDropRemainder = bDropRemainder
    Me.Reset
End Sub

Public Sub Reset()
    m_alIndices = GetIdentityPermutationArray(m_oDataset.NumSamples)
    m_lCursor = 1
End Sub

Public Sub Shuffle()
    m_alIndices = GetRandomPermutationArray(m_oDataset.NumSamples)
    m_lCursor = 1
End Sub

Public Function NextBatch(ByRef oBatch As TensorDataset) As Boolean
    Dim lLength As Long
    Dim alBatchIndices() As Long

    lLength = MinLng2(m_lBatchSize, m_oDataset.NumSamples - m_lCursor + 1)
    If lLength < IIf(m_bDropRemainder, m_lBatchSize, 1) Then
        Exit Function
    End If
    ReDim alBatchIndices(1 To lLength)
    CopyMemory alBatchIndices(1), m_alIndices(m_lCursor), lLength * SIZEOF_LONG
    m_lCursor = m_lCursor + lLength
    Set oBatch = m_oDataset.Gather(alBatchIndices)
    NextBatch = True
End Function

Private Sub Clear()
    Set m_oDataset = Nothing
    m_lBatchSize = 0
    m_bDropRemainder = False
    Erase m_alIndices
    m_lCursor = 0
End Sub
