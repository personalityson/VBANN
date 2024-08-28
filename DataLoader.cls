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

Private m_oDataset As IDataset
Private m_lBatchSize As Long
Private m_lNumBatches As Long
Private m_alItinerary() As Long

Public Property Get BatchSize() As Long
    BatchSize = m_lBatchSize
End Property

Public Property Get NumBatches() As Long
    NumBatches = m_lNumBatches
End Property

Public Sub Init(ByVal oDataset As IDataset, _
                ByVal lBatchSize As Long)
    Const PROCEDURE_NAME As String = "DataLoader.Init"
    
    If oDataset Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Dataset object is required."
    End If
    If lBatchSize < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Batch size must be greater than 0."
    End If
    Set m_oDataset = oDataset
    m_lBatchSize = lBatchSize
    m_lNumBatches = -Int(-m_oDataset.NumSamples / m_lBatchSize)
    InitItinerary
End Sub

'https://en.wikipedia.org/wiki/Fisher%E2%80%93Yates_shuffle
Public Sub Shuffle()
    Dim i As Long
    Dim j As Long
    Dim lTemp As Long
    
    For i = m_oDataset.NumSamples To 2 Step -1
        j = Int(Rnd() * i) + 1
        lTemp = m_alItinerary(i)
        m_alItinerary(i) = m_alItinerary(j)
        m_alItinerary(j) = lTemp
    Next i
End Sub

Public Function GetBatch(ByVal lIndex As Long) As SimpleDataset
    Const PROCEDURE_NAME As String = "DataLoader.GetBatch"
    Dim lStart As Long
    Dim lLength As Long
    Dim alIndices() As Long
    
    If lIndex < 1 Or lIndex > m_lNumBatches Then
        Err.Raise 9, PROCEDURE_NAME, "Batch index must be between 1 and the number of batches."
    End If
    lStart = (lIndex - 1) * m_lBatchSize + 1
    lLength = MinLng2(m_lBatchSize, m_oDataset.NumSamples - lStart + 1)
    ReDim alIndices(1 To lLength)
    CopyMemory alIndices(1), m_alItinerary(lStart), lLength * SIZEOF_LONG
    Set GetBatch = m_oDataset.Gather(alIndices)
End Function

Private Sub InitItinerary()
    Dim i As Long
    
    If m_oDataset.NumSamples < 1 Then
        Erase m_alItinerary
    Else
        ReDim m_alItinerary(1 To m_oDataset.NumSamples)
        For i = 1 To m_oDataset.NumSamples
            m_alItinerary(i) = i
        Next i
    End If
End Sub
