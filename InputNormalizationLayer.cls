VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputNormalizationLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_LATE_BINDING_IN_INTERFACES = True

Implements ILayer
Implements ISerializable

Private m_alInputShape() As Long
Private m_dblEpsilon As Double
Private m_m As Tensor
Private m_v As Tensor

Public Property Get Parameters() As Variant
    Parameters = Array()
End Property

Public Sub Init(ByVal oTrainingSet As DataLoader, _
                Optional ByVal dblEpsilon As Double = 0.00001)
    Const PROCEDURE_NAME As String = "InputNormalizationLayer.Init"
    Dim X As Tensor
    
    If oTrainingSet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid DataLoader object is required."
    End If
    If oTrainingSet.NumBatches < 1 Then
        Err.Raise 5, PROCEDURE_NAME, "Number of batches in the training set must be greater than 0."
    End If
    Set X = oTrainingSet.GetBatch(1).GetTensor(1)
    m_alInputShape = X.Shape
    m_alInputShape(X.NumDimensions) = -1
    m_dblEpsilon = dblEpsilon
    Adapt oTrainingSet
End Sub

Public Function Forward(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "InputNormalizationLayer.Forward"
    
    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not X.ShapeEquals(m_alInputShape) Then
        Err.Raise 5, PROCEDURE_NAME, "Input tensor shape does not match the expected shape for forward pass."
    End If
    Set Forward = CalculateOutput(X)
End Function

Public Function Backward(ByVal X As Tensor, _
                         ByVal Y As Tensor, _
                         ByVal dLdY As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "InputNormalizationLayer.Backward"
    
    If dLdY Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    If Not dLdY.ShapeEquals(m_alInputShape) Then
        Err.Raise 5, PROCEDURE_NAME, "Gradient tensor shape does not match the expected shape for backward pass."
    End If
    Set Backward = CalculateInputError(dLdY)
End Function

Public Sub Save(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "InputNormalizationLayer.Save"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        .WriteLongArray m_alInputShape
        .WriteDouble m_dblEpsilon
        .WriteObject m_m
        .WriteObject m_v
    End With
End Sub

Public Sub Load(ByVal oWorksheet As Serializer)
    Const PROCEDURE_NAME As String = "InputNormalizationLayer.Load"
    
    If oWorksheet Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oWorksheet
        m_alInputShape = .ReadLongArray()
        m_dblEpsilon = .ReadDouble()
        Set m_m = .ReadObject()
        Set m_v = .ReadObject()
    End With
End Sub

Private Sub Adapt(ByVal oTrainingSet As DataLoader)
    Dim lBatch As Long
    Dim lBatchSize As Long
    Dim X As Tensor
    Dim oBatchSum As Tensor
    Dim oBatchSquaredSum As Tensor
    Dim oTotalSum As Tensor
    Dim oTotalSquaredSum As Tensor
    Dim lNumSamples As Long
    
    With oTrainingSet
        For lBatch = 1 To .NumBatches
            With .GetBatch(lBatch)
                Set X = .GetTensor(1)
                lBatchSize = X.Size(X.NumDimensions)
                Set oBatchSum = X.ReduceSum(X.NumDimensions)
                Set oBatchSquaredSum = VecPow2(X).ReduceSum(X.NumDimensions)
                If lBatch = 1 Then
                    Set oTotalSum = oBatchSum
                    Set oTotalSquaredSum = oBatchSquaredSum
                Else
                    Set oTotalSum = VecAdd(oTotalSum, oBatchSum)
                    Set oTotalSquaredSum = VecAdd(oTotalSquaredSum, oBatchSquaredSum)
                End If
                lNumSamples = lNumSamples + lBatchSize
            End With
        Next lBatch
    End With
    If lNumSamples > 0 Then
        Set m_m = VecDivC(oTotalSum, lNumSamples)
        Set m_v = VecSub(VecDivC(oTotalSquaredSum, lNumSamples), VecPow2(m_m))
    End If
End Sub

Private Function CalculateOutput(ByVal X As Tensor) As Tensor
    Dim lBatchSize As Long
    Dim m As Tensor
    Dim v As Tensor
    
    lBatchSize = X.Size(X.NumDimensions)
    Set m = m_m.Tile(2, lBatchSize)
    Set v = m_v.Tile(2, lBatchSize)
    Set CalculateOutput = VecDivSqrtAddC(VecSub(X, m), v, m_dblEpsilon)
End Function

Private Function CalculateInputError(ByVal dLdY As Tensor) As Tensor
    Dim lBatchSize As Long
    Dim v As Tensor
    
    lBatchSize = dLdY.Size(dLdY.NumDimensions)
    Set v = m_v.Tile(2, lBatchSize)
    Set CalculateInputError = VecDivSqrtAddC(dLdY, v, m_dblEpsilon)
End Function

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get ILayer_Parameters() As Variant
        ILayer_Parameters = Parameters
    End Property
    
    Private Function ILayer_Forward(ByVal X As Object, _
                                    Optional ByVal bIsTraining As Boolean) As Object
        Set ILayer_Forward = Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal X As Object, _
                                     ByVal Y As Object, _
                                     ByVal dLdY As Object) As Object
        Set ILayer_Backward = Backward(X, Y, dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Object)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Object)
        Load oWorksheet
    End Sub
#Else
    Private Property Get ILayer_Parameters() As Variant
        ILayer_Parameters = Parameters
    End Property
    
    Private Function ILayer_Forward(ByVal X As Tensor, _
                                    Optional ByVal bIsTraining As Boolean) As Tensor
        Set ILayer_Forward = Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal X As Tensor, _
                                     ByVal Y As Tensor, _
                                     ByVal dLdY As Tensor) As Tensor
        Set ILayer_Backward = Backward(X, Y, dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oWorksheet As Serializer)
        Save oWorksheet
    End Sub
    
    Private Sub ISerializable_Load(ByVal oWorksheet As Serializer)
        Load oWorksheet
    End Sub
#End If
