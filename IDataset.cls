VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_LATE_BINDING_IN_INTERFACES = True

#If USE_LATE_BINDING_IN_INTERFACES Then
    Public Property Get NumSamples() As Long
        '
    End Property
    
    Public Function Item(ByVal lTensorIndex As Long) As Object
        '
    End Function
    
    Public Function Subset(ByVal vSampleIndices As Variant) As Object
        '
    End Function
#Else
    Public Property Get NumSamples() As Long
        '
    End Property
    
    Public Function Item(ByVal lTensorIndex As Long) As Tensor
        '
    End Function
    
    Public Function Subset(ByVal vSampleIndices As Variant) As IDataset
        '
    End Function
#End If


