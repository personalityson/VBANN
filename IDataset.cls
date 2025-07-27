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
    
    Public Function Cache() As Object
        '
    End Function
    
    Public Function Gather(ByVal vIndices As Variant) As Object
        '
    End Function
    
    Public Function Subset(ByVal vIndices As Variant) As Object
        '
    End Function
#Else
    Public Property Get NumSamples() As Long
        '
    End Property
    
    Public Function Cache() As TensorDataset
        '
    End Function
    
    Public Function Gather(ByVal vIndices As Variant) As TensorDataset
        '
    End Function
    
    Public Function Subset(ByVal vIndices As Variant) As SubsetDataset
        '
    End Function
#End If
