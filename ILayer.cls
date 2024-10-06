VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_LATE_BINDING_IN_INTERFACES = True

#If USE_LATE_BINDING_IN_INTERFACES Then
    Public Property Get LearnableParameters() As Variant
        '
    End Property
    
    Public Function Forward(ByVal X As Object, _
                            Optional ByVal bIsTraining As Boolean) As Object
        '
    End Function
    
    Public Function Backward(ByVal dLdY As Object) As Object
        '
    End Function
#Else
    Public Property Get LearnableParameters() As Variant
        '
    End Property
    
    Public Function Forward(ByVal X As Tensor, _
                            Optional ByVal bIsTraining As Boolean) As Tensor
        '
    End Function
    
    Public Function Backward(ByVal dLdY As Tensor) As Tensor
        '
    End Function
#End If
