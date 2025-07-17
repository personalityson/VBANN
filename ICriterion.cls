VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICriterion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_LATE_BINDING_IN_INTERFACES = True

#If USE_LATE_BINDING_IN_INTERFACES Then
    Public Function Loss(ByVal Y As Object, _
                         ByVal T As Object) As Object
        '
    End Function
    
    Public Function Gradient(ByVal Y As Object, _
                             ByVal T As Object) As Object
        '
    End Function
    
    Public Function Hessian(ByVal Y As Object, _
                            ByVal T As Object) As Object
        '
    End Function
#Else
    Public Function Loss(ByVal Y As Tensor, _
                         ByVal T As Tensor) As Tensor
        '
    End Function
    
    Public Function Gradient(ByVal Y As Tensor, _
                             ByVal T As Tensor) As Tensor
        '
    End Function
    
    Public Function Hessian(ByVal Y As Tensor, _
                            ByVal T As Tensor) As Tensor
        '
    End Function
#End If
